"""Run OpenAI analysis for configured Excel workbook rows."""

import argparse
from concurrent.futures import ThreadPoolExecutor
from concurrent.futures import FIRST_COMPLETED
from concurrent.futures import wait
import hashlib
import json
import random
import re
import time
from dataclasses import dataclass
from dataclasses import field
from datetime import datetime
from pathlib import Path
from typing import Any
from typing import Callable


TIMESTAMP_TOKEN = "{TIMESTAMP}"
TIMESTAMP_FORMAT = "%Y-%m-%d-%H-%M-%S"
MODEL_TOKEN = "{model}"
DEFAULT_MAX_WORKERS = 4
DEFAULT_OPENAI_REQUEST_TIMEOUT_SECONDS = 120.0
DEFAULT_STALL_LOG_SECONDS = 30.0
MAX_OPENAI_ATTEMPTS = 5
RETRYABLE_STATUS_CODES = {408, 409, 429, 500, 502, 503, 504}
EXCEL_ILLEGAL_CHARACTERS_RE = re.compile(r"[\000-\010]|[\013-\014]|[\016-\037]")
SENTIMENT_ANALYSIS_NAME = "sentiment"
REVIEWER_SENTIMENT_COLUMNS = ("Reviewer1_Sentiment", "Reviewer2_Sentiment")
SENTIMENT_LABELS = ("positive", "negative", "neutral")
ANALYSIS_RESPONSE_SCHEMA = {
    "type": "object",
    "properties": {
        "label": {"type": "string"},
        "confidence": {"type": "string", "enum": ["high", "medium", "low"]},
        "one_sentence_rationale": {"type": "string"},
        "evidence_quote": {"type": "string"},
    },
    "required": [
        "label",
        "confidence",
        "one_sentence_rationale",
        "evidence_quote",
    ],
    "additionalProperties": False,
}


class ConfigError(ValueError):
    """Error raised when the analysis configuration is invalid."""


@dataclass(frozen=True)
class AnalysisConfig:
    """Configuration for one worksheet analysis."""

    name: str
    sheet_name: str
    prompt_file: Path
    headline_column: str
    body_column: str | None
    output_label_column: str
    output_evidence_quote_column: str


@dataclass(frozen=True)
class AnalysisWorkItem:
    """Prepared article text for a future model call."""

    analysis_name: str
    worksheet_name: str
    row_number: int
    output_label_column: int
    output_evidence_quote_column: int
    text: str


@dataclass(frozen=True)
class AnalysisResult:
    """OpenAI result for a prepared worksheet row."""

    item: AnalysisWorkItem
    label: str
    evidence_quote: str


@dataclass(frozen=True)
class AnalysisBatch:
    """Prepared rows and prompt for one configured analysis."""

    config: AnalysisConfig
    prompt: str
    work_items: tuple[AnalysisWorkItem, ...]


@dataclass(frozen=True)
class RuntimeConfig:
    """Resolved runtime configuration for the analysis workflow."""

    input_file: Path
    output_file: Path
    openai_key_file: Path
    openai_cache_file: Path
    openai_model: str
    openai_request_timeout_seconds: float
    stall_log_seconds: float
    analyses: tuple[AnalysisConfig, ...]


@dataclass
class OpenAIAnalysisCache:
    """Local cache for parsed OpenAI analysis responses."""

    cache_file: Path
    responses: dict[str, dict[str, str]]
    unsaved_responses: dict[str, dict[str, str]] = field(default_factory=dict)
    dirty: bool = False

    @classmethod
    def load(cls, cache_file: Path) -> "OpenAIAnalysisCache":
        """Load cached OpenAI responses from disk.

        Args:
            cache_file: JSON cache file path.

        Returns:
            Loaded cache instance.

        Raises:
            ConfigError: If the cache file exists but is invalid.
        """
        cache_candidates = [cache_file]
        temporary_cache_file = cache_file.with_suffix(f"{cache_file.suffix}.tmp")
        if temporary_cache_file.exists():
            cache_candidates.append(temporary_cache_file)
        loaded_responses = {}
        for candidate_cache_file in cache_candidates:
            if not candidate_cache_file.exists():
                continue
            try:
                cache_data = json.loads(
                    candidate_cache_file.read_text(encoding="utf-8"))
            except json.JSONDecodeError as error:
                raise ConfigError(
                    f"OpenAI cache file is invalid JSON: {candidate_cache_file}"
                ) from error
            if not isinstance(cache_data, dict):
                raise ConfigError(
                    "OpenAI cache file must contain a mapping: "
                    f"{candidate_cache_file}")
            responses = cache_data.get("responses")
            if not isinstance(responses, dict):
                raise ConfigError(
                    "OpenAI cache file is missing a responses mapping: "
                    f"{candidate_cache_file}")
            if len(responses) > len(loaded_responses):
                loaded_responses = responses
        if not loaded_responses:
            loaded_responses = {}
        journal_file = cache_file.with_suffix(f"{cache_file.suffix}.jsonl")
        if journal_file.exists():
            for line_number, line in enumerate(
                    journal_file.read_text(encoding="utf-8").splitlines(),
                    start=1):
                if not line.strip():
                    continue
                try:
                    journal_entry = json.loads(line)
                except json.JSONDecodeError as error:
                    raise ConfigError(
                        f"OpenAI cache journal line {line_number} is invalid "
                        f"JSON: {journal_file}") from error
                cache_key = journal_entry.get("cache_key")
                response = journal_entry.get("response")
                if isinstance(cache_key, str) and isinstance(response, dict):
                    loaded_responses[cache_key] = response
                else:
                    raise ConfigError(
                        f"OpenAI cache journal line {line_number} is invalid: "
                        f"{journal_file}")
        return cls(cache_file=cache_file, responses=loaded_responses)

    def cache_key(self, model: str, prompt: str, article_text: str) -> str:
        """Build a cache key for an OpenAI analysis request.

        Args:
            model: OpenAI model name.
            prompt: System prompt text.
            article_text: User article text.

        Returns:
            SHA-256 cache key for the request.
        """
        cache_payload = json.dumps(
            {
                "model": model,
                "prompt": prompt,
                "article_text": article_text,
                "schema": ANALYSIS_RESPONSE_SCHEMA,
            },
            ensure_ascii=False,
            sort_keys=True,
            separators=(",", ":"),
        )
        return hashlib.sha256(cache_payload.encode("utf-8")).hexdigest()

    def save(self) -> None:
        """Write changed cache entries to disk."""
        if not self.dirty:
            return
        self.cache_file.parent.mkdir(parents=True, exist_ok=True)
        journal_file = self.cache_file.with_suffix(f"{self.cache_file.suffix}.jsonl")
        with journal_file.open("a", encoding="utf-8") as cache_journal:
            for cache_key, response in self.unsaved_responses.items():
                cache_journal.write(json.dumps(
                    {
                        "cache_key": cache_key,
                        "response": response,
                    },
                    ensure_ascii=False,
                    sort_keys=True,
                    separators=(",", ":"),
                ))
                cache_journal.write("\n")
            cache_journal.flush()
        self.unsaved_responses.clear()
        self.dirty = False


def parse_args() -> argparse.Namespace:
    """Parse command-line arguments.

    Returns:
        Parsed command-line arguments.
    """
    parser = argparse.ArgumentParser(
        description=(
            "Run configured OpenAI analyses for the seaweed public sentiment "
            "Excel workflow."
        )
    )
    parser.add_argument(
        "config",
        type=Path,
        help="Path to the YAML configuration file.",
    )
    parser.add_argument(
        "--limit-analysis-calls",
        type=int,
        default=None,
        help=(
            "Maximum OpenAI calls to make per configured analysis. Useful for "
            "debugging."
        ),
    )
    parser.add_argument(
        "--max-workers",
        type=int,
        default=DEFAULT_MAX_WORKERS,
        help=(
            "Maximum concurrent OpenAI calls. Lower this if you see rate "
            "limit errors."
        ),
    )
    return parser.parse_args()


def resolve_config_path(path_value: str, config_dir: Path) -> Path:
    """Resolve a configured path relative to the configuration file.

    Args:
        path_value: Path value from the YAML configuration.
        config_dir: Directory containing the YAML configuration file.

    Returns:
        Absolute path for the configured path.
    """
    configured_path = Path(path_value)
    if configured_path.is_absolute():
        return configured_path
    return (config_dir / configured_path).resolve()


def load_config(config_path: Path) -> dict[str, Any]:
    """Load a YAML configuration file.

    Args:
        config_path: Path to the YAML configuration file.

    Returns:
        Parsed configuration mapping.

    Raises:
        ConfigError: If the file is missing, empty, or does not contain a mapping.
    """
    if not config_path.exists():
        raise ConfigError(f"Configuration file not found: {config_path}")
    try:
        import yaml
    except ImportError as error:
        raise ConfigError(
            "PyYAML is required to parse configuration files. "
            "Install dependencies with `pip install -r requirements.txt`.") from error

    with config_path.open("r", encoding="utf-8") as config_file:
        config = yaml.safe_load(config_file)
    if not isinstance(config, dict):
        raise ConfigError(
            f"Configuration must be a YAML mapping: {config_path}")
    return config


def require_mapping(parent: dict[str, Any], key: str, context: str) -> dict[str, Any]:
    """Read a required nested mapping from the configuration.

    Args:
        parent: Mapping that should contain the requested key.
        key: Key to read.
        context: Human-readable configuration context for error messages.

    Returns:
        Nested mapping value.

    Raises:
        ConfigError: If the value is missing or is not a mapping.
    """
    value = parent.get(key)
    if not isinstance(value, dict):
        raise ConfigError(f"Missing or invalid mapping: {context}.{key}")
    return value


def require_string(parent: dict[str, Any], key: str, context: str) -> str:
    """Read a required string from the configuration.

    Args:
        parent: Mapping that should contain the requested key.
        key: Key to read.
        context: Human-readable configuration context for error messages.

    Returns:
        String value.

    Raises:
        ConfigError: If the value is missing or is not a string.
    """
    value = parent.get(key)
    if not isinstance(value, str) or not value.strip():
        raise ConfigError(f"Missing or invalid string: {context}.{key}")
    return value


def validate_required_file(path: Path, label: str) -> None:
    """Validate that a required configured file exists.

    Args:
        path: Resolved file path.
        label: Human-readable file label for error messages.

    Raises:
        ConfigError: If the file does not exist.
    """
    if not path.exists():
        raise ConfigError(f"{label} not found: {path}")


def model_column_prefix(model: str) -> str:
    """Convert a model name into a spreadsheet column-name prefix.

    Args:
        model: OpenAI model name from the configuration.

    Returns:
        Lowercase model name with non-alphanumeric characters converted to
        underscores.
    """
    sanitized_characters = [
        character.lower() if character.isalnum() else "_"
        for character in model
    ]
    return "_".join(
        segment for segment in "".join(sanitized_characters).split("_")
        if segment)


def resolve_output_column_name(column_name: str, model: str) -> str:
    """Resolve configured output column placeholders.

    Args:
        column_name: Configured output column name.
        model: OpenAI model name from the configuration.

    Returns:
        Output column name with supported placeholders resolved.
    """
    return column_name.replace(MODEL_TOKEN, model_column_prefix(model))


def validate_config(config: dict[str, Any], config_path: Path) -> RuntimeConfig:
    """Validate configured paths and return the resolved runtime configuration.

    Args:
        config: Parsed YAML configuration mapping.
        config_path: Path to the YAML configuration file.

    Returns:
        Resolved runtime configuration.

    Raises:
        ConfigError: If required configuration fields or files are missing.
    """
    config_dir = config_path.resolve().parent
    openai_config = require_mapping(config, "openai", "config")
    analyses_config = require_mapping(config, "analyses", "config")

    input_file = resolve_config_path(
        require_string(config, "input_file", "config"), config_dir)
    output_file_template = require_string(config, "output_file", "config")
    output_file = resolve_config_path(
        output_file_template.replace(
            TIMESTAMP_TOKEN,
            datetime.now().strftime(TIMESTAMP_FORMAT)),
        config_dir)
    key_file = resolve_config_path(
        require_string(openai_config, "key", "config.openai"), config_dir)
    openai_model = require_string(openai_config, "model", "config.openai")
    cache_file_value = openai_config.get("cache_file")
    if cache_file_value is None:
        openai_cache_file = config_path.resolve().with_suffix(".openai_cache.json")
    elif isinstance(cache_file_value, str) and cache_file_value.strip():
        openai_cache_file = resolve_config_path(cache_file_value, config_dir)
    else:
        raise ConfigError("Missing or invalid string: config.openai.cache_file")
    request_timeout_seconds = openai_config.get(
        "request_timeout_seconds", DEFAULT_OPENAI_REQUEST_TIMEOUT_SECONDS)
    if (
            not isinstance(request_timeout_seconds, (int, float)) or
            request_timeout_seconds <= 0):
        raise ConfigError(
            "Missing or invalid positive number: "
            "config.openai.request_timeout_seconds")
    stall_log_seconds = openai_config.get(
        "stall_log_seconds", DEFAULT_STALL_LOG_SECONDS)
    if not isinstance(stall_log_seconds, (int, float)) or stall_log_seconds <= 0:
        raise ConfigError(
            "Missing or invalid positive number: config.openai.stall_log_seconds")

    validate_required_file(input_file, "Input Excel file")
    validate_required_file(key_file, "OpenAI key file")

    analyses = []
    for analysis_name, analysis_config in analyses_config.items():
        if not isinstance(analysis_config, dict):
            raise ConfigError(
                f"Missing or invalid mapping: config.analyses.{analysis_name}")
        context = f"config.analyses.{analysis_name}"
        prompt_file = resolve_config_path(
            require_string(analysis_config, "prompt_file", context),
            config_dir)
        require_string(analysis_config, "sheet_name", context)
        input_columns = require_mapping(analysis_config, "input_columns", context)
        output_columns = require_mapping(
            analysis_config, "output_columns", context)
        require_string(input_columns, "headline", f"{context}.input_columns")
        body_column = None
        if "body" in input_columns and input_columns["body"] is not None:
            body_column = require_string(
                input_columns, "body", f"{context}.input_columns")
        require_string(output_columns, "label", f"{context}.output_columns")
        require_string(
            output_columns, "evidence_quote", f"{context}.output_columns")
        validate_required_file(prompt_file, f"{analysis_name} prompt file")
        analyses.append(AnalysisConfig(
            name=analysis_name,
            sheet_name=require_string(analysis_config, "sheet_name", context),
            prompt_file=prompt_file,
            headline_column=require_string(
                input_columns, "headline", f"{context}.input_columns"),
            body_column=body_column,
            output_label_column=resolve_output_column_name(
                require_string(
                    output_columns, "label", f"{context}.output_columns"),
                openai_model),
            output_evidence_quote_column=resolve_output_column_name(
                require_string(
                    output_columns,
                    "evidence_quote",
                    f"{context}.output_columns"),
                openai_model),
        ))

    return RuntimeConfig(
        input_file=input_file,
        output_file=output_file,
        openai_key_file=key_file,
        openai_cache_file=openai_cache_file,
        openai_model=openai_model,
        openai_request_timeout_seconds=float(request_timeout_seconds),
        stall_log_seconds=float(stall_log_seconds),
        analyses=tuple(analyses),
    )


def map_header_columns(worksheet: Any) -> dict[str, int]:
    """Map worksheet header names to one-based column indexes.

    Args:
        worksheet: openpyxl worksheet to inspect.

    Returns:
        Mapping from header name to one-based column index.
    """
    header_map = {}
    for cell in worksheet[1]:
        if cell.value is not None:
            header_map[str(cell.value)] = cell.column
    return header_map


def ensure_output_column(
        worksheet: Any, header_map: dict[str, int], column_name: str) -> int:
    """Ensure an output column exists on a worksheet.

    Args:
        worksheet: openpyxl worksheet to update.
        header_map: Existing worksheet header mapping.
        column_name: Output column header to create if missing.

    Returns:
        One-based column index for the output column.
    """
    if column_name in header_map:
        return header_map[column_name]
    column_index = worksheet.max_column + 1
    worksheet.cell(row=1, column=column_index).value = column_name
    header_map[column_name] = column_index
    return column_index


def add_sentiment_match_outputs(
        workbook: Any, runtime_config: RuntimeConfig) -> None:
    """Add reviewer-match validation outputs for the sentiment analysis.

    Args:
        workbook: openpyxl workbook after configured analyses have been written.
        runtime_config: Resolved runtime configuration used for the run.
    """
    sentiment_config = next(
        (
            analysis_config
            for analysis_config in runtime_config.analyses
            if analysis_config.name == SENTIMENT_ANALYSIS_NAME
        ),
        None,
    )
    if sentiment_config is None:
        return

    worksheet = workbook[sentiment_config.sheet_name]
    header_map = map_header_columns(worksheet)
    reviewer_columns = [
        header_map[column_name]
        for column_name in REVIEWER_SENTIMENT_COLUMNS
        if column_name in header_map
    ]
    if (
            sentiment_config.output_label_column not in header_map or
            len(reviewer_columns) != len(REVIEWER_SENTIMENT_COLUMNS)):
        return

    predicted_column = header_map[sentiment_config.output_label_column]
    match_column = ensure_output_column(
        worksheet,
        header_map,
        f"{sentiment_config.output_label_column}_matches_reviewer",
    )
    summary_column_names = {
        "total": f"{sentiment_config.output_label_column}_match_rate",
        **{
            label: f"{sentiment_config.output_label_column}_{label}_match_rate"
            for label in SENTIMENT_LABELS
        },
    }
    summary_columns = {
        "total": ensure_output_column(
            worksheet,
            header_map,
            summary_column_names["total"],
        ),
        **{
            label: ensure_output_column(
                worksheet,
                header_map,
                summary_column_names[label],
            )
            for label in SENTIMENT_LABELS
        },
    }
    summary_metric_column = ensure_output_column(
        worksheet,
        header_map,
        f"{sentiment_config.output_label_column}_summary_metric",
    )
    summary_value_column = ensure_output_column(
        worksheet,
        header_map,
        f"{sentiment_config.output_label_column}_summary_value",
    )
    match_counts = {
        label: {"matched": 0, "total": 0}
        for label in ("total", *SENTIMENT_LABELS)
    }

    for row_number in range(2, worksheet.max_row + 1):
        predicted = worksheet.cell(
            row=row_number, column=predicted_column).value
        predicted_label = (
            str(predicted).strip().lower()
            if predicted is not None and str(predicted).strip()
            else "")
        reviewer_labels = {
            str(reviewer_value).strip().lower()
            for reviewer_column in reviewer_columns
            if (
                reviewer_value := worksheet.cell(
                    row=row_number, column=reviewer_column).value
            ) is not None and str(reviewer_value).strip()
        }
        reviewer_labels &= set(SENTIMENT_LABELS)
        if predicted_label not in SENTIMENT_LABELS or not reviewer_labels:
            worksheet.cell(row=row_number, column=match_column).value = None
            continue

        matched_reviewer = predicted_label in reviewer_labels
        worksheet.cell(
            row=row_number, column=match_column).value = matched_reviewer
        match_counts["total"]["total"] += 1
        if matched_reviewer:
            match_counts["total"]["matched"] += 1
        for label in SENTIMENT_LABELS:
            if label in reviewer_labels:
                match_counts[label]["total"] += 1
                if predicted_label == label:
                    match_counts[label]["matched"] += 1

    for summary_row_number, (label, column_index) in enumerate(
            summary_columns.items(), start=2):
        counts = match_counts[label]
        match_rate = (
            counts["matched"] / counts["total"]
            if counts["total"]
            else None)
        cell = worksheet.cell(row=2, column=column_index)
        cell.value = match_rate
        cell.number_format = "0.0%"
        worksheet.cell(
            row=summary_row_number,
            column=summary_metric_column).value = summary_column_names[label]
        summary_value_cell = worksheet.cell(
            row=summary_row_number,
            column=summary_value_column)
        summary_value_cell.value = match_rate
        summary_value_cell.number_format = "0.0%"


def create_openai_client(
        api_key_file: Path, request_timeout_seconds: float) -> Any:
    """Create an OpenAI client from a configured API key file.

    Args:
        api_key_file: Path to a text file containing an OpenAI API key.
        request_timeout_seconds: Maximum seconds for each OpenAI request attempt.

    Returns:
        OpenAI client instance.

    Raises:
        ConfigError: If the OpenAI SDK is not installed or the key file is empty.
    """
    try:
        from openai import OpenAI
    except ImportError as error:
        raise ConfigError(
            "openai is required to run analysis calls. "
            "Install dependencies with `pip install -r requirements.txt`."
        ) from error

    api_key = api_key_file.read_text(encoding="utf-8").strip()
    if not api_key:
        raise ConfigError(f"OpenAI key file is empty: {api_key_file}")
    return OpenAI(
        api_key=api_key,
        timeout=request_timeout_seconds,
        max_retries=0,
    )


def is_retryable_openai_error(error: Exception) -> bool:
    """Determine whether an OpenAI exception should be retried.

    Args:
        error: Exception raised while calling OpenAI.

    Returns:
        Whether retrying the request is appropriate.
    """
    status_code = getattr(error, "status_code", None)
    if status_code is None:
        return True
    return status_code in RETRYABLE_STATUS_CODES


def call_openai_analysis(
        client: Any,
        model: str,
        prompt: str,
        article_text: str,
        request_label: str,
        retry_logger: Callable[[str], None] | None = None) -> dict[str, str]:
    """Classify one article with OpenAI and parse the JSON result.

    Args:
        client: OpenAI client instance.
        model: OpenAI model configured for the run.
        prompt: Analysis prompt text.
        article_text: Combined headline and body text to analyze.
        request_label: Human-readable row label for progress messages.
        retry_logger: Optional logger for retry/backoff messages.

    Returns:
        Parsed JSON response from the model.

    Raises:
        ConfigError: If the response is not valid JSON with required fields.
    """
    for attempt_index in range(MAX_OPENAI_ATTEMPTS):
        try:
            response = client.responses.create(
                model=model,
                input=[
                    {
                        "role": "system",
                        "content": prompt,
                    },
                    {
                        "role": "user",
                        "content": f"Article text:\n\n{article_text}",
                    },
                ],
                text={
                    "format": {
                        "type": "json_schema",
                        "name": "article_analysis_result",
                        "strict": True,
                        "schema": ANALYSIS_RESPONSE_SCHEMA,
                    }
                },
            )
        except Exception as error:
            final_attempt = attempt_index == MAX_OPENAI_ATTEMPTS - 1
            if final_attempt or not is_retryable_openai_error(error):
                raise ConfigError(f"OpenAI request failed: {error}") from error
            backoff_seconds = min(60, 2 ** attempt_index) + random.random()
            if retry_logger is not None:
                retry_logger(
                    f"{request_label}: OpenAI attempt {attempt_index + 1}/"
                    f"{MAX_OPENAI_ATTEMPTS} failed; retrying in "
                    f"{backoff_seconds:.1f}s ({error})")
            time.sleep(backoff_seconds)
            if retry_logger is not None:
                retry_logger(
                    f"{request_label}: starting OpenAI attempt "
                    f"{attempt_index + 2}/{MAX_OPENAI_ATTEMPTS}")
            continue

        final_attempt = attempt_index == MAX_OPENAI_ATTEMPTS - 1
        try:
            result = json.loads(response.output_text)
        except json.JSONDecodeError as error:
            if final_attempt:
                raise ConfigError(
                    "OpenAI response was not valid JSON after "
                    f"{MAX_OPENAI_ATTEMPTS} attempts: {response.output_text}"
                ) from error
            backoff_seconds = min(60, 2 ** attempt_index) + random.random()
            if retry_logger is not None:
                retry_logger(
                    f"{request_label}: OpenAI attempt {attempt_index + 1}/"
                    f"{MAX_OPENAI_ATTEMPTS} returned invalid JSON; retrying "
                    f"in {backoff_seconds:.1f}s")
            time.sleep(backoff_seconds)
            if retry_logger is not None:
                retry_logger(
                    f"{request_label}: starting OpenAI attempt "
                    f"{attempt_index + 2}/{MAX_OPENAI_ATTEMPTS}")
            continue

        if not isinstance(result, dict):
            if final_attempt:
                raise ConfigError(
                    "OpenAI response JSON must be an object after "
                    f"{MAX_OPENAI_ATTEMPTS} attempts: {response.output_text}")
            backoff_seconds = min(60, 2 ** attempt_index) + random.random()
            if retry_logger is not None:
                retry_logger(
                    f"{request_label}: OpenAI attempt {attempt_index + 1}/"
                    f"{MAX_OPENAI_ATTEMPTS} returned non-object JSON; "
                    f"retrying in {backoff_seconds:.1f}s")
            time.sleep(backoff_seconds)
            if retry_logger is not None:
                retry_logger(
                    f"{request_label}: starting OpenAI attempt "
                    f"{attempt_index + 2}/{MAX_OPENAI_ATTEMPTS}")
            continue

        required_fields = {
            "label",
            "confidence",
            "one_sentence_rationale",
            "evidence_quote",
        }
        missing_fields = sorted(required_fields.difference(result))
        if missing_fields:
            if final_attempt:
                raise ConfigError(
                    "OpenAI response missing required field(s) after "
                    f"{MAX_OPENAI_ATTEMPTS} attempts: "
                    f"{', '.join(missing_fields)}")
            backoff_seconds = min(60, 2 ** attempt_index) + random.random()
            if retry_logger is not None:
                retry_logger(
                    f"{request_label}: OpenAI attempt {attempt_index + 1}/"
                    f"{MAX_OPENAI_ATTEMPTS} returned incomplete JSON; "
                    f"retrying in {backoff_seconds:.1f}s")
            time.sleep(backoff_seconds)
            if retry_logger is not None:
                retry_logger(
                    f"{request_label}: starting OpenAI attempt "
                    f"{attempt_index + 2}/{MAX_OPENAI_ATTEMPTS}")
            continue
        return result

    raise ConfigError(f"OpenAI request failed after {MAX_OPENAI_ATTEMPTS} attempts")


def collect_analysis_items(
        worksheet: Any,
        analysis_config: AnalysisConfig,
        header_map: dict[str, int],
        output_label_column: int,
        output_evidence_quote_column: int,
        limit_analysis_calls: int | None) -> list[AnalysisWorkItem]:
    """Collect worksheet rows that need OpenAI analysis.

    Args:
        worksheet: openpyxl worksheet to inspect.
        analysis_config: Analysis configuration for this worksheet.
        header_map: Worksheet header mapping.
        output_label_column: One-based output label column index.
        output_evidence_quote_column: One-based output evidence quote column index.
        limit_analysis_calls: Optional maximum rows to analyze for this analysis.

    Returns:
        Prepared analysis work items.
    """
    headline_column = header_map[analysis_config.headline_column]
    body_column = (
        header_map[analysis_config.body_column]
        if analysis_config.body_column
        else None)
    work_items = []
    for row_number in range(2, worksheet.max_row + 1):
        if (
                limit_analysis_calls is not None and
                len(work_items) >= limit_analysis_calls):
            break
        existing_label = worksheet.cell(
            row=row_number, column=output_label_column).value
        existing_quote = worksheet.cell(
            row=row_number, column=output_evidence_quote_column).value
        if existing_label or existing_quote:
            continue
        headline = worksheet.cell(row=row_number, column=headline_column).value
        body = (
            worksheet.cell(row=row_number, column=body_column).value
            if body_column
            else None)
        text_parts = [
            str(value).strip()
            for value in (headline, body)
            if value is not None and str(value).strip()
        ]
        if text_parts:
            work_items.append(AnalysisWorkItem(
                analysis_name=analysis_config.name,
                worksheet_name=analysis_config.sheet_name,
                row_number=row_number,
                output_label_column=output_label_column,
                output_evidence_quote_column=output_evidence_quote_column,
                text="\n\n".join(text_parts),
            ))
    return work_items


def run_parallel_openai_analysis(
        client: Any,
        model: str,
        analysis_batches: list[AnalysisBatch],
        max_workers: int,
        cache: OpenAIAnalysisCache,
        stall_log_seconds: float) -> list[AnalysisResult]:
    """Analyze prepared rows across all analyses with progress reporting.

    Args:
        client: OpenAI client instance.
        model: OpenAI model configured for the run.
        analysis_batches: Prepared worksheet rows grouped by analysis.
        max_workers: Maximum concurrent OpenAI calls across all analyses.
        cache: OpenAI response cache.
        stall_log_seconds: Seconds without a completed request before logging.

    Returns:
        Results for completed OpenAI calls.

    Raises:
        ConfigError: If tqdm is unavailable or an OpenAI call fails.
    """
    try:
        from tqdm import tqdm
    except ImportError as error:
        raise ConfigError(
            "tqdm is required for progress reporting. "
            "Install dependencies with `pip install -r requirements.txt`."
        ) from error

    all_work_items = [
        item
        for analysis_batch in analysis_batches
        for item in analysis_batch.work_items
    ]
    if not all_work_items:
        return []

    worker_count = min(max_workers, len(all_work_items))
    results = []
    with ThreadPoolExecutor(max_workers=worker_count) as executor:
        future_to_items = {}
        future_to_cache_key = {}
        future_to_started_at = {}
        future_to_label = {}
        pending_futures_by_cache_key = {}
        progress_bars = {}
        try:
            for position, analysis_batch in enumerate(analysis_batches):
                if not analysis_batch.work_items:
                    continue
                progress_bars[analysis_batch.config.name] = tqdm(
                    total=len(analysis_batch.work_items),
                    desc=f"{analysis_batch.config.name} analysis",
                    unit="article",
                    position=position,
                )

            scheduled_work_items = []
            largest_batch_size = max(
                len(analysis_batch.work_items)
                for analysis_batch in analysis_batches
            )
            for item_index in range(largest_batch_size):
                for analysis_batch in analysis_batches:
                    if item_index < len(analysis_batch.work_items):
                        scheduled_work_items.append(
                            (
                                analysis_batch,
                                analysis_batch.work_items[item_index],
                            ))

            for analysis_batch, item in scheduled_work_items:
                cache_key = cache.cache_key(model, analysis_batch.prompt, item.text)
                cached_result = cache.responses.get(cache_key)
                if isinstance(cached_result, dict):
                    results.append(AnalysisResult(
                        item=item,
                        label=cached_result["label"],
                        evidence_quote=cached_result["evidence_quote"],
                    ))
                    progress_bars[item.analysis_name].update(1)
                    continue
                if cache_key in pending_futures_by_cache_key:
                    pending_future = pending_futures_by_cache_key[cache_key]
                    future_to_items[pending_future].append(item)
                    continue
                request_label = f"{item.analysis_name} row {item.row_number}"
                future = executor.submit(
                    call_openai_analysis,
                    client,
                    model,
                    analysis_batch.prompt,
                    item.text,
                    request_label,
                    tqdm.write)
                pending_futures_by_cache_key[cache_key] = future
                future_to_items[future] = [item]
                future_to_cache_key[future] = cache_key
                future_to_started_at[future] = time.monotonic()
                future_to_label[future] = request_label

            pending_futures = set(future_to_items)
            while pending_futures:
                done_futures, pending_futures = wait(
                    pending_futures,
                    timeout=stall_log_seconds,
                    return_when=FIRST_COMPLETED,
                )
                if not done_futures:
                    oldest_future = min(
                        pending_futures,
                        key=lambda pending_future: future_to_started_at[
                            pending_future],
                    )
                    oldest_elapsed_seconds = (
                        time.monotonic() - future_to_started_at[oldest_future])
                    tqdm.write(
                        "No OpenAI requests completed in "
                        f"{stall_log_seconds:g}s; "
                        f"{len(pending_futures)} request(s) still in flight. "
                        "Oldest pending request: "
                        f"{future_to_label[oldest_future]} "
                        f"({oldest_elapsed_seconds:g}s).")
                    continue
                for future in done_futures:
                    items = future_to_items[future]
                    try:
                        result = future.result()
                    except ConfigError as error:
                        item = items[0]
                        raise ConfigError(
                            f"{item.analysis_name} row {item.row_number} "
                            f"failed: {error}") from error
                    cache.responses[future_to_cache_key[future]] = dict(result)
                    cache.unsaved_responses[future_to_cache_key[future]] = dict(
                        result)
                    cache.dirty = True
                    cache.save()
                    for item in items:
                        results.append(AnalysisResult(
                            item=item,
                            label=result["label"],
                            evidence_quote=result["evidence_quote"],
                        ))
                        progress_bars[item.analysis_name].update(1)
        finally:
            for progress_bar in progress_bars.values():
                progress_bar.close()
    return results


def collect_analysis_batches(
        workbook: Any,
        runtime_config: RuntimeConfig,
        limit_analysis_calls: int | None) -> list[AnalysisBatch]:
    """Collect all configured worksheet rows that need OpenAI analysis.

    Args:
        workbook: openpyxl workbook to inspect and update with output headers.
        runtime_config: Resolved runtime configuration.
        limit_analysis_calls: Optional maximum OpenAI calls per analysis.

    Returns:
        Analysis batches with prompt text and prepared worksheet rows.

    Raises:
        ConfigError: If a configured sheet or input column is missing.
    """
    analysis_batches = []
    for analysis_config in runtime_config.analyses:
        prompt = analysis_config.prompt_file.read_text(encoding="utf-8").strip()
        if not prompt:
            raise ConfigError(f"Prompt file is empty: {analysis_config.prompt_file}")
        if analysis_config.sheet_name not in workbook.sheetnames:
            raise ConfigError(
                f"Worksheet not found for {analysis_config.name}: "
                f"{analysis_config.sheet_name}")
        worksheet = workbook[analysis_config.sheet_name]
        header_map = map_header_columns(worksheet)
        missing_columns = [
            column_name for column_name in (
                analysis_config.headline_column,
                analysis_config.body_column,
            )
            if column_name and column_name not in header_map
        ]
        if missing_columns:
            raise ConfigError(
                f"Missing input column(s) on {analysis_config.sheet_name}: "
                f"{', '.join(missing_columns)}")

        output_label_column = ensure_output_column(
            worksheet, header_map, analysis_config.output_label_column)
        output_evidence_quote_column = ensure_output_column(
            worksheet,
            header_map,
            analysis_config.output_evidence_quote_column)
        work_items = collect_analysis_items(
            worksheet,
            analysis_config,
            header_map,
            output_label_column,
            output_evidence_quote_column,
            limit_analysis_calls,
        )
        analysis_batches.append(AnalysisBatch(
            config=analysis_config,
            prompt=prompt,
            work_items=tuple(work_items),
        ))
    return analysis_batches


def process_workbook(
        runtime_config: RuntimeConfig,
        limit_analysis_calls: int | None,
        max_workers: int) -> list[AnalysisWorkItem]:
    """Run configured analyses and write results to an output workbook.

    Args:
        runtime_config: Resolved runtime configuration.
        limit_analysis_calls: Optional maximum OpenAI calls per analysis.
        max_workers: Maximum concurrent OpenAI calls.

    Returns:
        Article text items sent to OpenAI.

    Raises:
        ConfigError: If a configured sheet or input column is missing.
    """
    try:
        from openpyxl import load_workbook
    except ImportError as error:
        raise ConfigError(
            "openpyxl is required to process Excel files. "
            "Install dependencies with `pip install -r requirements.txt`."
        ) from error

    if max_workers < 1:
        raise ConfigError("--max-workers must be at least 1")
    client = create_openai_client(
        runtime_config.openai_key_file,
        runtime_config.openai_request_timeout_seconds,
    )
    cache = OpenAIAnalysisCache.load(runtime_config.openai_cache_file)
    workbook = load_workbook(runtime_config.input_file)
    analysis_batches = collect_analysis_batches(
        workbook, runtime_config, limit_analysis_calls)
    analysis_results = run_parallel_openai_analysis(
        client,
        runtime_config.openai_model,
        analysis_batches,
        max_workers,
        cache,
        runtime_config.stall_log_seconds,
    )
    cache.save()
    for analysis_result in analysis_results:
        item = analysis_result.item
        worksheet = workbook[item.worksheet_name]
        worksheet.cell(
            row=item.row_number,
            column=item.output_label_column).value = (
                EXCEL_ILLEGAL_CHARACTERS_RE.sub("", analysis_result.label))
        worksheet.cell(
            row=item.row_number,
            column=item.output_evidence_quote_column).value = (
                EXCEL_ILLEGAL_CHARACTERS_RE.sub(
                    "", analysis_result.evidence_quote))

    add_sentiment_match_outputs(workbook, runtime_config)
    runtime_config.output_file.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(runtime_config.output_file)
    return [analysis_result.item for analysis_result in analysis_results]


def main() -> None:
    """Run configured OpenAI analyses and save the output workbook."""
    args = parse_args()
    config_path = args.config.resolve()
    try:
        runtime_config = validate_config(load_config(config_path), config_path)
        work_items = process_workbook(
            runtime_config, args.limit_analysis_calls, args.max_workers)
    except ConfigError as error:
        raise SystemExit(f"Configuration error: {error}") from error

    print("Configuration parsed successfully.")
    print("Resolved paths:")
    print(f"- input_file: {runtime_config.input_file}")
    print(f"- output_file: {runtime_config.output_file}")
    print(f"- openai_key_file: {runtime_config.openai_key_file}")
    print(f"- openai_cache_file: {runtime_config.openai_cache_file}")
    print(
        "- openai_request_timeout_seconds: "
        f"{runtime_config.openai_request_timeout_seconds:g}")
    print(f"- stall_log_seconds: {runtime_config.stall_log_seconds:g}")
    for analysis_config in runtime_config.analyses:
        print(f"- {analysis_config.name}_prompt_file: {analysis_config.prompt_file}")
    print(f"Analyzed {len(work_items)} article text item(s).")
    print(f"Output workbook saved to: {runtime_config.output_file}")


if __name__ == "__main__":
    main()
