"""Validate configuration for the future OpenAI analysis workflow."""

import argparse
from datetime import datetime
from pathlib import Path
from typing import Any


TIMESTAMP_TOKEN = "{TIMESTAMP}"
TIMESTAMP_FORMAT = "%Y-%m-%d-%H-%M-%S"


class ConfigError(ValueError):
    """Error raised when the analysis configuration is invalid."""


def parse_args() -> argparse.Namespace:
    """Parse command-line arguments.

    Returns:
        Parsed command-line arguments.
    """
    parser = argparse.ArgumentParser(
        description=(
            "Validate the YAML configuration for the seaweed public sentiment "
            "analysis workflow."
        )
    )
    parser.add_argument(
        "config",
        type=Path,
        help="Path to the YAML configuration file.",
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


def validate_config(config: dict[str, Any], config_path: Path) -> dict[str, Path]:
    """Validate configured paths and return the resolved file locations.

    Args:
        config: Parsed YAML configuration mapping.
        config_path: Path to the YAML configuration file.

    Returns:
        Mapping of workflow path labels to resolved paths.

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
    require_string(openai_config, "model", "config.openai")

    resolved_paths = {
        "input_file": input_file,
        "output_file": output_file,
        "openai_key_file": key_file,
    }
    validate_required_file(input_file, "Input Excel file")
    validate_required_file(key_file, "OpenAI key file")

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
        require_string(input_columns, "body", f"{context}.input_columns")
        require_string(output_columns, "label", f"{context}.output_columns")
        require_string(
            output_columns, "evidence_quote", f"{context}.output_columns")
        validate_required_file(prompt_file, f"{analysis_name} prompt file")
        resolved_paths[f"{analysis_name}_prompt_file"] = prompt_file

    return resolved_paths


def main() -> None:
    """Validate the configured files and exit before analysis work begins."""
    args = parse_args()
    config_path = args.config.resolve()
    try:
        resolved_paths = validate_config(load_config(config_path), config_path)
    except ConfigError as error:
        raise SystemExit(f"Configuration error: {error}") from error

    print("Configuration parsed successfully.")
    print("Resolved paths:")
    for label, path in resolved_paths.items():
        print(f"- {label}: {path}")
    print("OpenAI calls and Excel processing are not implemented yet.")


if __name__ == "__main__":
    main()
