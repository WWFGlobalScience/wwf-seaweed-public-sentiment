# WWF Seaweed Public Sentiment

This repository contains a reproducible Excel-to-OpenAI workflow for classifying aquaculture news articles. The workflow reads a human-coded Excel workbook, applies fixed prompt files with a configured OpenAI GPT model, and writes model labels and supporting evidence quotes back into a copied output workbook.

## Methodology

This workflow uses a fixed OpenAI GPT model and fixed prompt instructions to classify articles from a human-coded validation spreadsheet. Each configured analysis combines the article headline with the article body text, sends that text to the model with a task-specific prompt, and requires the model to return structured JSON containing a label, confidence value, one-sentence rationale, and short evidence quote. The script writes only the label and evidence quote to the output workbook so model classifications can be compared directly with existing human reviewer columns.

Three article-coding tasks are currently implemented.

In the `Code for relevance` sheet, the relevance analysis classifies whether each article is substantively about seaweed aquaculture, primarily about other aquaculture, or irrelevant to the aquaculture coding task. The output columns are:

- `gpt_5_4_mini_relevance`: model relevance/category label, using `irrelevant`, `other aquaculture`, or `seaweed aquaculture`.
- `gpt_5_4_mini_relevance_quote`: short quote supporting the relevance/category label.

In the `Code for sentiment` sheet, the sentiment analysis classifies the sentiment represented in the article as positive, neutral, or negative. This task is intended to align with the human coding instructions for article sentiment rather than the author's writing tone alone. The output columns are:

- `gpt_5_4_mini_sentiment`: model sentiment label, using `positive`, `neutral`, or `negative`.
- `gpt_5_4_mini_sentiment_quote`: short quote supporting the sentiment label.

The `Code for sentiment` sheet also includes an article category analysis. This classifies whether the article is primarily about seaweed aquaculture or other aquaculture. The output columns are:

- `gpt_5_4_mini_category`: model article-category label, using `seaweed aquaculture` or `other aquaculture`.
- `gpt_5_4_mini_category_quote`: short quote supporting the article-category label.

For reproducibility, the model name, prompt files, input sheets, input columns, and output columns are all recorded in a YAML configuration file. The same configuration should be used for all years or article batches included in a single analysis. If the model or prompts are changed in the future, that change should be treated as a new version of the classification method and documented separately.

The workflow also caches OpenAI responses by model, prompt text, article text, and response schema. This prevents repeated API calls for identical requests and makes long runs resumable. Transient OpenAI or network failures are retried, malformed JSON responses are retried, and successful responses are written to the local cache as they complete.

## Required Files

Place these files in the repository before running the workflow:

- Excel input workbook, for example `data/main_coding_sheet.xlsx`.
- OpenAI API key text file, for example `secrets/openai_key.txt`.
- YAML configuration file, for example `configs/analysis_2026_06_05.yaml`.
- Prompt files in `prompts/`:
  - `article_aquaculture_subject_classification_prompt.txt`
  - `article_subject_public_view_sentiment_prompt.txt`
  - `headline_category_classification_prompt.txt`

The `data/`, `configs/`, and `secrets/` directories are local working directories. Input workbooks, generated output workbooks, API keys, and cache files should not be committed to the repository.

## Configuration

The YAML configuration defines the input workbook, output workbook, OpenAI model, API key path, cache behavior, and each analysis to run. Paths are resolved relative to the YAML file.

Example:

```yaml
input_file: "../data/main_coding_sheet.xlsx"
output_file: "../data/main_coding_sheet_analysis_{TIMESTAMP}.xlsx"

openai:
  model: "gpt-5.4-mini"
  key: "../secrets/openai_key.txt"
  cache_file: "../data/openai_analysis_cache.json"
  request_timeout_seconds: 120
  stall_log_seconds: 30

analyses:
  relevance:
    sheet_name: "Code for relevance"
    prompt_file: "../prompts/article_aquaculture_subject_classification_prompt.txt"
    input_columns:
      headline: "headline"
      body: "body"
    output_columns:
      label: "gpt_5_4_mini_relevance"
      evidence_quote: "gpt_5_4_mini_relevance_quote"

  sentiment:
    sheet_name: "Code for sentiment"
    prompt_file: "../prompts/article_subject_public_view_sentiment_prompt.txt"
    input_columns:
      headline: "headline"
      body: "body"
    output_columns:
      label: "gpt_5_4_mini_sentiment"
      evidence_quote: "gpt_5_4_mini_sentiment_quote"

  category:
    sheet_name: "Code for sentiment"
    prompt_file: "../prompts/headline_category_classification_prompt.txt"
    input_columns:
      headline: "headline"
      body: "body"
    output_columns:
      label: "gpt_5_4_mini_category"
      evidence_quote: "gpt_5_4_mini_category_quote"
```

`{TIMESTAMP}` in `output_file` is replaced with the run time using `YYYY-MM-DD-HH-MM-SS`.

`cache_file` is optional. If omitted, the script creates a cache next to the configuration file using the suffix `.openai_cache.json`. The cache also uses a `.jsonl` journal file for incremental saves.

## How To Run

From the repository root, install the Python dependencies:

```powershell
pip install -r requirements.txt
```

Create the local directories if they do not already exist:

```powershell
mkdir data
mkdir configs
mkdir secrets
```

Put the source workbook in `data/`:

```text
data/main_coding_sheet.xlsx
```

Put the OpenAI API key in a plain text file:

```text
secrets/openai_key.txt
```

Create the YAML configuration file:

```text
configs/analysis_2026_06_05.yaml
```

Run a small debug batch first:

```powershell
python analyze_public_sentiment.py .\configs\analysis_2026_06_05.yaml --limit-analysis-calls 5 --max-workers 2
```

Run the full analysis:

```powershell
python analyze_public_sentiment.py .\configs\analysis_2026_06_05.yaml --max-workers 4
```

The script writes a new Excel workbook to the configured `output_file` path. With the example configuration, output files will look like:

```text
data/main_coding_sheet_analysis_2026-06-05-18-30-00.xlsx
```

The output workbook is a copy of the input workbook with the configured GPT output columns added or populated in the relevant sheets.

## Runtime Notes

The script skips rows where either configured output column already has a value. To rerun an analysis from scratch, start from the original input workbook or clear the relevant GPT output columns.

OpenAI requests are cached. If a run is interrupted, the next run should reuse cached responses for requests that already completed. This is especially useful for long runs or unstable network connections.

Progress bars are shown for each configured analysis. Retry messages are printed when an OpenAI request times out, receives a transient connection error, or returns malformed JSON. A message such as `retrying in 1.5s` means the script will wait 1.5 seconds before starting the next attempt. It does not mean the next attempt will finish within 1.5 seconds.

The openpyxl warning `Data Validation extension is not supported and will be removed` can appear when reading workbooks with Excel data validation extensions. This warning is separate from the OpenAI analysis and does not indicate that model classification failed.
