# WWF Seaweed Public Sentiment

This repository contains a reproducible Excel-to-OpenAI workflow for classifying aquaculture news articles. The workflow reads a human-coded Excel workbook, applies fixed prompt files with a configured OpenAI GPT model, and writes model labels and supporting evidence quotes back into a copied output workbook.

## Methodology

This workflow treats automated article coding as a reproducible extension of the human review protocol, not as a retraining exercise. The source workbook contains double-coded validation examples for article relevance, article category, and sentiment. The human coding instructions were translated into fixed prompts for a specified OpenAI GPT model, then refined by reviewing cases where model outputs disagreed with reviewer labels.

The method has three classification tasks:

- Article relevance: determine whether an item is a real English news story, opinion piece, or editorial that substantively discusses aquaculture, and whether the primary subject is seaweed aquaculture, other aquaculture, or irrelevant to the coding task.
- Sentiment: classify the headline-level sentiment toward aquaculture as positive, negative, or neutral. The article body is used as context for interpreting the headline, but body-only claims do not override a neutral headline.
- Article category: classify articles in the sentiment sheet as primarily about seaweed aquaculture or other aquaculture.

The main methodological challenge is that the human categories are intentionally simple, while the articles are not. Many articles mix neutral reporting with promotional claims, regulatory process, controversy, or economic projections. Early prompt versions tended to over-classify positive sentiment when an article body included government or industry claims about growth, jobs, investment, or future benefits. Reviewing disagreements showed that the human sentiment task was anchored to headline sentiment, so the current sentiment method gives priority to the headline and treats the article body as supporting context only.

The neutral category is the hardest boundary. It often includes routine announcements, administrative updates, project planning, funding, permitting, formalization, business transactions, and factual descriptions. Those items can contain favorable or unfavorable material in the body, but the method classifies them as neutral unless the headline itself clearly frames aquaculture positively or negatively. This makes the automated coding more consistent with the human instructions and reduces the tendency to infer sentiment from background details.

The approach works best when the goal is consistent large-scale classification against a documented coding scheme. It is auditable because each model label is accompanied by a supporting quote and is compared directly with the human reviewer labels in the output workbook. It is also reproducible because the model, prompts, input sheets, input columns, and output columns are fixed in configuration. If the model or prompts are changed later, that should be treated as a new version of the coding method and documented separately.

The validation results should be interpreted in light of reviewer agreement. Positive and negative cases are usually clearer, while neutral cases have more human disagreement and remain the most difficult class. The workbook therefore reports whether the model matched at least one reviewer and includes per-label match rates, rather than hiding these differences behind a single overall accuracy number.

## Output Workbook

The output workbook is a copy of the input workbook with model-generated labels, evidence quotes, and reviewer-match summaries added to the configured sheets.

In the `Code for relevance` sheet, the relevance analysis classifies whether each article is substantively about seaweed aquaculture, primarily about other aquaculture, or irrelevant to the aquaculture coding task. The output columns are:

- `gpt_5_4_mini_relevance`: model relevance/category label, using `irrelevant`, `other aquaculture`, or `seaweed aquaculture`.
- `gpt_5_4_mini_relevance_quote`: short quote supporting the relevance/category label.

In the `Code for sentiment` sheet, the sentiment analysis classifies the sentiment expressed by the article headline as positive, neutral, or negative. The article body is provided only as context for interpreting ambiguous headline wording. This task is intended to align with the human coding instructions for headline sentiment rather than the author's writing tone or full-article sentiment. The output columns are:

- `gpt_5_4_mini_sentiment`: model sentiment label, using `positive`, `neutral`, or `negative`.
- `gpt_5_4_mini_sentiment_quote`: short quote supporting the sentiment label.

The `Code for sentiment` sheet also includes an article category analysis. This classifies whether the article is primarily about seaweed aquaculture or other aquaculture. The output columns are:

- `gpt_5_4_mini_category`: model article-category label, using `seaweed aquaculture` or `other aquaculture`.
- `gpt_5_4_mini_category_quote`: short quote supporting the article-category label.

When the workbook includes the expected human reviewer columns, the script also adds reviewer-match validation columns. For each configured analysis, `{model}_{analysis}_matches_reviewer` records whether the model label matched either reviewer on that row. The adjacent `{model}_{analysis}_summary_metric` and `{model}_{analysis}_summary_value` columns report the overall match rate and each label-specific match rate using visible Excel formulas. These formulas are intended to make the comparison auditable inside the workbook.

## Technical Details

For reproducibility, the model name, prompt files, input sheets, input columns, and output columns are all recorded in a YAML configuration file. The same configuration should be used for all years or article batches included in a single analysis.

Each configured analysis combines the article headline with the article body text and sends that text to the model with the task-specific prompt. The response is constrained to contain a label, confidence value, one-sentence rationale, and short evidence quote. The script writes the label and evidence quote to the output workbook.

The workflow caches OpenAI responses by model, prompt text, article text, and response schema. This prevents repeated API calls for identical requests and makes long runs resumable. Transient OpenAI or network failures are retried, malformed responses are retried, and successful responses are written to the local cache as they complete.

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

The output workbook is a copy of the input workbook with the configured GPT output columns added or populated in the relevant sheets. When reviewer columns are present, the workbook also includes reviewer-match and match-rate summary columns for relevance, sentiment, and category. All worksheets are saved at 100% zoom.

## Runtime Notes

The script skips rows where either configured output column already has a value. To rerun an analysis from scratch, start from the original input workbook or clear the relevant GPT output columns.

OpenAI requests are cached. If a run is interrupted, the next run should reuse cached responses for requests that already completed. This is especially useful for long runs or unstable network connections.

Progress bars are shown for each configured analysis. Retry messages are printed when an OpenAI request times out, receives a transient connection error, or returns malformed JSON. A message such as `retrying in 1.5s` means the script will wait 1.5 seconds before starting the next attempt. It does not mean the next attempt will finish within 1.5 seconds.

The openpyxl warning `Data Validation extension is not supported and will be removed` can appear when reading workbooks with Excel data validation extensions. This warning is separate from the OpenAI analysis and does not indicate that model classification failed.
