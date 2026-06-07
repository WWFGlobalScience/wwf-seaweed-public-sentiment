# WWF Seaweed Public Sentiment

This repository contains a reproducible workflow for classifying the subject and public sentiment of aquaculture news articles. The core classification routine, `call_openai_analysis()` in `analyze_public_sentiment.py`, is independent of the rest of the pipeline, which is currently engineered to process a specific Excel format.[^excel-format]

[^excel-format]: The current implementation reads an Excel workbook containing a `Code for relevance` sheet with `headline` and `body` columns and a `Code for sentiment` sheet with `headline` and `body` columns. It writes model labels and supporting evidence quotes back into configured output columns in a copied workbook.

## Methodology

Our method uses a fixed LLM-based classifier to reproduce the article-coding decisions defined in the human review protocol. Human coding instructions given to reviewers are translated into fixed prompts for a specified OpenAI GPT model (`gpt-5.4-mini`). We refined the prompts by reviewing cases where model outputs disagreed with reviewer labels, with revisions focused on better matching the written coding rules rather than retraining a model.

The method has three classification tasks related to analyzing a potential aquaculture news story:

1. Article relevance: determine whether an item is a news story, opinion piece, or editorial that substantively discusses aquaculture, and classify it as `seaweed aquaculture`, `other aquaculture`, or `irrelevant`.
2. Sentiment: classify the sentiment toward aquaculture as `positive`, `negative`, or `neutral`, following the headline-based coding protocol used by human reviewers. The article body is used as context for interpreting the headline, but body-only claims do not override a neutral headline.
3. Article category: classify articles selected for sentiment review as primarily about `seaweed aquaculture` or `other aquaculture`.

The main methodological challenge was aligning the model with the human coding protocol. The coding categories are intentionally simple, while the articles often combine neutral reporting with promotional claims, regulatory process, controversy, and economic projections. Our early prompt versions tended to classify articles as positive when the body included government or industry claims about growth, jobs, investment, or future benefits. Those labels were often plausible full-article sentiment classifications, but they did not always match the reviewer task, which asks reviewers to classify headline sentiment.

At the same time, reviewer labels suggest that article bodies are sometimes used to resolve ambiguous headlines. As an example, one reviewer classified the headline "Netherlands: Veramaris aims to reduce greenhouse gas emissions by nearly a quarter" as `other aquaculture`. The headline does not identify aquaculture on its own, but the article body discusses sustainable aquaculture growth, suggesting that the reviewer used body context to interpret the headline. We therefore designed the model to prioritize the headline for sentiment classification while using the body only to clarify ambiguous references. This preserves alignment with the human review protocol, even though a separate full-article sentiment task would also be technically straightforward.

The neutral category was the hardest category for the model to classify because it is defined mostly by the absence of clear positive or negative framing. It often includes routine announcements, administrative updates, project planning, funding, permitting, formalization, business transactions, and factual descriptions. These articles can contain favorable or unfavorable material in the body, but reviewers generally code them as neutral unless aquaculture itself is clearly framed positively or negatively. This main failure mode of the sentiment classifier is that it often treats contextual claims about growth, investment, jobs, environmental benefit, controversy, or risk as article sentiment (positive or negative), even when reviewers treat the headline as neutral.

Our validation results below are interpreted relative to human reviewer agreement. Positive and negative cases are generally more consistently coded, whereas neutral cases show greater disagreement and remain the most difficult class as discussed above. Accordingly, validation summaries report both whether the model matched at least one reviewer and label-specific match rates.

| Analysis | Match Rate |
|-|-:|
| Sentiment: overall | 79.0% |
| Sentiment: positive | 86.5% |
| Sentiment: negative | 75.0% |
| Sentiment: neutral | 25.5% |
| Category: overall | 94.3% |
| Category: seaweed aquaculture | 99.0% |
| Category: other aquaculture | 92.2% |
| Relevance: overall | 81.5% |
| Relevance: irrelevant | 60.7% |
| Relevance: other aquaculture | 60.5% |
| Relevance: seaweed aquaculture | 90.8% |

## Technical Details

For reproducibility, the model name, prompt files, input sheets, input columns, and output columns are all recorded in a YAML configuration file with an example given at `configs/example_analysis.yaml`. 

Each configured analysis combines the article headline with the article body text and sends that text to the model with the task-specific prompt. The response is constrained to contain a label, confidence value, one-sentence rationale, and short evidence quote. The script writes the label and evidence quote to the output workbook.

The workflow caches OpenAI responses by model, prompt text, article text, and response schema to prevent repeated API calls for identical requests and makes long runs resumable. Transient OpenAI or network failures are retried, malformed responses are retried, and successful responses are written to the local cache as they complete. [also say how the parallelization works]

### Configuration

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
python analyze_public_sentiment.py .\configs\analysis_2026_06_05.yaml -limit-analysis-calls 5 -max-workers 2
```

Run the full analysis:

```powershell
python analyze_public_sentiment.py .\configs\analysis_2026_06_05.yaml -max-workers 4
```

The script writes a new Excel workbook to the configured `output_file` path. With the example configuration, output files will look like:

```text
data/main_coding_sheet_analysis_2026-06-05-18-30-00.xlsx
```

The output workbook is a copy of the input workbook with the configured GPT output columns added or populated in the relevant sheets. When reviewer columns are present, the workbook also includes reviewer-match and match-rate summary columns for relevance, sentiment, and category.

## Runtime Notes

The script skips rows where either configured output column already has a value. To rerun an analysis from scratch, start from the original input workbook or clear the relevant GPT output columns.

As mentioned earlier, OpenAI requests are cached. If a run is interrupted, the next run should reuse cached responses for requests that already completed. This is especially useful for long runs or unstable network connections.

Progress bars are shown for each configured analysis. Retry messages are printed when an OpenAI request times out, receives a transient connection error, or returns malformed JSON. A message such as `retrying in 1.5s` means the script will wait 1.5 seconds before starting the next attempt. It does not mean the next attempt will finish within 1.5 seconds.

The openpyxl warning `Data Validation extension is not supported and will be removed` can appear when reading workbooks with Excel data validation extensions. This warning is separate from the OpenAI analysis and does not indicate that model classification failed.
