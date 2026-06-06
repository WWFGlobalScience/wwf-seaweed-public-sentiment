# WWF Seaweed Public Sentiment

This repository is being refactored into a single Excel-to-OpenAI workflow.

The command-line interface reads a YAML configuration file, validates the configured input spreadsheet, OpenAI key file, and prompt files, then creates an output workbook copied from the input workbook with OpenAI analysis outputs added.

## Placeholder Usage

```powershell
python analyze_public_sentiment.py configs\analysis_2026_06_05.yaml
```

Use `--limit-analysis-calls N` to limit OpenAI calls per configured analysis during debugging. Use `--max-workers N` to control concurrent OpenAI calls; the default is 4.

OpenAI responses are cached by model, prompt text, article text, and response schema. By default the cache file is created next to the configuration file using the suffix `.openai_cache.json`. Set `openai.cache_file` to choose a different cache path:

```yaml
openai:
  model: "gpt-5.4-mini"
  key: "../secrets/openai_key.txt"
  cache_file: "../data/openai_analysis_cache.json"
  request_timeout_seconds: 120
  stall_log_seconds: 30
```

An analysis can use headline-only input by omitting `input_columns.body`. Output column names may use `{model}`, which is expanded from the configured OpenAI model name:

```yaml
analyses:
  category:
    sheet_name: "Code for sentiment"
    prompt_file: "../prompts/headline_category_classification_prompt.txt"
    input_columns:
      headline: "headline"
    output_columns:
      label: "{model}_category"
      evidence_quote: "{model}_category_quote"
```
