# WWF Seaweed Public Sentiment

This repository is being refactored into a single Excel-to-OpenAI workflow.

The planned command-line interface will read an input Excel spreadsheet, an OpenAI API key file, a relevance prompt file, and a sentiment prompt file, then write a single output Excel spreadsheet.

## Placeholder Usage

```powershell
python analyze_public_sentiment.py `
  --input-excel path\to\input.xlsx `
  --openai-key-file path\to\openai_key.txt `
  --relevance-prompt-file path\to\relevance_prompt.txt `
  --sentiment-prompt-file path\to\sentiment_prompt.txt `
  --output-excel path\to\output.xlsx
```

The script currently validates the interface shape only. Model calls, prompt content, input column mapping, and output schema will be implemented in a later change.
