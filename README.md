# WWF Seaweed Public Sentiment

This repository is being refactored into a single Excel-to-OpenAI workflow.

The planned command-line interface will read a YAML configuration file, validate the configured input spreadsheet, OpenAI key file, and prompt files, then eventually write a single output Excel spreadsheet.

## Placeholder Usage

```powershell
python analyze_public_sentiment.py configs\analysis_2026_06_05.yaml
```

The script currently parses and validates the configuration only. OpenAI calls and Excel processing will be implemented in a later change.
