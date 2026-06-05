# WWF Seaweed Public Sentiment

This repository is being refactored into a single Excel-to-OpenAI workflow.

The command-line interface reads a YAML configuration file, validates the configured input spreadsheet, OpenAI key file, and prompt files, then creates an output workbook copied from the input workbook with the configured output columns added.

## Placeholder Usage

```powershell
python analyze_public_sentiment.py configs\analysis_2026_06_05.yaml
```

The script currently prepares the workbook scaffold only. OpenAI calls and model output population will be implemented in a later change.
