"""Placeholder command-line interface for the future OpenAI analysis workflow."""

import argparse
from pathlib import Path


def main() -> None:
    """Parse the planned Excel-to-OpenAI workflow arguments."""
    parser = argparse.ArgumentParser(
        description=(
            "Analyze seaweed public sentiment from an Excel spreadsheet using "
            "OpenAI models and prompt files."
        )
    )
    parser.add_argument(
        "--input-excel",
        required=True,
        type=Path,
        help="Path to the input Excel spreadsheet.",
    )
    parser.add_argument(
        "--openai-key-file",
        required=True,
        type=Path,
        help="Path to a text file containing the OpenAI API key.",
    )
    parser.add_argument(
        "--relevance-prompt-file",
        required=True,
        type=Path,
        help="Path to the text prompt used for relevance classification.",
    )
    parser.add_argument(
        "--sentiment-prompt-file",
        required=True,
        type=Path,
        help="Path to the text prompt used for sentiment classification.",
    )
    parser.add_argument(
        "--output-excel",
        required=True,
        type=Path,
        help="Path where the output Excel spreadsheet will be written.",
    )
    parser.parse_args()
    print("Excel-to-OpenAI sentiment analysis is not implemented yet.")


if __name__ == "__main__":
    main()
