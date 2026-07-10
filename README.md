# Advanced Meta-Prompting Chart Engine

This project is a Python tool that turns a webpage or YouTube link into a polished chart and Excel workbook. It uses Gemini to analyze tabular data or text, then generates a PNG chart and an editable Excel chart.

## Features

- Accepts a webpage URL or YouTube link
- Detects usable data tables on webpages
- Supports popup-based input for a more interactive experience
- Recommends a chart type such as bar, line, pie, scatter, area, or doughnut
- Creates:
  - a PNG chart image
  - an Excel workbook with an embedded chart

## Requirements

Install the required Python packages:

```bash
pip install requests beautifulsoup4 pandas matplotlib google-genai pydantic openpyxl yt-dlp
```

## Setup

1. Open the project folder.
2. Make sure your Gemini API key is set in the script before running it.
3. Run the script:

```bash
python main.py
```

## Usage

When the script starts:

1. Paste a webpage or YouTube URL.
2. Choose Simple or Advanced mode.
3. Pick a table if the page contains multiple usable tables.
4. Choose a chart type if prompted.
5. Choose how many items to include.

The script will generate:

- a chart image named `chart.png`
- an Excel file named like `chart_YYYYMMDD_HHMMSS.xlsx`

## Notes

- The popup workflow uses tkinter, so it works best in a normal desktop Python environment.
- If YouTube links are used, `yt-dlp` must be installed.
- The chart appearance is designed to be clearer and more presentation-friendly for users.
