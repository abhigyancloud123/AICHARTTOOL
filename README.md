# Advanced Meta-Prompting Chart Engine

This project is a Python script that fetches data from a webpage or YouTube link, analyzes it with Gemini, and generates:

- a chart image as PNG
- an Excel workbook with a chart

## Features

- Accepts a webpage URL or YouTube link
- Detects tables on webpages and extracts structured data
- Falls back to text extraction when no tables are found
- Recommends a chart type (bar, line, or pie)
- Saves a chart image and Excel output

## Requirements

Install the required Python packages:

```bash
pip install requests beautifulsoup4 pandas matplotlib google-genai pydantic openpyxl yt-dlp
```

## Usage

Run the script:

```bash
python main.py
```

Then provide:

1. a webpage or YouTube URL
2. the view mode (Simple or Advanced)
3. chart display choices when prompted

## Notes

- The script uses your Gemini API key from the code. Make sure it is valid and has access.
- For chart popups, a GUI backend such as TkAgg or QtAgg may be required depending on your environment.
- If YouTube support is used, `yt-dlp` must be installed.
