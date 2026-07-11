# Advanced Meta-Prompting Chart Engine

> BETA release. APIs and behavior may change; use for testing and development only.

This project is a Python tool that converts a webpage or YouTube link into a polished chart and an Excel workbook. It uses Gemini to analyze table or page text and then generates a PNG chart plus an editable Excel file.

## Features

- Accepts a webpage URL or YouTube link
- Detects usable data tables on webpages
- Analyzes text fallback when no table is available
- Recommends chart types such as bar, line, or pie
- Generates:
  - a PNG chart image
  - an Excel workbook with embedded chart data

## Requirements

Install the required Python packages:

```powershell
pip install requests beautifulsoup4 pandas matplotlib google-genai pydantic openpyxl yt-dlp
```

The script also uses `tkinter` for pop-up dialogs when available.

## Authentication

The code reads an API key from the environment variable `GEMINI_API_KEY`.

### Recommended: environment variable

In PowerShell:

```powershell
$env:GEMINI_API_KEY="PASTE_YOUR_API_KEY_HERE"
python main.py
```

To persist the key across sessions:

```powershell
setx GEMINI_API_KEY "PASTE_YOUR_API_KEY_HERE"
# restart terminal / VS Code
python main.py
```

### Alternative: Application Default Credentials (ADC)

If you prefer Google Cloud authentication:

```powershell
gcloud auth application-default login
gcloud services enable generativelanguage.googleapis.com
python main.py
```

### Note: do not hardcode secrets

Avoid placing API keys directly into source files. Use environment variables or secure local storage instead.

## Run

From the project folder:

```powershell
cd "C:\Users\DELL\Desktop\SPECIAL PROJECT"
python main.py
```

## Usage

1. Enter a webpage or YouTube URL.
2. Select `simple` or `advanced` mode.
3. If multiple tables are found, choose the one you want to chart.
4. Pick a chart type if prompted.
5. Choose how many items to include.

The script generates:

- `chart.png`
- `chart_YYYYMMDD_HHMMSS.xlsx`

## Troubleshooting

- If you see a missing table or label error, try a different page or URL.
- If `tkinter` is unavailable, the script falls back to terminal input.
- If Git or push commands are needed, use the integrated terminal in VS Code or Git Bash.

## License

This project is licensed under the MIT License — see [LICENSE](LICENSE) for details.
