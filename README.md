# Advanced Meta-Prompting Chart Engine

 
**License:** MIT — see [LICENSE](LICENSE)

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

2. Authentication options (pick one):

- API key (quick test):

  - Edit `main.py` and set `API_KEY` to your API key, or set the environment variable for the current session:

    ```powershell
    $env:GEMINI_API_KEY="PASTE_YOUR_API_KEY_HERE"
    python main.py
    ```

  - To persist the env var across sessions:

    ```powershell
    setx GEMINI_API_KEY "PASTE_YOUR_API_KEY_HERE"
    # restart terminal / VS Code
    python main.py
    ```

- Application Default Credentials (recommended for development/production):

  - Install and configure the Google Cloud SDK, then run:

    ```powershell
    gcloud auth application-default login
    gcloud services enable generativelanguage.googleapis.com
    python main.py
    ```

- Local key file (not recommended):

  - You can place your key in `API KEY FOR PROJECT.txt` (the script may read this file for convenience). Remove this file after testing.

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

## License

This project is licensed under the MIT License — see [LICENSE](LICENSE) for details.
