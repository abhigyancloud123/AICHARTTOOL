# Link → Chart Engine

Turn a webpage or YouTube link into a polished chart and an editable Excel workbook — powered by Gemini.

## What it does

1. **Paste a link** — a webpage (with a data table, or just plain text/prose) or a YouTube video
2. **Gemini analyzes it** using a step-by-step meta-prompting process:
   - Finds the right label/value columns (or extracts data directly from text/video)
   - Picks the best chart type (bar, line, or pie) using explicit, strict rules — not just a guess
   - Writes clear axis titles and decides whether a legend is needed
3. **You confirm or override** the chart type and how many items to include
4. **Two files are generated:**
   - `chart.png` — a polished, presentation-ready chart image with a legend always included
   - `chart_YYYYMMDD_HHMMSS.xlsx` — a real Excel file with your data *and* a native, fully editable Excel chart
5. **One-click access** — when it's done, click **"Open Excel File"** or **"Open Chart Image"** to jump straight to your result

## Features

- **Two data sources:** webpages (table-based or plain text) and YouTube videos (Gemini watches the video directly)
- **Smart table detection:** automatically filters out junk tables (navigation boxes, "authority control" footers) and lets you pick when a page has multiple real tables
- **Fuzzy column matching:** handles messy real-world column names (footnotes, extra whitespace, multi-level headers) so it doesn't break on typical Wikipedia-style tables
- **Meta-prompted chart selection:** Gemini is walked through an explicit rule order (pie → line → bar) instead of defaulting to bar out of habit
- **Always-on legends:** every bar/line chart includes a legend automatically, in both the PNG and the Excel chart
- **Back navigation:** every step (table choice, chart type, item count) has a Back button, so a misclick doesn't mean starting over
- **Popup-based UI:** everything runs through clean, styled popup windows — no need to type in a terminal
- **Graceful fallbacks:** if tkinter, a display backend, or yt-dlp isn't available, the tool degrades gracefully instead of crashing

## Requirements

Install the required Python packages:

```
pip install requests beautifulsoup4 pandas matplotlib google-genai pydantic openpyxl yt-dlp
```

## Setup

1. Open `main.py`
2. Find this line near the top:
   ```python
   API_KEY = "PASTE_YOUR_GEMINI_API_KEY_HERE"
   ```
3. Replace it with your own free Gemini API key (get one at [aistudio.google.com](https://aistudio.google.com))
4. Run it:
   ```
   python main.py
   ```

**Security note:** your API key will be sitting in plain text in `main.py` with this setup. Do **not** commit this file with your real key to a public GitHub repo — if you're sharing this project, replace the key with the placeholder text again before pushing.

## Usage

1. Run `python main.py`
2. Paste a webpage or YouTube link when the popup appears
3. Choose Simple or Advanced view mode (Advanced shows extra debug info in the terminal)
4. If the page has multiple tables, pick the one you want
5. Review Gemini's recommended chart type (or pick your own)
6. Choose how many items to include (or leave blank for a sensible default)
7. Click Open Excel File or Open Chart Image to see your result

## Model notes

- Default model: `gemini-2.0-flash-001` — chosen for broad availability, including on newer Google accounts that may not yet have access to the very latest preview models
- To use a different model, change `MODEL_NAME` near the top of `main.py`
- If you see a "not available to new users" error, your Google account may need phone verification, or you should stick with an established model like `gemini-2.0-flash-001` rather than the newest preview releases

## Known limitations

- This is a beta version — expect occasional rough edges in chart-type recommendations or data extraction on unusual pages
- YouTube video analysis works best on shorter videos with clearly visible on-screen data (rankings, stats, charts)
- Very large or heavily nested tables (e.g. pages with 10+ tables) may need manual table selection to get the right data

## License

MIT — see LICENSE for details.
