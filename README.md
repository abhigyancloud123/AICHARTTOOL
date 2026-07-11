# Advanced Meta-Prompting Chart Engine Generator

A Python application that automatically analyzes web pages and YouTube videos to extract tables, analyze data with Google Gemini AI, and generate professional charts and Excel files.

## Features

- **Webpage Table Extraction**: Automatically detects and extracts structured data tables from web pages
- **YouTube Video Analysis**: Downloads and analyzes YouTube videos to extract data insights
- **AI-Powered Analysis**: Uses Google Gemini to intelligently recommend chart types and labels
- **Interactive UI**: Popup-based interface with back-navigation support
- **Multiple Chart Types**: Bar, line, and pie charts with customizable item counts
- **Excel Export**: Generates formatted Excel files with embedded charts
- **Batch Testing**: `test_runner.py` for validating URLs against the pipeline
- **Robust Column Matching**: Fuzzy column name matching and normalization
- **Better Error Messages**: Clear feedback for quota, auth, and data issues

## Requirements

- Python 3.10+
- pandas
- requests
- beautifulsoup4
- matplotlib
- openpyxl
- google-generativeai (Gemini API)
- tkinter (usually included with Python)
- yt-dlp (optional, for YouTube support)
- PyQt5 (optional, fallback for matplotlib backend)

## Installation

1. **Clone or download the project**:
   ```bash
   cd "SPECIAL PROJECT"
   ```

2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```
   
   Or manually:
   ```bash
   pip install pandas requests beautifulsoup4 matplotlib openpyxl google-generativeai yt-dlp
   ```

3. **Set up your Gemini API key**:
   
   **Option A: Environment Variable** (Recommended)
   ```powershell
   $env:GEMINI_API_KEY = "your-api-key-here"
   ```
   
   **Option B: Local File**
   Create one of these files in the project folder with your API key:
   - `GEMINI_API_KEY.txt`
   - `api_key.txt`
   - `API KEY FOR PROJECT.txt`
   
   (The existing `API KEY FOR PROJECT.txt` file will be checked automatically.)

4. **Get a Gemini API Key**:
   - Go to [Google AI Studio](https://aistudio.google.com)
   - Click "Get API key"
   - Create a new API key for your project
   - Paste it in your environment or file

## Usage

### Main Application

Run the interactive chart generator:
```bash
python main.py
```

**Workflow**:
1. Enter a webpage URL or YouTube video link
2. Choose "simple" or "advanced" view mode
3. For webpages: select a table (if multiple exist)
4. Review the table preview and Gemini's recommendations
5. Choose your chart type (bar/line/pie)
6. Select how many items to display
7. Generated files appear in the project folder:
   - `chart_YYYYMMDD_HHMMSS.png` (image)
   - `chart_YYYYMMDD_HHMMSS.xlsx` (Excel with embedded chart)

**Back Button**: Press "Back" on any popup (except the initial URL prompt) to return to the previous screen.

### Batch Testing

Test multiple URLs at once using `test_runner.py`:

**Direct URLs**:
```bash
python test_runner.py "https://example.com/page1" "https://example.com/page2"
```

**From a File**:
```bash
python test_runner.py --file urls.txt
```
Create `urls.txt` with one URL per line (lines starting with `#` are comments).

**From stdin**:
```bash
python test_runner.py --stdin
```
Paste URLs one per line, then press Ctrl+Z and Enter.

**Output Example**:
```
Batch Test Summary
==================
- URL: https://en.wikipedia.org/wiki/List_of_countries_by_GDP_(nominal)
  Status: PASS
  Tables found: 5
  Gemini recommendation: bar

- URL: https://example.com/broken
  Status: FAIL
  Tables found: 0
  Reason: No usable tables found

Overall: 1 passed, 1 failed
```

## File Structure

```
SPECIAL PROJECT/
├── main.py                     # Main application (GUI and chart generation)
├── test_runner.py             # Batch URL testing script
├── chart.py                   # (Optional chart utilities)
├── README.md                  # This file
├── requirements.txt           # Python dependencies
├── LICENSE                    # Project license
├── API KEY FOR PROJECT.txt    # Gemini API key (local, optional)
└── chart_*.png / chart_*.xlsx # Generated output files
```

## Configuration

### Environment Variables

- `GEMINI_API_KEY`: Your Gemini API key (recommended for security)

### Default Limits

- **Bar/Line Charts**: Top 10 items by default
- **Pie Charts**: Top 6 items by default
- **YouTube Resolution**: 480p max (configurable in `download_youtube_video()`)

Modify these in `main.py`:
```python
DEFAULT_BAR_LINE_LIMIT = 10
DEFAULT_PIE_LIMIT = 6
```

## Troubleshooting

### "API Key missing" Error

**Cause**: No valid Gemini API key found.

**Fix**:
- Set `GEMINI_API_KEY` environment variable:
  ```powershell
  $env:GEMINI_API_KEY = "your-key"
  ```
- Or create `API KEY FOR PROJECT.txt` with your key
- Verify the key is valid at [Google AI Studio](https://aistudio.google.com)

### "Gemini API authentication failed" or "401 UNAUTHENTICATED"

**Cause**: API key is invalid or expired.

**Fix**:
- Get a new API key from [Google AI Studio](https://aistudio.google.com)
- Update your `GEMINI_API_KEY` environment variable or file
- Verify the key has no extra spaces or line breaks

### "Gemini API quota exceeded" or "429 RESOURCE_EXHAUSTED"

**Cause**: Your API quota is exhausted (free tier: 1,500 calls/day, 60 calls/min).

**Fix**:
- Upgrade your Gemini API plan at [Google AI Studio](https://aistudio.google.com)
- Wait for the quota to reset (usually next calendar day)
- Try with a new API key on a different project

### "No usable tables found"

**Cause**: The webpage doesn't have structured tables, or they were filtered out.

**Fix**:
- The app falls back to text extraction (parses the page text directly)
- Try a different URL with structured data (Wikipedia lists, statistics pages)
- Use "advanced" mode to see what was attempted

### "Couldn't reach that page"

**Cause**: Network error or invalid URL.

**Fix**:
- Check your internet connection
- Verify the URL is correct and publicly accessible
- Try a different URL

### Popup windows not appearing

**Cause**: tkinter not available or matplotlib backend issue.

**Fix**:
- Install tkinter:
  ```bash
  pip install tk
  ```
- On Ubuntu/Debian:
  ```bash
  sudo apt-get install python3-tk
  ```
- The app falls back to console input if tkinter isn't available

### YouTube download fails

**Cause**: yt-dlp not installed or video unavailable.

**Fix**:
- Install yt-dlp:
  ```bash
  pip install yt-dlp
  ```
- Verify the video is publicly accessible
- Check that it contains visual/statistical content for analysis

### Charts are blank or legend missing

**Cause**: Excel version compatibility with `openpyxl.chart.Legend`.

**Fix**:
- The app automatically detects and skips legend creation if unsupported
- Update openpyxl:
  ```bash
  pip install --upgrade openpyxl
  ```

## Advanced Features

### Column Name Fuzzy Matching

The app automatically matches Gemini's recommended columns to actual table columns using:
- Exact name matching
- Normalized text comparison (removes brackets, special chars)
- Partial string matching
- Falls back to the first non-label column if no match is found

### Back Navigation

Press "Back" on popup screens to return to the previous step without losing your data.

### Advanced Mode

Choose "advanced" view mode to see:
- Table preview (head rows)
- Gemini's selected columns
- Detailed column mapping info

### Chart Type Recommendations

Gemini uses these rules:
- **Line**: Multiple value columns tracking time series (years, quarters, dates)
- **Pie**: Single value column with ≤8 items representing parts of a whole
- **Bar**: Discrete entity comparisons, rankings, or flat categories

## API Limits & Costs

Google Gemini API has a [free tier](https://aistudio.google.com) with monthly quotas:
- **Free plan**: 60 API calls per minute, 1,500 calls per day
- **Paid plans**: Higher limits and priority support

Monitor your usage at [Google Cloud Console](https://console.cloud.google.com).

## Security Notes

- **Never commit your API key** to version control
- Use environment variables or local `.txt` files (added to `.gitignore`)
- The app reads from `API KEY FOR PROJECT.txt` by default for convenience
- For production, use a secure key management service (Google Secret Manager, etc.)

## Examples

### Test Wikipedia GDP Data
```bash
python test_runner.py "https://en.wikipedia.org/wiki/List_of_countries_by_GDP_(nominal)"
```

### Generate a chart from a webpage
```bash
python main.py
# Paste: https://en.wikipedia.org/wiki/List_of_best-selling_video_games
# Choose: simple
# Wait for analysis...
# Select chart type: bar
# Select item count: 10
```

### Batch test multiple URLs
```bash
# Create urls.txt with your list
python test_runner.py --file urls.txt
```

## Contributing

Feel free to modify and extend:
- Add new chart types in `save_chart_image()`
- Customize Gemini prompts in `analyze_table()` and `analyze_page_text()`
- Implement new export formats (PDF, SVG, etc.)

## License

See [LICENSE](LICENSE) file.

## Support

- Check error messages carefully — they describe the exact issue
- Review the `step()` and `info()` console output for detailed logs
- Test individual URLs with `test_runner.py` first before batch processing
- Run in "advanced" mode for more diagnostic info

---

**Last Updated**: 2026-07-11  
**Version**: 1.0
