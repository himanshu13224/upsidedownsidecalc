# Stock Price Fetcher — Streamlit Web App

Fetch NSE/BSE daily/monthly close prices and build an Excel workbook with Returns & Risk Analysis.
Data comes from tej-bazaar (official NSE/BSE Bhavcopy published on HuggingFace) — free, no API key.

## Files
- `app.py` — Streamlit application
- `requirements.txt` — Python dependencies
- `README.md` — this file

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```

Then open the URL Streamlit prints (usually `http://localhost:8501`).

## Deploy to Streamlit Cloud (free)
1. Push this folder (or just `app.py` + `requirements.txt`) to a GitHub repo.
2. Go to https://share.streamlit.io and connect your GitHub account.
3. Click "New app", select your repo, branch, and `app.py` as the main file.
4. Click Deploy. You'll get a public URL anyone can open.

## What the app does
1. Asks for stock symbol (e.g. `tcs`, `suzlon`), exchange (`NSE`/`BSE`), years, and frequency (Daily/Monthly).
2. Downloads the relevant year-Parquet files from tej-bazaar on HuggingFace.
3. Filters for the symbol and builds an Excel workbook (`.xlsx`) in memory.
4. Optionally adds:
   - **Returns sheet** — Date, Close Price, Returns (live Excel formula `=(B3-B2)/B2`)
   - **Risk Analysis sheet** — expected return, difference, negative/positive returns using `=IF(...)`, squared values, and downside/upside risk summary
5. Lets you download the Excel file directly from the browser.

## Notes
- Monthly mode returns the last close price of each month.
- If a symbol has less history than requested years, the app returns all available data.
- No `tkinter`, no file dialogs — everything runs in the browser.