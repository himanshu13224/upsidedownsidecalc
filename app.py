"""
streamlit_app.py
Stock Price Fetcher + Portfolio Analysis (Streamlit web app)
Fetches NSE/BSE daily/monthly close prices from tej-bazaar (official Bhavcopy via HuggingFace)
and builds an Excel workbook with Returns & Risk Analysis.
No API key required.
"""

import streamlit as st
import pandas as pd
import requests
import openpyxl
from openpyxl.styles import Font
from datetime import date, timedelta
import io

st.set_page_config(
    page_title="Stock Price Fetcher",
    page_icon="📈",
    layout="centered"
)

BOLD = Font(bold=True)
HF_BASE = "https://huggingface.co/datasets/tejhq/indian-markets/resolve/main"


def download_year(exchange, year):
    """Download one year's parquet from tej-bazaar HF dataset. Returns bytes or None."""
    fn = f"{exchange.lower()}_{year}.parquet"
    url = f"{HF_BASE}/{exchange.lower()}/year={year}/{fn}"
    try:
        r = requests.get(url, timeout=60)
        if r.status_code == 200:
            return r.content
        elif r.status_code == 404:
            return None
        else:
            st.warning(f"HTTP {r.status_code} for {exchange} {year}")
            return None
    except Exception as e:
        st.error(f"Error downloading {exchange} {year}: {e}")
        return None


def fetch_prices(symbol, exchange, years, freq="Daily"):
    """Fetch close prices for symbol from tej-bazaar, going back `years` years.
    freq: "Daily" or "Monthly" (last close of each month).
    """
    symbol = symbol.strip().upper()
    today = date.today()
    need_years = list(range(today.year, today.year - years - 1, -1))

    frames = []
    progress = st.progress(0)
    total_years = len(need_years)

    for i, y in enumerate(need_years):
        progress.progress((i + 1) / total_years, text=f"Downloading {exchange} {y} ...")
        content = download_year(exchange, y)
        if content is None:
            continue
        try:
            df = pd.read_parquet(io.BytesIO(content))
            df_sym = df[(df["symbol"] == symbol) & (df["series"].isin(["EQ", "BE", "BZ"]))]
            if not df_sym.empty:
                frames.append(df_sym[["date", "close"]].copy())
        except Exception as e:
            st.warning(f"Read error for {y}: {e}")

    progress.empty()

    if not frames:
        raise ValueError(f"No data found for {symbol} on {exchange}.")

    all_df = pd.concat(frames, ignore_index=True)
    all_df = all_df.sort_values("date").drop_duplicates(subset="date", keep="last")

    cutoff = today - timedelta(days=365 * years + 30)
    all_df = all_df[all_df["date"] >= cutoff]

    if freq == "Monthly":
        # last close of each month using pure Python (avoids pd.to_datetime crash on Py3.14)
        monthly = {}
        for _, row in all_df.iterrows():
            d = row["date"]
            key = (d.year, d.month)
            monthly[key] = (d, float(row["close"]))
        rows = [{"date": v[0], "close": v[1]} for v in monthly.values()]
        rows.sort(key=lambda x: x["date"])
        all_df = pd.DataFrame(rows)
    else:
        all_df = all_df[["date", "close"]].copy()

    return all_df


def build_workbook(symbol, exchange, years, freq, df, include_analysis=True):
    """Build an in-memory Excel workbook (.xlsx) and return BytesIO buffer."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"{symbol} {exchange} {years}Y {freq}"

    ws.cell(row=1, column=1, value="Date").font = BOLD
    ws.cell(row=1, column=2, value="Close Price").font = BOLD
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 14

    r = 2
    for _, row in df.iterrows():
        ws.cell(row=r, column=1, value=row["date"]).number_format = "DD-MM-YYYY"
        ws.cell(row=r, column=2, value=float(row["close"])).number_format = "0.00"
        r += 1

    if include_analysis and len(df) >= 2:
        # ----- Sheet: Returns -----
        ws2 = wb.create_sheet("Returns")
        ws2.cell(row=1, column=1, value="Date").font = BOLD
        ws2.cell(row=1, column=2, value="Close Price").font = BOLD
        ws2.cell(row=1, column=3, value="Returns").font = BOLD
        ws2.column_dimensions["A"].width = 14
        ws2.column_dimensions["B"].width = 14
        ws2.column_dimensions["C"].width = 14

        for i, row in df.iterrows():
            rr = i + 2
            ws2.cell(row=rr, column=1, value=row["date"]).number_format = "DD-MM-YYYY"
            ws2.cell(row=rr, column=2, value=float(row["close"])).number_format = "0.00"
            if rr > 2:
                ws2.cell(row=rr, column=3, value=f"=(B{rr}-B{rr-1})/B{rr-1}").number_format = "0.000000"

        last_ws2 = len(df) + 1  # last row with data in Returns sheet

        # ----- Sheet: Risk Analysis -----
        ws3 = wb.create_sheet("Risk Analysis")
        ret_range = f"'Returns'!C3:C{last_ws2}"
        headers = ["Returns", "expected return", "difference",
                   "negative returns", "Positive", "Square NR", "Square PR"]
        for c, h in enumerate(headers, start=1):
            ws3.cell(row=1, column=c, value=h).font = BOLD

        # expected return lives in J1 (referenced by every row as $J$1)
        ws3.cell(row=1, column=10, value="expected return")
        ws3.cell(row=1, column=11, value=f"=AVERAGE({ret_range})").number_format = "0.000000"

        rows_count = last_ws2 - 2  # number of return observations (rows 3..last)
        for i in range(rows_count):
            rr = i + 2
            ws3.cell(row=rr, column=1, value=f"='Returns'!C{i + 3}").number_format = "0.000000"
            ws3.cell(row=rr, column=2, value="=$J$1").number_format = "0.000000"
            ws3.cell(row=rr, column=3, value=f"=A{rr}-B{rr}").number_format = "0.000000"
            ws3.cell(row=rr, column=4, value=f"=IF(C{rr}<0,C{rr},0)").number_format = "0.000000"
            ws3.cell(row=rr, column=5, value=f"=IF(C{rr}>0,C{rr},0)").number_format = "0.000000"
            ws3.cell(row=rr, column=6, value=f"=D{rr}^2").number_format = "0.000000000000"
            ws3.cell(row=rr, column=7, value=f"=E{rr}^2").number_format = "0.000000000000"

        rng_nr = f"F2:F{rows_count + 1}"
        rng_pr = f"G2:G{rows_count + 1}"

        summary = [
            ("average", f"=AVERAGE({rng_nr})", "0.000000000000"),
            ("downside risk", f"=SQRT(AVERAGE({rng_nr}))", "0.000000"),
            ("average", f"=AVERAGE({rng_pr})", "0.000000000000"),
            ("upside", f"=SQRT(AVERAGE({rng_pr}))", "0.000000"),
        ]
        r0 = 6
        for i, (name, f, fmt) in enumerate(summary):
            ws3.cell(row=r0 + i * 3, column=10, value=name).font = BOLD
            ws3.cell(row=r0 + i * 3, column=11, value=f).number_format = fmt

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


# ---------------- UI ----------------
st.title("📈 Stock Price Fetcher & Portfolio Analysis")
st.markdown(
    "Fetch NSE/BSE daily/monthly close prices and build an Excel workbook with "
    "Returns & Risk Analysis. Data from official exchange Bhavcopy (tej-bazaar). "
    "No API key required."
)

with st.form("fetch_form"):
    col1, col2 = st.columns(2)
    with col1:
        symbol = st.text_input(
            "Stock Symbol",
            value="tcs",
            help="e.g. tcs, suzlon, reliance, hdfcbank"
        )
    with col2:
        exchange = st.selectbox("Exchange", ["NSE", "BSE"], index=0)

    col3, col4 = st.columns(2)
    with col3:
        years = st.number_input(
            "Years of data",
            min_value=1,
            max_value=20,
            value=1,
            step=1
        )
    with col4:
        freq = st.radio("Frequency", ["Daily", "Monthly"], horizontal=True)

    include_analysis = st.checkbox(
        "Include Returns & Risk Analysis",
        value=True,
        help="Adds Returns sheet with live formulas and Risk Analysis sheet"
    )

    submitted = st.form_submit_button("Fetch & Build Excel", type="primary")

if submitted:
    if not symbol.strip():
        st.error("Please enter a stock symbol.")
    else:
        try:
            with st.spinner("Fetching prices from exchange Bhavcopy..."):
                df = fetch_prices(symbol, exchange, int(years), freq)

            st.success(f"Fetched {len(df)} rows for {symbol.upper()} on {exchange}.")
            st.dataframe(df.head(10), use_container_width=True)
            if len(df) > 10:
                st.caption(f"Showing first 10 of {len(df)} rows")

            with st.spinner("Building Excel workbook..."):
                buffer = build_workbook(
                    symbol, exchange, int(years), freq, df, include_analysis
                )

            file_name = f"{symbol.lower()}_{exchange.lower()}_{years}y_{freq.lower()}.xlsx"
            st.download_button(
                label="📥 Download Excel Workbook",
                data=buffer,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        except Exception as e:
            st.error(f"Error: {e}")

st.markdown("---")
st.caption(
    "Data source: tej-bazaar / HuggingFace (official NSE & BSE Bhavcopy). "
    "Free, no API key required."
)