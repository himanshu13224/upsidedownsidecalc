"""
fetch_and_portfolio.py
Combined workflow:
  1. Fetch daily/monthly close prices for a stock from tej-bazaar (NSE/BSE Bhavcopy)
  2. Automatically run the Portfolio Returns & Risk analysis (returns_tool.py)
  3. Save the final portfolio Excel workbook (.xlsx) directly

Output: Column 1 = Date, Column 2 = Close Price, then Returns + Risk Analysis sheets.
"""

import os
import sys
import time
import tempfile
import tkinter as tk
from tkinter import filedialog, messagebox

# Make sibling modules importable
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import fetch_prices as fp
import returns_tool as rt


def main():
    print("=" * 64)
    print("        FETCH PRICES + PORTFOLIO RETURNS")
    print("=" * 64)

    symbol = input("\nEnter stock symbol (e.g. tcs): ").strip()
    while not symbol:
        symbol = input("Symbol cannot be empty. Enter stock symbol: ").strip()

    ex = input("Exchange [NSE/BSE] (default NSE): ").strip().upper()
    while ex not in ("", "NSE", "BSE"):
        ex = input("Enter NSE or BSE: ").strip().upper()
    exchange = ex if ex else "NSE"

    years_str = input("How many years of data? [1/2/3/5/10 or any number]: ").strip()
    while not years_str.isdigit() or int(years_str) < 1:
        years_str = input("Please enter a positive number of years: ").strip()
    years = int(years_str)

    freq = input("Close price frequency? [1=Daily / 2=Monthly]: ").strip()
    while freq not in ("1", "2"):
        freq = input("Enter 1 for Daily or 2 for Monthly: ").strip()
    freq = "Monthly" if freq == "2" else "Daily"

    print(f"\nFetching {symbol} on {exchange}, last {years} year(s), {freq} close...")

    root = tk.Tk()
    root.withdraw()

    default_name = f"{symbol.lower()}_{exchange.lower()}_portfolio_{years}y_{freq.lower()}.xlsx"
    out = filedialog.asksaveasfilename(
        title="Save portfolio Excel as",
        defaultextension=".xlsx",
        initialfile=default_name,
        filetypes=[("Excel file", "*.xlsx")])
    if not out:
        print("No output path chosen. Exiting.")
        time.sleep(1.5)
        return

    # --- Step 1: fetch prices to a temporary Excel file ---
    tmp_dir = tempfile.gettempdir()
    tmp_file = os.path.join(
        tmp_dir,
        f"{symbol.lower()}_{exchange.lower()}_fetch_{years}y_{freq.lower()}.xlsx")

    try:
        fp.build_workbook(symbol, exchange, years, freq, tmp_file)
        print(f"\nPrices fetched to temp file: {tmp_file}")

        # --- Step 2: run portfolio returns analysis on the fetched file ---
        print("Running portfolio returns & risk analysis...")
        # mode "2" = Returns + downside/upside risk analysis
        rt.build_output(tmp_file, "2", out)

        print("\nDone!")
        print(f"Portfolio file saved: {out}")
        try:
            os.startfile(out)
        except Exception:
            pass

        # Clean up temp fetch file
        try:
            os.remove(tmp_file)
        except Exception:
            pass

    except Exception as e:
        messagebox.showerror("Error", f"Something went wrong:\n{e}")
        print(f"\nError: {e}")
        input("Press Enter to exit...")
        return

    input("Press any key to exit...")


if __name__ == "__main__":
    main()