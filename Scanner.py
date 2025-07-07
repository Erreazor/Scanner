#!/usr/bin/env python3

import os
import time
from datetime import datetime
import pandas as pd
import yfinance as yf

# ─── CONFIG ─────────────────────────────────────────────────────────────────────
CSV_FILE       = os.path.expanduser("~/Desktop/SP500_updated.csv")
OUTPUT_DIR     = os.path.expanduser("~/Desktop")
MIN_MARKET_CAP = 500e6      # $500 million
MAX_TO_HIGH    = 0.05       # within 5%
MIN_AVG_VOLUME = 1e6        # 1 million shares/day

# ─── FETCH ONE TICKER ────────────────────────────────────────────────────────────
def fetch_info(ticker):
    try:
        tk   = yf.Ticker(ticker)
        info = tk.info
        mc   = info.get("marketCap", 0)
        adv  = info.get("averageVolume", 0)

        hist_full = tk.history(period="max")
        ath       = hist_full["High"].max() if not hist_full.empty else None

        hist52    = tk.history(period="52wk")
        w52       = hist52["High"].max() if not hist52.empty else None

        today_hist = tk.history(period="1d")
        curr       = today_hist["Close"].iloc[-1] if not today_hist.empty else None

        if None in (ath, w52, curr):
            return None

        return {
            "Ticker":      ticker,
            "Current":     curr,
            "AllTimeHigh": ath,
            "PctToATH":    (ath - curr) / ath,
            "52wHigh":     w52,
            "PctTo52w":    (w52  - curr) / w52,
            "MarketCap":   mc,
            "AvgVolume":   adv,
        }
    except Exception as e:
        print(f"⚠️  Error fetching {ticker}: {e}")
        return None

# ─── SCAN & SAVE WITH FORMATTING ────────────────────────────────────────────────
def scan_and_save(limit=None, sleep_interval=0.3):
    df_sp = pd.read_csv(CSV_FILE)
    # auto-detect ticker column
    possible = ["Ticker","ticker","Symbol","symbol","Code","code"]
    ticker_col = next((c for c in df_sp.columns if c in possible), None)
    if not ticker_col:
        print("❌ No ticker column found. Headers are:", df_sp.columns.tolist())
        return

    tickers = df_sp[ticker_col].dropna().astype(str).tolist()
    total   = len(tickers) if limit is None else min(limit, len(tickers))
    print(f"🔍 Scanning {total} tickers…")

    results = []
    for i, t in enumerate(tickers, start=1):
        if limit and i > limit:
            break
        print(f"[{i:03d}/{total:03d}] {t}", end="\r", flush=True)
        data = fetch_info(t.strip().upper())
        time.sleep(sleep_interval)
        if not data:
            continue
        if (data["MarketCap"] >= MIN_MARKET_CAP and
            data["AvgVolume"] >= MIN_AVG_VOLUME and
            (data["PctToATH"] <= MAX_TO_HIGH or data["PctTo52w"] <= MAX_TO_HIGH)):
            results.append(data)

    print(" " * 80, end="\r")  # clear progress line

    today = datetime.now().strftime("%Y%m%d")
    if not results:
        print("ℹ️ No matches today.")
        return

    df_out = pd.DataFrame(results)[[
        "Ticker","Current","AllTimeHigh","PctToATH",
        "52wHigh","PctTo52w","MarketCap","AvgVolume"
    ]].round(2)

    # write to Excel with conditional formatting
    outfile = os.path.join(OUTPUT_DIR, f"scan_results_{today}.xlsx")
    with pd.ExcelWriter(outfile, engine="xlsxwriter") as writer:
        df_out.to_excel(writer, index=False, sheet_name="Scan")
        wb  = writer.book
        ws  = writer.sheets["Scan"]

        # yellow fill format
        fmt_yellow = wb.add_format({"bg_color": "#FFFF00"})

        # apply yellow highlight if within 5% of all‐time high (PctToATH column is D)
        start_row = 2
        end_row   = len(df_out) + 1
        cell_range = f"A{start_row}:H{end_row}"

        ws.conditional_format(
            cell_range,
            {
                "type":     "formula",
                "criteria": f"=$D{start_row}<={MAX_TO_HIGH}",
                "format":   fmt_yellow
            }
        )

    print(f"✅ Wrote {len(df_out)} hits → {outfile}")

# ─── MAIN ───────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("➡️ Running scan now …")
    scan_and_save()   # remove `limit=` if you previously used it for testing
