#!/usr/bin/env python3

import os
import time
import random
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd
import yfinance as yf

import gspread
from google.oauth2.service_account import Credentials
from gspread.exceptions import WorksheetNotFound, APIError

# ─── CONFIG ─────────────────────────────────────────────────────────────────────
CSV_FILE              = os.path.expanduser("~/Desktop/SP500_updated.csv")
OUTPUT_DIR            = os.path.expanduser("~/Desktop")

# Google Sheets settings
SERVICE_ACCOUNT_FILE  = os.path.expanduser("~/Desktop/google-creds/scanner-465501-6c29eeebf51b.json")  # ← update this
SPREADSHEET_ID        = "1IdGxKSh0_6HCoIzSYHHYrGDdfGiixtDA19zcVxqyZLA"                           # ← your Sheet ID
WORKSHEET_NAME        = "Scanner"                                                             # ← your tab name

# Default thresholds
DEFAULT_MIN_MARKET_CAP = 500e6      # $500 million
DEFAULT_MIN_AVG_VOLUME = 1e6        # 1 million shares/day
DEFAULT_MAX_PCT        = 0.05       # within 5%

MIN_MARKET_CAP = DEFAULT_MIN_MARKET_CAP
MIN_AVG_VOLUME = DEFAULT_MIN_AVG_VOLUME
MAX_PCT        = DEFAULT_MAX_PCT

MAX_WORKERS    = 8
SLEEP_INTERVAL = 1.0  # seconds


def send_to_sheet(df: pd.DataFrame):
    """Push DataFrame to Google Sheet and apply formatting & conditional rules."""
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds  = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=scopes)
    gc     = gspread.authorize(creds)

    sh = gc.open_by_key(SPREADSHEET_ID)
    try:
        ws = sh.worksheet(WORKSHEET_NAME)
    except WorksheetNotFound:
        ws = sh.add_worksheet(
            title=WORKSHEET_NAME,
            rows=str(len(df) + 1),
            cols=str(len(df.columns))
        )
        print(f"➕ Created worksheet '{WORKSHEET_NAME}'")

    # Write data
    ws.clear()
    rows = [df.columns.tolist()] + df.values.tolist()
    ws.update(rows, value_input_option="USER_ENTERED")
    print(f"📝 Pushed {len(df)} rows to Google Sheet '{WORKSHEET_NAME}'")

    # Formatting parameters
    sheet_id = ws.id
    num_rows = len(df) + 1
    num_cols = len(df.columns)
    pct_idx  = df.columns.get_loc("PctToATH")

    # 1) Apply banded range (if not exists)
    banding_req = {
        "addBanding": {
            "bandedRange": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 0,
                    "endRowIndex": num_rows,
                    "startColumnIndex": 0,
                    "endColumnIndex": num_cols
                },
                "rowProperties": {
                    "headerColor":     {"red": 0.8, "green": 0.8, "blue": 0.8},
                    "firstBandColor":  {"red": 1.0, "green": 1.0, "blue": 1.0},
                    "secondBandColor": {"red": 0.95,"green": 0.95,"blue": 0.95}
                }
            }
        }
    }
    try:
        ws.spreadsheet.batch_update({"requests": [banding_req]})
        print("🎨 Applied banded table style")
    except APIError as e:
        if "already has alternating background colors" in e.response.text:
            print("ℹ️ Banding already exists—skipping")
        else:
            raise

    # 2) Other formatting & conditional rules
    other_reqs = [
        # freeze header row
        {
            "updateSheetProperties": {
                "properties": {
                    "sheetId": sheet_id,
                    "gridProperties": {"frozenRowCount": 1}
                },
                "fields": "gridProperties.frozenRowCount"
            }
        },
        # highlight rows near ATH
        {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [{
                        "sheetId": sheet_id,
                        "startRowIndex": 1,
                        "endRowIndex": num_rows,
                        "startColumnIndex": 0,
                        "endColumnIndex": num_cols
                    }],
                    "booleanRule": {
                        "condition": {
                            "type": "CUSTOM_FORMULA",
                            "values": [{
                                "userEnteredValue":
                                    f"=${chr(ord('A') + pct_idx)}2<={MAX_PCT}"
                            }]
                        },
                        "format": {
                            "backgroundColor": {"red": 1.0, "green": 1.0, "blue": 0.0}
                        }
                    }
                },
                "index": 0
            }
        }
    ]
    ws.spreadsheet.batch_update({"requests": other_reqs})
    print("🎨 Applied header freeze & highlights")


def fetch_info(ticker: str):
    """Fetch metrics with yfinance, including sector."""
    try:
        tk   = yf.Ticker(ticker)
        info = tk.info
        sector = info.get("sector", "N/A")
        mc   = info.get("marketCap", 0)
        adv  = info.get("averageVolume", 0)

        hist = tk.history(period="max")
        if hist.empty:
            return None

        ath = hist["High"].max()
        cutoff = (pd.Timestamp.now(tz=hist.index.tz) if hist.index.tz
                  else pd.Timestamp.now()) - pd.Timedelta(days=365)
        w52    = hist.loc[hist.index >= cutoff, "High"].max()
        curr   = hist["Close"].iloc[-1]
        if None in (ath, w52, curr):
            return None

        return {
            "Ticker":      ticker,
            "Sector":      sector,
            "Current":     curr,
            "AllTimeHigh": ath,
            "PctToATH":    (ath - curr) / ath,
            "52wHigh":     w52,
            "PctTo52w":    (w52  - curr) / w52,
            "MarketCap":   mc,
            "AvgVolume":   adv
        }
    except Exception as e:
        print(f"⚠️ {ticker} fetch failed: {e}")
        return None


def scan_and_save():
    df_sp = pd.read_csv(CSV_FILE)
    ticker_col = next((c for c in df_sp.columns if c.lower() in ("ticker","symbol","code")), None)
    if not ticker_col:
        print("❌ No ticker column found. Headers:", df_sp.columns.tolist())
        return

    tickers = df_sp[ticker_col].dropna().astype(str).tolist()
    total   = len(tickers)
    print(f"🔍 Scanning {total} tickers…")

    results = []
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {executor.submit(fetch_info, t): t for t in tickers}
        for i, fut in enumerate(as_completed(futures), 1):
            t = futures[fut]
            data = fut.result()
            time.sleep(SLEEP_INTERVAL + random.random())
            print(f"[{i}/{total}] {t}", end="\r", flush=True)
            if not isinstance(data, dict):
                continue
            if (
                data["MarketCap"]  >= MIN_MARKET_CAP and
                data["AvgVolume"]  >= MIN_AVG_VOLUME and
                (data["PctToATH"]  <= MAX_PCT or data["PctTo52w"] <= MAX_PCT)
            ):
                results.append(data)

    print(" " * 80, end="\r")
    today = datetime.now().strftime("%Y%m%d")
    if not results:
        print("ℹ️ No matches today.")
        return

    df_out = pd.DataFrame(results)[
        ["Ticker","Sector","Current","AllTimeHigh","PctToATH","52wHigh","PctTo52w","MarketCap","AvgVolume"]
    ].round(2)

    csv_file = os.path.join(OUTPUT_DIR, f"scan_results_{today}.csv")
    df_out.to_csv(csv_file, index=False)
    print(f"✅ Wrote {len(df_out)} hits → {csv_file}")

    send_to_sheet(df_out)


if __name__ == "__main__":
    inp = input("Market cap (preset >500M, type y): ")
    if inp.lower() != 'y' and inp.strip():
        try: MIN_MARKET_CAP = float(inp)
        except: print("Invalid—using default.")
    inp = input("Avg volume (preset >1M, type y): ")
    if inp.lower() != 'y' and inp.strip():
        try: MIN_AVG_VOLUME = float(inp)
        except: print("Invalid—using default.")
    inp = input("% from high (preset 0.05, type y): ")
    if inp.lower() != 'y' and inp.strip():
        try: MAX_PCT = float(inp)
        except: print("Invalid—using default.")

    print("➡️ Running scan …")
    scan_and_save()
