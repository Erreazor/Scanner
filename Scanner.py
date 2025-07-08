#!/usr/bin/env python3

import os
import time
from datetime import datetime
import pandas as pd
import yfinance as yf
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# ─── CONFIG ─────────────────────────────────────────────────────────────────────
CSV_FILE               = os.path.expanduser("~/Desktop/SP500_updated.csv")
OUTPUT_DIR             = os.path.expanduser("~/Desktop")
# Default thresholds
DEFAULT_MIN_MARKET_CAP = 500e6      # $500 million
DEFAULT_MIN_AVG_VOLUME = 1e6        # 1 million shares/day
DEFAULT_MAX_PCT        = 0.05       # within 5% of high

# Email settings (configure these!)
EMAIL_SENDER     = "trlm.scanner.bot@outlook.com"
EMAIL_PASSWORD   = "Scannerbot2025"
EMAIL_RECIPIENTS = ["erreazor@iu.edu"]
SMTP_SERVER      = "smtp.mail.outlook.com"
SMTP_PORT        = 587

# Global thresholds (will be set by user input)
MIN_MARKET_CAP = DEFAULT_MIN_MARKET_CAP
MIN_AVG_VOLUME = DEFAULT_MIN_AVG_VOLUME
MAX_PCT        = DEFAULT_MAX_PCT

# ─── EMAIL FUNCTION ─────────────────────────────────────────────────────────────
def send_email(subject, body):
    """Attempt to send an email, catch and log any errors."""
    try:
        msg = MIMEMultipart()
        msg['From']    = EMAIL_SENDER
        msg['To']      = ", ".join(EMAIL_RECIPIENTS)
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))
        # add timeout to SMTP connect
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=30) as server:
            server.starttls()
            server.login(EMAIL_SENDER, EMAIL_PASSWORD)
            server.send_message(msg)
        print(f"📧 Email sent to {', '.join(EMAIL_RECIPIENTS)}")
    except Exception as e:
        print(f"⚠️ Email failed: {e}")

# ─── FETCH INFO ─────────────────────────────────────────────────────────────────
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
    except Exception:
        return None

# ─── SCAN, SAVE, FORMAT, EMAIL ─────────────────────────────────────────────────
def scan_and_save():
    # Load tickers
    df_sp = pd.read_csv(CSV_FILE)
    possible = ["Ticker","ticker","Symbol","symbol","Code","code"]
    ticker_col = next((c for c in df_sp.columns if c in possible), None)
    if not ticker_col:
        print("❌ No ticker column found. Headers are:", df_sp.columns.tolist())
        return

    tickers = df_sp[ticker_col].dropna().astype(str).tolist()
    total   = len(tickers)

    print(f"🔍 Scanning {total} tickers with presets:\n"
          f"   • Min Market Cap: ${MIN_MARKET_CAP:,.0f}\n"
          f"   • Min Avg Volume: {MIN_AVG_VOLUME:,.0f} shares/day\n"
          f"   • Within {MAX_PCT*100:.2f}% of ATH or 52w High")

    results = []
    for i, t in enumerate(tickers, start=1):
        print(f"[{i:>3}/{total}] {t}", end="\r", flush=True)
        data = fetch_info(t.strip().upper())
        time.sleep(0.3)
        if not data:
            continue
        if (data["MarketCap"] >= MIN_MARKET_CAP and
            data["AvgVolume"] >= MIN_AVG_VOLUME and
            (data["PctToATH"] <= MAX_PCT or data["PctTo52w"] <= MAX_PCT)):
            results.append(data)

    print(" " * 80, end="\r")  # clear progress line

    today = datetime.now().strftime("%Y%m%d")
    if not results:
        print("ℹ️ No matches today.")
        send_email(f"Breakout Scan {today}", "No matches today.")
        return

    # Build DataFrame
    df_out = pd.DataFrame(results)[[
        "Ticker","Current","AllTimeHigh","PctToATH",
        "52wHigh","PctTo52w","MarketCap","AvgVolume"
    ]].round(2)

    # 1) Save CSV
    csv_file = os.path.join(OUTPUT_DIR, f"scan_results_{today}.csv")
    df_out.to_csv(csv_file, index=False)
    print(f"✅ Wrote {len(df_out)} hits → {csv_file}")

    # 2) Save Excel with highlight for All-Time High proximity
    xlsx_file = os.path.join(OUTPUT_DIR, f"scan_results_{today}.xlsx")
    with pd.ExcelWriter(xlsx_file, engine="xlsxwriter") as writer:
        df_out.to_excel(writer, index=False, sheet_name="Scan")
        wb  = writer.book
        ws  = writer.sheets["Scan"]
        fmt_yellow = wb.add_format({"bg_color": "#FFFF00"})

        start_row = 2
        end_row   = len(df_out) + 1
        cell_range = f"A{start_row}:H{end_row}"
        ws.conditional_format(cell_range, {
            "type":     "formula",
            "criteria": f"=$D{start_row}<={MAX_PCT}",
            "format":   fmt_yellow
        })
    print(f"✅ Wrote Excel with highlights → {xlsx_file}")

    # 3) Email results
    subject = f"Breakout Scan Results for {today}"
    body    = df_out.to_string(index=False)
    send_email(subject, body)

# ─── MAIN ───────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    inp = input(f"Market cap (want to use preset >500M type y): ")
    if inp.lower() != 'y' and inp.strip() != '':
        try:
            MIN_MARKET_CAP = float(inp)
        except ValueError:
            print("Invalid input for market cap, using preset.")

    inp = input(f"Avg volume (want to use preset >1M type y): ")
    if inp.lower() != 'y' and inp.strip() != '':
        try:
            MIN_AVG_VOLUME = float(inp)
        except ValueError:
            print("Invalid input for avg volume, using preset.")

    inp = input(f"% from high (want to use preset 0.05 type y): ")
    if inp.lower() != 'y' and inp.strip() != '':
        try:
            MAX_PCT = float(inp)
        except ValueError:
            print("Invalid input for % from high, using preset.")

    print("➡️ Running scan now …")
    scan_and_save()
