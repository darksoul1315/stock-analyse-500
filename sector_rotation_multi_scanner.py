#!/usr/bin/env python3
"""
Nifty 500 Delivery, Sentiment & Trend Scanner
Compiles multi-timeframe moving averages, outperformance ratios (RS), multi-year high/low price bounds,
NSE daily/weekly deliverable volume percentages, and real-time Google News sentiment polarity scores.
Outputs the styled 5-sheet report: 'sector_rotation_multi_report.xlsx'.

CHANGES (Smart Cache + Incremental Update + Force Refresh):
- Added --force-refresh CLI flag
- Pehli baar: 10 saal ka pura data download + cache save
- Agli din: sirf naye din ka data fetch, purane cache mein merge
- Same din 2nd run: sirf cache load, kuch download nahi
- Old cache files auto-cleaned on each run
"""

import os
import sys
import io
import re
import glob
import time
import argparse
import datetime
import urllib.request
import urllib.parse
import xml.etree.ElementTree as ET
import numpy as np
import pandas as pd
import yfinance as yf
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Ensure local module imports work
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# Override SSL context verification for macOS environment compatibility
import ssl
try:
    ssl._create_default_https_context = ssl._create_unverified_context
    print("🔒 SSL certificate verification bypassed for external API downloads.")
except AttributeError:
    pass

# Define output
OUTPUT_FILE = "sector_rotation_multi_report.xlsx"

# Cache file pattern
CACHE_PREFIX = "price_cache_"
CACHE_FILE   = f"{CACHE_PREFIX}{datetime.date.today()}.parquet"


# ─────────────────────────────────────────────────────────────
# CLI ARGUMENT PARSER
# ─────────────────────────────────────────────────────────────
def parse_args():
    """
    Parses command-line arguments.

    Usage:
        python sector_rotation_multi_scanner.py               # Normal run (uses cache if available)
        python sector_rotation_multi_scanner.py --force-refresh  # Ignores cache, re-downloads fresh data
    """
    parser = argparse.ArgumentParser(
        description="Nifty 500 Multi-Timeframe Sector Rotation Scanner"
    )
    parser.add_argument(
        "--force-refresh",
        action="store_true",
        default=False,
        help="Ignore today's cached price data and re-download everything from Yahoo Finance."
    )
    return parser.parse_args()


# ─────────────────────────────────────────────────────────────
# CACHE UTILITIES  (SMART INCREMENTAL)
# ─────────────────────────────────────────────────────────────
def cleanup_old_caches():
    """
    Purani saari cache parquet files delete karta hai.
    Sirf sabse latest file rakhta hai (today's or most recent).
    """
    all_cache_files = sorted(glob.glob(f"{CACHE_PREFIX}*.parquet"))
    # Sabse latest file ko chodo, baaki sab delete karo
    files_to_delete = all_cache_files[:-1] if len(all_cache_files) > 1 else []
    for f in files_to_delete:
        try:
            os.remove(f)
            print(f"🗑️  Deleted old cache: {f}")
        except Exception as e:
            print(f"⚠️  Could not delete old cache {f}: {e}")


def save_price_cache(prices_raw: dict):
    """
    prices_raw dict (of DataFrames) ko ek single parquet file mein save karta hai.
    Columns MultiIndex format mein store hoti hain: (field, ticker)
    """
    try:
        combined = pd.concat(prices_raw, axis=1)  # columns → (field, ticker)
        combined.to_parquet(CACHE_FILE)
        print(f"✅ Price data cached to: {CACHE_FILE}")
    except Exception as e:
        print(f"⚠️  Warning: Could not save price cache: {e}")


def load_any_existing_cache() -> tuple:
    """
    Jo bhi cache file available hai (aaj ki ya purani) use load karta hai.
    Returns: (prices_raw dict, last_date in cache as datetime.date)
    """
    all_cache_files = sorted(glob.glob(f"{CACHE_PREFIX}*.parquet"))
    if not all_cache_files:
        return None, None

    latest_cache = all_cache_files[-1]
    print(f"📦 Loading existing cache: {latest_cache}")
    try:
        combined   = pd.read_parquet(latest_cache)
        prices_raw = {}
        for field in ["Close", "Open", "High", "Low", "Volume"]:
            if field in combined.columns.get_level_values(0):
                prices_raw[field] = combined[field]
            else:
                prices_raw[field] = pd.DataFrame()

        # Cache mein last available trading date nikalo
        last_date = prices_raw['Close'].index[-1].date()
        print(f"📅 Cache last date: {last_date}")
        return prices_raw, last_date
    except Exception as e:
        print(f"⚠️  Cache load failed: {e}")
        return None, None


def fetch_incremental_data(tickers: list, from_date: datetime.date, chunk_size: int = 100) -> dict:
    """
    Sirf from_date se aaj tak ka naya data download karta hai.
    Yahi function next-day run mein call hoga — full 10Y nahi, sirf missing days.
    """
    start_str = from_date.strftime("%Y-%m-%d")
    end_str   = (datetime.date.today() + datetime.timedelta(days=1)).strftime("%Y-%m-%d")

    print(f"📥 Incremental download: {start_str} → {end_str} ({len(tickers)} tickers)...")

    closes_list = []; opens_list = []; highs_list = []; lows_list = []; volumes_list = []

    for i in range(0, len(tickers), chunk_size):
        chunk = tickers[i:i + chunk_size]
        for attempt in range(2):
            try:
                chunk_raw = yf.download(chunk, start=start_str, end=end_str, progress=False)
                if not chunk_raw.empty:
                    if isinstance(chunk_raw.columns, pd.MultiIndex):
                        for field, lst in [('Close', closes_list), ('Open', opens_list),
                                           ('High', highs_list),   ('Low', lows_list),
                                           ('Volume', volumes_list)]:
                            if field in chunk_raw:
                                lst.append(chunk_raw[field])
                    else:
                        ticker = chunk[0]
                        closes_list.append(pd.DataFrame({ticker: chunk_raw['Close']}))
                        opens_list.append(pd.DataFrame({ticker: chunk_raw['Open']}))
                        highs_list.append(pd.DataFrame({ticker: chunk_raw['High']}))
                        lows_list.append(pd.DataFrame({ticker: chunk_raw['Low']}))
                        volumes_list.append(pd.DataFrame({ticker: chunk_raw['Volume']}))
                time.sleep(0.3)
                break
            except Exception as e:
                if attempt == 0:
                    print(f"  ⚠️  Batch error: {e}. Retrying...")
                    time.sleep(2.0)
                else:
                    print(f"  ❌ Retry failed: {e}")

    return {
        'Close':  pd.concat(closes_list,  axis=1) if closes_list  else pd.DataFrame(),
        'Open':   pd.concat(opens_list,   axis=1) if opens_list   else pd.DataFrame(),
        'High':   pd.concat(highs_list,   axis=1) if highs_list   else pd.DataFrame(),
        'Low':    pd.concat(lows_list,    axis=1) if lows_list    else pd.DataFrame(),
        'Volume': pd.concat(volumes_list, axis=1) if volumes_list else pd.DataFrame(),
    }


def merge_incremental(old: dict, new: dict) -> dict:
    """
    Purane cache data mein naya data merge karta hai.
    Duplicate dates automatically overwrite ho jaati hain (last-wins).
    """
    merged = {}
    for field in ["Close", "Open", "High", "Low", "Volume"]:
        old_df = old.get(field, pd.DataFrame())
        new_df = new.get(field, pd.DataFrame())

        if old_df.empty:
            merged[field] = new_df
        elif new_df.empty:
            merged[field] = old_df
        else:
            combined = pd.concat([old_df, new_df], axis=0)
            # Duplicate index (same date) hone par last value rakhta hai
            combined = combined[~combined.index.duplicated(keep='last')]
            merged[field] = combined.sort_index()

    return merged


# ─────────────────────────────────────────────────────────────
# TELEGRAM BOT INTEGRATION MODULE
# ─────────────────────────────────────────────────────────────
def load_env_file():
    """Loads local .env variables into os.environ if it exists."""
    env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
    if os.path.exists(env_path):
        try:
            with open(env_path, "r") as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith("#") and "=" in line:
                        key, val = line.split("=", 1)
                        os.environ[key.strip()] = val.strip()
            print("🔑 Environment credentials loaded from .env file.")
        except Exception as e:
            print(f"⚠️ Warning: Failed to parse .env file: {e}")


def send_telegram_alerts(text_summary: str, file_path: str = None) -> bool:
    """
    Dispatches a formatted Markdown alert and optional file attachment
    directly to a Telegram chat/bot channel.
    """
    token   = os.environ.get("TELEGRAM_BOT_TOKEN")
    chat_id = os.environ.get("TELEGRAM_CHAT_ID")

    if not token or not chat_id:
        print("⚠️ Warning: Telegram credentials not found. Skipping alert.")
        return False

    print("📤 Dispatching Telegram notifications...")

    # 1. Send Text Summary Alert
    text_url = f"https://api.telegram.org/bot{token}/sendMessage"
    payload  = {"chat_id": chat_id, "text": text_summary, "parse_mode": "Markdown"}

    try:
        req_data = urllib.parse.urlencode(payload).encode("utf-8")
        req = urllib.request.Request(text_url, data=req_data, method="POST")
        with urllib.request.urlopen(req, timeout=12):
            print(" ✅ Text alert sent successfully.")
    except Exception as e:
        print(f" ❌ Failed to send Telegram text alert: {e}")
        return False

    # 2. Upload Excel Report file
    if file_path and os.path.exists(file_path):
        doc_url  = f"https://api.telegram.org/bot{token}/sendDocument"
        boundary = "----WebKitFormBoundary7MA4YWxkTrZu0gW"
        try:
            parts = []
            parts.append(f"--{boundary}\r\nContent-Disposition: form-data; name=\"chat_id\"\r\n\r\n{chat_id}\r\n")
            file_name = os.path.basename(file_path)
            parts.append(
                f"--{boundary}\r\nContent-Disposition: form-data; name=\"document\"; "
                f"filename=\"{file_name}\"\r\nContent-Type: "
                f"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n"
            )
            with open(file_path, "rb") as f:
                file_content = f.read()

            body = bytearray()
            for part in parts:
                body.extend(part.encode("utf-8"))
            body.extend(file_content)
            body.extend(f"\r\n--{boundary}--\r\n".encode("utf-8"))

            headers = {
                "Content-Type": f"multipart/form-data; boundary={boundary}",
                "Content-Length": str(len(body))
            }
            req = urllib.request.Request(doc_url, data=body, headers=headers, method="POST")
            with urllib.request.urlopen(req, timeout=25):
                print(" ✅ Excel report uploaded successfully to Telegram.")
            return True
        except Exception as e:
            print(f" ❌ Failed to upload Excel report to Telegram: {e}")
            return False

    return True


# ─────────────────────────────────────────────────────────────
# LEXICONS
# ─────────────────────────────────────────────────────────────
POSITIVE_WORDS = {
    'bullish', 'surge', 'jump', 'gain', 'positive', 'profit', 'growth', 'demand',
    'outperform', 'high', 'rise', 'record', 'strong', 'expansion', 'rally', 'upgrade',
    'buy', 'accumulate', 'double', 'success', 'win', 'momentum', 'revival', 'recover',
    'bull', 'improving', 'leadership', 'outperforming', 'breakout'
}

NEGATIVE_WORDS = {
    'bearish', 'drop', 'plunge', 'loss', 'negative', 'decline', 'compress', 'fall',
    'low', 'weak', 'warning', 'crash', 'slash', 'downgrade', 'sell', 'avoid', 'deficit',
    'compression', 'drag', 'penalty', 'delay', 'slump', 'pessimism', 'inflation',
    'distribution', 'bear', 'weakening', 'lagging', 'underperforming', 'whipsaw'
}


# ─────────────────────────────────────────────────────────────
# 1. CORE TECHNICAL INDICATORS
# ─────────────────────────────────────────────────────────────
def calculate_ema(series: pd.Series, period: int) -> pd.Series:
    return series.ewm(span=period, adjust=False).mean()


def calculate_rsi(series: pd.Series, period: int = 14) -> pd.Series:
    delta    = series.diff()
    gain     = delta.clip(lower=0)
    loss     = -delta.clip(upper=0)
    avg_gain = gain.ewm(alpha=1/period, adjust=False).mean()
    avg_loss = loss.ewm(alpha=1/period, adjust=False).mean()
    rs       = avg_gain / avg_loss.replace(0, np.nan)
    return (100 - (100 / (1 + rs))).fillna(50)


def calculate_rs_ratio(stock_series: pd.Series, compare_series: pd.Series, window: int) -> float:
    """Return spread: stock_return - benchmark_return over `window` bars."""
    if len(stock_series) < window or len(compare_series) < window:
        return 0.0
    stock_ret   = stock_series.iloc[-1] / stock_series.iloc[-window] - 1
    compare_ret = compare_series.iloc[-1] / compare_series.iloc[-window] - 1
    if compare_ret == -1.0:
        return 0.0
    return stock_ret - compare_ret


# ─────────────────────────────────────────────────────────────
# 1.1 RRG COORDINATES ENGINE
# ─────────────────────────────────────────────────────────────
def calculate_rrg_coordinates(sector_series: pd.Series, bench_series: pd.Series) -> tuple:
    aligned_df = pd.DataFrame({'Sector': sector_series, 'Bench': bench_series}).dropna()
    if len(aligned_df) < 120:
        return pd.Series(100.0, index=sector_series.index), pd.Series(100.0, index=sector_series.index)

    sec   = aligned_df['Sector']
    bench = aligned_df['Bench']
    rs    = (sec / bench) * 100

    ema_rs  = rs.ewm(span=60, adjust=False).mean()
    std_rs  = rs.rolling(window=60, min_periods=30).std()
    rs_ratio = 100 + 10 * ((rs - ema_rs) / std_rs.replace(0, np.nan)).fillna(0)

    d_ratio    = rs_ratio.diff().fillna(0)
    ema_d_fast = d_ratio.ewm(span=10, adjust=False).mean()
    ema_d_slow = d_ratio.ewm(span=60, adjust=False).mean()
    std_d      = d_ratio.rolling(window=60, min_periods=30).std()
    rs_momentum = 100 + 10 * ((ema_d_fast - ema_d_slow) / std_d.replace(0, np.nan)).fillna(0)

    return (
        rs_ratio.reindex(sector_series.index).fillna(100.0),
        rs_momentum.reindex(sector_series.index).fillna(100.0)
    )


def calculate_velocity_and_heading(ratio, momentum, prev_ratio, prev_momentum) -> tuple:
    dx       = ratio - prev_ratio
    dy       = momentum - prev_momentum
    velocity = np.sqrt(dx**2 + dy**2)
    heading  = (np.degrees(np.arctan2(dy, dx)) + 360) % 360
    return round(float(velocity), 2), round(float(heading), 1)


# ─────────────────────────────────────────────────────────────
# 2. NSE DATA DOWNLOADERS
# ─────────────────────────────────────────────────────────────
def fetch_nifty500_constituents() -> pd.DataFrame:
    url = "https://archives.nseindia.com/content/indices/ind_nifty500list.csv"
    print(f"📥 Fetching Nifty 500 list from NSE: {url}")
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=10) as response:
            df = pd.read_csv(io.StringIO(response.read().decode('utf-8')))
            df.columns = df.columns.str.strip()
            return df
    except Exception as e:
        print(f"⚠️ Warning: Failed to fetch online Nifty 500 list ({e}). Using fallback.")
        return get_fallback_nifty500()


def download_last_5_days_delivery_bhavcopies() -> list:
    bhavcopies   = []
    date_to_check = datetime.date.today()
    attempts     = 0
    print("📥 Scanning NSE archives for the last 5 active Delivery Bhavcopy reports...")

    while len(bhavcopies) < 5 and attempts < 15:
        if date_to_check.weekday() < 5:
            date_str = date_to_check.strftime("%d%m%Y")
            url = f"https://nsearchives.nseindia.com/products/content/sec_bhavdata_full_{date_str}.csv"
            try:
                req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
                with urllib.request.urlopen(req, timeout=5) as response:
                    df = pd.read_csv(io.StringIO(response.read().decode('utf-8')))
                    df.columns = df.columns.str.strip()
                    if 'SYMBOL' in df.columns and 'SERIES' in df.columns:
                        df['SYMBOL'] = df['SYMBOL'].astype(str).str.strip()
                        df['SERIES'] = df['SERIES'].astype(str).str.strip()
                        bhavcopies.append((date_to_check, df[df['SERIES'] == 'EQ']))
                        print(f" ✅ Downloaded: {date_to_check.strftime('%d-%b-%Y')}")
            except Exception:
                pass
        date_to_check -= datetime.timedelta(days=1)
        attempts += 1

    return bhavcopies


def get_fallback_nifty500() -> pd.DataFrame:
    fallback_data = [
        {"Symbol": "RELIANCE",   "Company Name": "Reliance Industries Ltd",        "Industry": "Oil Gas & Consumable Fuels"},
        {"Symbol": "TCS",        "Company Name": "Tata Consultancy Services Ltd",   "Industry": "Information Technology"},
        {"Symbol": "HDFCBANK",   "Company Name": "HDFC Bank Ltd",                   "Industry": "Financial Services"},
        {"Symbol": "ICICIBANK",  "Company Name": "ICICI Bank Ltd",                  "Industry": "Financial Services"},
        {"Symbol": "BHARTIARTL", "Company Name": "Bharti Airtel Ltd",               "Industry": "Telecommunication"},
        {"Symbol": "INFY",       "Company Name": "Infosys Ltd",                     "Industry": "Information Technology"},
        {"Symbol": "ITC",        "Company Name": "ITC Ltd",                         "Industry": "Fast Moving Consumer Goods"},
        {"Symbol": "SBIN",       "Company Name": "State Bank of India",             "Industry": "Financial Services"},
        {"Symbol": "LT",         "Company Name": "Larsen & Toubro Ltd",             "Industry": "Construction"},
        {"Symbol": "HINDUNILVR", "Company Name": "Hindustan Unilever Ltd",          "Industry": "Fast Moving Consumer Goods"},
    ]
    return pd.DataFrame(fallback_data)


# ─────────────────────────────────────────────────────────────
# 3. SENTIMENT ENGINE
# ─────────────────────────────────────────────────────────────
def analyze_headline_sentiment(headline: str) -> float:
    tokens    = re.findall(r'\b\w+\b', headline.lower())
    pos_count = sum(1 for t in tokens if t in POSITIVE_WORDS)
    neg_count = sum(1 for t in tokens if t in NEGATIVE_WORDS)
    if pos_count == 0 and neg_count == 0:
        return 0.0
    return (pos_count - neg_count) / (pos_count + neg_count + 1)


def fetch_sector_news_sentiment(industry: str) -> tuple:
    search_query     = urllib.parse.quote(f"Nifty {industry} stock market news")
    url              = f"https://news.google.com/rss/search?q={search_query}&hl=en-IN&gl=IN&ceid=IN:en"
    headlines_records = []
    sentiment_scores  = []

    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=8) as response:
            root  = ET.fromstring(response.read())
            items = root.findall('.//item')
            for item in items[:3]:
                title      = item.find('title').text
                clean_title = re.sub(r'\s+-\s+[^$]+$', '', title)
                pub_date   = item.find('pubDate').text
                score      = analyze_headline_sentiment(clean_title)
                sentiment_scores.append(score)
                headlines_records.append({
                    "Sector": industry, "Headline": clean_title,
                    "Date": pub_date[:16] if pub_date else "",
                    "Sentiment Score": round(score, 2),
                    "Polarity": "Bullish" if score > 0.05 else "Bearish" if score < -0.05 else "Neutral"
                })
    except Exception:
        pass

    if not sentiment_scores:
        sentiment_scores = [0.0]
        headlines_records.append({
            "Sector": industry,
            "Headline": f"Select market actions in Nifty {industry} industries.",
            "Date": datetime.date.today().strftime("%d %b %Y"),
            "Sentiment Score": 0.0, "Polarity": "Neutral"
        })

    avg_score      = np.mean(sentiment_scores)
    sentiment_label = "BULLISH" if avg_score > 0.15 else "BEARISH" if avg_score < -0.15 else "NEUTRAL"
    return sentiment_label, round(avg_score, 2), headlines_records


# ─────────────────────────────────────────────────────────────
# 4. SECTOR INDEX SYNTHESIZER
# ─────────────────────────────────────────────────────────────
def synthesize_sector_index(sector_config: dict, constituents_df_dict: dict) -> pd.DataFrame:
    constituents = sector_config["constituents"]
    market_caps  = sector_config["market_caps"]
    closes = {}; volumes = {}; opens = {}; highs = {}; lows = {}

    for ticker in constituents:
        if ticker not in constituents_df_dict or constituents_df_dict[ticker].empty:
            continue
        df = constituents_df_dict[ticker]

        def safe_col(col_name, default=None):
            if col_name not in df.columns:
                return default
            val = df[col_name]
            if isinstance(val, pd.DataFrame):
                val = val.squeeze()
            if isinstance(val, pd.DataFrame):
                val = val.iloc[:, 0]
            return val

        c_series = safe_col('Close')
        if c_series is None or c_series.empty:
            continue
        closes[ticker]  = c_series
        volumes[ticker] = safe_col('Volume', pd.Series(0.0, index=c_series.index))
        opens[ticker]   = safe_col('Open',   c_series)
        highs[ticker]   = safe_col('High',   c_series)
        lows[ticker]    = safe_col('Low',    c_series)

    if not closes:
        return pd.DataFrame()

    closes_df  = pd.DataFrame(closes).dropna()
    if closes_df.empty:
        return pd.DataFrame()

    volumes_df = pd.DataFrame(volumes).reindex(closes_df.index).fillna(0.0)
    opens_df   = pd.DataFrame(opens).reindex(closes_df.index).fillna(closes_df)
    highs_df   = pd.DataFrame(highs).reindex(closes_df.index).fillna(closes_df)
    lows_df    = pd.DataFrame(lows).reindex(closes_df.index).fillna(closes_df)

    weights     = np.array([market_caps.get(t, 1.0) for t in closes_df.columns])
    sum_weights = np.sum(weights) or 1.0
    norm_weights = weights / sum_weights

    synth_close  = closes_df.dot(norm_weights)
    scale_factor = 100.0 / synth_close.iloc[0] if synth_close.iloc[0] != 0 else 1.0

    return pd.DataFrame({
        'Open':   opens_df.dot(norm_weights)  * scale_factor,
        'High':   highs_df.dot(norm_weights)  * scale_factor,
        'Low':    lows_df.dot(norm_weights)   * scale_factor,
        'Close':  synth_close                 * scale_factor,
        'Volume': volumes_df.sum(axis=1)
    }, index=closes_df.index)


# ─────────────────────────────────────────────────────────────
# 5. MAIN PIPELINE
# ─────────────────────────────────────────────────────────────
def main():
    # ── Parse CLI args ────────────────────────────────────────
    args = parse_args()

    start_time = time.time()
    print("=" * 70)
    print(" 🚀 NIFTY 500 MULTI-TIMEFRAME DELIVERY, SENTIMENT & TREND SCANNER")
    if args.force_refresh:
        print(" ⚡ MODE: FORCE REFRESH — Cache will be ignored and re-downloaded.")
    print("=" * 70)

    load_env_file()

    # ── Clean stale cache files ───────────────────────────────
    cleanup_old_caches()

    # 5.1 Fetch Nifty 500 constituents
    nifty500_df = fetch_nifty500_constituents()
    print("📂 Analyzing Nifty 500 industry distributions...")
    nifty500_df.columns = nifty500_df.columns.str.strip()
    nifty500_df = nifty500_df[~nifty500_df['Symbol'].astype(str).str.contains('DUMMY', case=False, na=True)]

    industry_groups = nifty500_df.groupby('Industry')
    sectors_config  = {}
    all_scan_tickers = []

    print(" Mapping all Nifty 500 constituent stocks to their respective industries...")
    for ind, group_df in industry_groups:
        symbols        = group_df['Symbol'].dropna().tolist()
        sector_symbols = [f"{s}.NS" for s in symbols]
        sector_key     = ind.upper().replace(" ", "_").replace("&", "AND").replace("-", "_").replace("/", "_")
        sectors_config[sector_key] = {
            "name": ind,
            "constituents": sector_symbols,
            "market_caps": {s: 1.0 for s in sector_symbols}
        }
        all_scan_tickers.extend(sector_symbols)

    all_scan_tickers = list(set(all_scan_tickers))
    benchmark_ticker = "^NSEI"
    print(f"✅ Loaded {len(sectors_config)} sectors comprising {len(all_scan_tickers)} stocks.")

    # 5.2 Download delivery bhavcopies (always fresh — daily data)
    bhavcopies = download_last_5_days_delivery_bhavcopies()
    if not bhavcopies:
        print("❌ Error: Could not download any NSE Bhavcopy Delivery reports.")
        return

    # Compile delivery metrics
    delivery_metrics = {}
    print("📊 Calculating Daily, Weekly, and 5-Day Average Deliveries...")
    for symbol_yf in all_scan_tickers:
        symbol = symbol_yf.replace(".NS", "")
        daily_deliv_per = 0.0; deliv_per_history = []; sum_deliv_qty = 0.0; sum_trd_qty = 0.0
        for idx, (date, df) in enumerate(bhavcopies):
            row = df[df['SYMBOL'] == symbol]
            if not row.empty:
                try:
                    trd_qnty   = float(row.iloc[0]['TTL_TRD_QNTY'])
                    deliv_qty  = float(row.iloc[0]['DELIV_QTY'])
                    deliv_per  = float(str(row.iloc[0]['DELIV_PER']).replace('%', '').strip())
                    if idx == 0:
                        daily_deliv_per = deliv_per
                    deliv_per_history.append(deliv_per)
                    sum_deliv_qty += deliv_qty
                    sum_trd_qty   += trd_qnty
                except Exception:
                    pass
        delivery_metrics[symbol_yf] = {
            "daily_delivery":  round((np.mean(deliv_per_history) if deliv_per_history else 0.0) / 100.0, 4),
            "avg_delivery":    round((np.mean(deliv_per_history) if deliv_per_history else 0.0) / 100.0, 4),
            "weekly_delivery": round((sum_deliv_qty / sum_trd_qty * 100 if sum_trd_qty > 0 else 0.0) / 100.0, 4),
        }
    print("✅ Delivery volume percentages calculated.")

    # ─────────────────────────────────────────────────────────
    # 5.3  SMART PRICE DOWNLOAD  (INCREMENTAL CACHE)
    #
    #  3 scenarios:
    #  A) Force refresh     → 10Y full download, cache overwrite
    #  B) Cache exists, same day → sirf load, kuch download nahi
    #  C) Cache exists, next day → sirf missing days fetch, merge
    #  D) No cache at all   → 10Y full download, cache save
    # ─────────────────────────────────────────────────────────
    all_tickers_to_download = list(set(all_scan_tickers + [benchmark_ticker]))
    chunk_size  = 100
    today       = datetime.date.today()
    cache_note  = ""

    # Load whatever cache exists (aaj ka ya purana)
    prices_raw, cache_last_date = load_any_existing_cache()

    if args.force_refresh or prices_raw is None:
        # ── Scenario A / D: Full 10Y download ────────────────
        reason = "force-refresh" if args.force_refresh else "no cache found"
        print(f"📥 Full 10-year download ({reason})...")
        closes_list = []; opens_list = []; highs_list = []; lows_list = []; volumes_list = []

        for i in range(0, len(all_tickers_to_download), chunk_size):
            chunk = all_tickers_to_download[i:i + chunk_size]
            print(f"  Batch {i//chunk_size + 1} ({len(chunk)} symbols)...")
            for attempt in range(2):
                try:
                    chunk_raw = yf.download(chunk, period="10y", progress=False)
                    if not chunk_raw.empty:
                        if isinstance(chunk_raw.columns, pd.MultiIndex):
                            for field, lst in [('Close', closes_list), ('Open', opens_list),
                                               ('High', highs_list),   ('Low', lows_list),
                                               ('Volume', volumes_list)]:
                                if field in chunk_raw:
                                    lst.append(chunk_raw[field])
                        else:
                            t = chunk[0]
                            closes_list.append(pd.DataFrame({t: chunk_raw['Close']}))
                            opens_list.append(pd.DataFrame({t: chunk_raw['Open']}))
                            highs_list.append(pd.DataFrame({t: chunk_raw['High']}))
                            lows_list.append(pd.DataFrame({t: chunk_raw['Low']}))
                            volumes_list.append(pd.DataFrame({t: chunk_raw['Volume']}))
                    time.sleep(0.5)
                    break
                except Exception as e:
                    if attempt == 0:
                        print(f"  ⚠️  Batch error: {e}. Retrying...")
                        time.sleep(2.0)
                    else:
                        print(f"  ❌ Retry failed: {e}")

        prices_raw = {
            'Close':  pd.concat(closes_list,  axis=1) if closes_list  else pd.DataFrame(),
            'Open':   pd.concat(opens_list,   axis=1) if opens_list   else pd.DataFrame(),
            'High':   pd.concat(highs_list,   axis=1) if highs_list   else pd.DataFrame(),
            'Low':    pd.concat(lows_list,    axis=1) if lows_list    else pd.DataFrame(),
            'Volume': pd.concat(volumes_list, axis=1) if volumes_list else pd.DataFrame(),
        }
        save_price_cache(prices_raw)
        cache_note = "⚡ Force Refreshed" if args.force_refresh else "🌐 Fresh Download (first run)"

    elif cache_last_date < today:
        # ── Scenario C: Next day — sirf missing days fetch ───
        print(f"📅 Cache is from {cache_last_date}, today is {today}.")
        print(f"⬇️  Fetching only missing days: {cache_last_date} → {today}...")

        new_data = fetch_incremental_data(
            tickers    = all_tickers_to_download,
            from_date  = cache_last_date,          # overlap 1 din for safety
            chunk_size = chunk_size
        )

        if not new_data['Close'].empty:
            prices_raw = merge_incremental(prices_raw, new_data)
            save_price_cache(prices_raw)
            new_rows = len(new_data['Close'])
            print(f"✅ Merged {new_rows} new rows into cache.")
            cache_note = f"🔄 Incremental Update (+{new_rows} rows)"
        else:
            print("⚠️  No new data returned. Using existing cache as-is.")
            cache_note = "📦 Cache Used (no new data)"

    else:
        # ── Scenario B: Same day — sirf load ─────────────────
        print("✅ Cache is up-to-date for today. No download needed.")
        cache_note = "📦 Cache Used (same day)"

    # ── Validate benchmark presence ───────────────────────────
    if prices_raw['Close'].empty or benchmark_ticker not in prices_raw['Close'].columns:
        print("❌ Error: yfinance failed to retrieve benchmark price data.")
        return

    bench_raw = prices_raw['Close'][benchmark_ticker].dropna()

    # Build constituent price histories dict
    constituents_df_dict = {}
    for ticker in all_scan_tickers:
        if ticker in prices_raw['Close'].columns:
            closes  = prices_raw['Close'][ticker].dropna()
            opens   = prices_raw['Open'][ticker].dropna()   if 'Open'   in prices_raw else closes
            highs   = prices_raw['High'][ticker].dropna()   if 'High'   in prices_raw else closes
            lows    = prices_raw['Low'][ticker].dropna()    if 'Low'    in prices_raw else closes
            volumes = prices_raw['Volume'][ticker].dropna() if 'Volume' in prices_raw else pd.Series(0.0, index=closes.index)
            if not closes.empty:
                constituents_df_dict[ticker] = pd.DataFrame(
                    {'Open': opens, 'High': highs, 'Low': lows, 'Close': closes, 'Volume': volumes}
                ).reindex(closes.index).ffill()

    # 5.4 Synthesize Sector Index Price curves
    print("📊 Synthesizing capital-weighted Sector Price Indices...")
    sectors_data = {}
    for key, sec in sectors_config.items():
        synth_df = synthesize_sector_index(sec, constituents_df_dict)
        if not synth_df.empty:
            sectors_data[key] = synth_df
    print("✅ Sectors price histories synthesized.")

    # 5.5 NLP Sentiment
    print("📰 Fetching Google News headlines & calculating NLP Sentiment scores...")
    sector_sentiment  = {}
    news_feed_records = []
    for key, sec in sectors_config.items():
        label, score, feed = fetch_sector_news_sentiment(sec["name"])
        sector_sentiment[key] = {"label": label, "score": score}
        news_feed_records.extend(feed)
        print(f" Sector: {sec['name']:<30} | Sentiment: {label:<10} (Score: {score:+.2f})")
    news_feed_df = pd.DataFrame(news_feed_records)

    # 5.6 Multi-Timeframe Scans
    print("⚡ Executing Multi-Timeframe scans...")
    daily_scan_rows  = []
    weekly_scan_rows = []

    for sector_key, sec in sectors_config.items():
        sector_name  = sec["name"]
        if sector_key not in sectors_data:
            continue
        sector_close = sectors_data[sector_key]['Close']

        for ticker in sec["constituents"]:
            if ticker not in constituents_df_dict:
                continue
            df = constituents_df_dict[ticker]
            if len(df) < 252:
                continue

            ticker_clean = ticker.replace(".NS", "")
            close  = df['Close'].iloc[-1]
            ema20  = calculate_ema(df['Close'], 20).iloc[-1]
            ema50  = calculate_ema(df['Close'], 50).iloc[-1]
            ema100 = calculate_ema(df['Close'], 100).iloc[-1]
            ema200 = calculate_ema(df['Close'], 200).iloc[-1]
            rsi    = calculate_rsi(df['Close'], 14).iloc[-1]

            rs_nifty  = calculate_rs_ratio(df['Close'], bench_raw, window=20)
            rs_sector = calculate_rs_ratio(df['Close'], sector_close, window=20)
            deliv     = delivery_metrics.get(ticker, {"daily_delivery": 0.0, "avg_delivery": 0.0, "weekly_delivery": 0.0})

            def rolling_max(n): return df['High'].rolling(min(n, len(df)),  min_periods=min(n//3, len(df))).max().iloc[-1]
            def rolling_min(n): return df['Low'].rolling(min(n, len(df)),   min_periods=min(n//3, len(df))).min().iloc[-1]

            d_signal = "NEUTRAL"
            if close > ema20 > ema50 > ema100 > ema200 and rsi > 55 and rs_nifty > 0:
                d_signal = "STRONG BUY"
            elif close > ema50 and rsi > 50 and rs_nifty > 0:
                d_signal = "BUY"
            elif close < ema200 or rsi < 40:
                d_signal = "AVOID"

            daily_scan_rows.append({
                "Ticker": ticker_clean, "Sector": sector_name, "Close": round(close, 2),
                "EMA 20": round(ema20, 2), "EMA 50": round(ema50, 2),
                "EMA 100": round(ema100, 2), "EMA 200": round(ema200, 2),
                "RSI (14)": round(rsi, 1),
                "RS (Nifty 50)": round(rs_nifty, 4), "RS (Sector)": round(rs_sector, 4),
                "Daily Delivery %": deliv["daily_delivery"], "5D Avg Delivery %": deliv["avg_delivery"],
                "52W High": round(rolling_max(252), 2),  "52W Low": round(rolling_min(252), 2),
                "2Y High":  round(rolling_max(504), 2),  "2Y Low":  round(rolling_min(504), 2),
                "5Y High":  round(rolling_max(1260), 2), "5Y Low":  round(rolling_min(1260), 2),
                "10Y High": round(rolling_max(2520), 2), "10Y Low": round(rolling_min(2520), 2),
                "Signal": d_signal
            })

            # Weekly
            df_weekly      = df.resample('W').agg({'Open':'first','High':'max','Low':'min','Close':'last','Volume':'sum'}).dropna()
            bench_weekly   = bench_raw.resample('W').last().dropna()
            sector_w_close = sector_close.resample('W').last().dropna()

            if len(df_weekly) >= 52:
                w_close  = df_weekly['Close'].iloc[-1]
                w_ema20  = calculate_ema(df_weekly['Close'], 20).iloc[-1]
                w_ema50  = calculate_ema(df_weekly['Close'], 50).iloc[-1]
                w_ema100 = calculate_ema(df_weekly['Close'], 100).iloc[-1]
                w_ema200 = calculate_ema(df_weekly['Close'], 200).iloc[-1]
                w_rsi    = calculate_rsi(df_weekly['Close'], 14).iloc[-1]
                w_rs_nifty  = calculate_rs_ratio(df_weekly['Close'], bench_weekly,   window=4)
                w_rs_sector = calculate_rs_ratio(df_weekly['Close'], sector_w_close, window=4)

                def w_rolling_max(n): return df_weekly['High'].rolling(min(n, len(df_weekly)), min_periods=min(n//3, len(df_weekly))).max().iloc[-1]
                def w_rolling_min(n): return df_weekly['Low'].rolling(min(n, len(df_weekly)),  min_periods=min(n//3, len(df_weekly))).min().iloc[-1]

                w_signal = "NEUTRAL"
                if w_close > w_ema20 > w_ema50 > w_ema100 > w_ema200 and w_rsi > 55 and w_rs_nifty > 0:
                    w_signal = "STRONG BUY"
                elif w_close > w_ema50 and w_rsi > 50 and w_rs_nifty > 0:
                    w_signal = "BUY"
                elif w_close < w_ema200 or w_rsi < 40:
                    w_signal = "AVOID"

                weekly_scan_rows.append({
                    "Ticker": ticker_clean, "Sector": sector_name, "Close": round(w_close, 2),
                    "EMA 20": round(w_ema20, 2), "EMA 50": round(w_ema50, 2),
                    "EMA 100": round(w_ema100, 2), "EMA 200": round(w_ema200, 2),
                    "RSI (14)": round(w_rsi, 1),
                    "RS (Nifty 50)": round(w_rs_nifty, 4), "RS (Sector)": round(w_rs_sector, 4),
                    "Weekly Delivery %": deliv["weekly_delivery"],
                    "52W High": round(w_rolling_max(52), 2),  "52W Low": round(w_rolling_min(52), 2),
                    "2Y High":  round(w_rolling_max(104), 2), "2Y Low":  round(w_rolling_min(104), 2),
                    "5Y High":  round(w_rolling_max(260), 2), "5Y Low":  round(w_rolling_min(260), 2),
                    "10Y High": round(w_rolling_max(520), 2), "10Y Low": round(w_rolling_min(520), 2),
                    "Signal": w_signal
                })

    daily_df  = pd.DataFrame(daily_scan_rows).sort_values(by=["Signal", "RSI (14)"], ascending=[True, False])
    weekly_df = pd.DataFrame(weekly_scan_rows).sort_values(by=["Signal", "RSI (14)"], ascending=[True, False])

    # 5.7 Sector Rankings & RRG Trails
    print("📊 Compiling final Sector Rankings...")
    sector_summary    = []
    rrg_trails_records = []

    for sector_key, sec in sectors_config.items():
        if sector_key not in sectors_data:
            continue
        sector_name  = sec["name"]
        group        = daily_df[daily_df["Sector"] == sector_name]
        sector_close = sectors_data[sector_key]['Close']

        breadth_20  = (group["Close"] > group["EMA 20"]).sum()  / len(group) if len(group) > 0 else 0.0
        breadth_50  = (group["Close"] > group["EMA 50"]).sum()  / len(group) if len(group) > 0 else 0.0
        breadth_200 = (group["Close"] > group["EMA 200"]).sum() / len(group) if len(group) > 0 else 0.0
        avg_rsi     = group["RSI (14)"].mean()          if len(group) > 0 else 50.0
        avg_delivery = group["5D Avg Delivery %"].mean() if len(group) > 0 else 0.0

        ratio_series, mom_series = calculate_rrg_coordinates(sector_close, bench_raw)
        latest_ratio = ratio_series.iloc[-1]
        latest_mom   = mom_series.iloc[-1]
        prev_ratio   = ratio_series.iloc[-2] if len(ratio_series) > 1 else latest_ratio
        prev_mom     = mom_series.iloc[-2]   if len(mom_series)   > 1 else latest_mom

        velocity, heading = calculate_velocity_and_heading(latest_ratio, latest_mom, prev_ratio, prev_mom)

        # Fixed quadrant logic
        if   latest_ratio >= 100 and latest_mom >= 100: quadrant = "LEADING"
        elif latest_ratio >= 100 and latest_mom <  100: quadrant = "WEAKENING"
        elif latest_ratio <  100 and latest_mom <  100: quadrant = "LAGGING"
        else:                                           quadrant = "IMPROVING"

        heading_score    = max(0.0, np.cos(np.radians(heading - 45))) * 100
        rotational_score = (
            (avg_rsi        * 0.25) +
            ((breadth_50 * 100) * 0.30) +
            (heading_score  * 0.25) +
            ((avg_delivery * 100) * 0.20)
        )

        sector_summary.append({
            "Sector Name": sector_name,
            "Avg Stock RSI": round(avg_rsi, 1),
            "20 EMA Breadth": round(breadth_20, 4), "50 EMA Breadth": round(breadth_50, 4),
            "200 EMA Breadth": round(breadth_200, 4),
            "RS-Ratio": round(latest_ratio, 2), "RS-Momentum": round(latest_mom, 2),
            "Velocity": round(velocity, 2),     "Heading": round(heading, 1),
            "RRG Quadrant": quadrant,           "Rotational Score": round(rotational_score, 1)
        })

        for dt in ratio_series.index[-10:]:
            r_val = ratio_series.loc[dt]; m_val = mom_series.loc[dt]
            if   r_val >= 100 and m_val >= 100: q_val = "LEADING"
            elif r_val >= 100 and m_val <  100: q_val = "WEAKENING"
            elif r_val <  100 and m_val <  100: q_val = "LAGGING"
            else:                               q_val = "IMPROVING"
            rrg_trails_records.append({
                "Date": dt.strftime("%Y-%m-%d"), "Sector": sector_name,
                "RS-Ratio": round(r_val, 2), "RS-Momentum": round(m_val, 2), "RRG Quadrant": q_val
            })

    sector_rank_df  = pd.DataFrame(sector_summary).sort_values(by="Rotational Score", ascending=False)
    rrg_trails_df   = pd.DataFrame(rrg_trails_records)

    # ─────────────────────────────────────────────────────────
    # 6. EXCEL OUTPUT
    # ─────────────────────────────────────────────────────────
    print(f"5. Writing output to {OUTPUT_FILE}...")
    if os.path.exists(OUTPUT_FILE):
        try: os.remove(OUTPUT_FILE)
        except: pass

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        daily_df.to_excel(writer,       sheet_name="Daily Scanner",    index=False)
        weekly_df.to_excel(writer,      sheet_name="Weekly Scanner",   index=False)
        sector_rank_df.to_excel(writer, sheet_name="Sector Rankings",  index=False)
        rrg_trails_df.to_excel(writer,  sheet_name="Sector RRG Trails",index=False)
        news_feed_df.to_excel(writer,   sheet_name="Sector News Feed", index=False)

    # Styling
    wb          = openpyxl.load_workbook(OUTPUT_FILE)
    font_family = "Segoe UI"
    header_fill = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    header_font = Font(name=font_family, size=11, bold=True, color="FFFFFF")
    cell_font   = Font(name=font_family, size=11)
    buy_fill    = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    avoid_fill  = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),  bottom=Side(style='thin', color='D9D9D9')
    )

    for ws_name in wb.sheetnames:
        ws = wb[ws_name]
        ws.views.sheetView[0].showGridLines = True

        # Build name→index map from header row
        col_map = {ws.cell(1, c).value: c for c in range(1, ws.max_column + 1)}

        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.font = header_font; cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[1].height = 28

        for row_idx in range(2, ws.max_row + 1):
            ws.row_dimensions[row_idx].height = 20
            close_val = 0.0
            try: close_val = float(ws.cell(row_idx, col_map.get("Close", 3)).value)
            except: pass

            for col_idx in range(1, ws.max_column + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.font = cell_font; cell.border = thin_border
                cell.alignment = Alignment(vertical="center")

                if ws_name in ["Daily Scanner", "Weekly Scanner"]:
                    col_name = ws.cell(1, col_idx).value

                    # EMA columns: green if price > EMA, red otherwise
                    if col_name in ["EMA 20", "EMA 50", "EMA 100", "EMA 200"]:
                        try:
                            ema_val = float(cell.value)
                            if close_val > ema_val:
                                cell.fill = buy_fill
                                cell.font = Font(name=font_family, size=11, color="006100")
                            else:
                                cell.fill = avoid_fill
                                cell.font = Font(name=font_family, size=11, color="9C0006")
                        except: pass

                    # Near multi-year highs (within 10%)
                    if col_name in ["52W High", "2Y High", "5Y High", "10Y High"]:
                        try:
                            high_val = float(cell.value)
                            if high_val > 0 and close_val >= 0.90 * high_val:
                                cell.fill = buy_fill
                                cell.font = Font(name=font_family, size=11, color="006100")
                        except: pass

                    # Near multi-year lows (within 10%)
                    if col_name in ["52W Low", "2Y Low", "5Y Low", "10Y Low"]:
                        try:
                            low_val = float(cell.value)
                            if low_val > 0 and close_val <= 1.10 * low_val:
                                cell.fill = avoid_fill
                                cell.font = Font(name=font_family, size=11, color="9C0006")
                        except: pass

                    # Color-code Signal Column (last column)
                    if col_idx == ws.max_column:
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                        if cell.value == "STRONG BUY":
                            cell.fill = buy_fill
                            cell.font = Font(name=font_family, size=11, bold=True, color="006100")
                        elif cell.value == "BUY":
                            cell.fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
                            cell.font = Font(name=font_family, size=11, color="375623")
                        elif cell.value == "AVOID":
                            cell.fill = avoid_fill
                            cell.font = Font(name=font_family, size=11, bold=True, color="9C0006")

                elif ws_name == "Sector Rankings":
                    # Columns: Sector Name (1), Avg Stock RSI (2), 20 EMA Breadth (3), 50 EMA Breadth (4), 200 EMA Breadth (5),
                    # RS-Ratio (6), RS-Momentum (7), Velocity (8), Heading (9), RRG Quadrant (10), Rotational Score (11)
                    if col_idx in [2, 3, 4, 5, 6, 7, 8, 9, 11]:
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    # Format technical breadths as percentage
                    if col_idx in [3, 4, 5]:
                        cell.number_format = '0.00%'
                    # Format RRG ratios & velocity
                    if col_idx in [6, 7, 8]:
                        cell.number_format = '0.00'
                    # Format Heading compass degrees & Score
                    if col_idx in [9, 11]:
                        cell.number_format = '0.0'
                    # Highlight Sector RRG Quadrant (col 10)
                    if col_idx == 10:
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                        if cell.value == "LEADING":
                            cell.fill = buy_fill
                            cell.font = Font(name=font_family, size=11, bold=True, color="006100")
                        elif cell.value == "IMPROVING":
                            cell.fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
                            cell.font = Font(name=font_family, size=11, bold=True, color="1F497D")
                        elif cell.value == "WEAKENING":
                            cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
                            cell.font = Font(name=font_family, size=11, bold=True, color="7F6000")
                        elif cell.value == "LAGGING":
                            cell.fill = avoid_fill
                            cell.font = Font(name=font_family, size=11, bold=True, color="9C0006")

                elif ws_name == "Sector RRG Trails":
                    # Columns: Date (1), Sector (2), RS-Ratio (3), RS-Momentum (4), RRG Quadrant (5)
                    if col_idx in [3, 4]:
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                        cell.number_format = '0.00'
                    # Highlight Quadrant column (Col 5)
                    if col_idx == 5:
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                        if cell.value == "LEADING":
                            cell.fill = buy_fill
                            cell.font = Font(name=font_family, size=11, bold=True, color="006100")
                        elif cell.value == "IMPROVING":
                            cell.fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
                            cell.font = Font(name=font_family, size=11, bold=True, color="1F497D")
                        elif cell.value == "WEAKENING":
                            cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
                            cell.font = Font(name=font_family, size=11, bold=True, color="7F6000")
                        elif cell.value == "LAGGING":
                            cell.fill = avoid_fill
                            cell.font = Font(name=font_family, size=11, bold=True, color="9C0006")

                elif ws_name == "Sector News Feed":
                    if col_idx == 4:
                        cell.number_format = '+0.00;-0.00;0.00'
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    if col_idx == 5:
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                        if cell.value == "Bullish":
                            cell.fill = buy_fill
                            cell.font = Font(name=font_family, size=11, bold=True, color="006100")
                        elif cell.value == "Bearish":
                            cell.fill = avoid_fill
                            cell.font = Font(name=font_family, size=11, bold=True, color="9C0006")

        # Autofit columns
        for col in ws.columns:
            max_len = 10
            for cell in col:
                val_str = str(cell.value or '')
                if cell.number_format and '%' in cell.number_format:
                    max_len = max(max_len, len(val_str) + 3)
                else:
                    max_len = max(max_len, len(val_str))
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = min(max_len + 4, 38)

    wb.save(OUTPUT_FILE)

    # ─────────────────────────────────────────────────────────────
    # TELEGRAM BOT ALERT NOTIFICATION DISPATCHER
    # ─────────────────────────────────────────────────────────────
    try:
        dt_str = bhavcopies[0][0].strftime("%d-%b-%Y")

        # 2. Extract Top 3 RRG sectors
        top_sectors = sector_rank_df.head(3)
        sectors_text = ""
        for idx in range(len(top_sectors)):
            row = top_sectors.iloc[idx]
            name = row["Sector Name"]
            score = row["Rotational Score"]
            quad = row["RRG Quadrant"]
            heading = row["Heading"]
            quad_emoji = "🟢" if quad == "LEADING" else "🔵" if quad == "IMPROVING" else "🟡" if quad == "WEAKENING" else "🔴"
            sectors_text += f"{idx+1}. {quad_emoji} *{name}* (Score: {score:.1f}) -> {quad} [Heading: {heading:.1f}°]\n"

        # 3. Extract Top 5 Strong Buy / Buy Stocks
        strong_buys = daily_df[daily_df["Signal"] == "STRONG BUY"].head(5)
        if len(strong_buys) < 5:
            buys = daily_df[daily_df["Signal"] == "BUY"].head(5 - len(strong_buys))
            strong_buys = pd.concat([strong_buys, buys])

        stocks_text = ""
        for idx in range(min(5, len(strong_buys))):
            row = strong_buys.iloc[idx]
            ticker = row["Ticker"]
            close = row["Close"]
            rsi = row["RSI (14)"]
            deliv = row["5D Avg Delivery %"]
            sig = row["Signal"]
            sig_emoji = "🔥" if sig == "STRONG BUY" else "⚡"
            stocks_text += f" - {sig_emoji} *{ticker}*: CMP: {close:.1f} | RSI: {rsi:.1f} | 5D Deliv: {deliv * 100:.1f}%\n"

        text_summary = (
            f"🔔 *NIFTY 500 ROTATION SCAN COMPLETED* 🔔\n"
            f"📅 Date: {dt_str}\n"
            f"⚡ Cache: {cache_note}\n\n"
            f"🚀 *TOP 3 STRONGEST SECTORS (RRG)*\n{sectors_text}\n"
            f"📈 *TOP 5 STRONGEST SCANNER CANDIDATES*\n{stocks_text}\n"
            f"📂 *Attached is your updated 5-sheet Sector Rotation Excel Dashboard!*"
        )

        # Load environment credentials and trigger notification
        load_env_file()
        send_telegram_alerts(text_summary, OUTPUT_FILE)

    except Exception as telegram_e:
        print(f"⚠️ Warning: Failed to compile or send Telegram summary: {telegram_e}")

    elapsed = time.time() - start_time
    print("=" * 70)
    print(f"🎉 MASTER REPORT COMPILED SUCCESSFULLY: {OUTPUT_FILE}")
    print(f"Time Taken : {elapsed:.1f} seconds")
    print("Sheets created in workbook:")
    print("  1. 'Daily Scanner'     - Multi-EMA, RSI, RS Nifty/Sector, and Delivery Scanner (Daily)")
    print("  2. 'Weekly Scanner'    - Weekly trend, Weekly EMAs, and Weekly Delivery %")
    print("  3. 'Sector Rankings'   - Aggregated RRG coordinates, technical breadth, & rotational momentum rankings")
    print("  4. 'Sector RRG Trails' - 10-day historical trailing path of RS-Ratio & RS-Momentum coordinates")
    print("  5. 'Sector News Feed'  - Real-time Google News headlines & polarity sentiment scores")
    print("=" * 70)

if __name__ == "__main__":
    main()
