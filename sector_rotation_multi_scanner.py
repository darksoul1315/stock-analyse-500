#!/usr/bin/env python3
"""
Nifty 500 Delivery, Sentiment & Trend Scanner
Compiles multi-timeframe moving averages, outperformance ratios (RS), multi-year high/low price bounds,
NSE daily/weekly deliverable volume percentages, and real-time Google News sentiment polarity scores.
Outputs the styled 4-sheet report: 'sector_rotation_multi_report.xlsx'.
"""

import os
import sys
import io
import re
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

# -------------------------------------------------------------
# TELEGRAM BOT INTEGRATION MODULE
# -------------------------------------------------------------
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
    # Load credentials
    token = os.environ.get("TELEGRAM_BOT_TOKEN")
    chat_id = os.environ.get("TELEGRAM_CHAT_ID")
    
    if not token or not chat_id:
        print("⚠️ Warning: Telegram credentials (TELEGRAM_BOT_TOKEN / TELEGRAM_CHAT_ID) not found. Skipping alert.")
        return False
        
    print("📤 Dispatching Telegram notifications...")
    
    # 1. Send Text Summary Alert
    text_url = f"https://api.telegram.org/bot{token}/sendMessage"
    payload = {
        "chat_id": chat_id,
        "text": text_summary,
        "parse_mode": "Markdown"
    }
    
    try:
        req_data = urllib.parse.urlencode(payload).encode("utf-8")
        req = urllib.request.Request(text_url, data=req_data, method="POST")
        with urllib.request.urlopen(req, timeout=12) as response:
            print("   ✅ Text alert sent successfully.")
    except Exception as e:
        print(f"   ❌ Failed to send Telegram text alert: {e}")
        return False
        
    # 2. Upload Excel Report file
    if file_path and os.path.exists(file_path):
        doc_url = f"https://api.telegram.org/bot{token}/sendDocument"
        
        try:
            # Construct manual multipart/form-data payload in python
            boundary = "----WebKitFormBoundary7MA4YWxkTrZu0gW"
            parts = []
            
            # Chat ID field
            parts.append(f"--{boundary}\r\nContent-Disposition: form-data; name=\"chat_id\"\r\n\r\n{chat_id}\r\n")
            
            # Document file field
            file_name = os.path.basename(file_path)
            parts.append(f"--{boundary}\r\nContent-Disposition: form-data; name=\"document\"; filename=\"{file_name}\"\r\nContent-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n")
            
            # Read binary file content
            with open(file_path, "rb") as f:
                file_content = f.read()
                
            # Compile multipart raw payload
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
            with urllib.request.urlopen(req, timeout=25) as response:
                print("   ✅ Excel report uploaded successfully to Telegram.")
            return True
        except Exception as e:
            print(f"   ❌ Failed to upload Excel report to Telegram: {e}")
            return False
            
    return True

# Define output
OUTPUT_FILE = "sector_rotation_multi_report.xlsx"

# Lexicons for financial sentiment analysis
POSITIVE_WORDS = {
    'bullish', 'surge', 'jump', 'gain', 'positive', 'profit', 'growth', 'demand', 
    'outperform', 'high', 'rise', 'record', 'strong', 'expansion', 'rally', 'upgrade', 
    'buy', 'accumulate', 'double', 'success', 'win', 'momentum', 'revival', 'recover',
    'expansion', 'bull', 'improving', 'leadership', 'outperforming', 'breakout'
}
NEGATIVE_WORDS = {
    'bearish', 'drop', 'plunge', 'loss', 'negative', 'decline', 'compress', 'fall', 
    'low', 'weak', 'warning', 'crash', 'slash', 'downgrade', 'sell', 'avoid', 'deficit', 
    'compression', 'drag', 'penalty', 'delay', 'slump', 'pessimism', 'inflation',
    'distribution', 'bear', 'weakening', 'lagging', 'underperforming', 'whipsaw'
}

# -------------------------------------------------------------
# 1. CORE TECHNICAL INDICATORS (Vectorized)
# -------------------------------------------------------------
def calculate_ema(series: pd.Series, period: int) -> pd.Series:
    return series.ewm(span=period, adjust=False).mean()

def calculate_rsi(series: pd.Series, period: int = 14) -> pd.Series:
    delta = series.diff()
    gain = delta.clip(lower=0)
    loss = -delta.clip(upper=0)
    avg_gain = gain.ewm(alpha=1/period, adjust=False).mean()
    avg_loss = loss.ewm(alpha=1/period, adjust=False).mean()
    rs = avg_gain / avg_loss.replace(0, np.nan)
    rsi = 100 - (100 / (1 + rs))
    return rsi.fillna(50)

def calculate_rs_ratio(stock_series: pd.Series, compare_series: pd.Series, window: int) -> float:
    """Calculates performance outperformance ratio: Stock Return / Comparison Return - 1."""
    if len(stock_series) < window or len(compare_series) < window:
        return 0.0
    stock_ret = stock_series.iloc[-1] / stock_series.iloc[-window] - 1
    compare_ret = compare_series.iloc[-1] / compare_series.iloc[-window] - 1
    
    # Avoid division by zero
    if compare_ret == -1.0:
        return 0.0
    return stock_ret - compare_ret

# -------------------------------------------------------------
# 1.1 RRG COORDINATES ENGINE
# -------------------------------------------------------------
def calculate_rrg_coordinates(sector_series: pd.Series, bench_series: pd.Series) -> tuple:
    """
    Calculates Julius de Kempenaer style RRG coordinates (RS-Ratio and RS-Momentum)
    for a sector price series relative to a benchmark index.
    Returns two Series: (rs_ratio, rs_momentum).
    """
    # Align the two series
    aligned_df = pd.DataFrame({'Sector': sector_series, 'Bench': bench_series}).dropna()
    if len(aligned_df) < 120:  # Needs at least 120 points for reliable standard deviations
        # Fallback to neutral 100
        return pd.Series(100.0, index=sector_series.index), pd.Series(100.0, index=sector_series.index)
        
    sec = aligned_df['Sector']
    bench = aligned_df['Bench']
    
    # 1. Relative Strength Price Ratio
    rs = (sec / bench) * 100
    
    # 2. RS-Ratio (trend component)
    ema_rs = rs.ewm(span=60, adjust=False).mean()
    std_rs = rs.rolling(window=60, min_periods=30).std()
    rs_ratio = 100 + 10 * ((rs - ema_rs) / std_rs.replace(0, np.nan)).fillna(0)
    
    # 3. RS-Momentum (momentum component)
    d_ratio = rs_ratio.diff().fillna(0)
    ema_d_fast = d_ratio.ewm(span=10, adjust=False).mean()
    ema_d_slow = d_ratio.ewm(span=60, adjust=False).mean()
    std_d = d_ratio.rolling(window=60, min_periods=30).std()
    
    rs_momentum = 100 + 10 * ((ema_d_fast - ema_d_slow) / std_d.replace(0, np.nan)).fillna(0)
    
    # Reindex to match the original index
    return rs_ratio.reindex(sector_series.index).fillna(100.0), rs_momentum.reindex(sector_series.index).fillna(100.0)

def calculate_velocity_and_heading(ratio: float, momentum: float, prev_ratio: float, prev_momentum: float) -> tuple:
    """Calculates RRG coordinate rotational velocity and heading angle in degrees (0-360)."""
    dx = ratio - prev_ratio
    dy = momentum - prev_momentum
    
    velocity = np.sqrt(dx**2 + dy**2)
    
    # Heading angle in radians, convert to compass degrees
    angle_rad = np.arctan2(dy, dx)
    angle_deg = np.degrees(angle_rad)
    
    # Standardise angle into 0 to 360 degrees range
    heading = (angle_deg + 360) % 360
    
    return round(float(velocity), 2), round(float(heading), 1)

# -------------------------------------------------------------
# 2. DYNAMIC NIFTY 500 & NSE DELIVERY BHAVCOPY DOWNLOADERS
# -------------------------------------------------------------
def fetch_nifty500_constituents() -> pd.DataFrame:
    """Fetches the official Nifty 500 stock constituents list directly from the NSE website."""
    url = "https://archives.nseindia.com/content/indices/ind_nifty500list.csv"
    print(f"📥 Fetching Nifty 500 list from NSE: {url}")
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=10) as response:
            csv_data = response.read().decode('utf-8')
            df = pd.read_csv(io.StringIO(csv_data))
            df.columns = df.columns.str.strip()
            return df
    except Exception as e:
        print(f"⚠️ Warning: Failed to fetch online Nifty 500 list ({e}). Using hardcoded bluechip backup.")
        return get_fallback_nifty500()

def download_last_5_days_delivery_bhavcopies() -> list:
    """
    Loops backwards from today's date to download the last 5 active daily 
    NSE Security Bhavcopy reports. Automatically skips weekends and holidays.
    """
    bhavcopies = []
    date_to_check = datetime.date.today()
    attempts = 0
    
    print("📥 Scanning NSE archives for the last 5 active Delivery Bhavcopy reports...")
    
    while len(bhavcopies) < 5 and attempts < 15:
        # Check if date is weekday
        if date_to_check.weekday() < 5:
            date_str = date_to_check.strftime("%d%m%Y")
            url = f"https://nsearchives.nseindia.com/products/content/sec_bhavdata_full_{date_str}.csv"
            
            try:
                req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'})
                with urllib.request.urlopen(req, timeout=5) as response:
                    csv_data = response.read().decode('utf-8')
                    df = pd.read_csv(io.StringIO(csv_data))
                    df.columns = df.columns.str.strip()
                    
                    # Ensure columns EQ series exists
                    if 'SYMBOL' in df.columns and 'SERIES' in df.columns:
                        # Clean columns
                        df['SYMBOL'] = df['SYMBOL'].astype(str).str.strip()
                        df['SERIES'] = df['SERIES'].astype(str).str.strip()
                        # Filter standard equities series
                        df_eq = df[df['SERIES'] == 'EQ']
                        
                        bhavcopies.append((date_to_check, df_eq))
                        print(f"   ✅ Downloaded active report for: {date_to_check.strftime('%d-%b-%Y')}")
            except Exception as e:
                # Silently skip non-trading days
                pass
        date_to_check -= datetime.timedelta(days=1)
        attempts += 1
        
    return bhavcopies

def get_fallback_nifty500() -> pd.DataFrame:
    """Returns a robust fallback dataframe representing top NSE sectors."""
    fallback_data = [
        {"Symbol": "RELIANCE", "Company Name": "Reliance Industries Ltd", "Industry": "Oil Gas & Consumable Fuels"},
        {"Symbol": "TCS", "Company Name": "Tata Consultancy Services Ltd", "Industry": "Information Technology"},
        {"Symbol": "HDFCBANK", "Company Name": "HDFC Bank Ltd", "Industry": "Financial Services"},
        {"Symbol": "ICICIBANK", "Company Name": "ICICI Bank Ltd", "Industry": "Financial Services"},
        {"Symbol": "BHARTIARTL", "Company Name": "Bharti Airtel Ltd", "Industry": "Telecommunication"},
        {"Symbol": "INFY", "Company Name": "Infosys Ltd", "Industry": "Information Technology"},
        {"Symbol": "ITC", "Company Name": "ITC Ltd", "Industry": "Fast Moving Consumer Goods"},
        {"Symbol": "SBIN", "Company Name": "State Bank of India", "Industry": "Financial Services"},
        {"Symbol": "LENT", "Company Name": "Larsen & Toubro Ltd", "Industry": "Construction"},
        {"Symbol": "HINDUNILVR", "Company Name": "Hindustan Unilever Ltd", "Industry": "Fast Moving Consumer Goods"}
    ]
    return pd.DataFrame(fallback_data)

# -------------------------------------------------------------
# 3. FINANCIAL NEWS NLP SENTIMENT ENGINE
# -------------------------------------------------------------
def analyze_headline_sentiment(headline: str) -> float:
    """Lexicon-based stock sentiment scoring (-1.0 to +1.0)."""
    tokens = re.findall(r'\b\w+\b', headline.lower())
    pos_count = sum(1 for t in tokens if t in POSITIVE_WORDS)
    neg_count = sum(1 for t in tokens if t in NEGATIVE_WORDS)
    
    if pos_count == 0 and neg_count == 0:
        return 0.0
    return (pos_count - neg_count) / (pos_count + neg_count + 1)

def fetch_sector_news_sentiment(industry: str) -> tuple:
    """
    Fetches real-time financial headlines for an industry using Google News RSS,
    calculates NLP sentiment polarity scores, and outputs feed records.
    """
    # Sanitize industry name for search
    search_query = urllib.parse.quote(f"Nifty {industry} stock market news")
    url = f"https://news.google.com/rss/search?q={search_query}&hl=en-IN&gl=IN&ceid=IN:en"
    
    headlines_records = []
    sentiment_scores = []
    
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=8) as response:
            xml_data = response.read()
            root = ET.fromstring(xml_data)
            items = root.findall('.//item')
            
            # Extract top 3 headlines
            for item in items[:3]:
                title = item.find('title').text
                # Strip source string from title (usually ends with " - Source")
                clean_title = re.sub(r'\s+-\s+[^$]+$', '', title)
                pub_date = item.find('pubDate').text
                
                score = analyze_headline_sentiment(clean_title)
                sentiment_scores.append(score)
                
                headlines_records.append({
                    "Sector": industry,
                    "Headline": clean_title,
                    "Date": pub_date[:16] if pub_date else "",
                    "Sentiment Score": round(score, 2),
                    "Polarity": "Bullish" if score > 0.05 else "Bearish" if score < -0.05 else "Neutral"
                })
    except Exception as e:
        # Graceful fallback on network timeout
        pass
        
    if not sentiment_scores:
        # Default neutral if feed fails
        sentiment_scores = [0.0]
        headlines_records.append({
            "Sector": industry,
            "Headline": f"Select market actions in Nifty {industry} industries.",
            "Date": datetime.date.today().strftime("%d %b %Y"),
            "Sentiment Score": 0.0,
            "Polarity": "Neutral"
        })
        
    avg_score = np.mean(sentiment_scores)
    sentiment_label = "BULLISH" if avg_score > 0.15 else "BEARISH" if avg_score < -0.15 else "NEUTRAL"
    
    return sentiment_label, round(avg_score, 2), headlines_records

# -------------------------------------------------------------
# 4. CAP-WEIGHTED SECTOR INDEX SYNTHESIZER
# -------------------------------------------------------------
def synthesize_sector_index(sector_config: dict, constituents_df_dict: dict) -> pd.DataFrame:
    """Calculates a Market Cap-Weighted daily index price & volume curve from constituents."""
    constituents = sector_config["constituents"]
    market_caps = sector_config["market_caps"]
    
    closes = {}
    volumes = {}
    opens = {}
    highs = {}
    lows = {}
    
    for ticker in constituents:
        if ticker in constituents_df_dict and not constituents_df_dict[ticker].empty:
            df = constituents_df_dict[ticker]
            
            # Helper to extract a 1D Series safely regardless of MultiIndex/DataFrames
            def safe_col(col_name, default_series=None):
                if col_name not in df.columns:
                    return default_series
                val = df[col_name]
                if isinstance(val, pd.DataFrame):
                    val = val.squeeze()
                    if isinstance(val, pd.DataFrame):
                        val = val.iloc[:, 0]
                return val
                
            c_series = safe_col('Close')
            if c_series is None or c_series.empty:
                continue
                
            closes[ticker] = c_series
            volumes[ticker] = safe_col('Volume', pd.Series(0.0, index=c_series.index))
            opens[ticker] = safe_col('Open', c_series)
            highs[ticker] = safe_col('High', c_series)
            lows[ticker] = safe_col('Low', c_series)
            
    if not closes:
        return pd.DataFrame()
        
    closes_df = pd.DataFrame(closes).dropna()
    if closes_df.empty:
        return pd.DataFrame()
        
    volumes_df = pd.DataFrame(volumes).reindex(closes_df.index).fillna(0.0)
    opens_df = pd.DataFrame(opens).reindex(closes_df.index).fillna(closes_df)
    highs_df = pd.DataFrame(highs).reindex(closes_df.index).fillna(closes_df)
    lows_df = pd.DataFrame(lows).reindex(closes_df.index).fillna(closes_df)
    
    # Extract weights based on market cap
    weights = np.array([market_caps.get(t, 1.0) for t in closes_df.columns])
    sum_weights = np.sum(weights)
    if sum_weights == 0:
        sum_weights = 1.0
        weights = np.ones(len(closes_df.columns))
        
    norm_weights = weights / sum_weights
    
    # Calculate Cap-Weighted Closes
    synth_close = closes_df.dot(norm_weights)
    synth_open = opens_df.dot(norm_weights)
    synth_high = highs_df.dot(norm_weights)
    synth_low = lows_df.dot(norm_weights)
    
    synth_volume = volumes_df.sum(axis=1)
    
    # Scale index so the very first close is baseline 100.0
    scale_factor = 100.0 / synth_close.iloc[0] if len(synth_close) > 0 and synth_close.iloc[0] != 0 else 1.0
    
    synth_df = pd.DataFrame({
        'Open': synth_open * scale_factor,
        'High': synth_high * scale_factor,
        'Low': synth_low * scale_factor,
        'Close': synth_close * scale_factor,
        'Volume': synth_volume
    }, index=closes_df.index)
    
    return synth_df

# -------------------------------------------------------------
# 5. DYNAMIC MULTI-TIMEFRAME PIPELINE RUNNER
# -------------------------------------------------------------
def main():
    start_time = time.time()
    print("=" * 70)
    print(" 🚀 NIFTY 500 MULTI-TIMEFRAME DELIVERY, SENTIMENT & TREND SCANNER")
    print("=" * 70)
    
    # 5.1 Fetch dynamic Nifty 500 constituents
    nifty500_df = fetch_nifty500_constituents()
    
    # Group Nifty 500 stocks by Industry Sector
    # The columns are Symbol, Company Name, Industry
    print("📂 Analyzing Nifty 500 industry distributions...")
    
    # Standardize column names and strip spaces
    nifty500_df.columns = nifty500_df.columns.str.strip()
    
    # Sanitize the list by removing any dummy symbols or footnotes
    nifty500_df = nifty500_df[~nifty500_df['Symbol'].astype(str).str.contains('DUMMY', case=False, na=True)]
    
    industry_groups = nifty500_df.groupby('Industry')
    
    sectors_config = {}
    all_scan_tickers = []
    
    # Map all constituents and cap sizes per sector
    print("   Mapping all Nifty 500 constituent stocks to their respective industries...")
    for ind, group_df in industry_groups:
        symbols = group_df['Symbol'].dropna().tolist()
        sector_symbols = [f"{s}.NS" for s in symbols]
        
        # Save config (use equal weight proxy 1.0 since they are all part of the Nifty 500)
        sector_key = ind.upper().replace(" ", "_").replace("&", "AND").replace("-", "_").replace("/", "_")
        sectors_config[sector_key] = {
            "name": ind,
            "constituents": sector_symbols,
            "market_caps": {s: 1.0 for s in sector_symbols} # Equal weight index synthesis
        }
        all_scan_tickers.extend(sector_symbols)
        
    all_scan_tickers = list(set(all_scan_tickers))
    benchmark_ticker = "^NSEI"
    
    print(f"✅ Loaded {len(sectors_config)} major Nifty 500 sectors comprising {len(all_scan_tickers)} stocks.")
    
    # 5.2 Download active NSE Bhavcopy Delivery position files
    bhavcopies = download_last_5_days_delivery_bhavcopies()
    
    if not bhavcopies:
        print("❌ Error: Could not download any active NSE Bhavcopy Delivery reports.")
        return
        
    latest_bhav_date = bhavcopies[0][0]
    
    # Compile delivery metrics dictionary
    delivery_metrics = {}
    print("📊 Calculating Daily, Weekly, and 5-Day Average Deliveries...")
    
    # Loop over all scanned stocks
    for symbol_yf in all_scan_tickers:
        symbol = symbol_yf.replace(".NS", "")
        
        # Extract data across the 5 reports
        daily_deliv_per = 0.0
        deliv_per_history = []
        sum_deliv_qty = 0.0
        sum_trd_qty = 0.0
        
        for idx, (date, df) in enumerate(bhavcopies):
            row = df[df['SYMBOL'] == symbol]
            if not row.empty:
                try:
                    trd_qnty = float(row.iloc[0]['TTL_TRD_QNTY'])
                    deliv_qty = float(row.iloc[0]['DELIV_QTY'])
                    deliv_per_str = str(row.iloc[0]['DELIV_PER']).replace('%', '').strip()
                    deliv_per = float(deliv_per_str)
                    
                    if idx == 0:
                        daily_deliv_per = deliv_per
                        
                    deliv_per_history.append(deliv_per)
                    sum_deliv_qty += deliv_qty
                    sum_trd_qty += trd_qnty
                except:
                    pass
                    
        # Compute calculations
        avg_deliv_per = np.mean(deliv_per_history) if deliv_per_history else 0.0
        weekly_deliv_per = (sum_deliv_qty / sum_trd_qty * 100) if sum_trd_qty > 0 else 0.0
        
        delivery_metrics[symbol_yf] = {
            "daily_delivery": round(daily_deliv_per / 100.0, 4), # stored as decimal (e.g. 0.4500 for 45%)
            "avg_delivery": round(avg_deliv_per / 100.0, 4),
            "weekly_delivery": round(weekly_deliv_per / 100.0, 4)
        }
        
    print("✅ Delivery volume percentages calculated.")

    # 5.3 Fetch Stock & Benchmark Close Prices (10-Year window for extremes) in chunks
    print("📥 Downloading 10-year historical prices from Yahoo Finance in robust batches...")
    
    closes_list = []
    opens_list = []
    highs_list = []
    lows_list = []
    volumes_list = []
    
    all_tickers_to_download = list(set(all_scan_tickers + [benchmark_ticker]))
    chunk_size = 100
    
    for i in range(0, len(all_tickers_to_download), chunk_size):
        chunk = all_tickers_to_download[i:i + chunk_size]
        print(f"   Downloading batch {i//chunk_size + 1} ({len(chunk)} symbols)...")
        try:
            chunk_raw = yf.download(chunk, period="10y", progress=False)
            if not chunk_raw.empty:
                if isinstance(chunk_raw.columns, pd.MultiIndex):
                    if 'Close' in chunk_raw: closes_list.append(chunk_raw['Close'])
                    if 'Open' in chunk_raw: opens_list.append(chunk_raw['Open'])
                    if 'High' in chunk_raw: highs_list.append(chunk_raw['High'])
                    if 'Low' in chunk_raw: lows_list.append(chunk_raw['Low'])
                    if 'Volume' in chunk_raw: volumes_list.append(chunk_raw['Volume'])
                else:
                    # Single ticker downloaded
                    ticker = chunk[0]
                    closes_list.append(pd.DataFrame({ticker: chunk_raw['Close']}))
                    opens_list.append(pd.DataFrame({ticker: chunk_raw['Open']}))
                    highs_list.append(pd.DataFrame({ticker: chunk_raw['High']}))
                    lows_list.append(pd.DataFrame({ticker: chunk_raw['Low']}))
                    volumes_list.append(pd.DataFrame({ticker: chunk_raw['Volume']}))
            time.sleep(0.5)
        except Exception as e:
            print(f"   ⚠️ Warning: Batch download encountered error: {e}. Retrying after brief pause...")
            time.sleep(2.0)
            try:
                chunk_raw = yf.download(chunk, period="10y", progress=False)
                if not chunk_raw.empty:
                    if isinstance(chunk_raw.columns, pd.MultiIndex):
                        if 'Close' in chunk_raw: closes_list.append(chunk_raw['Close'])
                        if 'Open' in chunk_raw: opens_list.append(chunk_raw['Open'])
                        if 'High' in chunk_raw: highs_list.append(chunk_raw['High'])
                        if 'Low' in chunk_raw: lows_list.append(chunk_raw['Low'])
                        if 'Volume' in chunk_raw: volumes_list.append(chunk_raw['Volume'])
            except Exception as retry_e:
                print(f"   ❌ Retry failed for this batch: {retry_e}")
                
    # Combine list DataFrames
    prices_raw = {
        'Close': pd.concat(closes_list, axis=1) if closes_list else pd.DataFrame(),
        'Open': pd.concat(opens_list, axis=1) if opens_list else pd.DataFrame(),
        'High': pd.concat(highs_list, axis=1) if highs_list else pd.DataFrame(),
        'Low': pd.concat(lows_list, axis=1) if lows_list else pd.DataFrame(),
        'Volume': pd.concat(volumes_list, axis=1) if volumes_list else pd.DataFrame()
    }
    
    if prices_raw['Close'].empty or benchmark_ticker not in prices_raw['Close'].columns:
        print("❌ Error: yfinance failed to retrieve price data.")
        return
        
    bench_raw = prices_raw['Close'][benchmark_ticker].dropna()
        
    # Build constituent stock price histories dict
    constituents_df_dict = {}
    for ticker in all_scan_tickers:
        if ticker in prices_raw['Close'].columns:
            closes = prices_raw['Close'][ticker].dropna()
            opens = prices_raw['Open'][ticker].dropna() if 'Open' in prices_raw else closes
            highs = prices_raw['High'][ticker].dropna() if 'High' in prices_raw else closes
            lows = prices_raw['Low'][ticker].dropna() if 'Low' in prices_raw else closes
            volumes = prices_raw['Volume'][ticker].dropna() if 'Volume' in prices_raw else pd.Series(0.0, index=closes.index)
            
            if not closes.empty:
                constituents_df_dict[ticker] = pd.DataFrame({
                    'Open': opens, 'High': highs, 'Low': lows, 'Close': closes, 'Volume': volumes
                }).reindex(closes.index).ffill()
                
    # 5.4 Synthesize Sector Index Price curves
    print("📊 Synthesizing capital-weighted Sector Price Indices...")
    sectors_data = {}
    for key, sec in sectors_config.items():
        synth_df = synthesize_sector_index(sec, constituents_df_dict)
        if not synth_df.empty:
            sectors_data[key] = synth_df
            
    print("✅ Sectors price histories synthesized.")

    # 5.5 Real-time NLP News Sentiment Engine
    print("📰 Fetching real-time Google News headlines & calculating NLP Sentiment scores...")
    sector_sentiment = {}
    news_feed_records = []
    
    for key, sec in sectors_config.items():
        label, score, feed = fetch_sector_news_sentiment(sec["name"])
        sector_sentiment[key] = {
            "label": label,
            "score": score
        }
        news_feed_records.extend(feed)
        print(f"   Sector: {sec['name']:<30} | Sentiment: {label:<10} (Score: {score:+.2f})")
        
    news_feed_df = pd.DataFrame(news_feed_records)

    # 5.6 Execute Multi-Timeframe Scans
    print("⚡ Executing Multi-Timeframe scans (EMA, RSI, RS Ratios, Multi-Year Extremes)...")
    daily_scan_rows = []
    weekly_scan_rows = []
    
    for sector_key, sec in sectors_config.items():
        sector_name = sec["name"]
        sector_close = sectors_data[sector_key]['Close']
        
        for ticker in sec["constituents"]:
            if ticker not in constituents_df_dict:
                continue
                
            df = constituents_df_dict[ticker]
            if len(df) < 252:
                continue
                
            ticker_clean = ticker.replace(".NS", "")
            
            # --- DAILY TIME-SERIES SCANS ---
            close = df['Close'].iloc[-1]
            ema20 = calculate_ema(df['Close'], 20).iloc[-1]
            ema50 = calculate_ema(df['Close'], 50).iloc[-1]
            ema100 = calculate_ema(df['Close'], 100).iloc[-1]
            ema200 = calculate_ema(df['Close'], 200).iloc[-1]
            rsi = calculate_rsi(df['Close'], 14).iloc[-1]
            
            # RS Ratios
            rs_nifty = calculate_rs_ratio(df['Close'], bench_raw, window=20)
            rs_sector = calculate_rs_ratio(df['Close'], sector_close, window=20)
            
            # Daily Delivery
            deliv = delivery_metrics.get(ticker, {"daily_delivery": 0.0, "avg_delivery": 0.0, "weekly_delivery": 0.0})
            
            # Extreme Bounds (Daily)
            high_52w = df['High'].rolling(min(252, len(df)), min_periods=min(100, len(df))).max().iloc[-1]
            low_52w = df['Low'].rolling(min(252, len(df)), min_periods=min(100, len(df))).min().iloc[-1]
            high_2y = df['High'].rolling(min(504, len(df)), min_periods=min(200, len(df))).max().iloc[-1]
            low_2y = df['Low'].rolling(min(504, len(df)), min_periods=min(200, len(df))).min().iloc[-1]
            high_5y = df['High'].rolling(min(1260, len(df)), min_periods=min(500, len(df))).max().iloc[-1]
            low_5y = df['Low'].rolling(min(1260, len(df)), min_periods=min(500, len(df))).min().iloc[-1]
            high_10y = df['High'].rolling(min(2520, len(df)), min_periods=min(1000, len(df))).max().iloc[-1]
            low_10y = df['Low'].rolling(min(2520, len(df)), min_periods=min(1000, len(df))).min().iloc[-1]
            
            # Daily signal criteria
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
                "RSI (14)": round(rsi, 1), "RS (Nifty 50)": round(rs_nifty, 4), "RS (Sector)": round(rs_sector, 4),
                "Daily Delivery %": deliv["daily_delivery"], "5D Avg Delivery %": deliv["avg_delivery"],
                "52W High": round(high_52w, 2), "52W Low": round(low_52w, 2),
                "2Y High": round(high_2y, 2), "2Y Low": round(low_2y, 2),
                "5Y High": round(high_5y, 2), "5Y Low": round(low_5y, 2),
                "10Y High": round(high_10y, 2), "10Y Low": round(low_10y, 2),
                "Signal": d_signal
            })
            
            # --- WEEKLY TIME-SERIES SCANS ---
            # Resample Stock Daily Dataframe to Weekly Close
            df_weekly = df.resample('W').agg({
                'Open': 'first', 'High': 'max', 'Low': 'min', 'Close': 'last', 'Volume': 'sum'
            }).dropna()
            
            bench_weekly = bench_raw.resample('W').last().dropna()
            sector_weekly_close = sector_close.resample('W').last().dropna()
            
            if len(df_weekly) >= 52:
                w_close = df_weekly['Close'].iloc[-1]
                w_ema20 = calculate_ema(df_weekly['Close'], 20).iloc[-1]
                w_ema50 = calculate_ema(df_weekly['Close'], 50).iloc[-1]
                w_ema100 = calculate_ema(df_weekly['Close'], 100).iloc[-1]
                w_ema200 = calculate_ema(df_weekly['Close'], 200).iloc[-1]
                w_rsi = calculate_rsi(df_weekly['Close'], 14).iloc[-1]
                
                w_rs_nifty = calculate_rs_ratio(df_weekly['Close'], bench_weekly, window=4)
                w_rs_sector = calculate_rs_ratio(df_weekly['Close'], sector_weekly_close, window=4)
                
                # Extreme Bounds (Weekly)
                w_high_52w = df_weekly['High'].rolling(min(52, len(df_weekly)), min_periods=min(20, len(df_weekly))).max().iloc[-1]
                w_low_52w = df_weekly['Low'].rolling(min(52, len(df_weekly)), min_periods=min(20, len(df_weekly))).min().iloc[-1]
                w_high_2y = df_weekly['High'].rolling(min(104, len(df_weekly)), min_periods=min(40, len(df_weekly))).max().iloc[-1]
                w_low_2y = df_weekly['Low'].rolling(min(104, len(df_weekly)), min_periods=min(40, len(df_weekly))).min().iloc[-1]
                w_high_5y = df_weekly['High'].rolling(min(260, len(df_weekly)), min_periods=min(100, len(df_weekly))).max().iloc[-1]
                w_low_5y = df_weekly['Low'].rolling(min(260, len(df_weekly)), min_periods=min(100, len(df_weekly))).min().iloc[-1]
                w_high_10y = df_weekly['High'].rolling(min(520, len(df_weekly)), min_periods=min(200, len(df_weekly))).max().iloc[-1]
                w_low_10y = df_weekly['Low'].rolling(min(520, len(df_weekly)), min_periods=min(200, len(df_weekly))).min().iloc[-1]
                
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
                    "RSI (14)": round(w_rsi, 1), "RS (Nifty 50)": round(w_rs_nifty, 4), "RS (Sector)": round(w_rs_sector, 4),
                    "Weekly Delivery %": deliv["weekly_delivery"],
                    "52W High": round(w_high_52w, 2), "52W Low": round(w_low_52w, 2),
                    "2Y High": round(w_high_2y, 2), "2Y Low": round(w_low_2y, 2),
                    "5Y High": round(w_high_5y, 2), "5Y Low": round(w_low_5y, 2),
                    "10Y High": round(w_high_10y, 2), "10Y Low": round(w_low_10y, 2),
                    "Signal": w_signal
                })
                
    daily_df = pd.DataFrame(daily_scan_rows).sort_values(by=["Signal", "RSI (14)"], ascending=[True, False])
    weekly_df = pd.DataFrame(weekly_scan_rows).sort_values(by=["Signal", "RSI (14)"], ascending=[True, False])
    
    # 5.7 Aggregated Sector Strength rankings & trails
    print("📊 Compiling final Sector Rankings strength index...")
    sector_summary = []
    rrg_trails_records = []
    
    for sector_key, sec in sectors_config.items():
        sector_name = sec["name"]
        group = daily_df[daily_df["Sector"] == sector_name]
        
        # Sector technical breadth
        breadth_20 = (group["Close"] > group["EMA 20"]).sum() / len(group) if len(group) > 0 else 0.0
        breadth_50 = (group["Close"] > group["EMA 50"]).sum() / len(group) if len(group) > 0 else 0.0
        breadth_200 = (group["Close"] > group["EMA 200"]).sum() / len(group) if len(group) > 0 else 0.0
        
        avg_rsi = group["RSI (14)"].mean() if len(group) > 0 else 50.0
        avg_delivery = group["5D Avg Delivery %"].mean() if len(group) > 0 else 0.0
        
        sent = sector_sentiment[sector_key]
        
        # Sector price trails for RRG
        sector_close = sectors_data[sector_key]['Close']
        ratio_series, mom_series = calculate_rrg_coordinates(sector_close, bench_raw)
        
        # Latest coordinates
        latest_ratio = ratio_series.iloc[-1]
        latest_mom = mom_series.iloc[-1]
        
        # Previous coordinates
        prev_ratio = ratio_series.iloc[-2] if len(ratio_series) > 1 else latest_ratio
        prev_mom = mom_series.iloc[-2] if len(mom_series) > 1 else latest_mom
        
        # Velocity and Heading compass vector
        velocity, heading = calculate_velocity_and_heading(latest_ratio, latest_mom, prev_ratio, prev_mom)
        
        # Quadrant classification
        quadrant = "LEADING" if latest_ratio >= 100 and latest_mom >= 100 else "WEAKENING" if latest_ratio >= 100 else "LAGGING" if latest_mom < 100 else "IMPROVING"
        
        # Rotational score (Multi-Factor Rank)
        heading_score = max(0.0, np.cos(np.radians(heading - 45))) * 100
        rotational_score = (
            (avg_rsi * 0.25) +
            ((breadth_50 * 100) * 0.30) +
            (heading_score * 0.25) +
            ((avg_delivery * 100) * 0.20)
        )
        
        sector_summary.append({
            "Sector Name": sector_name,
            "Avg Stock RSI": round(avg_rsi, 1),
            "20 EMA Breadth": round(breadth_20, 4),
            "50 EMA Breadth": round(breadth_50, 4),
            "200 EMA Breadth": round(breadth_200, 4),
            "RS-Ratio": round(latest_ratio, 2),
            "RS-Momentum": round(latest_mom, 2),
            "Velocity": round(velocity, 2),
            "Heading": round(heading, 1),
            "RRG Quadrant": quadrant,
            "Rotational Score": round(rotational_score, 1)
        })
        
        # Compile last 10 days of RRG trails for this sector
        last_10_dates = ratio_series.index[-10:]
        for dt in last_10_dates:
            r_val = ratio_series.loc[dt]
            m_val = mom_series.loc[dt]
            q_val = "LEADING" if r_val >= 100 and m_val >= 100 else "WEAKENING" if r_val >= 100 else "LAGGING" if m_val < 100 else "IMPROVING"
            
            rrg_trails_records.append({
                "Date": dt.strftime("%Y-%m-%d"),
                "Sector": sector_name,
                "RS-Ratio": round(r_val, 2),
                "RS-Momentum": round(m_val, 2),
                "RRG Quadrant": q_val
            })
            
    sector_rank_df = pd.DataFrame(sector_summary).sort_values(by="Rotational Score", ascending=False)
    rrg_trails_df = pd.DataFrame(rrg_trails_records)

    # -------------------------------------------------------------
    # 6. EXCEL COMPILER & CORPORATE STYLE SHEETS
    # -------------------------------------------------------------
    print(f"5. Writing output report to {OUTPUT_FILE} and applying styling...")
    
    if os.path.exists(OUTPUT_FILE):
        try:
            os.remove(OUTPUT_FILE)
        except:
            pass
            
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        daily_df.to_excel(writer, sheet_name="Daily Scanner", index=False)
        weekly_df.to_excel(writer, sheet_name="Weekly Scanner", index=False)
        sector_rank_df.to_excel(writer, sheet_name="Sector Rankings", index=False)
        rrg_trails_df.to_excel(writer, sheet_name="Sector RRG Trails", index=False)
        news_feed_df.to_excel(writer, sheet_name="Sector News Feed", index=False)
        
    # Styling layer
    wb = openpyxl.load_workbook(OUTPUT_FILE)
    
    # Font & Fills
    font_family = "Segoe UI"
    header_fill = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid") # Dark Corporate Navy Blue
    header_font = Font(name=font_family, size=11, bold=True, color="FFFFFF")
    cell_font = Font(name=font_family, size=11)
    
    # Custom fills
    buy_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid") # Soft green
    avoid_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid") # Soft red
    neutral_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid") # Soft yellow
    
    bold_green_font = Font(name=font_family, size=11, bold=True, color="006100")
    bold_red_font = Font(name=font_family, size=11, bold=True, color="9C0006")
    bold_yellow_font = Font(name=font_family, size=11, bold=True, color="7F6000")
    
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
    )
    
    for ws_name in wb.sheetnames:
        ws = wb[ws_name]
        # Force grid lines
        ws.views.sheetView[0].showGridLines = True
        
        # Style Header
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            
        ws.row_dimensions[1].height = 28
        
        # Style Cells
        for row_idx in range(2, ws.max_row + 1):
            ws.row_dimensions[row_idx].height = 20
            for col_idx in range(1, ws.max_column + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.font = cell_font
                cell.border = thin_border
                cell.alignment = Alignment(vertical="center")
                
                # Apply numbers alignments & percentage formats
                if ws_name in ["Daily Scanner", "Weekly Scanner"]:
                    # Clean stock numbers right
                    if col_idx in [3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21]:
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    # Format RS Percent outperformance columns (columns 9, 10)
                    if col_idx in [9, 10]:
                        cell.number_format = '+0.00%;-0.00%;0.00%'
                    # Format Delivery Columns (columns 11, 12 on daily, column 11 on weekly)
                    if ws_name == "Daily Scanner" and col_idx in [11, 12]:
                        cell.number_format = '0.00%'
                    if ws_name == "Weekly Scanner" and col_idx == 11:
                        cell.number_format = '0.00%'
                        
                    # Extract Close price (CMP) for pricing proximity checks
                    close_cell = ws.cell(row=row_idx, column=3)
                    try:
                        close_val = float(close_cell.value)
                    except:
                        close_val = 0.0
                        
                    # 1. Highlight EMAs (columns 4, 5, 6, 7) based on Price vs EMA relation
                    if col_idx in [4, 5, 6, 7]:
                        try:
                            ema_val = float(cell.value)
                            if close_val > ema_val:
                                cell.fill = buy_fill
                                cell.font = Font(name=font_family, size=11, color="006100")
                            elif close_val < ema_val:
                                cell.fill = avoid_fill
                                cell.font = Font(name=font_family, size=11, color="9C0006")
                        except:
                            pass
                            
                    # 2. Highlight price proximity to Multi-Year Highs (within 10% of high)
                    high_cols = [13, 15, 17, 19] if ws_name == "Daily Scanner" else [12, 14, 16, 18]
                    if col_idx in high_cols:
                        try:
                            high_val = float(cell.value)
                            if high_val > 0 and close_val >= 0.90 * high_val:
                                cell.fill = buy_fill
                                cell.font = Font(name=font_family, size=11, color="006100")
                        except:
                            pass
                            
                    # 3. Highlight price proximity to Multi-Year Lows (within 10% of low)
                    low_cols = [14, 16, 18, 20] if ws_name == "Daily Scanner" else [13, 15, 17, 19]
                    if col_idx in low_cols:
                        try:
                            low_val = float(cell.value)
                            if low_val > 0 and close_val <= 1.10 * low_val:
                                cell.fill = avoid_fill
                                cell.font = Font(name=font_family, size=11, color="9C0006")
                        except:
                            pass
                        
                    # Color-code Signal Column (last column)
                    if col_idx == ws.max_column:
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                        if cell.value == "STRONG BUY":
                            cell.fill = buy_fill
                            cell.font = bold_green_font
                        elif cell.value == "BUY":
                            cell.fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
                            cell.font = Font(name=font_family, size=11, color="375623")
                        elif cell.value == "AVOID":
                            cell.fill = avoid_fill
                            cell.font = bold_red_font
                            
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
                            cell.font = bold_green_font
                        elif cell.value == "IMPROVING":
                            cell.fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
                            cell.font = Font(name=font_family, size=11, bold=True, color="1F497D")
                        elif cell.value == "WEAKENING":
                            cell.fill = neutral_fill
                            cell.font = bold_yellow_font
                        elif cell.value == "LAGGING":
                            cell.fill = avoid_fill
                            cell.font = bold_red_font
                            
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
                            cell.font = bold_green_font
                        elif cell.value == "IMPROVING":
                            cell.fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
                            cell.font = Font(name=font_family, size=11, bold=True, color="1F497D")
                        elif cell.value == "WEAKENING":
                            cell.fill = neutral_fill
                            cell.font = bold_yellow_font
                        elif cell.value == "LAGGING":
                            cell.fill = avoid_fill
                            cell.font = bold_red_font
                            
                elif ws_name == "Sector News Feed":
                    if col_idx == 4:
                        cell.number_format = '+0.00;-0.00;0.00'
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    if col_idx == 5:
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                        if cell.value == "Bullish":
                            cell.fill = buy_fill
                            cell.font = bold_green_font
                        elif cell.value == "Bearish":
                            cell.fill = avoid_fill
                            cell.font = bold_red_font
                            
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
    
    # -------------------------------------------------------------
    # TELEGRAM BOT ALERT NOTIFICATION DISPATCHER
    # -------------------------------------------------------------
    try:
        # 1. Fetch latest data date
        dt_str = latest_bhav_date.strftime("%d-%b-%Y")
        
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
            f"📅 Date: {dt_str}\n\n"
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
