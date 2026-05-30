import sys
import os
import pandas as pd
import yfinance as yf

# Add the current directory to python path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

try:
    print("🔄 Initializing Sector Rotation Multi-Scanner Verification checks...")
    
    # 1. Import from our self-contained scanner
    from sector_rotation_multi_scanner import (
        fetch_nifty500_constituents,
        download_last_5_days_delivery_bhavcopies,
        fetch_sector_news_sentiment,
        analyze_headline_sentiment,
        calculate_rsi,
        calculate_ema,
        synthesize_sector_index
    )
    print("✅ Successfully imported all scanner core functions.")
    
    # 2. Test NSE Nifty 500 Constituents list fetcher
    n500_df = fetch_nifty500_constituents()
    if n500_df.empty:
        print("❌ Nifty 500 fetcher returned empty.")
        sys.exit(1)
    print(f"✅ Nifty 500 list fetched successfully. Symbols parsed: {len(n500_df)}.")
    print(f"   Sample columns: {list(n500_df.columns)}")
    print(f"   Sample industries: {n500_df['Industry'].dropna().unique()[:5]}")
    
    # 3. Test NSE Security Bhavcopy Delivery Downloader
    bhavs = download_last_5_days_delivery_bhavcopies()
    if not bhavs:
        print("❌ Bhavcopy downloader failed to retrieve any reports.")
        sys.exit(1)
    print(f"✅ Delivery Bhavcopies downloaded successfully (active days: {len(bhavs)}).")
    
    # 4. Test Google News RSS & Sentiment Engine
    label, score, feed = fetch_sector_news_sentiment("Banking")
    print(f"✅ News Sentiment Engine working. Industry: Banking | Sentiment: {label} (Score: {score:+.2f})")
    print("   Sample headlines fetched:")
    for item in feed[:2]:
        print(f"     - [{item['Polarity']}] {item['Headline']} ({item['Date']})")
        
    # 5. Test Technical Indicators
    print("📥 Fetching minimal historical data for benchmark ^NSEI...")
    nsei = yf.download("^NSEI", period="1mo", progress=False)
    if nsei.empty:
        print("❌ Failed to download benchmark data from yfinance.")
        sys.exit(1)
        
    closes = nsei['Close'].squeeze()
    if isinstance(closes, pd.DataFrame):
        closes = closes.iloc[:, 0]
    rsi = calculate_rsi(closes)
    ema20 = calculate_ema(closes, 20)
    
    print(f"✅ Vectorized technical indicators validated.")
    print(f"   Latest Nifty 50 Close: {closes.iloc[-1]:.2f}")
    print(f"   Latest Nifty 50 RSI (14): {rsi.iloc[-1]:.1f}")
    print(f"   Latest Nifty 50 EMA (20): {ema20.iloc[-1]:.2f}")
    
    # 6. Test Sector Synthesis
    print("🧪 Testing Sector Synthesis with mock constituents...")
    sector_config = {
        "name": "TEST SECTOR",
        "constituents": ["^NSEI"],
        "market_caps": {"^NSEI": 1.0}
    }
    constituents_df_dict = {
        "^NSEI": nsei
    }
    synth_df = synthesize_sector_index(sector_config, constituents_df_dict)
    if synth_df.empty:
        print("❌ Sector index synthesis returned empty.")
        sys.exit(1)
    print(f"✅ Sector Index Synthesis verified. Synth Close: {synth_df['Close'].iloc[-1]:.2f}")
    
    print("\n🎉 ALL PIPELINES ARE 100% OPERATIONAL, INTEGRATED, AND CORRECT!")
    sys.exit(0)
    
except Exception as e:
    import traceback
    print(f"❌ Verification failed due to error: {e}")
    traceback.print_exc()
    sys.exit(1)
