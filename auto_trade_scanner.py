#!/usr/bin/env python3
"""
Automated 1-Minute Price Action Trade Scanner
Runs in the background, checks Klines, calculates price action signals,
and writes any valid trade entry setups to a JSON trigger file.
"""

import json
import os
import sys
import pandas as pd
from price_action_1m import fetch_1m_data, analyze_price_action

SIGNAL_FILE = "active_signal.json"

def main():
    # Clear any stale signals
    if os.path.exists(SIGNAL_FILE):
        try:
            os.remove(SIGNAL_FILE)
        except Exception:
            pass

    # Fetch and analyze data
    df = fetch_1m_data("BTCUSDT", 100)
    if df is None:
        print("Error: Could not retrieve 1m data.")
        sys.exit(1)
        
    analysis = analyze_price_action(df)
    signal = analysis.get("signal", "WAIT")
    
    print(f"Current Price Action Scan: {signal} at ${analysis.get('current_price', 0.0):,.2f}")
    print(f"Reason: {analysis.get('reason', 'N/A')}")
    
    if signal in ["BUY", "SELL"]:
        print(f"\n🚨 [SIGNAL_TRIGGERED] {signal} setup detected!")
        # Save signal parameters to trigger file
        with open(SIGNAL_FILE, "w") as f:
            json.dump({
                "symbol": "BTCUSDT",
                "signal": signal,
                "entry_price": analysis.get("current_price"),
                "sl_price": analysis.get("sl_price"),
                "tp_price": analysis.get("tp_price"),
                "reason": analysis.get("reason"),
                "timestamp": pd.Timestamp.now().isoformat()
            }, f, indent=2)
        # Force a non-zero exit code or print keyword to trigger the main agent notification
        sys.exit(2)
        
    sys.exit(0)

if __name__ == "__main__":
    main()
