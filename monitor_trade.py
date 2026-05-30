#!/usr/bin/env python3
"""
Trade Monitoring & Decision Advisor
Loads active trade state, analyzes live market technicals, and provides 
actionable trade management advice (Take Profit, Cut Loss, Hold, or Market Entry guidance).
"""

import json
import urllib.request
import ssl
import sys
from typing import Dict, Any

# Bypass SSL certificate verification (needed on some macOS setups)
_SSL_CTX = ssl.create_default_context()
_SSL_CTX.check_hostname = False
_SSL_CTX.verify_mode = ssl.CERT_NONE

def get_live_price(symbol="BTCUSDT") -> float:
    """Fetches real-time price from Binance API."""
    url = f"https://api.binance.com/api/v3/ticker/price?symbol={symbol}"
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, context=_SSL_CTX) as response:
            data = json.loads(response.read().decode())
        return float(data['price'])
    except Exception as e:
        print(f"Error fetching live price: {e}", file=sys.stderr)
        return 0.0

def load_trade_state(file_path="trade_state.json") -> Dict[str, Any]:
    """Loads active trade parameters from local JSON."""
    try:
        with open(file_path, "r") as f:
            return json.load(f)
    except Exception as e:
        print(f"Error loading trade state: {e}", file=sys.stderr)
        return {}

def analyze_market_indicators() -> Dict[str, Any]:
    """Dynamically imports and queries the local TradingView server tools to get live indicators."""
    try:
        from tradingview_mcp_server import get_analysis, get_indicators
        
        # 1-Minute Indicators (Trend & Momentum)
        ta_1m = get_analysis("BTCUSDT", "BINANCE", "crypto", "1m")
        indicators_1m = get_indicators("BTCUSDT", "BINANCE", "crypto", "1m", ["RSI", "ADX", "EMA20", "EMA50"])
        
        return {
            "summary_1h": ta_1m.get("summary", {}).get("RECOMMENDATION", "NEUTRAL"),
            "ma_recommendation_1h": ta_1m.get("moving_averages", {}).get("recommendation", "NEUTRAL"),
            "rsi": indicators_1m.get("indicators", {}).get("RSI", 50.0),
            "adx": indicators_1m.get("indicators", {}).get("ADX", 20.0),
            "ema20": indicators_1m.get("indicators", {}).get("EMA20", 0.0),
            "ema50": indicators_1m.get("indicators", {}).get("EMA50", 0.0)
        }
    except Exception as e:
        # Fallback values if TradingView library fails
        return {
            "summary_1h": "NEUTRAL",
            "ma_recommendation_1h": "NEUTRAL",
            "rsi": 50.0,
            "adx": 20.0,
            "ema20": 0.0,
            "ema50": 0.0
        }

def monitor():
    state = load_trade_state()
    if not state:
        print("### ⚠️ Monitor Alert\nNo active trade state found. Place a trade first.")
        return
        
    if state.get("status") == "CLOSED":
        indicators = analyze_market_indicators()
        print("## 👁️ Live Trade Monitor Report")
        print(f"**Asset:** `{state.get('symbol', 'BTCUSDT')}` | **Position:** `NONE` | **State:** `FLAT` (Closed)\n")
        print(f"### 📊 Final PnL")
        print(f"*   **Realized PnL:** 🔴 **${state.get('realized_pnl', 0.0):.2f} USD**\n")
        print(f"### 🎯 Market Entry Guidance")
        if indicators["adx"] < 25:
            print("*   Market is in **Consolidation (No-Trade Zone)**. Avoid entering new trend-following trades due to high chop risk.")
        else:
            print(f"*   Market structure is `{indicators['summary_1h']}`. Good for trend strategies.")
        return
        
    live_price = get_live_price(state["symbol"])
    if live_price == 0.0:
        print("### ⚠️ Monitor Alert\nCould not retrieve live price.")
        return
        
    entry_price = state["entry_price"]
    qty = state["qty"]
    sl_price = state["sl_price"]
    tp_price = state["tp_price"]
    
    trade_type = state.get("trade_type", "LONG").upper()
    
    # Calculate Profit / Loss based on trade type
    if trade_type == "SHORT":
        pnl_usd = (entry_price - live_price) * qty
        pnl_pct = ((entry_price - live_price) / entry_price) * 100
    else:
        pnl_usd = (live_price - entry_price) * qty
        pnl_pct = ((live_price - entry_price) / entry_price) * 100
    
    # Get live indicators
    indicators = analyze_market_indicators()
    rsi = indicators["rsi"]
    adx = indicators["adx"]
    summary = indicators["summary_1h"]
    ma_rec = indicators["ma_recommendation_1h"]
    
    # Assess State & Generate Advice
    status = "HOLD"
    advice = "Hold position. Technical indicators are stable."
    urgency = "LOW"
    
    # Check absolute boundary triggers based on trade type
    is_tp_hit = (live_price <= tp_price) if trade_type == "SHORT" else (live_price >= tp_price)
    is_sl_hit = (live_price >= sl_price) if trade_type == "SHORT" else (live_price <= sl_price)
    
    if is_tp_hit:
        status = "TAKE PROFIT (TP TRIGGERED)"
        advice = f"Target hit! Close the position immediately to secure your profit of **${pnl_usd:.2f}**."
        urgency = "CRITICAL"
    elif is_sl_hit:
        status = "STOP LOSS (SL TRIGGERED)"
        advice = f"Stop Loss hit! Cut the position immediately to limit loss to **${pnl_usd:.2f}**."
        urgency = "CRITICAL"
    else:
        # Check indicator warnings
        if rsi > 70:
            status = "CAUTION"
            advice = "RSI is entering overbought territory (>70). Consider locking in partial profits or raising your Stop Loss." if trade_type == "LONG" else "RSI is overbought (>70). Momentum is favorable for your short, but watch for a reversal."
            urgency = "MEDIUM"
        elif rsi < 30:
            status = "CAUTION"
            advice = "RSI is oversold (<30). Monitor support closely; a bounce is expected but momentum is weak." if trade_type == "LONG" else "RSI is entering oversold territory (<30). Consider locking in profits or lowering your Stop Loss."
            urgency = "MEDIUM"
        elif trade_type == "LONG" and ma_rec == "SELL" and live_price < indicators["ema50"]:
            status = "WARNING"
            advice = "Price has broken below the 1-Minute EMA50. Short-term momentum is turning bearish; consider closing early if you wish to minimize risk."
            urgency = "HIGH"
        elif trade_type == "SHORT" and ma_rec == "BUY" and live_price > indicators["ema50"]:
            status = "WARNING"
            advice = "Price has reclaimed the 1-Minute EMA50. Short-term momentum is turning bullish; consider closing early if you wish to minimize risk."
            urgency = "HIGH"
            
    # Evaluate Market Entry / Trade vs No-Trade zone
    new_trade_advice = "No action"
    if adx < 25:
        new_trade_advice = "Market is in **Consolidation (No-Trade Zone)**. Avoid entering new trend-following trades due to high chop risk."
    else:
        if summary in ["BUY", "STRONG_BUY"]:
            new_trade_advice = "Market is trending **Bullish**. Good environment for long pullbacks."
        elif summary in ["SELL", "STRONG_SELL"]:
            new_trade_advice = "Market is trending **Bearish**. Good environment for short pullbacks."
            
    # Print Markdown Output
    print(f"## 👁️ Live Trade Monitor Report")
    print(f"**Asset:** `{state['symbol']}` | **Position:** `{trade_type}` | **Quantity:** `{qty} BTC`\n")
    
    print(f"### 📊 Trade Status & PnL")
    pnl_sign = "+" if pnl_usd >= 0 else ""
    pnl_color = "🟢" if pnl_usd >= 0 else "🔴"
    print(f"*   **Current Price:** `${live_price:,.2f}`")
    print(f"*   **Entry Price:** `${entry_price:,.2f}`")
    print(f"*   **PnL:** {pnl_color} **{pnl_sign}${pnl_usd:.2f}** ({pnl_sign}{pnl_pct:.2f}%)")
    print(f"*   **Stop Loss:** `${sl_price:,.2f}` | **Take Profit:** `${tp_price:,.2f}`\n")
    
    print(f"### 📈 Technical Metrics (1-Minute)")
    print(f"*   **1M Recommendation:** `{summary}` (Moving Averages: `{ma_rec}`)")
    print(f"*   **RSI (14):** `{rsi:.2f}` | **ADX (Trend Strength):** `{adx:.2f}`\n")
    
    print(f"### 💡 Actionable Decision & Advice")
    alert_emoji = "🚨" if urgency in ["HIGH", "CRITICAL"] else "⚠️" if urgency == "MEDIUM" else "ℹ️"
    print(f"> {alert_emoji} **DECISION: {status}** (Urgency: `{urgency}`)")
    print(f"> {advice}\n")
    
    print(f"### 🎯 Market Entry Guidance")
    print(f"*   {new_trade_advice}")

if __name__ == "__main__":
    monitor()
