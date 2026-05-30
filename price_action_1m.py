#!/usr/bin/env python3
"""
BTCUSDT 1-Minute Price Action Scanner
Analyzes 1-minute candlestick data to detect structural trends, breakouts, 
and key price action reversal patterns (Engulfing, Hammers, Stars).
Generates precise trade entry signals.
"""

import urllib.request
import urllib.error
import json
import ssl
import pandas as pd
import numpy as np

# Bypass SSL certificate verification (needed on some macOS setups)
_SSL_CTX = ssl.create_default_context()
_SSL_CTX.check_hostname = False
_SSL_CTX.verify_mode = ssl.CERT_NONE

def fetch_1m_data(symbol="BTCUSDT", limit=100) -> pd.DataFrame:
    """Fetches last 100 candles of 1-minute Klines from Binance."""
    url = f"https://api.binance.com/api/v3/klines?symbol={symbol}&interval=1m&limit={limit}"
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, context=_SSL_CTX) as response:
            data = json.loads(response.read().decode())
        
        df = pd.DataFrame(data, columns=[
            'open_time', 'open', 'high', 'low', 'close', 'volume',
            'close_time', 'quote_asset_volume', 'number_of_trades',
            'taker_buy_base_asset_volume', 'taker_buy_quote_asset_volume', 'ignore'
        ])
        
        numeric_cols = ['open', 'high', 'low', 'close', 'volume']
        df[numeric_cols] = df[numeric_cols].astype(float)
        df['open_time'] = pd.to_datetime(df['open_time'], unit='ms')
        return df
    except Exception as e:
        print(f"Error fetching 1m Klines: {e}")
        return None

def analyze_price_action(df) -> dict:
    """
    Analyzes price action structure on the 1-minute chart:
    1. Trend bias (EMA 20 & EMA 50)
    2. Hammer / Shooting Star candles
    3. Engulfing patterns
    4. Support/Resistance Breakouts
    """
    if df is None or len(df) < 50:
        return {"signal": "WAIT", "reason": "Insufficient data"}
        
    # Calculate fast and slow EMAs
    df['ema_20'] = df['close'].ewm(span=20, adjust=False).mean()
    df['ema_50'] = df['close'].ewm(span=50, adjust=False).mean()
    
    # Get last 3 candles
    c0 = df.iloc[-1]  # Current forming/just closed candle
    c1 = df.iloc[-2]  # Previous candle
    c2 = df.iloc[-3]  # Two candles ago
    
    # General structural bias
    is_bullish_bias = c0['close'] > c0['ema_20'] and c0['ema_20'] > c0['ema_50']
    is_bearish_bias = c0['close'] < c0['ema_20'] and c0['ema_20'] < c0['ema_50']
    
    # Calculate body and shadow sizes for pattern recognition
    def candle_metrics(row):
        body = abs(row['close'] - row['open'])
        candle_range = row['high'] - row['low']
        upper_shadow = row['high'] - max(row['close'], row['open'])
        lower_shadow = min(row['close'], row['open']) - row['low']
        is_green = row['close'] > row['open']
        return body, candle_range, upper_shadow, lower_shadow, is_green

    b1, r1, u1, l1, g1 = candle_metrics(c1)
    
    # 1. HAMMER detection (Bullish reversal from low)
    # Body is small, lower shadow is at least 2x the body, very small upper shadow
    is_hammer = (l1 >= 2 * b1) and (u1 <= 0.2 * r1) and (b1 / (r1 + 1e-5) < 0.4)
    
    # 2. SHOOTING STAR detection (Bearish rejection of high)
    # Body is small, upper shadow is at least 2x the body, very small lower shadow
    is_shooting_star = (u1 >= 2 * b1) and (l1 <= 0.2 * r1) and (b1 / (r1 + 1e-5) < 0.4)
    
    # 3. ENGULFING patterns
    g2 = c2['close'] > c2['open']
    is_bullish_engulfing = (not g2) and g1 and (c1['close'] > c2['open']) and (c1['open'] < c2['close'])
    is_bearish_engulfing = g2 and (not g1) and (c1['close'] < c2['open']) and (c1['open'] > c2['close'])
    
    # 4. BREAKOUTS (Recent 20-period range breakout)
    recent_high = df['high'].iloc[-21:-1].max()
    recent_low = df['low'].iloc[-21:-1].min()
    
    is_bullish_breakout = c0['close'] > recent_high
    is_bearish_breakout = c0['close'] < recent_low
    
    # --- SIGNAL GENERATION DECISION LOGIC ---
    signal = "WAIT"
    reason = "Consolidation; no reliable pattern triggered."
    sl_price = 0.0
    tp_price = 0.0
    
    if is_bullish_bias:
        if is_bullish_breakout:
            signal = "BUY"
            reason = f"Bullish Breakout: Price closed at ${c0['close']:.2f} above the 20-minute local resistance of ${recent_high:.2f}."
        elif is_hammer and c1['low'] <= c1['ema_20']:
            signal = "BUY"
            reason = f"Trend Pullback Hammer: Bullish rejection candle at the EMA 20 during an uptrend."
        elif is_bullish_engulfing:
            signal = "BUY"
            reason = "Bullish Engulfing: Strong buying pressure engulfing the previous down candle."
            
        if signal == "BUY":
            # SL = local swing low of last 3 candles - tiny buffer
            sl_price = min(c0['low'], c1['low'], c2['low']) - 3.0
            # Protect against extreme tight SL
            if c0['close'] - sl_price < c0['close'] * 0.0005: 
                sl_price = c0['close'] * 0.999 # Minimum 0.1% stop loss
            risk = c0['close'] - sl_price
            tp_price = c0['close'] + (risk * 2.0) # 1:2 Risk-Reward
            
    elif is_bearish_bias:
        if is_bearish_breakout:
            signal = "SELL"
            reason = f"Bearish Breakout: Price closed at ${c0['close']:.2f} below the 20-minute local support of ${recent_low:.2f}."
        elif is_shooting_star and c1['high'] >= c1['ema_20']:
            signal = "SELL"
            reason = f"Trend Pullback Shooting Star: Bearish rejection candle at the EMA 20 during a downtrend."
        elif is_bearish_engulfing:
            signal = "SELL"
            reason = "Bearish Engulfing: Strong selling pressure engulfing the previous up candle."
            
        if signal == "SELL":
            # SL = local swing high of last 3 candles + tiny buffer
            sl_price = max(c0['high'], c1['high'], c2['high']) + 3.0
            # Protect against extreme tight SL
            if sl_price - c0['close'] < c0['close'] * 0.0005:
                sl_price = c0['close'] * 1.001 # Minimum 0.1% stop loss
            risk = sl_price - c0['close']
            tp_price = c0['close'] - (risk * 2.0) # 1:2 Risk-Reward
            
    return {
        "signal": signal,
        "reason": reason,
        "current_price": c0['close'],
        "sl_price": round(sl_price, 2),
        "tp_price": round(tp_price, 2),
        "ema20": c0['ema_20'],
        "ema50": c0['ema_50'],
        "volume": c0['volume'],
        "rsi_14": df['close'].diff().where(df['close'].diff() > 0, 0).rolling(14).mean().iloc[-1] # Simple raw RSI
    }

if __name__ == "__main__":
    df = fetch_1m_data("BTCUSDT", 100)
    if df is not None:
        analysis = analyze_price_action(df)
        print(json.dumps(analysis, indent=2))
