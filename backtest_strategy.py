import urllib.request
import json
import pandas as pd
import numpy as np

def fetch_binance_data(symbol="BTCUSDT", interval="15m", limit=1000):
    """Fetches historical candlestick data from Binance API using standard library urllib."""
    url = f"https://api.binance.com/api/v3/klines?symbol={symbol}&interval={interval}&limit={limit}"
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req) as response:
            data = json.loads(response.read().decode())
        
        # Columns: Open time, Open, High, Low, Close, Volume, Close time, etc.
        df = pd.DataFrame(data, columns=[
            'open_time', 'open', 'high', 'low', 'close', 'volume',
            'close_time', 'quote_asset_volume', 'number_of_trades',
            'taker_buy_base_asset_volume', 'taker_buy_quote_asset_volume', 'ignore'
        ])
        
        # Convert numeric columns to float
        numeric_cols = ['open', 'high', 'low', 'close', 'volume']
        df[numeric_cols] = df[numeric_cols].astype(float)
        df['open_time'] = pd.to_datetime(df['open_time'], unit='ms')
        return df
    except Exception as e:
        print(f"Error fetching data from Binance: {e}")
        return None

def calculate_indicators(df):
    """Calculates EMAs and RSI on the dataframe."""
    # 15m EMA 50 (Pullback zone)
    df['ema_50'] = df['close'].ewm(span=50, adjust=False).mean()
    
    # 15m EMA 800 (Equivalent to 1-Hour EMA 200 trend filter)
    df['ema_800'] = df['close'].ewm(span=800, adjust=False).mean()
    
    # Relative Strength Index (RSI 14)
    delta = df['close'].diff()
    gain = (delta.where(delta > 0, 0)).copy()
    loss = (-delta.where(delta < 0, 0)).copy()
    
    avg_gain = gain.rolling(window=14).mean()
    avg_loss = loss.rolling(window=14).mean()
    
    # Smooth RSI
    for i in range(14, len(df)):
        avg_gain.iloc[i] = (avg_gain.iloc[i-1] * 13 + gain.iloc[i]) / 14
        avg_loss.iloc[i] = (avg_loss.iloc[i-1] * 13 + loss.iloc[i]) / 14
        
    rs = avg_gain / avg_loss
    df['rsi'] = 100 - (100 / (1 + rs))
    return df

def run_backtest(df, sl_pct=0.005, rr_ratio=2.0):
    """
    Simulates a high-probability pullback strategy:
    - Long Setup:
      1. Price above EMA 800 (Bullish structural bias)
      2. Price pulls back to touch or go below EMA 50
      3. RSI 14 is oversold (< 38)
    - Short Setup:
      1. Price below EMA 800 (Bearish structural bias)
      2. Price pulls back to touch or go above EMA 50
      3. RSI 14 is overbought (> 62)
    """
    tp_pct = sl_pct * rr_ratio
    trades = []
    in_trade = False
    trade_type = None  # 'LONG' or 'SHORT'
    entry_price = 0
    sl_price = 0
    tp_price = 0
    entry_time = None
    
    # We skip first 800 candles to allow EMA 800 to stabilize
    for idx in range(800, len(df)):
        row = df.iloc[idx]
        
        if not in_trade:
            # --- LONG ENTRY CONDITIONS ---
            is_bullish_trend = row['close'] > row['ema_800']
            is_pullback = row['low'] <= row['ema_50']
            is_rsi_oversold = row['rsi'] < 38
            
            if is_bullish_trend and is_pullback and is_rsi_oversold:
                in_trade = True
                trade_type = 'LONG'
                entry_price = row['close']
                sl_price = entry_price * (1.0 - sl_pct)
                tp_price = entry_price * (1.0 + tp_pct)
                entry_time = row['open_time']
                continue
                
            # --- SHORT ENTRY CONDITIONS ---
            is_bearish_trend = row['close'] < row['ema_800']
            is_pullback_short = row['high'] >= row['ema_50']
            is_rsi_overbought = row['rsi'] > 62
            
            if is_bearish_trend and is_pullback_short and is_rsi_overbought:
                in_trade = True
                trade_type = 'SHORT'
                entry_price = row['close']
                sl_price = entry_price * (1.0 + sl_pct)
                tp_price = entry_price * (1.0 - tp_pct)
                entry_time = row['open_time']
                continue
                
        else:
            # --- MANAGE ACTIVE TRADE ---
            if trade_type == 'LONG':
                # Check Stop Loss
                if row['low'] <= sl_price:
                    trades.append({
                        'type': 'LONG',
                        'entry_time': entry_time,
                        'exit_time': row['open_time'],
                        'entry_price': entry_price,
                        'exit_price': sl_price,
                        'result': 'LOSE',
                        'return': -sl_pct
                    })
                    in_trade = False
                # Check Take Profit
                elif row['high'] >= tp_price:
                    trades.append({
                        'type': 'LONG',
                        'entry_time': entry_time,
                        'exit_time': row['open_time'],
                        'entry_price': entry_price,
                        'exit_price': tp_price,
                        'result': 'WIN',
                        'return': tp_pct
                    })
                    in_trade = False
                    
            elif trade_type == 'SHORT':
                # Check Stop Loss
                if row['high'] >= sl_price:
                    trades.append({
                        'type': 'SHORT',
                        'entry_time': entry_time,
                        'exit_time': row['open_time'],
                        'entry_price': entry_price,
                        'exit_price': sl_price,
                        'result': 'LOSE',
                        'return': -sl_pct
                    })
                    in_trade = False
                # Check Take Profit
                elif row['low'] <= tp_price:
                    trades.append({
                        'type': 'SHORT',
                        'entry_time': entry_time,
                        'exit_time': row['open_time'],
                        'entry_price': entry_price,
                        'exit_price': tp_price,
                        'result': 'WIN',
                        'return': tp_pct
                    })
                    in_trade = False
                    
    return trades

def print_performance_metrics(trades):
    """Helper to calculate and print detailed performance metrics."""
    if not trades:
        print("No trades triggered in the backtested historical window.")
        return
        
    trades_df = pd.DataFrame(trades)
    total_trades = len(trades_df)
    wins = len(trades_df[trades_df['result'] == 'WIN'])
    losses = len(trades_df[trades_df['result'] == 'LOSE'])
    win_rate = (wins / total_trades) * 100 if total_trades > 0 else 0
    
    total_return = trades_df['return'].sum() * 100
    profit_factor = abs(trades_df[trades_df['return'] > 0]['return'].sum() / 
                        trades_df[trades_df['return'] < 0]['return'].sum()) if losses > 0 else float('inf')
    
    print("\n" + "="*50)
    print(" 📈 HIGH-PROBABILITY PULLBACK BACKTEST RESULTS")
    print("="*50)
    print(f"Total Trades Triggered : {total_trades}")
    print(f"Wins                   : {wins} ✅")
    print(f"Losses                 : {losses} ❌")
    print(f"Calculated Win Rate    : {win_rate:.2f}%")
    print(f"Risk-to-Reward Ratio   : 1:2.0")
    print(f"Stop Loss (SL)         : 0.50%")
    print(f"Take Profit (TP)       : 1.00%")
    print(f"Profit Factor          : {profit_factor:.2f}")
    print(f"Net Cumulative Return  : {total_return:.2f}%")
    print("="*50)

if __name__ == "__main__":
    print("Fetching last 1000 candles (15-minute timeframe) for BTCUSDT...")
    df = fetch_binance_data("BTCUSDT", "15m", 1000)
    if df is not None:
        df = calculate_indicators(df)
        trades = run_backtest(df, sl_pct=0.005, rr_ratio=2.0)
        print_performance_metrics(trades)
