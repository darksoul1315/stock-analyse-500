#!/usr/bin/env python3
"""
TradingView MCP Server
Provides real-time technical analysis data, moving averages, oscillators,
and individual indicators using the Model Context Protocol (MCP) and tradingview-ta.
"""

import sys
import logging
from typing import Dict, List, Any, Optional
from fastmcp import FastMCP
from tradingview_ta import TA_Handler, Interval, Exchange

# Configure logging
logging.basicConfig(level=logging.INFO, stream=sys.stderr)
logger = logging.getLogger("tradingview_mcp")

# Initialize FastMCP Server
mcp = FastMCP("TradingView")

# Dictionary to map interval strings to tradingview_ta Interval objects
INTERVAL_MAP = {
    "1m": Interval.INTERVAL_1_MINUTE,
    "5m": Interval.INTERVAL_5_MINUTES,
    "15m": Interval.INTERVAL_15_MINUTES,
    "30m": Interval.INTERVAL_30_MINUTES,
    "1h": Interval.INTERVAL_1_HOUR,
    "2h": Interval.INTERVAL_2_HOURS,
    "4h": Interval.INTERVAL_4_HOURS,
    "1d": Interval.INTERVAL_1_DAY,
    "1w": Interval.INTERVAL_1_WEEK,
    "1M": Interval.INTERVAL_1_MONTH,
}

# Helper mapping for common exchange/screener guesses
ASSET_SUGGESTIONS = {
    # Cryptocurrencies
    "BTC": {"exchange": "BINANCE", "screener": "crypto"},
    "ETH": {"exchange": "BINANCE", "screener": "crypto"},
    "SOL": {"exchange": "BINANCE", "screener": "crypto"},
    "ADA": {"exchange": "BINANCE", "screener": "crypto"},
    "DOT": {"exchange": "BINANCE", "screener": "crypto"},
    "DOGE": {"exchange": "BINANCE", "screener": "crypto"},
    
    # US Stocks
    "AAPL": {"exchange": "NASDAQ", "screener": "america"},
    "MSFT": {"exchange": "NASDAQ", "screener": "america"},
    "GOOGL": {"exchange": "NASDAQ", "screener": "america"},
    "AMZN": {"exchange": "NASDAQ", "screener": "america"},
    "NVDA": {"exchange": "NASDAQ", "screener": "america"},
    "TSLA": {"exchange": "NASDAQ", "screener": "america"},
    "META": {"exchange": "NASDAQ", "screener": "america"},
    "SPY": {"exchange": "AMEX", "screener": "america"},
    "QQQ": {"exchange": "NASDAQ", "screener": "america"},
    
    # Indian Stocks (NSE)
    "RELIANCE": {"exchange": "NSE", "screener": "india"},
    "TCS": {"exchange": "NSE", "screener": "india"},
    "INFY": {"exchange": "NSE", "screener": "india"},
    "HDFCBANK": {"exchange": "NSE", "screener": "india"},
    "ICICIBANK": {"exchange": "NSE", "screener": "india"},
    "NIFTY": {"exchange": "NSE", "screener": "india"},
    "BANKNIFTY": {"exchange": "NSE", "screener": "india"},
}

def parse_interval(interval_str: str) -> str:
    """Parses user interval string and returns tradingview-ta compatible interval."""
    # Normalize interval input
    clean_val = interval_str.strip().lower()
    if clean_val == "1day" or clean_val == "daily" or clean_val == "d":
        return "1d"
    if clean_val == "1week" or clean_val == "weekly" or clean_val == "w":
        return "1w"
    if clean_val == "1month" or clean_val == "monthly" or clean_val == "m":
        return "1M"
    if clean_val == "30m" or clean_val == "30min":
        return "30m"
    if clean_val == "15m" or clean_val == "15min":
        return "15m"
    if clean_val == "5m" or clean_val == "5min":
        return "5m"
    if clean_val == "1m" or clean_val == "1min":
        return "1m"
    if clean_val == "1h" or clean_val == "1hour" or clean_val == "hourly":
        return "1h"
    if clean_val == "4h" or clean_val == "4hour":
        return "4h"
    if clean_val == "2h" or clean_val == "2hour":
        return "2h"
    return clean_val

def get_ta_handler(symbol: str, exchange: str, screener: str, interval: str) -> TA_Handler:
    """Helper to initialize the tradingview-ta handler."""
    parsed_int = parse_interval(interval)
    tv_interval = INTERVAL_MAP.get(parsed_int, Interval.INTERVAL_1_DAY)
    
    return TA_Handler(
        symbol=symbol.upper(),
        exchange=exchange.upper(),
        screener=screener.lower(),
        interval=tv_interval
    )

@mcp.tool
def get_analysis(
    symbol: str,
    exchange: str,
    screener: str = "america",
    interval: str = "1d"
) -> Dict[str, Any]:
    """
    Get full technical analysis summary for a specific symbol on TradingView.
    
    Args:
        symbol: The stock, crypto, or index ticker symbol (e.g., TSLA, BTCUSDT, RELIANCE, NIFTY).
        exchange: The listing exchange (e.g., NASDAQ, NYSE, AMEX, BINANCE, NSE, BSE).
        screener: TradingView screener region or market type. Defaults to 'america'. Use 'crypto' for cryptocurrencies, 'india' for Indian stock market.
        interval: Timeframe for calculation. Options: '1m', '5m', '15m', '30m', '1h', '2h', '4h', '1d' (default), '1w', '1M'.
    """
    try:
        handler = get_ta_handler(symbol, exchange, screener, interval)
        analysis = handler.get_analysis()
        
        # Prepare clean return structure
        result = {
            "symbol": symbol.upper(),
            "exchange": exchange.upper(),
            "screener": screener.lower(),
            "interval": interval,
            "summary": analysis.summary,
            "moving_averages": {
                "recommendation": analysis.moving_averages.get("RECOMMENDATION"),
                "buy": analysis.moving_averages.get("BUY"),
                "sell": analysis.moving_averages.get("SELL"),
                "neutral": analysis.moving_averages.get("NEUTRAL"),
            },
            "oscillators": {
                "recommendation": analysis.oscillators.get("RECOMMENDATION"),
                "buy": analysis.oscillators.get("BUY"),
                "sell": analysis.oscillators.get("SELL"),
                "neutral": analysis.oscillators.get("NEUTRAL"),
            },
            "time": analysis.time.isoformat() if hasattr(analysis.time, "isoformat") else str(analysis.time)
        }
        return result
    except Exception as e:
        logger.error(f"Error fetching analysis for {symbol}: {str(e)}")
        return {
            "error": str(e),
            "tip": f"Ensure the ticker '{symbol}' is valid on the exchange '{exchange}' under the '{screener}' screener."
        }

@mcp.tool
def get_indicators(
    symbol: str,
    exchange: str,
    screener: str = "america",
    interval: str = "1d",
    indicators_list: Optional[List[str]] = None
) -> Dict[str, Any]:
    """
    Fetch specific indicator values (e.g. RSI, MACD, EMA20, SMA200, Bollinger Bands) for a symbol.
    
    Args:
        symbol: The stock, crypto, or index ticker symbol (e.g., AAPL, BTCUSDT, RELIANCE).
        exchange: The listing exchange (e.g., NASDAQ, BINANCE, NSE).
        screener: TradingView screener type ('america', 'crypto', 'india', etc.). Defaults to 'america'.
        interval: Timeframe ('1m', '5m', '15m', '1h', '1d', etc.). Defaults to '1d'.
        indicators_list: List of specific indicators to retrieve (e.g., ["RSI", "MACD.macd", "EMA20"]). If omitted, returns all 90+ standard indicators.
    """
    try:
        handler = get_ta_handler(symbol, exchange, screener, interval)
        analysis = handler.get_analysis()
        
        all_indicators = analysis.indicators
        
        # If requested a subset, filter them
        if indicators_list:
            filtered = {}
            for name in indicators_list:
                # Support case-insensitive key lookup
                matched_key = next((k for k in all_indicators if k.upper() == name.upper()), None)
                if matched_key:
                    filtered[matched_key] = all_indicators[matched_key]
                else:
                    # Try partial match or exact
                    filtered[name] = "Indicator not found or not computed by TV TA."
            indicators_data = filtered
        else:
            indicators_data = all_indicators
            
        return {
            "symbol": symbol.upper(),
            "exchange": exchange.upper(),
            "interval": interval,
            "indicators": indicators_data
        }
    except Exception as e:
        logger.error(f"Error fetching indicators for {symbol}: {str(e)}")
        return {"error": str(e)}

@mcp.tool
def get_bulk_analysis(
    symbols: List[str],
    exchange: str,
    screener: str = "america",
    interval: str = "1d"
) -> Dict[str, Any]:
    """
    Fetch technical analysis summaries for a list of symbols on the same exchange concurrently.
    
    Args:
        symbols: List of symbols (e.g., ["AAPL", "MSFT", "GOOGL"]).
        exchange: The listing exchange (e.g., NASDAQ, NYSE, NSE, BINANCE).
        screener: TradingView screener type ('america', 'crypto', 'india', etc.). Defaults to 'america'.
        interval: Timeframe ('1m', '5m', '15m', '1h', '1d', etc.). Defaults to '1d'.
    """
    results = {}
    for sym in symbols:
        sym = sym.strip().upper()
        try:
            handler = get_ta_handler(sym, exchange, screener, interval)
            analysis = handler.get_analysis()
            results[sym] = {
                "recommendation": analysis.summary.get("RECOMMENDATION"),
                "summary": analysis.summary,
                "status": "success"
            }
        except Exception as e:
            results[sym] = {
                "status": "error",
                "error": str(e)
            }
    return {
        "exchange": exchange.upper(),
        "screener": screener,
        "interval": interval,
        "results": results
    }

@mcp.tool
def search_symbols(query: str) -> Dict[str, Any]:
    """
    Helper to search/lookup exchange and screener mapping for common assets.
    
    Args:
        query: Ticker name or asset query (e.g., 'AAPL', 'Bitcoin', 'Nifty', 'Reliance').
    """
    query_upper = query.upper().strip()
    
    # Exact check
    if query_upper in ASSET_SUGGESTIONS:
        return {
            "found": True,
            "query": query,
            "best_match": {
                "symbol": query_upper,
                **ASSET_SUGGESTIONS[query_upper]
            }
        }
        
    # Heuristics based matches
    matches = []
    
    # Check if ends with USDT, BTC, ETH -> likely crypto/BINANCE
    if any(query_upper.endswith(suffix) for suffix in ["USDT", "BTC", "ETH", "USD"]):
        base_sym = query_upper
        matches.append({
            "symbol": base_sym,
            "exchange": "BINANCE",
            "screener": "crypto",
            "confidence": "high (crypto ending match)"
        })
        
    # Check for Indian stocks or index keywords
    if any(k in query_upper for k in ["NIFTY", "RELIANCE", "TCS", "INFY", "NSE", "BSE"]):
        matches.append({
            "symbol": query_upper.replace("NSE:", "").replace("BSE:", ""),
            "exchange": "NSE",
            "screener": "india",
            "confidence": "high (Indian market keyword)"
        })

    # Standard suggestions list scan
    for key, val in ASSET_SUGGESTIONS.items():
        if query_upper in key or key in query_upper:
            matches.append({
                "symbol": key,
                "exchange": val["exchange"],
                "screener": val["screener"],
                "confidence": "medium (substring match)"
            })
            
    if matches:
        return {
            "found": True,
            "query": query,
            "matches": matches,
            "message": "Verify the exchange and symbol match what you need."
        }
        
    return {
        "found": False,
        "query": query,
        "message": "No direct matching suggestions. Commonly: \n"
                   "- For US Equities: screener='america', exchange='NASDAQ'/'NYSE'\n"
                   "- For Crypto: screener='crypto', exchange='BINANCE'/'COINBASE'\n"
                   "- For Indian Equities: screener='india', exchange='NSE'/'BSE'"
    }

if __name__ == "__main__":
    # Start the fastmcp stdio server
    mcp.run()
