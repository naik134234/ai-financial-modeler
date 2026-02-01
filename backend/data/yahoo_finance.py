import logging
import requests
import json
from typing import Dict, Any, Optional, List, Union
from datetime import datetime, timedelta
import random
import re
import time

logger = logging.getLogger(__name__)

# Constants for Yahoo Finance API
BASE_URL = "https://query1.finance.yahoo.com/v10/finance/quoteSummary/"
CHART_URL = "https://query1.finance.yahoo.com/v8/finance/chart/"

# List of user agents to avoid rate limiting
USER_AGENTS = [
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36',
    'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/118.0.0.0 Safari/537.36',
]

# Common US stock exchanges - symbols from these don't need suffix
US_EXCHANGES = ['NYSE', 'NASDAQ', 'AMEX']

# Session manager for Yahoo Finance with crumb authentication
class YahooSession:
    _instance = None
    _session = None
    _crumb = None
    _last_refresh = None
    
    @classmethod
    def get_session(cls):
        """Get or create a session with valid crumb"""
        current_time = time.time()
        
        # Refresh session every 15 minutes
        if cls._session is None or cls._crumb is None or \
           (cls._last_refresh and current_time - cls._last_refresh > 900):
            cls._refresh_session()
        
        return cls._session, cls._crumb
    
    @classmethod
    def _refresh_session(cls):
        """Refresh the session and get a new crumb"""
        cls._session = requests.Session()
        headers = {
            'User-Agent': random.choice(USER_AGENTS),
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.5',
        }
        cls._session.headers.update(headers)
        
        try:
            # First, visit main page to get cookies
            cls._session.get('https://finance.yahoo.com', timeout=10)
            
            # Try to get crumb
            crumb_url = "https://query1.finance.yahoo.com/v1/test/getcrumb"
            response = cls._session.get(crumb_url, timeout=10)
            
            if response.status_code == 200:
                cls._crumb = response.text
                logger.info(f"Got Yahoo crumb: {cls._crumb[:10]}...")
            else:
                logger.warning(f"Failed to get crumb: {response.status_code}")
                cls._crumb = None
            
            cls._last_refresh = time.time()
        except Exception as e:
            logger.error(f"Failed to refresh Yahoo session: {e}")
            cls._crumb = None

def _is_us_stock(symbol: str) -> bool:
    """Check if symbol is likely a US stock (no suffix needed)"""
    # If already has a suffix like .NS, .BO, .L, etc., it's not a plain US stock
    if '.' in symbol:
        return False
    # US stocks are typically 1-5 uppercase letters without suffix
    # Common indicators: no numbers, all caps, short length
    return symbol.isalpha() and symbol.isupper() and len(symbol) <= 5

def _format_symbol(symbol: str, exchange: str = None) -> str:
    """Format symbol with appropriate suffix based on exchange"""
    # If already has a suffix, return as-is
    if '.' in symbol:
        return symbol
    
    # If exchange is specified as US-based, return without suffix
    if exchange and exchange.upper() in ['NYSE', 'NASDAQ', 'AMEX', 'US']:
        return symbol
    
    # If exchange is Indian, add appropriate suffix
    if exchange and exchange.upper() in ['NSE', 'BSE']:
        return f"{symbol}.NS" if exchange.upper() == 'NSE' else f"{symbol}.BO"
    
    # Try to auto-detect: if it looks like a US stock, don't add suffix
    # Otherwise default to Indian NSE
    if _is_us_stock(symbol):
        return symbol
    
    return f"{symbol}.NS"

def _get_headers():
    return {
        'User-Agent': random.choice(USER_AGENTS),
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
        'Accept-Language': 'en-US,en;q=0.5',
    }

def _make_yahoo_request(url: str, timeout: int = 15) -> Optional[requests.Response]:
    """Make a request to Yahoo Finance with session and crumb"""
    try:
        session, crumb = YahooSession.get_session()
        
        # Add crumb if available
        if crumb:
            separator = '&' if '?' in url else '?'
            url = f"{url}{separator}crumb={crumb}"
        
        response = session.get(url, timeout=timeout)
        
        # If unauthorized, refresh session and retry
        if response.status_code == 401:
            logger.info("Session expired, refreshing...")
            YahooSession._refresh_session()
            session, crumb = YahooSession.get_session()
            if crumb:
                # Rebuild URL with new crumb
                base_url = url.split('crumb=')[0].rstrip('&?')
                separator = '&' if '?' in base_url else '?'
                url = f"{base_url}{separator}crumb={crumb}"
            response = session.get(url, timeout=timeout)
        
        return response
    except Exception as e:
        logger.error(f"Yahoo request failed: {e}")
        # Fallback to simple request without session
        try:
            return requests.get(url.split('crumb=')[0].rstrip('&?'), 
                              headers=_get_headers(), timeout=timeout)
        except:
            return None

def _get_basic_info_from_chart(symbol: str, exchange: str = None) -> Optional[Dict[str, Any]]:
    """Fallback: Get basic stock info from Chart API which doesn't require auth"""
    try:
        formatted_symbol = _format_symbol(symbol, exchange)
        url = f"{CHART_URL}{formatted_symbol}?range=1d&interval=1d"
        
        response = requests.get(url, headers=_get_headers(), timeout=15)
        if response.status_code != 200:
            return None
        
        data = response.json()
        chart_result = data.get('chart', {}).get('result', [{}])
        if not chart_result:
            return None
        
        meta = chart_result[0].get('meta', {})
        
        # Get latest price
        indicators = chart_result[0].get('indicators', {})
        quote = indicators.get('quote', [{}])[0] if indicators else {}
        closes = quote.get('close', [])
        current_price = closes[-1] if closes else meta.get('regularMarketPrice', 0)
        
        return {
            "symbol": meta.get('symbol', symbol).replace('.NS', '').replace('.BO', ''),
            "name": meta.get('shortName', meta.get('longName', symbol)),
            "currency": meta.get('currency', 'USD'),
            "exchange": meta.get('exchangeName', exchange or 'Unknown'),
            "current_price": current_price,
            "previous_close": meta.get('chartPreviousClose', 0),
            "regular_market_price": meta.get('regularMarketPrice', current_price),
            # Mark as basic info only
            "_source": "chart_api"
        }
    except Exception as e:
        logger.error(f"Chart API fallback failed for {symbol}: {e}")
        return None

class YahooFinanceCollector:
    """Class to fetch data from Yahoo Finance without yfinance"""
    
    def __init__(self, symbol: str, exchange: str = None):
        self.symbol = symbol
        self.exchange = exchange
        # Use the new _format_symbol helper for proper US/Indian stock handling
        self.ticker_symbol = _format_symbol(symbol, exchange)
            
    def get_data(self) -> Dict[str, Any]:
        """Get all available data for the symbol"""
        # Pass exchange to helper functions for proper symbol formatting
        info = get_stock_info(self.symbol, self.exchange)
        financials = get_historical_financials(self.symbol, exchange=self.exchange)
        price_history = get_price_history(self.symbol, exchange=self.exchange)
        
        return {
            "company_info": info if info else {},
            "info": info if info else {}, 
            "financials": financials if financials else {},
            "price_history": price_history if price_history else {},
            "income_statement": financials.get('income_statement', {}) if financials else {},
            "balance_sheet": financials.get('balance_sheet', {}) if financials else {},
            "cash_flow": financials.get('cash_flow', {}) if financials else {},
        }

async def fetch_stock_data(symbol: str, exchange: str = "NSE") -> Dict[str, Any]:
    collector = YahooFinanceCollector(symbol, exchange)
    return collector.get_data()

def get_stock_info(symbol: str, exchange: str = None) -> Optional[Dict[str, Any]]:
    """Fetch basic stock info directly from Yahoo API
    
    Args:
        symbol: Stock symbol (e.g., AAPL for US, RELIANCE for India)
        exchange: Optional exchange hint (NYSE, NASDAQ, NSE, BSE)
    """
    original_symbol = symbol
    try:
        # Format symbol based on exchange or auto-detect
        formatted_symbol = _format_symbol(symbol, exchange)
        symbol = formatted_symbol
            
        modules = "financialData,quoteType,summaryDetail,price,defaultKeyStatistics,summaryProfile"
        url = f"{BASE_URL}{symbol}?modules={modules}"
        
        response = _make_yahoo_request(url, timeout=15)
        if not response or response.status_code != 200:
            logger.warning(f"Yahoo QuoteSummary API failed for {symbol}, trying chart API fallback...")
            # Try chart API fallback
            return _get_basic_info_from_chart(original_symbol, exchange)
        
        data = response.json()
        
        if 'quoteSummary' not in data or not data['quoteSummary']['result']:
            # Try chart API fallback
            logger.info(f"No quoteSummary data for {symbol}, trying chart API fallback...")
            return _get_basic_info_from_chart(original_symbol, exchange)
            
        result = data['quoteSummary']['result'][0]
        
        # Helper to safely get nested values
        def get_v(module, key, default=0):
            return result.get(module, {}).get(key, {}).get('raw', default)

        # Map to common structure
        info = {
            "symbol": symbol.replace('.NS', '').replace('.BO', ''),
            "name": result.get('price', {}).get('longName', symbol),
            "sector": result.get('summaryProfile', {}).get('sector', 'Unknown'),
            "industry": result.get('summaryProfile', {}).get('industry', 'Unknown'),
            "website": result.get('summaryProfile', {}).get('website', ''),
            "description": result.get('summaryProfile', {}).get('longBusinessSummary', ''),
            
            "market_cap": get_v('price', 'marketCap') / 10000000,
            "enterprise_value": get_v('defaultKeyStatistics', 'enterpriseValue') / 10000000,
            "current_price": get_v('financialData', 'currentPrice'),
            "52_week_high": get_v('summaryDetail', 'fiftyTwoWeekHigh'),
            "52_week_low": get_v('summaryDetail', 'fiftyTwoWeekLow'),
            "avg_volume": get_v('summaryDetail', 'averageVolume'),
            
            "shares_outstanding": get_v('defaultKeyStatistics', 'sharesOutstanding') / 10000000,
            "held_percent_insiders": get_v('defaultKeyStatistics', 'heldPercentInsiders'),
            "held_percent_institutions": get_v('defaultKeyStatistics', 'heldPercentInstitutions'),
            
            "pe_ratio": get_v('summaryDetail', 'trailingPE'),
            "forward_pe": get_v('summaryDetail', 'forwardPE'),
            "pb_ratio": get_v('defaultKeyStatistics', 'priceToBook'),
            "ps_ratio": get_v('summaryDetail', 'priceToSalesTrailing12Months'),
            
            "beta": get_v('defaultKeyStatistics', 'beta', 1.0),
            
            "profit_margin": get_v('financialData', 'profitMargins'),
            "operating_margin": get_v('financialData', 'operatingMargins'),
            "return_on_equity": get_v('financialData', 'returnOnEquity'),
            "return_on_assets": get_v('financialData', 'returnOnAssets'),
            
            "total_revenue": get_v('financialData', 'totalRevenue') / 10000000,
            "revenue_growth": get_v('financialData', 'revenueGrowth'),
            "ebitda": get_v('financialData', 'ebitda') / 10000000,
            "total_debt": get_v('financialData', 'totalDebt') / 10000000,
            "total_cash": get_v('financialData', 'totalCash') / 10000000,
            "free_cash_flow": get_v('financialData', 'freeCashflow') / 10000000,
            "earnings_per_share": get_v('defaultKeyStatistics', 'trailingEps'),
            
            "dividend_yield": get_v('summaryDetail', 'dividendYield'),
            "dividend_rate": get_v('summaryDetail', 'dividendRate'),
            "debt_to_equity": get_v('financialData', 'debtToEquity'),
            "current_ratio": get_v('financialData', 'currentRatio'),
        }
        return info
    except Exception as e:
        logger.error(f"Error in get_stock_info for {symbol}: {e}")
        # Try chart API fallback on exception
        return _get_basic_info_from_chart(original_symbol, exchange)

def get_historical_financials(symbol: str, years: int = 5, exchange: str = None) -> Optional[Dict[str, Any]]:
    """Fetch historical financials from Yahoo API
    
    Args:
        symbol: Stock symbol
        years: Number of years of data
        exchange: Optional exchange hint (NYSE, NASDAQ, NSE, BSE)
    """
    try:
        # Format symbol based on exchange or auto-detect
        formatted_symbol = _format_symbol(symbol, exchange)
        symbol = formatted_symbol

        modules = "incomeStatementHistory,balanceSheetHistory,cashflowStatementHistory"
        url = f"{BASE_URL}{symbol}?modules={modules}"
        
        response = _make_yahoo_request(url, timeout=15)
        if not response or response.status_code != 200:
            logger.warning(f"Yahoo API returned status {response.status_code if response else 'None'} for {symbol} financials")
            return None
        
        data = response.json()
        
        if 'quoteSummary' not in data or not data['quoteSummary']['result']:
            return None
            
        result = data['quoteSummary']['result'][0]
        
        def parse_statement(module_name):
            stmt_data = {}
            history = result.get(module_name, {}).get(module_name.replace('History', 'Statements'), [])
            for item in history:
                date = item.get('endDate', {}).get('fmt', '')[:4]
                if not date: continue
                
                vals = {}
                for k, v in item.items():
                    if isinstance(v, dict) and 'raw' in v:
                        vals[k] = v['raw'] / 10000000 # To Crores
                stmt_data[date] = vals
            return stmt_data

        income_stmt = parse_statement('incomeStatementHistory')
        balance_sheet = parse_statement('balanceSheetHistory')
        cash_flow = parse_statement('cashflowStatementHistory')
        
        # Normalize keys for the engine
        normalized_income = {}
        for yr, vals in income_stmt.items():
            normalized_income[yr] = {
                "revenue": vals.get('totalRevenue', 0),
                "gross_profit": vals.get('grossProfit', 0),
                "ebitda": vals.get('ebitda', 0),
                "operating_income": vals.get('operatingIncome', 0),
                "net_income": vals.get('netIncome', 0),
                "interest_expense": vals.get('interestExpense', 0),
                "tax_expense": vals.get('incomeTaxExpense', 0),
            }

        normalized_balance = {}
        for yr, vals in balance_sheet.items():
            normalized_balance[yr] = {
                "total_assets": vals.get('totalAssets', 0),
                "total_liabilities": vals.get('totalLiab', 0),
                "total_equity": vals.get('totalStockholderEquity', 0),
                "cash": vals.get('cash', 0),
                "total_debt": vals.get('longTermDebt', 0) + vals.get('shortLongTermDebt', 0),
                "current_assets": vals.get('totalCurrentAssets', 0),
                "current_liabilities": vals.get('totalCurrentLiabilities', 0),
            }

        normalized_cash = {}
        for yr, vals in cash_flow.items():
            normalized_cash[yr] = {
                "operating_cash_flow": vals.get('totalCashFromOperatingActivities', 0),
                "capex": abs(vals.get('capitalExpenditures', 0)),
                "depreciation": vals.get('depreciation', 0),
                "free_cash_flow": vals.get('totalCashFromOperatingActivities', 0) + vals.get('capitalExpenditures', 0),
            }

        return {
            "income_statement": normalized_income,
            "balance_sheet": normalized_balance,
            "cash_flow": normalized_cash,
            "years_available": len(normalized_income),
        }
    except Exception as e:
        logger.error(f"Error in get_historical_financials for {symbol}: {e}")
        return None

def get_price_history(symbol: str, period: str = "5y", exchange: str = None) -> Optional[Dict[str, Any]]:
    """Fetch price history using Chart API
    
    Args:
        symbol: Stock symbol
        period: Time period (1y, 2y, 5y, 10y, max)
        exchange: Optional exchange hint (NYSE, NASDAQ, NSE, BSE)
    """
    try:
        # Format symbol based on exchange or auto-detect
        formatted_symbol = _format_symbol(symbol, exchange)
        symbol = formatted_symbol
            
        # Map period to YF range
        range_map = {"1y": "1y", "2y": "2y", "5y": "5y", "10y": "10y", "max": "max"}
        r = range_map.get(period, "5y")
        
        url = f"{CHART_URL}{symbol}?range={r}&interval=1d"
        response = _make_yahoo_request(url, timeout=15)
        if not response or response.status_code != 200:
            logger.warning(f"Yahoo Chart API returned status {response.status_code if response else 'None'} for {symbol}")
            return None
        
        data = response.json()
        
        chart = data.get('chart', {}).get('result', [{}])[0]
        if not chart: return None
        
        closes = chart.get('indicators', {}).get('quote', [{}])[0].get('close', [])
        closes = [c for c in closes if c is not None]
        
        if not closes: return None
        
        returns = [(closes[i] / closes[i-1]) - 1 for i in range(1, len(closes))]
        avg_ret = sum(returns) / len(returns) if returns else 0
        std_ret = (sum([(x - avg_ret)**2 for x in returns]) / len(returns))**0.5 if returns else 0
        
        return {
            "current_price": closes[-1],
            "start_price": closes[0],
            "high": max(closes),
            "low": min(closes),
            "total_return": (closes[-1] / closes[0]) - 1,
            "annualized_return": avg_ret * 252,
            "volatility": std_ret * (252 ** 0.5),
            "sharpe_ratio": (avg_ret * 252 - 0.07) / (std_ret * (252 ** 0.5)) if std_ret > 0 else 0,
        }
    except Exception as e:
        logger.error(f"Error in get_price_history for {symbol}: {e}")
        return None

def get_peer_comparison(symbol: str, peers: Optional[List[str]] = None) -> Optional[List[Dict[str, Any]]]:
    """Fetch peer data using multiple API calls"""
    if not peers:
        peers = []
    
    results = []
    # Add main company
    main_info = get_stock_info(symbol)
    if main_info:
        results.append({
            "symbol": main_info["symbol"],
            "name": main_info["name"],
            "market_cap": main_info["market_cap"],
            "pe_ratio": main_info["pe_ratio"],
            "pb_ratio": main_info["pb_ratio"],
            "roe": main_info["return_on_equity"] * 100,
            "is_main": True,
        })
    
    for p in peers:
        p_info = get_stock_info(p)
        if p_info:
            results.append({
                "symbol": p_info["symbol"],
                "name": p_info["name"],
                "market_cap": p_info["market_cap"],
                "pe_ratio": p_info["pe_ratio"],
                "pb_ratio": p_info["pb_ratio"],
                "roe": p_info["return_on_equity"] * 100,
                "is_main": False,
            })
    return results

def search_stocks(query: str, limit: int = 10) -> List[Dict[str, str]]:
    """Simple offline search for common Indian stocks"""
    common_stocks = [
        {"symbol": "RELIANCE", "name": "Reliance Industries Ltd"},
        {"symbol": "TCS", "name": "Tata Consultancy Services"},
        {"symbol": "HDFCBANK", "name": "HDFC Bank Ltd"},
        {"symbol": "INFY", "name": "Infosys Ltd"},
        {"symbol": "ICICIBANK", "name": "ICICI Bank Ltd"},
        {"symbol": "HINDUNILVR", "name": "Hindustan Unilever Ltd"},
        {"symbol": "SBIN", "name": "State Bank of India"},
        {"symbol": "BHARTIARTL", "name": "Bharti Airtel Ltd"},
        {"symbol": "ITC", "name": "ITC Ltd"},
        {"symbol": "WIPRO", "name": "Wipro Ltd"},
        {"symbol": "ADANIPOWER", "name": "Adani Power"},
    ]
    query_lower = query.lower()
    return [s for s in common_stocks if query_lower in s["symbol"].lower() or query_lower in s["name"].lower()][:limit]


def _get_value(df, col, possible_keys: List[str]) -> float:
    """Helper to get value from dataframe with multiple possible keys"""
    for key in possible_keys:
        if key in df.index:
            val = df.loc[key, col]
            if val is not None and not (hasattr(val, '__iter__') and not isinstance(val, str)):
                return float(val) / 10000000  # Convert to Crores
    return 0.0


def search_stocks(query: str, limit: int = 10) -> List[Dict[str, str]]:
    """
    Search for stocks by name or symbol
    
    Args:
        query: Search query
        limit: Maximum results
    
    Returns:
        List of matching stocks
    """
    # For Yahoo Finance, we'd need a separate search API
    # This is a placeholder that returns common Indian stocks
    common_stocks = [
        {"symbol": "RELIANCE", "name": "Reliance Industries Ltd"},
        {"symbol": "TCS", "name": "Tata Consultancy Services"},
        {"symbol": "HDFCBANK", "name": "HDFC Bank Ltd"},
        {"symbol": "INFY", "name": "Infosys Ltd"},
        {"symbol": "ICICIBANK", "name": "ICICI Bank Ltd"},
        {"symbol": "HINDUNILVR", "name": "Hindustan Unilever Ltd"},
        {"symbol": "SBIN", "name": "State Bank of India"},
        {"symbol": "BHARTIARTL", "name": "Bharti Airtel Ltd"},
        {"symbol": "ITC", "name": "ITC Ltd"},
        {"symbol": "KOTAKBANK", "name": "Kotak Mahindra Bank"},
        {"symbol": "LT", "name": "Larsen & Toubro Ltd"},
        {"symbol": "AXISBANK", "name": "Axis Bank Ltd"},
        {"symbol": "WIPRO", "name": "Wipro Ltd"},
        {"symbol": "ASIANPAINT", "name": "Asian Paints Ltd"},
        {"symbol": "MARUTI", "name": "Maruti Suzuki India Ltd"},
        {"symbol": "TITAN", "name": "Titan Company Ltd"},
        {"symbol": "SUNPHARMA", "name": "Sun Pharmaceutical"},
        {"symbol": "ULTRACEMCO", "name": "UltraTech Cement Ltd"},
        {"symbol": "TATAMOTORS", "name": "Tata Motors Ltd"},
        {"symbol": "POWERGRID", "name": "Power Grid Corporation"},
    ]
    
    query_lower = query.lower()
    results = [
        stock for stock in common_stocks
        if query_lower in stock["symbol"].lower() or query_lower in stock["name"].lower()
    ]
    
    return results[:limit]
