"""
Yahoo Finance Data Fetcher
Real-time stock data and multi-year historical financials
"""

import logging
from typing import Dict, Any, Optional, List
from datetime import datetime, timedelta

logger = logging.getLogger(__name__)

# Try to import yfinance
try:
    import yfinance as yf
    YFINANCE_AVAILABLE = True
except ImportError:
    YFINANCE_AVAILABLE = False
    logger.warning("yfinance not installed. Run: pip install yfinance")


# Backward compatibility class
class YahooFinanceCollector:
    """Legacy class for backward compatibility"""
    
    def __init__(self, symbol: str, exchange: str = "NSE"):
        self.symbol = symbol
        self.exchange = exchange
        self.ticker_symbol = f"{symbol}.NS" if exchange == "NSE" else f"{symbol}.BO"
    
    def get_data(self) -> Dict[str, Any]:
        """Get all available data for the symbol"""
        info = get_stock_info(self.symbol)
        financials = get_historical_financials(self.symbol)
        price_history = get_price_history(self.symbol)
        
        # Use consistent key names that the generator expects
        return {
            "company_info": info if info else {},  # Generator expects company_info
            "info": info if info else {},  # Keep for backward compatibility
            "financials": financials if financials else {},
            "price_history": price_history if price_history else {},
            "income_statement": financials.get('income_statement', {}) if financials else {},
            "balance_sheet": financials.get('balance_sheet', {}) if financials else {},
            "cash_flow": financials.get('cash_flow', {}) if financials else {},
        }


async def fetch_stock_data(symbol: str, exchange: str = "NSE") -> Dict[str, Any]:
    """Fetch stock data from Yahoo Finance - async wrapper for backward compatibility"""
    collector = YahooFinanceCollector(symbol, exchange)
    return collector.get_data()


def get_stock_info(symbol: str) -> Optional[Dict[str, Any]]:
    """
    Get basic stock information
    
    Args:
        symbol: Stock symbol (e.g., 'RELIANCE.NS' for NSE stocks)
    
    Returns:
        Dictionary with stock info
    """
    if not YFINANCE_AVAILABLE:
        return None
    
    try:
        # Add .NS suffix for Indian stocks if not present
        if not symbol.endswith(('.NS', '.BO')):
            symbol = f"{symbol}.NS"
        
        ticker = yf.Ticker(symbol)
        info = ticker.info
        
        # Extract real company financial data
        shares_outstanding_raw = info.get('sharesOutstanding', 0) or 0
        shares_outstanding = shares_outstanding_raw / 10000000 if shares_outstanding_raw > 0 else 0  # Convert to Crores
        
        market_cap_raw = info.get('marketCap', 0) or 0
        market_cap = market_cap_raw / 10000000 if market_cap_raw > 0 else 0  # Convert to Crores
        
        enterprise_value_raw = info.get('enterpriseValue', 0) or 0
        enterprise_value = enterprise_value_raw / 10000000 if enterprise_value_raw > 0 else 0
        
        total_revenue_raw = info.get('totalRevenue', 0) or 0
        total_revenue = total_revenue_raw / 10000000 if total_revenue_raw > 0 else 0
        
        ebitda_raw = info.get('ebitda', 0) or 0
        ebitda = ebitda_raw / 10000000 if ebitda_raw > 0 else 0
        
        total_debt_raw = info.get('totalDebt', 0) or 0
        total_debt = total_debt_raw / 10000000 if total_debt_raw > 0 else 0
        
        total_cash_raw = info.get('totalCash', 0) or 0
        total_cash = total_cash_raw / 10000000 if total_cash_raw > 0 else 0
        
        free_cash_flow_raw = info.get('freeCashflow', 0) or 0
        free_cash_flow = free_cash_flow_raw / 10000000 if free_cash_flow_raw > 0 else 0
        
        net_income_raw = info.get('netIncomeToCommon', 0) or 0
        net_income = net_income_raw / 10000000 if net_income_raw else 0
        
        return {
            # Basic info
            "symbol": symbol.replace('.NS', '').replace('.BO', ''),
            "name": info.get('longName', info.get('shortName', symbol)),
            "sector": info.get('sector', 'Unknown'),
            "industry": info.get('industry', 'Unknown'),
            "website": info.get('website', ''),
            "description": info.get('longBusinessSummary', ''),
            
            # Market data (in Crores for Indian stocks)
            "market_cap": market_cap,
            "enterprise_value": enterprise_value,
            "current_price": info.get('currentPrice', info.get('regularMarketPrice', 0)),
            "52_week_high": info.get('fiftyTwoWeekHigh', 0),
            "52_week_low": info.get('fiftyTwoWeekLow', 0),
            "avg_volume": info.get('averageVolume', 0),
            
            # Shares and ownership (REAL DATA)
            "shares_outstanding": shares_outstanding,
            "shares_outstanding_raw": shares_outstanding_raw,  # Raw value for verification
            "float_shares": (info.get('floatShares', 0) or 0) / 10000000,
            "held_percent_insiders": info.get('heldPercentInsiders', 0),
            "held_percent_institutions": info.get('heldPercentInstitutions', 0),
            
            # Valuation ratios
            "pe_ratio": info.get('trailingPE', 0) or 0,
            "forward_pe": info.get('forwardPE', 0) or 0,
            "pb_ratio": info.get('priceToBook', 0) or 0,
            "ps_ratio": info.get('priceToSalesTrailing12Months', 0) or 0,
            "ev_to_revenue": info.get('enterpriseToRevenue', 0) or 0,
            "ev_to_ebitda": info.get('enterpriseToEbitda', 0) or 0,
            "peg_ratio": info.get('pegRatio', 0) or 0,
            
            # Risk metrics
            "beta": info.get('beta', 1.0) or 1.0,
            
            # Profitability (REAL DATA)
            "profit_margin": info.get('profitMargins', 0) or 0,
            "operating_margin": info.get('operatingMargins', 0) or 0,
            "gross_margin": info.get('grossMargins', 0) or 0,
            "ebitda_margin": (ebitda / total_revenue) if total_revenue > 0 else 0,
            "return_on_equity": info.get('returnOnEquity', 0) or 0,
            "return_on_assets": info.get('returnOnAssets', 0) or 0,
            
            # Financial data (in Crores - REAL DATA)
            "total_revenue": total_revenue,
            "revenue_growth": info.get('revenueGrowth', 0) or 0,
            "ebitda": ebitda,
            "net_income": net_income,
            "total_debt": total_debt,
            "total_cash": total_cash,
            "free_cash_flow": free_cash_flow,
            "book_value": (info.get('bookValue', 0) or 0),
            "earnings_per_share": info.get('trailingEps', 0) or 0,
            
            # Dividend data
            "dividend_yield": info.get('dividendYield', 0) or 0,
            "dividend_rate": info.get('dividendRate', 0) or 0,
            "payout_ratio": info.get('payoutRatio', 0) or 0,
            
            # Debt metrics
            "debt_to_equity": info.get('debtToEquity', 0) or 0,
            "current_ratio": info.get('currentRatio', 0) or 0,
            "quick_ratio": info.get('quickRatio', 0) or 0,
        }
    except Exception as e:
        logger.error(f"Error fetching stock info for {symbol}: {e}")
        return None


def get_historical_financials(symbol: str, years: int = 5) -> Optional[Dict[str, Any]]:
    """
    Get multi-year historical financial data
    
    Args:
        symbol: Stock symbol
        years: Number of years of historical data
    
    Returns:
        Dictionary with historical financials
    """
    if not YFINANCE_AVAILABLE:
        return None
    
    try:
        if not symbol.endswith(('.NS', '.BO')):
            symbol = f"{symbol}.NS"
        
        ticker = yf.Ticker(symbol)
        
        # Get financial statements
        income_stmt = ticker.income_stmt
        balance_sheet = ticker.balance_sheet
        cash_flow = ticker.cashflow
        
        # Process income statement
        income_data = {}
        if income_stmt is not None and not income_stmt.empty:
            for col in income_stmt.columns[:years]:
                year = col.year if hasattr(col, 'year') else str(col)[:4]
                income_data[str(year)] = {
                    "revenue": _get_value(income_stmt, col, ['Total Revenue', 'Revenue', 'Operating Revenue']),
                    "gross_profit": _get_value(income_stmt, col, ['Gross Profit']),
                    "ebitda": _get_value(income_stmt, col, ['EBITDA', 'Ebitda']),
                    "operating_income": _get_value(income_stmt, col, ['Operating Income', 'EBIT']),
                    "net_income": _get_value(income_stmt, col, ['Net Income', 'Net Income Common Stockholders']),
                    "interest_expense": _get_value(income_stmt, col, ['Interest Expense']),
                    "tax_expense": _get_value(income_stmt, col, ['Tax Provision', 'Income Tax Expense']),
                }
        
        # Process balance sheet
        balance_data = {}
        if balance_sheet is not None and not balance_sheet.empty:
            for col in balance_sheet.columns[:years]:
                year = col.year if hasattr(col, 'year') else str(col)[:4]
                balance_data[str(year)] = {
                    "total_assets": _get_value(balance_sheet, col, ['Total Assets']),
                    "total_liabilities": _get_value(balance_sheet, col, ['Total Liabilities Net Minority Interest', 'Total Liab']),
                    "total_equity": _get_value(balance_sheet, col, ['Total Equity Gross Minority Interest', 'Stockholders Equity']),
                    "cash": _get_value(balance_sheet, col, ['Cash And Cash Equivalents', 'Cash']),
                    "total_debt": _get_value(balance_sheet, col, ['Total Debt', 'Long Term Debt']),
                    "current_assets": _get_value(balance_sheet, col, ['Current Assets']),
                    "current_liabilities": _get_value(balance_sheet, col, ['Current Liabilities']),
                    "inventory": _get_value(balance_sheet, col, ['Inventory']),
                    "receivables": _get_value(balance_sheet, col, ['Net Receivables', 'Accounts Receivable']),
                    "payables": _get_value(balance_sheet, col, ['Accounts Payable']),
                }
        
        # Process cash flow
        cashflow_data = {}
        if cash_flow is not None and not cash_flow.empty:
            for col in cash_flow.columns[:years]:
                year = col.year if hasattr(col, 'year') else str(col)[:4]
                cashflow_data[str(year)] = {
                    "operating_cash_flow": _get_value(cash_flow, col, ['Operating Cash Flow', 'Total Cash From Operating Activities']),
                    "capex": abs(_get_value(cash_flow, col, ['Capital Expenditure', 'Capital Expenditures'])),
                    "depreciation": _get_value(cash_flow, col, ['Depreciation And Amortization', 'Depreciation']),
                    "free_cash_flow": _get_value(cash_flow, col, ['Free Cash Flow']),
                }
        
        return {
            "income_statement": income_data,
            "balance_sheet": balance_data,
            "cash_flow": cashflow_data,
            "years_available": len(income_data),
        }
        
    except Exception as e:
        logger.error(f"Error fetching historical financials for {symbol}: {e}")
        return None


def get_price_history(symbol: str, period: str = "5y") -> Optional[Dict[str, Any]]:
    """
    Get historical price data
    
    Args:
        symbol: Stock symbol
        period: Time period (1y, 2y, 5y, 10y, max)
    
    Returns:
        Dictionary with price history
    """
    if not YFINANCE_AVAILABLE:
        return None
    
    try:
        if not symbol.endswith(('.NS', '.BO')):
            symbol = f"{symbol}.NS"
        
        ticker = yf.Ticker(symbol)
        hist = ticker.history(period=period)
        
        if hist.empty:
            return None
        
        # Calculate returns
        returns = hist['Close'].pct_change().dropna()
        
        return {
            "current_price": float(hist['Close'].iloc[-1]),
            "start_price": float(hist['Close'].iloc[0]),
            "high": float(hist['High'].max()),
            "low": float(hist['Low'].min()),
            "avg_volume": float(hist['Volume'].mean()),
            "total_return": float((hist['Close'].iloc[-1] / hist['Close'].iloc[0]) - 1),
            "annualized_return": float(returns.mean() * 252),
            "volatility": float(returns.std() * (252 ** 0.5)),
            "sharpe_ratio": float((returns.mean() * 252 - 0.07) / (returns.std() * (252 ** 0.5))) if returns.std() > 0 else 0,
        }
        
    except Exception as e:
        logger.error(f"Error fetching price history for {symbol}: {e}")
        return None


def get_peer_comparison(symbol: str, peers: Optional[List[str]] = None) -> Optional[List[Dict[str, Any]]]:
    """
    Get comparison data for peer companies
    
    Args:
        symbol: Main stock symbol
        peers: List of peer symbols (auto-detected if not provided)
    
    Returns:
        List of dictionaries with peer comparison data
    """
    if not YFINANCE_AVAILABLE:
        return None
    
    try:
        if not symbol.endswith(('.NS', '.BO')):
            symbol = f"{symbol}.NS"
        
        # Get main stock info for sector
        main_ticker = yf.Ticker(symbol)
        main_info = main_ticker.info
        
        # If no peers provided, try to find them (limited without a database)
        if not peers:
            # Return just the main company for now
            peers = []
        
        results = []
        
        # Add main company
        results.append({
            "symbol": symbol.replace('.NS', ''),
            "name": main_info.get('shortName', symbol),
            "market_cap": main_info.get('marketCap', 0) / 10000000,
            "pe_ratio": main_info.get('trailingPE', 0),
            "pb_ratio": main_info.get('priceToBook', 0),
            "roe": main_info.get('returnOnEquity', 0) * 100 if main_info.get('returnOnEquity') else 0,
            "debt_equity": main_info.get('debtToEquity', 0) / 100 if main_info.get('debtToEquity') else 0,
            "is_main": True,
        })
        
        # Add peers
        for peer in peers:
            try:
                if not peer.endswith(('.NS', '.BO')):
                    peer = f"{peer}.NS"
                
                peer_ticker = yf.Ticker(peer)
                peer_info = peer_ticker.info
                
                results.append({
                    "symbol": peer.replace('.NS', ''),
                    "name": peer_info.get('shortName', peer),
                    "market_cap": peer_info.get('marketCap', 0) / 10000000,
                    "pe_ratio": peer_info.get('trailingPE', 0),
                    "pb_ratio": peer_info.get('priceToBook', 0),
                    "roe": peer_info.get('returnOnEquity', 0) * 100 if peer_info.get('returnOnEquity') else 0,
                    "debt_equity": peer_info.get('debtToEquity', 0) / 100 if peer_info.get('debtToEquity') else 0,
                    "is_main": False,
                })
            except Exception:
                continue
        
        return results
        
    except Exception as e:
        logger.error(f"Error fetching peer comparison for {symbol}: {e}")
        return None


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
