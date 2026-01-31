"""
Alpha Vantage API Data Fetcher
Fetches real financial data using Alpha Vantage API
API Documentation: https://www.alphavantage.co/documentation/
"""

import requests
import logging
import os
from typing import Dict, Any, Optional, List

logger = logging.getLogger(__name__)

# Get API key from environment variable or use provided key
ALPHA_VANTAGE_API_KEY = os.environ.get('ALPHA_VANTAGE_API_KEY', 'LBE8AQPKX0SYIXWN')
BASE_URL = "https://www.alphavantage.co/query"


class AlphaVantageAPI:
    """Fetch financial data from Alpha Vantage API"""
    
    def __init__(self, api_key: str = None):
        self.api_key = api_key or ALPHA_VANTAGE_API_KEY
        if not self.api_key:
            raise ValueError("Alpha Vantage API key is required. Set ALPHA_VANTAGE_API_KEY environment variable.")
    
    def _make_request(self, function: str, symbol: str, **kwargs) -> Optional[Dict[str, Any]]:
        """Make API request to Alpha Vantage"""
        params = {
            'function': function,
            'symbol': symbol,
            'apikey': self.api_key,
            **kwargs
        }
        
        try:
            response = requests.get(BASE_URL, params=params, timeout=30)
            response.raise_for_status()
            data = response.json()
            
            # Check for API errors
            if 'Error Message' in data:
                logger.error(f"Alpha Vantage error: {data['Error Message']}")
                return None
            if 'Note' in data:
                logger.warning(f"Alpha Vantage rate limit: {data['Note']}")
                return None
            
            return data
        except requests.RequestException as e:
            logger.error(f"Alpha Vantage API request failed: {e}")
            return None
    
    def get_company_overview(self, symbol: str) -> Optional[Dict[str, Any]]:
        """
        Get company overview with fundamental data
        
        Returns: Market cap, PE ratio, EPS, Book Value, Dividend Yield, etc.
        """
        # For Indian stocks, append .BSE or .NSE
        av_symbol = f"{symbol}.BSE"
        
        data = self._make_request('OVERVIEW', av_symbol)
        if not data or 'Symbol' not in data:
            # Try NSE suffix
            av_symbol = f"{symbol}.NSE"
            data = self._make_request('OVERVIEW', av_symbol)
        
        if not data or 'Symbol' not in data:
            logger.warning(f"No overview data for {symbol}")
            return None
        
        return {
            'symbol': symbol,
            'name': data.get('Name', symbol),
            'description': data.get('Description', ''),
            'sector': data.get('Sector', 'Unknown'),
            'industry': data.get('Industry', 'Unknown'),
            'market_cap': self._parse_number(data.get('MarketCapitalization', 0)) / 10000000,  # Convert to Cr
            'pe_ratio': self._parse_number(data.get('PERatio', 0)),
            'peg_ratio': self._parse_number(data.get('PEGRatio', 0)),
            'book_value': self._parse_number(data.get('BookValue', 0)),
            'dividend_yield': self._parse_number(data.get('DividendYield', 0)),
            'eps': self._parse_number(data.get('EPS', 0)),
            'revenue_per_share': self._parse_number(data.get('RevenuePerShareTTM', 0)),
            'profit_margin': self._parse_number(data.get('ProfitMargin', 0)),
            'operating_margin': self._parse_number(data.get('OperatingMarginTTM', 0)),
            'return_on_assets': self._parse_number(data.get('ReturnOnAssetsTTM', 0)),
            'return_on_equity': self._parse_number(data.get('ReturnOnEquityTTM', 0)),
            'revenue': self._parse_number(data.get('RevenueTTM', 0)) / 10000000,  # Convert to Cr
            'gross_profit': self._parse_number(data.get('GrossProfitTTM', 0)) / 10000000,
            'ebitda': self._parse_number(data.get('EBITDA', 0)) / 10000000,
            'beta': self._parse_number(data.get('Beta', 1.0)),
            'shares_outstanding': self._parse_number(data.get('SharesOutstanding', 0)) / 10000000,  # Convert to Cr
            '52_week_high': self._parse_number(data.get('52WeekHigh', 0)),
            '52_week_low': self._parse_number(data.get('52WeekLow', 0)),
            'analyst_target_price': self._parse_number(data.get('AnalystTargetPrice', 0)),
            'forward_pe': self._parse_number(data.get('ForwardPE', 0)),
            'price_to_sales': self._parse_number(data.get('PriceToSalesRatioTTM', 0)),
            'price_to_book': self._parse_number(data.get('PriceToBookRatio', 0)),
            'ev_to_revenue': self._parse_number(data.get('EVToRevenue', 0)),
            'ev_to_ebitda': self._parse_number(data.get('EVToEBITDA', 0)),
        }
    
    def get_income_statement(self, symbol: str) -> Optional[Dict[str, Any]]:
        """Get income statement data (annual and quarterly)"""
        av_symbol = f"{symbol}.BSE"
        data = self._make_request('INCOME_STATEMENT', av_symbol)
        
        if not data or 'annualReports' not in data:
            av_symbol = f"{symbol}.NSE"
            data = self._make_request('INCOME_STATEMENT', av_symbol)
        
        if not data or 'annualReports' not in data:
            return None
        
        annual_reports = data.get('annualReports', [])
        quarterly_reports = data.get('quarterlyReports', [])
        
        # Get latest annual data
        latest = annual_reports[0] if annual_reports else {}
        
        return {
            'annual_reports': annual_reports,
            'quarterly_reports': quarterly_reports,
            'latest': {
                'revenue': self._parse_number(latest.get('totalRevenue', 0)) / 10000000,
                'gross_profit': self._parse_number(latest.get('grossProfit', 0)) / 10000000,
                'operating_income': self._parse_number(latest.get('operatingIncome', 0)) / 10000000,
                'net_income': self._parse_number(latest.get('netIncome', 0)) / 10000000,
                'ebitda': self._parse_number(latest.get('ebitda', 0)) / 10000000,
                'eps': self._parse_number(latest.get('eps', 0)),
            }
        }
    
    def get_balance_sheet(self, symbol: str) -> Optional[Dict[str, Any]]:
        """Get balance sheet data"""
        av_symbol = f"{symbol}.BSE"
        data = self._make_request('BALANCE_SHEET', av_symbol)
        
        if not data or 'annualReports' not in data:
            av_symbol = f"{symbol}.NSE"
            data = self._make_request('BALANCE_SHEET', av_symbol)
        
        if not data or 'annualReports' not in data:
            return None
        
        latest = data.get('annualReports', [{}])[0]
        
        return {
            'total_assets': self._parse_number(latest.get('totalAssets', 0)) / 10000000,
            'total_liabilities': self._parse_number(latest.get('totalLiabilities', 0)) / 10000000,
            'total_equity': self._parse_number(latest.get('totalShareholderEquity', 0)) / 10000000,
            'cash': self._parse_number(latest.get('cashAndCashEquivalentsAtCarryingValue', 0)) / 10000000,
            'total_debt': self._parse_number(latest.get('shortLongTermDebtTotal', 0)) / 10000000,
            'shares_outstanding': self._parse_number(latest.get('commonStockSharesOutstanding', 0)) / 10000000,
        }
    
    def get_cash_flow(self, symbol: str) -> Optional[Dict[str, Any]]:
        """Get cash flow statement data"""
        av_symbol = f"{symbol}.BSE"
        data = self._make_request('CASH_FLOW', av_symbol)
        
        if not data or 'annualReports' not in data:
            av_symbol = f"{symbol}.NSE"
            data = self._make_request('CASH_FLOW', av_symbol)
        
        if not data or 'annualReports' not in data:
            return None
        
        latest = data.get('annualReports', [{}])[0]
        
        return {
            'operating_cash_flow': self._parse_number(latest.get('operatingCashflow', 0)) / 10000000,
            'capital_expenditure': self._parse_number(latest.get('capitalExpenditures', 0)) / 10000000,
            'free_cash_flow': (self._parse_number(latest.get('operatingCashflow', 0)) - 
                              abs(self._parse_number(latest.get('capitalExpenditures', 0)))) / 10000000,
            'dividend_payout': self._parse_number(latest.get('dividendPayout', 0)) / 10000000,
        }
    
    def get_quote(self, symbol: str) -> Optional[Dict[str, Any]]:
        """Get real-time quote"""
        av_symbol = f"{symbol}.BSE"
        data = self._make_request('GLOBAL_QUOTE', av_symbol)
        
        if not data or 'Global Quote' not in data:
            av_symbol = f"{symbol}.NSE"
            data = self._make_request('GLOBAL_QUOTE', av_symbol)
        
        if not data or 'Global Quote' not in data:
            return None
        
        quote = data['Global Quote']
        return {
            'current_price': self._parse_number(quote.get('05. price', 0)),
            'open': self._parse_number(quote.get('02. open', 0)),
            'high': self._parse_number(quote.get('03. high', 0)),
            'low': self._parse_number(quote.get('04. low', 0)),
            'volume': self._parse_number(quote.get('06. volume', 0)),
            'previous_close': self._parse_number(quote.get('08. previous close', 0)),
            'change': self._parse_number(quote.get('09. change', 0)),
            'change_percent': quote.get('10. change percent', '0%').replace('%', ''),
        }
    
    def get_all_data(self, symbol: str) -> Dict[str, Any]:
        """Get all available financial data for a symbol"""
        logger.info(f"Fetching Alpha Vantage data for {symbol}")
        
        overview = self.get_company_overview(symbol) or {}
        income = self.get_income_statement(symbol) or {}
        balance = self.get_balance_sheet(symbol) or {}
        cash_flow = self.get_cash_flow(symbol) or {}
        quote = self.get_quote(symbol) or {}
        
        # Merge all data
        company_info = {
            **overview,
            **quote,
        }
        
        return {
            'company_info': company_info,
            'income_statement': income,
            'balance_sheet': balance,
            'cash_flow': cash_flow,
            'real_financials': income.get('latest', {}),
            'data_source': 'Alpha Vantage API',
        }
    
    @staticmethod
    def _parse_number(value) -> float:
        """Parse various number formats to float"""
        if value is None or value == 'None' or value == '-':
            return 0.0
        try:
            return float(value)
        except (ValueError, TypeError):
            return 0.0


# Convenience function
def fetch_alpha_vantage_data(symbol: str, api_key: str = None) -> Dict[str, Any]:
    """Fetch all financial data for a symbol from Alpha Vantage"""
    api = AlphaVantageAPI(api_key)
    return api.get_all_data(symbol)
