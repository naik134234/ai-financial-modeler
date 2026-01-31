"""
Yahoo Finance Data Collector
Fetches stock prices, financials, and key metrics for Indian stocks
"""

import yfinance as yf
import pandas as pd
from typing import Dict, Any, Optional
from datetime import datetime, timedelta
import logging

logger = logging.getLogger(__name__)


class YahooFinanceCollector:
    """Collect financial data from Yahoo Finance for Indian stocks"""
    
    # NSE/BSE suffix mapping
    EXCHANGE_SUFFIX = {
        'NSE': '.NS',
        'BSE': '.BO'
    }
    
    def __init__(self, symbol: str, exchange: str = 'NSE'):
        """
        Initialize with stock symbol
        
        Args:
            symbol: Stock symbol (e.g., 'ADANIPOWER', 'RELIANCE')
            exchange: 'NSE' or 'BSE'
        """
        self.base_symbol = symbol.upper().strip()
        self.exchange = exchange.upper()
        self.symbol = f"{self.base_symbol}{self.EXCHANGE_SUFFIX.get(self.exchange, '.NS')}"
        self.ticker = yf.Ticker(self.symbol)
        self._cache: Dict[str, Any] = {}
    
    def get_company_info(self) -> Dict[str, Any]:
        """Get basic company information"""
        if 'info' not in self._cache:
            try:
                info = self.ticker.info
                self._cache['info'] = {
                    'name': info.get('longName', info.get('shortName', self.base_symbol)),
                    'sector': info.get('sector', 'Unknown'),
                    'industry': info.get('industry', 'Unknown'),
                    'market_cap': info.get('marketCap', 0),
                    'currency': info.get('currency', 'INR'),
                    'website': info.get('website', ''),
                    'description': info.get('longBusinessSummary', ''),
                    'employees': info.get('fullTimeEmployees', 0),
                    'country': info.get('country', 'India'),
                }
            except Exception as e:
                logger.error(f"Error fetching company info: {e}")
                self._cache['info'] = {'name': self.base_symbol, 'sector': 'Unknown'}
        
        return self._cache['info']
    
    def get_historical_prices(self, period: str = '5y') -> pd.DataFrame:
        """
        Get historical stock prices
        
        Args:
            period: '1y', '2y', '5y', '10y', 'max'
        """
        try:
            df = self.ticker.history(period=period)
            df = df.reset_index()
            df['Date'] = pd.to_datetime(df['Date']).dt.date
            return df[['Date', 'Open', 'High', 'Low', 'Close', 'Volume']]
        except Exception as e:
            logger.error(f"Error fetching historical prices: {e}")
            return pd.DataFrame()
    
    def get_income_statement(self, quarterly: bool = False) -> pd.DataFrame:
        """Get income statement data"""
        try:
            if quarterly:
                return self.ticker.quarterly_income_stmt.T
            return self.ticker.income_stmt.T
        except Exception as e:
            logger.error(f"Error fetching income statement: {e}")
            return pd.DataFrame()
    
    def get_balance_sheet(self, quarterly: bool = False) -> pd.DataFrame:
        """Get balance sheet data"""
        try:
            if quarterly:
                return self.ticker.quarterly_balance_sheet.T
            return self.ticker.balance_sheet.T
        except Exception as e:
            logger.error(f"Error fetching balance sheet: {e}")
            return pd.DataFrame()
    
    def get_cash_flow(self, quarterly: bool = False) -> pd.DataFrame:
        """Get cash flow statement"""
        try:
            if quarterly:
                return self.ticker.quarterly_cashflow.T
            return self.ticker.cashflow.T
        except Exception as e:
            logger.error(f"Error fetching cash flow: {e}")
            return pd.DataFrame()
    
    def get_key_metrics(self) -> Dict[str, Any]:
        """Get key financial metrics"""
        try:
            info = self.ticker.info
            return {
                'pe_ratio': info.get('trailingPE'),
                'forward_pe': info.get('forwardPE'),
                'pb_ratio': info.get('priceToBook'),
                'ev_ebitda': info.get('enterpriseToEbitda'),
                'ev_revenue': info.get('enterpriseToRevenue'),
                'profit_margin': info.get('profitMargins'),
                'operating_margin': info.get('operatingMargins'),
                'roe': info.get('returnOnEquity'),
                'roa': info.get('returnOnAssets'),
                'debt_to_equity': info.get('debtToEquity'),
                'current_ratio': info.get('currentRatio'),
                'quick_ratio': info.get('quickRatio'),
                'dividend_yield': info.get('dividendYield'),
                'payout_ratio': info.get('payoutRatio'),
                'beta': info.get('beta'),
                'fifty_two_week_high': info.get('fiftyTwoWeekHigh'),
                'fifty_two_week_low': info.get('fiftyTwoWeekLow'),
                'current_price': info.get('currentPrice', info.get('regularMarketPrice')),
            }
        except Exception as e:
            logger.error(f"Error fetching key metrics: {e}")
            return {}
    
    def get_financials_summary(self) -> Dict[str, Any]:
        """Get comprehensive financial summary for modeling"""
        income = self.get_income_statement()
        balance = self.get_balance_sheet()
        cashflow = self.get_cash_flow()
        info = self.get_company_info()
        metrics = self.get_key_metrics()
        
        # Extract key figures for the model
        summary = {
            'company_info': info,
            'key_metrics': metrics,
            'income_statement': {},
            'balance_sheet': {},
            'cash_flow': {},
        }
        
        # Process income statement
        if not income.empty:
            latest = income.iloc[0] if len(income) > 0 else {}
            summary['income_statement'] = {
                'revenue': self._safe_get(latest, 'Total Revenue'),
                'cost_of_revenue': self._safe_get(latest, 'Cost Of Revenue'),
                'gross_profit': self._safe_get(latest, 'Gross Profit'),
                'operating_income': self._safe_get(latest, 'Operating Income'),
                'ebitda': self._safe_get(latest, 'EBITDA'),
                'net_income': self._safe_get(latest, 'Net Income'),
                'eps': self._safe_get(latest, 'Basic EPS'),
            }
            # Historical data for trends
            summary['income_statement']['historical'] = income.to_dict('records')
        
        # Process balance sheet
        if not balance.empty:
            latest = balance.iloc[0] if len(balance) > 0 else {}
            summary['balance_sheet'] = {
                'total_assets': self._safe_get(latest, 'Total Assets'),
                'total_liabilities': self._safe_get(latest, 'Total Liabilities Net Minority Interest'),
                'total_equity': self._safe_get(latest, 'Total Equity Gross Minority Interest'),
                'total_debt': self._safe_get(latest, 'Total Debt'),
                'cash': self._safe_get(latest, 'Cash And Cash Equivalents'),
                'current_assets': self._safe_get(latest, 'Current Assets'),
                'current_liabilities': self._safe_get(latest, 'Current Liabilities'),
                'inventory': self._safe_get(latest, 'Inventory'),
                'receivables': self._safe_get(latest, 'Accounts Receivable'),
                'payables': self._safe_get(latest, 'Accounts Payable'),
            }
            summary['balance_sheet']['historical'] = balance.to_dict('records')
        
        # Process cash flow
        if not cashflow.empty:
            latest = cashflow.iloc[0] if len(cashflow) > 0 else {}
            summary['cash_flow'] = {
                'operating_cash_flow': self._safe_get(latest, 'Operating Cash Flow'),
                'capital_expenditure': self._safe_get(latest, 'Capital Expenditure'),
                'free_cash_flow': self._safe_get(latest, 'Free Cash Flow'),
                'dividends_paid': self._safe_get(latest, 'Cash Dividends Paid'),
            }
            summary['cash_flow']['historical'] = cashflow.to_dict('records')
        
        return summary
    
    def _safe_get(self, data: Any, key: str, default: float = 0.0) -> float:
        """Safely get a value from data"""
        try:
            if isinstance(data, dict):
                val = data.get(key, default)
            elif hasattr(data, key):
                val = getattr(data, key, default)
            elif hasattr(data, 'get'):
                val = data.get(key, default)
            else:
                val = default
            
            if pd.isna(val):
                return default
            return float(val) if val is not None else default
        except:
            return default


def fetch_stock_data(symbol: str, exchange: str = 'NSE') -> Dict[str, Any]:
    """
    Convenience function to fetch all stock data
    
    Args:
        symbol: Stock symbol (e.g., 'ADANIPOWER')
        exchange: 'NSE' or 'BSE'
    
    Returns:
        Dictionary with all financial data
    """
    collector = YahooFinanceCollector(symbol, exchange)
    return collector.get_financials_summary()


if __name__ == "__main__":
    # Test with Adani Power
    data = fetch_stock_data("ADANIPOWER")
    print(f"Company: {data['company_info']['name']}")
    print(f"Sector: {data['company_info']['sector']}")
    print(f"Revenue: {data['income_statement'].get('revenue', 'N/A')}")
