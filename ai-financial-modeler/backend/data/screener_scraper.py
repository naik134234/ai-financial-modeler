"""
Screener.in Web Scraper
Extracts detailed Indian stock financials from Screener.in
"""

import requests
from bs4 import BeautifulSoup
import pandas as pd
from typing import Dict, Any, List, Optional
import logging
import re
import time

logger = logging.getLogger(__name__)


class ScreenerScraper:
    """Scrape financial data from Screener.in"""
    
    BASE_URL = "https://www.screener.in/company"
    
    HEADERS = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
        'Accept-Language': 'en-US,en;q=0.5',
        'Connection': 'keep-alive',
    }
    
    def __init__(self, symbol: str):
        """
        Initialize with stock symbol
        
        Args:
            symbol: Stock symbol (e.g., 'ADANIPOWER', 'RELIANCE')
        """
        self.symbol = symbol.upper().strip()
        self.url = f"{self.BASE_URL}/{self.symbol}/"
        self._soup: Optional[BeautifulSoup] = None
        self._data: Dict[str, Any] = {}
    
    def _fetch_page(self) -> bool:
        """Fetch and parse the company page"""
        if self._soup is not None:
            return True
        
        try:
            response = requests.get(self.url, headers=self.HEADERS, timeout=15)
            response.raise_for_status()
            self._soup = BeautifulSoup(response.content, 'lxml')
            return True
        except requests.RequestException as e:
            logger.error(f"Failed to fetch Screener page for {self.symbol}: {e}")
            return False
    
    def get_company_info(self) -> Dict[str, Any]:
        """Get basic company information"""
        if not self._fetch_page():
            return {'name': self.symbol}
        
        info = {'symbol': self.symbol}
        
        try:
            # Company name
            name_elem = self._soup.select_one('h1.margin-0')
            if name_elem:
                info['name'] = name_elem.get_text(strip=True)
            
            # Sector and industry from warehouse data
            warehouse = self._soup.select_one('#warehouse-data')
            if warehouse:
                sector_elem = warehouse.select_one('[data-sector]')
                if sector_elem:
                    info['sector'] = sector_elem.get('data-sector', 'Unknown')
            
            # Ratios from the top section
            ratios = self._soup.select('.company-ratios li')
            for ratio in ratios:
                name_elem = ratio.select_one('.name')
                value_elem = ratio.select_one('.value')
                if name_elem and value_elem:
                    name = name_elem.get_text(strip=True).lower().replace(' ', '_')
                    value = self._parse_number(value_elem.get_text(strip=True))
                    info[name] = value
            
        except Exception as e:
            logger.error(f"Error parsing company info: {e}")
        
        return info
    
    def get_quarterly_results(self) -> List[Dict[str, Any]]:
        """Get quarterly financial results"""
        if not self._fetch_page():
            return []
        
        return self._parse_financial_table('quarters')
    
    def get_annual_results(self) -> List[Dict[str, Any]]:
        """Get annual profit & loss statement"""
        if not self._fetch_page():
            return []
        
        return self._parse_financial_table('profit-loss')
    
    def get_balance_sheet(self) -> List[Dict[str, Any]]:
        """Get balance sheet data"""
        if not self._fetch_page():
            return []
        
        return self._parse_financial_table('balance-sheet')
    
    def get_cash_flow(self) -> List[Dict[str, Any]]:
        """Get cash flow statement"""
        if not self._fetch_page():
            return []
        
        return self._parse_financial_table('cash-flow')
    
    def get_ratios(self) -> List[Dict[str, Any]]:
        """Get financial ratios over time"""
        if not self._fetch_page():
            return []
        
        return self._parse_financial_table('ratios')
    
    def get_shareholding(self) -> Dict[str, Any]:
        """Get shareholding pattern"""
        if not self._fetch_page():
            return {}
        
        shareholding = {}
        try:
            section = self._soup.select_one('#shareholding')
            if section:
                rows = section.select('table tr')
                for row in rows:
                    cells = row.select('td')
                    if len(cells) >= 2:
                        name = cells[0].get_text(strip=True)
                        value = self._parse_number(cells[-1].get_text(strip=True))
                        shareholding[name] = value
        except Exception as e:
            logger.error(f"Error parsing shareholding: {e}")
        
        return shareholding
    
    def get_peers(self) -> List[Dict[str, Any]]:
        """Get peer comparison data"""
        if not self._fetch_page():
            return []
        
        peers = []
        try:
            section = self._soup.select_one('#peers')
            if section:
                table = section.select_one('table')
                if table:
                    headers = [th.get_text(strip=True) for th in table.select('thead th')]
                    for row in table.select('tbody tr'):
                        cells = row.select('td')
                        if len(cells) == len(headers):
                            peer = {}
                            for i, cell in enumerate(cells):
                                if i == 0:
                                    link = cell.select_one('a')
                                    peer['name'] = link.get_text(strip=True) if link else cell.get_text(strip=True)
                                else:
                                    peer[headers[i]] = self._parse_number(cell.get_text(strip=True))
                            peers.append(peer)
        except Exception as e:
            logger.error(f"Error parsing peers: {e}")
        
        return peers
    
    def _parse_financial_table(self, section_id: str) -> List[Dict[str, Any]]:
        """Parse a financial data table from Screener"""
        try:
            section = self._soup.select_one(f'#{section_id}')
            if not section:
                return []
            
            table = section.select_one('table')
            if not table:
                return []
            
            # Get headers (years/quarters)
            headers = []
            header_row = table.select_one('thead tr')
            if header_row:
                headers = [th.get_text(strip=True) for th in header_row.select('th')]
            
            if not headers:
                return []
            
            # Parse data rows
            data = []
            for row in table.select('tbody tr'):
                cells = row.select('td')
                if not cells:
                    continue
                
                row_data = {}
                row_name = cells[0].get_text(strip=True) if cells else ''
                
                for i, cell in enumerate(cells):
                    if i == 0:
                        row_data['metric'] = cell.get_text(strip=True)
                    elif i < len(headers):
                        value = self._parse_number(cell.get_text(strip=True))
                        row_data[headers[i]] = value
                
                if row_data.get('metric'):
                    data.append(row_data)
            
            return data
            
        except Exception as e:
            logger.error(f"Error parsing {section_id} table: {e}")
            return []
    
    def _parse_number(self, value: str) -> Optional[float]:
        """Parse a number string, handling Cr, Lakh, %, etc."""
        if not value or value == '-':
            return None
        
        try:
            # Remove commas and whitespace
            value = value.replace(',', '').strip()
            
            # Handle percentages
            if '%' in value:
                return float(value.replace('%', '')) / 100
            
            # Handle Cr (Crore = 10 million)
            if 'Cr' in value:
                return float(value.replace('Cr', '').strip()) * 10000000
            
            # Handle Lakh (100,000)
            if 'Lakh' in value or 'L' in value:
                return float(re.sub(r'[LakhLAKH]', '', value).strip()) * 100000
            
            # Handle negative numbers in parentheses
            if value.startswith('(') and value.endswith(')'):
                return -float(value[1:-1])
            
            return float(value)
            
        except (ValueError, TypeError):
            return None
    
    def get_all_data(self) -> Dict[str, Any]:
        """Get all available data for the company"""
        return {
            'company_info': self.get_company_info(),
            'annual_results': self.get_annual_results(),
            'quarterly_results': self.get_quarterly_results(),
            'balance_sheet': self.get_balance_sheet(),
            'cash_flow': self.get_cash_flow(),
            'ratios': self.get_ratios(),
            'shareholding': self.get_shareholding(),
            'peers': self.get_peers(),
        }


def fetch_screener_data(symbol: str) -> Dict[str, Any]:
    """
    Convenience function to fetch all data from Screener.in
    
    Args:
        symbol: Stock symbol (e.g., 'ADANIPOWER')
    
    Returns:
        Dictionary with all scraped data
    """
    scraper = ScreenerScraper(symbol)
    return scraper.get_all_data()


if __name__ == "__main__":
    # Test with Adani Power
    data = fetch_screener_data("ADANIPOWER")
    print(f"Company: {data['company_info'].get('name', 'Unknown')}")
    print(f"Annual Results: {len(data['annual_results'])} rows")
    print(f"Peers: {len(data['peers'])} companies")
