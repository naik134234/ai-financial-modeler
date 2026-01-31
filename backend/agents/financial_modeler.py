"""
Financial Modeler Agent
Core AI agent that builds financial model structure and generates Excel formulas
"""

import google.generativeai as genai
from typing import Dict, Any, List, Optional, Tuple
import json
import os
import logging
from datetime import datetime

logger = logging.getLogger(__name__)


class FinancialModeler:
    """AI agent that builds financial model structure and formulas"""
    
    MODEL_STRUCTURE_PROMPT = """You are an elite financial modeler from McKinsey/Goldman Sachs. Given the company data and industry, design a comprehensive Excel financial model structure.

Company: {company_name}
Industry: {industry_name}
Model Type: {model_type}
Historical Years: {historical_years}
Forecast Years: {forecast_years}

Available Historical Data:
{historical_data_summary}

Design the model with these sheets:
1. Cover - Company info, model date, key metrics summary
2. Assumptions - All input parameters (editable cells highlighted)
3. Historical - Historical financial data
4. Forecast IS - Projected Income Statement
5. Forecast BS - Projected Balance Sheet
6. Forecast CF - Projected Cash Flow
7. Working Capital - Detailed WC schedule
8. Debt Schedule - Loan amortization
9. Valuation - DCF, sensitivity analysis
10. Dashboard - Charts and key metrics

For each sheet, provide:
- Key line items
- Formula logic (not actual Excel formulas, but the calculation approach)
- Cross-references to other sheets

Respond with ONLY a JSON object in this format:
{{
    "sheets": [
        {{
            "name": "Cover",
            "purpose": "Summary page with key info",
            "sections": [
                {{
                    "name": "Company Information",
                    "items": ["Company Name", "Industry", "Model Date", "Analyst"]
                }}
            ]
        }}
    ],
    "key_assumptions": [
        {{
            "name": "Revenue Growth",
            "default_value": 0.10,
            "unit": "percent",
            "driver_logic": "Historical CAGR adjusted for industry outlook"
        }}
    ],
    "valuation_approach": {{
        "primary": "DCF",
        "secondary": "Trading Multiples",
        "wacc_components": ["Cost of Equity", "Cost of Debt", "Target D/E"]
    }}
}}"""

    FORMULA_GENERATION_PROMPT = """You are an Excel formula expert. Generate the exact Excel formula for this financial calculation.

Context:
- Sheet: {sheet_name}
- Row: {row_number}
- Column: {column_letter} (representing {period})
- Cell Purpose: {cell_purpose}
- Related Cells: {related_cells}

Rules:
1. Use proper Excel syntax
2. Reference cells by sheet!cell format for cross-sheet references
3. Use named ranges where appropriate
4. Include IFERROR for robustness
5. Add appropriate rounding

Return ONLY the Excel formula starting with ="""

    def __init__(self, api_key: Optional[str] = None):
        """Initialize with Gemini API key"""
        self.api_key = api_key or os.getenv('GEMINI_API_KEY')
        if self.api_key:
            genai.configure(api_key=self.api_key)
            self.model = genai.GenerativeModel('gemini-2.0-flash')
        else:
            self.model = None
            logger.warning("No Gemini API key. Using template-based modeling.")
    
    def design_model_structure(
        self,
        company_name: str,
        industry_info: Dict[str, Any],
        historical_data: Dict[str, Any],
        forecast_years: int = 5
    ) -> Dict[str, Any]:
        """
        Design the financial model structure
        
        Args:
            company_name: Name of the company
            industry_info: Industry classification result
            historical_data: Historical financial data
            forecast_years: Number of forecast years
        
        Returns:
            Model structure with sheets, assumptions, and valuation approach
        """
        # Try AI-based design
        if self.model:
            try:
                return self._ai_design_structure(
                    company_name, industry_info, historical_data, forecast_years
                )
            except Exception as e:
                logger.error(f"AI model design failed: {e}")
        
        # Fallback to template-based design
        return self._template_based_design(industry_info, forecast_years)
    
    def _ai_design_structure(
        self,
        company_name: str,
        industry_info: Dict[str, Any],
        historical_data: Dict[str, Any],
        forecast_years: int
    ) -> Dict[str, Any]:
        """Use AI to design model structure"""
        # Summarize historical data for the prompt
        historical_summary = self._summarize_historical_data(historical_data)
        
        prompt = self.MODEL_STRUCTURE_PROMPT.format(
            company_name=company_name,
            industry_name=industry_info.get('industry_name', 'General'),
            model_type=industry_info.get('model_type', 'general'),
            historical_years=5,
            forecast_years=forecast_years,
            historical_data_summary=historical_summary
        )
        
        response = self.model.generate_content(prompt)
        response_text = response.text.strip()
        
        # Clean up JSON
        if '```json' in response_text:
            response_text = response_text.split('```json')[1].split('```')[0]
        elif '```' in response_text:
            response_text = response_text.split('```')[1].split('```')[0]
        
        return json.loads(response_text)
    
    def _summarize_historical_data(self, data: Dict[str, Any]) -> str:
        """Create a summary of historical data for the prompt"""
        summary_lines = []
        
        if 'income_statement' in data:
            is_data = data['income_statement']
            summary_lines.append(f"Revenue: {is_data.get('revenue', 'N/A')}")
            summary_lines.append(f"EBITDA: {is_data.get('ebitda', 'N/A')}")
            summary_lines.append(f"Net Income: {is_data.get('net_income', 'N/A')}")
        
        if 'balance_sheet' in data:
            bs_data = data['balance_sheet']
            summary_lines.append(f"Total Assets: {bs_data.get('total_assets', 'N/A')}")
            summary_lines.append(f"Total Debt: {bs_data.get('total_debt', 'N/A')}")
            summary_lines.append(f"Total Equity: {bs_data.get('total_equity', 'N/A')}")
        
        return '\n'.join(summary_lines) if summary_lines else "Limited historical data available"
    
    def _template_based_design(
        self,
        industry_info: Dict[str, Any],
        forecast_years: int
    ) -> Dict[str, Any]:
        """Use template-based model design"""
        model_type = industry_info.get('model_type', 'general')
        
        # Base model structure applicable to all industries
        base_structure = {
            'sheets': [
                {
                    'name': 'Cover',
                    'purpose': 'Summary page with company info and key metrics',
                    'sections': [
                        {'name': 'Company Information', 'items': ['Company Name', 'Industry', 'Model Date', 'Analyst']},
                        {'name': 'Key Metrics', 'items': ['Market Cap', 'Enterprise Value', 'Revenue', 'EBITDA', 'Net Income']},
                        {'name': 'Valuation Summary', 'items': ['DCF Value', 'Implied Upside', 'Target Price']},
                    ]
                },
                {
                    'name': 'Assumptions',
                    'purpose': 'All input parameters and assumptions',
                    'sections': [
                        {'name': 'Growth Assumptions', 'items': ['Revenue Growth Rate', 'Volume Growth', 'Price Growth']},
                        {'name': 'Margin Assumptions', 'items': ['Gross Margin', 'EBITDA Margin', 'Net Margin']},
                        {'name': 'Working Capital', 'items': ['Receivable Days', 'Inventory Days', 'Payable Days']},
                        {'name': 'Capex & D&A', 'items': ['Capex % of Revenue', 'D&A % of PPE']},
                        {'name': 'Valuation', 'items': ['Risk-free Rate', 'Equity Risk Premium', 'Beta', 'Cost of Debt', 'Tax Rate', 'Terminal Growth']},
                    ]
                },
                {
                    'name': 'Historical',
                    'purpose': 'Historical financial statements',
                    'sections': [
                        {'name': 'Income Statement', 'items': ['Revenue', 'COGS', 'Gross Profit', 'Operating Expenses', 'EBITDA', 'D&A', 'EBIT', 'Interest', 'PBT', 'Tax', 'Net Income']},
                        {'name': 'Balance Sheet', 'items': ['Cash', 'Receivables', 'Inventory', 'Current Assets', 'PPE', 'Total Assets', 'Payables', 'Short-term Debt', 'Current Liabilities', 'Long-term Debt', 'Total Liabilities', 'Equity']},
                        {'name': 'Cash Flow', 'items': ['Operating CF', 'Capex', 'Free Cash Flow', 'Dividends', 'Net Borrowing']},
                    ]
                },
                {
                    'name': 'Forecast_IS',
                    'purpose': 'Projected Income Statement',
                    'sections': [
                        {'name': 'Revenue Build-up', 'items': ['Volume', 'Price/Mix', 'Revenue']},
                        {'name': 'Cost Structure', 'items': ['COGS', 'Gross Profit', 'SG&A', 'Other OpEx', 'EBITDA']},
                        {'name': 'Below EBITDA', 'items': ['D&A', 'EBIT', 'Interest Expense', 'Interest Income', 'PBT', 'Tax', 'Net Income']},
                    ]
                },
                {
                    'name': 'Forecast_BS',
                    'purpose': 'Projected Balance Sheet',
                    'sections': [
                        {'name': 'Assets', 'items': ['Cash', 'Receivables', 'Inventory', 'Other Current', 'Total Current Assets', 'Gross PPE', 'Accumulated D&A', 'Net PPE', 'Intangibles', 'Other Non-current', 'Total Assets']},
                        {'name': 'Liabilities', 'items': ['Payables', 'Accrued Expenses', 'Short-term Debt', 'Total Current Liabilities', 'Long-term Debt', 'Other Non-current', 'Total Liabilities']},
                        {'name': 'Equity', 'items': ['Share Capital', 'Retained Earnings', 'Total Equity', 'Total L&E', 'Balance Check']},
                    ]
                },
                {
                    'name': 'Forecast_CF',
                    'purpose': 'Projected Cash Flow Statement',
                    'sections': [
                        {'name': 'Operating Activities', 'items': ['Net Income', 'D&A', 'Change in WC', 'Other Operating', 'Operating Cash Flow']},
                        {'name': 'Investing Activities', 'items': ['Capex', 'Acquisitions', 'Other Investing', 'Investing Cash Flow']},
                        {'name': 'Financing Activities', 'items': ['Debt Raised', 'Debt Repaid', 'Dividends', 'Equity Issuance', 'Financing Cash Flow']},
                        {'name': 'Cash Movement', 'items': ['Net Cash Flow', 'Opening Cash', 'Closing Cash']},
                    ]
                },
                {
                    'name': 'Working_Capital',
                    'purpose': 'Detailed working capital schedule',
                    'sections': [
                        {'name': 'Receivables', 'items': ['Opening Balance', 'Revenue', 'Collections', 'Closing Balance', 'Days']},
                        {'name': 'Inventory', 'items': ['Opening Balance', 'COGS', 'Production', 'Closing Balance', 'Days']},
                        {'name': 'Payables', 'items': ['Opening Balance', 'Purchases', 'Payments', 'Closing Balance', 'Days']},
                        {'name': 'Net Working Capital', 'items': ['Total WC', 'Change in WC']},
                    ]
                },
                {
                    'name': 'Debt_Schedule',
                    'purpose': 'Loan amortization and interest calculation',
                    'sections': [
                        {'name': 'Existing Debt', 'items': ['Opening Balance', 'Drawdown', 'Repayment', 'Closing Balance', 'Interest Rate', 'Interest Expense']},
                        {'name': 'New Debt', 'items': ['Facility Size', 'Drawdown', 'Repayment', 'Balance', 'Interest']},
                        {'name': 'Summary', 'items': ['Total Debt', 'Total Interest', 'Net Debt', 'Leverage Ratios']},
                    ]
                },
                {
                    'name': 'Valuation',
                    'purpose': 'DCF and sensitivity analysis',
                    'sections': [
                        {'name': 'WACC Calculation', 'items': ['Cost of Equity', 'Cost of Debt', 'Tax Shield', 'WACC']},
                        {'name': 'DCF Valuation', 'items': ['FCFF', 'Discount Factor', 'PV of FCFF', 'Terminal Value', 'Enterprise Value', 'Net Debt', 'Equity Value', 'Shares Outstanding', 'Implied Share Price']},
                        {'name': 'Sensitivity', 'items': ['WACC vs Terminal Growth Matrix', 'WACC vs EBITDA Margin Matrix']},
                        {'name': 'Trading Multiples', 'items': ['EV/EBITDA', 'EV/Revenue', 'P/E', 'P/B']},
                    ]
                },
                {
                    'name': 'Dashboard',
                    'purpose': 'Visual summary with charts',
                    'sections': [
                        {'name': 'Revenue & Growth', 'items': ['Revenue Chart', 'Growth Rate Chart']},
                        {'name': 'Profitability', 'items': ['EBITDA Margin Chart', 'Net Margin Chart']},
                        {'name': 'Returns', 'items': ['ROE Chart', 'ROCE Chart']},
                        {'name': 'Valuation', 'items': ['DCF Sensitivity Chart', 'Football Field Chart']},
                    ]
                },
            ],
            'key_assumptions': self._get_industry_assumptions(model_type),
            'valuation_approach': {
                'primary': 'DCF',
                'secondary': 'Trading Multiples',
                'wacc_components': ['Cost of Equity (CAPM)', 'Cost of Debt (after-tax)', 'Target D/E Ratio'],
            },
            'forecast_years': forecast_years,
        }
        
        # Add industry-specific sheets
        if model_type == 'power':
            base_structure['sheets'].insert(3, {
                'name': 'Power_Operations',
                'purpose': 'Power sector specific operating metrics',
                'sections': [
                    {'name': 'Capacity', 'items': ['Installed Capacity (MW)', 'Operational Capacity', 'Under Construction']},
                    {'name': 'Generation', 'items': ['PLF (%)', 'Availability Factor', 'Units Generated (MU)']},
                    {'name': 'Revenue Build-up', 'items': ['Capacity Charges', 'Energy Charges', 'Other Revenue']},
                    {'name': 'Fuel Costs', 'items': ['Coal Consumption', 'Coal Price', 'Total Fuel Cost']},
                    {'name': 'Project Finance', 'items': ['DSCR', 'LLCR', 'IRR']},
                ]
            })
        
        return base_structure
    
    def _get_industry_assumptions(self, model_type: str) -> List[Dict[str, Any]]:
        """Get industry-specific assumptions"""
        base_assumptions = [
            {'name': 'Revenue Growth Rate', 'default_value': 0.10, 'unit': 'percent', 'driver_logic': 'Industry growth + market share gains'},
            {'name': 'Gross Margin', 'default_value': 0.40, 'unit': 'percent', 'driver_logic': 'Historical average with efficiency improvements'},
            {'name': 'EBITDA Margin', 'default_value': 0.20, 'unit': 'percent', 'driver_logic': 'Operating leverage and cost optimization'},
            {'name': 'D&A % of Gross PPE', 'default_value': 0.05, 'unit': 'percent', 'driver_logic': 'Based on asset life'},
            {'name': 'Capex % of Revenue', 'default_value': 0.05, 'unit': 'percent', 'driver_logic': 'Maintenance + growth capex'},
            {'name': 'Working Capital Days', 'default_value': 45, 'unit': 'days', 'driver_logic': 'Industry benchmark'},
            {'name': 'Tax Rate', 'default_value': 0.25, 'unit': 'percent', 'driver_logic': 'Statutory rate'},
            {'name': 'Risk-free Rate', 'default_value': 0.07, 'unit': 'percent', 'driver_logic': '10-year G-Sec yield'},
            {'name': 'Equity Risk Premium', 'default_value': 0.06, 'unit': 'percent', 'driver_logic': 'India market risk premium'},
            {'name': 'Beta', 'default_value': 1.0, 'unit': 'number', 'driver_logic': 'Industry average'},
            {'name': 'Cost of Debt', 'default_value': 0.10, 'unit': 'percent', 'driver_logic': 'Current borrowing rate'},
            {'name': 'Terminal Growth', 'default_value': 0.04, 'unit': 'percent', 'driver_logic': 'Long-term GDP growth'},
        ]
        
        if model_type == 'power':
            base_assumptions.extend([
                {'name': 'PLF (%)', 'default_value': 0.70, 'unit': 'percent', 'driver_logic': 'Plant efficiency'},
                {'name': 'Tariff per kWh (₹)', 'default_value': 4.5, 'unit': 'currency', 'driver_logic': 'Regulatory tariff'},
                {'name': 'Coal Price Growth', 'default_value': 0.03, 'unit': 'percent', 'driver_logic': 'Fuel price escalation'},
                {'name': 'Capacity Addition (MW)', 'default_value': 0, 'unit': 'number', 'driver_logic': 'Expansion plans'},
            ])
        
        return base_assumptions
    
    def generate_formula(
        self,
        sheet_name: str,
        row_number: int,
        column_letter: str,
        period: str,
        cell_purpose: str,
        related_cells: Dict[str, str]
    ) -> str:
        """
        Generate an Excel formula for a specific cell
        
        Returns:
            Excel formula string
        """
        if self.model:
            try:
                return self._ai_generate_formula(
                    sheet_name, row_number, column_letter, period, cell_purpose, related_cells
                )
            except Exception as e:
                logger.error(f"AI formula generation failed: {e}")
        
        return self._template_formula(cell_purpose, related_cells)
    
    def _ai_generate_formula(
        self,
        sheet_name: str,
        row_number: int,
        column_letter: str,
        period: str,
        cell_purpose: str,
        related_cells: Dict[str, str]
    ) -> str:
        """Use AI to generate Excel formula"""
        prompt = self.FORMULA_GENERATION_PROMPT.format(
            sheet_name=sheet_name,
            row_number=row_number,
            column_letter=column_letter,
            period=period,
            cell_purpose=cell_purpose,
            related_cells=json.dumps(related_cells)
        )
        
        response = self.model.generate_content(prompt)
        formula = response.text.strip()
        
        # Ensure formula starts with =
        if not formula.startswith('='):
            formula = '=' + formula
        
        return formula
    
    def _template_formula(self, cell_purpose: str, related_cells: Dict[str, str]) -> str:
        """Generate formula from templates"""
        purpose_lower = cell_purpose.lower()
        
        # Common formula templates
        if 'growth' in purpose_lower and 'revenue' in related_cells:
            return f"=IFERROR({related_cells.get('revenue', 'B5')}*(1+{related_cells.get('growth_rate', '$B$3')}),0)"
        
        if 'ebitda' in purpose_lower and 'margin' in purpose_lower:
            return f"={related_cells.get('revenue', 'B5')}*{related_cells.get('ebitda_margin', '$B$4')}"
        
        if 'fcff' in purpose_lower or 'free cash flow' in purpose_lower:
            return f"={related_cells.get('ebitda', 'B10')}-{related_cells.get('capex', 'B15')}-{related_cells.get('delta_wc', 'B16')}"
        
        if 'wacc' in purpose_lower:
            return "=($B$20*$B$21)+($B$22*$B$23*(1-$B$24))*($B$25/(1+$B$25))"
        
        if 'terminal value' in purpose_lower:
            return f"=IFERROR({related_cells.get('fcff', 'G15')}*(1+{related_cells.get('terminal_growth', '$B$26')})/({related_cells.get('wacc', '$B$27')}-{related_cells.get('terminal_growth', '$B$26')}),0)"
        
        return "=0"


def create_model_structure(
    company_name: str,
    industry_info: Dict[str, Any],
    historical_data: Dict[str, Any],
    forecast_years: int = 5,
    api_key: Optional[str] = None
) -> Dict[str, Any]:
    """
    Convenience function to create model structure
    """
    modeler = FinancialModeler(api_key)
    return modeler.design_model_structure(
        company_name, industry_info, historical_data, forecast_years
    )


if __name__ == "__main__":
    # Test model structure design
    from industry_classifier import classify_company
    
    company_info = {
        'name': 'Adani Power Limited',
        'sector': 'Utilities',
        'industry': 'Electric Utilities',
    }
    
    industry = classify_company(company_info)
    structure = create_model_structure(
        'Adani Power Limited',
        industry,
        {},
        forecast_years=5
    )
    
    print(f"Model has {len(structure['sheets'])} sheets")
    for sheet in structure['sheets']:
        print(f"  - {sheet['name']}: {sheet['purpose']}")
