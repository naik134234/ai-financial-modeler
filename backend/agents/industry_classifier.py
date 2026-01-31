"""
Industry Classifier Agent
Uses Gemini AI to identify company sector and appropriate modeling approach
"""

import google.generativeai as genai
from typing import Dict, Any, Optional
import json
import os
import logging

logger = logging.getLogger(__name__)


# Industry templates mapping
INDUSTRY_TEMPLATES = {
    'power': {
        'name': 'Power & Utilities',
        'model_type': 'power_sector',
        'key_metrics': ['PLF', 'Installed Capacity', 'Tariff', 'DSCR', 'LLCR'],
        'revenue_drivers': ['capacity_mw', 'plf_percent', 'tariff_per_kwh'],
        'cost_drivers': ['fuel_cost', 'employee_cost', 'om_cost'],
        'specific_ratios': ['DSCR', 'LLCR', 'IRR'],
    },
    'banking': {
        'name': 'Banking & Financial Services',
        'model_type': 'banking',
        'key_metrics': ['NIM', 'CASA', 'NPA', 'CAR', 'ROA', 'ROE'],
        'revenue_drivers': ['interest_income', 'fee_income', 'trading_income'],
        'cost_drivers': ['interest_expense', 'provisions', 'opex'],
        'specific_ratios': ['NIM', 'Cost-to-Income', 'NPA Ratio'],
    },
    'nbfc': {
        'name': 'NBFC & Financial Services',
        'model_type': 'nbfc',
        'key_metrics': ['NIM', 'AUM', 'NPA', 'CAR', 'ROA'],
        'revenue_drivers': ['interest_income', 'fee_income'],
        'cost_drivers': ['interest_expense', 'provisions', 'opex'],
        'specific_ratios': ['Spread', 'Cost-to-Income', 'NPA Ratio'],
    },
    'fmcg': {
        'name': 'FMCG & Consumer Goods',
        'model_type': 'fmcg',
        'key_metrics': ['Revenue Growth', 'Gross Margin', 'EBITDA Margin', 'ROCE'],
        'revenue_drivers': ['volume_growth', 'price_growth', 'new_products'],
        'cost_drivers': ['raw_material', 'advertising', 'distribution'],
        'specific_ratios': ['Inventory Days', 'Receivable Days', 'Payable Days'],
    },
    'it_services': {
        'name': 'IT Services & Technology',
        'model_type': 'it_services',
        'key_metrics': ['Revenue per Employee', 'Utilization', 'Attrition', 'EBIT Margin'],
        'revenue_drivers': ['headcount', 'utilization', 'billing_rate'],
        'cost_drivers': ['employee_cost', 'subcontracting', 'sga'],
        'specific_ratios': ['Employee Utilization', 'Offshore Mix', 'Revenue per Employee'],
    },
    'pharma': {
        'name': 'Pharmaceuticals',
        'model_type': 'pharma',
        'key_metrics': ['R&D Spend', 'ANDA Filings', 'US Revenue Share', 'EBITDA Margin'],
        'revenue_drivers': ['domestic_formulations', 'exports', 'api'],
        'cost_drivers': ['raw_material', 'rd_expense', 'marketing'],
        'specific_ratios': ['R&D to Revenue', 'US Sales Mix', 'Gross Margin'],
    },
    'infrastructure': {
        'name': 'Infrastructure & Construction',
        'model_type': 'infrastructure',
        'key_metrics': ['Order Book', 'Revenue Growth', 'EBITDA Margin', 'Working Capital Days'],
        'revenue_drivers': ['order_book', 'execution_rate', 'new_orders'],
        'cost_drivers': ['raw_material', 'subcontracting', 'employee_cost'],
        'specific_ratios': ['Book-to-Bill', 'Order Book/Revenue', 'Working Capital Days'],
    },
    'manufacturing': {
        'name': 'General Manufacturing',
        'model_type': 'manufacturing',
        'key_metrics': ['Capacity Utilization', 'Revenue Growth', 'EBITDA Margin', 'ROCE'],
        'revenue_drivers': ['volume', 'realization', 'product_mix'],
        'cost_drivers': ['raw_material', 'power_fuel', 'employee_cost'],
        'specific_ratios': ['Capacity Utilization', 'Asset Turnover', 'Working Capital Days'],
    },
    'general': {
        'name': 'General Corporate',
        'model_type': 'general',
        'key_metrics': ['Revenue Growth', 'EBITDA Margin', 'ROE', 'ROCE'],
        'revenue_drivers': ['volume', 'price', 'new_segments'],
        'cost_drivers': ['cogs', 'employee_cost', 'sga'],
        'specific_ratios': ['Gross Margin', 'Operating Margin', 'Net Margin'],
    },
}


class IndustryClassifier:
    """AI-powered industry classification agent"""
    
    CLASSIFICATION_PROMPT = """You are a financial analyst expert. Analyze the following company information and classify it into the most appropriate industry category.

Company Information:
- Name: {company_name}
- Sector: {sector}
- Industry: {industry}
- Description: {description}

Available Industry Categories:
1. power - Power generation, utilities, renewable energy companies
2. banking - Commercial banks, private banks, public sector banks
3. nbfc - Non-banking financial companies, housing finance, microfinance
4. fmcg - Fast moving consumer goods, food & beverages, personal care
5. it_services - IT services, software, technology companies
6. pharma - Pharmaceuticals, healthcare, medical devices
7. infrastructure - Infrastructure, construction, real estate, EPC
8. manufacturing - General manufacturing, auto, chemicals, metals
9. general - Companies that don't fit other categories

Respond with ONLY a JSON object in this exact format:
{{
    "industry_code": "power",
    "confidence": 0.95,
    "reasoning": "Brief explanation of classification",
    "sub_sector": "Thermal Power Generation",
    "key_characteristics": ["Large installed capacity", "Coal-based generation", "High debt levels"]
}}

Make sure to return only valid JSON, no additional text."""

    def __init__(self, api_key: Optional[str] = None):
        """Initialize with Gemini API key"""
        self.api_key = api_key or os.getenv('GEMINI_API_KEY')
        if self.api_key:
            genai.configure(api_key=self.api_key)
            self.model = genai.GenerativeModel('gemini-2.0-flash')
        else:
            self.model = None
            logger.warning("No Gemini API key provided. Classification will use rule-based fallback.")
    
    def classify(self, company_info: Dict[str, Any]) -> Dict[str, Any]:
        """
        Classify company into industry category
        
        Args:
            company_info: Dictionary with company name, sector, industry, description
        
        Returns:
            Classification result with industry code, confidence, and template info
        """
        # Try AI classification first
        if self.model:
            try:
                result = self._ai_classify(company_info)
                if result:
                    return result
            except Exception as e:
                logger.error(f"AI classification failed: {e}")
        
        # Fallback to rule-based classification
        return self._rule_based_classify(company_info)
    
    def _ai_classify(self, company_info: Dict[str, Any]) -> Optional[Dict[str, Any]]:
        """Use Gemini AI to classify the company"""
        prompt = self.CLASSIFICATION_PROMPT.format(
            company_name=company_info.get('name', 'Unknown'),
            sector=company_info.get('sector', 'Unknown'),
            industry=company_info.get('industry', 'Unknown'),
            description=company_info.get('description', 'No description available')[:500]
        )
        
        response = self.model.generate_content(prompt)
        response_text = response.text.strip()
        
        # Clean up response to extract JSON
        if '```json' in response_text:
            response_text = response_text.split('```json')[1].split('```')[0]
        elif '```' in response_text:
            response_text = response_text.split('```')[1].split('```')[0]
        
        result = json.loads(response_text)
        industry_code = result.get('industry_code', 'general')
        
        # Get template for this industry
        template = INDUSTRY_TEMPLATES.get(industry_code, INDUSTRY_TEMPLATES['general'])
        
        return {
            'industry_code': industry_code,
            'industry_name': template['name'],
            'model_type': template['model_type'],
            'confidence': result.get('confidence', 0.8),
            'reasoning': result.get('reasoning', ''),
            'sub_sector': result.get('sub_sector', ''),
            'key_characteristics': result.get('key_characteristics', []),
            'template': template,
        }
    
    def _rule_based_classify(self, company_info: Dict[str, Any]) -> Dict[str, Any]:
        """Fallback rule-based classification when AI is not available"""
        sector = company_info.get('sector', '').lower()
        industry = company_info.get('industry', '').lower()
        name = company_info.get('name', '').lower()
        
        # Simple keyword matching
        if any(kw in sector + industry + name for kw in ['power', 'energy', 'electricity', 'utility']):
            code = 'power'
        elif any(kw in sector + industry + name for kw in ['bank', 'banking']):
            code = 'banking'
        elif any(kw in sector + industry + name for kw in ['nbfc', 'finance', 'lending', 'housing finance']):
            code = 'nbfc'
        elif any(kw in sector + industry + name for kw in ['fmcg', 'consumer', 'food', 'beverage']):
            code = 'fmcg'
        elif any(kw in sector + industry + name for kw in ['software', 'it ', 'technology', 'tech', 'infosys', 'tcs', 'wipro']):
            code = 'it_services'
        elif any(kw in sector + industry + name for kw in ['pharma', 'drug', 'healthcare', 'medical']):
            code = 'pharma'
        elif any(kw in sector + industry + name for kw in ['infra', 'construction', 'real estate', 'epc']):
            code = 'infrastructure'
        elif any(kw in sector + industry + name for kw in ['auto', 'manufacturing', 'chemical', 'metal', 'steel']):
            code = 'manufacturing'
        else:
            code = 'general'
        
        template = INDUSTRY_TEMPLATES[code]
        
        return {
            'industry_code': code,
            'industry_name': template['name'],
            'model_type': template['model_type'],
            'confidence': 0.6,
            'reasoning': 'Rule-based classification from sector/industry keywords',
            'sub_sector': '',
            'key_characteristics': [],
            'template': template,
        }
    
    def get_template(self, industry_code: str) -> Dict[str, Any]:
        """Get the template for a specific industry"""
        return INDUSTRY_TEMPLATES.get(industry_code, INDUSTRY_TEMPLATES['general'])


def classify_company(company_info: Dict[str, Any], api_key: Optional[str] = None) -> Dict[str, Any]:
    """
    Convenience function to classify a company
    
    Args:
        company_info: Dictionary with company name, sector, industry, description
        api_key: Optional Gemini API key
    
    Returns:
        Classification result
    """
    classifier = IndustryClassifier(api_key)
    return classifier.classify(company_info)


if __name__ == "__main__":
    # Test classification
    test_company = {
        'name': 'Adani Power Limited',
        'sector': 'Utilities',
        'industry': 'Electric Utilities',
        'description': 'Adani Power Limited is engaged in power generation and related activities.'
    }
    
    result = classify_company(test_company)
    print(f"Industry: {result['industry_name']}")
    print(f"Model Type: {result['model_type']}")
    print(f"Confidence: {result['confidence']}")
