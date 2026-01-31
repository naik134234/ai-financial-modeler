"""
Model Templates - Pre-built configurations for different sectors
Provides quick-start templates with industry-specific assumptions
"""

from typing import Dict, Any, List

# LBO Templates for different sectors
LBO_TEMPLATES = {
    "tech_saas": {
        "name": "Tech / SaaS",
        "description": "High-growth software companies with recurring revenue",
        "icon": "💻",
        "assumptions": {
            "holding_period": 5,
            "entry_multiple": 12.0,
            "exit_multiple": 14.0,
            "senior_debt_multiple": 2.0,
            "senior_interest_rate": 0.07,
            "mezz_debt_multiple": 1.0,
            "mezz_interest_rate": 0.11,
            "sub_debt_multiple": 0.5,
            "sub_interest_rate": 0.13,
            "revenue_growth": 0.20,
            "ebitda_margin": 0.30,
        },
        "key_metrics": ["ARR", "Net Revenue Retention", "Rule of 40", "CAC Payback"],
    },
    "manufacturing": {
        "name": "Manufacturing",
        "description": "Industrial and manufacturing companies",
        "icon": "🏭",
        "assumptions": {
            "holding_period": 5,
            "entry_multiple": 7.0,
            "exit_multiple": 7.5,
            "senior_debt_multiple": 3.5,
            "senior_interest_rate": 0.08,
            "mezz_debt_multiple": 1.5,
            "mezz_interest_rate": 0.12,
            "sub_debt_multiple": 0.5,
            "sub_interest_rate": 0.14,
            "revenue_growth": 0.05,
            "ebitda_margin": 0.15,
        },
        "key_metrics": ["Capacity Utilization", "Gross Margin", "Working Capital Days"],
    },
    "banking": {
        "name": "Banking / NBFC",
        "description": "Financial services and lending companies",
        "icon": "🏦",
        "assumptions": {
            "holding_period": 5,
            "entry_multiple": 10.0,
            "exit_multiple": 11.0,
            "senior_debt_multiple": 1.5,
            "senior_interest_rate": 0.09,
            "mezz_debt_multiple": 0.5,
            "mezz_interest_rate": 0.12,
            "sub_debt_multiple": 0.0,
            "sub_interest_rate": 0.14,
            "revenue_growth": 0.12,
            "ebitda_margin": 0.35,
        },
        "key_metrics": ["NIM", "ROE", "CAR", "NPA Ratio", "Cost-to-Income"],
    },
    "healthcare": {
        "name": "Healthcare / Pharma",
        "description": "Healthcare services and pharmaceutical companies",
        "icon": "💊",
        "assumptions": {
            "holding_period": 5,
            "entry_multiple": 10.0,
            "exit_multiple": 11.0,
            "senior_debt_multiple": 2.5,
            "senior_interest_rate": 0.08,
            "mezz_debt_multiple": 1.0,
            "mezz_interest_rate": 0.11,
            "sub_debt_multiple": 0.5,
            "sub_interest_rate": 0.13,
            "revenue_growth": 0.10,
            "ebitda_margin": 0.22,
        },
        "key_metrics": ["R&D as % of Revenue", "Pipeline Value", "Gross Margin"],
    },
    "consumer_retail": {
        "name": "Consumer / Retail",
        "description": "Consumer goods and retail companies",
        "icon": "🛒",
        "assumptions": {
            "holding_period": 5,
            "entry_multiple": 8.0,
            "exit_multiple": 9.0,
            "senior_debt_multiple": 3.0,
            "senior_interest_rate": 0.08,
            "mezz_debt_multiple": 1.5,
            "mezz_interest_rate": 0.12,
            "sub_debt_multiple": 0.5,
            "sub_interest_rate": 0.14,
            "revenue_growth": 0.08,
            "ebitda_margin": 0.12,
        },
        "key_metrics": ["Same-Store Sales", "Inventory Turnover", "Gross Margin"],
    },
    "infrastructure": {
        "name": "Infrastructure / Power",
        "description": "Power, utilities, and infrastructure companies",
        "icon": "⚡",
        "assumptions": {
            "holding_period": 7,
            "entry_multiple": 8.0,
            "exit_multiple": 8.5,
            "senior_debt_multiple": 4.0,
            "senior_interest_rate": 0.085,
            "mezz_debt_multiple": 1.5,
            "mezz_interest_rate": 0.11,
            "sub_debt_multiple": 0.5,
            "sub_interest_rate": 0.13,
            "revenue_growth": 0.06,
            "ebitda_margin": 0.35,
        },
        "key_metrics": ["PLF", "Asset Utilization", "DSCR", "Project IRR"],
    },
}


# M&A Templates
MA_TEMPLATES = {
    "horizontal": {
        "name": "Horizontal Merger",
        "description": "Same industry, market share expansion",
        "icon": "🔗",
        "assumptions": {
            "offer_premium": 0.25,
            "percent_stock": 0.50,
            "percent_cash": 0.50,
            "synergies_revenue": 100,
            "synergies_cost": 200,
            "acquirer_growth_rate": 0.08,
            "target_growth_rate": 0.05,
        },
    },
    "vertical": {
        "name": "Vertical Integration",
        "description": "Supply chain integration",
        "icon": "📊",
        "assumptions": {
            "offer_premium": 0.20,
            "percent_stock": 0.30,
            "percent_cash": 0.70,
            "synergies_revenue": 50,
            "synergies_cost": 300,
            "acquirer_growth_rate": 0.06,
            "target_growth_rate": 0.04,
        },
    },
    "strategic": {
        "name": "Strategic Acquisition",
        "description": "Capability/technology acquisition",
        "icon": "🎯",
        "assumptions": {
            "offer_premium": 0.35,
            "percent_stock": 0.60,
            "percent_cash": 0.40,
            "synergies_revenue": 200,
            "synergies_cost": 100,
            "acquirer_growth_rate": 0.10,
            "target_growth_rate": 0.15,
        },
    },
}


def get_lbo_templates() -> List[Dict[str, Any]]:
    """Get all LBO templates"""
    return [
        {
            "id": key,
            **value
        }
        for key, value in LBO_TEMPLATES.items()
    ]


def get_ma_templates() -> List[Dict[str, Any]]:
    """Get all M&A templates"""
    return [
        {
            "id": key,
            **value
        }
        for key, value in MA_TEMPLATES.items()
    ]


def get_lbo_template(template_id: str) -> Dict[str, Any]:
    """Get a specific LBO template by ID"""
    if template_id not in LBO_TEMPLATES:
        return None
    return {
        "id": template_id,
        **LBO_TEMPLATES[template_id]
    }


def get_ma_template(template_id: str) -> Dict[str, Any]:
    """Get a specific M&A template by ID"""
    if template_id not in MA_TEMPLATES:
        return None
    return {
        "id": template_id,
        **MA_TEMPLATES[template_id]
    }
