"""
AI Assistant Module for Financial Modeling
Uses OpenAI GPT or Gemini AI for smart assumptions, valuation commentary, and NLP model building
"""

import os
import json
import logging
from typing import Dict, Any, Optional, List

logger = logging.getLogger(__name__)

# OpenAI API Key (from environment variables)
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY")

# Try to import OpenAI
try:
    from openai import OpenAI
    OPENAI_AVAILABLE = True
except ImportError:
    OPENAI_AVAILABLE = False
    logger.warning("openai not installed. Trying Gemini as fallback.")

# Try to import Google Generative AI
try:
    import google.generativeai as genai
    GEMINI_AVAILABLE = True
except ImportError:
    GEMINI_AVAILABLE = False
    logger.warning("google-generativeai not installed. AI features will use fallbacks.")


# OpenAI client
_openai_client = None

def get_openai_client():
    """Get or create OpenAI client"""
    global _openai_client
    if _openai_client is None and OPENAI_AVAILABLE:
        api_key = os.environ.get("OPENAI_API_KEY") or OPENAI_API_KEY
        if api_key:
            _openai_client = OpenAI(api_key=api_key)
            logger.info("OpenAI client initialized")
    return _openai_client


def setup_gemini():
    """Configure Gemini AI (fallback)"""
    if not GEMINI_AVAILABLE:
        return None
    
    api_key = os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
    if not api_key:
        logger.warning("No Gemini API key found. Set GEMINI_API_KEY environment variable.")
        return None
    
    genai.configure(api_key=api_key)
    return genai.GenerativeModel('gemini-1.5-flash')



# Import Bytez client
try:
    from agents.bytez_client import get_bytez_client
    BYTEZ_AVAILABLE = True
except ImportError:
    BYTEZ_AVAILABLE = False
    logger.warning("Bytez client import failed")

def call_ai(prompt: str, system_prompt: str = None) -> Optional[str]:
    """
    Call AI API (Bytez preferred, then OpenAI, then Gemini)
    
    Args:
        prompt: The user prompt
        system_prompt: Optional system prompt
    
    Returns:
        AI response text or None if failed
    """
    # Try Bytez first (Preferred Provider)
    if BYTEZ_AVAILABLE:
        bytez = get_bytez_client()
        response = bytez.generate_content(prompt, system_prompt)
        if response:
            return response

    # Try OpenAI second
    client = get_openai_client()
    if client:
        try:
            messages = []
            if system_prompt:
                messages.append({"role": "system", "content": system_prompt})
            messages.append({"role": "user", "content": prompt})
            
            response = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=messages,
                temperature=0.7,
                max_tokens=2000
            )
            return response.choices[0].message.content
        except Exception as e:
            logger.error(f"OpenAI API call failed: {e}")
    
    # Fallback to Gemini
    model = setup_gemini()
    if model:
        try:
            full_prompt = f"{system_prompt}\n\n{prompt}" if system_prompt else prompt
            response = model.generate_content(full_prompt)
            return response.text
        except Exception as e:
            logger.error(f"Gemini API call failed: {e}")
    
    return None


# Industry benchmark data (fallback when AI unavailable)
INDUSTRY_BENCHMARKS = {
    "it": {
        "revenue_growth": 0.15,
        "ebitda_margin": 0.25,
        "net_margin": 0.18,
        "tax_rate": 0.25,
        "capex_ratio": 0.05,
        "working_capital_days": 60,
        "terminal_growth": 0.04,
        "wacc_range": (0.10, 0.14),
    },
    "banking": {
        "nim": 0.035,
        "cost_income_ratio": 0.45,
        "credit_cost": 0.015,
        "crar": 0.16,
        "roe_target": 0.15,
        "terminal_growth": 0.05,
    },
    "pharma": {
        "revenue_growth": 0.12,
        "ebitda_margin": 0.22,
        "rd_ratio": 0.08,
        "net_margin": 0.14,
        "terminal_growth": 0.04,
    },
    "power": {
        "revenue_growth": 0.08,
        "ebitda_margin": 0.35,
        "capex_ratio": 0.15,
        "debt_equity": 1.5,
        "terminal_growth": 0.03,
    },
    "fmcg": {
        "revenue_growth": 0.10,
        "ebitda_margin": 0.20,
        "net_margin": 0.12,
        "working_capital_days": 30,
        "terminal_growth": 0.05,
    },
    "auto": {
        "revenue_growth": 0.10,
        "ebitda_margin": 0.12,
        "capex_ratio": 0.08,
        "working_capital_days": 45,
        "terminal_growth": 0.04,
    },
    "metals": {
        "revenue_growth": 0.08,
        "ebitda_margin": 0.18,
        "capex_ratio": 0.12,
        "debt_equity": 0.8,
        "terminal_growth": 0.03,
    },
    "general": {
        "revenue_growth": 0.10,
        "ebitda_margin": 0.18,
        "net_margin": 0.10,
        "tax_rate": 0.25,
        "capex_ratio": 0.06,
        "working_capital_days": 45,
        "terminal_growth": 0.04,
        "wacc_range": (0.10, 0.14),
    }
}


async def generate_smart_assumptions(
    industry: str,
    company_name: str,
    historical_data: Optional[Dict] = None,
    market_cap: Optional[float] = None
) -> Dict[str, Any]:
    """
    Generate AI-powered assumptions based on industry and company data
    
    Args:
        industry: Industry classification
        company_name: Name of the company
        historical_data: Historical financial data if available
        market_cap: Market capitalization in crores
    
    Returns:
        Dictionary of recommended assumptions with explanations
    """
    # Get industry benchmarks as base
    base_assumptions = INDUSTRY_BENCHMARKS.get(industry.lower(), INDUSTRY_BENCHMARKS["general"])
    
    # Build prompt for AI
    prompt = f"""You are a financial analyst. Generate optimal financial model assumptions for {company_name} in the {industry} sector.

Historical Data: {json.dumps(historical_data) if historical_data else 'Not available'}
Market Cap: {market_cap if market_cap else 'Not available'} Crores

Based on Indian market conditions and industry benchmarks, provide assumptions in this exact JSON format:
{{
    "revenue_growth": <decimal 0-0.30>,
    "ebitda_margin": <decimal 0-0.50>,
    "net_margin": <decimal 0-0.30>,
    "tax_rate": <decimal 0.20-0.30>,
    "capex_ratio": <decimal 0.02-0.20>,
    "working_capital_days": <integer 20-120>,
    "terminal_growth": <decimal 0.02-0.06>,
    "risk_free_rate": <decimal 0.06-0.08>,
    "equity_risk_premium": <decimal 0.05-0.08>,
    "beta": <decimal 0.6-1.8>,
    "explanation": "<brief 2-3 sentence rationale>"
}}

Return ONLY the JSON object, no other text."""

    # Try AI (OpenAI first, then Gemini)
    response_text = call_ai(prompt, "You are a senior financial analyst specializing in Indian markets.")
    
    if response_text:
        try:
            # Remove markdown code blocks if present
            if response_text.startswith("```"):
                response_text = response_text.split("```")[1]
                if response_text.startswith("json"):
                    response_text = response_text[4:]
            
            ai_assumptions = json.loads(response_text.strip())
            
            return {
                "source": "ai",
                "assumptions": ai_assumptions,
                "industry_benchmark": base_assumptions,
            }
            
        except Exception as e:
            logger.error(f"AI assumptions parsing failed: {e}")
    
    # Fallback to benchmarks
    return {
        "source": "benchmark",
        "assumptions": {
            "revenue_growth": base_assumptions.get("revenue_growth", 0.10),
            "ebitda_margin": base_assumptions.get("ebitda_margin", 0.18),
            "net_margin": base_assumptions.get("net_margin", 0.10),
            "tax_rate": base_assumptions.get("tax_rate", 0.25),
            "capex_ratio": base_assumptions.get("capex_ratio", 0.06),
            "working_capital_days": base_assumptions.get("working_capital_days", 45),
            "terminal_growth": base_assumptions.get("terminal_growth", 0.04),
            "risk_free_rate": 0.07,
            "equity_risk_premium": 0.06,
            "beta": 1.0,
            "explanation": f"Using {industry} industry benchmark values for Indian markets."
        },
        "industry_benchmark": base_assumptions,
    }


async def generate_valuation_commentary(
    company_name: str,
    industry: str,
    valuation_metrics: Dict[str, Any],
    assumptions: Dict[str, Any]
) -> Dict[str, str]:
    """
    Generate AI-powered investment thesis and risk factors
    
    Args:
        company_name: Name of the company
        industry: Industry classification
        valuation_metrics: Key metrics (EV, share price, WACC, etc.)
        assumptions: Model assumptions used
    
    Returns:
        Dictionary with investment_thesis, key_risks, and recommendation
    """
    prompt = f"""You are a senior equity research analyst. Write a brief investment commentary for {company_name} ({industry} sector).

Valuation Metrics:
- Enterprise Value: ₹{valuation_metrics.get('enterprise_value', 'N/A')} Crores
- Implied Share Price: ₹{valuation_metrics.get('share_price', 'N/A')}
- WACC: {valuation_metrics.get('wacc', 'N/A')}%

Key Assumptions:
- Revenue Growth: {assumptions.get('revenue_growth', 'N/A')}%
- EBITDA Margin: {assumptions.get('ebitda_margin', 'N/A')}%
- Terminal Growth: {assumptions.get('terminal_growth', 'N/A')}%

Provide a brief analysis in this exact JSON format:
{{
    "investment_thesis": "<2-3 sentences on value drivers>",
    "key_risks": "<2-3 key risks to monitor>",
    "recommendation": "<BUY/HOLD/SELL with brief rationale>",
    "sensitivity_note": "<1 sentence on key sensitivity drivers>"
}}

Return ONLY the JSON object."""

    response_text = call_ai(prompt, "You are a senior equity research analyst at a top investment bank.")
    
    if response_text:
        try:
            if response_text.startswith("```"):
                response_text = response_text.split("```")[1]
                if response_text.startswith("json"):
                    response_text = response_text[4:]
            
            return json.loads(response_text.strip())
            
        except Exception as e:
            logger.error(f"AI commentary parsing failed: {e}")
    
    # Fallback commentary
    return {
        "investment_thesis": f"{company_name} operates in the {industry} sector. The valuation is based on a DCF model using industry-standard assumptions.",
        "key_risks": "Key risks include market conditions, regulatory changes, and execution risk on growth projections.",
        "recommendation": "HOLD - Fair value based on current assumptions. Monitor quarterly results.",
        "sensitivity_note": "Valuation is most sensitive to terminal growth rate and WACC assumptions."
    }


async def parse_natural_language_request(prompt: str) -> Dict[str, Any]:
    """
    Parse natural language request to extract model parameters
    
    Args:
        prompt: User's natural language request
    
    Returns:
        Extracted parameters for model generation
    """
    system_prompt = """You are a financial modeling assistant. Parse the user's request and extract parameters for building a financial model.

User request: "{prompt}"

Extract the following in JSON format:
{{
    "company_name": "<extracted or null>",
    "industry": "<one of: it, banking, pharma, power, fmcg, auto, metals, general>",
    "forecast_years": <integer 3-10, default 5>,
    "custom_assumptions": {{
        "revenue_growth": <decimal or null>,
        "ebitda_margin": <decimal or null>,
        ... any mentioned assumptions
    }},
    "model_type": "<dcf, lbo, comps, or merger>",
    "special_instructions": "<any other requirements>"
}}

Return ONLY the JSON object."""

    response_text = call_ai(system_prompt.format(prompt=prompt), "You are a helpful financial assistant.")
    
    if response_text:
        try:
            if response_text.startswith("```"):
                response_text = response_text.split("```")[1]
                if response_text.startswith("json"):
                    response_text = response_text[4:]
            
            return json.loads(response_text.strip())
            
        except Exception as e:
            logger.error(f"NLP parsing failed: {e}")
    
    # Basic fallback parsing
    prompt_lower = prompt.lower()
    
    # Try to extract industry
    industry = "general"
    for ind in ["it", "banking", "pharma", "power", "fmcg", "auto", "metals"]:
        if ind in prompt_lower:
            industry = ind
            break
    
    # Try to extract years
    forecast_years = 5
    for word in prompt.split():
        if word.isdigit() and 3 <= int(word) <= 10:
            forecast_years = int(word)
            break
    
    return {
        "company_name": None,
        "industry": industry,
        "forecast_years": forecast_years,
        "custom_assumptions": {},
        "model_type": "dcf",
        "special_instructions": prompt
    }


def get_industry_benchmarks(industry: str) -> Dict[str, Any]:
    """Get industry benchmark values"""
    return INDUSTRY_BENCHMARKS.get(industry.lower(), INDUSTRY_BENCHMARKS["general"])
