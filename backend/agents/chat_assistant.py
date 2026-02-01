"""
AI Chat Assistant
Interactive chat for model Q&A and what-if scenarios
Uses Claude API via OpenRouter
"""

import os
import json
import logging
import requests
from typing import Dict, Any, List, Optional

logger = logging.getLogger(__name__)


GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
OPENAI_URL = "https://api.openai.com/v1/chat/completions"

class ChatAssistant:
    """AI-powered chat assistant for financial model Q&A"""
    

    def __init__(self, model: str = "gpt-4o"):
        self.model = model
        
        # Explicitly load .env from backend root
        current_dir = os.path.dirname(os.path.abspath(__file__))
        backend_dir = os.path.dirname(os.path.dirname(current_dir)) # agents -> backend -> root (wait, agents is in backend)
        # agents is in backend. so dirname(agents) is backend.
        backend_dir = os.path.dirname(current_dir)
        env_path = os.path.join(backend_dir, '.env')
        
        from dotenv import load_dotenv
        load_dotenv(env_path)
        
        # Re-fetch to ensure fresh env vars
        self.openai_key = os.getenv("OPENAI_API_KEY")
        self.gemini_key = os.getenv("GEMINI_API_KEY")
        self.conversation_history: List[Dict] = []
        
        print(f"--------- CHAT ASSISTANT INIT ---------")
        print(f"OpenAI Key Present: {bool(self.openai_key)}")
        print(f"Gemini Key Present: {bool(self.gemini_key)}")
        if self.gemini_key:
             print(f"Gemini Key: {self.gemini_key[:5]}...{self.gemini_key[-5:]}")

        # Configure Gemini if available
        if self.gemini_key:
            try:
                import google.generativeai as genai
                genai.configure(api_key=self.gemini_key)
                self.gemini_model = genai.GenerativeModel('gemini-pro')
                print("✅ ChatAssistant configured with Gemini")
                logger.info("ChatAssistant configured with Gemini")
            except Exception as e:
                print(f"❌ Failed to configure Gemini: {e}")
                logger.error(f"Failed to configure Gemini: {e}")
                self.gemini_model = None
        else:
            print("⚠️ No Gemini Key found, skipping Gemini config")
            self.gemini_model = None
    
    def create_model_context(self, job_data: Dict) -> str:
        """Create context string from job data for the AI"""
        company = job_data.get('company_name', 'Unknown Company')
        industry = job_data.get('industry', 'general')
        assumptions = job_data.get('assumptions', {})
        valuation = job_data.get('valuation_data', {})
        
        def safe_fmt(val, fmt="{:,.2f}", prefix=""):
            if val is None or val == 'N/A':
                return 'N/A'
            try:
                return f"{prefix}{fmt.format(float(val))}"
            except (ValueError, TypeError):
                return str(val)

        context = f"""You are analyzing a financial model for {company} ({industry} sector).

KEY ASSUMPTIONS:
- Revenue Growth: {safe_fmt(assumptions.get('revenue_growth'), '{:.1%}', '')}
- EBITDA Margin: {safe_fmt(assumptions.get('ebitda_margin'), '{:.1%}', '')}
- WACC: {safe_fmt(assumptions.get('wacc'), '{:.1%}', '')}
- Terminal Growth: {safe_fmt(assumptions.get('terminal_growth'), '{:.1%}', '')}
- Tax Rate: {safe_fmt(assumptions.get('tax_rate'), '{:.1%}', '')}

VALUATION METRICS:
- Enterprise Value: {safe_fmt(valuation.get('enterprise_value'), '{:,.0f}', '₹')} Cr
- Equity Value: {safe_fmt(valuation.get('equity_value'), '{:,.0f}', '₹')} Cr
- Implied Share Price: {safe_fmt(valuation.get('share_price'), '{:,.2f}', '₹')}
- Current Market Price: {safe_fmt(valuation.get('current_price'), '{:,.2f}', '₹')}

COMPANY FINANCIALS:
- Revenue: {safe_fmt(assumptions.get('base_revenue'), '{:,.0f}', '₹')} Cr
- EBITDA: {safe_fmt(assumptions.get('base_ebitda'), '{:,.0f}', '₹')} Cr
- Net Debt: {safe_fmt(valuation.get('net_debt'), '{:,.0f}', '₹')} Cr

You are a helpful financial analyst assistant. Answer questions about this model concisely.
When discussing valuation, explain the key drivers and risks.
For what-if scenarios, estimate the directional impact on valuation."""
        
        return context
    
    def chat(
        self, 
        message: str, 
        job_data: Dict,
        chat_history: Optional[List[Dict]] = None
    ) -> Dict[str, Any]:
        """
        Send a chat message and get AI response
        Prioritizes Gemini if available, falls back to OpenAI
        """
        try:
            # Try Gemini First
            if self.gemini_model:
                try:
                    context = self.create_model_context(job_data)
                    
                    # Construct prompt with context
                    prompt = f"{context}\n\nUSER QUESTION: {message}"
                    
                    # Simple generation for now (history handling depends on user pref)
                    response = self.gemini_model.generate_content(prompt)
                    
                    return {
                        "success": True,
                        "response": response.text,
                        "model": "gemini-pro",
                        "tokens_used": 0
                    }
                except Exception as e:
                    logger.error(f"Gemini chat error: {e}")
                    # Fallthrough to OpenAI if valid key
            
            # Fallback to OpenAI
            if not self.openai_key:
                 return {
                    "success": False,
                    "response": "Error: configured AI providers unavailable. Check API keys.",
                    "error": "Missing API Keys"
                }

            context = self.create_model_context(job_data)
            
            messages = [{"role": "system", "content": context}]
            
            # Add chat history
            if chat_history:
                for msg in chat_history[-6:]:  # Last 6 messages for context
                    messages.append({
                        "role": msg.get("role", "user"),
                        "content": msg.get("content", "")
                    })
            
            # Add current message
            messages.append({"role": "user", "content": message})
            
            response = requests.post(
                OPENAI_URL,
                headers={
                    "Authorization": f"Bearer {self.openai_key}",
                    "Content-Type": "application/json"
                },
                json={
                    "model": self.model,
                    "messages": messages,
                    "max_tokens": 500,
                    "temperature": 0.7
                },
                timeout=30
            )
            
            if response.status_code == 200:
                result = response.json()
                ai_message = result['choices'][0]['message']['content']
                
                return {
                    "success": True,
                    "response": ai_message,
                    "model": self.model,
                    "tokens_used": result.get('usage', {}).get('total_tokens', 0)
                }
            else:
                logger.error(f"Chat API error: {response.status_code} - {response.text}")
                return {
                    "success": False,
                    "response": "I'm having trouble connecting. Please try again.",
                    "error": response.text
                }
                
        except Exception as e:
            logger.error(f"Chat error: {e}")
            return {
                "success": False,
                "response": f"Error: {str(e)}",
                "error": str(e)
            }
    
    def suggest_questions(self, job_data: Dict) -> List[str]:
        """Generate suggested questions based on the model"""
        company = job_data.get('company_name', 'the company')
        valuation = job_data.get('valuation_data', {})
        
        share_price = valuation.get('share_price', 0)
        current_price = valuation.get('current_price', 0)
        
        suggestions = [
            f"What are the key value drivers for {company}?",
            "What are the main risks to this valuation?",
            "How sensitive is the valuation to WACC changes?",
            "What would happen if revenue growth dropped to 5%?",
        ]
        
        if share_price > current_price:
            suggestions.append("Why is the model showing upside potential?")
        else:
            suggestions.append("Why is the current price above intrinsic value?")
        
        return suggestions


# Singleton instance
chat_assistant = ChatAssistant()


def process_chat_message(message: str, job_data: Dict, history: List = None) -> Dict:
    """Main entry point for chat processing"""
    return chat_assistant.chat(message, job_data, history)


def get_suggested_questions(job_data: Dict) -> List[str]:
    """Get AI-generated question suggestions"""
    return chat_assistant.suggest_questions(job_data)
