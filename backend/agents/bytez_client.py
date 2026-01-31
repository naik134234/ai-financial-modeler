
import os
import logging
from typing import Optional, Dict, Any, List

logger = logging.getLogger(__name__)

class BytezClient:
    """Wrapper for Bytez API"""
    
    def __init__(self, api_key: str = None):
        self.api_key = api_key or os.environ.get("BYTEZ_API_KEY")
        self.client = None
        self.model = None
        
        if self.api_key:
            try:
                from bytez import Bytez
                self.client = Bytez(self.api_key)
                # Initialize default model
                self.model = self.client.model("anthropic/claude-opus-4-5")
                logger.info("Bytez client initialized with model: anthropic/claude-opus-4-5")
            except ImportError:
                logger.error("bytez package not installed. Run: pip install bytez")
            except Exception as e:
                logger.error(f"Failed to initialize Bytez client: {e}")
        else:
            logger.warning("No Bytez API key found")

    def generate_content(self, prompt: str, system_prompt: str = None) -> Optional[str]:
        """
        Generate content using Bytez API
        
        Args:
            prompt: User prompt
            system_prompt: Optional system prompt
            
        Returns:
            Generated text content or None on failure
        """
        if not self.client or not self.model:
            logger.warning("Bytez client not initialized")
            return None
            
        try:
            messages = []
            if system_prompt:
                # Bytez/Claude usually accepts system prompt differently or as part of context, 
                # but for simplicity we'll prepend it or include it if the API supports it.
                # Based on the user snippet, it takes a list of messages.
                # We'll stick to a simple user message for now, or prepend system prompt.
                messages.append({
                    "role": "user",
                    "content": f"System: {system_prompt}\n\nUser: {prompt}"
                })
            else:
                messages.append({
                    "role": "user",
                    "content": prompt
                })
                
            logger.info("Sending request to Bytez API...")
            results = self.model.run(messages)
            
            if results.error:
                logger.error(f"Bytez API returned error: {results.error}")
                return None
                
            return results.output
            
        except Exception as e:
            logger.error(f"Error calling Bytez API: {e}")
            return None

# Singleton instance
_bytez_client = None

def get_bytez_client():
    global _bytez_client
    if _bytez_client is None:
        _bytez_client = BytezClient()
    return _bytez_client
