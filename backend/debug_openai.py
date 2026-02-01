
import os
import requests
from dotenv import load_dotenv

load_dotenv()

api_key = os.getenv("OPENAI_API_KEY")
print(f"API Key found: {api_key[:10]}...{api_key[-5:] if api_key else 'None'}")
print(f"Key length: {len(api_key) if api_key else 0}")

if not api_key:
    print("❌ No API Key found")
    exit(1)

url = "https://api.openai.com/v1/chat/completions"
headers = {
    "Authorization": f"Bearer {api_key}",
    "Content-Type": "application/json"
}
data = {
    "model": "gpt-4o",
    "messages": [{"role": "user", "content": "Hello"}],
    "max_tokens": 10
}

print(f"Sending request to {url}...")
try:
    resp = requests.post(url, headers=headers, json=data, timeout=30)
    print(f"Status: {resp.status_code}")
    print(f"Response: {resp.text}")
except Exception as e:
    print(f"Exception: {e}")
