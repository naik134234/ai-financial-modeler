
import os
import google.generativeai as genai
from dotenv import load_dotenv

load_dotenv()

api_key = os.getenv("GEMINI_API_KEY")
print(f"Gemini API Key found: {api_key[:10]}...{api_key[-5:] if api_key else 'None'}")

if not api_key:
    print("❌ No Gemini API Key found")
    exit(1)

genai.configure(api_key=api_key)
model = genai.GenerativeModel('gemini-pro')

print("Sending request to Gemini...")
try:
    response = model.generate_content("Hello, reflect back this message.")
    print(f"Response: {response.text}")
except Exception as e:
    print(f"Exception: {e}")
