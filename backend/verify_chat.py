import requests
import time
import json

API_URL = "http://127.0.0.1:8000"

def verify_chat():
    print("1. Creating dummy job...")
    payload = {
        "symbol": "MSFT",
        "exchange": "NASDAQ",
        "forecast_years": 5,
        "model_types": ["dcf"]
    }
    try:
        resp = requests.post(f"{API_URL}/api/model/generate", json=payload)
        if resp.status_code != 200:
            print(f"Failed to create job: {resp.status_code} {resp.text}")
            return
        
        job_id = resp.json().get("job_id")
        print(f"Job created: {job_id}")
        
        # Wait a bit for job to initialize (though chat should work even if pending, usually)
        time.sleep(2)
        
        print("2. Sending chat message...")
        chat_payload = {
            "message": "What is the company's revenue growth?",
            "history": []
        }
        
        chat_resp = requests.post(f"{API_URL}/api/chat/{job_id}", json=chat_payload)
        if chat_resp.status_code == 200:
            print("✅ Chat Success!")
            print("Response:", json.dumps(chat_resp.json(), indent=2))
        else:
            print(f"❌ Chat Failed: {chat_resp.status_code}")
            print(chat_resp.text)
            
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    verify_chat()
