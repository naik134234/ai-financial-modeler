import requests
import time
import json
import sys

API_URL = "http://127.0.0.1:8000"

def run_test():
    with open("repro.log", "w", encoding="utf-8") as f:
        sys.stdout = f
        sys.stderr = f
        
        print("Starting reproduction test...")
        
        # 1. Generate Model
        payload = {
            "symbol": "TATASTEEL",
            "exchange": "NSE",
            "forecast_years": 5,
            "model_types": ["dcf"]
        }
        
        try:
            resp = requests.post(f"{API_URL}/api/model/generate", json=payload)
            if resp.status_code != 200:
                print(f"❌ Failed to start job: {resp.text}")
                return
                
            job_id = resp.json()["job_id"]
            print(f"Job started: {job_id}")
            
            # 2. Poll for completion
            max_retries = 60 # Increased retries
            for i in range(max_retries):
                status_resp = requests.get(f"{API_URL}/api/job/{job_id}")
                if status_resp.status_code != 200:
                    print(f"❌ Failed to get status: {status_resp.text}")
                    return
                    
                status_data = status_resp.json()
                status = status_data["status"]
                print(f"Status: {status}, Progress: {status_data.get('progress')}%")
                
                if status == "completed":
                    print("Job completed!")
                    break
                elif status == "failed":
                    print(f"❌ Job failed: {status_data.get('message')}")
                    return
                    
                time.sleep(2)
            else:
                print("❌ Timeout waiting for job completion")
                return

            # 3. Test Sensitivity Endpoint
            print("\nTesting Sensitivity Endpoint...")
            sens_resp = requests.get(f"{API_URL}/api/analysis/sensitivity/{job_id}")
            if sens_resp.status_code == 200:
                print(f"✅ Sensitivity Data: {json.dumps(sens_resp.json())[:100]}...")
            else:
                print(f"❌ Sensitivity Error: {sens_resp.status_code} - {sens_resp.text}")

            # 4. Test Chat Endpoint
            print("\nTesting Chat Endpoint...")
            chat_payload = {"message": "What is the implied share price?", "history": []}
            chat_resp = requests.post(f"{API_URL}/api/chat/{job_id}", json=chat_payload)
            if chat_resp.status_code == 200:
                print(f"✅ Chat Response: {chat_resp.json()}")
            else:
                print(f"❌ Chat Error: {chat_resp.status_code} - {chat_resp.text}")
                
            # 5. Test PDF Export (guessing endpoint)
            print("\nTesting PDF Export Endpoint (Guessing /api/export/pdf/{job_id})...")
            pdf_resp = requests.get(f"{API_URL}/api/export/pdf/{job_id}")
            if pdf_resp.status_code == 200:
                print("✅ PDF Endpoint found (unexpected!)")
            else:
                print(f"❌ PDF Endpoint Error: {pdf_resp.status_code}")

        except Exception as e:
            print(f"❌ Exception: {e}")

if __name__ == "__main__":
    run_test()
