import requests
import time
import json
import sys

API_URL = "http://127.0.0.1:8000"

def log(msg):
    print(msg)
    sys.stdout.flush()

def run_test():
    log("Starting validation test...")
    
    # 1. Generate Model
    payload = {
        "symbol": "TATASTEEL",
        "exchange": "NSE",
        "forecast_years": 5,
        "model_types": ["dcf"]
    }
    
    try:
        log(f"Submitting job for {payload['symbol']}...")
        resp = requests.post(f"{API_URL}/api/model/generate", json=payload)
        if resp.status_code != 200:
            log(f"❌ Failed to start job: {resp.text}")
            return
            
        job_id = resp.json()["job_id"]
        log(f"Job started: {job_id}")
        
        # 2. Poll for completion
        max_retries = 40
        for i in range(max_retries):
            status_resp = requests.get(f"{API_URL}/api/job/{job_id}")
            if status_resp.status_code != 200:
                log(f"❌ Failed to get status: {status_resp.text}")
                return
                
            status_data = status_resp.json()
            status = status_data["status"]
            log(f"Status: {status}, Progress: {status_data.get('progress')}%")
            
            if status == "completed":
                log("Job completed!")
                break
            elif status == "failed":
                log(f"❌ Job failed: {status_data.get('message')}")
                return
                
            time.sleep(3)
        else:
            log("❌ Timeout waiting for job completion")
            return

        # 3. Test Sensitivity Endpoint (Check if valuation_data exists)
        log("\nTesting Sensitivity Endpoint...")
        sens_resp = requests.get(f"{API_URL}/api/analysis/sensitivity/{job_id}")
        if sens_resp.status_code == 200:
            data = sens_resp.json()
            if data.get("sensitivity"):
                log("✅ Sensitivity Data present")
            else:
                log("⚠️ Sensitivity Data empty (but 200 OK)")
        else:
            log(f"❌ Sensitivity Error: {sens_resp.status_code} - {sens_resp.text}")

        # 4. Test Chat Endpoint
        log("\nTesting Chat Endpoint...")
        chat_payload = {"message": "What is the implied share price?", "history": []}
        chat_resp = requests.post(f"{API_URL}/api/chat/{job_id}", json=chat_payload)
        if chat_resp.status_code == 200:
            log(f"✅ Chat Response: {chat_resp.json().get('response')[:50]}...")
        else:
            log(f"❌ Chat Error: {chat_resp.status_code} - {chat_resp.text}")
            
        # 5. Test PDF Export
        log(f"\nTesting PDF Export Endpoint /api/export/{job_id}?format=pdf...")
        pdf_resp = requests.get(f"{API_URL}/api/export/{job_id}?format=pdf")
        if pdf_resp.status_code == 200:
            content_type = pdf_resp.headers.get("Content-Type")
            log(f"✅ PDF Endpoint works! Content-Type: {content_type}, Size: {len(pdf_resp.content)} bytes")
        else:
            log(f"❌ PDF Endpoint Error: {pdf_resp.status_code}")

    except Exception as e:
        log(f"❌ Exception: {e}")

if __name__ == "__main__":
    run_test()
