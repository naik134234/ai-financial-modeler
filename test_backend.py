import requests
import sys
import time

def check_health():
    print("Checking backend health...")
    try:
        response = requests.get("http://127.0.0.1:8000/api/health")
        if response.status_code == 200:
            print("Backend is HEALTHY")
            print(response.json())
            return True
        else:
            print(f"Backend returned status check {response.status_code}")
            return False
    except Exception as e:
        print(f"Failed to connect to backend: {e}")
        return False

if __name__ == "__main__":
    success = check_health()
    if not success:
        sys.exit(1)
