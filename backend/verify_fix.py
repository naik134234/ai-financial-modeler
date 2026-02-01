import requests
import json
import time

API_URL = "http://127.0.0.1:8000"

def test_us_stock_data_fetching():
    print("\n[Testing US Stock Data Fetching]")
    try:
        from data.yahoo_finance import get_stock_info, get_historical_financials
        
        # Test 1: Basic Info for AAPL
        print("Fetching info for AAPL...")
        info = get_stock_info("AAPL", exchange="NASDAQ")
        if info and info.get("symbol") == "AAPL":
            print(f"✅ Success: Got info for {info['name']} ({info['symbol']}) - Price: {info.get('current_price')}")
        else:
            print(f"❌ Failed: Could not get info for AAPL. Result: {info}")

        # Test 2: Historical Data for TSLA
        print("\nFetching financials for TSLA...")
        financials = get_historical_financials("TSLA", exchange="NASDAQ")
        if financials and financials.get("income_statement"):
            print("✅ Success: Got historical financials for TSLA")
            print(f"   Shape: {len(financials['income_statement'].get('revenue', []))} years of data")
        else:
            print(f"❌ Failed: Could not get financials for TSLA. Result type: {type(financials)}")
            
    except ImportError:
        print("⚠️ Could not import backend modules directly. Skipping internal test.")
    except Exception as e:
        print(f"❌ Error during internal test: {e}")

def test_api_endpoints():
    print("\n[Testing API Endpoints]")
    try:
        # Test 1: US Stocks Search
        print("Testing GET /api/stocks/us?search=APP...")
        resp = requests.get(f"{API_URL}/api/stocks/us", params={"search": "APP", "limit": 5})
        if resp.status_code == 200:
            data = resp.json()
            stocks = data.get("stocks", [])
            print(f"✅ Success: Found {len(stocks)} stocks matching 'APP'")
            for s in stocks:
                print(f"   - {s['symbol']}: {s['name']}")
        else:
            print(f"❌ Failed: API returned {resp.status_code}")

        # Test 2: Generate Model for MSFT
        print("\nTesting POST /api/model/generate for MSFT...")
        payload = {
            "symbol": "MSFT",
            "exchange": "NASDAQ",
            "forecast_years": 5,
            "model_types": ["dcf"]
        }
        resp = requests.post(f"{API_URL}/api/model/generate", json=payload)
        if resp.status_code == 200:
            data = resp.json()
            job_id = data.get("job_id")
            print(f"✅ Success: Job started with ID {job_id}")
            
            # Optional: Poll status briefly
            print("   Polling status...")
            for _ in range(3):
                time.sleep(2)
                status_resp = requests.get(f"{API_URL}/api/jobs/{job_id}")
                if status_resp.status_code == 200:
                    status_data = status_resp.json()
                    print(f"   Status: {status_data.get('status')} - Progress: {status_data.get('progress')}%")
                    if status_data.get('status') == 'completed':
                        break
        else:
            print(f"❌ Failed: API returned {resp.status_code} - {resp.text}")

    except requests.exceptions.ConnectionError:
        print("❌ Failed: Could not connect to backend server. Is it running?")
    except Exception as e:
        print(f"❌ Error during API test: {e}")

if __name__ == "__main__":
    test_us_stock_data_fetching()
    test_api_endpoints()
