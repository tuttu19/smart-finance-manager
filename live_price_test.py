import requests

# ✅ Replace these with your actual Capital.com credentials
CAPITAL_COM_API_KEY = "FezXw3h3IIxE7IUJ"  # Your API Key
EMAIL = "tuttu19@gmail.com"                 # Your login email
PASSWORD = "Rahim@2014!"          # Your login password

symbol = "AAPL"

# Step 1: Authenticate with email & password
auth_url = "https://api-capital.backend-capital.com/api/v1/session"
headers = {
    "X-CAP-API-KEY": CAPITAL_COM_API_KEY,
    "Content-Type": "application/json"
}
payload = {
    "identifier": EMAIL,
    "password": PASSWORD
}

auth_response = requests.post(auth_url, headers=headers, json=payload)

if auth_response.status_code == 200:
    cst = auth_response.headers.get("CST")
    x_security_token = auth_response.headers.get("X-SECURITY-TOKEN")

    if not cst or not x_security_token:
        print("❌ Failed to extract security tokens.")
    else:
        # Step 2: Fetch price data
        price_url = f"https://api-capital.backend-capital.com/api/v1/prices/{symbol}"
        price_headers = {
            "CST": cst,
            "X-SECURITY-TOKEN": x_security_token,
            "X-CAP-API-KEY": CAPITAL_COM_API_KEY
        }
        price_response = requests.get(price_url, headers=price_headers)

        if price_response.status_code == 200:
            price_data = price_response.json()
            print(f"🔔 Latest price data for {symbol}:")
            print(price_data)
        else:
            print(f"❌ Failed to fetch price data. Status code: {price_response.status_code}")
            print(price_response.text)
else:
    print(f"❌ Authentication failed. Status code: {auth_response.status_code}")
    print(auth_response.text)
