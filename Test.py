import requests
from datetime import datetime
import pytz  # if you need timezone

# --- Step 1: Get access token ---
token_url = "https://tools.uniuni.com:8888/token"
form_data = {
    "username": "uniuni",
    "password": "Uniexpress!@!@"
}

token_response = requests.post(token_url, data=form_data)

# Check if request succeeded
if token_response.status_code != 200:
    print("Failed to get token:", token_response.text)
    exit()

token_json = token_response.json()
access_token = token_json['access_token']
print("Access token:", access_token)