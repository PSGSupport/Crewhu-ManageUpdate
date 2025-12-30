import requests
import base64
import os
import pandas as pd
from dotenv import load_dotenv

load_dotenv()

# === CSV File Path ===
# Change this to the actual file location
CSV_FILE_PATH = r"/content/Lost Surveys(Survey History (5)) (1).csv"

# === ConnectWise Credentials (from environment variables) ===
COMPANY_ID = os.environ.get("CW_COMPANY_ID")
PUBLIC_KEY = os.environ.get("CW_PUBLIC_KEY")
PRIVATE_KEY = os.environ.get("CW_PRIVATE_KEY")
CLIENT_ID = os.environ.get("CW_CLIENT_ID")

API_BASE = os.environ.get("CW_API_BASE", "https://na.myconnectwise.net/v4_6_release/apis/3.0")

# === AUTH HEADER ===
auth_string = f"{COMPANY_ID}+{PUBLIC_KEY}:{PRIVATE_KEY}"
auth_base64 = base64.b64encode(auth_string.encode()).decode()

headers = {
    "Authorization": f"Basic {auth_base64}",
    "clientId": CLIENT_ID,
    "Content-Type": "application/json",
    "Accept": "application/json"
}

# === LOAD CSV FILE ===
df = pd.read_csv('/content/Lost Surveys(Survey History (5)) (1).csv')

print("Starting ticket update process...")

# === PROCESS EACH ROW ===
for index, row in df.iterrows():
    ticket_id = str(row['Ticket#']).strip()
    rating = str(row['Rating']).strip()

    # Only send "AWESOME" as "Positive"
    if rating.upper() == "AWESOME":
        rating_value = "Positive"
    else:
        print(f"Skipping ticket {ticket_id} — Rating is not AWESOME ({rating})")
        continue

    url = f"{API_BASE}/service/tickets/{ticket_id}"

    payload = [
        {
            "op": "replace",
            "path": "customFields",
            "value": [
                {
                    "id": 18,
                    "caption": "Latest Crewhu Rating",
                    "type": "Text",
                    "entryMethod": "EntryField",
                    "numberOfDecimals": 0,
                    "value": rating_value
                }
            ]
        }
    ]

    try:
        response = requests.patch(url, headers=headers, json=payload)

        if response.status_code == 200:
            print(f" Updated Ticket {ticket_id} — Rating set to {rating_value}")
        else:
            print(f" ERROR updating Ticket {ticket_id}: {response.status_code} — {response.text}")

    except requests.exceptions.RequestException as e:
        print(f" NETWORK ERROR on Ticket {ticket_id}: {e}")

print("\nFinished processing tickets.")
