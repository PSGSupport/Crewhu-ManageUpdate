# get and post links

import csv
import json
import re
import base64
import requests
from pathlib import Path

# ==========
# CONFIG: FILE PATHS
# ==========
# Change these to match your actual uploaded file names.
CSV_FILE = Path("Lost Surveys(Survey History (5)) (1).csv")
JSON_FILE = Path("crewhu_notifications_NEW.json")

# ==========
# CONFIG: CONNECTWISE CREDS
# ==========
COMPANY_ID = "pearlsolves"
PUBLIC_KEY = "AVk7JkjVIiTjjnle"
PRIVATE_KEY = "HVy911UWKBMVGAwb"
CLIENT_ID = "43c39678-9ed1-4fd4-8ccf-db2d5a9dab10"

API_BASE = "https://api-na.myconnectwise.net/v2025_1/apis/3.0"


# ==========
# AUTH / HEADERS
# ==========
def build_headers():
    auth_string = f"{COMPANY_ID}+{PUBLIC_KEY}:{PRIVATE_KEY}"
    auth_base64 = base64.b64encode(auth_string.encode()).decode()
    return {
        "Authorization": f"Basic {auth_base64}",
        "clientId": CLIENT_ID,
        "Accept": "application/json",
        "Content-Type": "application/json",
    }


# ==========
# STEP 1: LOAD CSV TICKETS
# ==========
def load_ticket_numbers_from_csv(csv_path):
    ticket_numbers = set()

    with csv_path.open("r", newline="", encoding="utf-8-sig") as f:
        sample = f.read(2048)
        f.seek(0)
        try:
            dialect = csv.Sniffer().sniff(sample)
        except csv.Error:
            dialect = csv.get_dialect("excel")

        reader = csv.DictReader(f, dialect=dialect)

        possible_ticket_keys = ["Ticket#", "Ticket", "Ticket #", "ticket", "ticket#"]

        for row in reader:
            ticket_value = None
            for key in possible_ticket_keys:
                if key in row and row[key]:
                    ticket_value = row[key]
                    break

            if not ticket_value:
                continue

            ticket_value = str(ticket_value).strip()
            # Handle things like "497225.0"
            ticket_value = ticket_value.split(".")[0]

            if ticket_value.isdigit():
                ticket_numbers.add(ticket_value)

    return sorted(ticket_numbers, key=int)


# ==========
# STEP 2: LOAD JSON NOTIFICATIONS
# ==========
def load_notifications_from_json(json_path):
    with json_path.open("r", encoding="utf-8") as f:
        data = json.load(f)
    return data  # expecting a list of dicts


# ==========
# STEP 3: EXTRACT SURVEY LINK FOR A TICKET
# ==========
SURVEY_LINK_PATTERN = re.compile(
    r"https://web\.crewhu\.com/#/managesurvey/form/[^\s<>\"']+"
)

def get_survey_link_for_ticket(ticket_number, notifications):
    """
    Find the first Crewhu survey link for a given ticket number.
    We look for:
      - 'ticket# {ticket_number}' in FullBody
      - and a link containing '.../managesurvey/form/...'
    """
    ticket_pattern = f"ticket# {ticket_number}"

    for notif in notifications:
        full_body = notif.get("FullBody", "")
        if ticket_pattern not in full_body:
            continue

        match = SURVEY_LINK_PATTERN.search(full_body)
        if match:
            link = match.group(0).rstrip(">.")  # trim any trailing markup chars
            return link

    return None


# ==========
# STEP 4: UPDATE CONNECTWISE TICKET FIELD
# ==========
def update_ticket_crewhu_field(ticket_number, survey_link, headers):
    """
    For a given ConnectWise ticket:
      - GET the ticket
      - Find the 'Latest Crewhu Survey' custom field (or any field with 'crewhu' in caption)
      - PATCH its value to survey_link
    """
    ticket_id = int(ticket_number)
    url = f"{API_BASE}/service/tickets/{ticket_id}"

    # --- GET the ticket ---
    resp = requests.get(url, headers=headers)
    if resp.status_code != 200:
        print(f"[{ticket_number}] GET ticket failed: {resp.status_code} {resp.text[:200]}")
        return

    ticket = resp.json()
    custom_fields = ticket.get("customFields", [])

    if not custom_fields:
        print(f"[{ticket_number}] No customFields on ticket.")
        return

    # Find the Crewhu field (prefer explicit caption, otherwise any 'crewhu'-related)
    candidates = [
        f for f in custom_fields
        if f.get("caption", "").strip().lower() == "latest crewhu survey"
        or "crewhu" in f.get("caption", "").lower()
    ]

    if not candidates:
        print(f"[{ticket_number}] No Crewhu-related custom field found.")
        return

    latest_field = candidates[-1]
    idx = custom_fields.index(latest_field)

    patch_body = [
        {
            "op": "replace",
            "path": f"/customFields/{idx}/value",
            "value": survey_link,
        }
    ]

    patch_resp = requests.patch(url, headers=headers, json=patch_body)

    if patch_resp.status_code in (200, 204):
        print(f"[{ticket_number}]  Updated Crewhu field with: {survey_link}")
    else:
        print(f"[{ticket_number}]  PATCH failed: {patch_resp.status_code} {patch_resp.text[:200]}")


# ==========
# MAIN FLOW
# ==========
def main():
    headers = build_headers()

    print("Loading CSV ticket list...")
    tickets = load_ticket_numbers_from_csv(CSV_FILE)
    print(f"Found {len(tickets)} tickets in CSV: {tickets}")

    print("Loading Crewhu JSON notifications...")
    notifications = load_notifications_from_json(JSON_FILE)
    print(f"Loaded {len(notifications)} notification records from JSON.")

    processed = 0
    missing_link = 0
    api_failures = 0  # counted via messages

    for ticket_number in tickets:
        print(f"\n=== Processing ticket {ticket_number} ===")
        survey_link = get_survey_link_for_ticket(ticket_number, notifications)

        if not survey_link:
            print(f"[{ticket_number}] No Crewhu survey link found in JSON.")
            missing_link += 1
            continue

        print(f"[{ticket_number}] Found survey link: {survey_link}")
        update_ticket_crewhu_field(ticket_number, survey_link, headers)
        processed += 1

    print("\n======== SUMMARY ========")
    print(f"Total tickets in CSV: {len(tickets)}")
    print(f"Tickets with survey link found & attempted update: {processed}")
    print(f"Tickets with NO survey link found in JSON: {missing_link}")
    print("Check above logs for any API failure messages.")


if __name__ == "__main__":
    main()
