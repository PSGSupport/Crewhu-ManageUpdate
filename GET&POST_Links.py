# get and post links
"""
Update ConnectWise tickets with Crewhu survey links.

Reads ticket numbers from a CSV file, finds matching survey links
from Crewhu notification JSON, and updates the ticket's custom field.
"""

import csv
import json
import re
import requests
from pathlib import Path

from connectwise_api import (
    require_credentials,
    get_headers,
    tickets_url,
    API_BASE
)

# ==========
# CONFIG
# ==========
DRY_RUN = False  # Set to True to preview changes without making them

# File paths - change these to match your actual files
CSV_FILE = Path("Lost Surveys(Survey History (5)) (1).csv")
JSON_FILE = Path("crewhu_notifications_NEW.json")


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
    url = tickets_url(ticket_id)

    # --- GET the ticket ---
    resp = requests.get(url, headers=headers)
    if resp.status_code != 200:
        print(f"[{ticket_number}] GET ticket failed: {resp.status_code} {resp.text[:200]}")
        return False

    ticket = resp.json()
    custom_fields = ticket.get("customFields", [])

    if not custom_fields:
        print(f"[{ticket_number}] No customFields on ticket.")
        return False

    # Find the Crewhu field (prefer explicit caption, otherwise any 'crewhu'-related)
    candidates = [
        f for f in custom_fields
        if f.get("caption", "").strip().lower() == "latest crewhu survey"
        or "crewhu" in f.get("caption", "").lower()
    ]

    if not candidates:
        print(f"[{ticket_number}] No Crewhu-related custom field found.")
        return False

    latest_field = candidates[-1]
    idx = custom_fields.index(latest_field)

    patch_body = [
        {
            "op": "replace",
            "path": f"/customFields/{idx}/value",
            "value": survey_link,
        }
    ]

    if DRY_RUN:
        print(f"[{ticket_number}] (DRY RUN) Would update field to: {survey_link}")
        return True

    patch_resp = requests.patch(url, headers=headers, json=patch_body)

    if patch_resp.status_code in (200, 204):
        print(f"[{ticket_number}] Updated Crewhu field with: {survey_link}")
        return True
    else:
        print(f"[{ticket_number}] PATCH failed: {patch_resp.status_code} {patch_resp.text[:200]}")
        return False


# ==========
# MAIN FLOW
# ==========
def main():
    # Validate credentials
    require_credentials()

    headers = get_headers()

    # Check files exist
    if not CSV_FILE.exists():
        print(f"ERROR: CSV file not found: {CSV_FILE}")
        return

    if not JSON_FILE.exists():
        print(f"ERROR: JSON file not found: {JSON_FILE}")
        return

    print("Loading CSV ticket list...")
    tickets = load_ticket_numbers_from_csv(CSV_FILE)
    print(f"Found {len(tickets)} tickets in CSV")

    print("Loading Crewhu JSON notifications...")
    notifications = load_notifications_from_json(JSON_FILE)
    print(f"Loaded {len(notifications)} notification records from JSON.")

    if DRY_RUN:
        print("\n!!! DRY RUN MODE - No changes will be made !!!\n")

    processed = 0
    missing_link = 0
    errors = 0

    for ticket_number in tickets:
        print(f"\n=== Processing ticket {ticket_number} ===")
        survey_link = get_survey_link_for_ticket(ticket_number, notifications)

        if not survey_link:
            print(f"[{ticket_number}] No Crewhu survey link found in JSON.")
            missing_link += 1
            continue

        print(f"[{ticket_number}] Found survey link: {survey_link}")
        if update_ticket_crewhu_field(ticket_number, survey_link, headers):
            processed += 1
        else:
            errors += 1

    print("\n" + "=" * 40)
    print("SUMMARY")
    print("=" * 40)
    print(f"Total tickets in CSV: {len(tickets)}")
    print(f"Successfully updated: {processed}")
    print(f"No survey link found: {missing_link}")
    print(f"Errors: {errors}")

    if DRY_RUN:
        print("\nDRY RUN COMPLETE - Set DRY_RUN = False to make changes.")


if __name__ == "__main__":
    main()
