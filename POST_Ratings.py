"""
Update ConnectWise tickets with Crewhu ratings.

Reads ratings from a CSV file and updates the "Latest Crewhu Rating"
custom field on matching tickets.
"""

import requests
import pandas as pd
from pathlib import Path

from connectwise_api import (
    require_credentials,
    get_headers,
    tickets_url,
    API_BASE
)

# ==========================================
# CONFIGURATION
# ==========================================
DRY_RUN = False  # Set to True to preview changes without making them

# CSV file path - change this to match your file location
CSV_FILE = Path("Lost Surveys(Survey History (5)) (1).csv")


# ==========================================
# RATING MAPPING
# ==========================================
RATING_MAP = {
    "AWESOME": "Positive",
    # Add more mappings here if needed
    # "GOOD": "Positive",
    # "BAD": "Negative",
}


# ==========================================
# UPDATE LOGIC
# ==========================================
def update_ticket_rating(ticket_id, rating_value, headers):
    """Update the Latest Crewhu Rating custom field on a ticket."""
    url = tickets_url(ticket_id)

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

    if DRY_RUN:
        print(f"[{ticket_id}] (DRY RUN) Would set rating to: {rating_value}")
        return True

    try:
        response = requests.patch(url, headers=headers, json=payload)

        if response.status_code == 200:
            print(f"[{ticket_id}] Updated rating to: {rating_value}")
            return True
        else:
            print(f"[{ticket_id}] ERROR: {response.status_code} - {response.text[:200]}")
            return False

    except requests.exceptions.RequestException as e:
        print(f"[{ticket_id}] NETWORK ERROR: {e}")
        return False


# ==========================================
# MAIN
# ==========================================
def main():
    # Validate credentials before starting
    require_credentials()

    # Check CSV file exists
    if not CSV_FILE.exists():
        print(f"ERROR: CSV file not found: {CSV_FILE}")
        print("Please update CSV_FILE path in the script.")
        return

    headers = get_headers()

    # Load CSV
    try:
        df = pd.read_csv(CSV_FILE)
    except Exception as e:
        print(f"ERROR reading CSV: {e}")
        return

    print(f"Loaded {len(df)} rows from {CSV_FILE}")

    if DRY_RUN:
        print("!!! DRY RUN MODE - No changes will be made !!!\n")

    # Track statistics
    updated = 0
    skipped = 0
    errors = 0

    # Process each row
    for index, row in df.iterrows():
        ticket_id = str(row['Ticket#']).strip()

        # Handle float ticket IDs like "497225.0"
        if '.' in ticket_id:
            ticket_id = ticket_id.split('.')[0]

        rating = str(row['Rating']).strip().upper()

        # Map rating to value
        if rating not in RATING_MAP:
            print(f"[{ticket_id}] Skipping - Rating '{rating}' not mapped")
            skipped += 1
            continue

        rating_value = RATING_MAP[rating]

        if update_ticket_rating(ticket_id, rating_value, headers):
            updated += 1
        else:
            errors += 1

    # Summary
    print("\n" + "=" * 40)
    print("SUMMARY")
    print("=" * 40)
    print(f"Total rows:  {len(df)}")
    print(f"Updated:     {updated}")
    print(f"Skipped:     {skipped}")
    print(f"Errors:      {errors}")

    if DRY_RUN:
        print("\nDRY RUN COMPLETE - Set DRY_RUN = False to make changes.")


if __name__ == "__main__":
    main()
