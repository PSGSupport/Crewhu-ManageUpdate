"""
Post Crewhu survey feedback as internal notes on ConnectWise tickets.

Reads parsed survey data from JSON and creates internal analysis notes
on the corresponding tickets.
"""

import json
import requests
from datetime import datetime, timezone
from pathlib import Path

from connectwise_api import (
    require_credentials,
    get_headers,
    ticket_notes_url
)

# ==============================
# CONFIG
# ==============================
DRY_RUN = False  # Set to True to preview changes without making them

PARSED_JSON = Path("crewhu_surveys_clean.json")


# ==============================
# DELETE OLD AUTOMATED NOTES
# ==============================
def delete_automated_notes(ticket_id, headers):
    """
    Find and delete any existing notes created by this automation
    to prevent duplicates.
    """
    if DRY_RUN:
        print(f"[{ticket_id}] (DRY RUN) Would delete old automation notes.")
        return 0

    notes_url = ticket_notes_url(ticket_id)

    try:
        resp = requests.get(notes_url, headers=headers)
        if resp.status_code != 200:
            print(f"[{ticket_id}] Failed to fetch notes: {resp.status_code}")
            return 0

        notes = resp.json()
        deleted = 0

        for note in notes:
            note_id = note.get("id")
            text = (note.get("text") or "")

            # Delete notes created by our script to prevent duplicates
            if "just gave a" in text and "Customer feedback:" in text:
                del_url = ticket_notes_url(ticket_id, note_id)
                d = requests.delete(del_url, headers=headers)
                if d.status_code in (200, 204):
                    print(f"[{ticket_id}] Deleted old auto note {note_id}")
                    deleted += 1

        if deleted == 0:
            print(f"[{ticket_id}] No old automation notes found.")

        return deleted

    except Exception as e:
        print(f"[{ticket_id}] Error during deletion: {e}")
        return 0


# ==============================
# POST NEW NOTE (INTERNAL)
# ==============================
def post_note(ticket_id, summary, feedback, headers):
    """Post a new internal analysis note to a ticket."""
    url = ticket_notes_url(ticket_id)

    # Construct the final note text
    final_note_text = f"{summary}\n\nCustomer feedback:\n{feedback}"

    if DRY_RUN:
        print(f"\n[{ticket_id}] === DRY RUN NOTE (INTERNAL) ===")
        print(final_note_text)
        print("=" * 40 + "\n")
        return True

    payload = {
        "text": final_note_text,
        "detailDescriptionFlag": False,  # False = Do not put in Discussion
        "internalAnalysisFlag": True,    # True = Put in Internal Analysis
        "resolutionFlag": False,         # False = Do not put in Resolution
        "createdBy": "Crewhu API",
        "dateCreated": datetime.now(timezone.utc).isoformat()
    }

    try:
        response = requests.post(url, headers=headers, json=payload)

        if response.status_code == 201:
            print(f"[{ticket_id}] Posted internal note")
            return True
        else:
            print(f"[{ticket_id}] Error posting note: {response.status_code} {response.text[:200]}")
            return False
    except Exception as e:
        print(f"[{ticket_id}] Connection Error: {e}")
        return False


# ==============================
# MAIN
# ==============================
def main():
    # Validate credentials
    require_credentials()

    headers = get_headers()

    if not PARSED_JSON.exists():
        print(f"ERROR: {PARSED_JSON} not found.")
        print("Please run reformatJSON.py first to generate this file.")
        return

    with PARSED_JSON.open("r", encoding="utf-8") as f:
        parsed_data = json.load(f)

    print(f"Loaded {len(parsed_data)} parsed Crewhu entries.")

    if DRY_RUN:
        print("!!! DRY RUN MODE - No changes will be made !!!\n")

    # Track statistics
    posted = 0
    errors = 0

    for entry in parsed_data:
        ticket_id = entry.get("ticket_number")
        summary = entry.get("summary", "").strip()
        feedback = entry.get("customer_feedback", "No feedback provided.").strip()

        if not ticket_id:
            print("Skipping entry with missing ticket number.")
            continue

        print(f"\n=== Processing ticket {ticket_id} ===")

        # Delete old automated notes (prevents duplicates)
        delete_automated_notes(ticket_id, headers)

        # Post the new note
        if post_note(ticket_id, summary, feedback, headers):
            posted += 1
        else:
            errors += 1

    # Summary
    print("\n" + "=" * 40)
    print("SUMMARY")
    print("=" * 40)
    print(f"Total entries: {len(parsed_data)}")
    print(f"Notes posted:  {posted}")
    print(f"Errors:        {errors}")

    if DRY_RUN:
        print("\nDRY RUN COMPLETE - Set DRY_RUN = False to make changes.")


if __name__ == "__main__":
    main()
