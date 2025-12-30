import json
import base64
import os
import requests
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

# ==========================================
# CONFIGURATION
# ==========================================
DRY_RUN = False  # <--- Set to False to actually delete the notes

INPUT_FILE = Path("crewhu_surveys_clean.json")

# ConnectWise Credentials (from environment variables)
COMPANY_ID = os.environ.get("CW_COMPANY_ID")
PUBLIC_KEY = os.environ.get("CW_PUBLIC_KEY")
PRIVATE_KEY = os.environ.get("CW_PRIVATE_KEY")
CLIENT_ID = os.environ.get("CW_CLIENT_ID")

API_BASE = os.environ.get("CW_API_BASE", "https://na.myconnectwise.net/v4_6_release/apis/3.0")

# ==========================================
# AUTHENTICATION
# ==========================================
def get_headers():
    auth_string = f"{COMPANY_ID}+{PUBLIC_KEY}:{PRIVATE_KEY}"
    auth_base64 = base64.b64encode(auth_string.encode()).decode()
    return {
        "Authorization": f"Basic {auth_base64}",
        "clientId": CLIENT_ID,
        "Accept": "application/json",
        "Content-Type": "application/json"
    }

# ==========================================
# DELETE LOGIC
# ==========================================
def process_ticket_deletion(ticket_id, headers):
    notes_url = f"{API_BASE}/service/tickets/{ticket_id}/notes"

    try:
        # 1. Get all notes for the ticket
        response = requests.get(notes_url, headers=headers)

        if response.status_code == 404:
            print(f"[{ticket_id}] Ticket not found.")
            return
        elif response.status_code != 200:
            print(f"[{ticket_id}] Error fetching notes: {response.status_code}")
            return

        notes = response.json()
        notes_deleted = 0

        # 2. Loop through notes to find matches
        for note in notes:
            note_id = note.get('id')
            text = note.get('text', '')

            # SAFETY CHECK: Only delete if it matches the specific format generated previously
            if "just gave a" in text and "Customer feedback:" in text:

                if DRY_RUN:
                    print(f"[{ticket_id}] [DRY RUN] Would delete Note ID {note_id}:")
                    print(f"    Content: {text[:50]}...")
                else:
                    # Perform Deletion
                    del_url = f"{notes_url}/{note_id}"
                    del_resp = requests.delete(del_url, headers=headers)

                    if del_resp.status_code in [200, 204]:
                        print(f"[{ticket_id}] 🗑️ Deleted Note ID {note_id}")
                        notes_deleted += 1
                    else:
                        print(f"[{ticket_id}] ❌ Failed to delete Note {note_id}: {del_resp.status_code}")

        if notes_deleted == 0 and not DRY_RUN:
            print(f"[{ticket_id}] No matching Crewhu notes found.")

    except Exception as e:
        print(f"[{ticket_id}] Connection error: {e}")

# ==========================================
# MAIN EXECUTION
# ==========================================
def main():
    if not INPUT_FILE.exists():
        print(f"Error: {INPUT_FILE} not found. Make sure the JSON file is in the same folder.")
        return

    headers = get_headers()

    with open(INPUT_FILE, 'r', encoding='utf-8') as f:
        data = json.load(f)

    print(f"--- Starting Cleanup Process for {len(data)} tickets ---")
    if DRY_RUN:
        print("!!! DRY RUN MODE ACTIVE: No data will be deleted !!!\n")

    for entry in data:
        ticket_id = entry.get('ticket_number')
        if ticket_id:
            process_ticket_deletion(ticket_id, headers)

    if DRY_RUN:
        print("\n--- Dry Run Complete ---")
        print("Set DRY_RUN = False in the script to perform actual deletions.")
    else:
        print("\n--- Deletion Process Complete ---")

if __name__ == "__main__":
    main()
