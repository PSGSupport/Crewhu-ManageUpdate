"""
Shared ConnectWise API utilities for Crewhu integration scripts.

This module provides common functionality for authenticating and
interacting with the ConnectWise Manage API.
"""

import os
import base64
import sys
from dotenv import load_dotenv

load_dotenv()

# ==========================================
# CONFIGURATION
# ==========================================
COMPANY_ID = os.environ.get("CW_COMPANY_ID")
PUBLIC_KEY = os.environ.get("CW_PUBLIC_KEY")
PRIVATE_KEY = os.environ.get("CW_PRIVATE_KEY")
CLIENT_ID = os.environ.get("CW_CLIENT_ID")
API_BASE = os.environ.get("CW_API_BASE", "https://na.myconnectwise.net/v4_6_release/apis/3.0")


# ==========================================
# CREDENTIAL VALIDATION
# ==========================================
def validate_credentials():
    """
    Validate that all required environment variables are set.
    Returns tuple (is_valid, missing_vars).
    """
    required = {
        "CW_COMPANY_ID": COMPANY_ID,
        "CW_PUBLIC_KEY": PUBLIC_KEY,
        "CW_PRIVATE_KEY": PRIVATE_KEY,
        "CW_CLIENT_ID": CLIENT_ID,
    }

    missing = [name for name, value in required.items() if not value]

    return (len(missing) == 0, missing)


def require_credentials():
    """
    Validate credentials and exit with error if any are missing.
    Call this at the start of scripts that need API access.
    """
    is_valid, missing = validate_credentials()

    if not is_valid:
        print("ERROR: Missing required environment variables:")
        for var in missing:
            print(f"  - {var}")
        print("\nPlease set these in your .env file or environment.")
        sys.exit(1)


# ==========================================
# AUTHENTICATION
# ==========================================
def get_headers():
    """
    Build and return the authentication headers for ConnectWise API calls.
    """
    auth_string = f"{COMPANY_ID}+{PUBLIC_KEY}:{PRIVATE_KEY}"
    auth_base64 = base64.b64encode(auth_string.encode()).decode()

    return {
        "Authorization": f"Basic {auth_base64}",
        "clientId": CLIENT_ID,
        "Accept": "application/json",
        "Content-Type": "application/json"
    }


# ==========================================
# API URL HELPERS
# ==========================================
def tickets_url(ticket_id=None):
    """Get the URL for tickets endpoint."""
    if ticket_id:
        return f"{API_BASE}/service/tickets/{ticket_id}"
    return f"{API_BASE}/service/tickets"


def ticket_notes_url(ticket_id, note_id=None):
    """Get the URL for ticket notes endpoint."""
    base = f"{API_BASE}/service/tickets/{ticket_id}/notes"
    if note_id:
        return f"{base}/{note_id}"
    return base
