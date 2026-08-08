"""
setup_airports.py — One-time addition: Airports worksheet for the main RG donation form
==========================================================================================
Adds an "Airports" worksheet (single column: AirportName) to the existing
"TMO Admin" Google Sheet, seeded with the airport list the form used to have
hardcoded. From then on, editing this worksheet directly (add/remove/rename a
row) changes the Airport Name dropdown in the form — no code change needed.

Idempotent — skipped if the "Airports" worksheet already exists.

Prerequisites:
    You've already run `python setup.py` (sheets_config.json with admin_sheet_id,
    and token.pickle, must already exist).

Run once:
    python setup_airports.py
"""

import json
import os
import pickle
import sys

try:
    import gspread
    from google.auth.transport.requests import Request
    from google_auth_oauthlib.flow import InstalledAppFlow
except ImportError:
    print("Missing packages. Run:")
    print("  python -m pip install gspread google-auth google-auth-oauthlib")
    sys.exit(1)

CLIENT_SECRET = "client_secret.json"
TOKEN_FILE = "token.pickle"
CONFIG_FILE = "sheets_config.json"
AIRPORTS_WORKSHEET = "Airports"
SEED_AIRPORTS = ["Vizag", "Coimbatore", "Trichy", "Goa", "Jaipur", "Indore", "Ahmedabad", "Chennai"]
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


def get_credentials():
    """Authenticate via OAuth — opens a browser on first run, caches the token after."""
    creds = None
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, "rb") as f:
            creds = pickle.load(f)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(CLIENT_SECRET):
                print(f"ERROR: {CLIENT_SECRET} not found in this folder.")
                sys.exit(1)
            flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, "wb") as f:
            pickle.dump(creds, f)
    return creds


def main():
    if not os.path.exists(CONFIG_FILE):
        print(f"ERROR: {CONFIG_FILE} not found. Run `python setup.py` first.")
        sys.exit(1)
    with open(CONFIG_FILE) as f:
        config = json.load(f)
    if "admin_sheet_id" not in config:
        print(f"ERROR: 'admin_sheet_id' missing from {CONFIG_FILE}. Run `python setup.py` first.")
        sys.exit(1)

    print("Authenticating with your Google account...\n")
    creds = get_credentials()
    gc = gspread.authorize(creds)
    print("Signed in successfully!\n")

    admin_sh = gc.open_by_key(config["admin_sheet_id"])
    existing_titles = [ws.title for ws in admin_sh.worksheets()]
    if AIRPORTS_WORKSHEET in existing_titles:
        print(f"'{AIRPORTS_WORKSHEET}' worksheet already exists in the Admin sheet — nothing to do.")
        sys.exit(0)

    print(f"Creating '{AIRPORTS_WORKSHEET}' worksheet...")
    airports_ws = admin_sh.add_worksheet(AIRPORTS_WORKSHEET, rows=100, cols=1)
    airports_ws.update([["AirportName"]] + [[a] for a in SEED_AIRPORTS], value_input_option="RAW")
    print(f"  Airports: {len(SEED_AIRPORTS)} rows uploaded")

    print("\n" + "=" * 60)
    print("SETUP COMPLETE!")
    print("=" * 60)
    print(f"\n  Admin sheet: https://docs.google.com/spreadsheets/d/{admin_sh.id}")


if __name__ == "__main__":
    main()
