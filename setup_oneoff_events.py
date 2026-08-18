"""
setup_oneoff_events.py — One-time addition: Events worksheet for the One-Off form
====================================================================================
Adds an "Events" worksheet to the Admin One-Off Google Sheet, seeded with a
single "Adhoc" entry. Unlike the main RG form's Events worksheet (which maps
each event to a specific owner), this list is global — every owner sees the
same Event Name options when SOD is "Events". Edit this worksheet directly
(add/remove/rename a row) to change what shows up in the form's dropdown —
no code change needed.

Idempotent — skipped if the "Events" worksheet already exists.

Prerequisites:
    You've already run `python setup_oneoff.py` (sheets_config.json with
    oneoff_admin_sheet_id, and token.pickle, must already exist).

Run once:
    python setup_oneoff_events.py
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
ADMIN_SHEET_KEY = "oneoff_admin_sheet_id"
EVENTS_WORKSHEET = "Events"
SEED_EVENTS = ["Adhoc"]
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


def get_credentials():
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
        print(f"ERROR: {CONFIG_FILE} not found. Run `python setup_oneoff.py` first.")
        sys.exit(1)
    with open(CONFIG_FILE) as f:
        config = json.load(f)
    if ADMIN_SHEET_KEY not in config:
        print(f"ERROR: '{ADMIN_SHEET_KEY}' missing from {CONFIG_FILE}.")
        sys.exit(1)

    print("Authenticating with your Google account...\n")
    creds = get_credentials()
    gc = gspread.authorize(creds)
    print("Signed in successfully!\n")

    admin_sh = gc.open_by_key(config[ADMIN_SHEET_KEY])
    existing_titles = [ws.title for ws in admin_sh.worksheets()]
    if EVENTS_WORKSHEET in existing_titles:
        print(f"'{EVENTS_WORKSHEET}' worksheet already exists — nothing to do.")
        sys.exit(0)

    print(f"Creating '{EVENTS_WORKSHEET}' worksheet...")
    events_ws = admin_sh.add_worksheet(EVENTS_WORKSHEET, rows=100, cols=1)
    events_ws.update([["EventName"]] + [[e] for e in SEED_EVENTS], value_input_option="RAW")
    print(f"  Events: {len(SEED_EVENTS)} row(s) uploaded: {SEED_EVENTS}")

    print("\n" + "=" * 60)
    print("SETUP COMPLETE!")
    print("=" * 60)
    print(f"\n  Admin One-Off sheet: https://docs.google.com/spreadsheets/d/{admin_sh.id}")


if __name__ == "__main__":
    main()
