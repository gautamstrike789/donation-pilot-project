"""
setup_oneoff.py — One-time setup for the One-Off Donation Entry app (OAuth version)
=====================================================================================
Reuses the existing Admin sheet (Owners/BAs) created by setup.py, and creates a new
"TMO Donations One-Off" Google Sheet for app_oneoff.py to write to. Extends the
existing sheets_config.json with the new sheet's ID — admin_sheet_id and
donations_sheet_id (from setup.py) are left untouched.

Prerequisites:
    1. You've already run `python setup.py` at least once (sheets_config.json and
       token.pickle must already exist).

Run once:
    python setup_oneoff.py
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
DONATIONS_SHEET_KEY = "oneoff_donations_sheet_id"
HEADERS = ["SigninDT", "OWNCODE", "BAName", "BACode", "Forms", "Supports", "ADS"]
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
                print("Download the OAuth Client ID JSON from Google Cloud Console")
                print("and rename it to client_secret.json")
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
    if DONATIONS_SHEET_KEY in config:
        print(f"'{DONATIONS_SHEET_KEY}' already present in {CONFIG_FILE}:")
        print(f"  https://docs.google.com/spreadsheets/d/{config[DONATIONS_SHEET_KEY]}")
        print("Delete that key from the config file first if you want to create a fresh sheet.")
        sys.exit(0)

    print("Authenticating with your Google account...\n")
    creds = get_credentials()
    gc = gspread.authorize(creds)
    print("Signed in successfully!\n")

    print("Creating 'TMO Donations One-Off' Google Sheet...")
    don_sh = gc.create("TMO Donations One-Off")
    don_ws = don_sh.sheet1
    don_ws.update_title("Donations One-Off")
    don_ws.update([HEADERS], value_input_option="RAW")

    # Save the config immediately after creation succeeds, before any further
    # output — a print() encoding error here must not lose track of the new sheet.
    config[DONATIONS_SHEET_KEY] = don_sh.id
    with open(CONFIG_FILE, "w") as f:
        json.dump(config, f, indent=2)
    print(f"  Sheet ID: {don_sh.id}")
    print(f"\nConfig updated: {CONFIG_FILE}")

    print("\n" + "=" * 60)
    print("SETUP COMPLETE!")
    print("=" * 60)
    print(f"\n  Donations One-Off sheet: https://docs.google.com/spreadsheets/d/{don_sh.id}")
    print(f"\n  To start the app:  python -m streamlit run app_oneoff.py")


if __name__ == "__main__":
    main()
