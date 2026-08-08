"""
setup_events.py — One-time addition: Events worksheet for the main RG donation form
=====================================================================================
Adds an "Events" worksheet to the existing "TMO Admin" Google Sheet (positioned
after "BAs"), uploaded from EVENTS_XLSX. Maps Event Name -> Owner code so the
main form (app.py) can offer an Event dropdown filtered to the selected owner.

Idempotent — skipped if the "Events" worksheet already exists.

Prerequisites:
    1. You've already run `python setup.py` (sheets_config.json with
       admin_sheet_id, and token.pickle, must already exist).
    2. EVENTS_XLSX ("EventDataForForm.xlsx") is present in this folder, with a
       sheet containing columns "Owner code", "OwnerName", "Event Name".

Run once:
    python setup_events.py
"""

import json
import os
import pickle
import sys

try:
    import gspread
    import pandas as pd
    from google.auth.transport.requests import Request
    from google_auth_oauthlib.flow import InstalledAppFlow
except ImportError:
    print("Missing packages. Run:")
    print("  python -m pip install gspread pandas google-auth google-auth-oauthlib")
    sys.exit(1)

CLIENT_SECRET = "client_secret.json"
TOKEN_FILE = "token.pickle"
CONFIG_FILE = "sheets_config.json"
EVENTS_XLSX = "EventDataForForm.xlsx"
EVENTS_WORKSHEET = "Events"
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
    if not os.path.exists(EVENTS_XLSX):
        print(f"ERROR: {EVENTS_XLSX} not found in this folder.")
        sys.exit(1)

    print("Authenticating with your Google account...\n")
    creds = get_credentials()
    gc = gspread.authorize(creds)
    print("Signed in successfully!\n")

    admin_sh = gc.open_by_key(config["admin_sheet_id"])
    existing_titles = [ws.title for ws in admin_sh.worksheets()]
    if EVENTS_WORKSHEET in existing_titles:
        print(f"'{EVENTS_WORKSHEET}' worksheet already exists in the Admin sheet — nothing to do.")
        print("Delete that worksheet first if you want to re-upload fresh data.")
        sys.exit(0)

    print(f"Reading {EVENTS_XLSX}...")
    df = pd.read_excel(EVENTS_XLSX, dtype=str).fillna("")
    df = df.rename(columns={"Owner code": "OWNCODE", "Event Name": "EventName"})
    df["OWNCODE"] = df["OWNCODE"].str.strip()
    df = df[["OWNCODE", "OwnerName", "EventName"]]

    print(f"Creating '{EVENTS_WORKSHEET}' worksheet (after BAs)...")
    events_ws = admin_sh.add_worksheet(EVENTS_WORKSHEET, rows=max(len(df) + 1, 100), cols=3)
    events_ws.update([["OWNCODE", "OwnerName", "EventName"]], value_input_option="RAW")

    rows = df.values.tolist()
    batch = 500
    for i in range(0, len(rows), batch):
        events_ws.append_rows(rows[i:i + batch], value_input_option="RAW")
    print(f"  Events: {len(rows)} rows uploaded")

    print("\n" + "=" * 60)
    print("SETUP COMPLETE!")
    print("=" * 60)
    print(f"\n  Admin sheet: https://docs.google.com/spreadsheets/d/{admin_sh.id}")


if __name__ == "__main__":
    main()
