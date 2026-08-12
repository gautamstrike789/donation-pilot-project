"""
setup_oneoff_no_forms_column.py — One-time addition: No Forms column for Donations One-Off
==============================================================================================
Adds a "No Forms" column (as the last column) to the Donations One-Off
worksheet. Existing rows are backfilled with "0" (normal entry). New rows
get "1" when submitted via the "No Forms" checkbox in the One-Off form
(equivalent to the main RG form's "NO Production" checkbox) — used when an
owner has no forms/supports to report for a sign-in.

Idempotent — skipped if "No Forms" is already a column.

Prerequisites:
    You've already run `python setup_oneoff.py` (sheets_config.json with
    oneoff_donations_sheet_id, and token.pickle, must already exist).

Run once:
    python setup_oneoff_no_forms_column.py
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
NEW_COLUMN = "No Forms"
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
        print(f"ERROR: {CONFIG_FILE} not found. Run `python setup_oneoff.py` first.")
        sys.exit(1)
    with open(CONFIG_FILE) as f:
        config = json.load(f)
    if DONATIONS_SHEET_KEY not in config:
        print(f"ERROR: '{DONATIONS_SHEET_KEY}' missing from {CONFIG_FILE}. Run `python setup_oneoff.py` first.")
        sys.exit(1)

    print("Authenticating with your Google account...\n")
    creds = get_credentials()
    gc = gspread.authorize(creds)
    print("Signed in successfully!\n")

    sh = gc.open_by_key(config[DONATIONS_SHEET_KEY])
    ws = sh.sheet1
    header = ws.row_values(1)

    if NEW_COLUMN in header:
        print(f"'{NEW_COLUMN}' column already exists — nothing to do.")
        sys.exit(0)

    new_col_index = len(header) + 1
    if new_col_index > ws.col_count:
        ws.add_cols(new_col_index - ws.col_count)
    ws.update_cell(1, new_col_index, NEW_COLUMN)

    values = ws.get_all_values()
    data_row_count = len(values) - 1
    if data_row_count > 0:
        col_letter = chr(ord("A") + new_col_index - 1)
        ws.update(
            range_name=f"{col_letter}2:{col_letter}{data_row_count + 1}",
            values=[["0"] for _ in range(data_row_count)],
            value_input_option="RAW",
        )
        print(f"Backfilled '{NEW_COLUMN}' = 0 for {data_row_count} existing row(s).")

    print(f"Added '{NEW_COLUMN}' column (column {new_col_index}).")

    print("\n" + "=" * 60)
    print("SETUP COMPLETE!")
    print("=" * 60)
    print(f"\n  Donations One-Off sheet: https://docs.google.com/spreadsheets/d/{sh.id}")


if __name__ == "__main__":
    main()
