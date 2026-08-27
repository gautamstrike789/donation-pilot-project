"""
setup_timestamp_column.py — One-time addition: Timestamp column for both forms
==================================================================================
Adds a "Timestamp" column as the new LAST column on the RG Donations sheet
and the One-Off Donations One-Off sheet, recording exactly when each row was
received from the form (separate from SigninDT, which is a date the BA
picks, not a submission time).

This only ever sets ONE header cell text in a column that's already blank
past the current data (both sheets are provisioned with far more grid
columns than are in use) — it does not insert or shift any existing
column, so it carries none of the risk a mid-sheet column insert would.
Existing historical rows are left blank in this column; there's no way to
know their real submission time, so nothing is backfilled/guessed.

Idempotent — skipped per-sheet if its header already ends with "Timestamp".

Run once:
    python setup_timestamp_column.py
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

TARGETS = [
    ("donations_sheet_id", "RG"),
    ("oneoff_donations_sheet_id", "One-Off"),
]

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
        print(f"ERROR: {CONFIG_FILE} not found.")
        sys.exit(1)
    with open(CONFIG_FILE) as f:
        config = json.load(f)

    print("Authenticating with your Google account...\n")
    creds = get_credentials()
    gc = gspread.authorize(creds)
    print("Signed in successfully!\n")

    for key, label in TARGETS:
        if key not in config:
            print(f"[{label}] '{key}' missing from {CONFIG_FILE} — skipping.")
            continue
        sh = gc.open_by_key(config[key])
        ws = sh.sheet1
        header = ws.row_values(1)
        if header and header[-1] == "Timestamp":
            print(f"[{label}] header already ends with Timestamp — nothing to do.")
            continue
        next_col = len(header) + 1
        print(f"[{label}] current header ({len(header)} cols): {header}")
        ws.update_cell(1, next_col, "Timestamp")
        after = ws.row_values(1)
        assert after == header + ["Timestamp"], f"unexpected header after update: {after}"
        print(f"[{label}] set column {next_col} header to 'Timestamp'. New header: {after}")

    print("\nSETUP COMPLETE!")


if __name__ == "__main__":
    main()
