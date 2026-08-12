"""
setup_ba_joinee_date.py — One-time addition: New Joinee Date column for BAs
==============================================================================
Adds a "New Joinee Date" column (as the last column) to the "BAs" worksheet
in the Admin Google Sheet. Existing rows are left blank. From now on, when
someone adds a brand-new BA through the form's "Add a new BA" section, the
SignIn Date used for that submission is written into this column — purely
for the Admin sheet's own record-keeping. It is never shown in the form's
BA dropdown (which still only shows "BA Name · BA Code", unchanged).

Idempotent — skipped if "New Joinee Date" is already a column in "BAs".

Prerequisites:
    You've already run `python setup.py` (sheets_config.json with
    admin_sheet_id, and token.pickle, must already exist).

Run once:
    python setup_ba_joinee_date.py
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
BAS_WORKSHEET = "BAs"
NEW_COLUMN = "New Joinee Date"
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
    bas_ws = admin_sh.worksheet(BAS_WORKSHEET)
    header = bas_ws.row_values(1)

    if NEW_COLUMN in header:
        print(f"'{NEW_COLUMN}' column already exists in '{BAS_WORKSHEET}' — nothing to do.")
        sys.exit(0)

    new_col_index = len(header) + 1
    if new_col_index > bas_ws.col_count:
        bas_ws.add_cols(new_col_index - bas_ws.col_count)
    bas_ws.update_cell(1, new_col_index, NEW_COLUMN)
    print(f"Added '{NEW_COLUMN}' column to '{BAS_WORKSHEET}' (column {new_col_index}).")
    print("All existing BAs are left blank — only newly-added BAs get this filled in going forward.")

    print("\n" + "=" * 60)
    print("SETUP COMPLETE!")
    print("=" * 60)
    print(f"\n  Admin sheet: https://docs.google.com/spreadsheets/d/{admin_sh.id}")


if __name__ == "__main__":
    main()
