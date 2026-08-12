"""
setup_owner_passcodes.py — One-time addition: Passcode column for Owners
==========================================================================
Adds a "Passcode" column (as the last column) to the "Owners" worksheet in
the Admin Google Sheet. Existing rows are left blank on purpose — passcodes
are set from inside the form itself, self-service: the first time anyone
picks an owner whose Passcode cell is blank, the form asks them to set one
(typed twice to confirm) and writes it to this column automatically. No
admin action is needed, before or after adding a new owner to the sheet.

From then on, picking that owner shows a small passcode box that blocks the
rest of the form (SignIn Date, BA, Donations, and the Owners Data History
view) until the correct passcode is entered. Clear an owner's Passcode cell
yourself if you want the "set a passcode" prompt to run again for them.

Idempotent — skipped if "Passcode" is already a column in "Owners".

Prerequisites:
    You've already run `python setup.py` (sheets_config.json with
    admin_sheet_id, and token.pickle, must already exist).

Run once:
    python setup_owner_passcodes.py
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
OWNERS_WORKSHEET = "Owners"
NEW_COLUMN = "Passcode"
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
    owners_ws = admin_sh.worksheet(OWNERS_WORKSHEET)
    header = owners_ws.row_values(1)

    if NEW_COLUMN in header:
        print(f"'{NEW_COLUMN}' column already exists in '{OWNERS_WORKSHEET}' — nothing to do.")
        sys.exit(0)

    new_col_index = len(header) + 1
    if new_col_index > owners_ws.col_count:
        owners_ws.add_cols(new_col_index - owners_ws.col_count)
    owners_ws.update_cell(1, new_col_index, NEW_COLUMN)
    print(f"Added '{NEW_COLUMN}' column to '{OWNERS_WORKSHEET}' (column {new_col_index}).")
    print("All existing owners are left blank — the form will prompt each owner to set")
    print("their own passcode the first time they're picked. No manual entry needed.")

    print("\n" + "=" * 60)
    print("SETUP COMPLETE!")
    print("=" * 60)
    print(f"\n  Admin sheet: https://docs.google.com/spreadsheets/d/{admin_sh.id}")


if __name__ == "__main__":
    main()
