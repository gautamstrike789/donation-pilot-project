"""
setup_submitlog.py — One-time addition: SubmitLog worksheet for both forms
==============================================================================
Adds a "SubmitLog" worksheet to the RG Admin sheet and the One-Off Admin
sheet, header: SubmitID | OWNCODE | RowsWritten | Timestamp.

This is the durable half of the duplicate-submission fix: each staged batch
of entries gets a random ID the moment it's first built, which stays the
same no matter how many times Submit is clicked on it (a genuine retry, a
double-click, a "please wait" warning followed by another click). Right
after a batch is confirmed written to the Donations sheet, its ID gets
logged here. Before writing, the app checks this log first — if the ID is
already present, the batch was already written by an earlier attempt, so
the write is skipped and the same success message is shown instead.

Two different batches that happen to contain identical-looking rows (same
BA, amount, date, source) are NOT affected — each gets its own random ID
the moment it's built, regardless of what its rows contain, so this can
never mistake two separate real donations for a duplicate.

Idempotent — skipped per-sheet if its "SubmitLog" worksheet already exists.

Run once:
    python setup_submitlog.py
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
SUBMITLOG_WORKSHEET = "SubmitLog"
HEADER = ["SubmitID", "OWNCODE", "RowsWritten", "Timestamp"]

TARGETS = [
    ("admin_sheet_id", "RG"),
    ("oneoff_admin_sheet_id", "One-Off"),
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
        existing_titles = [ws.title for ws in sh.worksheets()]
        if SUBMITLOG_WORKSHEET in existing_titles:
            print(f"[{label}] '{SUBMITLOG_WORKSHEET}' already exists — nothing to do.")
            continue
        print(f"[{label}] creating '{SUBMITLOG_WORKSHEET}' worksheet...")
        ws = sh.add_worksheet(SUBMITLOG_WORKSHEET, rows=1000, cols=4)
        ws.update([HEADER], value_input_option="RAW")
        print(f"[{label}] done: https://docs.google.com/spreadsheets/d/{sh.id}")

    print("\nSETUP COMPLETE!")


if __name__ == "__main__":
    main()
