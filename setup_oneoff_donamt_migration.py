"""
setup_oneoff_donamt_migration.py — One-time schema change: Forms/ADS -> DonAmt/SOD/etc.
==========================================================================================
Restructures the Donations One-Off worksheet's columns:
    Before: SigninDT, OWNCODE, BAName, BACode, Clients, Forms, Supports, ADS, No Forms
    After:  SigninDT, OWNCODE, BAName, BACode, Clients, DonAmt, Supports, SOD,
            Event Name, Mode of Payment, No Production

- "Forms" is renamed to "DonAmt" (existing values are left as-is; they were the old
  Forms count, now sitting under the DonAmt header — this is a one-time semantic
  rename of pre-existing historical rows, not a value backfill).
- "Supports" keeps its position and existing values; the form now computes it
  automatically as DonAmt / 1200 for new rows instead of taking it as input.
- "ADS" is removed entirely (no longer tracked).
- "SOD", "Event Name", "Mode of Payment" are added as new blank columns for
  existing rows.
- "No Forms" is renamed to "No Production" (existing 0/1 values kept as-is).

Idempotent — skipped if the header already matches the target shape.

Prerequisites:
    You've already run `python setup_oneoff.py` (sheets_config.json with
    oneoff_donations_sheet_id, and token.pickle, must already exist).

Run once:
    python setup_oneoff_donamt_migration.py
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

OLD_HEADER = ["SigninDT", "OWNCODE", "BAName", "BACode", "Clients", "Forms", "Supports", "ADS", "No Forms"]
NEW_HEADER = ["SigninDT", "OWNCODE", "BAName", "BACode", "Clients", "DonAmt", "Supports",
              "SOD", "Event Name", "Mode of Payment", "No Production"]

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
    if DONATIONS_SHEET_KEY not in config:
        print(f"ERROR: '{DONATIONS_SHEET_KEY}' missing from {CONFIG_FILE}.")
        sys.exit(1)

    print("Authenticating with your Google account...\n")
    creds = get_credentials()
    gc = gspread.authorize(creds)
    print("Signed in successfully!\n")

    sh = gc.open_by_key(config[DONATIONS_SHEET_KEY])
    ws = sh.sheet1

    header = ws.row_values(1)
    if header == NEW_HEADER:
        print("Header already matches the target shape — nothing to do.")
        sys.exit(0)
    if header != OLD_HEADER:
        print(f"ERROR: current header doesn't match the expected old shape.")
        print(f"  expected: {OLD_HEADER}")
        print(f"  actual:   {header}")
        print("Refusing to proceed automatically — check the sheet manually.")
        sys.exit(1)

    print(f"Current header confirmed: {header}")

    # 1) Rename "Forms" -> "DonAmt" (column F)
    ws.update_cell(1, 6, "DonAmt")
    print("Renamed column F header: Forms -> DonAmt")

    # 2) Delete "ADS" column (column H) — shifts "No Forms" from I to H
    ws.delete_columns(8, 8)
    print("Deleted column H (ADS)")

    after_delete = ws.row_values(1)
    print(f"Header after delete: {after_delete}")
    assert after_delete == ["SigninDT", "OWNCODE", "BAName", "BACode", "Clients", "DonAmt", "Supports", "No Forms"], \
        f"Unexpected header after delete: {after_delete}"

    # 3) Insert 3 new blank columns before "No Forms" (now column H), pushing it to K
    ws.insert_cols([[""], [""], [""]], col=8)
    print("Inserted 3 blank columns at H (for SOD, Event Name, Mode of Payment)")

    ws.update_cell(1, 8, "SOD")
    ws.update_cell(1, 9, "Event Name")
    ws.update_cell(1, 10, "Mode of Payment")
    print("Set headers: H=SOD, I=Event Name, J=Mode of Payment")

    # 4) Rename "No Forms" (now column K) -> "No Production"
    ws.update_cell(1, 11, "No Production")
    print("Renamed column K header: No Forms -> No Production")

    final_header = ws.row_values(1)
    print(f"\nFinal header: {final_header}")
    assert final_header == NEW_HEADER, f"Final header mismatch! Got: {final_header}"

    print("\n" + "=" * 60)
    print("MIGRATION COMPLETE!")
    print("=" * 60)
    print(f"\n  Donations One-Off sheet: https://docs.google.com/spreadsheets/d/{sh.id}")


if __name__ == "__main__":
    main()
