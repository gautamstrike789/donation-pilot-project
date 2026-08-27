"""
setup_oneoff_ownername_column.py — One-time addition: OwnerName column for the One-Off form
================================================================================================
Inserts an "OwnerName" column into the Donations One-Off worksheet, right
after OWNCODE (matching the main RG form's column order):
    Before: SigninDT, OWNCODE, BAName, BACode, Clients, DonAmt, Supports,
            SOD, Event Name, Mode of Payment, No Production
    After:  SigninDT, OWNCODE, OwnerName, BAName, BACode, Clients, DonAmt,
            Supports, SOD, Event Name, Mode of Payment, No Production

Existing rows are backfilled by looking up each row's OWNCODE against the
Admin One-Off sheet's Owners worksheet (OWNCODE|OwnerName|Client|Passcode) —
an OWNCODE can appear more than once there (once per Client), but they all
share the same OwnerName, so the first match found is used.

Idempotent — skipped if the header already matches the target shape.

Prerequisites:
    You've already run `python setup_oneoff.py` (sheets_config.json with
    oneoff_admin_sheet_id and oneoff_donations_sheet_id, and token.pickle,
    must already exist).

Run once:
    python setup_oneoff_ownername_column.py
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
DONATIONS_SHEET_KEY = "oneoff_donations_sheet_id"

OLD_HEADER = ["SigninDT", "OWNCODE", "BAName", "BACode", "Clients", "DonAmt", "Supports",
              "SOD", "Event Name", "Mode of Payment", "No Production"]
NEW_HEADER = ["SigninDT", "OWNCODE", "OwnerName", "BAName", "BACode", "Clients", "DonAmt",
              "Supports", "SOD", "Event Name", "Mode of Payment", "No Production"]

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
    for key in (ADMIN_SHEET_KEY, DONATIONS_SHEET_KEY):
        if key not in config:
            print(f"ERROR: '{key}' missing from {CONFIG_FILE}.")
            sys.exit(1)

    print("Authenticating with your Google account...\n")
    creds = get_credentials()
    gc = gspread.authorize(creds)
    print("Signed in successfully!\n")

    don_sh = gc.open_by_key(config[DONATIONS_SHEET_KEY])
    ws = don_sh.sheet1

    header = ws.row_values(1)
    if header == NEW_HEADER:
        print("Header already matches the target shape — nothing to do.")
        sys.exit(0)
    if header != OLD_HEADER:
        print("ERROR: current header doesn't match the expected old shape.")
        print(f"  expected: {OLD_HEADER}")
        print(f"  actual:   {header}")
        print("Refusing to proceed automatically — check the sheet manually.")
        sys.exit(1)

    print(f"Current header confirmed: {header}")

    values = ws.get_all_values()
    data_rows = values[1:]
    print(f"Backfilling OwnerName for {len(data_rows)} existing row(s)...")

    admin_sh = gc.open_by_key(config[ADMIN_SHEET_KEY])
    owners_values = admin_sh.worksheet("Owners").get_all_values()
    owners_header = owners_values[0]
    code_idx = owners_header.index("OWNCODE")
    name_idx = owners_header.index("OwnerName")
    owner_name_by_code = {}
    for row in owners_values[1:]:
        c = row[code_idx].strip() if len(row) > code_idx else ""
        if c and c not in owner_name_by_code:
            owner_name_by_code[c] = row[name_idx].strip() if len(row) > name_idx else ""

    # 1) Insert a blank column at position 3 (C), pushing BAName and
    #    everything after it one column to the right.
    ws.insert_cols([[""]], col=3)
    print("Inserted blank column at C")

    after_insert = ws.row_values(1)
    expected_after_insert = ["SigninDT", "OWNCODE", "", "BAName", "BACode", "Clients", "DonAmt",
                              "Supports", "SOD", "Event Name", "Mode of Payment", "No Production"]
    assert after_insert == expected_after_insert, f"Unexpected header after insert: {after_insert}"

    # 2) Set the new column's header
    ws.update_cell(1, 3, "OwnerName")
    print("Set header C1 = OwnerName")

    # 3) Backfill OwnerName for every existing data row from the Owners lookup
    if data_rows:
        owner_names = [[owner_name_by_code.get(row[1].strip() if len(row) > 1 else "", "")]
                        for row in data_rows]
        ws.update(f"C2:C{len(data_rows) + 1}", owner_names, value_input_option="RAW")
        matched = sum(1 for r in owner_names if r[0])
        print(f"Backfilled OwnerName for {matched} of {len(data_rows)} row(s) "
              f"({len(data_rows) - matched} had an OWNCODE not found in Owners).")

    final_header = ws.row_values(1)
    print(f"\nFinal header: {final_header}")
    assert final_header == NEW_HEADER, f"Final header mismatch! Got: {final_header}"

    print("\n" + "=" * 60)
    print("MIGRATION COMPLETE!")
    print("=" * 60)
    print(f"\n  Donations One-Off sheet: https://docs.google.com/spreadsheets/d/{don_sh.id}")


if __name__ == "__main__":
    main()
