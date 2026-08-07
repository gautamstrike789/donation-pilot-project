"""
setup_oneoff.py — One-time setup for the One-Off Donation Entry app (OAuth version)
=====================================================================================
Creates two Google Sheets dedicated to the One-Off form (fully independent from the
main donation form's "TMO Admin" / "TMO Donations" sheets):
    • "TMO Admin One-Off"      — uploaded from ADMIN_XLSX ("Owners" + "BAs" sheets).
      An OWNCODE can appear on multiple Owners rows with a different Client
      (e.g. HAI, STC) — each row is its own selectable entry in the form.
    • "TMO Donations One-Off"  — one worksheet with columns:
        SigninDT, OWNCODE, BAName, BACode, Clients, Forms, Supports, ADS

Extends sheets_config.json with oneoff_admin_sheet_id / oneoff_donations_sheet_id.
Each step is idempotent — it's skipped if its key is already present in the config.

Prerequisites:
    1. You've already run `python setup.py` at least once (token.pickle must exist).
    2. ADMIN_XLSX ("TMO Admin One-Off.xlsx") is present in this folder, with "Owners"
       and "BAs" sheets in the same column order as the main admin.xlsx.

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
ADMIN_XLSX = "TMO Admin One-Off.xlsx"
ADMIN_SHEET_KEY = "oneoff_admin_sheet_id"
DONATIONS_SHEET_KEY = "oneoff_donations_sheet_id"
DONATIONS_HEADERS = ["SigninDT", "OWNCODE", "BAName", "BACode", "Clients", "Forms", "Supports", "ADS"]
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


def save_config(config):
    with open(CONFIG_FILE, "w") as f:
        json.dump(config, f, indent=2)


def create_admin_sheet(gc, config):
    if ADMIN_SHEET_KEY in config:
        print(f"'{ADMIN_SHEET_KEY}' already present in {CONFIG_FILE}:")
        print(f"  https://docs.google.com/spreadsheets/d/{config[ADMIN_SHEET_KEY]}")
        print("Delete that key from the config file first if you want to create a fresh sheet.\n")
        return
    if not os.path.exists(ADMIN_XLSX):
        print(f"ERROR: {ADMIN_XLSX} not found in this folder — can't create the Admin One-Off sheet.")
        sys.exit(1)

    import pandas as pd
    print("Creating 'TMO Admin One-Off' Google Sheet...")
    admin_sh = gc.create("TMO Admin One-Off")

    owners_ws = admin_sh.sheet1
    owners_ws.update_title("Owners")
    owners_ws.update([["OWNCODE", "OwnerName", "Client"]], value_input_option="RAW")
    bas_ws = admin_sh.add_worksheet("BAs", rows=1000, cols=3)
    bas_ws.update([["OWNCODE", "BACode", "BAName"]], value_input_option="RAW")

    print(f"Uploading data from {ADMIN_XLSX}...")
    owners_df = pd.read_excel(ADMIN_XLSX, sheet_name="Owners", dtype=str).fillna("")
    bas_df = pd.read_excel(ADMIN_XLSX, sheet_name="BAs", dtype=str).fillna("")

    batch = 2000
    if len(owners_df):
        rows = owners_df.values.tolist()
        for i in range(0, len(rows), batch):
            owners_ws.append_rows(rows[i:i + batch], value_input_option="RAW")
        print(f"  Owners: {len(owners_df)} rows uploaded")

    if len(bas_df):
        rows = bas_df.values.tolist()
        for i in range(0, len(rows), batch):
            bas_ws.append_rows(rows[i:i + batch], value_input_option="RAW")
        print(f"  BAs: {len(bas_df)} rows uploaded")

    config[ADMIN_SHEET_KEY] = admin_sh.id
    save_config(config)
    print(f"  Sheet ID: {admin_sh.id}\n")


def create_donations_sheet(gc, config):
    if DONATIONS_SHEET_KEY in config:
        print(f"'{DONATIONS_SHEET_KEY}' already present in {CONFIG_FILE}:")
        print(f"  https://docs.google.com/spreadsheets/d/{config[DONATIONS_SHEET_KEY]}")
        print("Delete that key from the config file first if you want to create a fresh sheet.\n")
        return

    print("Creating 'TMO Donations One-Off' Google Sheet...")
    don_sh = gc.create("TMO Donations One-Off")
    don_ws = don_sh.sheet1
    don_ws.update_title("Donations One-Off")
    don_ws.update([DONATIONS_HEADERS], value_input_option="RAW")

    config[DONATIONS_SHEET_KEY] = don_sh.id
    save_config(config)
    print(f"  Sheet ID: {don_sh.id}\n")


def main():
    if not os.path.exists(CONFIG_FILE):
        print(f"ERROR: {CONFIG_FILE} not found. Run `python setup.py` first.")
        sys.exit(1)
    with open(CONFIG_FILE) as f:
        config = json.load(f)

    print("Authenticating with your Google account...\n")
    creds = get_credentials()
    gc = gspread.authorize(creds)
    print("Signed in successfully!\n")

    create_admin_sheet(gc, config)
    create_donations_sheet(gc, config)

    print("=" * 60)
    print("SETUP COMPLETE!")
    print("=" * 60)
    print(f"\n  To start the app:  python -m streamlit run app_oneoff.py")


if __name__ == "__main__":
    main()
