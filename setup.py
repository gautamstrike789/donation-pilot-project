"""
setup.py — One-time setup for the Donation Entry app (OAuth version)
=====================================================================
Uses YOUR Google account (OAuth login) — no service account keys needed.

Creates two Google Sheets (Admin + Donations), uploads admin.xlsx data,
and writes sheets_config.json so the app knows where to read/write.

Prerequisites:
    1. A Google Cloud project with Sheets API + Drive API enabled.
    2. An OAuth 2.0 Client ID (Desktop type) downloaded as  client_secret.json.

Run once:
    python setup.py

A browser window will open for you to sign in with your Google account.
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
ADMIN_XLSX = "admin.xlsx"
CONFIG_FILE = "sheets_config.json"
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
    print("Authenticating with your Google account...\n")
    creds = get_credentials()
    gc = gspread.authorize(creds)
    print("Signed in successfully!\n")

    # ---- create Admin sheet ----
    print("Creating 'TMO Admin' Google Sheet...")
    admin_sh = gc.create("TMO Admin")
    print(f"  → Sheet ID: {admin_sh.id}")

    # rename default sheet to Owners, create BAs sheet
    owners_ws = admin_sh.sheet1
    owners_ws.update_title("Owners")
    owners_ws.update([["OWNCODE", "OwnerName", "City"]], value_input_option="RAW")
    bas_ws = admin_sh.add_worksheet("BAs", rows=5000, cols=3)
    bas_ws.update([["OWNCODE", "BACode", "BAName"]], value_input_option="RAW")

    # upload from admin.xlsx if it exists
    if os.path.exists(ADMIN_XLSX):
        import pandas as pd
        print(f"\nUploading data from {ADMIN_XLSX}...")
        owners_df = pd.read_excel(ADMIN_XLSX, sheet_name="Owners", dtype=str).fillna("")
        bas_df = pd.read_excel(ADMIN_XLSX, sheet_name="BAs", dtype=str).fillna("")

        if len(owners_df):
            owners_ws.append_rows(owners_df.values.tolist(), value_input_option="RAW")
            print(f"  Owners: {len(owners_df)} rows uploaded")

        if len(bas_df):
            rows = bas_df.values.tolist()
            batch = 500
            for i in range(0, len(rows), batch):
                bas_ws.append_rows(rows[i:i + batch], value_input_option="RAW")
            print(f"  BAs: {len(bas_df)} rows uploaded")
    else:
        print(f"\n{ADMIN_XLSX} not found — created empty Admin sheet (add owners/BAs manually).")

    # ---- create Donations sheet ----
    print("\nCreating 'TMO Donations' Google Sheet...")
    don_sh = gc.create("TMO Donations")
    don_ws = don_sh.sheet1
    don_ws.update_title("Donations")
    don_ws.update([["SigninDT", "OWNCODE", "BAName", "BACode", "Amount(Amt)", "Age", "SOD"]],
                  value_input_option="RAW")
    print(f"  → Sheet ID: {don_sh.id}")

    # ---- save config ----
    config = {
        "admin_sheet_id": admin_sh.id,
        "donations_sheet_id": don_sh.id,
    }
    with open(CONFIG_FILE, "w") as f:
        json.dump(config, f, indent=2)
    print(f"\nConfig saved to {CONFIG_FILE}")

    # ---- summary ----
    print("\n" + "=" * 60)
    print("SETUP COMPLETE!")
    print("=" * 60)
    print(f"\n  Admin sheet:     https://docs.google.com/spreadsheets/d/{admin_sh.id}")
    print(f"  Donations sheet: https://docs.google.com/spreadsheets/d/{don_sh.id}")
    print(f"\n  Both sheets are in YOUR Google Drive — open them anytime!")
    print(f"  You can view them live while submitting donations.")
    print(f"\n  To start the app:  python -m streamlit run app.py")


if __name__ == "__main__":
    main()
