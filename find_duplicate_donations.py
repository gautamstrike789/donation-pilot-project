"""
find_duplicate_donations.py — Read-only duplicate-candidate report for both forms
=====================================================================================
Scans the RG and One-Off Donations sheets for groups of rows that share the
same SigninDT + OWNCODE + BAName + Amount + SOD, and writes them into a new
"Possible Duplicates" tab in each Donations spreadsheet — a plain report, not
a formula, and not a new column on the main sheet, so it can never interfere
with how the forms detect where to append new rows.

Rows landing in the same group are NOT necessarily duplicates — two separate,
real donations can legitimately share every one of those fields (same BA,
same amount, same day, same source). This report only narrows down which
rows are *worth a human look*; nothing is ever auto-deleted or auto-merged.

Safe to re-run any time — it only ever rewrites the "Possible Duplicates"
tab's own content, never touches the "Donations" / "Donations One-Off"
worksheet.

Run:
    python find_duplicate_donations.py
"""

import json
import os
import pickle
import sys
from collections import defaultdict

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
REPORT_WORKSHEET = "Possible Duplicates"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# (config key, amount column name)
TARGETS = [
    ("donations_sheet_id", "Amount(Amt)", "RG"),
    ("oneoff_donations_sheet_id", "DonAmt", "One-Off"),
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


def build_report(ws_donations, amount_col):
    values = ws_donations.get_all_values()
    if len(values) < 2:
        return [], 0
    header = values[0]
    idx = {name: header.index(name) for name in
           ["SigninDT", "OWNCODE", "OwnerName", "BAName", amount_col, "SOD"]}

    groups = defaultdict(list)
    for row_num, row in enumerate(values[1:], start=2):
        key = tuple(row[idx[name]].strip() if len(row) > idx[name] else "" for name in
                    ["SigninDT", "OWNCODE", "BAName", amount_col, "SOD"])
        groups[key].append((row_num, row))

    report_rows = []
    group_num = 0
    for key, rows in groups.items():
        if len(rows) < 2:
            continue
        group_num += 1
        for row_num, row in rows:
            report_rows.append([
                group_num,
                row_num,
                row[idx["SigninDT"]],
                row[idx["OWNCODE"]],
                row[idx["OwnerName"]] if len(row) > idx["OwnerName"] else "",
                row[idx["BAName"]],
                row[idx[amount_col]],
                row[idx["SOD"]],
            ])
    return report_rows, group_num


BAND_FILL = {"backgroundColor": {"red": 1.0, "green": 0.949, "blue": 0.800}}  # soft amber
NO_FILL = {"backgroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}}


def write_report(sh, report_rows, group_count, amount_col_label):
    existing_titles = [ws.title for ws in sh.worksheets()]
    if REPORT_WORKSHEET in existing_titles:
        ws = sh.worksheet(REPORT_WORKSHEET)
        ws.clear()
        # clear any earlier run's banding before laying down fresh formats
        ws.format("A1:H10000", NO_FILL)
    else:
        ws = sh.add_worksheet(REPORT_WORKSHEET, rows=max(len(report_rows) + 10, 100), cols=8)

    header = ["Group #", "Row # in Donations sheet", "SigninDT", "OWNCODE", "OwnerName",
              "BAName", amount_col_label, "SOD"]
    if not report_rows:
        ws.update([header, ["No matching groups found — nothing to review right now."]],
                  value_input_option="RAW")
        ws.format("A1:H1", {"textFormat": {"bold": True}})
        return

    ws.update([header] + report_rows, value_input_option="RAW")
    ws.format("A1:H1", {"textFormat": {"bold": True}})
    ws.freeze(rows=1)

    # Alternate a soft amber band per Group #, so consecutive duplicate-candidate
    # clusters are visually separated at a glance — every other group gets a
    # fill, the rest stay white. Sent as one batch_format call regardless of
    # how many groups there are.
    formats = []
    sheet_row = 2  # row 1 is the header
    for group_num in range(1, group_count + 1):
        group_size = sum(1 for r in report_rows if r[0] == group_num)
        if group_num % 2 == 0:
            formats.append({
                "range": f"A{sheet_row}:H{sheet_row + group_size - 1}",
                "format": BAND_FILL,
            })
        sheet_row += group_size
    if formats:
        ws.batch_format(formats)


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

    for key, amount_col, label in TARGETS:
        if key not in config:
            print(f"[{label}] '{key}' missing from {CONFIG_FILE} — skipping.")
            continue
        print(f"[{label}] scanning Donations sheet...")
        sh = gc.open_by_key(config[key])
        ws_donations = sh.sheet1
        report_rows, group_count = build_report(ws_donations, amount_col)
        write_report(sh, report_rows, group_count, amount_col)
        print(f"[{label}] {group_count} group(s), {len(report_rows)} row(s) flagged for review.")
        print(f"[{label}] written to the \"{REPORT_WORKSHEET}\" tab: "
              f"https://docs.google.com/spreadsheets/d/{sh.id}")
        print()

    print("Done. Remember: rows in the same group are CANDIDATES to review, not")
    print("confirmed duplicates — two separate real donations can share every one")
    print("of these fields. Nothing was changed in the Donations sheets themselves.")


if __name__ == "__main__":
    main()
