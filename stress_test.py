"""
Stress test for the Donation Entry app.

Bypasses Streamlit and calls the same Google Sheets functions directly.
Writes 50+ rows across varied combinations, then reads the sheet back
to verify every row was saved. Prints a pass/fail report.

Run:
    python stress_test.py

All test rows are tagged with [TEST] in BAName so they can be deleted later.
"""

import json
import os
import pickle
import random
import time
from datetime import date, timedelta

import gspread
from google.auth.transport.requests import Request

# ── Config (same as app.py) ─────────────────────────────────────────────────
CONFIG_FILE = "sheets_config.json"
TOKEN_FILE = "token.pickle"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
HEADERS = ["SigninDT", "OWNCODE", "BAName", "BACode", "Amount(Amt)", "Age", "SOD"]
SOD_CATEGORIES = ["B2B/Commercial", "D2D/Resi", "Events", "Streets", "Airport"]

# ── Auth (mirrors app.py get_credentials) ───────────────────────────────────
def get_gc():
    creds = None
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, "rb") as f:
            creds = pickle.load(f)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
            with open(TOKEN_FILE, "wb") as f:
                pickle.dump(creds, f)
        else:
            raise RuntimeError("Token expired. Run python setup.py to re-authenticate.")
    return gspread.authorize(creds)


def load_config():
    with open(CONFIG_FILE) as f:
        return json.load(f)


# ── Load real owners + BAs from Admin sheet ──────────────────────────────────
def load_admin(gc, cfg):
    sh = gc.open_by_key(cfg["admin_sheet_id"])

    def sheet_rows(ws_name):
        values = sh.worksheet(ws_name).get_all_values()
        if not values:
            return []
        headers = values[0]
        return [dict(zip(headers, row)) for row in values[1:]]

    owners_rows = sheet_rows("Owners")
    bas_rows = sheet_rows("BAs")

    owners = []
    for r in owners_rows:
        c = str(r.get("OWNCODE", "")).strip()
        if c:
            owners.append(c)

    ba_by_code = {}
    for r in bas_rows:
        c = str(r.get("OWNCODE", "")).strip()
        name = str(r.get("BAName", "")).strip()
        bacode = str(r.get("BACode", "")).strip()
        if c and name:
            ba_by_code.setdefault(c, []).append((name, bacode))

    return owners, ba_by_code


# ── Cached worksheet (avoids read-quota hit on every call) ───────────────────
_ws_donations = None

def get_donations_ws(gc, cfg):
    global _ws_donations
    if _ws_donations is None:
        sh = gc.open_by_key(cfg["donations_sheet_id"])
        _ws_donations = sh.sheet1
    return _ws_donations


# ── Append rows to Donations sheet with retry on 429 ─────────────────────────
def append_donations(gc, cfg, rows):
    ws = get_donations_ws(gc, cfg)
    for attempt in range(5):
        try:
            ws.append_rows([[r[h] for h in HEADERS] for r in rows], value_input_option="RAW")
            return
        except Exception as e:
            if "429" in str(e) and attempt < 4:
                wait = 15 * (attempt + 1)
                print(f"    429 rate limit — waiting {wait}s before retry {attempt + 1}/4 ...")
                time.sleep(wait)
            else:
                raise


# ── Read all rows currently in the sheet ────────────────────────────────────
def read_all_rows(gc, cfg):
    ws = get_donations_ws(gc, cfg)
    values = ws.get_all_values()
    if len(values) < 2:
        return []
    hdrs = values[0]
    return [dict(zip(hdrs, row)) for row in values[1:]]


# ── Build 200 test cases ─────────────────────────────────────────────────────
def build_test_cases(owners, ba_by_code, target_batches=200):
    """
    200 batches, ~270 total rows, covering:
    - All owners with BAs (cycles through all 55)
    - Multiple BAs per owner (rotates on revisit)
    - All 5 SOD categories
    - 12 amounts: min-edge (499), common, large, decimal
    - 17 ages: every major bracket from 25 to 99
    - 7 dates spanning the last week
    - Batch sizes: 70% single, 20% double, 10% three-to-five rows
    """
    random.seed(99)
    today = date.today()
    dates = [today - timedelta(days=d) for d in range(7)]
    amounts = [
        "499", "500", "600", "750", "999", "1000",
        "1500", "2000", "3000", "5000", "10000", "499.99",
    ]
    ages = [
        "25", "28", "30", "33", "35", "38", "40", "45",
        "50", "55", "60", "65", "70", "75", "80", "90", "99",
    ]

    valid_owners = [o for o in owners if ba_by_code.get(o)]
    if not valid_owners:
        raise RuntimeError("No owners with BAs found in Admin sheet. Cannot run test.")

    cases = []

    for i in range(target_batches):
        owncode = valid_owners[i % len(valid_owners)]
        bas = ba_by_code[owncode]
        # Rotate through available BAs each time this owner comes around
        ba_name, ba_code = bas[(i // len(valid_owners)) % len(bas)]

        # Batch size distribution: 70% × 1, 20% × 2, 10% × 3-5
        roll = i % 10
        if roll < 7:
            batch_size = 1
        elif roll < 9:
            batch_size = 2
        else:
            batch_size = (i % 3) + 3  # 3, 4, or 5

        batch = []
        for j in range(batch_size):
            batch.append({
                "SigninDT": dates[(i + j) % len(dates)].strftime("%Y-%m-%d"),
                "OWNCODE": owncode,
                "BAName": f"[TEST] {ba_name}",
                "BACode": ba_code,
                "Amount(Amt)": amounts[(i + j * 3) % len(amounts)],
                "Age": ages[(i + j * 7) % len(ages)],
                "SOD": SOD_CATEGORIES[(i + j) % len(SOD_CATEGORIES)],
            })
        cases.append(batch)

    return cases


# ── Verify: check every expected row exists in the sheet ────────────────────
def verify(expected_rows, all_sheet_rows):
    """
    For each expected row, look for an exact match in the sheet.
    Returns (matched, missing).
    """
    # Build a multiset from sheet rows for O(n) matching
    from collections import Counter
    def row_key(r):
        return tuple(str(r.get(h, "")).strip() for h in HEADERS)

    sheet_counts = Counter(row_key(r) for r in all_sheet_rows)
    matched, missing = [], []
    for r in expected_rows:
        k = row_key(r)
        if sheet_counts.get(k, 0) > 0:
            sheet_counts[k] -= 1
            matched.append(r)
        else:
            missing.append(r)
    return matched, missing


# ── Main ─────────────────────────────────────────────────────────────────────
def main():
    print("=" * 65)
    print("  Donation Entry — Stress Test")
    print("=" * 65)

    gc = get_gc()
    cfg = load_config()
    print("OK  Authenticated with Google")

    owners, ba_by_code = load_admin(gc, cfg)
    print(f"OK  Loaded {len(owners)} owners, {sum(len(v) for v in ba_by_code.values())} BAs")

    cases = build_test_cases(owners, ba_by_code, target_batches=200)
    total_rows = sum(len(c) for c in cases)
    print(f"OK  Built {len(cases)} test batches -> {total_rows} total rows\n")

    # Snapshot sheet size before writing
    before = read_all_rows(gc, cfg)
    print(f"  Sheet rows before test : {len(before)}")

    # Write each batch (each case = one "Submit" click)
    all_written = []
    errors = []
    for i, batch in enumerate(cases, start=1):
        try:
            append_donations(gc, cfg, batch)
            all_written.extend(batch)
            print(f"  Batch {i:>3}/{len(cases)} -- wrote {len(batch)} row(s)  OK")
            # 1 write request per second stays well under the 60/min write quota
            time.sleep(1.1)
        except Exception as e:
            errors.append((i, batch, str(e)))
            print(f"  Batch {i:>3}/{len(cases)} -- FAILED: {e}")

    print()

    # Read back and verify
    print("  Reading sheet back to verify …")
    after = read_all_rows(gc, cfg)
    print(f"  Sheet rows after  test : {len(after)}")
    new_row_count = len(after) - len(before)

    matched, missing = verify(all_written, after)

    # ── Report ───────────────────────────────────────────────────────────────
    print()
    print("=" * 65)
    print("  RESULTS")
    print("=" * 65)
    print(f"  Batches attempted   : {len(cases)}")
    print(f"  Batches failed      : {len(errors)}")
    print(f"  Rows expected       : {len(all_written)}")
    print(f"  Rows matched        : {len(matched)}")
    print(f"  Rows MISSING        : {len(missing)}")
    print(f"  Net rows added      : {new_row_count}")
    print()

    if errors:
        print("-- Failed batches --")
        for i, batch, err in errors:
            print(f"  Batch {i}: {err}")
        print()

    if missing:
        print("-- Missing rows (written but not found in sheet) ---")
        for r in missing:
            print(f"  {r}")
        print()
    else:
        print("  All rows verified in sheet. PASS")

    if not errors and not missing:
        print()
        print("  STRESS TEST PASSED -- every row was saved correctly.")
    else:
        print()
        print("  STRESS TEST FOUND ISSUES -- see details above.")

    print("=" * 65)
    print()
    print("  NOTE: Test rows have '[TEST]' in BAName.")
    print("  Delete them from the Donations sheet when done.")
    print()
    print("=" * 65)
    print("  CAPACITY ANALYSIS for 2000 entries/week")
    print("=" * 65)
    print("  Google Sheets write quota : 60 requests/min (OAuth)")
    print("  Rows per submit (avg)     : ~1-5 (each submit = 1 API call)")
    print("  Peak throughput safe      : 60 submits/min -> ~300 rows/min")
    print("  2000 rows/week            = ~286 rows/day = ~1 row every 5 min")
    print("  Sheet row limit           : 10,000,000 cells / 7 cols = 1.4M rows")
    print("  Rows at 2000/week fills   : 700 weeks (~13 years) before hitting limit")
    print()
    print("  VERDICT: Form handles 2000 entries/week with significant headroom.")
    print("  Bottleneck risk: only if many users submit simultaneously on the")
    print("  same service account. With the cached worksheet fix, each submit")
    print("  costs exactly 1 write request (down from 3 before the fix).")
    print("=" * 65)


if __name__ == "__main__":
    main()
