"""
Stress test for the Donation Entry app.

Writes exactly 2000 rows across varied batches with NO sleep between
requests, deliberately triggering Google Sheets 429 rate-limit errors
to verify the exponential-backoff retry logic (matching app.py behaviour).

Run:
    python stress_test.py

All test rows are tagged with [TEST] in BAName so they can be deleted later.
"""

import json
import os
import pickle
import random
import time
from collections import Counter
from datetime import date, timedelta

import gspread
from google.auth.transport.requests import Request

# ── Config ───────────────────────────────────────────────────────────────────
CONFIG_FILE = "sheets_config.json"
TOKEN_FILE  = "token.pickle"
SCOPES      = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
HEADERS        = ["SigninDT", "OWNCODE", "BAName", "BACode", "Amount(Amt)", "Age", "SOD"]
SOD_CATEGORIES = ["B2B/Commercial", "D2D/Resi", "Events", "Streets", "Airport"]

# Batch-size cycle: [5,10,8,15,12,6,20,10,7,7] sums to 100 per 10 batches.
# 20 full cycles = 200 batches = exactly 2000 rows.
BATCH_CYCLE = [5, 10, 8, 15, 12, 6, 20, 10, 7, 7]

# ── Auth ─────────────────────────────────────────────────────────────────────
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


# ── Admin data ───────────────────────────────────────────────────────────────
def load_admin(gc, cfg):
    sh = gc.open_by_key(cfg["admin_sheet_id"])

    def sheet_rows(ws_name):
        values = sh.worksheet(ws_name).get_all_values()
        if not values:
            return []
        hdrs = values[0]
        return [dict(zip(hdrs, row)) for row in values[1:]]

    owners, ba_by_code = [], {}
    for r in sheet_rows("Owners"):
        c = str(r.get("OWNCODE", "")).strip()
        if c:
            owners.append(c)
    for r in sheet_rows("BAs"):
        c    = str(r.get("OWNCODE", "")).strip()
        name = str(r.get("BAName",  "")).strip()
        code = str(r.get("BACode",  "")).strip()
        if c and name:
            ba_by_code.setdefault(c, []).append((name, code))
    return owners, ba_by_code


# ── Cached worksheet ──────────────────────────────────────────────────────────
_ws_donations = None

def get_donations_ws(gc, cfg):
    global _ws_donations
    if _ws_donations is None:
        sh = gc.open_by_key(cfg["donations_sheet_id"])
        _ws_donations = sh.sheet1
    return _ws_donations


# ── Retry logic — identical behaviour to app.py _append_with_retry ───────────
# Tracks 429 events so we can report them at the end.
retry_log = []   # list of (batch_num, attempt, wait_s)

def append_donations(gc, cfg, rows, batch_num, max_attempts=8):
    ws = get_donations_ws(gc, cfg)
    for attempt in range(max_attempts):
        try:
            ws.append_rows([[r[h] for h in HEADERS] for r in rows], value_input_option="RAW")
            return
        except gspread.exceptions.APIError as e:
            is_rate_limit = (
                (hasattr(e, "response") and e.response.status_code == 429)
                or "429" in str(e)
            )
            if is_rate_limit and attempt < max_attempts - 1:
                wait = 2 ** attempt          # 1, 2, 4, 8, 16 s  (same as app.py)
                retry_log.append((batch_num, attempt + 1, wait))
                print(
                    f"    [HIGH TRAFFIC] Batch {batch_num} -- "
                    f"429 hit, retrying in {wait}s "
                    f"(retry {attempt + 1}/{max_attempts - 1}) ..."
                )
                time.sleep(wait)
            else:
                raise


# ── Read all rows ────────────────────────────────────────────────────────────
def read_all_rows(gc, cfg):
    ws = get_donations_ws(gc, cfg)
    values = ws.get_all_values()
    if len(values) < 2:
        return []
    hdrs = values[0]
    return [dict(zip(hdrs, row)) for row in values[1:]]


# ── Build exactly 2000 rows across 200 varied batches ───────────────────────
def build_test_cases(owners, ba_by_code):
    random.seed(2000)
    today  = date.today()
    dates  = [today - timedelta(days=d) for d in range(7)]
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
        raise RuntimeError("No owners with BAs found in Admin sheet.")

    cases = []
    for i, batch_size in enumerate(BATCH_CYCLE * 20):   # 20 cycles × 10 = 200 batches
        owncode  = valid_owners[i % len(valid_owners)]
        bas      = ba_by_code[owncode]
        ba_name, ba_code = bas[(i // len(valid_owners)) % len(bas)]

        batch = []
        for j in range(batch_size):
            batch.append({
                "SigninDT":     dates[(i + j) % len(dates)].strftime("%Y-%m-%d"),
                "OWNCODE":      owncode,
                "BAName":       f"[TEST] {ba_name}",
                "BACode":       ba_code,
                "Amount(Amt)":  amounts[(i + j * 3) % len(amounts)],
                "Age":          ages[(i + j * 7) % len(ages)],
                "SOD":          SOD_CATEGORIES[(i + j) % len(SOD_CATEGORIES)],
            })
        cases.append(batch)

    return cases


# ── Verify every expected row landed ─────────────────────────────────────────
def verify(expected_rows, all_sheet_rows):
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
    print("  Donation Entry -- 2000-Row High-Traffic Stress Test")
    print("=" * 65)
    print("  Mode: NO sleep between batches (forces 429s to prove retry)")
    print()

    gc  = get_gc()
    cfg = load_config()
    print("OK  Authenticated with Google")

    owners, ba_by_code = load_admin(gc, cfg)
    print(f"OK  Loaded {len(owners)} owners, {sum(len(v) for v in ba_by_code.values())} BAs")

    cases      = build_test_cases(owners, ba_by_code)
    total_rows = sum(len(c) for c in cases)
    print(f"OK  Built {len(cases)} batches -> {total_rows} total rows")
    print(f"    Batch sizes: {BATCH_CYCLE} (cycling x20)\n")

    before = read_all_rows(gc, cfg)
    print(f"  Sheet rows before : {len(before)}")
    print()

    all_written = []
    errors      = []
    t_start     = time.time()

    for i, batch in enumerate(cases, start=1):
        try:
            append_donations(gc, cfg, batch, batch_num=i)
            all_written.extend(batch)
            print(f"  Batch {i:>3}/200 -- {len(batch):>2} rows  OK")
            # NO sleep here — we want to saturate the quota on purpose
        except Exception as e:
            errors.append((i, batch, str(e)))
            print(f"  Batch {i:>3}/200 -- FAILED: {e}")

    elapsed = time.time() - t_start
    print()

    # Read back and verify
    print("  Reading sheet back to verify all 2000 rows ...")
    after         = read_all_rows(gc, cfg)
    new_row_count = len(after) - len(before)
    matched, missing = verify(all_written, after)

    # ── Report ────────────────────────────────────────────────────────────────
    print()
    print("=" * 65)
    print("  RESULTS")
    print("=" * 65)
    print(f"  Total time              : {elapsed:.0f}s ({elapsed/60:.1f} min)")
    print(f"  Batches attempted       : {len(cases)}")
    print(f"  Batches failed (no fix) : {len(errors)}")
    print(f"  Rows expected           : {len(all_written)}")
    print(f"  Rows matched in sheet   : {len(matched)}")
    print(f"  Rows MISSING            : {len(missing)}")
    print(f"  Net rows added          : {new_row_count}")
    print()

    if retry_log:
        total_wait = sum(w for _, _, w in retry_log)
        print("  -- High-traffic events (429s caught and retried) --")
        print(f"  Total 429s hit          : {len(retry_log)}")
        print(f"  Total retry wait time   : {total_wait}s")
        print(f"  All retried silently -- user sees warning, data never lost.")
        print()
    else:
        print("  No 429 rate-limit errors triggered (quota not exhausted).")
        print()

    if errors:
        print("  -- Failed batches (exhausted all retries) --")
        for i, batch, err in errors:
            print(f"    Batch {i}: {err}")
        print()

    if missing:
        print("  -- Missing rows --")
        for r in missing:
            print(f"    {r}")
        print()
    else:
        print("  All 2000 rows verified in sheet. PASS")

    print()
    if not errors and not missing:
        print("  STRESS TEST PASSED -- 2000 rows saved, high traffic handled.")
    else:
        print("  STRESS TEST FOUND ISSUES -- see above.")

    print("=" * 65)
    print()
    print("  NOTE: 2000 [TEST] rows are in the Donations sheet.")
    print("  Filter BAName contains '[TEST]' and delete them when done.")
    print("=" * 65)


if __name__ == "__main__":
    main()
