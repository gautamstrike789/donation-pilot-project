"""
Donation Entry — One-Off Owner submission portal (Streamlit + Google Sheets)
=============================================================================
Form order:  Owner Code  →  Passcode  →  SignIn Date  →  BA Name  →  Add-new-BA  →  Entries (+ rows)

Data source: its own dedicated sheets, independent from the main donation form,
configured via sheets_config.json (created by setup_oneoff.py):
    • Admin One-Off sheet     — worksheets "Owners" (OWNCODE|OwnerName|Client|Passcode,
      Passcode added by setup_oneoff_owner_passcodes.py), "BAs"
      (OWNCODE|BACode|BAName|Client|New Joinee Date, New Joinee Date added by
      setup_oneoff_ba_joinee_date.py), and "Events" (EventName, added by
      setup_oneoff_events.py — a flat, global list, not owner-scoped; edit that
      worksheet directly to add/remove/rename events). An OWNCODE can have
      multiple Owners rows with a different Client (e.g. HAI, STC) — each is
      its own selectable entry; picking one sets both together. A Passcode
      protects the OWNCODE as a whole, regardless of which Client row was
      picked. The BA Name dropdown is scoped to BAs whose Client matches the
      currently selected owner's Client.
    • Donations One-Off sheet — one worksheet with columns:
        SigninDT, OWNCODE, OwnerName, BAName, BACode, Clients, DonAmt, Supports,
        SOD, Event Name, Mode of Payment, No Production, Timestamp
      (OwnerName added by setup_oneoff_ownername_column.py; Timestamp added by
      setup_timestamp_column.py and records when each row was actually
      received from the form, separate from SigninDT which is a date the BA
      picks — blank for historical rows submitted before it existed; the rest
      added/renamed by setup_oneoff_donamt_migration.py)

Supports is calculated automatically as DonAmt / 1200 — it is not a form
input, and is shown live next to the DonAmt field as it's typed.

The first time an owner with a blank "Passcode" cell (Owners sheet) is
picked, the form asks them to set one (typed twice to confirm) and saves it
to every row sharing that OWNCODE. From then on, picking that owner shows a
passcode box that gates everything below it (SignIn Date onward, plus the
Owners Data History view) until the correct passcode is entered — for the
rest of that browser session, it isn't asked again.

The "No Production" checkbox (right after SignIn Date) is for an owner with
nothing to report for that sign-in — it hides BA/Entries and records a
single row with only SigninDT, OWNCODE, and Clients filled in, no preview
step. When a brand-new BA is added through the form, the SignIn Date used
for that submission is written into the BAs sheet's "New Joinee Date"
column — never shown in the form itself.

Uses OAuth (your Google account) — no service account keys needed.

Run:
    python -m pip install -r requirements.txt
    python setup.py                        # one-time: signs in with your Google account
    python setup_oneoff.py                 # one-time: creates the Admin + Donations One-Off sheets + extends config
    python setup_oneoff_owner_passcodes.py # one-time: adds the Passcode column to Owners
    python setup_oneoff_ba_joinee_date.py  # one-time: adds the New Joinee Date column to BAs
    python setup_oneoff_donamt_migration.py# one-time: DonAmt/SOD/Event Name/Mode of Payment/No Production columns
    python setup_oneoff_events.py          # one-time: adds the global Events worksheet
    python setup_oneoff_ownername_column.py# one-time: adds the OwnerName column to Donations One-Off
    python -m streamlit run app_oneoff.py
"""

import json
import os
import pickle
import time
import uuid
from datetime import datetime, timedelta, timezone
from io import BytesIO

import gspread
import pandas as pd
import streamlit as st
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from google.oauth2.service_account import Credentials as ServiceAccountCredentials
from openpyxl import Workbook

LOGO_FILE = "logo.png"
CONFIG_FILE = "sheets_config.json"
CLIENT_SECRET = "client_secret.json"
TOKEN_FILE = "token.pickle"
SECRETS_SECTION = "google_sheets"
SERVICE_ACCOUNT_SECTION = "gcp_service_account"
ADMIN_SHEET_KEY = "oneoff_admin_sheet_id"
DONATIONS_SHEET_KEY = "oneoff_donations_sheet_id"
HEADERS = ["SigninDT", "OWNCODE", "OwnerName", "BAName", "BACode", "Clients", "DonAmt", "Supports",
           "SOD", "Event Name", "Mode of Payment", "No Production", "Timestamp"]
SOD_CATEGORIES = ["B2B/Commercial", "D2D/Resi", "Events", "Streets", "Roadtrip", "Telesales"]
MODE_OF_PAYMENT = ["Online", "Cheque"]
SUPPORTS_DIVISOR = 1200
SUPPORTS_DECIMALS = 2
IST = timezone(timedelta(hours=5, minutes=30))
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

st.set_page_config(
    page_title="Donation Entry — One-Off Owner Portal",
    page_icon=LOGO_FILE if os.path.exists(LOGO_FILE) else "🎗️",
    layout="centered",
)


# --------------------------------------------------------------------------- #
#  Google OAuth + Sheets connection
# --------------------------------------------------------------------------- #
def _has_secret(section):
    """Safely check for a secrets section. Returns False when no secrets.toml
    exists (e.g. local runs), instead of raising StreamlitSecretNotFoundError."""
    try:
        return section in st.secrets
    except Exception:  # noqa: BLE001 — no secrets file / not configured
        return False


def get_credentials():
    """Load cached OAuth token, refresh if expired, or prompt re-auth."""
    if _has_secret(SERVICE_ACCOUNT_SECTION):
        return ServiceAccountCredentials.from_service_account_info(
            dict(st.secrets[SERVICE_ACCOUNT_SECTION]),
            scopes=SCOPES,
        )

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
            st.error(
                f"Google login token expired or missing. "
                f"Run **`python setup.py`** again in your terminal to re-authenticate."
            )
            st.stop()
    return creds


def get_gc():
    """Return an authorized gspread client (cached per session)."""
    if "gc" not in st.session_state:
        creds = get_credentials()
        st.session_state.gc = gspread.authorize(creds)
    return st.session_state.gc


def get_ws(sheet_id, cache_key, ws_selector):
    """Return a cached worksheet — avoids a read-quota hit on every submit."""
    if cache_key not in st.session_state:
        gc = get_gc()
        sh = gc.open_by_key(sheet_id)
        st.session_state[cache_key] = sh.sheet1 if ws_selector == 0 else sh.worksheet(ws_selector)
    return st.session_state[cache_key]


def load_config():
    if _has_secret(SECRETS_SECTION):
        secret_cfg = st.secrets[SECRETS_SECTION]
        return {
            ADMIN_SHEET_KEY: secret_cfg[ADMIN_SHEET_KEY],
            DONATIONS_SHEET_KEY: secret_cfg[DONATIONS_SHEET_KEY],
        }
    with open(CONFIG_FILE) as f:
        cfg = json.load(f)
    if ADMIN_SHEET_KEY not in cfg or DONATIONS_SHEET_KEY not in cfg:
        st.error(
            f"**{ADMIN_SHEET_KEY}** / **{DONATIONS_SHEET_KEY}** missing from {CONFIG_FILE}. "
            "Run **`python setup_oneoff.py`** first to create the Admin One-Off and "
            "Donations One-Off sheets and extend the config file."
        )
        st.stop()
    return cfg


def cloud_secrets_ready():
    return _has_secret(SERVICE_ACCOUNT_SECTION) and _has_secret(SECRETS_SECTION)


# --------------------------------------------------------------------------- #
#  Read dropdown data from the Admin Google Sheet (cached 2 min)
# --------------------------------------------------------------------------- #
@st.cache_data(ttl=120, show_spinner="Loading dropdown data…")
def load_admin():
    cfg = load_config()
    gc = get_gc()
    sh = gc.open_by_key(cfg[ADMIN_SHEET_KEY])

    def sheet_rows(ws_name):
        values = sh.worksheet(ws_name).get_all_values()
        if not values:
            return []
        headers = values[0]
        return [dict(zip(headers, row)) for row in values[1:]]

    owners_rows = sheet_rows("Owners")
    bas_rows = sheet_rows("BAs")

    # Every (OWNCODE, Client) row is its own dropdown entry — the same owner
    # code can appear multiple times with a different Client (e.g. HAI, STC),
    # and selecting one picks both the owner and the client together. The
    # Passcode protects the OWNCODE as a whole, so we take the first
    # non-blank Passcode found across all of that owner's rows.
    owner_labels, label_to_owner, owner_codes = [], {}, set()
    owner_name_by_code, passcode_by_code = {}, {}
    for r in owners_rows:
        c = str(r.get("OWNCODE", "")).strip()
        if not c:
            continue
        name = str(r.get("OwnerName", "")).strip()
        client = str(r.get("Client", "")).strip()
        meta = " — ".join(p for p in [name, client] if p)
        label = f"{c}  ·  {meta}" if meta else c
        if label not in label_to_owner:
            owner_labels.append(label)
            label_to_owner[label] = (c, client)
        owner_codes.add(c)
        owner_name_by_code.setdefault(c, name)
        pc = str(r.get("Passcode", "")).strip()
        if pc and not passcode_by_code.get(c):
            passcode_by_code[c] = pc
        else:
            passcode_by_code.setdefault(c, "")

    # BA Name options are scoped by (OWNCODE, Client) — an owner code with
    # several Client rows (e.g. HAI, STC) only sees the BAs registered under
    # the Client actually selected.
    ba_by_owner_client, ba_codes = {}, set()
    for r in bas_rows:
        c = str(r.get("OWNCODE", "")).strip()
        name = str(r.get("BAName", "")).strip()
        bacode = str(r.get("BACode", "")).strip()
        ba_client = str(r.get("Client", "")).strip()
        if not c or not name:
            continue
        key = (c, ba_client)
        ba_by_owner_client.setdefault(key, [])
        if name not in [t[0] for t in ba_by_owner_client[key]]:
            ba_by_owner_client[key].append((name, bacode))
        if bacode:
            ba_codes.add(bacode)
    for key in ba_by_owner_client:
        ba_by_owner_client[key] = sorted(ba_by_owner_client[key], key=lambda t: t[0].lower())

    events_values = sh.worksheet("Events").get_all_values()
    event_options = sorted(
        {row[0].strip() for row in events_values[1:] if row and row[0].strip()},
        key=str.lower,
    ) if events_values else []

    return {"owner_labels": owner_labels, "label_to_owner": label_to_owner,
            "owner_count": len(owner_codes), "owner_name_by_code": owner_name_by_code,
            "passcode_by_code": passcode_by_code,
            "ba_by_owner_client": ba_by_owner_client, "ba_codes": ba_codes,
            "event_options": event_options}


@st.cache_data(ttl=60, show_spinner=False)
def load_donations_df():
    """Whole Donations One-Off sheet, cached briefly — backs the Owners Data History view."""
    cfg = load_config()
    ws = get_ws(cfg[DONATIONS_SHEET_KEY], "_ws_donations_oneoff", 0)
    values = ws.get_all_values()
    if len(values) < 2:
        return pd.DataFrame(columns=HEADERS)
    return pd.DataFrame(values[1:], columns=values[0])


def _get_status_code(e):
    """Extract HTTP status code from a gspread APIError."""
    if hasattr(e, "response") and hasattr(e.response, "status_code"):
        return e.response.status_code
    for code in (429, 401, 403, 500, 503):
        if str(code) in str(e):
            return code
    return None


def _append_with_retry(ws, rows, cache_keys, max_attempts=8, max_verify_rounds=5):
    """Append rows with exponential backoff for 429 (rate limit), automatic
    session reset for 401/403 (expired token), and — critically — a
    verify-by-content pass after every write.

    Under a real concurrent-submission stress test on the main RG form,
    ws.append_rows() was observed to occasionally report success while a
    batch of rows never actually landed in the sheet — no exception, no
    error, just silently fewer rows than expected. So after writing, we
    re-read the sheet and check, by exact content, which of the rows we
    sent are actually present; anything missing gets resent — and only the
    missing ones, so a resend can never create a duplicate of a row that
    did land."""
    notice = st.empty()
    try:
        remaining = list(rows)
        for verify_round in range(max_verify_rounds):
            if not remaining:
                return
            for attempt in range(max_attempts):
                try:
                    ws.append_rows(remaining, value_input_option="RAW")
                    break
                except gspread.exceptions.APIError as e:
                    code = _get_status_code(e)
                    if code == 429 and attempt < max_attempts - 1:
                        wait = 2 ** attempt  # 1, 2, 4, 8, 16, 32, 64 s
                        notice.warning(
                            f"High traffic detected — your data is safe and will be saved "
                            f"in {wait} second{'s' if wait > 1 else ''}. "
                            f"Please do not close this tab. "
                            f"(Retry {attempt + 1} of {max_attempts - 1}, "
                            f"max wait {2 ** (max_attempts - 1) - 1}s)"
                        )
                        time.sleep(wait)
                    elif code in (401, 403) and attempt == 0:
                        # Token expired mid-session: drop cached gc + worksheet and retry once
                        notice.warning("Re-authenticating with Google — please wait a moment...")
                        for k in ("gc", *cache_keys):
                            st.session_state.pop(k, None)
                        creds = get_credentials()
                        st.session_state.gc = gspread.authorize(creds)
                        cfg = load_config()
                        # Rebuild whichever worksheet we're writing to
                        for ck in cache_keys:
                            if ck == "_ws_donations_oneoff":
                                sh = st.session_state.gc.open_by_key(cfg[DONATIONS_SHEET_KEY])
                                st.session_state[ck] = sh.sheet1
                                ws = st.session_state[ck]
                            elif ck == "_ws_bas":
                                sh = st.session_state.gc.open_by_key(cfg[ADMIN_SHEET_KEY])
                                st.session_state[ck] = sh.worksheet("BAs")
                                ws = st.session_state[ck]
                            elif ck == "_ws_owners":
                                sh = st.session_state.gc.open_by_key(cfg[ADMIN_SHEET_KEY])
                                st.session_state[ck] = sh.worksheet("Owners")
                                ws = st.session_state[ck]
                            elif ck == "_ws_submitlog":
                                sh = st.session_state.gc.open_by_key(cfg[ADMIN_SHEET_KEY])
                                st.session_state[ck] = sh.worksheet("SubmitLog")
                                ws = st.session_state[ck]
                    else:
                        raise

            time.sleep(0.5)
            after_set = {tuple(r) for r in ws.get_all_values()}
            remaining = [r for r in remaining if tuple(r) not in after_set]
            if remaining and verify_round < max_verify_rounds - 1:
                time.sleep(0.5 * (verify_round + 1))

        if remaining:
            raise RuntimeError(
                f"{len(remaining)} of {len(rows)} row(s) could not be confirmed as "
                "written to the sheet after several attempts."
            )
    finally:
        notice.empty()


def _update_cell_with_retry(ws, row, col, value, cache_keys, max_attempts=8):
    """Same retry/backoff/re-auth handling as _append_with_retry, for a
    single-cell write to an existing row (no table-detection risk here —
    the target cell is addressed directly)."""
    notice = st.empty()
    try:
        for attempt in range(max_attempts):
            try:
                ws.update_cell(row, col, value)
                return
            except gspread.exceptions.APIError as e:
                code = _get_status_code(e)
                if code == 429 and attempt < max_attempts - 1:
                    wait = 2 ** attempt
                    notice.warning(
                        f"High traffic detected — retrying in {wait} second{'s' if wait > 1 else ''}. "
                        f"Please do not close this tab."
                    )
                    time.sleep(wait)
                elif code in (401, 403) and attempt == 0:
                    notice.warning("Re-authenticating with Google — please wait a moment...")
                    for k in ("gc", *cache_keys):
                        st.session_state.pop(k, None)
                    creds = get_credentials()
                    st.session_state.gc = gspread.authorize(creds)
                    cfg = load_config()
                    for ck in cache_keys:
                        if ck == "_ws_owners":
                            sh = st.session_state.gc.open_by_key(cfg[ADMIN_SHEET_KEY])
                            st.session_state[ck] = sh.worksheet("Owners")
                            ws = st.session_state[ck]
                else:
                    raise
    finally:
        notice.empty()


def append_bas(new_rows):
    """new_rows: list of (OWNCODE, BACode, BAName, Client, NewJoineeDate) -> appended to the BAs worksheet."""
    cfg = load_config()
    ws = get_ws(cfg[ADMIN_SHEET_KEY], "_ws_bas", "BAs")
    _append_with_retry(ws, [list(r) for r in new_rows], cache_keys=("_ws_bas",))


def set_owner_passcode(code, passcode):
    """Writes `passcode` into the Owners sheet's Passcode column for every
    row whose OWNCODE is `code` (an owner can have several rows, one per
    Client) — the passcode protects the owner as a whole."""
    cfg = load_config()
    ws = get_ws(cfg[ADMIN_SHEET_KEY], "_ws_owners", "Owners")
    values = ws.get_all_values()
    header = values[0]
    code_idx = header.index("OWNCODE")
    pc_idx = header.index("Passcode")
    matched = False
    for i, row in enumerate(values[1:], start=2):
        if len(row) > code_idx and row[code_idx].strip() == code:
            _update_cell_with_retry(ws, i, pc_idx + 1, passcode, cache_keys=("_ws_owners",))
            matched = True
    if not matched:
        raise ValueError(f"Owner code {code!r} not found in the Owners sheet.")


def already_submitted(submit_id):
    """True if this exact batch (by its random submit id, not its content)
    was already confirmed written to the Donations One-Off sheet by an
    earlier attempt — catches double-clicks and retries without ever
    treating two separate, genuinely identical-looking donations as
    duplicates, since each batch gets its own id regardless of content."""
    cfg = load_config()
    ws = get_ws(cfg[ADMIN_SHEET_KEY], "_ws_submitlog", "SubmitLog")
    return submit_id in ws.col_values(1)


def log_submission(submit_id, code, row_count):
    """Records that `submit_id` was successfully written, so a later
    retry/double-click of the same batch can be recognized and skipped."""
    cfg = load_config()
    ws = get_ws(cfg[ADMIN_SHEET_KEY], "_ws_submitlog", "SubmitLog")
    _append_with_retry(
        ws, [[submit_id, code, str(row_count), datetime.now(timezone.utc).isoformat()]],
        cache_keys=("_ws_submitlog",),
    )


def append_donations(rows):
    """rows: list of dicts keyed by HEADERS -> appended to the Donations One-Off sheet.
    Returns (rows_before, rows_after) read straight back from the sheet so the
    caller can prove the write actually landed (data row count excludes header)."""
    cfg = load_config()
    ws = get_ws(cfg[DONATIONS_SHEET_KEY], "_ws_donations_oneoff", 0)
    before = max(len(ws.get_all_values()) - 1, 0)
    _append_with_retry(ws, [[r[h] for h in HEADERS] for r in rows], cache_keys=("_ws_donations_oneoff",))
    after = max(len(ws.get_all_values()) - 1, 0)
    return before, after


def session_xlsx_bytes(entries):
    wb = Workbook()
    ws = wb.active
    ws.title = "Submitted"
    ws.append(HEADERS)
    for e in entries:
        ws.append([e[h] for h in HEADERS])
    bio = BytesIO()
    wb.save(bio)
    return bio.getvalue()


def df_to_xlsx_bytes(df, sheet_title="Sheet1"):
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_title
    ws.append(list(df.columns))
    for row in df.itertuples(index=False):
        ws.append(list(row))
    bio = BytesIO()
    wb.save(bio)
    return bio.getvalue()


# --------------------------------------------------------------------------- #
#  Guard + load
# --------------------------------------------------------------------------- #
if not cloud_secrets_ready() and not os.path.exists(CONFIG_FILE):
    st.error(
        f"**{CONFIG_FILE}** not found. Run **`python setup.py`** first to create "
        "Google Sheets and generate the config file, or deploy with Streamlit secrets."
    )
    st.stop()

if not cloud_secrets_ready() and not os.path.exists(TOKEN_FILE):
    st.error(
        f"**{TOKEN_FILE}** not found. Run **`python setup.py`** first to sign in "
        "with your Google account, or deploy with Streamlit secrets."
    )
    st.stop()

try:
    A = load_admin()
except Exception as e:  # noqa: BLE001
    st.error(f"Couldn't load admin data from Google Sheets: {e}")
    st.stop()

# --------------------------------------------------------------------------- #
#  Visual styling
# --------------------------------------------------------------------------- #
st.markdown(
    """
    <style>
      @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@500;600;700&display=swap');
      .stApp { background: linear-gradient(180deg,#f7f9fd 0%, #eaeefb 100%); }
      [data-testid="stHeader"] { background: transparent; }
      .block-container { padding-top: 2.2rem; padding-bottom: 3rem; max-width: 780px; }
      h1, h2, h3 { font-family: 'Poppins','Segoe UI',sans-serif !important; color:#1f2a5a; letter-spacing:-.01em; }
      h1 { font-weight:700 !important; }
      [data-testid="stVerticalBlockBorderWrapper"]{
        background:#ffffff; border:1px solid #e4e9f6; border-radius:16px;
        box-shadow:0 2px 12px rgba(31,42,90,.06); padding:.55rem .55rem;
      }
      [data-testid="stWidgetLabel"] p { font-weight:600; color:#41496b; }
      [data-baseweb="input"] input, [data-baseweb="select"] > div, .stDateInput input{
        border-radius:10px !important;
      }
      .stButton > button{
        border-radius:10px; font-weight:600; padding:.5rem 1.05rem; border:1px solid #d4dbf0;
        transition:transform .12s ease, box-shadow .12s ease;
      }
      .stButton > button:hover{ transform:translateY(-1px); box-shadow:0 4px 14px rgba(31,42,90,.12); }
      .stButton > button[kind="primary"]{
        background:linear-gradient(135deg,#2a3a7a,#1f2a5a); border:none; color:#fff;
      }
      .stDownloadButton > button{
        border-radius:10px; font-weight:600; border:1px solid #d4dbf0;
      }
      hr { border-color:#e0e6f4; }
      [data-testid="stDataFrame"]{ border-radius:12px; overflow:hidden; border:1px solid #e4e9f6; }
      [data-testid="stAlert"]{ border-radius:12px; }
    </style>
    """,
    unsafe_allow_html=True,
)

# --------------------------------------------------------------------------- #
#  Session state
# --------------------------------------------------------------------------- #
st.session_state.setdefault("nonce", 0)
st.session_state.setdefault("ba_nonce", 0)
st.session_state.setdefault("rows", [0])
st.session_state.setdefault("next_id", 1)
st.session_state.setdefault("session_entries", [])
st.session_state.setdefault("new_bas", {})
st.session_state.setdefault("pending_preview", [])    # validated rows awaiting confirm (accumulates across BAs)
st.session_state.setdefault("pending_new_bas", [])    # new BAs staged for the Admin sheet, written on Submit
st.session_state.setdefault("flash_success", None)    # confirmation message that survives the post-submit rerun
st.session_state.setdefault("preview_nonce", 0)       # rotated whenever pending_preview resets, so the
                                                       # preview editor widget doesn't carry over stale edits
st.session_state.setdefault("submitting", False)      # guards against a duplicate Submit click resubmitting
                                                       # the same staged rows before the first write finishes
st.session_state.setdefault("submitting_since", None)  # when the current submit lock was taken, so a
                                                        # stalled attempt (e.g. a hung network call that
                                                        # never raises or returns) can't leave every later
                                                        # click permanently stuck on "already in progress"
st.session_state.setdefault("submit_batch_id", None)   # random id for the batch currently staged for
                                                        # submission — stable across repeat Submit clicks
                                                        # on the same batch, so a durable check against
                                                        # SubmitLog can tell "already written" from "new"
n = st.session_state.nonce
bn = st.session_state.ba_nonce

# _append_with_retry's own bounded retry math tops out at roughly 5 verify
# rounds * ~127s of 429 backoff each, plus a few seconds of sleeps between
# rounds — call it ~11 minutes worst case. A submit lock held longer than
# that can only mean the attempt that took it never reached its own
# finally block (killed thread, hard network hang, etc.), not that it's
# still legitimately working — so treat it as abandoned and let a new
# click through instead of blocking the user forever.
SUBMIT_LOCK_TIMEOUT = timedelta(minutes=15)


def submit_lock_is_stale():
    started = st.session_state.get("submitting_since")
    return bool(started) and (datetime.now(timezone.utc) - started) > SUBMIT_LOCK_TIMEOUT


def parse_positive_amount(raw):
    text = str(raw or "").strip()
    if not text:
        return None
    try:
        value = float(text)
    except ValueError:
        return None
    return value


def fmt_number(v):
    """Render a numeric cell (possibly numpy float64 from the data editor) back to
    a plain string, without a spurious trailing .0 for whole numbers."""
    try:
        f = float(v)
    except (TypeError, ValueError):
        return str(v) if v is not None else ""
    return str(int(f)) if f.is_integer() else str(f)


def compute_supports(donamt_v):
    return round(donamt_v / SUPPORTS_DIVISOR, SUPPORTS_DECIMALS)


def validate_edited_rows(records):
    """Re-validate rows coming back from the editable preview grid (values may
    have been typed/changed there). Returns (errors, clean_rows) — clean_rows
    is only meaningful when errors is empty. Supports and OwnerName are always
    recomputed from DonAmt/OWNCODE rather than trusted from the grid, since
    they're derived (OwnerName also guards against the grid's data_editor
    occasionally round-tripping a blank cell as float NaN, which would
    otherwise land in the sheet as the literal text "nan")."""
    errors, clean_rows = [], []
    for i, r in enumerate(records, start=1):
        signindt = str(r.get("SigninDT") or "").strip()
        ownc = str(r.get("OWNCODE") or "").strip()
        ban = str(r.get("BAName") or "").strip()
        bac = str(r.get("BACode") or "").strip()
        clients = str(r.get("Clients") or "").strip()
        donamt_raw = r.get("DonAmt")
        sod_v = str(r.get("SOD") or "").strip()
        event_v = str(r.get("Event Name") or "").strip()
        mop_v = str(r.get("Mode of Payment") or "").strip()

        # a blank row added via the grid's "+" control — skip silently
        if not any([signindt, ownc, ban, donamt_raw not in (None, ""), sod_v]):
            continue

        donamt_v = parse_positive_amount(donamt_raw)

        if not signindt:
            errors.append(f"Row {i}: **SigninDT** is required.")
        if not ownc:
            errors.append(f"Row {i}: **OWNCODE** is required.")
        if not ban:
            errors.append(f"Row {i}: **BAName** is required.")
        if donamt_v is None or donamt_v <= 0:
            errors.append(f"Row {i}: **DonAmt** must be a number greater than 0.")
        if not sod_v or sod_v not in SOD_CATEGORIES:
            errors.append(f"Row {i}: **SOD** must be one of {', '.join(SOD_CATEGORIES)}.")
        elif sod_v == "Events" and not event_v:
            errors.append(f"Row {i}: **Event Name** is required when SOD is Events.")
        if not mop_v or mop_v not in MODE_OF_PAYMENT:
            errors.append(f"Row {i}: **Mode of Payment** must be one of {', '.join(MODE_OF_PAYMENT)}.")

        supports_v = compute_supports(donamt_v) if donamt_v else ""
        ownername = A["owner_name_by_code"].get(ownc, "")
        clean_rows.append({
            "SigninDT": signindt, "OWNCODE": ownc, "OwnerName": ownername, "BAName": ban, "BACode": bac,
            "Clients": clients, "DonAmt": fmt_number(donamt_raw), "Supports": fmt_number(supports_v),
            "SOD": sod_v, "Event Name": event_v if sod_v == "Events" else "",
            "Mode of Payment": mop_v, "No Production": "0",
            # set fresh here (not trusted from the grid) so it reflects the moment this
            # row was actually validated for submission, not whenever it was first staged
            "Timestamp": datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S"),
        })
    return errors, clean_rows


def sync_preview_from_grid(records):
    """Rebuild pending_preview from the editable grid's current output so that
    edits/deletions made there persist across reruns. Without this, deleting a
    row only ever changed what the widget displayed — st.session_state.pending_preview
    (the actual source of truth used to rebuild the grid on every rerun, e.g.
    when a new entry gets appended for another BA) never learned about the
    deletion, so the "deleted" row would silently come back. Lenient: doesn't
    validate, just mirrors whatever is currently showing (blank filler rows
    from the grid's own "+" control are dropped)."""
    synced = []
    for r in records:
        row = {h: r.get(h, "") for h in HEADERS}
        if not any([str(row.get("SigninDT") or "").strip(), str(row.get("OWNCODE") or "").strip(),
                    str(row.get("BAName") or "").strip(), row.get("DonAmt") not in (None, ""),
                    str(row.get("SOD") or "").strip()]):
            continue
        row["DonAmt"] = fmt_number(row["DonAmt"]) if row["DonAmt"] not in (None, "") else ""
        row["Supports"] = fmt_number(row["Supports"]) if row["Supports"] not in (None, "") else ""
        # OwnerName is derived from OWNCODE, not user-entered — re-derive it here too so a
        # data_editor round-trip that turns a blank cell into float NaN (which would render
        # as the literal text "nan") never gets persisted back into pending_preview.
        row["OwnerName"] = A["owner_name_by_code"].get(str(row.get("OWNCODE") or "").strip(), "")
        synced.append(row)
    return synced


# --------------------------------------------------------------------------- #
#  Header
# --------------------------------------------------------------------------- #
total_bas = sum(len(v) for v in A["ba_by_owner_client"].values())
hc1, hc2 = st.columns([1, 4], vertical_alignment="center")
if os.path.exists(LOGO_FILE):
    hc1.image(LOGO_FILE, width=130)
hc2.title("Donation Entry — One-Off")
hc2.caption(f"{A['owner_count']} owners · {total_bas:,} BAs · source: Google Sheets")

# ---- 1) Owner Code ----
owner_label = st.selectbox(
    "1 · Owner Code (OWNCODE) *",
    A["owner_labels"],
    index=None,
    placeholder="Search your owner code…",
    key="owner",
)
code, client = A["label_to_owner"].get(owner_label, (None, None)) if owner_label else (None, None)
owner_name = A["owner_name_by_code"].get(code, "") if code else ""

st.session_state.setdefault("verified_owners", set())

passcode_ok = False
if code:
    stored_passcode = A["passcode_by_code"].get(code, "")
    if code in st.session_state.verified_owners:
        passcode_ok = True
        st.caption("🔓 Passcode verified for this owner.")
    elif not stored_passcode:
        st.info(
            "🔐 No passcode set for this owner yet — set one now. "
            "You'll be asked for it every time you pick this owner from now on."
        )
        spc1, spc2 = st.columns(2)
        new_passcode = spc1.text_input("Set a passcode", type="password", key=f"newpc_{code}")
        confirm_passcode = spc2.text_input("Confirm passcode", type="password", key=f"newpc_confirm_{code}")
        if st.button("💾 Save passcode", key=f"savepc_{code}"):
            if not new_passcode or len(new_passcode) < 4:
                st.error("Passcode must be at least 4 characters.")
            elif new_passcode != confirm_passcode:
                st.error("Passcodes don't match.")
            else:
                try:
                    set_owner_passcode(code, new_passcode)
                    load_admin.clear()
                    st.session_state.verified_owners.add(code)
                    st.session_state.flash_success = f"🔐 Passcode set for owner {code}."
                    st.rerun()
                except Exception as e:  # noqa: BLE001
                    st.error(f"Couldn't save passcode: {e}")
    else:
        entered_passcode = st.text_input(
            "🔒 Enter this owner's passcode to continue",
            type="password",
            key=f"passcode_{code}",
        )
        if entered_passcode:
            if entered_passcode == stored_passcode:
                st.session_state.verified_owners.add(code)
                passcode_ok = True
            else:
                st.error("Incorrect passcode.")

history_slot = st.empty()

if code and not passcode_ok:
    st.markdown(
        '<style>div[class*="st-key-entry_form"] { display: none; }</style>',
        unsafe_allow_html=True,
    )

with st.container(key="entry_form"):
    # ---- 2) SignIn Date ----
    signin = st.date_input("2 · SignIn Date *", value=None, format="YYYY-MM-DD", key="signin")

    if not signin:
        st.info("Select the **SignIn Date** above to continue.")
    else:
        # ---- NO Production ----
        no_production = st.checkbox(
            "NO Production — this owner has no donations to report for this sign-in",
            key=f"noprod_{n}",
        )

        if not no_production:
            # ---- 3) BA Name (shows "Name · BACode"; key includes code so it resets when owner changes) ----
            ba_key = (code, client) if code else None
            combined = {}
            for nm_, cd_ in (A["ba_by_owner_client"].get(ba_key, []) if ba_key else []):
                combined[nm_] = cd_
            for nm_, cd_ in (st.session_state.new_bas.get(ba_key, []) if ba_key else []):
                combined.setdefault(nm_, cd_)
            if ba_key:
                for owner_c, owner_cl, cd_, nm_, _jd_ in st.session_state.pending_new_bas:
                    if (owner_c, owner_cl) == ba_key:
                        combined.setdefault(nm_, cd_)
            ba_pairs = sorted(combined.items(), key=lambda t: t[0].lower())
            ba_labels = [f"{nm_}  ·  {cd_}" if cd_ else nm_ for nm_, cd_ in ba_pairs]
            label_to_name = {(f"{nm_}  ·  {cd_}" if cd_ else nm_): nm_ for nm_, cd_ in ba_pairs}
            name_to_code = {nm_: cd_ for nm_, cd_ in ba_pairs}

            new_name_typed = bool(str(st.session_state.get(f"newname_{n}", "") or "").strip())
            if not code:
                ba_ph = "Select owner code first…"
            elif new_name_typed:
                ba_ph = "Disabled — you're adding a new BA below"
            else:
                ba_ph = "Search BA name or code…"

            ba_sel_label = st.selectbox(
                "3 · BA Name *",
                ba_labels,
                index=None,
                placeholder=ba_ph,
                disabled=(not code) or new_name_typed,
                key=f"ba_{bn}_{code or 'x'}",
            )
            ba_sel = None if new_name_typed else (label_to_name.get(ba_sel_label) if ba_sel_label else None)

            # ---- 4) Add a new BA (optional) ----
            with st.container(border=True):
                st.markdown("**➕ Add a new BA**  — fill these only if the BA isn't in the list above")
                nb1, nb2 = st.columns(2)
                new_ba_name = nb1.text_input("New BA Name", key=f"newname_{n}", disabled=not code,
                                             placeholder="Full name")
                code_mode = nb2.selectbox("New BA Code", ["Unassigned", "Enter code manually"],
                                          index=None, placeholder="Select…", disabled=not code,
                                          key=f"codemode_{n}")
                manual_code = ""
                if code_mode == "Enter code manually":
                    manual_code = st.text_input("Enter BA Code", key=f"manualcode_{n}", disabled=not code,
                                                placeholder="e.g. MMUN011-09999")

            # ---- 5) Entries (dynamic rows) ----
            st.markdown("#### Entries")
            st.caption("Add one or more entries for this BA, then save them all at once.")

            row_inputs = []
            to_remove = None
            # keyed by the same nonce that rotates whenever a fresh round of entries
            # starts, so these two checkboxes reset to unchecked for the next BA/round
            # without ever needing to write to an already-instantiated widget's key
            same_sod_all = st.session_state.get(f"same_sod_all_{n}", False)
            same_mop_all = st.session_state.get(f"same_mop_all_{n}", False)
            first_rid = st.session_state.rows[0]
            entry1_sod = st.session_state.get(f"sod_{first_rid}")
            entry1_mop = st.session_state.get(f"mop_{first_rid}")
            for idx, rid in enumerate(st.session_state.rows, start=1):
                with st.container(border=True):
                    h = st.columns([6, 1])
                    h[0].markdown(f"**Entry #{idx}**")
                    if len(st.session_state.rows) > 1 and h[1].button("✕", key=f"rm_{rid}", help="Remove"):
                        to_remove = rid
                    d1, d2 = st.columns(2)
                    donamt = d1.text_input("DonAmt *", key=f"donamt_{rid}", placeholder="amount > 0")
                    _d = (donamt or "").strip()
                    donamt_v = parse_positive_amount(_d) if _d else None
                    if _d and (donamt_v is None or donamt_v <= 0):
                        d1.caption(":red[⚠ Enter a number greater than 0]")
                    elif donamt_v:
                        d1.caption(f"Supports = **{compute_supports(donamt_v)}**")

                    # entries after #1 get their SOD/Mode of Payment pre-filled (and locked)
                    # from Entry #1 whenever the matching checkbox below is checked — this
                    # must happen before the widget is created so it takes effect this render
                    sod_locked = idx > 1 and same_sod_all
                    mop_locked = idx > 1 and same_mop_all
                    if sod_locked and entry1_sod:
                        st.session_state[f"sod_{rid}"] = entry1_sod
                    if mop_locked and entry1_mop:
                        st.session_state[f"mop_{rid}"] = entry1_mop

                    sod = d2.selectbox("Source of Donation (SOD) *", SOD_CATEGORIES, index=None,
                                       placeholder="Select a source…", key=f"sod_{rid}",
                                       disabled=sod_locked)
                    if sod_locked:
                        d2.caption("Matches Entry #1")

                    event_name = ""
                    if sod == "Events":
                        if not A["event_options"]:
                            st.selectbox("Event Name *", [], index=None, disabled=True,
                                         placeholder="No events set up yet", key=f"event_ph_{rid}")
                        else:
                            event_name = st.selectbox("Event Name *", A["event_options"], index=None,
                                                      placeholder="Select an event…", key=f"event_{rid}")

                    mop = st.selectbox("Mode of Payment *", MODE_OF_PAYMENT, index=None,
                                       placeholder="Select a mode…", key=f"mop_{rid}",
                                       disabled=mop_locked)
                    if mop_locked:
                        st.caption("Matches Entry #1")

                    if idx == 1:
                        sc1, sc2 = st.columns(2)
                        same_sod_all = sc1.checkbox("Use same SOD for all entries", key=f"same_sod_all_{n}")
                        same_mop_all = sc2.checkbox("Use same Mode of Payment for all entries",
                                                    key=f"same_mop_all_{n}")

                    row_inputs.append((idx, donamt, sod, event_name, mop))

            if to_remove is not None:
                st.session_state.rows.remove(to_remove)
                st.rerun()

            ca, cs = st.columns(2)
            if ca.button("➕ Add another entry", use_container_width=True):
                st.session_state.rows.append(st.session_state.next_id)
                st.session_state.next_id += 1
                st.rerun()
            save_clicked = cs.button("💾 Save all entries", type="primary", use_container_width=True)

            # --------------------------------------------------------------------------- #
            #  Step 1: Validate + show preview (does NOT write to Google Sheets yet)
            # --------------------------------------------------------------------------- #
            if save_clicked:
                errors = []
                if not code:
                    errors.append("Select a valid **owner code**.")

                # resolve BA: new vs existing
                nm = (new_ba_name or "").strip()
                if code_mode == "Unassigned":
                    cd = "Unassigned"
                elif code_mode == "Enter code manually":
                    cd = (manual_code or "").strip()
                else:
                    cd = ""
                existing_lower = {t[0].lower() for t in (A["ba_by_owner_client"].get(ba_key, []) if ba_key else [])}
                existing_lower |= {t[0].lower() for t in (st.session_state.new_bas.get(ba_key, []) if ba_key else [])}
                existing_lower |= {
                    nm_.lower() for owner_c, owner_cl, cd_, nm_, _jd_ in st.session_state.pending_new_bas
                    if (owner_c, owner_cl) == ba_key
                }
                is_new, effective_ba = False, ""
                if nm:
                    effective_ba = nm
                    if nm.lower() not in existing_lower:
                        is_new = True
                        if not code_mode:
                            errors.append("Choose a **New BA Code** option — *Unassigned* or *Enter code manually*.")
                        elif code_mode == "Enter code manually" and not cd:
                            errors.append("Enter the **BA Code**, or choose *Unassigned*.")
                elif ba_sel:
                    effective_ba = ba_sel
                else:
                    errors.append("Select a **BA Name**, or add a new one.")

                row_code = cd if is_new else name_to_code.get(effective_ba, "")

                # validate every entry row
                valid_rows = []
                for idx, donamt, sod, event_name, mop in row_inputs:
                    donamt = (donamt or "").strip()
                    donamt_v = parse_positive_amount(donamt)

                    if not any([donamt, sod, mop]):
                        continue
                    if donamt_v is None or donamt_v <= 0:
                        errors.append(f"Entry #{idx}: **DonAmt** must be a number greater than 0.")
                    if not sod:
                        errors.append(f"Entry #{idx}: select a **source of donation**.")
                    elif sod == "Events" and not event_name:
                        errors.append(f"Entry #{idx}: select an **event name**.")
                    if not mop:
                        errors.append(f"Entry #{idx}: select a **mode of payment**.")
                    if (donamt_v is not None and donamt_v > 0 and sod
                            and (sod != "Events" or event_name) and mop):
                        supports = str(compute_supports(donamt_v))
                        valid_rows.append({"SigninDT": signin.strftime("%Y-%m-%d"), "OWNCODE": code,
                                           "OwnerName": owner_name, "BAName": effective_ba, "BACode": row_code,
                                           "Clients": client or "", "DonAmt": donamt, "Supports": supports,
                                           "SOD": sod, "Event Name": event_name if sod == "Events" else "",
                                           "Mode of Payment": mop, "No Production": "0",
                                           # placeholder — validate_edited_rows() sets the real
                                           # value fresh at submit time, this is just staging
                                           "Timestamp": ""})

                if not errors and not valid_rows:
                    errors.append("Add at least one entry (DonAmt, SOD, mode of payment) before saving.")

                if errors:
                    for e in errors:
                        st.error(e)
                    # keep whatever is already staged in the preview; just don't add this invalid batch
                else:
                    # accumulate validated rows into the running preview (NOT saved to Sheets yet)
                    st.session_state.pending_preview.extend(valid_rows)
                    # stage a new BA (if any) for the Admin sheet on Submit — dedup by owner + client + name
                    if is_new and effective_ba:
                        already = {
                            (o, ocl, nm.lower()) for o, ocl, _c, nm, _jd in st.session_state.pending_new_bas
                        }
                        if (code, client, effective_ba.lower()) not in already:
                            st.session_state.pending_new_bas.append(
                                (code, client, cd, effective_ba, signin.strftime("%Y-%m-%d"))
                            )
                    # clear for the next batch: keep owner code, sign-in date, and BA name;
                    # reset only the "add a new BA" fields and the entry rows
                    st.session_state.nonce += 1
                    st.session_state.rows = [st.session_state.next_id]
                    st.session_state.next_id += 1
                    # pending_preview's shape just changed (rows appended) — rotate the
                    # preview editor's widget key so Streamlit doesn't try to reapply
                    # stale per-row edit/delete state (from the old shape) onto the new
                    # data, which is what let a deleted row silently come back to life.
                    st.session_state.preview_nonce += 1
                    st.rerun()

            # --------------------------------------------------------------------------- #
            #  Step 2: Preview + Submit button (writes to Google Sheets only on confirm)
            # --------------------------------------------------------------------------- #
            if st.session_state.pending_preview:
                if not st.session_state.get("submit_batch_id"):
                    # a stable id for this whole staged batch, regardless of how many
                    # rows it has or how many times Submit gets clicked on it — see
                    # already_submitted()'s docstring for why this is content-blind.
                    st.session_state.submit_batch_id = uuid.uuid4().hex
                preview = st.session_state.pending_preview
                staged_new = st.session_state.pending_new_bas

                st.divider()
                st.subheader("📋 Preview — review before submitting")
                st.caption("Click any cell to fix a wrong entry, or remove a row with the 🗑 icon on the right.")

                preview_df = pd.DataFrame(preview)[HEADERS].copy()
                preview_df["DonAmt"] = pd.to_numeric(preview_df["DonAmt"], errors="coerce")
                preview_df["Supports"] = pd.to_numeric(preview_df["Supports"], errors="coerce")
                preview_df.insert(0, "S.No", range(1, len(preview_df) + 1))
                edited_df = st.data_editor(
                    preview_df,
                    use_container_width=True,
                    hide_index=True,
                    num_rows="dynamic",
                    key=f"preview_editor_{st.session_state.preview_nonce}",
                    column_config={
                        "S.No": st.column_config.NumberColumn("S.No", disabled=True),
                        "DonAmt": st.column_config.NumberColumn("DonAmt", min_value=0.01),
                        "Supports": st.column_config.NumberColumn("Supports", disabled=True),
                        "SOD": st.column_config.SelectboxColumn("SOD", options=SOD_CATEGORIES),
                        "Event Name": st.column_config.SelectboxColumn("Event Name", options=A["event_options"]),
                        "Mode of Payment": st.column_config.SelectboxColumn("Mode of Payment", options=MODE_OF_PAYMENT),
                        "No Production": st.column_config.TextColumn("No Production", disabled=True),
                        "Timestamp": st.column_config.TextColumn("Timestamp", disabled=True),
                    },
                )
                edited_records = edited_df.to_dict("records")
                # persist the grid's current state (including any deletion) back to
                # pending_preview immediately, so it survives the next rerun instead
                # of being silently forgotten — see sync_preview_from_grid's docstring.
                st.session_state.pending_preview = sync_preview_from_grid(edited_records)
                ba_count = len({(r.get("OWNCODE"), r.get("BAName")) for r in edited_records if r.get("OWNCODE")})
                st.info(
                    f"**{len(edited_records)} entry(s)** across **{ba_count} BA(s)** — not saved yet. "
                    "Add more BAs with **Save all entries**, or click **Submit** to write them all."
                )

                pc1, pc2 = st.columns(2)
                submit_clicked = pc1.button("✅ Submit", type="primary", use_container_width=True)
                cancel_clicked = pc2.button("✕ Cancel", use_container_width=True)

                if cancel_clicked:
                    st.session_state.pending_preview = []
                    st.session_state.pending_new_bas = []
                    st.session_state.preview_nonce += 1
                    st.session_state.submit_batch_id = None
                    st.rerun()

                if submit_clicked:
                    edit_errors, final_rows = validate_edited_rows(edited_records)
                    if not edit_errors and not final_rows:
                        edit_errors = ["Add at least one entry before submitting."]

                    batch_id = st.session_state.submit_batch_id

                    if edit_errors:
                        for e in edit_errors:
                            st.error(e)
                    elif st.session_state.get("submitting") and not submit_lock_is_stale():
                        # A submission from a moment ago (e.g. a double-click, or a slow
                        # 429 retry) is still in flight for this session. Processing this
                        # click too would submit the same staged rows a second time —
                        # skip instead of resubmitting.
                        st.warning("A submission is already in progress — please wait a moment.")
                    else:
                        st.session_state.submitting = True
                        st.session_state.submitting_since = datetime.now(timezone.utc)
                        try:
                            # Durable check: this exact batch (identified by its random
                            # id, not its row content) may already have been written by
                            # an earlier click whose success message got missed — e.g. a
                            # slow retry that finished after the user gave up and clicked
                            # again. If so, skip re-writing and just show success again.
                            try:
                                already_done = already_submitted(batch_id)
                            except Exception:  # noqa: BLE001
                                already_done = False  # can't confirm either way — fall through to a normal attempt

                            if already_done:
                                st.session_state.session_entries.extend(final_rows)
                                final_ba_count = len({(r["OWNCODE"], r["BAName"]) for r in final_rows})
                                st.session_state.flash_success = (
                                    f"✅ Submitted {len(final_rows)} entry(s) across {final_ba_count} BA(s)."
                                )
                                st.session_state.pending_preview = []
                                st.session_state.pending_new_bas = []
                                st.session_state.submit_batch_id = None
                                st.session_state.nonce += 1
                                st.session_state.ba_nonce += 1
                                st.session_state.preview_nonce += 1
                                st.session_state.rows = [st.session_state.next_id]
                                st.session_state.next_id += 1
                                st.rerun()

                            ok_to_save = True

                            # 1) register any staged new BAs in the Admin sheet first
                            if staged_new:
                                try:
                                    for owner_code, owner_client, cd, effective_ba, joinee_date in staged_new:
                                        if cd and cd != "Unassigned" and cd in A["ba_codes"]:
                                            st.warning(f"BA code **{cd}** already exists in the Admin sheet; saving anyway.")
                                    append_bas([(o, cd, nm, ocl, jd) for o, ocl, cd, nm, jd in staged_new])
                                    load_admin.clear()
                                    for owner_code, owner_client, cd, effective_ba, joinee_date in staged_new:
                                        st.session_state.new_bas.setdefault((owner_code, owner_client), []).append(
                                            (effective_ba, cd)
                                        )
                                    added = ", ".join(f"{nm} ({cd})" for _o, _ocl, cd, nm, _jd in staged_new)
                                    st.success(f"➕ New BA(s) added to Admin sheet: {added}")
                                except Exception as e:  # noqa: BLE001
                                    ok_to_save = False
                                    st.error(f"Couldn't update Admin sheet: {e}")

                            # 2) write all entries
                            if ok_to_save:
                                try:
                                    before, after = append_donations(final_rows)
                                    added = after - before
                                    if added == len(final_rows):
                                        # Confirmed: exactly what we sent landed. Log this batch's
                                        # id right away so a later retry of the same batch is
                                        # recognized as already-done, then clear the preview —
                                        # otherwise a shortfall would silently vanish from the
                                        # UI while the "submitted entries" download still claimed
                                        # it was saved.
                                        try:
                                            log_submission(batch_id, code, len(final_rows))
                                        except Exception:  # noqa: BLE001
                                            pass  # logging failure shouldn't hide a successful write
                                        st.session_state.session_entries.extend(final_rows)
                                        final_ba_count = len({(r["OWNCODE"], r["BAName"]) for r in final_rows})
                                        st.session_state.flash_success = (
                                            f"✅ Submitted {len(final_rows)} entry(s) across {final_ba_count} BA(s)."
                                        )
                                        st.session_state.pending_preview = []
                                        st.session_state.pending_new_bas = []
                                        st.session_state.submit_batch_id = None
                                        st.session_state.nonce += 1
                                        st.session_state.ba_nonce += 1
                                        st.session_state.preview_nonce += 1
                                        st.session_state.rows = [st.session_state.next_id]
                                        st.session_state.next_id += 1
                                        st.rerun()
                                    else:
                                        st.error(
                                            f"Expected to add {len(final_rows)} row(s) but the sheet's row "
                                            f"count only increased by {added}. Nothing has been cleared from "
                                            "your preview — check the Donations One-Off sheet before "
                                            "retrying, and contact the admin if you're unsure."
                                        )
                                except Exception as e:  # noqa: BLE001
                                    st.error(f"Couldn't save to Donations One-Off sheet: {e}")
                                    if "403" in str(e) or "PERMISSION_DENIED" in str(e):
                                        st.caption(
                                            "**403 / permission denied** — the account the app signs in as cannot "
                                            "edit that sheet. Open the Donations One-Off sheet, click **Share**, and give "
                                            "**Editor** access to the right account (your Google login locally, or the "
                                            "service-account email on Streamlit Cloud)."
                                        )
                        finally:
                            st.session_state.submitting = False
                            st.session_state.submitting_since = None
        else:
            if not st.session_state.get("submit_batch_id"):
                st.session_state.submit_batch_id = uuid.uuid4().hex
            st.info(
                "BA and Entry details are disabled while **NO Production** is checked. "
                "Submitting will record a single entry marking no production for this owner "
                "and date — there is no preview step."
            )
            nf_submit_clicked = st.button("✅ Submit", type="primary", use_container_width=True,
                                          key=f"nf_submit_{n}")
            if nf_submit_clicked:
                batch_id = st.session_state.submit_batch_id
                if not code:
                    st.error("Select a valid **owner code**.")
                elif st.session_state.get("submitting") and not submit_lock_is_stale():
                    st.warning("A submission is already in progress — please wait a moment.")
                else:
                    st.session_state.submitting = True
                    st.session_state.submitting_since = datetime.now(timezone.utc)
                    try:
                        try:
                            already_done = already_submitted(batch_id)
                        except Exception:  # noqa: BLE001
                            already_done = False

                        if already_done:
                            st.session_state.flash_success = (
                                f"✅ Submitted 1 'NO Production' entry for owner {code}."
                            )
                            st.session_state.submit_batch_id = None
                            st.session_state.nonce += 1
                            st.rerun()

                        no_prod_row = {
                            "SigninDT": signin.strftime("%Y-%m-%d"), "OWNCODE": code, "OwnerName": owner_name,
                            "BAName": "", "BACode": "", "Clients": client or "",
                            "DonAmt": "", "Supports": "", "SOD": "", "Event Name": "",
                            "Mode of Payment": "", "No Production": "1",
                            "Timestamp": datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S"),
                        }
                        try:
                            before, after = append_donations([no_prod_row])
                            added = after - before
                            if added == 1:
                                try:
                                    log_submission(batch_id, code, 1)
                                except Exception:  # noqa: BLE001
                                    pass  # logging failure shouldn't hide a successful write
                                st.session_state.session_entries.append(no_prod_row)
                                st.session_state.flash_success = (
                                    f"✅ Submitted 1 'NO Production' entry for owner {code}."
                                )
                                st.session_state.submit_batch_id = None
                                st.session_state.nonce += 1
                                st.rerun()
                            else:
                                st.error(
                                    "Expected to add 1 row but the sheet's row count didn't increase by 1. "
                                    "Nothing has been recorded as submitted — check the Donations One-Off "
                                    "sheet before retrying, and contact the admin if you're unsure."
                                )
                        except Exception as e:  # noqa: BLE001
                            st.error(f"Couldn't save to Donations One-Off sheet: {e}")
                            if "403" in str(e) or "PERMISSION_DENIED" in str(e):
                                st.caption(
                                    "**403 / permission denied** — the account the app signs in as cannot "
                                    "edit that sheet. Open the Donations One-Off sheet, click **Share**, and give "
                                    "**Editor** access to the right account (your Google login locally, or the "
                                    "service-account email on Streamlit Cloud)."
                                )
                    finally:
                        st.session_state.submitting = False
                        st.session_state.submitting_since = None


with history_slot.container():
    st.session_state.setdefault("show_owner_history", False)

    if code and passcode_ok:
        if st.session_state.show_owner_history:
            if st.button("← Back to form"):
                st.session_state.show_owner_history = False
                st.rerun()
        else:
            if st.button("📊 View Owners Data History"):
                st.session_state.show_owner_history = True
                st.rerun()
    else:
        st.session_state.show_owner_history = False

    if code and passcode_ok and st.session_state.show_owner_history:
        owner_name_hist = A["owner_name_by_code"].get(code, "")
        st.markdown(f"### Owners Data History — {code} · {owner_name_hist}")
        st.markdown(
            '<style>div[class*="st-key-entry_form"] { display: none; }</style>',
            unsafe_allow_html=True,
        )
        try:
            hist_df = load_donations_df()
            owner_hist = hist_df[hist_df["OWNCODE"].astype(str).str.strip() == code]
            owner_hist = owner_hist.sort_values("SigninDT", ascending=False)
        except Exception as e:  # noqa: BLE001
            st.error(f"Couldn't load history: {e}")
            owner_hist = pd.DataFrame()
        if owner_hist.empty:
            st.info(f"No previous entries found for **{code} · {owner_name_hist}**.")
        else:
            st.caption(
                f"{len(owner_hist)} previous entr{'y' if len(owner_hist) == 1 else 'ies'} "
                f"for **{code} · {owner_name_hist}**"
            )
            st.dataframe(owner_hist[HEADERS], hide_index=True, use_container_width=True)

            st.markdown("##### Download this owner's data for a date")
            available_dates = sorted(owner_hist["SigninDT"].unique(), reverse=True)
            dl1, dl2 = st.columns([2, 1], vertical_alignment="bottom")
            selected_date = dl1.selectbox(
                "Select a date",
                available_dates,
                key=f"hist_dl_date_{code}",
            )
            date_rows = owner_hist[owner_hist["SigninDT"] == selected_date][HEADERS]
            dl2.download_button(
                f"⬇ Download {selected_date} (.xlsx)",
                data=df_to_xlsx_bytes(date_rows, sheet_title="History"),
                file_name=f"{code}_{selected_date}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key=f"hist_dl_btn_{code}",
            )

# --------------------------------------------------------------------------- #
#  Submitted entries (this session) + downloads
# --------------------------------------------------------------------------- #
st.divider()
se = st.session_state.session_entries
st.subheader(f"Submitted entries — this session ({len(se)})")
if se:
    st.dataframe(pd.DataFrame(se)[HEADERS].iloc[::-1], use_container_width=True, hide_index=True)
    st.download_button(
        "⬇  Download “submitted entries” (this session)",
        data=session_xlsx_bytes(se),
        file_name="submitted entries oneoff.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
else:
    st.info("Entries you save in this session will appear here, updating live as you submit.")

# Confirmation shown last, at the very bottom, right after the submitted-entries list.
if st.session_state.get("flash_success"):
    st.success(st.session_state.pop("flash_success"))
