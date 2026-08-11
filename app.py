"""
Donation Entry — Owner submission portal (Streamlit + Google Sheets)
=====================================================================
Form order:  Owner Code  →  SignIn Date  →  BA Name  →  Add-new-BA  →  Donations (+ rows)

Data source: two Google Sheets configured via sheets_config.json (created by setup.py):
    • Admin sheet  — worksheets "Owners" (OWNCODE|OwnerName|City), "BAs" (OWNCODE|BACode|BAName),
      "Events" (OWNCODE|OwnerName|EventName, added by setup_events.py), and "Airports"
      (AirportName, added by setup_airports.py)
    • Donations sheet — one worksheet with columns:
        SigninDT, OWNCODE, OwnerName, BAName, BACode, Amount(Amt), Age, SOD, Event Name, Airport Name

When SOD is "Events", a second dropdown offers the Event Names mapped to the
selected owner (from the Events worksheet). When SOD is "Airport", a second
dropdown offers the airport list from the Airports worksheet — edit that sheet
directly to add/remove/rename an airport. Either selection is saved into its
own column in the Donations sheet.

Uses OAuth (your Google account) — no service account keys needed.
No file-locking issues — you can keep both sheets open in your browser while submitting.

Run:
    python -m pip install -r requirements.txt
    python setup.py          # one-time: creates sheets + config
    python -m streamlit run app.py
"""

import json
import os
import pickle
import time
from datetime import date
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
HEADERS = ["SigninDT", "OWNCODE", "OwnerName", "BAName", "BACode", "Amount(Amt)", "Age", "SOD", "Event Name", "Airport Name", "NO Production"]
SOD_CATEGORIES = ["B2B/Commercial", "D2D/Resi", "Events", "Streets", "Airport", "Roadtrip"]
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

st.set_page_config(
    page_title="Donation Entry — Owner Portal",
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
            "admin_sheet_id": secret_cfg["admin_sheet_id"],
            "donations_sheet_id": secret_cfg["donations_sheet_id"],
        }
    with open(CONFIG_FILE) as f:
        return json.load(f)


def cloud_secrets_ready():
    return _has_secret(SERVICE_ACCOUNT_SECTION) and _has_secret(SECRETS_SECTION)


# --------------------------------------------------------------------------- #
#  Read dropdown data from the Admin Google Sheet (cached 2 min)
# --------------------------------------------------------------------------- #
@st.cache_data(ttl=120, show_spinner="Loading dropdown data…")
def load_admin():
    cfg = load_config()
    gc = get_gc()
    sh = gc.open_by_key(cfg["admin_sheet_id"])

    def sheet_rows(ws_name):
        values = sh.worksheet(ws_name).get_all_values()
        if not values:
            return []
        headers = values[0]
        return [dict(zip(headers, row)) for row in values[1:]]

    owners_rows = sheet_rows("Owners")
    bas_rows = sheet_rows("BAs")

    label_by_code, code_by_label, owner_meta = {}, {}, {}
    for r in owners_rows:
        c = str(r.get("OWNCODE", "")).strip()
        if not c:
            continue
        meta = " — ".join(p for p in [str(r.get("OwnerName", "")).strip(),
                                      str(r.get("City", "")).strip()] if p)
        label = f"{c}  ·  {meta}" if meta else c
        label_by_code[c] = label
        code_by_label[label] = c
        owner_meta[c] = (str(r.get("OwnerName", "")).strip(), str(r.get("City", "")).strip())

    ba_by_code, ba_codes = {}, set()
    for r in bas_rows:
        c = str(r.get("OWNCODE", "")).strip()
        name = str(r.get("BAName", "")).strip()
        bacode = str(r.get("BACode", "")).strip()
        if not c or not name:
            continue
        ba_by_code.setdefault(c, [])
        if name not in [t[0] for t in ba_by_code[c]]:
            ba_by_code[c].append((name, bacode))
        if bacode:
            ba_codes.add(bacode)
    for c in ba_by_code:
        ba_by_code[c] = sorted(ba_by_code[c], key=lambda t: t[0].lower())

    events_rows = sheet_rows("Events")
    events_by_code = {}
    for r in events_rows:
        c = str(r.get("OWNCODE", "")).strip()
        name = str(r.get("EventName", "")).strip()
        if not c or not name:
            continue
        events_by_code.setdefault(c, [])
        if name not in events_by_code[c]:
            events_by_code[c].append(name)
    for c in events_by_code:
        events_by_code[c] = sorted(events_by_code[c], key=str.lower)

    airports_rows = sheet_rows("Airports")
    airport_options = []
    for r in airports_rows:
        name = str(r.get("AirportName", "")).strip()
        if name and name not in airport_options:
            airport_options.append(name)

    return {"label_by_code": label_by_code, "code_by_label": code_by_label,
            "owner_meta": owner_meta, "ba_by_code": ba_by_code, "ba_codes": ba_codes,
            "events_by_code": events_by_code, "airport_options": airport_options}


def _get_status_code(e):
    """Extract HTTP status code from a gspread APIError."""
    if hasattr(e, "response") and hasattr(e.response, "status_code"):
        return e.response.status_code
    for code in (429, 401, 403, 500, 503):
        if str(code) in str(e):
            return code
    return None


def _append_with_retry(ws, rows, cache_keys, max_attempts=8):
    """Append rows with exponential backoff for 429 (rate limit) and
    automatic session reset for 401/403 (expired token).
    cache_keys: tuple of session_state keys to clear on auth failure."""
    notice = st.empty()
    try:
        for attempt in range(max_attempts):
            try:
                ws.append_rows(rows, value_input_option="RAW")
                return
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
                        if ck == "_ws_donations":
                            sh = st.session_state.gc.open_by_key(cfg["donations_sheet_id"])
                            st.session_state[ck] = sh.sheet1
                            ws = st.session_state[ck]
                        elif ck == "_ws_bas":
                            sh = st.session_state.gc.open_by_key(cfg["admin_sheet_id"])
                            st.session_state[ck] = sh.worksheet("BAs")
                            ws = st.session_state[ck]
                else:
                    raise
    finally:
        notice.empty()


def append_bas(new_rows):
    """new_rows: list of (OWNCODE, BACode, BAName) -> appended to the BAs worksheet."""
    cfg = load_config()
    ws = get_ws(cfg["admin_sheet_id"], "_ws_bas", "BAs")
    _append_with_retry(ws, [list(r) for r in new_rows], cache_keys=("_ws_bas",))


def append_donations(rows):
    """rows: list of dicts keyed by HEADERS -> appended to the Donations sheet.
    Returns (rows_before, rows_after) read straight back from the sheet so the
    caller can prove the write actually landed (data row count excludes header)."""
    cfg = load_config()
    ws = get_ws(cfg["donations_sheet_id"], "_ws_donations", 0)
    before = max(len(ws.get_all_values()) - 1, 0)
    _append_with_retry(ws, [[r[h] for h in HEADERS] for r in rows], cache_keys=("_ws_donations",))
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
n = st.session_state.nonce
bn = st.session_state.ba_nonce


def parse_age_value(raw_age):
    text = str(raw_age or "").strip()
    if not text:
        return None
    try:
        age_value = float(text)
    except ValueError:
        return None
    if not age_value.is_integer():
        return None
    return int(age_value)


def fmt_number(v):
    """Render a numeric cell (possibly numpy float64 from the data editor) back to
    a plain string, without a spurious trailing .0 for whole numbers."""
    try:
        f = float(v)
    except (TypeError, ValueError):
        return str(v) if v is not None else ""
    return str(int(f)) if f.is_integer() else str(f)


def validate_edited_rows(records):
    """Re-validate rows coming back from the editable preview grid (values may
    have been typed/changed there). Returns (errors, clean_rows) — clean_rows
    is only meaningful when errors is empty."""
    errors, clean_rows = [], []
    for i, r in enumerate(records, start=1):
        signindt = str(r.get("SigninDT") or "").strip()
        ownc = str(r.get("OWNCODE") or "").strip()
        ownername = str(r.get("OwnerName") or "").strip()
        ban = str(r.get("BAName") or "").strip()
        bac = str(r.get("BACode") or "").strip()
        sod_v = str(r.get("SOD") or "").strip()
        event_v = str(r.get("Event Name") or "").strip()
        airport_v = str(r.get("Airport Name") or "").strip()
        amt_raw, age_raw = r.get("Amount(Amt)"), r.get("Age")

        # a blank row added via the grid's "+" control — skip silently
        if not any([signindt, ownc, ban, sod_v, amt_raw not in (None, ""), age_raw not in (None, "")]):
            continue

        try:
            amt_v = float(amt_raw)
        except (TypeError, ValueError):
            amt_v = None
        age_v = parse_age_value(age_raw)

        if not signindt:
            errors.append(f"Row {i}: **SigninDT** is required.")
        if not ownc:
            errors.append(f"Row {i}: **OWNCODE** is required.")
        if not ban:
            errors.append(f"Row {i}: **BAName** is required.")
        if amt_v is None or amt_v < 499:
            errors.append(f"Row {i}: **Amount(Amt)** must be 499 or more.")
        if age_v is None or not (24 < age_v < 100):
            errors.append(f"Row {i}: **Age** must be a whole number between 25 and 99.")
        if not sod_v or sod_v not in SOD_CATEGORIES:
            errors.append(f"Row {i}: **SOD** must be one of {', '.join(SOD_CATEGORIES)}.")
        elif sod_v == "Events" and not event_v:
            errors.append(f"Row {i}: **Event Name** is required when SOD is Events.")
        elif sod_v == "Airport" and not airport_v:
            errors.append(f"Row {i}: **Airport Name** is required when SOD is Airport.")

        clean_rows.append({
            "SigninDT": signindt, "OWNCODE": ownc, "OwnerName": ownername, "BAName": ban, "BACode": bac,
            "Amount(Amt)": fmt_number(amt_raw), "Age": fmt_number(age_raw), "SOD": sod_v,
            "Event Name": event_v if sod_v == "Events" else "",
            "Airport Name": airport_v if sod_v == "Airport" else "",
            "NO Production": "0",
        })
    return errors, clean_rows

# --------------------------------------------------------------------------- #
#  Header
# --------------------------------------------------------------------------- #
total_bas = sum(len(v) for v in A["ba_by_code"].values())
hc1, hc2 = st.columns([1, 4], vertical_alignment="center")
if os.path.exists(LOGO_FILE):
    hc1.image(LOGO_FILE, width=130)
hc2.title("Donation Entry")
hc2.caption(f"{len(A['label_by_code'])} owners · {total_bas:,} BAs · source: Google Sheets")

# ---- 1) Owner Code ----
owner_label = st.selectbox(
    "1 · Owner Code (OWNCODE) *",
    list(A["label_by_code"].values()),
    index=None,
    placeholder="Search your owner code…",
    key="owner",
)
code = A["code_by_label"].get(owner_label) if owner_label else None

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
        combined = {}
        for nm_, cd_ in (A["ba_by_code"].get(code, []) if code else []):
            combined[nm_] = cd_
        for nm_, cd_ in (st.session_state.new_bas.get(code, []) if code else []):
            combined.setdefault(nm_, cd_)
        if code:
            for owner_c, cd_, nm_ in st.session_state.pending_new_bas:
                if owner_c == code:
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

        # ---- 5) Donations (dynamic rows) ----
        st.markdown("#### Donations")
        st.caption("Add one or more donations for this BA, then save them all at once.")

        row_inputs = []
        to_remove = None
        for idx, rid in enumerate(st.session_state.rows, start=1):
            with st.container(border=True):
                h = st.columns([6, 1])
                h[0].markdown(f"**Donation #{idx}**")
                if len(st.session_state.rows) > 1 and h[1].button("✕", key=f"rm_{rid}", help="Remove"):
                    to_remove = rid
                d1, d2 = st.columns(2)
                amt = d1.text_input("Amount (Amt) *", key=f"amt_{rid}", placeholder="min 499")
                _a = (amt or "").strip()
                if _a:
                    try:
                        if float(_a) < 499:
                            d1.caption(":red[⚠ Amount must be 499 or more]")
                    except ValueError:
                        d1.caption(":red[⚠ Enter a valid number]")
                age = d2.text_input("Age *", key=f"age_{rid}", placeholder="25–99")
                _g = (age or "").strip()
                if _g:
                    age_hint_value = parse_age_value(_g)
                    if age_hint_value is None or not (24 < age_hint_value < 100):
                        d2.caption(":red[⚠ Enter a whole number]")
                sod = st.selectbox("Source of Donation (SOD) *", SOD_CATEGORIES, index=None,
                                   placeholder="Select a source…", key=f"sod_{rid}")

                event_name, airport_name = "", ""
                if sod == "Events":
                    owner_events = A["events_by_code"].get(code, []) if code else []
                    if not code:
                        st.selectbox("Event Name *", [], index=None, disabled=True,
                                     placeholder="Select owner code first…", key=f"event_ph_{rid}")
                    elif not owner_events:
                        st.selectbox("Event Name *", [], index=None, disabled=True,
                                     placeholder="No events mapped for this owner", key=f"event_ph_{rid}")
                    else:
                        event_name = st.selectbox("Event Name *", owner_events, index=None,
                                                  placeholder="Select an event…", key=f"event_{rid}")
                elif sod == "Airport":
                    airport_name = st.selectbox("Airport Name *", A["airport_options"], index=None,
                                                placeholder="Select an airport…", key=f"airport_{rid}")

                row_inputs.append((idx, amt, age, sod, event_name, airport_name))

        if to_remove is not None:
            st.session_state.rows.remove(to_remove)
            st.rerun()

        ca, cs = st.columns(2)
        if ca.button("➕ Add another donation", use_container_width=True):
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
            existing_lower = {t[0].lower() for t in (A["ba_by_code"].get(code, []) if code else [])}
            existing_lower |= {t[0].lower() for t in (st.session_state.new_bas.get(code, []) if code else [])}
            existing_lower |= {nm_.lower() for owner_c, cd_, nm_ in st.session_state.pending_new_bas if owner_c == code}
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

            # validate every donation row
            valid_rows = []
            for idx, amt, age, sod, event_name, airport_name in row_inputs:
                amt = (amt or "").strip()
                age = (age or "").strip()
                try:
                    amt_v = float(amt)
                except ValueError:
                    amt_v = None
                age_v = parse_age_value(age)

                if not any([amt, age, sod]):
                    continue
                if amt_v is None or amt_v < 499:
                    errors.append(f"Donation #{idx}: amount must be **499 or more**.")
                if age_v is None or not (24 < age_v < 100):
                    errors.append(f"Donation #{idx}: age must be **between 25 and 99**.")
                if not sod:
                    errors.append(f"Donation #{idx}: select a **source of donation**.")
                elif sod == "Events" and not event_name:
                    errors.append(f"Donation #{idx}: select an **event name**.")
                elif sod == "Airport" and not airport_name:
                    errors.append(f"Donation #{idx}: select an **airport**.")
                if (amt_v is not None and amt_v >= 499 and age_v is not None and 24 < age_v < 100 and sod
                        and (sod != "Events" or event_name) and (sod != "Airport" or airport_name)):
                    owner_name = A["owner_meta"].get(code, ("", ""))[0]
                    valid_rows.append({"SigninDT": signin.strftime("%Y-%m-%d"), "OWNCODE": code,
                                       "OwnerName": owner_name, "BAName": effective_ba, "BACode": row_code,
                                       "Amount(Amt)": amt, "Age": age, "SOD": sod,
                                       "Event Name": event_name, "Airport Name": airport_name,
                                       "NO Production": "0"})

            if not errors and not valid_rows:
                errors.append("Add at least one donation (amount, age, source) before saving.")

            if errors:
                for e in errors:
                    st.error(e)
                # keep whatever is already staged in the preview; just don't add this invalid batch
            else:
                # accumulate validated rows into the running preview (NOT saved to Sheets yet)
                st.session_state.pending_preview.extend(valid_rows)
                # stage a new BA (if any) for the Admin sheet on Submit — dedup by owner + name
                if is_new and effective_ba:
                    already = {(o, nm.lower()) for o, _c, nm in st.session_state.pending_new_bas}
                    if (code, effective_ba.lower()) not in already:
                        st.session_state.pending_new_bas.append((code, cd, effective_ba))
                # clear for the next batch: keep owner code, sign-in date, and BA name;
                # reset only the "add a new BA" fields and the donation rows
                st.session_state.nonce += 1
                st.session_state.rows = [st.session_state.next_id]
                st.session_state.next_id += 1
                st.rerun()

        # --------------------------------------------------------------------------- #
        #  Step 2: Preview + Submit button (writes to Google Sheets only on confirm)
        # --------------------------------------------------------------------------- #
        if st.session_state.pending_preview:
            preview = st.session_state.pending_preview
            staged_new = st.session_state.pending_new_bas

            st.divider()
            st.subheader("📋 Preview — review before submitting")
            st.caption("Click any cell to fix a wrong entry, or remove a row with the 🗑 icon on the right.")

            preview_df = pd.DataFrame(preview)[HEADERS].copy()
            preview_df["Amount(Amt)"] = pd.to_numeric(preview_df["Amount(Amt)"], errors="coerce")
            preview_df["Age"] = pd.to_numeric(preview_df["Age"], errors="coerce")
            preview_df.insert(0, "S.No", range(1, len(preview_df) + 1))
            edited_df = st.data_editor(
                preview_df,
                use_container_width=True,
                hide_index=True,
                num_rows="dynamic",
                key=f"preview_editor_{st.session_state.preview_nonce}",
                column_config={
                    "S.No": st.column_config.NumberColumn("S.No", disabled=True),
                    "Amount(Amt)": st.column_config.NumberColumn("Amount(Amt)", min_value=0, step=1),
                    "Age": st.column_config.NumberColumn("Age", min_value=0, max_value=120, step=1, format="%d"),
                    "SOD": st.column_config.SelectboxColumn("SOD", options=SOD_CATEGORIES),
                    "Airport Name": st.column_config.SelectboxColumn("Airport Name", options=A["airport_options"]),
                    "NO Production": st.column_config.TextColumn("NO Production", disabled=True),
                },
            )
            edited_records = edited_df.to_dict("records")
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
                st.rerun()

            if submit_clicked:
                edit_errors, final_rows = validate_edited_rows(edited_records)
                if not edit_errors and not final_rows:
                    edit_errors = ["Add at least one donation before submitting."]

                if edit_errors:
                    for e in edit_errors:
                        st.error(e)
                else:
                    ok_to_save = True

                    # 1) register any staged new BAs in the Admin sheet first
                    if staged_new:
                        try:
                            for owner_code, cd, effective_ba in staged_new:
                                if cd and cd != "Unassigned" and cd in A["ba_codes"]:
                                    st.warning(f"BA code **{cd}** already exists in the Admin sheet; saving anyway.")
                            append_bas([(o, cd, nm) for o, cd, nm in staged_new])
                            load_admin.clear()
                            for owner_code, cd, effective_ba in staged_new:
                                st.session_state.new_bas.setdefault(owner_code, []).append((effective_ba, cd))
                            added = ", ".join(f"{nm} ({cd})" for _o, cd, nm in staged_new)
                            st.success(f"➕ New BA(s) added to Admin sheet: {added}")
                        except Exception as e:  # noqa: BLE001
                            ok_to_save = False
                            st.error(f"Couldn't update Admin sheet: {e}")

                    # 2) write all donations
                    if ok_to_save:
                        try:
                            before, after = append_donations(final_rows)
                            added = after - before
                            st.session_state.session_entries.extend(final_rows)
                            if added < len(final_rows):
                                # The API returned success but fewer rows than expected landed.
                                st.warning(
                                    f"Expected to add {len(final_rows)} row(s) but only {added} landed. "
                                    "Please contact the admin to verify your data was saved correctly."
                                )
                            else:
                                final_ba_count = len({(r["OWNCODE"], r["BAName"]) for r in final_rows})
                                st.session_state.flash_success = (
                                    f"✅ Submitted {len(final_rows)} donation(s) across {final_ba_count} BA(s)."
                                )
                            st.session_state.pending_preview = []
                            st.session_state.pending_new_bas = []
                            st.session_state.nonce += 1
                            st.session_state.ba_nonce += 1
                            st.session_state.preview_nonce += 1
                            st.session_state.rows = [st.session_state.next_id]
                            st.session_state.next_id += 1
                            st.rerun()
                        except Exception as e:  # noqa: BLE001
                            st.error(f"Couldn't save to Donations sheet: {e}")
                            if "403" in str(e) or "PERMISSION_DENIED" in str(e):
                                st.caption(
                                    "**403 / permission denied** — the account the app signs in as cannot "
                                    "edit that sheet. Open the Donations sheet, click **Share**, and give "
                                    "**Editor** access to the right account (your Google login locally, or the "
                                    "service-account email on Streamlit Cloud)."
                                )
    else:
        st.info(
            "BA and Donation details are disabled while **NO Production** is checked. "
            "Submitting will record a single entry marking no production for this owner "
            "and date — there is no preview step."
        )
        np_submit_clicked = st.button("✅ Submit", type="primary", use_container_width=True,
                                      key=f"np_submit_{n}")
        if np_submit_clicked:
            if not code:
                st.error("Select a valid **owner code**.")
            else:
                owner_name = A["owner_meta"].get(code, ("", ""))[0]
                no_prod_row = {
                    "SigninDT": signin.strftime("%Y-%m-%d"), "OWNCODE": code, "OwnerName": owner_name,
                    "BAName": "", "BACode": "", "Amount(Amt)": "", "Age": "", "SOD": "",
                    "Event Name": "", "Airport Name": "", "NO Production": "1",
                }
                try:
                    before, after = append_donations([no_prod_row])
                    added = after - before
                    st.session_state.session_entries.append(no_prod_row)
                    if added < 1:
                        st.warning(
                            "Expected to add 1 row but it did not land. "
                            "Please contact the admin to verify your data was saved correctly."
                        )
                    else:
                        st.session_state.flash_success = (
                            f"✅ Submitted 1 'NO Production' entry for owner {code}."
                        )
                    st.session_state.nonce += 1
                    st.rerun()
                except Exception as e:  # noqa: BLE001
                    st.error(f"Couldn't save to Donations sheet: {e}")
                    if "403" in str(e) or "PERMISSION_DENIED" in str(e):
                        st.caption(
                            "**403 / permission denied** — the account the app signs in as cannot "
                            "edit that sheet. Open the Donations sheet, click **Share**, and give "
                            "**Editor** access to the right account (your Google login locally, or the "
                            "service-account email on Streamlit Cloud)."
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
        "⬇  Download \u201csubmitted entries\u201d (this session)",
        data=session_xlsx_bytes(se),
        file_name="submitted entries.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
else:
    st.info("Entries you save in this session will appear here, updating live as you submit.")

# Confirmation shown last, at the very bottom, right after the submitted-entries list.
if st.session_state.get("flash_success"):
    st.success(st.session_state.pop("flash_success"))
