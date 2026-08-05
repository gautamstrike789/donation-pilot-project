"""
Donation Entry — One-Off Owner submission portal (Streamlit + Google Sheets)
=============================================================================
Form order:  Owner Code  →  SignIn Date  →  BA Name  →  Add-new-BA  →  Entries (+ rows)

Data source: same Admin sheet as the main donation form, configured via
sheets_config.json (created by setup.py / extended by setup_oneoff.py):
    • Admin sheet          — worksheets "Owners" (OWNCODE|OwnerName|City) and "BAs" (OWNCODE|BACode|BAName)
    • Donations One-Off sheet — one worksheet with columns:
        SigninDT, OWNCODE, BAName, BACode, Forms, Supports, ADS

ADS (Average Donation per Support) is calculated automatically as
Supports / Forms — it is not a form input.

Uses OAuth (your Google account) — no service account keys needed.

Run:
    python -m pip install -r requirements.txt
    python setup.py            # one-time: creates Admin + Donations sheets + config (if not already done)
    python setup_oneoff.py     # one-time: creates the Donations One-Off sheet + extends config
    python -m streamlit run app_oneoff.py
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
DONATIONS_SHEET_KEY = "oneoff_donations_sheet_id"
HEADERS = ["SigninDT", "OWNCODE", "BAName", "BACode", "Forms", "Supports", "ADS"]
ADS_DECIMALS = 1
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
            "admin_sheet_id": secret_cfg["admin_sheet_id"],
            DONATIONS_SHEET_KEY: secret_cfg[DONATIONS_SHEET_KEY],
        }
    with open(CONFIG_FILE) as f:
        cfg = json.load(f)
    if DONATIONS_SHEET_KEY not in cfg:
        st.error(
            f"**{DONATIONS_SHEET_KEY}** missing from {CONFIG_FILE}. "
            "Run **`python setup_oneoff.py`** first to create the Donations One-Off "
            "sheet and extend the config file."
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

    return {"label_by_code": label_by_code, "code_by_label": code_by_label,
            "owner_meta": owner_meta, "ba_by_code": ba_by_code, "ba_codes": ba_codes}


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
                        if ck == "_ws_donations_oneoff":
                            sh = st.session_state.gc.open_by_key(cfg[DONATIONS_SHEET_KEY])
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
n = st.session_state.nonce
bn = st.session_state.ba_nonce


def parse_positive_int(raw):
    text = str(raw or "").strip()
    if not text:
        return None
    try:
        value = float(text)
    except ValueError:
        return None
    if not value.is_integer():
        return None
    return int(value)


# --------------------------------------------------------------------------- #
#  Header
# --------------------------------------------------------------------------- #
total_bas = sum(len(v) for v in A["ba_by_code"].values())
hc1, hc2 = st.columns([1, 4], vertical_alignment="center")
if os.path.exists(LOGO_FILE):
    hc1.image(LOGO_FILE, width=130)
hc2.title("Donation Entry — One-Off")
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
signin = st.date_input("2 · SignIn Date *", value=date.today(), format="YYYY-MM-DD", key="signin")

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

# ---- 5) Entries (dynamic rows) ----
st.markdown("#### Entries")
st.caption("Add one or more entries for this BA, then save them all at once.")

row_inputs = []
to_remove = None
for idx, rid in enumerate(st.session_state.rows, start=1):
    with st.container(border=True):
        h = st.columns([6, 1])
        h[0].markdown(f"**Entry #{idx}**")
        if len(st.session_state.rows) > 1 and h[1].button("✕", key=f"rm_{rid}", help="Remove"):
            to_remove = rid
        d1, d2 = st.columns(2)
        forms = d1.text_input("Forms *", key=f"forms_{rid}", placeholder="whole number > 0")
        _f = (forms or "").strip()
        forms_v = parse_positive_int(_f) if _f else None
        if _f and (forms_v is None or forms_v <= 0):
            d1.caption(":red[⚠ Enter a whole number greater than 0]")
        supports = d2.text_input("Supports *", key=f"supports_{rid}", placeholder="whole number > 0")
        _s = (supports or "").strip()
        supports_v = parse_positive_int(_s) if _s else None
        if _s and (supports_v is None or supports_v <= 0):
            d2.caption(":red[⚠ Enter a whole number greater than 0]")
        if forms_v and forms_v > 0 and supports_v is not None and supports_v > 0:
            st.caption(f"ADS (Supports ÷ Forms) = **{round(supports_v / forms_v, ADS_DECIMALS)}**")
        row_inputs.append((idx, forms, supports))

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

    # validate every entry row
    valid_rows = []
    for idx, forms, supports in row_inputs:
        forms = (forms or "").strip()
        supports = (supports or "").strip()
        forms_v = parse_positive_int(forms)
        supports_v = parse_positive_int(supports)

        if not any([forms, supports]):
            continue
        if forms_v is None or forms_v <= 0:
            errors.append(f"Entry #{idx}: **Forms** must be a whole number greater than 0.")
        if supports_v is None or supports_v <= 0:
            errors.append(f"Entry #{idx}: **Supports** must be a whole number greater than 0.")
        if forms_v is not None and forms_v > 0 and supports_v is not None and supports_v > 0:
            ads = round(supports_v / forms_v, ADS_DECIMALS)
            valid_rows.append({"SigninDT": signin.strftime("%Y-%m-%d"), "OWNCODE": code,
                               "BAName": effective_ba, "BACode": row_code,
                               "Forms": forms, "Supports": supports, "ADS": ads})

    if not errors and not valid_rows:
        errors.append("Add at least one entry (Forms and Supports) before saving.")

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
        # reset only the "add a new BA" fields and the entry rows
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
    ba_count = len({(r["OWNCODE"], r["BAName"]) for r in preview})

    st.divider()
    st.subheader("📋 Preview — review before submitting")
    st.info(
        f"**{len(preview)} entry(s)** across **{ba_count} BA(s)** — not saved yet. "
        "Add more BAs with **Save all entries**, or click **Submit** to write them all."
    )
    st.dataframe(pd.DataFrame(preview)[HEADERS], use_container_width=True, hide_index=True)

    pc1, pc2 = st.columns(2)
    submit_clicked = pc1.button("✅ Submit", type="primary", use_container_width=True)
    cancel_clicked = pc2.button("✕ Cancel", use_container_width=True)

    if cancel_clicked:
        st.session_state.pending_preview = []
        st.session_state.pending_new_bas = []
        st.rerun()

    if submit_clicked:
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

        # 2) write all entries
        if ok_to_save:
            try:
                before, after = append_donations(preview)
                added = after - before
                st.session_state.session_entries.extend(preview)
                if added < len(preview):
                    # The API returned success but fewer rows than expected landed.
                    st.warning(
                        f"Expected to add {len(preview)} row(s) but only {added} landed. "
                        "Please contact the admin to verify your data was saved correctly."
                    )
                else:
                    st.session_state.flash_success = (
                        f"✅ Submitted {len(preview)} entry(s) across {ba_count} BA(s)."
                    )
                st.session_state.pending_preview = []
                st.session_state.pending_new_bas = []
                st.session_state.nonce += 1
                st.session_state.ba_nonce += 1
                st.session_state.rows = [st.session_state.next_id]
                st.session_state.next_id += 1
                st.rerun()
            except Exception as e:  # noqa: BLE001
                st.error(f"Couldn't save to Donations One-Off sheet: {e}")
                if "403" in str(e) or "PERMISSION_DENIED" in str(e):
                    st.caption(
                        "**403 / permission denied** — the account the app signs in as cannot "
                        "edit that sheet. Open the Donations One-Off sheet, click **Share**, and give "
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
