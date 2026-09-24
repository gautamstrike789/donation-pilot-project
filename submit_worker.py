"""
Background submission worker shared by app.py (RG form) and app_oneoff.py
(One-Off form).

Why this exists
---------------
Streamlit runs the page script on a thread that is thrown away the moment the
same browser session triggers another rerun (any click, a second Submit
click, a widget change, ...). With `runner.fastReruns` (Streamlit's default)
a *new* script thread is started immediately, while the old one is told to
stop at its next `st.*` call. A submit that was doing its Google Sheets
writes on that script thread — often sleeping through 429 "rate limit"
backoff when many owners submit at once — therefore:

    * left the old "submitting" flag set, so the new run showed
      "A submission is already in progress", and then
    * got aborted by Streamlit at its next `st.*` call (e.g. the "High
      traffic" notice), so the rows were never written.

Here the Sheets writes run on a plain Python thread owned by this module
instead, which Streamlit reruns never touch. The page only starts the job
and polls its state. Nothing in this module calls `st.*` or reads
`st.session_state`, so it is safe to run off the script thread.

Duplicate protection (unchanged in spirit from the forms' earlier logic):
    * one job per batch id at a time, process-wide (a second click while the
      first is running just keeps waiting on the first),
    * the SubmitLog check, so a batch already confirmed written is never
      written again,
    * every append is verified by exact row content afterwards and only the
      rows actually missing are resent, counted as a multiset so a batch
      that legitimately contains two identical rows is handled correctly.
"""

import random
import threading
import time
from collections import Counter
from datetime import datetime, timezone

import gspread
import requests

try:  # google-auth network errors (token refresh failures etc.)
    from google.auth.exceptions import TransportError as _AuthTransportError
except Exception:  # noqa: BLE001
    _AuthTransportError = ()

# HTTP status codes worth retrying: rate limit + transient server errors.
TRANSIENT_STATUS = {429, 500, 502, 503, 504}
HTTP_TIMEOUT_SECONDS = 60          # a hung request can't block a job forever
MAX_ATTEMPTS = 8                   # 1+2+4+...+64 s ≈ 2 min of backoff per call
MAX_VERIFY_ROUNDS = 5
JOB_RETENTION_SECONDS = 60 * 60    # finished jobs nobody collected are dropped after 1 h

HIGH_TRAFFIC_NOTICE = (
    "High traffic detected — your data is safe and will be saved in {wait} second{s}. "
    "Please do not close this tab. (Retry {n} of {total})"
)


# --------------------------------------------------------------------------- #
#  Errors + retry helpers
# --------------------------------------------------------------------------- #
class SubmitError(Exception):
    """Raised by a job to report which step failed, and whether new BAs had
    already been written to the Admin sheet before the failure."""

    def __init__(self, stage, cause, bas_written=False):
        super().__init__(str(cause))
        self.stage = stage            # "admin" or "donations"
        self.cause = cause
        self.bas_written = bas_written


def status_code(e):
    """Extract an HTTP status code from a gspread APIError (or None)."""
    resp = getattr(e, "response", None)
    if isinstance(getattr(resp, "status_code", None), int):
        return resp.status_code
    code = getattr(e, "code", None)
    if isinstance(code, int) and code > 0:  # gspread uses -1 for an unparseable error body
        return code
    for c in (429, 401, 403, 500, 502, 503, 504):
        if str(c) in str(e):
            return c
    return None


def is_transient(e):
    if isinstance(e, gspread.exceptions.APIError):
        return status_code(e) in TRANSIENT_STATUS
    return isinstance(e, (requests.exceptions.ConnectionError, requests.exceptions.Timeout)) or (
        bool(_AuthTransportError) and isinstance(e, _AuthTransportError)
    )


def _backoff(attempt):
    # exponential backoff with jitter, so many owners hitting the rate limit
    # at the same moment don't all retry in lock-step and collide again
    return min(2 ** attempt, 64) + random.uniform(0, 1)


# --------------------------------------------------------------------------- #
#  Worksheet handles (process-wide, so reopening a sheet doesn't cost a read
#  request on every submit — all sessions share the same Sheets API quota)
# --------------------------------------------------------------------------- #
_client = None
_ws_cache = {}
_ws_lock = threading.Lock()


class SheetRef:
    """Lazily opened worksheet: `ws_name=None` means the first worksheet."""

    def __init__(self, creds, sheet_id, ws_name=None):
        self.creds = creds
        self.sheet_id = sheet_id
        self.ws_name = ws_name

    def get(self):
        global _client
        key = (self.sheet_id, self.ws_name)
        with _ws_lock:
            ws = _ws_cache.get(key)
            if ws is not None:
                return ws
            if _client is None:
                _client = _authorize(self.creds)
        # open outside the lock so one slow open doesn't stall every job
        ws = _with_retry(lambda: _open(_client, self.sheet_id, self.ws_name))
        with _ws_lock:
            return _ws_cache.setdefault(key, ws)

    def reopen(self):
        """Drop every cached handle and client (e.g. after a 401) and reopen."""
        global _client
        with _ws_lock:
            _ws_cache.clear()
            _client = _authorize(self.creds)
        return self.get()


def _authorize(creds):
    gc = gspread.authorize(creds)
    try:
        gc.set_timeout(HTTP_TIMEOUT_SECONDS)
    except Exception:  # noqa: BLE001 — older gspread without set_timeout
        pass
    return gc


def _open(gc, sheet_id, ws_name):
    sh = gc.open_by_key(sheet_id)
    return sh.sheet1 if ws_name is None else sh.worksheet(ws_name)


def _with_retry(fn, ref=None, notify=None):
    """Call fn() with backoff on 429/5xx/network errors, and one reopen of the
    worksheet on 401 (expired token). Anything else is raised straight away."""
    reauthed = False
    for attempt in range(MAX_ATTEMPTS):
        try:
            return fn()
        except Exception as e:  # noqa: BLE001
            code = status_code(e) if isinstance(e, gspread.exceptions.APIError) else None
            if code == 401 and ref is not None and not reauthed:
                reauthed = True
                ref.reopen()
                continue
            if not is_transient(e) or attempt == MAX_ATTEMPTS - 1:
                raise
            wait = _backoff(attempt)
            if notify and code == 429:
                w = int(round(wait))
                notify(HIGH_TRAFFIC_NOTICE.format(wait=w, s="s" if w != 1 else "",
                                                  n=attempt + 1, total=MAX_ATTEMPTS - 1))
            time.sleep(wait)
    raise RuntimeError("unreachable")


def read_col(ref, col, notify=None):
    return _with_retry(lambda: ref.get().col_values(col), ref, notify)


def _append_once(ref, rows, notify):
    """One append, retried only when Google certainly did NOT apply it (429
    rate limit, 401 expired token). A 5xx, timeout or dropped connection may
    have been applied anyway, so that's returned to the caller to verify
    instead of being blindly resent (which is how duplicates happen)."""
    reauthed = False
    for attempt in range(MAX_ATTEMPTS):
        try:
            ref.get().append_rows(rows, value_input_option="RAW")
            return
        except Exception as e:  # noqa: BLE001
            code = status_code(e) if isinstance(e, gspread.exceptions.APIError) else None
            if code == 401 and not reauthed:
                reauthed = True
                ref.reopen()
                continue
            if code == 429 and attempt < MAX_ATTEMPTS - 1:
                wait = _backoff(attempt)
                if notify:
                    w = int(round(wait))
                    notify(HIGH_TRAFFIC_NOTICE.format(wait=w, s="s" if w != 1 else "",
                                                      n=attempt + 1, total=MAX_ATTEMPTS - 1))
                time.sleep(wait)
                continue
            if is_transient(e):
                return  # outcome unknown — let the verify pass decide
            raise


def append_verified(ref, rows, notify=None, check_first=False):
    """Append `rows` and confirm, by exact content, that every one of them is
    in the sheet; resend only the ones that are missing.

    `check_first` verifies before the first append too — used when this
    batch was attempted before (e.g. its rows landed but the confirmation
    read failed), so a retry never writes a row that's already there. Rows
    are compared on their own columns only, so extra columns an admin adds to
    the right of the sheet can't make every row look "missing"."""
    rows = [[("" if v is None else str(v)) for v in r] for r in rows]
    if not rows:
        return
    width = max(len(r) for r in rows)
    need = Counter(tuple(r + [""] * (width - len(r))) for r in rows)

    def missing_rows():
        values = _with_retry(lambda: ref.get().get_all_values(), ref, notify)
        have = Counter(tuple((list(v[:width]) + [""] * width)[:width]) for v in values)
        out = []
        for key, count in need.items():
            out.extend([list(key)] * (count - min(count, have.get(key, 0))))
        return out

    remaining = missing_rows() if check_first else list(rows)
    for verify_round in range(MAX_VERIFY_ROUNDS):
        if not remaining:
            return
        _append_once(ref, remaining, notify)
        time.sleep(0.5 + 0.5 * verify_round)
        remaining = missing_rows()
    if not remaining:
        return

    raise RuntimeError(
        f"{len(remaining)} of {len(rows)} row(s) could not be confirmed as "
        "written to the sheet after several attempts."
    )


# --------------------------------------------------------------------------- #
#  The submission itself
# --------------------------------------------------------------------------- #
def run_submission(notify, *, bas, donations, submitlog, batch_id, code, donation_rows, new_ba_rows):
    """Everything the Submit button used to do inline, in the same order:
    SubmitLog check → new BAs to the Admin sheet → rows to the Donations
    sheet → log the batch id. Returns {"already": bool, "bas_written": bool};
    raises SubmitError on failure."""
    try:
        already = batch_id in read_col(submitlog, 1, notify)
    except Exception:  # noqa: BLE001 — can't confirm either way; fall through to a normal attempt
        already = False
    if already:
        return {"already": True, "bas_written": False}

    # a batch tried before may have partly landed — look before writing again
    with _jobs_lock:
        retry = batch_id in _attempted
        _attempted.add(batch_id)

    bas_written = False
    if new_ba_rows:
        try:
            append_verified(bas, new_ba_rows, notify, check_first=retry)
        except Exception as e:  # noqa: BLE001
            raise SubmitError("admin", e) from e
        bas_written = True

    try:
        append_verified(donations, donation_rows, notify, check_first=retry)
    except Exception as e:  # noqa: BLE001
        raise SubmitError("donations", e, bas_written=bas_written) from e

    try:
        append_verified(
            submitlog,
            [[batch_id, code, str(len(donation_rows)), datetime.now(timezone.utc).isoformat()]],
            notify,
        )
    except Exception:  # noqa: BLE001 — logging failure shouldn't hide a successful write
        pass
    return {"already": False, "bas_written": bas_written}


# --------------------------------------------------------------------------- #
#  Job registry (process-wide; keyed by the batch id, which is a random uuid)
# --------------------------------------------------------------------------- #
_jobs = {}
_jobs_lock = threading.Lock()
_attempted = set()  # batch ids run_submission has already tried (only ever tiny uuids)


def is_running(job):
    return job is not None and job["thread"].is_alive()


def get_job(job_id):
    with _jobs_lock:
        return _jobs.get(job_id)


def pop_job(job_id):
    with _jobs_lock:
        return _jobs.pop(job_id, None)


def start_job(job_id, meta, target, **kwargs):
    """Start `target(notify, **kwargs)` on a background thread under
    `job_id`, unless a job with that id is already running — in which case
    the running one is returned untouched. `meta` is stored on the job for
    the page to use once it finishes."""
    with _jobs_lock:
        now = time.time()
        for jid in [j for j, jb in _jobs.items()
                    if not is_running(jb) and now - jb["started"] > JOB_RETENTION_SECONDS]:
            _jobs.pop(jid, None)

        existing = _jobs.get(job_id)
        if is_running(existing):
            return existing

        job = {"meta": meta, "notice": "", "result": None, "error": None, "started": now}

        def notify(msg):
            job["notice"] = msg

        def run():
            try:
                job["result"] = target(notify, **kwargs)
            except BaseException as e:  # noqa: BLE001 — surfaced to the page, never lost
                job["error"] = e

        job["thread"] = threading.Thread(target=run, name=f"submit-{job_id[:8]}", daemon=True)
        _jobs[job_id] = job
        job["thread"].start()
        return job
