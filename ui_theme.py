"""
Visual theme for the donation forms (RG + One-Off): colours, fonts, glass
cards, buttons, inputs and the entrance/hover animations. Only the look —
the forms keep their own titles, text and logo, and nothing here reads or
writes data, so they behave exactly as before.
"""

import base64
import html
import os
import time
from functools import lru_cache

import streamlit as st

THEME_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Sora:wght@400;600;700;800&family=Manrope:wght@400;500;600;700;800&display=swap');

:root{
  --bg:#060a1c; --bg-2:#0b1130;
  --glow-a:rgba(99,72,230,.32); --glow-b:rgba(245,180,40,.13); --glow-c:rgba(17,162,111,.14);
  --surface:rgba(22,29,66,.58); --surface-strong:rgba(26,34,76,.9); --surface-sunk:rgba(255,255,255,.04);
  --border:rgba(255,255,255,.09); --border-strong:rgba(255,255,255,.16);
  --text:#f3f4ff; --text-2:#c0c5e6; --text-3:#9298c0;
  --gold:#f6c95c; --gold-fill:#f2b632; --gold-soft:rgba(246,201,92,.14);
  --gold-grad:linear-gradient(120deg,#ffe7a3 0%,#f2b632 40%,#e09a1c 100%);
  --btn-grad:linear-gradient(135deg,#ffe08a 0%,#f2b632 45%,#e08d12 100%);
  --danger:#fb7185; --danger-soft:rgba(251,113,133,.12);
  --shadow:0 1px 1px rgba(0,0,0,.3), 0 18px 40px -18px rgba(0,0,0,.7);
  --ease:cubic-bezier(.16,1,.3,1);
  --font-display:'Sora',ui-sans-serif,system-ui,sans-serif;
  --font-sans:'Manrope',ui-sans-serif,system-ui,sans-serif;
}

/* ---------- page ---------- */
.stApp{ background:var(--bg); color:var(--text); font-family:var(--font-sans);
  -webkit-font-smoothing:antialiased; }
[data-testid="stAppViewContainer"]::before{ content:""; position:fixed; inset:-20vh -20vw; z-index:0;
  pointer-events:none;
  background:
    radial-gradient(40vmax 40vmax at 12% 8%, var(--glow-a), transparent 60%),
    radial-gradient(36vmax 36vmax at 92% 18%, var(--glow-b), transparent 60%),
    radial-gradient(44vmax 44vmax at 70% 100%, var(--glow-c), transparent 60%),
    linear-gradient(180deg, var(--bg) 0%, var(--bg-2) 100%);
  animation:ui-drift 22s ease-in-out infinite alternate; }
[data-testid="stMain"], [data-testid="stAppViewContainer"] > section{ position:relative; z-index:1; }
[data-testid="stHeader"]{ background:transparent; }
.block-container{ padding-top:1.4rem; padding-bottom:2.5rem; max-width:800px; }
::selection{ background:var(--gold-soft); }

h1,h2,h3,h4{ font-family:var(--font-display) !important; color:var(--text) !important; letter-spacing:-.02em; }
h4{ font-weight:700 !important; }
.stApp p, .stApp li, .stApp label, .stApp input, .stApp textarea{ font-family:var(--font-sans); }
[data-testid="stWidgetLabel"] p{ font-weight:700; color:var(--text-2); font-size:.86rem; }
[data-testid="stCaptionContainer"]{ color:var(--text-3); }
hr{ border-color:var(--border) !important; }

/* ---------- header (logo + title + caption) ---------- */
.ui-top{ display:flex; align-items:center; gap:1.1rem; flex-wrap:wrap; padding:1rem 1.15rem; margin:0 0 1.2rem;
  border:1px solid var(--border); border-radius:22px; background:var(--surface);
  backdrop-filter:blur(14px) saturate(140%); -webkit-backdrop-filter:blur(14px) saturate(140%);
  box-shadow:var(--shadow); animation:ui-rise .5s var(--ease) both; }
.ui-logo{ border-radius:14px; padding:.35rem .6rem; background:rgba(255,255,255,.94); line-height:0; flex:none;
  box-shadow:0 1px 0 rgba(255,255,255,.6) inset, 0 6px 18px -10px rgba(0,0,0,.45);
  animation:ui-pop .5s .05s var(--ease) both; transition:transform .25s var(--ease); }
.ui-top:hover .ui-logo{ transform:rotate(-2deg) scale(1.03); }
.ui-logo img{ height:52px; width:auto; }
.ui-titles{ min-width:0; animation:ui-rise .5s .12s var(--ease) both; }
.ui-titles h1{ font-size:clamp(1.55rem, 4.5vw, 2.05rem) !important; font-weight:800 !important;
  margin:0 0 .2rem !important; padding:0 !important; line-height:1.15; }
.ui-titles p{ margin:0; color:var(--text-3); font-weight:600; font-size:.88rem; }
.ui-titles [data-testid="stHeaderActionElements"], .ui-titles h1 a{ display:none !important; }

/* ---------- cards (bordered containers) ---------- */
[data-testid="stVerticalBlockBorderWrapper"]{
  background:var(--surface) !important; border:1px solid var(--border) !important; border-radius:22px !important;
  box-shadow:var(--shadow); backdrop-filter:blur(14px) saturate(140%); -webkit-backdrop-filter:blur(14px) saturate(140%);
  padding:.6rem .6rem; animation:ui-rise .45s var(--ease) both;
  transition:box-shadow .2s ease, border-color .2s ease; }
[data-testid="stVerticalBlockBorderWrapper"]:focus-within{ border-color:rgba(242,182,50,.55) !important;
  box-shadow:0 0 0 4px var(--gold-soft), var(--shadow); }

/* ---------- inputs ---------- */
[data-baseweb="input"], [data-baseweb="select"] > div, [data-baseweb="textarea"]{
  border-radius:14px !important; background:var(--surface-strong) !important;
  border-color:var(--border-strong) !important; transition:border-color .15s, box-shadow .15s; }
[data-baseweb="input"]:focus-within, [data-baseweb="select"] > div:focus-within, [data-baseweb="textarea"]:focus-within{
  border-color:var(--gold-fill) !important; box-shadow:0 0 0 4px var(--gold-soft) !important; }
[data-baseweb="popover"] li[role="option"]{ transition:background .12s ease; }

/* ---------- buttons ---------- */
.stButton > button, .stDownloadButton > button{
  border-radius:14px; font-weight:700; font-family:var(--font-sans); min-height:46px; padding:0 1.15rem;
  color:var(--text); background:var(--surface-sunk); border:1px solid var(--border);
  transition:transform .14s var(--ease), box-shadow .18s ease, filter .15s ease, background .15s ease; }
.stButton > button:hover:not(:disabled), .stDownloadButton > button:hover:not(:disabled){
  background:var(--surface-strong); border-color:var(--border-strong); color:var(--text);
  transform:translateY(-1px); box-shadow:0 8px 20px -12px rgba(0,0,0,.7); }
.stButton > button:active:not(:disabled), .stDownloadButton > button:active:not(:disabled){
  transform:translateY(1px) scale(.99); }
.stButton > button[kind="primary"], button[data-testid="stBaseButton-primary"]{
  position:relative; overflow:hidden; color:#1d1300 !important; background:var(--btn-grad) !important; border:none !important;
  box-shadow:0 8px 24px -10px rgba(242,182,50,.8), inset 0 1px 0 rgba(255,255,255,.5); }
.stButton > button[kind="primary"] p, button[data-testid="stBaseButton-primary"] p{ color:#1d1300 !important; font-weight:800; }
.stButton > button[kind="primary"]::after, button[data-testid="stBaseButton-primary"]::after{
  content:""; position:absolute; inset:0; pointer-events:none;
  background:linear-gradient(110deg,transparent 30%,rgba(255,255,255,.45) 45%,transparent 60%);
  transform:translateX(-100%); }
.stButton > button[kind="primary"]:hover:not(:disabled)::after,
button[data-testid="stBaseButton-primary"]:hover:not(:disabled)::after{ animation:ui-sheen .9s ease; }
.stButton > button[kind="primary"]:hover:not(:disabled), button[data-testid="stBaseButton-primary"]:hover:not(:disabled){
  filter:brightness(1.05); transform:translateY(-1px); box-shadow:0 12px 28px -10px rgba(242,182,50,.9); }
.stButton > button:disabled{ opacity:.55; }

/* ---------- alerts, tables ---------- */
[data-testid="stAlert"]{ border-radius:14px; animation:ui-pop .35s var(--ease) both; }
[data-testid="stAlertContentSuccess"]{ font-weight:700; }
[data-testid="stDataFrame"], [data-testid="stDataEditor"]{ border-radius:16px; overflow:hidden;
  border:1px solid var(--border); box-shadow:var(--shadow); animation:ui-rise .45s var(--ease) both; }

/* ---------- weekly summary table ---------- */
.we-summary-wrap{ overflow-x:auto; margin:.25rem 0 1rem; border-radius:18px; animation:ui-rise .45s var(--ease) both; }
.we-summary{ width:100%; min-width:620px; table-layout:fixed; border-collapse:separate; border-spacing:0;
  border:1px solid var(--border); border-radius:18px; overflow:hidden; background:var(--surface-strong);
  box-shadow:var(--shadow); font-size:.9rem; }
.we-summary th, .we-summary td{ padding:.6rem .35rem; text-align:center; border-bottom:1px solid var(--border);
  word-wrap:break-word; }
.we-summary th{ background:var(--surface-sunk); color:var(--text-2); font-weight:800; font-size:.8rem;
  overflow-wrap:normal; word-wrap:normal; }
.we-summary th.lblh{ width:11%; }
.we-summary th span{ display:block; font-weight:600; font-size:.74rem; color:var(--text-3); }
.we-summary tr:last-child td{ border-bottom:none; }
.we-summary td{ animation:ui-cell .5s var(--ease) both; }
.we-summary td:nth-child(2){animation-delay:.04s} .we-summary td:nth-child(3){animation-delay:.08s}
.we-summary td:nth-child(4){animation-delay:.12s} .we-summary td:nth-child(5){animation-delay:.16s}
.we-summary td:nth-child(6){animation-delay:.20s} .we-summary td:nth-child(7){animation-delay:.24s}
.we-summary td:nth-child(8){animation-delay:.28s} .we-summary td:nth-child(9){animation-delay:.34s}
.we-summary td.lbl{ color:var(--text-2); font-weight:700; text-align:left; padding-left:.7rem; }
.we-summary td.val{ font-family:var(--font-display); font-variant-numeric:tabular-nums lining-nums;
  color:var(--text); font-weight:700; }
.we-summary td.np{ color:var(--danger); font-weight:800; background:var(--danger-soft); font-size:.8rem;
  word-wrap:normal; overflow-wrap:normal; word-break:keep-all; hyphens:none; }
.we-summary td.nd{ color:var(--text-3); font-size:.8rem; }
.we-summary th.tot{ background:var(--gold-soft); color:var(--gold); }
.we-summary th.tot span{ color:var(--gold); }
.we-summary td.tot{ background:var(--gold-soft); color:var(--gold); font-family:var(--font-display);
  font-weight:800; font-variant-numeric:tabular-nums lining-nums; }

/* ---------- monthly total (KPI tile) ---------- */
.month-total{ position:relative; overflow:hidden; display:flex; align-items:baseline; gap:.7rem; flex-wrap:wrap;
  margin:.25rem 0 1rem; padding:1rem 1.15rem; background:var(--surface-strong); border:1px solid var(--border);
  border-radius:18px; box-shadow:var(--shadow); animation:ui-rise .45s var(--ease) both; }
.month-total::after{ content:""; position:absolute; inset:0; pointer-events:none;
  background:linear-gradient(110deg,transparent 30%,rgba(255,255,255,.10) 45%,transparent 60%);
  transform:translateX(-100%); animation:ui-shimmer 4.5s ease-in-out 1s infinite; }
.month-total b{ font-family:var(--font-display); font-size:1.9rem; font-weight:800; line-height:1;
  font-variant-numeric:tabular-nums lining-nums; background:var(--gold-grad); -webkit-background-clip:text;
  background-clip:text; color:transparent; animation:ui-pop .5s .1s var(--ease) both; }
.month-total span{ color:var(--text-2); font-weight:600; }

/* ---------- submitting indicator ---------- */
.ui-progress{ display:flex; align-items:center; gap:.75rem; padding:.85rem 1rem; margin:.4rem 0;
  border-radius:14px; background:var(--gold-soft); border:1px solid rgba(242,182,50,.3);
  color:var(--text); font-weight:700; animation:ui-pop .3s var(--ease) both; }
.ui-spinner{ width:20px; height:20px; flex:none; border-radius:50%; border:2.5px solid rgba(255,255,255,.18);
  border-top-color:#e08d12; animation:ui-spin .8s linear infinite; }
.ui-progress small{ display:block; font-weight:600; color:var(--text-2); margin-top:.15rem; }

/* ---------- opening reminder popup ---------- */
.ui-popup{ position:fixed; inset:0; z-index:1000000; display:flex; align-items:flex-start; justify-content:center;
  padding:18vh 16px 0; pointer-events:none; background:rgba(4,7,22,.5);
  backdrop-filter:blur(3px); -webkit-backdrop-filter:blur(3px);
  animation:ui-popup-life var(--t) linear var(--d) both; }
.ui-popup-card{ position:relative; overflow:hidden; width:min(440px, 100%); padding:1.25rem 1.35rem 1.4rem;
  display:flex; align-items:center; gap:.9rem; border-radius:20px; background:var(--surface-strong);
  border:1px solid rgba(242,182,50,.45); box-shadow:0 24px 60px -20px rgba(0,0,0,.8), 0 0 0 6px var(--gold-soft);
  animation:ui-popup-in .5s var(--ease) var(--d) both; }
.ui-popup-icon{ width:44px; height:44px; flex:none; border-radius:14px; display:grid; place-items:center;
  font-size:1.35rem; background:var(--gold-soft); animation:ui-popup-nudge 1.2s ease-in-out var(--d) 2; }
.ui-popup-text{ font-family:var(--font-display); font-weight:700; font-size:1.12rem; color:var(--text); line-height:1.3; }
.ui-popup-bar{ position:absolute; left:0; bottom:0; height:4px; width:100%; background:var(--btn-grad);
  transform-origin:left; animation:ui-popup-bar var(--t) linear var(--d) both; }
@keyframes ui-popup-life{ 0%{opacity:0} 7%{opacity:1} 88%{opacity:1} 100%{opacity:0; visibility:hidden} }
@keyframes ui-popup-in{ from{opacity:0; transform:translateY(-18px) scale(.94)} to{opacity:1; transform:none} }
@keyframes ui-popup-bar{ from{transform:scaleX(1)} to{transform:scaleX(0)} }
@keyframes ui-popup-nudge{ 0%,100%{transform:none} 25%{transform:rotate(-10deg)} 75%{transform:rotate(10deg)} }

/* ---------- SignIn Date confirm dialog ---------- */
div[role="dialog"]{ background:var(--surface-strong) !important; border:1px solid var(--border-strong) !important;
  border-radius:22px !important; box-shadow:0 30px 80px -24px rgba(0,0,0,.85), 0 0 0 6px var(--gold-soft) !important;
  animation:ui-popup-in .4s var(--ease) both; }
div[role="dialog"] h2, div[role="dialog"] [data-testid="stHeading"]{ font-family:var(--font-display) !important; }
.ui-confirm-date{ display:flex; flex-direction:column; gap:.55rem; padding:1rem 1.1rem; margin:.1rem 0 .9rem;
  border-radius:16px; background:var(--surface-sunk); border:1px solid var(--border); }
.ui-confirm-date span{ color:var(--text-2); font-weight:700; }
.ui-confirm-date b{ font-family:var(--font-display); font-size:1.45rem; font-weight:800; line-height:1.2;
  background:var(--gold-grad); -webkit-background-clip:text; background-clip:text; color:transparent;
  animation:ui-pop .45s .08s var(--ease) both; }

/* ---------- motion ---------- */
@keyframes ui-rise{ from{opacity:0; transform:translateY(14px)} to{opacity:1; transform:none} }
@keyframes ui-cell{ from{opacity:0; transform:translateY(6px)} to{opacity:1; transform:none} }
@keyframes ui-pop{ from{opacity:0; transform:scale(.96)} to{opacity:1; transform:none} }
@keyframes ui-sheen{ from{transform:translateX(-100%)} to{transform:translateX(100%)} }
@keyframes ui-shimmer{ 0%,55%{transform:translateX(-100%)} 100%{transform:translateX(100%)} }
@keyframes ui-spin{ to{transform:rotate(360deg)} }
@keyframes ui-drift{ from{transform:translate3d(0,0,0)} to{transform:translate3d(-2vw,1.5vh,0)} }
@media (prefers-reduced-motion: reduce){
  *, *::before, *::after{ animation-duration:.001ms !important; animation-iteration-count:1 !important;
    transition-duration:.001ms !important; }
  .ui-popup, .ui-popup-bar{ animation-duration:var(--t) !important; }
}
@media (max-width: 520px){
  .ui-logo img{ height:40px; }
}
</style>
"""


@lru_cache(maxsize=4)
def _logo_data_uri(path):
    if not os.path.isabs(path):
        # next to this file, so it works whatever folder the app is started from
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)), path)
    if not os.path.exists(path):
        return ""
    with open(path, "rb") as f:
        return "data:image/png;base64," + base64.b64encode(f.read()).decode()


def apply_theme():
    st.markdown(THEME_CSS, unsafe_allow_html=True)


def render_header(title, subtitle, logo_path="logo.png"):
    """The form's own header — logo, title and caption — in a glass card."""
    logo = _logo_data_uri(logo_path)
    logo_html = f'<div class="ui-logo"><img src="{logo}" alt="logo"></div>' if logo else ""
    st.markdown(
        f'<div class="ui-top">{logo_html}<div class="ui-titles"><h1>{title}</h1><p>{subtitle}</p></div></div>',
        unsafe_allow_html=True,
    )


def progress_html(notice=""):
    """The "Submitting…" indicator with a spinner (shown while a submit is being written)."""
    extra = f"<small>{notice}</small>" if notice else ""
    return (f'<div class="ui-progress" role="status"><span class="ui-spinner"></span>'
            f'<div>Submitting… please wait and don\'t close this tab.{extra}</div></div>')


def reminder_popup(message, seconds=4):
    """A popup shown when the form is opened, fading away by itself after
    `seconds`. Once per browser session; keeps its own clock across reruns
    (owner typing straight away doesn't restart or cut it short), and never
    blocks clicks underneath it."""
    shown_at = st.session_state.setdefault("_reminder_popup_at", time.time())
    elapsed = time.time() - shown_at
    if elapsed >= seconds:
        return
    st.markdown(
        f'<div class="ui-popup" role="alert" style="--t:{seconds}s; --d:-{elapsed:.2f}s">'
        f'<div class="ui-popup-card"><div class="ui-popup-icon">⚠️</div>'
        f'<div class="ui-popup-text">{html.escape(message)}</div><div class="ui-popup-bar"></div></div></div>',
        unsafe_allow_html=True,
    )
