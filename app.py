import io
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime
from zoneinfo import ZoneInfo

import requests
import streamlit as st
from openpyxl import load_workbook

REFRESH_SECONDS = 60
TIMEZONE = "America/Chicago"
URL = st.secrets["EXCEL_URL"]          # set in the app's Secrets, not in code

st.set_page_config(page_title="Data Collection", layout="wide")
st.markdown("""
<style>
#MainMenu, header, footer, [data-testid="stToolbar"], [data-testid="stDecoration"],
[data-testid="stStatusWidget"] { display: none !important; }
.stApp { background: #fff; }
.block-container { max-width: 100%; padding: 1rem 2rem 0.5rem 2rem; }
.title { background:#1f3a5f; color:#fff; text-align:center; font-size:2.2rem; font-weight:700; padding:.6rem; }
.label { text-align:center; color:#595959; font-weight:700; font-size:1.3rem; margin-top:.8rem; }
.big   { text-align:center; color:#1f3a5f; font-weight:800; font-size:6rem; line-height:1.1; }
.mid   { text-align:center; color:#1f3a5f; font-weight:800; font-size:3rem; line-height:1.1; }
.sub   { text-align:center; color:#000; font-size:1.1rem; margin-bottom:.4rem; }
.bar   { display:flex; gap:3px; height:2.6rem; }
.seg   { flex:1; background:#d9d9d9; }
.tiles { display:grid; grid-template-columns:repeat(4,1fr); gap:6px; margin-top:.6rem; }
.tile  { background:#f2f2f2; text-align:center; padding:.3rem; }
.tile .k { color:#595959; font-weight:700; font-size:.85rem; }
.tile .v { color:#1f3a5f; font-weight:800; font-size:2rem; }
.tests { display:grid; grid-template-columns:repeat(4,1fr); gap:4px; margin-top:.3rem; }
.test  { text-align:center; font-weight:700; padding:.55rem; background:#f2f2f2; color:#595959; }
.test.part { background:#ffeb9c; color:#9c5700; }
.test.done { background:#c6efce; color:#006100; }
.status { text-align:right; color:#888; font-size:.8rem; margin-top:.3rem; }
.err    { text-align:right; color:#e74c3c; font-size:.9rem; font-weight:700; }
</style>
""", unsafe_allow_html=True)


@st.cache_data(ttl=REFRESH_SECONDS, show_spinner=False)
def load():
    r = requests.get(URL, timeout=30, allow_redirects=True)
    r.raise_for_status()
    if r.content[:2] != b"PK":
        raise ValueError("Link did not return an Excel file")
    wb = load_workbook(io.BytesIO(r.content), data_only=True)
    ws = wb["Totals"]
    # When the file was last saved: workbook property first, HTTP header as backup
    tz = ZoneInfo(TIMEZONE)
    saved = wb.properties.modified
    if saved is not None:
        saved = saved.replace(tzinfo=timezone.utc)
    elif r.headers.get("Last-Modified"):
        saved = parsedate_to_datetime(r.headers["Last-Modified"])
    changed = saved.astimezone(tz).strftime("%b %-d, %-I:%M %p") if saved else "unknown"
    for cell in ("E21", "F16", "F17"):
        if not isinstance(ws[cell].value, (int, float)):
            raise ValueError(f"Totals!{cell} is not a number")
    return {
        "overall": float(ws["E21"].value), "workload": float(ws["F16"].value),
        "control": float(ws["F17"].value),
        "w_done": ws["F10"].value, "c_done": ws["F11"].value,
        "w_need": ws["F13"].value, "c_need": ws["F14"].value,
        "hours": ws["E23"].value,
        "tests": [(ws[f"C{r}"].value, ws[f"D{r}"].value or 0) for r in range(5, 21)],
        "ctrl": [ws[f"F{r}"].value or 0 for r in range(5, 9)],
        "changed": changed,
        "fetched": datetime.now(tz).strftime("%-I:%M:%S %p"),
    }


def color(p):
    return "#e74c3c" if p < 0.33 else "#f1c40f" if p < 0.66 else "#2ecc71"


def bar(p, n):
    c = color(p)
    segs = "".join(
        f'<div class="seg" style="background:{c}"></div>' if p * n >= i + 0.5
        else '<div class="seg"></div>' for i in range(n))
    return f'<div class="bar">{segs}</div>'


@st.fragment(run_every=REFRESH_SECONDS)
def dashboard():
    error = False
    try:
        d = load()
        st.session_state["last"] = d
    except Exception:
        d = st.session_state.get("last")
        error = True
    if d is None:
        st.markdown('<div class="big">Can\'t reach file</div>', unsafe_allow_html=True)
        return

    total_done, total_need = d["w_done"] + d["c_done"], d["w_need"] + d["c_need"]
    per_test = d["w_need"] / len(d["tests"])
    per_ctrl = d["c_need"] / len(d["ctrl"])
    tests_done = sum(1 for _, n in d["tests"] if n >= per_test)
    ctrl_done = sum(1 for n in d["ctrl"] if n >= per_ctrl)

    h = '<div class="title">PILOT WORKLOAD MANEUVERS \u2014 DATA COLLECTION</div>'
    h += '<div class="label">TOTAL DATA COLLECTED</div>'
    h += f'<div class="big">{d["overall"]:.1%}</div>' + bar(d["overall"], 20)
    h += f'<div class="sub">{total_done:g} of {total_need:g} maneuvers complete</div>'

    cols = ""
    for name, key, done, need in (("WORKLOAD", "workload", d["w_done"], d["w_need"]),
                                  ("CONTROL", "control", d["c_done"], d["c_need"])):
        cols += (f'<div><div class="label">{name}</div><div class="mid">{d[key]:.1%}</div>'
                 f'{bar(d[key], 10)}<div class="sub">{done:g} of {need:g} complete</div></div>')
    h += f'<div style="display:grid;grid-template-columns:1fr 1fr;gap:2rem">{cols}</div>'

    tiles = [("HOURS COLLECTED", f'{d["hours"]:.2f} hrs'),
             ("MANEUVERS REMAINING", f"{total_need - total_done:g}"),
             ("WORKLOAD TESTS DONE", f'{tests_done} of {len(d["tests"])}'),
             ("CONTROL SETS DONE", f'{ctrl_done} of {len(d["ctrl"])}')]
    h += '<div class="tiles">' + "".join(
        f'<div class="tile"><div class="k">{k}</div><div class="v">{v}</div></div>'
        for k, v in tiles) + '</div>'

    h += '<div class="label">PROGRESS BY TEST</div><div class="tests">'
    for num, n in d["tests"]:
        cls = "done" if n >= per_test else "part" if n > 0 else ""
        h += f'<div class="test {cls}">Test {num}: {n:g} / {per_test:g}</div>'
    h += '</div>'

    if error:
        h += f'<div class="err">Connection lost \u2014 showing data from {d["changed"]}</div>'
    else:
        h += (f'<div class="status">Data last changed {d["changed"]} '
              f'&nbsp;\u00b7&nbsp; checked {d["fetched"]}</div>')
    st.markdown(h, unsafe_allow_html=True)


dashboard()
