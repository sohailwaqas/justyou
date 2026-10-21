"""Toll Calculator - PDF -> trips -> filters -> tick sets -> total (UI only).
Logic lives in tools/toll_core.py (engine) and tools/toll_service.py (helpers).
Session keys are all prefixed `toll_`."""
from __future__ import annotations

import calendar
import datetime as dt

import pandas as pd
import streamlit as st

from tools import da_logic as da
from tools import toll_core as core
from tools.toll_service import read_trips
from utils import ui
from utils.files import MAX_PDF_MB, check_size, file_id

CURRENCY = "Rs."
ss = st.session_state

# ---------------------------------------------------------------- page
ui.page_header("Toll Calculator",
               "Read a toll-transaction PDF and calculate the toll for any period. "
               "Scanned or 'Microsoft Print to PDF' files are read with OCR.",
               icon="route", eyebrow="Tool")

# ---- 1. upload --------------------------------------------------------------
ui.section_header("Upload the toll data", "A PDF of your toll transactions.", step=1)
with ui.card("toll-upload"):
    up = st.file_uploader("Toll data (PDF)", type="pdf", key="toll_upload",
                          help=f"Up to {MAX_PDF_MB} MB. Reading takes a few seconds per page.")
if up is None:
    ui.empty_state("No PDF yet", "Upload a toll-transaction PDF to begin.", icon="upload")
    st.stop()

too_big = check_size(up, MAX_PDF_MB)
if too_big:
    ui.show_error(too_big)
    st.stop()

data = up.getvalue()
fid = file_id(data)
if ss.get("toll_file_id") != fid:                     # OCR only once per file
    status = st.empty()
    try:
        with st.spinner("Reading the PDF with OCR …"):
            ss.toll_trips = read_trips(data, lambda m: status.info(m))
    except Exception as e:
        status.empty()
        ui.show_error("Could not read this PDF. Please check that it is a toll-transaction "
                      f"report and try again. ({type(e).__name__})", e)
        st.stop()
    status.empty()
    ss.toll_file_id = fid
    ss.toll_unticked, ss.toll_reset, ss.toll_result = set(), 0, None
    st.toast(f"Read {len(ss.toll_trips)} rows from {up.name}", icon=":material/check_circle:")
trips = ss.toll_trips

real = [t for t in trips if core.is_toll_trip(t)]
dated = [t["date"] for t in real if t["date"]]
if not dated:
    ui.show_error("No toll trips were found in this PDF.")
    st.stop()

bad = [t for t in trips if t["flag"]]
if bad:
    st.warning("Please double-check these trips in the PDF:\n\n" + "\n".join(
        f"- {core.describe_trip(t)} — {t['flag']}" for t in bad[:15]),
        icon=":material/warning:")

# ---- 2. filters -------------------------------------------------------------
ui.section_header("Month & exclusions",
                  "Choose the month, skip Sundays, and list leave days to leave out.", step=2)
latest = max(dated)
with ui.card("toll-filters"):
    c1, c2, c3 = st.columns([1.4, 1, 2.2], gap="large")
    month_name = c1.selectbox("Month", [calendar.month_name[i] for i in range(1, 13)],
                              index=latest.month - 1, key=f"toll_month_{fid}")
    year = c2.number_input("Year", min_value=2000, max_value=2100, value=latest.year, step=1,
                           key=f"toll_year_{fid}")
    leave_text = c3.text_input("Leave / excluded dates", key=f"toll_leaves_{fid}",
                               placeholder="e.g. 14, 20-22",
                               help="Day numbers of the chosen month, e.g. 3, 5-7, 14 "
                                    "(or full dates like 3-9-2026). These days are left out of the total.")
    no_sun = st.checkbox("Exclude Sundays", key="toll_no_sun")

month = list(calendar.month_name).index(month_name)
year = int(year)
lo = dt.date(year, month, 1)
hi = dt.date(year, month, calendar.monthrange(year, month)[1])

leave: set = set()
try:
    for a, b in da.parse_leaves(leave_text, month, year):
        leave |= core.date_span(a, b)
except (da.ToolError, ValueError) as e:
    ui.show_error(str(e))
    st.stop()

classified = core.classify_trips(trips, lo, hi, no_sun, leave)
groups = core.group_counted(classified)

# ---- 3. trip sets with tick boxes -----------------------------------------------
ui.section_header("Trip sets", "Tick the sets to include in the total.", step=3)
if "toll_unticked" not in ss:
    ss.toll_unticked, ss.toll_reset = set(), 0

b1, b2, _ = st.columns([1, 1, 4])
if b1.button("Select all", key="toll_sel_all", width="stretch"):
    ss.toll_unticked = set()
    ss.toll_reset += 1
if b2.button("Select none", key="toll_sel_none", width="stretch"):
    ss.toll_unticked = set(groups)
    ss.toll_reset += 1

if not groups:
    st.warning("No trips match the current filters.", icon=":material/warning:")
    st.stop()

df = pd.DataFrame([{"Select": k not in ss.toll_unticked,
                    "Entry Station": k[0], "Exit Station": k[1], "Fare": k[2],
                    "Frequency": n, "Amount": k[2] * n} for k, n in groups.items()])
sig = hash((lo, hi, no_sun, tuple(sorted(leave)), ss.toll_reset))
edited = st.data_editor(
    df, hide_index=True, width="stretch", key=f"toll_sets_{sig}",
    disabled=["Entry Station", "Exit Station", "Fare", "Frequency", "Amount"],
    column_config={"Select": st.column_config.CheckboxColumn("✔", width="small"),
                   "Fare": st.column_config.NumberColumn(format="%d"),
                   "Amount": st.column_config.NumberColumn(format="%d")})

# remember ticks across filter changes
keys = list(groups)
for k, ticked in zip(keys, edited["Select"]):
    (ss.toll_unticked.discard if ticked else ss.toll_unticked.add)(k)

chosen = edited[edited["Select"]]
total = int(chosen["Amount"].sum())
n_trips = int(chosen["Frequency"].sum())
state = (sig, tuple(sorted(ss.toll_unticked)))

# ---- 4. calculate -----------------------------------------------------------------
ui.section_header("Result", step=4)
if st.button("Calculate", type="primary", key="toll_calc", icon=":material/calculate:"):
    ss.toll_result = (state, total, len(chosen), n_trips)

res = ss.get("toll_result")
if res and res[0] == state:
    m1, m2 = st.columns([2, 3], gap="medium")
    with m1:
        ui.metric_card("Total amount", f"{CURRENCY} {res[1]:,}",
                       f"{res[2]} of {len(groups)} sets ticked · {res[3]} trips")
elif res:
    st.info("Filters or ticks changed – press **Calculate** again.", icon=":material/info:")
else:
    st.caption("Press **Calculate** to total the ticked sets.")

# ---- audit ------------------------------------------------------------------------
st.write("")
counts = pd.Series([s for _, s in classified]).value_counts()
st.caption("Trips in file: " + " · ".join(f"{n} {s.lower()}" for s, n in counts.items()))
with st.expander("Audit: every trip and why it was counted or left out", icon=":material/search:"):
    audit = pd.DataFrame([{
        "Date": core.fmt_date(t["date"]) if t["date"] else "",
        "Day": t["date"].strftime("%a") if t["date"] else "",
        "Time": t["time"], "Entry Station": t["entry"] or "—",
        "Exit Station": t["exit"] or "—", "Fare": t["fare"], "Status": s}
        for t, s in classified])
    only = st.multiselect("Show status", sorted(audit["Status"].unique()),
                          default=sorted(audit["Status"].unique()))
    view = audit[audit["Status"].isin(only)]
    st.dataframe(view, hide_index=True, width="stretch")
    st.download_button("Download audit as CSV", view.to_csv(index=False),
                       "toll_audit.csv", "text/csv", key="toll_audit_dl",
                       icon=":material/download:")
