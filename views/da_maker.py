"""DA Maker - creates DA bills for sales staff + the final summary (UI only).
Logic lives in tools/da_logic.py.  Session keys are all prefixed `da_`."""
from __future__ import annotations

import calendar
import json
from datetime import date

import pandas as pd
import streamlit as st

from tools import da_logic as da
from utils import ui
from utils.files import MAX_SHEET_MB, check_size, file_id

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# ---------------------------------------------------------------- state
ss = st.session_state
ss.setdefault("da_all_sun", da.DEFAULT_EXCLUDE_SUNDAYS)
ss.setdefault("da_all_fri", da.DEFAULT_EXCLUDE_FRIDAYS)
ss.setdefault("da_result", None)


def _apply_everyone(kind: str, flag_key: str) -> None:
    """'Everyone' tick boxes set that column for every person (personal leaves are kept)."""
    for i in range(len(ss.get("da_people", []))):
        ss[f"da_{kind}_{i}"] = bool(ss[flag_key])


def _cell_text(v) -> str:
    return "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v).strip()


# ---------------------------------------------------------------- page
ui.page_header("DA Maker",
               "Create Daily Allowance bills for your sales staff, plus one final summary "
               "sheet, in a single Excel workbook.",
               icon="calendar-check", eyebrow="Tool")

# ---- 1. data sheet --------------------------------------------------------
ui.section_header("Upload the data sheet",
                  "Columns: Code, Name, Designation, Department, any number of area columns, then DA.",
                  step=1)
with ui.card("da-upload"):
    up = st.file_uploader("DA data sheet", type=["xlsx", "xls", "ods", "csv"], key="da_upload",
                          help="Everything between 'Department' and 'DA' is treated as a visit area.")
    st.download_button("Download a blank data-sheet template (CSV)", da.TEMPLATE_CSV,
                       "da_data_sheet_template.csv", "text/csv", key="da_template",
                       help="Optional: shows the column layout the tool expects.")

if up is None:
    ui.empty_state("No data sheet yet",
                   "Upload an .xlsx, .xls, .ods or .csv file to list your staff and start.")
    st.stop()

too_big = check_size(up, MAX_SHEET_MB)
if too_big:
    ui.show_error(too_big)
    st.stop()

data = up.getvalue()
fid = file_id(data)
if ss.get("da_file_id") != fid:                       # (re)load only when the file changes
    try:
        with st.spinner("Reading the data sheet …"):
            people = da.load_data(data, up.name)
    except da.ToolError as e:
        ui.show_error(str(e))
        st.stop()
    except Exception as e:                            # unexpected
        ui.show_error("Sorry, this file could not be read. Please check it and try again.", e)
        st.stop()
    ss.da_file_id, ss.da_people = fid, people
    for i in range(len(people)):                      # fresh per-person settings for the new file
        ss[f"da_sun_{i}"] = bool(ss.da_all_sun)
        ss[f"da_fri_{i}"] = bool(ss.da_all_fri)
        ss[f"da_leave_{i}"] = ""
    ss.da_result = None
    st.toast(f"Loaded {len(people)} people from {up.name}", icon=":material/check_circle:")
people = ss.da_people

# ---- 2. unit details ----------------------------------------------------------
ui.section_header("Unit details", "These appear on the Summary sheet and in the file name.", step=2)
with ui.card("da-unit"):
    c1, c2 = st.columns(2)
    unit_code = c1.text_input("Unit Code", key="da_unit_code", placeholder="e.g. 1234")
    unit_name = c2.text_input("Unit Name", key="da_unit_name", placeholder="e.g. Muridke")
    c3, c4 = st.columns(2)
    month_name = c3.selectbox("Month", [calendar.month_name[i] for i in range(1, 13)],
                              index=date.today().month - 1, key="da_month")
    year = c4.number_input("Year", min_value=2000, max_value=2100, value=date.today().year,
                           step=1, key="da_year")

# ---- 3. exclusions & leaves ---------------------------------------------------
ui.section_header("Exclusions & leaves",
                  "Set per person before generating. Areas are picked at random for each visit day.",
                  step=3)
with ui.card("da-rules"):
    st.markdown("**Everyone**")
    e1, e2, e3 = st.columns([1, 1, 2])
    e1.checkbox("Exclude Sundays", key="da_all_sun", on_change=_apply_everyone,
                args=("sun", "da_all_sun"), help="Ticks or unticks Sundays for every person below.")
    e2.checkbox("Exclude Fridays", key="da_all_fri", on_change=_apply_everyone,
                args=("fri", "da_all_fri"), help="Ticks or unticks Fridays for every person below.")
    common_leaves = e3.text_input("Common leaves (e.g. public holidays)", key="da_common_leaves",
                                  placeholder="e.g. 14, 20-22",
                                  help="Added to every person's own leaves.")
    st.caption("Leaves: type day numbers of the chosen month, e.g. **3, 5-7, 14** "
               "(or full dates like 3-9-2026). Use each person's own cell for personal leaves.")
    # One row of ordinary widgets per person. (A grid cell editor detaches and floats away when the page
    # is scrolled mid-edit; a normal text box does not.)
    W = [3, 1.2, 1.2, 2.6]
    h = st.columns(W, vertical_alignment="center")
    for col, label in zip(h, ["Name", "Exclude Sundays", "Exclude Fridays", "Leaves"]):
        col.markdown(f"<span class='hy-th'>{label}</span>", unsafe_allow_html=True)
    for i, person in enumerate(people):
        r = st.columns(W, vertical_alignment="center")
        r[0].markdown(f"{i + 1}. {person['name']}")
        r[1].checkbox("Exclude Sundays", key=f"da_sun_{i}", label_visibility="collapsed")
        r[2].checkbox("Exclude Fridays", key=f"da_fri_{i}", label_visibility="collapsed")
        r[3].text_input("Leaves", key=f"da_leave_{i}", label_visibility="collapsed",
                        placeholder="e.g. 3, 5-7, 14")

# ---- 4. generate --------------------------------------------------------------
month = list(calendar.month_name).index(month_name)
person_settings = [dict(sun=bool(ss.get(f"da_sun_{i}", False)), fri=bool(ss.get(f"da_fri_{i}", False)),
                        leaves=", ".join(x for x in (_cell_text(common_leaves),
                                                     _cell_text(ss.get(f"da_leave_{i}"))) if x))
                   for i in range(len(people))]
signature = json.dumps([fid, unit_code.strip(), unit_name.strip(), month, int(year), person_settings])

st.write("")
if st.button("Generate DA sheets", type="primary", key="da_generate", icon=":material/bolt:"):
    if not unit_code.strip():
        ui.show_error("Please enter the Unit Code.")
    elif not unit_name.strip():
        ui.show_error("Please enter the Unit Name.")
    else:
        try:
            with st.spinner("Building DA sheets and summary …"):
                result = da.generate_bytes(people, unit_code.strip(), unit_name.strip(),
                                           month, int(year), person_settings)
            ss.da_result = (signature, result)
            st.toast("DA sheets and summary created", icon=":material/check_circle:")
        except da.ToolError as e:
            ui.show_error(str(e))
        except Exception as e:
            ui.show_error("Something went wrong while building the workbook. "
                          "Please check your inputs and try again.", e)

# ---- 5. results & download ----------------------------------------------------
res = ss.da_result
if res:
    sig, result = res
    ui.section_header("Your workbook", "One DA bill per person plus the summary sheet.", step=4)
    if sig != signature:
        st.info("Inputs changed since this workbook was built - press **Generate DA sheets** "
                "again to refresh it.", icon=":material/info:")
    rows = result.people
    m1, m2, m3 = st.columns(3, gap="medium")
    with m1:
        ui.metric_card("People", str(len(rows)), f"{len(rows)} DA sheets + 1 summary")
    with m2:
        ui.metric_card("Claimed days", f"{sum(r.days for r in rows):,}", "Across all bills")
    with m3:
        ui.metric_card("Total D.A", f"Rs. {sum(r.total for r in rows):,.0f}", "Before any T.A")
    st.write("")
    st.dataframe(pd.DataFrame([{"Name": r.name, "Sheet": r.sheet, "Days": r.days,
                                "Rate": r.rate, "D.A total": r.total} for r in rows]),
                 hide_index=True, width="stretch")
    st.caption(f"Summary sheet: **{result.summary_sheet}**. T.A is an input (blue cells) on the "
               "Summary sheet; D.A is linked to each person's sheet.")
    st.download_button("Download Excel workbook", result.data, result.filename, XLSX_MIME,
                       type="primary", key="da_download", icon=":material/download:")
