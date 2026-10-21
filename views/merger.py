"""Merger - UI page (display only). Logic lives in tools/merger.py.

Flow: upload + order files -> output details -> Merge -> result + download."""
from __future__ import annotations

import datetime

import pandas as pd
import streamlit as st

from tools import merger as mg
from utils import ui
from utils.files import MAX_SHEET_MB, check_size

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
TODAY = datetime.date.today()


# ---------------------------------------------------------------- state
st.session_state.setdefault("merger_up_ver", 0)      # bump -> fresh (empty) uploader
st.session_state.setdefault("merger_order", [])      # upload file_ids, in merge order


def _move(fid: str, step: int) -> None:
    """Swap a file with its neighbour in the merge order (button callback)."""
    order = st.session_state.merger_order
    i = order.index(fid)
    j = i + step
    if 0 <= j < len(order):
        order[i], order[j] = order[j], order[i]


def _clear() -> None:
    """Remove every uploaded file and any previous result (button callback)."""
    st.session_state.merger_up_ver += 1
    st.session_state.merger_order = []
    st.session_state.pop("merger_result", None)


# ---------------------------------------------------------------- header
ui.page_header(
    "Merger",
    "Combine per-distribution Excel files into one clean workbook: filters removed, "
    "links broken, only each distribution's own rows kept, uniform formatting.",
    icon="layers", eyebrow="Tool")

# ---------------------------------------------------------------- 1. files
ui.section_header("Distribution files", "The file name is the Dist. Code, e.g. LESCO.xlsx.", step=1)
with ui.card("merger-files"):
    uploads = st.file_uploader(
        "Select distribution Excel files", type=["xlsx", "xlsm"], accept_multiple_files=True,
        key=f"merger_upload_{st.session_state.merger_up_ver}",
        help=f"One workbook per distribution (.xlsx / .xlsm, up to {MAX_SHEET_MB} MB each). "
             "The file name must contain the Dist. Code.")

    valid = {}
    for f in uploads or []:
        problem = check_size(f, MAX_SHEET_MB)
        if problem:
            ui.show_error(problem)
        else:
            valid[f.file_id] = f

    # keep the user's order for files that are still present, append new ones
    kept = [i for i in st.session_state.merger_order if i in valid]
    st.session_state.merger_order = kept + [i for i in valid if i not in kept]
    order: list[str] = st.session_state.merger_order

    if order:
        names = [valid[i].name for i in order]
        if len(set(n.lower() for n in names)) < len(names):
            st.warning("Two files share the same name, so they point to the same Dist. Code "
                       "and its rows would be added twice.", icon=":material/warning:")
        st.caption("Merge order, top first. The first file sets the layout and formatting.")
        for pos, fid in enumerate(order):
            c_no, c_name, c_up, c_dn = st.columns([0.5, 6, 1.4, 1.4], vertical_alignment="center")
            c_no.markdown(f"**{pos + 1}.**")
            c_name.markdown(f"{valid[fid].name}  \n<small>{valid[fid].size / 1024:,.0f} KB</small>",
                            unsafe_allow_html=True)
            c_up.button("Up", key=f"merger_up_{fid}", icon=":material/arrow_upward:",
                        disabled=pos == 0, on_click=_move, args=(fid, -1), width="stretch")
            c_dn.button("Down", key=f"merger_dn_{fid}", icon=":material/arrow_downward:",
                        disabled=pos == len(order) - 1, on_click=_move, args=(fid, 1),
                        width="stretch")
        st.button("Clear all", key="merger_clear", icon=":material/delete:", on_click=_clear,
                  help="Remove every file from the list")

if not order:
    ui.empty_state("No files yet", "Upload at least one distribution workbook to begin.")
    st.stop()

# ---------------------------------------------------------------- 2. details
ui.section_header("Output details", "These form the output file name.", step=2)
with ui.card("merger-details"):
    c1, c2 = st.columns(2)
    unit = c1.text_input("Unit", key="merger_unit", placeholder="e.g. Lahore")
    nature = c2.text_input("Nature", key="merger_nature", placeholder="e.g. Sales")
    c3, c4 = st.columns(2)
    month = c3.selectbox("Month", mg.MONTHS, index=TODAY.month - 1, key="merger_month")
    year = c4.number_input("Year", min_value=2000, max_value=2100, value=TODAY.year, step=1,
                           key="merger_year")
    sheet = st.text_input("Sheet name", key="merger_sheet",
                          placeholder="Blank = first sheet of every file",
                          help="Leave blank to use the first sheet of every file.")
    with st.expander("Advanced (change only if needed)"):
        a1, a2 = st.columns(2)
        dist_col = a1.text_input("Dist. Code header", value="Dist. Code", key="merger_dist_col",
                                 help="Header text of the column that holds the Dist. Code.")
        hdr_row = a2.text_input("Header row (blank = auto)", key="merger_hdr_row",
                                placeholder="Auto",
                                help="Row number of the header. Leave blank to detect it.")
    out_label = mg.clean_filename_part(f"{unit} {nature} {month} {int(year)}") + ".xlsx"
    st.caption(f"Output file name: **{out_label}**")

signature = (tuple(order), unit.strip(), nature.strip(), month, int(year), sheet.strip(),
             dist_col.strip(), hdr_row.strip())

# ---------------------------------------------------------------- 3. action
ui.section_header("Merge", "Combines the files in the order listed above.", step=3)
with ui.card("merger-run"):
    run = st.button("Merge files", key="merger_run", type="primary",
                    icon=":material/merge:", width="stretch")

if run:
    problem = None
    if not unit.strip() or not nature.strip():
        problem = "Please enter both Unit and Nature."
    elif hdr_row.strip() and not hdr_row.strip().isdigit():
        problem = "Header row must be a number, or blank."
    if problem:
        st.warning(problem, icon=":material/warning:")
    else:
        lines: list[str] = []
        bar = st.progress(0.0, text="Starting...")

        def _progress(i: int, n: int) -> None:
            bar.progress(i / n, text=f"Processed {i} of {n} file(s)")

        result: dict = {"error": None, "log": lines}
        try:
            with st.spinner("Merging files..."):
                name, data, summary = mg.merge_bytes(
                    [(valid[i].name, valid[i].getvalue()) for i in order],
                    unit.strip(), nature.strip(), month, str(int(year)),
                    sheet_name=sheet.strip() or None,
                    dist_col=dist_col.strip() or "Dist. Code",
                    header_row=int(hdr_row) if hdr_row.strip() else None,
                    log=lines.append, progress=_progress)
            result.update(name=name, data=data, summary=summary)
        except mg.MergeError as e:
            lines.append(f"ERROR: {e}")
            result["error"] = str(e)
        except Exception as e:                      # unexpected: traceback to server log only
            ui.log.exception("Merger failed", exc_info=e)
            lines.append(f"UNEXPECTED ERROR: {e!r}")
            result["error"] = "Something went wrong while merging. Please check the files and try again."
        finally:
            bar.empty()
        st.session_state.merger_result = (signature, result)

# ---------------------------------------------------------------- 4. results
saved = st.session_state.get("merger_result")
if saved:
    saved_sig, res = saved
    ui.section_header("Result", step=4)
    if saved_sig != signature:
        st.info("Inputs changed since this result was created. Click Merge files to refresh it.",
                icon=":material/info:")
    if res["error"]:
        ui.show_error(res["error"])
    else:
        summary = res["summary"]
        merged = [s for s in summary if s[3] > 0]
        skipped = len(summary) - len(merged)
        ui.show_success(f"Merge complete: {res['name']}")
        m1, m2, m3 = st.columns(3)
        with m1:
            ui.metric_card("Files merged", str(len(merged)), f"of {len(summary)} uploaded")
        with m2:
            ui.metric_card("Total data rows", f"{sum(s[3] for s in summary):,}")
        with m3:
            ui.metric_card("Files skipped", str(skipped), "no rows for their Dist. Code")
        with ui.card("merger-summary"):
            st.dataframe(pd.DataFrame(
                [(f, d, t, k, "Merged" if k else "Skipped") for f, d, t, k in summary],
                columns=["File", "Dist. Code", "Rows in sheet", "Rows kept", "Status"]),
                hide_index=True, width="stretch")
        st.download_button("Download merged workbook", data=res["data"], file_name=res["name"],
                           mime=XLSX_MIME, key="merger_download", type="primary",
                           icon=":material/download:", width="stretch")
    with st.expander("Processing log"):
        st.code("\n".join(res["log"]) or "(empty)", language=None)
