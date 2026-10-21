"""NKA Filler page (UI only). Logic lives in tools/nka_filler.py."""
from __future__ import annotations

import os

import pandas as pd
import streamlit as st

from tools import nka_filler as nka
from utils import ui
from utils.files import MAX_SHEET_MB, check_size, file_id

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
XLSM_MIME = "application/vnd.ms-excel.sheet.macroEnabled.12"

ui.page_header(
    "NKA Filler",
    "Upload the NKA sales data and the NKA file. Each product's cases are placed "
    "into the key-account column(s) you choose, and you download the filled file.",
    icon="table",
    eyebrow="Tool",
)

# ---------------------------------------------------------------- 1. Upload
ui.section_header("Upload your files", "Both files stay in this session and are not stored.", step=1)
with ui.card("nka-upload"):
    c1, c2 = st.columns(2)
    with c1:
        sale_up = st.file_uploader(
            "Sale file", type=["xlsx", "xlsm"], key="nka_sale_upload",
            help="Needs columns: Brand NAME, PKG_NAME, Sales_Raw Cases. "
                 "A Month / Year / Date column is optional.",
        )
    with c2:
        nka_up = st.file_uploader(
            "NKA Sale file", type=["xlsx", "xlsm"], key="nka_nka_upload",
            help="Needs BRAND in cell B6 and the key-account names in row 6.",
        )

if not sale_up or not nka_up:
    ui.empty_state(
        "Add both files to begin",
        "Choose the Sale file and the NKA Sale file (.xlsx or .xlsm, up to "
        f"{MAX_SHEET_MB} MB each).",
        icon="upload",
    )
    st.stop()

for up in (sale_up, nka_up):
    problem = check_size(up, MAX_SHEET_MB)
    if problem:
        ui.show_error(problem)
        st.stop()

sale_bytes, nka_bytes = sale_up.getvalue(), nka_up.getvalue()
sale_fid, nka_fid = file_id(sale_bytes), file_id(nka_bytes)

# ---------------------------------------------------------------- 2. Options
prev_m, prev_y = nka.previous_month()
ui.section_header("Choose period and key accounts", "Defaults to last month.", step=2)
with ui.card("nka-options"):
    m_col, y_col = st.columns(2)
    with m_col:
        month_name = st.selectbox(
            "Month", nka.MONTH_NAMES, index=prev_m - 1, key="nka_month",
            help="Used to pick the NKA tab and to filter the sale rows when the sale file has a period column.",
        )
    with y_col:
        year = int(st.number_input("Year", min_value=2000, max_value=2100, value=prev_y,
                                   step=1, key="nka_year"))
    month = nka.MONTH_NAMES.index(month_name) + 1

    try:
        info_key = (nka_fid, month, year)
        if st.session_state.get("nka_info_key") != info_key:
            st.session_state["nka_info"] = nka.inspect_nka_bytes(nka_bytes, month, year)
            st.session_state["nka_info_key"] = info_key
        info = st.session_state["nka_info"]
    except ValueError as e:
        st.session_state.pop("nka_info_key", None)
        ui.show_error(str(e))
        st.stop()
    except Exception as e:  # noqa: BLE001 - never show raw tracebacks
        st.session_state.pop("nka_info_key", None)
        ui.show_error("The NKA file could not be read. Please check it is a valid Excel file.", e)
        st.stop()

    default_acc = [a for a in info["accounts"] if nka.norm(a) == nka.CFG["channel"]] or info["accounts"][:1]
    accounts = st.multiselect(
        "Key accounts to fill", info["accounts"], default=default_acc,
        key=f"nka_accounts_{nka_fid}_{info['tab']}",
        help="Only the ticked columns are changed; everything else in the NKA file is left as it is.",
    )
    st.caption(f"NKA tab detected: **{info['tab']}**")
    if info["note"]:
        st.info(info["note"])

# ---------------------------------------------------------------- 3. Action
ui.section_header("Fill the NKA file", step=3)
signature = (sale_fid, nka_fid, month, year, tuple(accounts))
ext = os.path.splitext(nka_up.name)[1]
with ui.card("nka-action"):
    run = st.button("Fill NKA file", type="primary", icon=":material/play_arrow:",
                    disabled=not accounts, key="nka_run",
                    help="Tick at least one key account first." if not accounts else None)
    if not accounts:
        st.caption("Tick at least one key account to continue.")

if run:
    try:
        with st.spinner("Matching products and filling the NKA file..."):
            result, out_bytes = nka.fill_bytes(sale_bytes, nka_bytes, accounts, month, year, ext)
        out_name = os.path.basename(nka.default_output_path(nka_up.name, month, year))
        st.session_state["nka_result"] = (signature, result, out_bytes, out_name)
    except ValueError as e:
        st.session_state.pop("nka_result", None)
        ui.show_error(str(e))
    except Exception as e:  # noqa: BLE001
        st.session_state.pop("nka_result", None)
        ui.show_error("Something went wrong while filling the file. Please check both files and try again.", e)

# ---------------------------------------------------------------- 4. Results
saved = st.session_state.get("nka_result")
if not saved:
    st.stop()
sig, r, out_bytes, out_name = saved
if sig != signature:
    st.info("Files or options changed since the last run. Press **Fill NKA file** to refresh the result.")
    st.stop()

ui.section_header("Result", step=4)
ok, message = nka.summary_message(r, out_name)
message = message.replace(" next to the NKA file", "").replace("see the red lines below", "see the \"Not placed\" table below")
if ok:
    ui.show_success(message)
else:
    st.warning(message)

m1, m2, m3, m4 = st.columns(4)
with m1:
    ui.metric_card("Cases in sale file", nka.fmt_num(r["sale_total"]))
with m2:
    ui.metric_card("Cases placed", nka.fmt_num(r["placed_total"]), ", ".join(r["accounts"]))
with m3:
    ui.metric_card("Unmatched lines", str(len(r["unmatched"])))
with m4:
    ui.metric_card("Cells cleared", str(r["cleared"]), "existing values reset")

for note in r["notes"]:
    st.info(note)
if r["bad"]:
    st.warning("Sale rows with a non-numeric quantity were ignored: row "
               + ", ".join(map(str, r["bad"])))

if r["unmatched"]:
    ui.section_header("Not placed", "These sale lines have no matching NKA row. Fix them in the NKA file or add a rule.")
    with ui.card("nka-unmatched"):
        st.dataframe(
            pd.DataFrame([{"Sale brand": u["sb"], "Sale pack": u["sp"],
                           "Cases": nka.fmt_num(u["qty"]), "Reason": u["why"]} for u in r["unmatched"]]),
            hide_index=True, width="stretch",
        )

ui.section_header("Placed lines", f"Sale tab: {r['sale_tab']}  |  NKA tab: {r['nka_tab']}")
with ui.card("nka-placed"):
    rows = [{
        "Sale brand": l["sb"], "Sale pack": l["sp"], "Cases": nka.fmt_num(l["qty"]),
        "NKA row": l["row"]["row"], "Placed in NKA as": f'{l["row"]["brand"]} / {l["row"]["pack"]}',
        "Note": l["note"] or "",
    } for l in r["lines"]]
    if rows:
        st.dataframe(pd.DataFrame(rows), hide_index=True, width="stretch")
    else:
        st.caption("No lines were placed.")

ui.section_header("Download")
with ui.card("nka-download"):
    st.download_button(
        "Download filled NKA file", data=out_bytes, file_name=out_name,
        mime=XLSM_MIME if out_name.lower().endswith(".xlsm") else XLSX_MIME,
        type="primary", icon=":material/download:", key="nka_download",
    )
    st.caption("Excel recalculates all totals when you open the file.")
