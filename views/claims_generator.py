"""Claims Generator - UI page (display only). Logic lives in tools/claims/ (entry point: tools/claims/service.py).

Five reports in workflow order, each in its own tab. Shared inputs (Month, Year, Unit, sales file, wholesale data
file, HTH / Card Discount PDFs) are entered once at the top. A report made in one tab is kept in session_state and
picked up automatically by the next one (this replaces the desktop tool's `output` folder); a file uploaded inside
a tab overrides it. Session keys are all prefixed `claims_`."""
from __future__ import annotations

import datetime
import re

import pandas as pd
import streamlit as st

from tools.claims import config as cfg
from tools.claims import processing as proc
from tools.claims import service as sv
from tools.claims import ws_processing
from tools.claims.readers import FriendlyError
from utils import ui
from utils.files import MAX_PDF_MB, MAX_SHEET_MB, check_size, file_id

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
ss = st.session_state
TODAY = datetime.date.today()
CHECK = ":material/check_circle:"
# ---------------------------------------------------------------- small helpers
def _clean(text) -> str:
    return " ".join(str(text or "").split())


def _warn(message: str) -> None:
    st.warning(message, icon=":material/warning:")


def _fid(file) -> str | None:
    """Content hash of a (name, bytes) file, or None."""
    return file_id(file[1]) if file else None


def _number(text):
    """A number typed by the user (commas allowed), or None. Zero is allowed."""
    try:
        v = float(str(text).strip().replace(",", ""))
        return v if v >= 0 else None
    except ValueError:
        return None


def _load_upload(up, loader, key: str, max_mb: int, what: str):
    """Parse an uploaded file once (again only when its content changes). Returns the parsed object or None."""
    rec_key = f"claims_{key}_rec"
    if up is None:
        ss.pop(rec_key, None)
        return None
    problem = check_size(up, max_mb)
    if problem:
        ui.show_error(problem)
        return None
    data = up.getvalue()
    fid = file_id(data)
    rec = ss.get(rec_key)
    if rec is None or rec[0] != fid:
        obj, err = None, None
        try:
            with st.spinner(f"Reading the {what} …"):
                obj = loader(data, up.name)
        except FriendlyError as e:
            err = str(e)
        except Exception as e:                                   # unexpected: traceback to the server log only
            ui.log.exception("Claims Generator: could not read the %s", what, exc_info=e)
            err = f"Sorry, the {what} could not be read. Please check that it is the right file and try again."
        rec = ss[rec_key] = (fid, obj, err)
    if rec[2]:
        ui.show_error(rec[2])
        return None
    return rec[1]


def _pick(kind: str, label: str, override, year, month, unit: str | None = None):
    """Which file feeds a report. An uploaded file wins; otherwise the report of that kind made earlier in this
    session (only when it is for the same month and year, and the same unit when `unit` is given).
    Returns ((file name, bytes) or None, note)."""
    if override is not None:
        problem = check_size(override, MAX_SHEET_MB)
        if problem:
            return None, problem
        return (override.name, override.getvalue()), f"Using the file you uploaded: **{override.name}**"
    g = ss.get(f"claims_gen_{kind}")
    if g is None:
        return None, f"No {label} file yet. Make it in its own tab first, or upload one here."
    if month is None or g.year != year or g.month != month:
        return None, (f"The {label} file made in this session is for {proc.date_label(g.year, g.month)}, "
                      "which is not the chosen month, so it is not used.")
    if unit is not None:
        safe = lambda u: " ".join(re.sub(r'[<>:"/\\|?*]', " ", u).split()).lower()
        if safe(unit) and safe(g.unit) != safe(unit):
            return None, (f"The {label} file made in this session is for unit '{g.unit}', not '{unit}', "
                          "so it is not used.")
    return (g.name, g.data), f"Using the {label} file made in this session: **{g.name}**"


def _need_period(month, unit_required: bool, unit: str):
    """Month / Unit checks shared by every report. Returns a message or None."""
    if month is None:
        return "Please choose the Month (step 1 above)."
    if unit_required and not unit:
        return "Please type the Unit name (step 1 above)."
    return None


def _confirm_missing(key: str, missing: list, consequence: str) -> bool:
    """The desktop tool asked 'Continue anyway?'. Here: a tick box that has to be ticked first."""
    if not missing:
        return True
    _warn(f"No {' and no '.join(missing)} is chosen, so {consequence} will be EMPTY.")
    return st.checkbox("Continue anyway", key=key)


def _result(tab: str, sig, noun: str) -> None:
    """Show the saved result of a tab: status, download, and the notes the desktop pop-up used to show."""
    rec = ss.get(f"claims_res_{tab}")
    if not rec:
        return
    saved_sig, out = rec
    ui.section_header("Result")
    if saved_sig != sig:
        st.info("Inputs changed since this file was created. Press the button again to refresh it.",
                icon=":material/info:")
    warn = [l for l in out.lines if "WARNING" in l]
    ui.show_success(f"{noun} created: {out.name}")
    if warn:
        _warn(f"{len(warn)} warning(s). Please read them before using the file.")
    st.download_button(f"Download {noun.lower()}", out.data, out.name, XLSX_MIME, key=f"claims_dl_{tab}",
                       type="primary", icon=":material/download:", width="stretch")
    st.caption("These notes are also saved on the 'Check' sheet inside the workbook.")
    if warn:
        with st.expander(f"Warnings ({len(warn)})", expanded=True):
            st.code("\n\n".join(warn), language=None, wrap_lines=True)
    with st.expander("All notes and details"):
        st.code("\n\n".join(out.lines), language=None, wrap_lines=True)


def _remember(tab: str, kind: str | None, sig, out, year: int, month: int, unit: str) -> None:
    """Keep the finished report for display and (for freight / wholesale / summary) for the next report."""
    ss[f"claims_res_{tab}"] = (sig, out)
    if kind:
        ss[f"claims_gen_{kind}"] = sv.GeneratedFile(out.name, out.data, year, month, unit)
    st.toast(f"{out.name} created", icon=CHECK)


def _run(fn, *args, **kwargs):
    """Run a report; show friendly errors. Returns the Output or None."""
    try:
        with st.spinner("Building the report …"):
            return fn(*args, **kwargs)
    except FriendlyError as e:
        ui.show_error(str(e))
    except Exception as e:
        ui.show_error("Something went wrong while building the report. Please check your files and inputs "
                      "and try again.", e)
    return None


def _table_height(n: int) -> int:
    return int(min(60 + 36 * max(n, 1), 520))


# ---------------------------------------------------------------- page header
with st.container(key="claims-header"):               # CSS hook for the subtle premium accent (assets/premium.css)
    ui.page_header("Claims Generator",
                   "Generate the monthly claims along with related summaries: Freight Charges, Wholesale Claim, "
                   "Discount Summary, Distribution Cost and Wholesale Claims Detail.",
                   icon="receipt", eyebrow="Premium tool")

# ---------------------------------------------------------------- 1. shared details
ui.section_header("Month, year and unit", "Used by every report below. Each report can use the file made by the "
                  "report before it.", step=1)
with ui.card("claims-period"):
    c1, c2, c3 = st.columns(3)
    month_name = c1.selectbox("Month", proc.MONTHS, index=None, placeholder="Choose month", key="claims_month")
    year = int(c2.number_input("Year", min_value=2000, max_value=2100, value=TODAY.year, step=1, key="claims_year"))
    unit = _clean(c3.text_input("Unit", key="claims_unit", placeholder="e.g. Muridke"))
month = proc.MONTHS.index(month_name) + 1 if month_name else None

# ---------------------------------------------------------------- 2. shared source files
ui.section_header("Source files", "Upload what you have; each report tells you what it still needs. "
                  "Files stay in this session and are not stored.", step=2)
with ui.card("claims-files"):
    f1, f2, f3 = st.columns(3, gap="medium")
    with f1:
        up_sales = st.file_uploader("Sales file", type=["xls", "xlsx", "csv"], key="claims_sales_upload",
                                    help="Used by Freight Charges, Distribution Cost and WS Claims Detail. "
                                         f"Up to {MAX_SHEET_MB} MB.")
    with f2:
        up_wsdata = st.file_uploader("Wholesale data file", type=["xls", "xlsx", "csv"], key="claims_wsdata_upload",
                                     help=f"Used by Wholesale Claim. Up to {MAX_SHEET_MB} MB.")
    with f3:
        up_pdfs = st.file_uploader("HTH and Card Discount PDFs", type="pdf", accept_multiple_files=True,
                                   key="claims_pdf_upload",
                                   help="Used by Discount Summary and WS Claims Detail. Many files at once, or one "
                                        f"combined PDF. Up to {MAX_PDF_MB} MB each.")

    sales = _load_upload(up_sales, sv.load_sales, "sales", MAX_SHEET_MB, "sales file")
    wsdata = _load_upload(up_wsdata, sv.load_wholesale, "wsdata", MAX_SHEET_MB, "wholesale data file")
    sales_file = ((up_sales.name, up_sales.getvalue())
                  if up_sales is not None and not check_size(up_sales, MAX_SHEET_MB) else None)

    pdf_files, pdf_months, pdf_rows = [], set(), []
    info = ss.setdefault("claims_pdf_info", {})
    for f in up_pdfs or []:
        problem = check_size(f, MAX_PDF_MB)
        if problem:
            ui.show_error(problem)
            continue
        data = f.getvalue()
        fid = file_id(data)
        if fid not in info:
            with st.spinner(f"Looking at {f.name} …"):
                info[fid] = sv.describe_pdf(f.name, data)
        label, months = info[fid]
        pdf_files.append((f.name, data))
        pdf_months |= months
        pdf_rows.append({"PDF file": f.name, "What was found": label})

    if sales is not None:
        st.caption(f"{CHECK} Sales file: {len(sales.distributors)} distributors from sheet '{sales.sheet_used}'.")
    if wsdata is not None:
        st.caption(f"{CHECK} Wholesale data: {len(wsdata.distributors)} distributors, {len(wsdata.rows)} data lines "
                   f"from sheet '{wsdata.sheet_used}'.")
    if pdf_rows:
        st.dataframe(pd.DataFrame(pdf_rows), hide_index=True, width="stretch")
        if pdf_months:
            names = ", ".join(f"{proc.MONTHS[m - 1]}-{y}" for m, y in sorted(pdf_months, key=lambda t: (t[1], t[0])))
            st.caption(f"The PDFs say: {names}. Make sure the Month and Year above are the same.")


# ================================================================ 1. FREIGHT CHARGES
def tab_freight() -> None:
    st.caption("One printable Freight Charges sheet per ticked distributor, from the sales file (cases x Freight Rate).")
    if sales is None:
        st.info("Upload the **Sales file** above to list the distributors.", icon=":material/upload:")
        return
    df = pd.DataFrame({
        "Tick": False,
        "Code": [d.code for d in sales.distributors],
        "Distributor": [d.name + ("   (no sales)" if sum(d.packs.get(c, 0.0) for _, srcs in cfg.PACKS for c in srcs) == 0
                                  else "") for d in sales.distributors],
        "Freight Rate": "", "Route": "", "Area Incharge": ""})
    with ui.card("claims-fr-table"):
        st.caption("Tick the distributors you need and fill their Freight Rate, Route and Area Incharge.")
        edited = st.data_editor(
            df, hide_index=True, width="stretch", height=_table_height(len(df)), key=f"claims_fr_editor_{file_id(up_sales.getvalue())}",
            disabled=["Code", "Distributor"],
            column_config={"Tick": st.column_config.CheckboxColumn("Tick", width="small"),
                           "Code": st.column_config.TextColumn("Code", width="small"),
                           "Distributor": st.column_config.TextColumn("Distributor", width="large"),
                           "Freight Rate": st.column_config.TextColumn("Freight Rate", width="small",
                                                                       help="A number above 0, e.g. 1.5"),
                           "Route": st.column_config.TextColumn("Route"),
                           "Area Incharge": st.column_config.TextColumn("Area Incharge")})
        go = st.button("Generate Freight Report", type="primary", key="claims_fr_run", icon=":material/bolt:")
    sig = (file_id(up_sales.getvalue()), month, year, unit, tuple(map(tuple, edited.itertuples(index=False))))
    if go:
        problem = _need_period(month, True, unit)
        selections = []
        if not problem:
            for i, r in edited.iterrows():
                if not r["Tick"]:
                    continue
                d = sales.distributors[i]
                try:
                    rate = float(str(r["Freight Rate"]).strip())
                    if rate <= 0:
                        raise ValueError
                except ValueError:
                    problem = f"Please enter a valid Freight Rate (a number above 0) for: {d.name}"
                    break
                route, incharge = _clean(r["Route"]), _clean(r["Area Incharge"])
                if not route or not incharge:
                    problem = f"Please fill in both Route and Area Incharge for: {d.name}"
                    break
                selections.append((d, rate, route, incharge))
            if not problem and not selections:
                problem = "Tick at least one distributor."
        if problem:
            _warn(problem)
        else:
            out = _run(sv.make_freight, sales, selections, year, month, unit, up_sales.name)
            if out:
                _remember("freight", "freight", sig, out, year, month, unit)
    _result("freight", sig, "Freight Charges report")


# ================================================================ 2. WHOLESALE CLAIM
def _ws_add() -> None:
    ss.claims_ws_ids.append(ss.claims_ws_next)
    ss.claims_ws_next += 1


def _ws_remove(cid: int) -> None:
    ss.claims_ws_ids.remove(cid)


def _ws_read_conditions(brands: list):
    """Read the condition lines. Returns (list of Condition, None) or (None, message)."""
    if not ss.claims_ws_ids:
        return None, "Add at least one condition with the Add condition button."
    conds = []
    for i, cid in enumerate(ss.claims_ws_ids, 1):
        pack = ss.get(f"claims_ws_pack_{cid}")
        if not pack:
            return None, f"Condition {i}: please choose the Pack."
        d1, d2 = ss.get(f"claims_ws_from_{cid}"), ss.get(f"claims_ws_to_{cid}")
        if d1 is None or d2 is None:
            return None, f"Condition {i}: please type the 'From day' and 'To day' as whole numbers, e.g. 1 and 15."
        rate = ss.get(f"claims_ws_rate_{cid}")
        if rate is None or rate <= 0:
            return None, f"Condition {i}: please type a Discount Rate (a number above 0)."
        chosen = [b for b in brands if b in (ss.get(f"claims_ws_brand_{cid}") or [])]
        if len(chosen) == len(brands):                         # every brand ticked = "All"
            chosen = []
        conds.append(ws_processing.Condition(pack, tuple(chosen), int(d1), int(d2), float(rate)))
    return conds, None


def tab_wholesale() -> None:
    st.caption("One Wholesale Claim sheet per ticked distributor. A sale is counted by its BOOKING date; "
               "sales delivered next month are flagged with a WARNING.")
    if wsdata is None:
        st.info("Upload the **Wholesale data file** above to set the conditions and list the distributors.",
                icon=":material/upload:")
        return
    ws_fid = file_id(up_wsdata.getvalue())
    if ss.get("claims_ws_fid") != ws_fid:                      # new data file: old conditions belong to the old packs / brands
        ss.claims_ws_fid = ws_fid
        ss.setdefault("claims_ws_next", 0)
        ss.claims_ws_ids = [ss.claims_ws_next]
        ss.claims_ws_next += 1
    packs, brands = sv.ws_pack_choices(wsdata), sv.ws_brand_choices(wsdata)

    with ui.card("claims-ws-conditions"):
        st.markdown("**Conditions** — a sale is counted by its BOOKING date. Days are days of the chosen month.")
        widths = [2.2, 3, 1.1, 1.1, 1.4, 0.6]
        head = st.columns(widths)
        for col, text in zip(head, ("Pack", "Brand (none = All)", "From day", "To day", "Discount Rate", "")):
            col.caption(text)
        for cid in list(ss.claims_ws_ids):
            c = st.columns(widths, vertical_alignment="center")
            c[0].selectbox("Pack", packs, index=None, placeholder="Choose pack", key=f"claims_ws_pack_{cid}",
                           label_visibility="collapsed")
            c[1].multiselect("Brand", brands, key=f"claims_ws_brand_{cid}", format_func=sv.brand_label,
                             placeholder="All", label_visibility="collapsed")
            c[2].number_input("From day", min_value=1, max_value=31, value=None, step=1, placeholder="1",
                              key=f"claims_ws_from_{cid}", label_visibility="collapsed")
            c[3].number_input("To day", min_value=1, max_value=31, value=None, step=1, placeholder="15",
                              key=f"claims_ws_to_{cid}", label_visibility="collapsed")
            c[4].number_input("Discount Rate", min_value=0.0, value=None, step=0.5, format="%g", placeholder="Rs",
                              key=f"claims_ws_rate_{cid}", label_visibility="collapsed")
            c[5].button("✕", key=f"claims_ws_rm_{cid}", on_click=_ws_remove, args=(cid,), help="Remove this condition")
        st.button("Add condition", key="claims_ws_add", icon=":material/add:", on_click=_ws_add)

        conds, cond_problem = _ws_read_conditions(brands)
        overlaps, overlap_ok = [], True
        if conds and month is not None:
            try:
                ws_processing.validate_conditions(conds, year, month)
                overlaps = ws_processing.find_overlaps(conds, month)
            except FriendlyError:
                pass                                           # shown when the button is pressed
        if overlaps:
            _warn("These conditions overlap, so the same cases would be counted twice:\n\n" + "\n\n".join(overlaps))
            overlap_ok = st.checkbox("Continue anyway", key="claims_ws_overlap_ok")

    df = pd.DataFrame({"Tick": False, "Code": [d.code for d in wsdata.distributors],
                       "Distributor": [d.name + ("   (no sales)" if d.total_cases == 0 else "") for d in wsdata.distributors],
                       "Unit": [d.unit for d in wsdata.distributors], "SDM/O/S": [d.sdm for d in wsdata.distributors]})
    with ui.card("claims-ws-table"):
        st.caption("Tick the distributors. Unit and SDM/O/S are filled from the data; you can change them.")
        edited = st.data_editor(
            df, hide_index=True, width="stretch", height=_table_height(len(df)), key=f"claims_ws_editor_{ws_fid}",
            disabled=["Code", "Distributor"],
            column_config={"Tick": st.column_config.CheckboxColumn("Tick", width="small"),
                           "Code": st.column_config.TextColumn("Code", width="small"),
                           "Distributor": st.column_config.TextColumn("Distributor", width="large"),
                           "Unit": st.column_config.TextColumn("Unit"),
                           "SDM/O/S": st.column_config.TextColumn("SDM/O/S")})
        go = st.button("Generate Wholesale Claim", type="primary", key="claims_ws_run", icon=":material/bolt:")
    sig = (ws_fid, month, year, tuple(conds or ()), tuple(map(tuple, edited.itertuples(index=False))))
    if go:
        problem = _need_period(month, False, unit)
        if not problem and cond_problem:
            problem = cond_problem
        if not problem:
            try:
                ws_processing.validate_conditions(conds, year, month)
            except FriendlyError as e:
                problem = str(e)
        if not problem and overlaps and not overlap_ok:
            problem = "Some conditions overlap. Tick 'Continue anyway' to count the same cases in both rows."
        selections = []
        if not problem:
            for i, r in edited.iterrows():
                if not r["Tick"]:
                    continue
                d = wsdata.distributors[i]
                u, s = _clean(r["Unit"]), _clean(r["SDM/O/S"])
                if not u or not s:
                    problem = f"Please fill in both Unit and SDM/O/S for: {d.name}"
                    break
                selections.append((d, u, s))
            if not problem and not selections:
                problem = "Tick at least one distributor."
        if problem:
            _warn(problem)
        else:
            out = _run(sv.make_wholesale, wsdata, selections, conds, year, month, up_wsdata.name)
            if out:
                _remember("wholesale", "wholesale", sig, out, year, month, unit)
    _result("wholesale", sig, "Wholesale Claim")


# ================================================================ 3. DISCOUNT SUMMARY
def tab_summary() -> None:
    st.caption("One row per distributor: Card Discount and HTH (from the PDFs) plus Freight and Wholesale "
               "(from the two reports made in tabs 1 and 2). Non GTM, Periodic, Promo and Special Approval are "
               "left empty to fill in by hand.")
    with ui.card("claims-ds-inputs"):
        if pdf_files:
            st.caption(f"{CHECK} {len(pdf_files)} PDF file(s) added above.")
        else:
            st.info("Add the **HTH and Card Discount PDFs** above.", icon=":material/upload:")
        st.markdown("**Reports made earlier** (Freight and Wholesale amounts come from here)")
        o1, o2 = st.columns(2, gap="medium")
        with o1:
            up_fr = st.file_uploader("Use a different Freight Charges file (optional)", type="xlsx", key="claims_ds_fr_up")
        with o2:
            up_wh = st.file_uploader("Use a different Wholesale Claim file (optional)", type="xlsx", key="claims_ds_wh_up")
        freight, fr_note = _pick("freight", "Freight Charges", up_fr, year, month)
        wholesale, wh_note = _pick("wholesale", "Wholesale Claim", up_wh, year, month)
        st.caption(f"Freight Charges: {fr_note}")
        st.caption(f"Wholesale Claim: {wh_note}")
        missing = [n for n, f in (("Freight Charges file", freight), ("Wholesale Claim file", wholesale)) if f is None]
        ok = _confirm_missing("claims_ds_missing_ok", missing, "that column")
        go = st.button("Generate Discount Summary", type="primary", key="claims_ds_run", icon=":material/bolt:")
    sig = (month, year, unit, tuple(file_id(d) for _, d in pdf_files), _fid(freight), _fid(wholesale))
    if go:
        problem = _need_period(month, True, unit)
        if not problem and not pdf_files:
            problem = "Add the HTH and Card Discount PDF files first."
        if not problem and not ok:
            problem = "Tick 'Continue anyway' to build the summary without the missing file(s)."
        if problem:
            _warn(problem)
        else:
            out = _run(sv.make_discount_summary, pdf_files, freight, wholesale, year, month, unit)
            if out:
                _remember("summary", "summary", sig, out, year, month, unit)
    _result("summary", sig, "Discount Summary")


# ================================================================ 4. DISTRIBUTION COST
def tab_cost() -> None:
    st.caption("The Discount Summary table plus Raw Cases and Per Unit Cost for every distributor, and the unit's "
               "Per Case Cost (sampling included). Raw Cases come from the sales file.")
    with ui.card("claims-co-inputs"):
        st.info("The claim amounts come from the **Discount Summary** file. If you fill any columns there by hand "
                "(Non GTM, Periodic, Promo, Special Approval), save your copy in Excel and upload it below. "
                "Otherwise the copy made in tab 3 is used, with those columns still empty.", icon=":material/info:")
        if sales_file is None:
            st.info("Upload the **Sales file** above (the Raw Cases come from it).", icon=":material/upload:")
        up_sm = st.file_uploader("Use your own saved Discount Summary file (optional)", type="xlsx", key="claims_co_sm_up")
        summary, sm_note = _pick("summary", "Discount Summary", up_sm, year, month, unit)
        st.caption(f"Discount Summary: {sm_note}")
        k1, k2 = st.columns(2)
        cases_txt = k1.text_input("Sampling Cases", key="claims_co_cases", placeholder="a number, 0 or more")
        price_txt = k2.text_input("Sampling Rate (Rs per case)", key="claims_co_price", placeholder="a number, 0 or more")
        go = st.button("Generate Cost Sheet", type="primary", key="claims_co_run", icon=":material/bolt:")
    sig = (month, year, unit, _fid(sales_file), _fid(summary), cases_txt.strip(), price_txt.strip())
    if go:
        problem = _need_period(month, True, unit)
        if not problem and sales_file is None:
            problem = "Upload the sales file first (the Raw Cases come from it)."
        if not problem and summary is None:
            problem = "Choose the Discount Summary file first (make it in tab 3, or upload it)."
        cases = price = None
        if not problem:
            cases = _number(cases_txt)
            if cases is None:
                problem = "Please type the Sampling Cases (a number, 0 or more)."
        if not problem:
            price = _number(price_txt)
            if price is None:
                problem = "Please type the Sampling Rate (a number, 0 or more)."
        if problem:
            _warn(problem)
        else:
            out = _run(sv.make_cost, summary, sales_file, cases, price, year, month, unit)
            if out:
                _remember("cost", None, sig, out, year, month, unit)
    _result("cost", sig, "Cost sheet")


# ================================================================ 5. WS CLAIMS DETAIL
def tab_wcd() -> None:
    st.caption("Per distributor and pack (1500 / 1000 / 2000 ML): Primary Lifting, Variable Disc. W.S and Committed "
               "Discount (Card + HTH). BLUE = Committed is the sum of both claims; RED = W.S above 50% of Primary Lifting.")
    if sales is None:
        st.info("Upload the **Sales file** above (Primary Lifting comes from it).", icon=":material/upload:")
        return
    with ui.card("claims-wcd-inputs"):
        region = _clean(st.text_input("Region", key="claims_wcd_region", placeholder="e.g. Central"))
        up_ws = st.file_uploader("Use a different Wholesale Claim file (optional)", type="xlsx", key="claims_wcd_ws_up")
        ws_book, ws_note = _pick("wholesale", "Wholesale Claim", up_ws, year, month)
        st.caption(f"Wholesale Claim: {ws_note}")
        if pdf_files:
            st.caption(f"{CHECK} {len(pdf_files)} HTH / Card Discount PDF file(s) added above.")
        ok_ws = _confirm_missing("claims_wcd_ws_ok", ["Wholesale Claim file"] if ws_book is None else [],
                                 "Variable Disc. W.S")
        ok_pdf = _confirm_missing("claims_wcd_pdf_ok", [] if pdf_files else ["HTH / Card Discount PDF"],
                                  "Committed Discount")
    df = pd.DataFrame({"Tick": False, "Code": [d.code for d in sales.distributors],
                       "Distributor": [d.name for d in sales.distributors]})
    with ui.card("claims-wcd-table"):
        st.caption("Tick the distributors for the report (all boxes start empty).")
        edited = st.data_editor(
            df, hide_index=True, width="stretch", height=_table_height(len(df)),
            key=f"claims_wcd_editor_{file_id(up_sales.getvalue())}", disabled=["Code", "Distributor"],
            column_config={"Tick": st.column_config.CheckboxColumn("Tick", width="small"),
                           "Code": st.column_config.TextColumn("Code", width="small"),
                           "Distributor": st.column_config.TextColumn("Distributor (from the sales file)", width="large")})
        go = st.button("Generate WS Claims Detail", type="primary", key="claims_wcd_run", icon=":material/bolt:")
    sig = (month, year, unit, region, file_id(up_sales.getvalue()), _fid(ws_book),
           tuple(file_id(d) for _, d in pdf_files), tuple(edited["Tick"]))
    if go:
        problem = _need_period(month, True, unit)
        if not problem and not region:
            problem = "Please type the Region."
        selected = [sales.distributors[i] for i, t in enumerate(edited["Tick"]) if t]
        if not problem and not selected:
            problem = "Tick at least one distributor."
        if not problem and not (ok_ws and ok_pdf):
            problem = "Tick 'Continue anyway' to build the report without the missing file(s)."
        if problem:
            _warn(problem)
        else:
            out = _run(sv.make_wcd, selected, sales, ws_book, pdf_files, year, month, unit, region, up_sales.name)
            if out:
                _remember("wcd", None, sig, out, year, month, unit)
    _result("wcd", sig, "WS Claims Detail report")


# ---------------------------------------------------------------- 3. the five reports
ui.section_header("Reports", "Work left to right: the Freight and Wholesale reports feed the Discount Summary, "
                  "which feeds the Cost sheet.", step=3)
t1, t2, t3, t4, t5 = st.tabs(["1 · Freight Charges", "2 · Wholesale Claim", "3 · Discount Summary",
                              "4 · Distribution Cost", "5 · WS Claims Detail"])
with t1:
    tab_freight()
with t2:
    tab_wholesale()
with t3:
    tab_summary()
with t4:
    tab_cost()
with t5:
    tab_wcd()

