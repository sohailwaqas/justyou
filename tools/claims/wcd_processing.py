"""
wcd_processing.py  -  Puts Primary Lifting (sales file), Variable Disc. W.S (Wholesale Claim workbook) and
Committed Discount (HTH + Card Discount PDFs) together, one line per distributor and pack, and collects everything
the user should check.  No Excel formatting here (see wcd_writer.py).

Rules (v0.7):
  * Primary Lifting = the pack's column(s) of the sales file (several columns are added, e.g. 2000 ML 4-B + 6-B).
  * Variable Disc. W.S = raw cases of that pack on the distributor's sheet of the Wholesale Claim workbook.
  * Committed Discount = Card Discount cases + HTH cases of that package (the printed "Sales(CS)" figures).
    The cell is BLUE when both claims have cases for the pack (it is written as =card+hth).
  * "% age of W.S Claim" = W.S / Primary Lifting; above 50% it is RED (done in wcd_writer.py).
  * A distributor is matched by its CODE. Nothing is guessed: a missing figure stays EMPTY and is listed on the Check sheet.
  * Carry Forward Stock is left 0 (filled by hand).
"""
import calendar
import os
import re
from dataclasses import dataclass, field

from tools.claims import config
from tools.claims import processing
from tools.claims import wcd_reader
from tools.claims import wcd_writer
from tools.claims.readers import FriendlyError
from tools.claims.wcd_reader import pack_key


@dataclass
class PackFigures:
    primary: float
    ws: float = None          # None = no Wholesale Claim figure (EMPTY cell)
    card: float = None        # None = the distributor has no Card Discount page
    hth: float = None         # None = the distributor has no HTH page

    def committed(self):
        got = [v for v in (self.card, self.hth) if v is not None]
        return sum(got) if got else None

    def is_sum_of_two(self):
        """Both claims hold cases for this pack -> the cell is shown as =card+hth on a blue background."""
        return bool(self.card and self.hth and self.card > 0 and self.hth > 0)


@dataclass
class WcdRow:
    code: str
    name: str
    packs: list = field(default_factory=list)      # [PackFigures] in the order of config.WCD_PACKS


def _key(name):
    return re.sub(r"[^A-Z0-9]", "", name.upper())


def _mon(m):
    return f"{processing.MONTHS[m[0] - 1]}-{m[1]}"


def _fmt(v):
    return "-" if v is None else f"{v:,.0f}"


def build_rows(selected, sales, ws_map, pages, year, month):
    """selected = [Distributor]; ws_map = {code: WsSheet} or None (no file); pages = [PackPage] or None (no PDFs).
    Returns (rows, issues, run_notes). issues = [(distributor, code, type, details)]; run_notes = [(type, details)]."""
    issues, run_notes = [], []
    want = (month, year)
    mapped_ws = {pack_key(n) for _, _, names, _ in config.WCD_PACKS for n in names}
    mapped_pdf = {pack_key(n) for _, _, _, names in config.WCD_PACKS for n in names}

    # ---------- sales file: columns used for Primary Lifting ----------
    for _, cols, _, _ in config.WCD_PACKS:
        for c in cols:
            if c not in sales.pack_columns:
                run_notes.append(("Sales column missing",
                                  f"Pack column '{c}' is not in the sales file; it is counted as 0 for every distributor."))
    run_notes += [("Sales file note", n) for n in sales.notes if n.startswith("Code ")]

    # ---------- PDF pages: one per (kind, code); repeated pages are reported ----------
    by = {"CARD": {}, "HTH": {}}
    label = {"CARD": "Card Discount", "HTH": "HTH"}
    seen_names = set()
    if pages is None:
        run_notes.append(("PDFs not provided", "No HTH / Card Discount PDF was added, so Committed Discount is EMPTY."))
    else:
        for p in pages:
            lst = by[p.kind].setdefault(p.code, [])
            lst.append(p)
        ticked = {d.code for d in selected}
        for kind, d in by.items():
            for code, lst in d.items():
                if code not in ticked:                      # pages of distributors not on the report: one run note below
                    continue
                first = lst[0]
                for dup in lst[1:]:
                    same = abs((dup.total_cases or 0) - (first.total_cases or 0)) < 0.5
                    if same:
                        issues.append((first.name, code, f"{label[kind]} page repeated",
                                       f"The {label[kind]} page of this distributor was found twice ({first.file} p.{first.page} "
                                       f"and {dup.file} p.{dup.page}) with the same cases. It is counted ONCE."))
                    else:
                        issues.append((first.name, code, f"WARNING {label[kind]} pages differ",
                                       f"Two {label[kind]} pages for this distributor have different cases "
                                       f"({first.total_cases:,.0f} in {first.file} p.{first.page} and {dup.total_cases:,.0f} in "
                                       f"{dup.file} p.{dup.page}). The FIRST one is used. Please check."))
                for e in lst:
                    if e.month and e.month != want:
                        issues.append((e.name, code, f"WARNING {label[kind]} month",
                                       f"This {label[kind]} page is for {_mon(e.month)} but the report month is "
                                       f"{processing.MONTHS[month - 1]}-{year} ({e.file} p.{e.page})."))
                    elif not e.month:
                        issues.append((e.name, code, f"{label[kind]} month not found",
                                       f"No 'Month:' was found on this page, so the month could not be checked ({e.file})."))
                if first.total_cases is not None and abs(first.lines_cases - first.total_cases) > 0.5:
                    issues.append((first.name, code, f"{label[kind]} lines differ from total",
                                   f"The package lines add up to {first.lines_cases:,.0f} cases but the printed Total is "
                                   f"{first.total_cases:,.0f}. The package lines are used."))
                for k, (nm, cs) in first.packs.items():
                    if k not in mapped_pdf:
                        seen_names.add(nm)
        if seen_names:
            run_notes.append(("PDF packages not on this report",
                              "These package names in the PDFs are not one of the report's packs and were left out "
                              "(as designed): " + ", ".join(sorted(seen_names)) + ". If one of them should feed a pack, "
                              "add its name to WCD_PACKS in config.py."))

    # ---------- Wholesale Claim sheets ----------
    if ws_map is None:
        run_notes.append(("Wholesale Claim not provided",
                          "No Wholesale Claim file was chosen, so Variable Disc. W.S is EMPTY."))

    rows = []
    for d in selected:
        row = WcdRow(d.code, d.name)
        ws = ws_map.get(d.code) if ws_map is not None else None
        card = by["CARD"].get(d.code, [None])[0]
        hth = by["HTH"].get(d.code, [None])[0]

        if ws_map is not None:
            if ws is None:
                issues.append((d.name, d.code, "No Wholesale Claim sheet",
                               "No sheet with this code was found in the Wholesale Claim file; Variable Disc. W.S is left EMPTY."))
            else:
                if ws.month and ws.month != want:
                    issues.append((d.name, d.code, "WARNING Wholesale month",
                                   f"The Wholesale Claim sheet '{ws.sheet}' is for {_mon(ws.month)} but the report month is "
                                   f"{processing.MONTHS[month - 1]}-{year}."))
                elif not ws.month:
                    issues.append((d.name, d.code, "Wholesale month not found",
                                   f"No month was found on the sheet '{ws.sheet}', so it could not be checked."))
                for k, (lab, cs) in ws.packs.items():
                    if k not in mapped_ws and abs(cs) > 1e-12:
                        issues.append((d.name, d.code, "WARNING Wholesale pack not on report",
                                       f"The Wholesale Claim sheet has {cs:,.0f} raw cases of '{lab}', which is not one of the "
                                       "report's packs, so they are NOT in the report. (Is the name in WCD_PACKS right?)"))
        if pages is not None and card is None and hth is None:
            issues.append((d.name, d.code, "No HTH / Card Discount page",
                           "This distributor has no HTH and no Card Discount page in the PDFs; Committed Discount is left EMPTY."))
        elif pages is not None and (card is None or hth is None):
            issues.append((d.name, d.code, f"No {'HTH' if hth is None else 'Card Discount'} page",
                           f"This distributor has no {'HTH' if hth is None else 'Card Discount'} page in the PDFs; "
                           "only the other claim is counted."))

        # names written differently between the files
        names = {"Sales file": d.name}
        if ws:
            names["Wholesale Claim"] = ws.name
        for tag, p in (("Card Discount", card), ("HTH", hth)):
            if p:
                names[tag] = p.name
        distinct = {}
        for s, n in names.items():
            distinct.setdefault(_key(n), []).append(f"{s}: '{n}'")
        if len(distinct) > 1:
            issues.append((d.name, d.code, "Name differs",
                           "The name is written differently in the files (" + "; ".join(sum(distinct.values(), [])) +
                           f"). Matched by code {d.code}; '{d.name}' is used. Please make sure the code is right."))

        for heading, cols, ws_names, pdf_names in config.WCD_PACKS:
            for c in cols:
                if c in d.blank_cols:
                    issues.append((d.name, d.code, "Blank cell", f"'{c}' was empty in the sales file and is counted as 0."))
            for c, text in d.bad_cells:
                if c in cols:
                    issues.append((d.name, d.code, "Not a number", f"'{c}' contains '{text}'; counted as 0. Please check."))
            fig = PackFigures(sum(d.packs.get(c, 0.0) for c in cols))
            if ws is not None:
                fig.ws = sum(ws.packs.get(pack_key(n), [None, 0.0])[1] for n in ws_names)
            for attr, pg in (("card", card), ("hth", hth)):
                if pg is not None:
                    setattr(fig, attr, sum(pg.packs.get(pack_key(n), [None, 0.0])[1] for n in pdf_names))
            row.packs.append(fig)
            if fig.ws is not None and fig.primary > 0 and fig.ws / fig.primary > config.WCD_RED_LIMIT:
                issues.append((d.name, d.code, "W.S above limit",
                               f"{heading}: Variable Disc. W.S {fig.ws:,.0f} is {fig.ws / fig.primary:.1%} of Primary Lifting "
                               f"{fig.primary:,.0f} (limit {config.WCD_RED_LIMIT:.0%}); the cell is RED."))
            if fig.primary == 0 and ((fig.ws or 0) > 0 or (fig.committed() or 0) > 0):
                issues.append((d.name, d.code, "WARNING Primary Lifting is 0",
                               f"{heading}: Primary Lifting is 0 but there are claim cases (W.S {_fmt(fig.ws)}, "
                               f"Committed {_fmt(fig.committed())}); the % cells show **."))
        rows.append(row)

    # ---------- claims that belong to nobody on the report ----------
    chosen = {d.code for d in selected}
    other_ws = sorted(c for c in (ws_map or {}) if c not in chosen)
    if other_ws:
        run_notes.append(("Wholesale sheets not on the report",
                          "These Wholesale Claim sheets belong to distributors you did not tick, so they are NOT in the report: "
                          + ", ".join(f"{ws_map[c].name} ({c})" for c in other_ws)))
    other_pdf = sorted({p.code: p.name for p in (pages or []) if p.code not in chosen}.items())
    if other_pdf:
        run_notes.append(("PDF pages not on the report",
                          "These HTH / Card Discount pages belong to distributors you did not tick, so they are NOT in the report: "
                          + ", ".join(f"{n} ({c})" for c, n in other_pdf)))
    return rows, issues, run_notes


def make_file_name(unit, year, month):
    safe = " ".join(re.sub(r'[<>:"/\\|?*]', " ", unit).split())
    return f"{safe} {config.WCD_FILE_WORD} {calendar.month_name[month]}-{year}.xlsx"


def generate_wcd_report(selected, sales, ws_path, pdf_paths, year, month, unit, region, output_folder=None, sales_name=""):
    """Reads the Wholesale Claim workbook and the PDFs, builds the report workbook and returns (saved path, summary lines).
    selected = [Distributor] (any order; the report is sorted by distributor code)."""
    if not selected:
        raise FriendlyError("No distributor was ticked.")
    output_folder = output_folder or config.OUTPUT_FOLDER
    os.makedirs(output_folder, exist_ok=True)

    ws_map, notes, pages, files = None, [], None, []
    if ws_path:
        ws_map, n = wcd_reader.read_claim_book(ws_path)
        notes += n
    if pdf_paths:
        pages, n, files = wcd_reader.read_pdf_packs(pdf_paths)
        notes += n
    selected = sorted(selected, key=lambda d: (len(d.code), d.code))
    rows, issues, run_notes = build_rows(selected, sales, ws_map, pages, year, month)
    run_notes = notes + run_notes
    sources = [f"Sales file: {sales_name or '(not named)'}   (sheet '{sales.sheet_used}')",
               "Wholesale Claim file: " + (os.path.basename(ws_path) if ws_path else "(not chosen)")]
    sources += [f"PDF: {n} - {s}" for n, s in files]
    if not pdf_paths:
        sources.append("PDF files: (none added)")

    path = processing.unique_path(output_folder, make_file_name(unit, year, month))
    try:
        wcd_writer.write_wcd_workbook(rows, issues, run_notes, sources, unit, region, year, month, path)
    except PermissionError:
        raise FriendlyError(f"Could not save '{path}'. Close it in Excel and try again.")

    warn = [(t, d) for t, d in run_notes if t.startswith("WARNING")] + \
           [(f"{n} ({c})", f"[{t}] {d}") for n, c, t, d in issues if t.startswith("WARNING")]
    lines = [f"Saved: {path}", f"Distributors (rows): {len(rows)}"]
    if warn:
        lines.insert(0, f"WARNING: {len(warn)} warning(s) below (in red). Please read them before using the report.")
    for r in rows:
        parts = []
        for (heading, *_), f in zip(config.WCD_PACKS, r.packs):
            pct = "**" if not f.primary or f.ws is None else f"{f.ws / f.primary:.1%}"
            parts.append(f"{heading}: Primary {_fmt(f.primary)}, W.S {_fmt(f.ws)} ({pct}), Committed {_fmt(f.committed())}"
                         + (" [card+HTH]" if f.is_sum_of_two() else ""))
        lines.append(f"{r.code} {r.name}: " + " | ".join(parts))
    for t, d in run_notes:
        lines.append(f"[Whole run] {t}: {d}")
    for n, c, t, d in issues:
        lines.append(f"[{t}] {n} ({c}): {d}")
    return path, lines
