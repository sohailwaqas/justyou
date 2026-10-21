"""
ds_processing.py  -  Puts the PDF amounts (HTH, Card Discount) and the Freight / Wholesale report amounts together,
one line per distributor, and collects everything the user should check.  No Excel formatting here (see ds_writer.py).

Rules (v0.5):
  * Card Discount amount = the printed "Total:" amount on the distributor's Card Discount page.
  * HTH amount           = the printed "Total:" amount on the distributor's HTH page.
  * Freight / Wholesale  = computed from the workbooks this tool created earlier (Cases x Rate).
  * A distributor is matched by its CODE (names can differ in spelling or wrap over two lines).
  * Nothing is guessed: a missing figure stays EMPTY and is listed on the Check sheet.
"""
import calendar
import os
import re
from dataclasses import dataclass, field

from tools.claims import config
from tools.claims import ds_reader
from tools.claims import ds_writer
from tools.claims import processing
from tools.claims.readers import FriendlyError


@dataclass
class DsRow:
    code: str
    name: str
    card: float = None
    hth: float = None
    freight: float = None
    wholesale: float = None

    def total(self):
        return sum(v for v in (self.card, self.hth, self.freight, self.wholesale) if v is not None)


def _key(name):
    return re.sub(r"[^A-Z0-9]", "", name.upper())


def _fmt(v):
    return "-" if v is None else f"{v:,.0f}"


def _mon(m):
    return f"{processing.MONTHS[m[0] - 1]}-{m[1]}"


def build_summary(pdf, freight, wholesale, year, month, freight_name="", wholesale_name=""):
    """pdf = PdfResult; freight / wholesale = list of SheetAmount (or None when no file was chosen).
    Returns (rows, issues, run_notes).  issues = [(distributor, code, type, details)]."""
    issues, run_notes = [], []
    want = (month, year)

    # ---------- PDF entries: one per (kind, code); duplicates are reported ----------
    by_kind = {"CARD": {}, "HTH": {}}
    for e in pdf.entries:
        slot = by_kind[e.kind].setdefault(e.code, [])
        slot.append(e)
    label = {"CARD": "Card Discount", "HTH": "HTH"}
    names = {}                                   # code -> {source: name}
    figures = {"CARD": {}, "HTH": {}}
    for kind, d in by_kind.items():
        for code, lst in d.items():
            first = lst[0]
            figures[kind][code] = first.amount
            names.setdefault(code, {})[label[kind]] = first.name
            for dup in lst[1:]:
                if abs(dup.amount - first.amount) < 0.5:
                    issues.append((first.name, code, f"{label[kind]} page repeated",
                                   f"The {label[kind]} page of this distributor was found twice ({first.file} p.{first.page} "
                                   f"and {dup.file} p.{dup.page}) with the same amount. It is counted ONCE."))
                else:
                    issues.append((first.name, code, f"WARNING {label[kind]} amounts differ",
                                   f"Two {label[kind]} pages for this distributor have different amounts: "
                                   f"{first.amount:,.0f} ({first.file} p.{first.page}) and {dup.amount:,.0f} "
                                   f"({dup.file} p.{dup.page}). The FIRST one ({first.amount:,.0f}) is used. Please check."))
            for e in lst:
                if e.month and e.month != want:
                    issues.append((e.name, code, f"WARNING {label[kind]} month",
                                   f"This {label[kind]} page is for {_mon(e.month)} but the report month is "
                                   f"{processing.MONTHS[month - 1]}-{year} ({e.file} p.{e.page})."))
                elif not e.month:
                    issues.append((e.name, code, f"{label[kind]} month not found",
                                   f"No 'Month:' was found on this page, so the month could not be checked ({e.file})."))
            if abs(first.lines_sum - first.amount) > 0.5:
                issues.append((first.name, code, f"{label[kind]} lines differ from total",
                               f"The {first.lines_count} package lines add up to {first.lines_sum:,.0f} but the printed "
                               f"Total is {first.amount:,.0f}. The printed Total is used (difference "
                               f"{first.amount - first.lines_sum:+,.0f})."))
            if kind == "CARD" and first.discount_amount is not None and abs(first.discount_amount - first.amount) > 0.5:
                issues.append((first.name, code, "WARNING Card amounts differ",
                               f"'Discount Amount' ({first.discount_amount:,.0f}) is not the same as the 'Total:' "
                               f"({first.amount:,.0f}) on this Card Discount page. The 'Total:' is used. Please check."))
    for fname, printed in pdf.grand_total.items():
        mine = sum(e.amount for e in pdf.entries if e.file == fname and e.kind == "CARD")
        if abs(mine - printed) > 0.5:
            run_notes.append(("Card Discount grand total",
                              f"{fname}: the distributor totals add up to {mine:,.0f} but the printed Grand Total is "
                              f"{printed:,.0f} (difference {mine - printed:+,.0f}). The distributor totals are used."))

    # ---------- Freight / Wholesale sheets ----------
    sheet_fig = {"Freight": {}, "Wholesale": {}}
    for kind, items, fname in (("Freight", freight, freight_name), ("Wholesale", wholesale, wholesale_name)):
        if items is None:
            continue
        for it in items:
            if it.code in sheet_fig[kind]:
                issues.append((it.name, it.code, f"WARNING {kind} sheet repeated",
                               f"Two {kind} sheets have the code {it.code} ('{sheet_fig[kind][it.code][1]}' and "
                               f"'{it.sheet}'). The first is used. Please check."))
                continue
            sheet_fig[kind][it.code] = (it.amount, it.sheet)
            names.setdefault(it.code, {})[kind] = it.name or it.sheet
            if it.month and it.month != want:
                issues.append((it.name or it.sheet, it.code, f"WARNING {kind} month",
                               f"The {kind} sheet '{it.sheet}' is for {_mon(it.month)} but the report month is "
                               f"{processing.MONTHS[month - 1]}-{year}."))
            elif not it.month:
                issues.append((it.name or it.sheet, it.code, f"{kind} month not found",
                               f"No month was found on the sheet '{it.sheet}', so it could not be checked."))

    # ---------- one row per distributor ----------
    codes = set(names)
    rows = []
    for code in codes:
        src = names[code]
        nm = src.get("Card Discount") or src.get("HTH") or src.get("Freight") or src.get("Wholesale") or code
        row = DsRow(code, nm,
                    figures["CARD"].get(code), figures["HTH"].get(code),
                    sheet_fig["Freight"].get(code, (None,))[0], sheet_fig["Wholesale"].get(code, (None,))[0])
        rows.append(row)
        distinct = {}
        for s, n in src.items():
            distinct.setdefault(_key(n), []).append(f"{s}: '{n}'")
        if len(distinct) > 1:
            issues.append((nm, code, "Name differs",
                           "The name is written differently in the files (" + "; ".join(sum(distinct.values(), [])) +
                           f"). Matched by code {code}; '{nm}' is used. Please make sure the code is right."))
    if config.DS_SORT == "name":
        rows.sort(key=lambda r: r.name.lower())
    else:
        rows.sort(key=lambda r: (len(r.code), r.code))

    # ---------- missing figures (only for the sources that were provided) ----------
    sources = [("Card Discount", bool(by_kind["CARD"]), "card", "No Card Discount PDF page was provided"),
               ("HTH", bool(by_kind["HTH"]), "hth", "No HTH PDF page was provided"),
               ("Freight", freight is not None, "freight", "No Freight Charges file was chosen"),
               ("Wholesale", wholesale is not None, "wholesale", "No Wholesale Claim file was chosen")]
    for title, given, attr, why in sources:
        if not given:
            run_notes.append((f"{title} not provided", f"{why}, so the {title} column is EMPTY for every distributor."))
            continue
        for r in rows:
            if getattr(r, attr) is None:
                issues.append((r.name, r.code, f"{title} missing",
                               f"No {title} figure was found for this distributor; the cell is left EMPTY."))
    return rows, issues, run_notes


def make_file_name(unit, year, month):
    safe = " ".join(re.sub(r'[<>:"/\\|?*]', " ", unit).split())
    return f"{safe} {config.DS_FILE_WORD} {calendar.month_name[month]}-{year}.xlsx"


def generate_discount_summary(pdf_paths, freight_path, wholesale_path, year, month, unit, output_folder=None):
    """Reads everything, builds the summary workbook and returns (saved path, list of summary lines)."""
    output_folder = output_folder or config.OUTPUT_FOLDER
    os.makedirs(output_folder, exist_ok=True)
    pdf = ds_reader.read_pdfs(pdf_paths)
    book_notes = []
    freight = wholesale = None
    if freight_path:
        freight, n = ds_reader.read_freight_book(freight_path)
        book_notes += n
    if wholesale_path:
        wholesale, n = ds_reader.read_wholesale_book(wholesale_path)
        book_notes += n
    rows, issues, run_notes = build_summary(pdf, freight, wholesale, year, month,
                                            os.path.basename(freight_path or ""), os.path.basename(wholesale_path or ""))
    run_notes = list(pdf.notes) + list(book_notes) + run_notes
    sources = [f"PDF: {n} - {s}" for n, s in pdf.files]
    sources.append("Freight Charges file: " + (os.path.basename(freight_path) if freight_path else "(not chosen)"))
    sources.append("Wholesale Claim file: " + (os.path.basename(wholesale_path) if wholesale_path else "(not chosen)"))

    path = processing.unique_path(output_folder, make_file_name(unit, year, month))
    try:
        ds_writer.write_summary_workbook(rows, issues, run_notes, sources, unit, year, month, path)
    except PermissionError:
        raise FriendlyError(f"Could not save '{path}'. Close it in Excel and try again.")

    warn = [(t, d) for t, d in run_notes if t.startswith("WARNING")] + \
           [(f"{n} ({c})", f"[{t}] {d}") for n, c, t, d in issues if t.startswith("WARNING")]
    lines = [f"Saved: {path}", f"Distributors (rows): {len(rows)}"]
    if warn:
        lines.insert(0, f"WARNING: {len(warn)} warning(s) below (in red). Please read them before using the summary.")
    for r in rows:
        lines.append(f"{r.code} {r.name}: Card {_fmt(r.card)} | HTH {_fmt(r.hth)} | Freight {_fmt(r.freight)} | "
                     f"Wholesale {_fmt(r.wholesale)} | Total {r.total():,.0f}")
    lines.append(f"ALL: Card {sum(r.card or 0 for r in rows):,.0f} | HTH {sum(r.hth or 0 for r in rows):,.0f} | "
                 f"Freight {sum(r.freight or 0 for r in rows):,.0f} | Wholesale {sum(r.wholesale or 0 for r in rows):,.0f} | "
                 f"Total {sum(r.total() for r in rows):,.0f}")
    for t, d in run_notes:
        lines.append(f"[Whole run] {t}: {d}")
    for n, c, t, d in issues:
        lines.append(f"[{t}] {n} ({c}): {d}")
    return path, lines
