"""
cost_processing.py  -  Puts the claim amounts (Discount Summary workbook) and the Raw Cases (sales file) together,
one line per distributor, adds the unit totals and collects everything the user should check.
No Excel formatting here (see cost_writer.py).

Rules (v0.6):
  * Claim amounts (columns D..K) = copied from the Discount Summary workbook made earlier by this tool,
    including the columns you filled in by hand there. Total Amount = D + E + ... + K.
  * Raw Cases = the "Grand Total Raw Cases Vol." column of the sales file, matched by distributor CODE.
  * Per Unit Cost (per distributor) = Total Amount / Raw Cases.
  * Unit block:  Sampling Amount = Sampling Cases x Sampling Price;  Total Amount = Sampling Amount + Claims Amount
                 (= Total of all distributors);  Per Case Cost = Total Amount / Raw Cases (= Raw Cases of all distributors).
  * Nothing is guessed: a missing figure stays EMPTY and is listed on the Check sheet.
"""
import calendar
import os
import re
from dataclasses import dataclass, field

from tools.claims import config
from tools.claims import cost_reader
from tools.claims import cost_writer
from tools.claims import processing
from tools.claims.readers import FriendlyError


@dataclass
class CostRow:
    code: str
    name: str
    amounts: list = field(default_factory=list)      # 8 values (columns D..K), None = empty
    raw: float = None                                # Raw Cases, None = not found

    def total(self):
        return sum(v for v in self.amounts if v is not None)

    def per_case(self):
        return self.total() / self.raw if self.raw else None


def _key(name):
    return re.sub(r"[^A-Z0-9]", "", name.upper())


def _mon(m):
    return f"{processing.MONTHS[m[0] - 1]}-{m[1]}"


def build_cost(summary_rows, summary_notes, summary_month, sales_map, sales_notes, year, month):
    """Returns (rows, issues, run_notes). issues = [(distributor, code, type, details)]; run_notes = [(type, details)]."""
    issues, run_notes = [], list(summary_notes) + list(sales_notes)
    if summary_month and summary_month != (month, year):
        run_notes.append(("WARNING Discount Summary month",
                          f"The Discount Summary file is for {_mon(summary_month)} but the report month is "
                          f"{processing.MONTHS[month - 1]}-{year}."))
    elif not summary_month:
        run_notes.append(("Discount Summary month not found",
                          "No 'Claim Month' date was found in the Discount Summary file, so its month could not be checked."))

    rows, seen = [], set()
    for sr in summary_rows:
        if sr.code in seen:
            issues.append((sr.name, sr.code, "WARNING Code repeated in summary",
                           "This code appears twice in the Discount Summary file. The FIRST row is used. Please check."))
            continue
        seen.add(sr.code)
        row = CostRow(sr.code, sr.name, list(sr.amounts))
        rows.append(row)
        if all(v is None for v in sr.amounts):
            issues.append((sr.name, sr.code, "No claim amounts", "This distributor has no amount in any claim column."))
        ent = sales_map.get(sr.code)
        if ent is None:
            issues.append((sr.name, sr.code, "WARNING Raw cases missing",
                           f"Code {sr.code} was not found in the sales file, so Raw Cases is EMPTY. Its claims are still "
                           "added to the unit total, so the unit Per Case Cost is too HIGH until this is fixed."))
            continue
        if ent.problem == "blank":
            issues.append((sr.name, sr.code, "WARNING Raw cases missing",
                           f"'{config.CO_COL_RAW_CASES}' is empty in the sales file; Raw Cases is left EMPTY."))
        elif ent.problem:
            issues.append((sr.name, sr.code, "WARNING Raw cases not a number",
                           f"'{config.CO_COL_RAW_CASES}' contains '{ent.problem[5:]}'; Raw Cases is left EMPTY. Please check."))
        else:
            row.raw = ent.raw
            if ent.raw == 0:
                issues.append((sr.name, sr.code, "Raw cases zero", "Raw Cases is 0, so Per Unit Cost shows ****."))
            if abs(ent.raw - ent.pack_sum) > 0.5:
                issues.append((sr.name, sr.code, "Raw cases differ from pack columns",
                               f"'{config.CO_COL_RAW_CASES}' is {ent.raw:,.1f} but the pack columns add up to "
                               f"{ent.pack_sum:,.1f}. The '{config.CO_COL_RAW_CASES}' value is used."))
        if _key(ent.name) != _key(sr.name):
            issues.append((sr.name, sr.code, "Name differs",
                           f"Summary says '{sr.name}', sales file says '{ent.name}'. Matched by code {sr.code}; "
                           f"'{sr.name}' is used. Please make sure the code is right."))

    others = {c: e for c, e in sales_map.items() if c not in seen}
    if others:
        cases = sum(e.raw or 0 for e in others.values())
        run_notes.append(("Distributors not in the summary",
                          f"{len(others)} distributor(s) of the sales file are not in the Discount Summary "
                          f"({cases:,.0f} raw cases). They are NOT included in the unit Raw Cases or cost."))
    return rows, issues, run_notes


def make_file_name(unit, year, month):
    safe = " ".join(re.sub(r'[<>:"/\\|?*]', " ", unit).split())
    return f"{safe} {config.CO_FILE_WORD} {calendar.month_name[month]}-{year}.xlsx"


def generate_cost_sheet(summary_path, sales_path, sampling_cases, sampling_price, year, month, unit, output_folder=None):
    """Reads both files, builds the cost workbook and returns (saved path, list of summary lines)."""
    output_folder = output_folder or config.OUTPUT_FOLDER
    os.makedirs(output_folder, exist_ok=True)
    summary_rows, summary_notes, summary_month = cost_reader.read_summary_book(summary_path)
    sales_map, sales_notes = cost_reader.read_raw_cases(sales_path)
    rows, issues, run_notes = build_cost(summary_rows, summary_notes, summary_month, sales_map, sales_notes, year, month)
    sources = ["Discount Summary file: " + os.path.basename(summary_path),
               "Sales file: " + os.path.basename(sales_path),
               f"Sampling Cases: {sampling_cases:g}   |   Sampling Price: {sampling_price:g}"]

    path = processing.unique_path(output_folder, make_file_name(unit, year, month))
    try:
        cost_writer.write_cost_workbook(rows, issues, run_notes, sources, unit, year, month,
                                        sampling_cases, sampling_price, path)
    except PermissionError:
        raise FriendlyError(f"Could not save '{path}'. Close it in Excel and try again.")

    claims = sum(r.total() for r in rows)
    raw_total = sum(r.raw or 0 for r in rows)
    sampling_amount = sampling_cases * sampling_price
    unit_cost = (sampling_amount + claims) / raw_total if raw_total else None
    warn = [(t, d) for t, d in run_notes if t.startswith("WARNING")] + \
           [(f"{n} ({c})", f"[{t}] {d}") for n, c, t, d in issues if t.startswith("WARNING")]
    lines = [f"Saved: {path}", f"Distributors (rows): {len(rows)}"]
    if warn:
        lines.insert(0, f"WARNING: {len(warn)} warning(s) below (in red). Please read them before using the cost sheet.")
    for r in rows:
        pc = r.per_case()
        lines.append(f"{r.code} {r.name}: Claims {r.total():,.0f} | Raw Cases {('-' if r.raw is None else f'{r.raw:,.0f}')} | "
                     f"Cost per case {('****' if pc is None else f'{pc:.4f}')}")
    lines.append(f"UNIT: Claims {claims:,.0f} + Sampling {sampling_amount:,.0f} = {claims + sampling_amount:,.0f} | "
                 f"Raw Cases {raw_total:,.0f} | Per Case Cost {('****' if unit_cost is None else f'{unit_cost:.4f}')}")
    for t, d in run_notes:
        lines.append(f"[Whole run] {t}: {d}")
    for n, c, t, d in issues:
        lines.append(f"[{t}] {n} ({c}): {d}")
    return path, lines
