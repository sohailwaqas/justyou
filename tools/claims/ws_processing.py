"""
ws_processing.py  -  Turns wholesale data + the user's CONDITIONS into report data and checks the data.
No Excel formatting happens here (that is ws_writer.py).

A CONDITION = Pack + Brand(s) + Claim dates (day from - day to) + Discount rate.
A sale is counted in a condition when ALL of these are true:
   - its pack matches,  - its brand matches (or the condition says "All"),
   - its BOOKING date (ORDER_DATE) is inside the report month and between the condition's days.
Delivery date is NOT used to decide, but sales delivered after the end of the month get a WARNING.
"""
import calendar
import os
import re
from dataclasses import dataclass, field

import numpy as np
import pandas as pd

from tools.claims import config
from tools.claims import processing
from tools.claims import readers
from tools.claims import ws_writer
from tools.claims.readers import FriendlyError


@dataclass(frozen=True)
class Condition:
    pack: str            # pack name exactly as in the data, e.g. "1000 ML PET"
    brands: tuple        # brand names as in the data; empty tuple = "All"
    day_from: int
    day_to: int
    rate: float

    def pack_label(self):
        return config.WS_PACK_LABELS.get(self.pack.upper(), self.pack)

    def brand_text(self):
        if not self.brands:
            return "All"
        order = list(config.WS_BRAND_LABELS)
        chosen = sorted(self.brands, key=lambda b: (order.index(b.upper()) if b.upper() in order else len(order), b))
        return ", ".join(config.WS_BRAND_LABELS.get(b.upper(), b.title()) for b in chosen)

    def date_text(self, month):
        mon = processing.MONTHS[month - 1]
        if self.day_from == self.day_to:
            return f"{self.day_from:02d} {mon}"
        return f"{self.day_from:02d}-{self.day_to:02d} {mon}"


@dataclass
class WsOutlet:
    code: str
    name: str
    address: str
    cases: list = field(default_factory=list)      # cases for each condition (same order as the conditions)


@dataclass
class WsReport:
    code: str
    name: str
    unit: str
    sdm: str
    sheet_name: str
    outlets: list = field(default_factory=list)
    issues: list = field(default_factory=list)     # [(type, details)]; types starting with "WARNING" are highlighted
    total_cases: float = 0.0
    total_amount: float = 0.0
    late_lines: int = 0                            # claimed lines delivered after the month end
    late_cases: float = 0.0
    late_amount: float = 0.0


# ---------------------------------------------------------------- conditions
def order_conditions(conds):
    """Keep the user's order, but group conditions of the same pack together (first pack typed comes first)."""
    packs = []
    for c in conds:
        if c.pack.upper() not in packs:
            packs.append(c.pack.upper())
    return sorted(conds, key=lambda c: packs.index(c.pack.upper()))      # sort is stable


def validate_conditions(conds, year, month):
    """Raises FriendlyError with a clear message if a condition is not usable."""
    if not conds:
        raise FriendlyError("Add at least one condition.")
    last = calendar.monthrange(year, month)[1]
    for i, c in enumerate(conds, 1):
        if not c.pack:
            raise FriendlyError(f"Condition {i}: please choose the Pack.")
        if not (1 <= c.day_from <= last and 1 <= c.day_to <= last):
            raise FriendlyError(f"Condition {i}: days must be between 1 and {last} "
                                f"({processing.MONTHS[month - 1]}-{year}).")
        if c.day_from > c.day_to:
            raise FriendlyError(f"Condition {i}: 'From' day ({c.day_from}) is after 'To' day ({c.day_to}).")
        if c.rate <= 0:
            raise FriendlyError(f"Condition {i}: the Discount Rate must be above 0.")


def find_overlaps(conds, month):
    """Pairs of conditions that would count the same cases twice (same pack, shared days, shared brand)."""
    out = []
    for i, a in enumerate(conds):
        for b in conds[i + 1:]:
            if a.pack.upper() != b.pack.upper():
                continue
            if a.day_to < b.day_from or b.day_to < a.day_from:
                continue
            if a.brands and b.brands and not ({x.upper() for x in a.brands} & {x.upper() for x in b.brands}):
                continue
            out.append(f"{a.pack_label()}: [{a.brand_text()} | {a.date_text(month)} | Rs {a.rate:g}] overlaps "
                       f"[{b.brand_text()} | {b.date_text(month)} | Rs {b.rate:g}]. "
                       "The same cases are counted in both rows.")
    return out


# ---------------------------------------------------------------- helpers
def _fd(ts):
    return "(no date)" if pd.isna(ts) else f"{ts:%d-%b-%Y}"


def make_ws_sheet_name(dist_name, used):
    """Sheet name = distributor name (Excel allows 31 characters, and not  / \\ [ ] : * ? ).
    Returns (name, note or None)."""
    clean = " ".join(re.sub(r"[\[\]:*?\\]", " ", dist_name.replace("/", "-")).split()) or "Distributor"
    name, note = clean, None
    if len(name) > config.MAX_SHEET_NAME:
        name = clean[:config.MAX_SHEET_NAME].rstrip()
        note = f"Distributor name was too long for an Excel sheet name (31 characters max), so it was trimmed to '{name}'."
    base, n = name, 2
    while name.lower() in used or name.lower() == "check":
        suffix = f" ({n})"
        name = base[:config.MAX_SHEET_NAME - len(suffix)] + suffix
        n += 1
    used.add(name.lower())
    return name, note


# ---------------------------------------------------------------- one distributor
def build_ws_report(dist, unit, sdm, data, conds, year, month, used_names):
    """Cases per outlet per condition for one distributor, plus the list of things the user should check."""
    sheet_name, name_note = make_ws_sheet_name(dist.name, used_names)
    rep = WsReport(dist.code, dist.name, unit, sdm, sheet_name)
    if name_note:
        rep.issues.append(("Sheet name shortened", name_note))

    sub = data.rows[data.rows["dist_code"] == dist.code].reset_index(drop=True)
    last_day = calendar.monthrange(year, month)[1]
    month_end = pd.Timestamp(year, month, last_day)
    od = sub["order_date"]
    in_month = ((od.dt.year == year) & (od.dt.month == month)).to_numpy()
    day = od.dt.day.fillna(0).to_numpy()
    pack_key = sub["pack_key"].to_numpy()
    brand_key = sub["brand_key"].to_numpy()
    cases = sub["cases"].to_numpy()

    claimed = np.zeros(len(sub), dtype=int)
    masks = []
    for c in conds:
        m = in_month & (pack_key == c.pack.upper()) & (day >= c.day_from) & (day <= c.day_to)
        if c.brands:
            m &= np.isin(brand_key, [b.upper() for b in c.brands])
        masks.append(m)
        claimed += m

    # --- outlets that have claimable cases ---
    per_cond = []
    for m in masks:
        per_cond.append(sub[m].groupby("outlet_code")["cases"].sum() if m.any() else pd.Series(dtype=float))
    codes = sorted({code for s in per_cond for code in s.index}, key=lambda x: (len(x), x))
    for code in codes:
        vals = [float(s.get(code, 0.0)) for s in per_cond]
        if all(abs(v) < 1e-12 for v in vals):
            continue
        first = sub[sub["outlet_code"] == code].iloc[0]
        rep.outlets.append(WsOutlet(code, first["outlet_name"], first["outlet_address"], vals))
        for c, v in zip(conds, vals):
            if abs(v - round(v)) > 1e-9:
                rep.issues.append(("Decimal cases", f"{first['outlet_name']} ({code}), {c.pack_label()} "
                                   f"{c.brand_text()} {c.date_text(month)}: {v:g} cases (shown rounded on the sheet; "
                                   "the exact value is used in Amount)."))
    rep.total_cases = sum(sum(o.cases) for o in rep.outlets)
    rep.total_amount = sum(v * c.rate for o in rep.outlets for v, c in zip(o.cases, conds))
    if not rep.outlets:
        rep.issues.append(("No claim", "No sales matched any condition for this distributor; the sheet shows zeros."))

    # --- WARNING: booked in the month, but delivered after the month end (still counted: booking-date rule) ---
    dd = sub["delivery_date"]
    late = (dd > month_end).to_numpy()
    for m, c in zip(masks, conds):
        for i in np.flatnonzero(m & late):
            r = sub.iloc[i]
            rep.late_lines += 1
            rep.late_cases += r["cases"]
            rep.late_amount += r["cases"] * c.rate
            rep.issues.append((
                "WARNING Delivered next month",
                f"{r['outlet_name']} ({r['outlet_code']}), invoice {r['invoice'] or '-'}: booked {_fd(r['order_date'])}, "
                f"delivered {_fd(r['delivery_date'])}. {r['cases']:g} cases of {r['pack']} {r['brand']} "
                f"(Rs {r['cases'] * c.rate:,.0f}) ARE INCLUDED in this claim because the booking date is inside the dates."))
    for m in masks:
        for i in np.flatnonzero(m & dd.isna().to_numpy()):
            r = sub.iloc[i]
            rep.issues.append(("WARNING Delivery date missing",
                               f"{r['outlet_name']} ({r['outlet_code']}), invoice {r['invoice'] or '-'}, Excel row "
                               f"{r['excel_row']}: no valid delivery date, so it cannot be checked. Cases are included."))

    # --- cases that are NOT claimed (so nothing disappears silently) ---
    cond_packs = {c.pack.upper() for c in conds}
    unc = sub[(claimed == 0) & (np.abs(cases) > 1e-12)].copy()
    unc["why"] = np.where(~in_month[unc.index], "Outside month",
                          np.where(~unc["pack_key"].isin(cond_packs), "Pack not in conditions", "Not covered"))
    reason = {"Outside month": "booked outside the report month",
              "Pack not in conditions": "this pack has no condition",
              "Not covered": "no condition matches this brand / these days"}
    for (code, pack, brand, why), g in unc.groupby(["outlet_code", "pack", "brand", "why"], sort=False):
        dts = g["order_date"].dropna()
        span = "" if dts.empty else f", booked {_fd(dts.min())}" + ("" if dts.min() == dts.max() else f" to {_fd(dts.max())}")
        rep.issues.append((why, f"{g['outlet_name'].iloc[0]} ({code}): {pack} {brand}: {g['cases'].sum():g} cases{span} "
                                f"NOT claimed ({reason[why]})."))

    # --- data problems ---
    for _, r in sub[sub["cases_flag"] != ""].iterrows():
        if r["cases_flag"] == "blank":
            rep.issues.append(("Blank cell", f"Excel row {r['excel_row']} ({r['outlet_name']}): Sales_Raw Cases is empty; counted as 0."))
        else:
            rep.issues.append(("Not a number", f"Excel row {r['excel_row']} ({r['outlet_name']}): Sales_Raw Cases contains "
                                               f"'{r['cases_flag'][5:]}'; counted as 0. Please check."))
    for _, r in sub[(sub["order_date_flag"] != "")].iterrows():
        rep.issues.append(("Bad booking date", f"Excel row {r['excel_row']} ({r['outlet_name']}): booking date is "
                                               f"{'empty' if r['order_date_flag'] == 'blank' else 'not a valid date'}; "
                                               "this line cannot be claimed."))
    neg = sub[sub["cases"] < 0]
    for _, r in neg.iterrows():
        rep.issues.append(("Negative cases", f"Excel row {r['excel_row']} ({r['outlet_name']}): {r['cases']:g} cases "
                                             "(a return?). It is used as it is."))
    promo = sub["promo_cases"].sum()
    if abs(promo) > 1e-12:
        rep.issues.append(("Promo cases", f"The file has {promo:g} Promo_Raw_Cases for this distributor. "
                                          "They are NOT included in the claim."))
    return rep


# ---------------------------------------------------------------- whole run
def generate_wholesale_report(selections, data, conds, year, month, output_folder=None, source_name=""):
    """selections = [(WsDistributor, unit, sdm), ...]; conds = [Condition, ...]
    Returns (saved file path, list of summary lines)."""
    if not selections:
        raise FriendlyError("No distributor was ticked.")
    validate_conditions(conds, year, month)
    conds = order_conditions(conds)
    output_folder = output_folder or config.OUTPUT_FOLDER
    os.makedirs(output_folder, exist_ok=True)

    used, reports = set(), []
    for dist, unit, sdm in selections:
        reports.append(build_ws_report(dist, unit, sdm, data, conds, year, month, used))
    present = set(data.rows["pack_key"])                      # a pack offered in the window may not exist in this data file
    for c in conds:
        if c.pack.upper() not in present:
            for rep in reports:
                rep.issues.append(("WARNING Pack not in data file",
                                   f"The pack '{c.pack}' of a condition does not appear anywhere in the wholesale data file "
                                   f"(packs in the file: {', '.join(data.packs)}), so it has 0 cases. Check the spelling."))

    run_notes = list(data.notes)
    for text in find_overlaps(conds, month):
        run_notes.append("WARNING Overlapping conditions: " + text)

    path = processing.unique_path(output_folder, f"Wholesale Claim {processing.date_label(year, month)}.xlsx")
    try:
        ws_writer.write_ws_workbook(reports, conds, year, month, path, source_name, run_notes)
    except PermissionError:
        raise FriendlyError(f"Could not save '{path}'. Close it in Excel and try again.")

    lines = [f"Saved: {path}", f"Sheets created: {len(reports)}"]
    late_lines = sum(r.late_lines for r in reports)
    if late_lines:
        lines.append(f"WARNING: {late_lines} line(s) ({sum(r.late_cases for r in reports):g} cases, "
                     f"Rs {sum(r.late_amount for r in reports):,.0f}) were booked inside {processing.MONTHS[month - 1]}-{year} "
                     "but DELIVERED IN THE NEXT MONTH. They ARE included in the claim (booking date decides). "
                     "Please confirm they should be claimed this month. Details are listed below and on the 'Check' sheet.")
    for r in reports:
        lines.append(f"{r.name} ({r.code}): {len(r.outlets)} outlet(s), {r.total_cases:g} cases, "
                     f"Total Amount Rs {r.total_amount:,.0f}")
    for text in run_notes:
        lines.append(f"[Whole file] {text}")
    for r in reports:
        for kind, text in r.issues:
            lines.append(f"[{kind}] {r.name} ({r.code}): {text}")
    return path, lines
