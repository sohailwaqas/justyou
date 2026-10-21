#!/usr/bin/env python3
"""
NKA Filler logic (from nka_pso_filler.py) - no UI code, importable by the handy app.

Takes the Sale file and the NKA Sale file, puts each product's sale into the
chosen key-account column(s) of the NKA file and writes a filled copy.
Standard library only (Python 3.8+).

The original matching / XML-patching logic is unchanged. Added at the bottom:
  inspect_nka_bytes() and fill_bytes()  -> in-memory wrappers used by the web page.
CLI (kept from the original):
  python -m tools.nka_filler SALE.xlsx NKA.xlsx -m 10 -y 2026 -a PSO,METRO [-o OUT.xlsx]
The desktop (tkinter) window was not carried over: the Streamlit page replaces it.
"""
import argparse
import datetime
import io
import math
import os
import re
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET

# --------------------------------------------------------------------------
# Configuration (same rules as the HTML tool)
# --------------------------------------------------------------------------
CFG = {
    "header_row": 6,
    "channel": "PSO",
    "stop_at": "LRB TOTAL",
    # first match wins (tested against UPPER-CASE text)
    "brand_rules": [
        (re.compile(r"7\s*-?\s*UP.*MINT|MINT.*7\s*-?\s*UP"), "7Up Mint"),   # 7UP MINT has its own rows
        (re.compile(r"7\s*-?\s*UP"), "7Up"),                      # 7UP ZERO / FREE -> 7Up
        (re.compile(r"PEPSI.*DIET|DIET.*PEPSI"), "Diet Pepsi"),
        (re.compile(r"PEPSI"), "Pepsi Cola"),                     # PEPSI ZERO / BLACK -> Pepsi Cola
        (re.compile(r"MIRINDA"), "Mirinda Orange"),
        (re.compile(r"M\.?\s*DEW|MTN|MOUNTAIN"), "Mountain Dew"),
        (re.compile(r"TEEM"), "Teem"),
        (re.compile(r"S-M-JUICE|SLICE"), "Slice Mango"),
        (re.compile(r"STING"), "Sting Red"),
        (re.compile(r"AQUAFINA"), "Aquafina"),
    ],
    "pack_pref": {"Aquafina|500|PET": 24, "Aquafina|1500|PET": 6},
}

MONTH_NAMES = ["January", "February", "March", "April", "May", "June", "July",
               "August", "September", "October", "November", "December"]

R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


# --------------------------------------------------------------------------
# Small helpers
# --------------------------------------------------------------------------
def js_str(x):
    """String conversion that behaves like JavaScript's String()."""
    if x is None:
        return ""
    if isinstance(x, bool):
        return "true" if x else "false"
    if isinstance(x, float):
        if not math.isfinite(x):
            return "NaN" if math.isnan(x) else ("Infinity" if x > 0 else "-Infinity")
        if x == int(x):
            return str(int(x))
    return str(x)


def clean(s):
    return re.sub(r"\s+", " ", js_str(s).replace("\u00a0", " ")).strip()


def norm(s):
    return clean(s).upper()


def is_num(x):
    return isinstance(x, (int, float)) and not isinstance(x, bool) and not math.isnan(x)


def col_num(ref):
    n = 0
    for ch in re.sub(r"[0-9]", "", ref):
        n = n * 26 + ord(ch) - 64
    return n


def col_name(n):
    s = ""
    while n > 0:
        m = (n - 1) % 26
        s = chr(65 + m) + s
        n = (n - 1) // 26
    return s


def local(tag):
    return tag.rsplit("}", 1)[-1]


def fmt_num(x):
    return str(int(round(x))) if abs(x - round(x)) < 1e-9 else f"{x:.2f}"


def brand_key(s):
    """Brand label compared ignoring case, spaces and punctuation ('7-Up Mint' == '7UP MINT')."""
    return re.sub(r"[^A-Z0-9]", "", norm(s))


def period_label(month, year):
    return f"{MONTH_NAMES[month - 1][:3]} {year}"


def parse_period(v):
    """Read a month/year/date cell value -> (year or None, month or None)."""
    if v is None or v == "":
        return None, None
    if is_num(v):
        if v >= 20000:                                   # Excel date serial
            d = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=float(v))
            return d.year, d.month
        if v == int(v):
            iv = int(v)
            if 1 <= iv <= 12:
                return None, iv
            if 1900 <= iv <= 2100:
                return iv, None
        return None, None
    t = norm(v)
    m = re.match(r"^(\d{4})[-/.](\d{1,2})", t)
    if m:
        return int(m.group(1)), int(m.group(2))
    m = re.match(r"^(\d{1,2})[-/.](\d{4})$", t)
    if m:
        return int(m.group(2)), int(m.group(1))
    m = re.match(r"^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})", t)      # day/month/year
    if m:
        return int(m.group(3)), int(m.group(2))
    if re.fullmatch(r"\d{1,2}", t) and 1 <= int(t) <= 12:
        return None, int(t)
    if re.fullmatch(r"(19|20)\d\d", t):
        return int(t), None
    mon = re.search(r"\b(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)[A-Z]*", t)
    if not mon:
        return None, None
    month = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"].index(mon.group(1)) + 1
    y = re.search(r"(?<!\d)((?:19|20)\d\d)(?!\d)", t)
    if y:
        return int(y.group(1)), month
    y = re.search(r"[-/' ](\d\d)$", t)
    return (2000 + int(y.group(1)) if y else None), month


# --------------------------------------------------------------------------
# Generic xlsx reading
# --------------------------------------------------------------------------
class Book:
    def __init__(self, zf):
        self.zf = zf
        names = set(zf.namelist())
        wb = ET.fromstring(zf.read("xl/workbook.xml"))
        rel = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        rels = {r.get("Id"): r.get("Target") for r in rel.iter() if local(r.tag) == "Relationship"}

        self.sst = []
        if "xl/sharedStrings.xml" in names:
            sx = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in sx.iter():
                if local(si.tag) == "si":
                    self.sst.append("".join((t.text or "") for t in si.iter() if local(t.tag) == "t"))

        self.sheets = []
        for s in wb.iter():
            if local(s.tag) != "sheet":
                continue
            rid = s.get(R_NS + "id")
            if rid is None:
                rid = next((v for k, v in s.attrib.items() if k.endswith("}id")), None)
            target = rels[rid].lstrip("/")
            if not target.startswith("xl/"):
                target = "xl/" + target
            self.sheets.append({"name": s.get("name"), "path": target})


class Sheet:
    """rows: {row_number: {col_number: (value, has_formula)}}  plus raw xml text."""

    def __init__(self, book, info):
        raw = book.zf.read(info["path"])
        self.text = raw.decode("utf-8")
        self.rows = {}
        root = ET.fromstring(raw)
        for row in root.iter():
            if local(row.tag) != "row":
                continue
            r = int(row.get("r"))
            cells = {}
            for c in row:
                if local(c.tag) != "c":
                    continue
                ref, t = c.get("r"), c.get("t")
                v = next((e for e in c.iter() if local(e.tag) == "v"), None)
                f = next((e for e in c.iter() if local(e.tag) == "f"), None)
                val = None
                if t == "inlineStr":
                    val = "".join((e.text or "") for e in c.iter() if local(e.tag) == "t")
                elif v is not None:
                    rawv = v.text or ""
                    if t == "s":
                        try:
                            val = book.sst[int(rawv)]
                        except (ValueError, IndexError):
                            val = None
                    elif t in ("str", "e"):
                        val = rawv
                    elif t == "b":
                        val = rawv == "1"
                    else:
                        try:
                            val = float(rawv)
                        except ValueError:
                            val = float("nan")
                cells[col_num(ref)] = (val, f is not None)
            self.rows[r] = cells

    def val(self, r, c):
        x = self.rows.get(r, {}).get(c)
        return x[0] if x else None


# --------------------------------------------------------------------------
# Sale file
# --------------------------------------------------------------------------
PERIOD_HEADERS = {"MONTH": "month", "YEAR": "year", "DATE": "date", "PERIOD": "date",
                  "MONTH YEAR": "date", "MONTH_YEAR": "date", "MONTH-YEAR": "date"}


def read_sale(zf, month=None, year=None):
    book = Book(zf)
    want = {norm("Brand NAME"): "brand", norm("PKG_NAME"): "pack", norm("Sales_Raw Cases"): "qty"}
    for info in book.sheets:
        S = Sheet(book, info)
        for hr in range(1, 16):
            cols, pcols = {}, {}
            for c, (v, _f) in S.rows.get(hr, {}).items():
                k = want.get(norm(v))
                if k:
                    cols[k] = c
                pk = PERIOD_HEADERS.get(norm(v))
                if pk:
                    pcols[pk] = c
            if len(cols) < 3:
                continue
            totals, bad = {}, []
            used = other = unparsed = 0
            for r, cells in S.rows.items():
                if r <= hr:
                    continue
                g = lambda c: (cells.get(c) or (None,))[0]
                b, p, q = g(cols["brand"]), g(cols["pack"]), g(cols["qty"])
                if b is None or b == "" or p is None or p == "":
                    continue
                # --- month / year filter (only if the sale file has such a column) ---
                if pcols and month and year:
                    verdict = True
                    for pc in pcols.values():
                        y, m = parse_period(g(pc))
                        if y is None and m is None:
                            verdict = None
                            break
                        if (m is not None and m != month) or (y is not None and y != year):
                            verdict = False
                            break
                    if verdict is None:
                        unparsed += 1
                        continue
                    if verdict is False:
                        other += 1
                        continue
                if not is_num(q):
                    if q is not None and q != "":
                        bad.append(r)
                    continue
                key = (clean(b), clean(p))
                totals[key] = totals.get(key, 0) + q
                used += 1
            if pcols and month and year and used == 0:
                raise ValueError(f"The sale file has a {'/'.join(sorted(pcols))} column but no rows "
                                 f"for {period_label(month, year)}.")
            return {"totals": totals, "tab": info["name"], "bad": bad,
                    "period": {"cols": sorted(pcols), "used": used, "other": other, "unparsed": unparsed}}
    raise ValueError("Could not find a tab in the SALE file with columns: "
                     "Brand NAME, PKG_NAME, Sales_Raw Cases.")


# --------------------------------------------------------------------------
# NKA file
# --------------------------------------------------------------------------
def container_of(d):
    d = d.upper()
    if "GB_R" in d:
        return "RB"
    if "PET_NR" in d:
        return "PET"
    if "CAN_NR" in d:
        return "CAN"
    if "TETRA" in d:
        return "TETRA"
    return "OTHER"


def header_accounts(S):
    """Key-account columns = text headers in the header row (formula / TOTAL headers skipped)."""
    out, seen = [], set()
    for c in sorted(S.rows.get(CFG["header_row"], {})):
        v, has_f = S.rows[CFG["header_row"]][c]
        name = clean(v)
        if has_f or not isinstance(v, str) or not name or "TOTAL" in name.upper():
            continue
        if c <= 7 and norm(name) != CFG["channel"]:      # B..G hold brand / pack details
            continue
        if norm(name) in seen:
            continue
        seen.add(norm(name))
        out.append((name, c))
    return out


def tab_matches_period(tab, month, year):
    t = norm(tab)
    if not month or MONTH_NAMES[month - 1][:3].upper() not in t:
        return False
    if str(year) in t or re.search(r"(?<!\d)%02d(?!\d)" % (year % 100), t):
        return True
    return not re.search(r"(?<!\d)(\d{2}|\d{4})(?!\d)", t)     # month named, no other year in the name


def scan_nka(zf, month=None, year=None):
    book = Book(zf)
    hdr = CFG["header_row"]
    cands = []
    for info in book.sheets:
        S = Sheet(book, info)
        if norm(S.val(hdr, 2)) != "BRAND":
            continue
        accts = header_accounts(S)
        if accts:
            cands.append((info, S, accts))
    if not cands:
        raise ValueError("Could not find the NKA tab (needs BRAND in B6 and the account columns in row 6).")
    info, S, accts = cands[0]
    note = None
    if len(cands) > 1:
        hit = [c for c in cands if tab_matches_period(c[0]["name"], month, year)]
        if hit:
            info, S, accts = hit[0]
            note = f'Used NKA tab "{info["name"]}" for {period_label(month, year)}.'
        else:
            note = (f'The NKA file has {len(cands)} similar tabs and none is named for '
                    f'{period_label(month, year)}; used "{info["name"]}".')
    rows = []
    max_r = max(S.rows) if S.rows else 0
    for r in range(hdr + 1, max_r + 1):
        b, d, e = S.val(r, 2), S.val(r, 4), S.val(r, 5)
        if norm(b) == CFG["stop_at"]:
            break
        if not (isinstance(b, str) and isinstance(d, str) and isinstance(e, str)) or not b.strip():
            continue
        if S.rows[r][4][1]:          # column D is a formula -> skip
            continue
        ml, bpc = S.val(r, 6), S.val(r, 7)
        rows.append({
            "row": r, "brand": clean(b), "pack": clean(e),
            "ml": ml if is_num(ml) else None,
            "bpc": bpc if is_num(bpc) else None,
            "cont": container_of(d),
        })
    return {"info": info, "S": S, "accounts": accts, "rows": rows, "note": note}


def inspect_nka(nka_path, month=None, year=None):
    """Used by the window: which key-account columns does this NKA file have?"""
    with zipfile.ZipFile(nka_path) as z:
        n = scan_nka(z, month, year)
    return {"tab": n["info"]["name"], "accounts": [a for a, _c in n["accounts"]], "note": n["note"]}


# --------------------------------------------------------------------------
# Matching
# --------------------------------------------------------------------------
def map_brand(b):
    t = norm(b)
    for rx, nb in CFG["brand_rules"]:
        if rx.search(t):
            return nb
    return None


def parse_pack(p):
    t = norm(p)
    m = re.match(r"^(\d+(?:\.\d+)?)\s*(ML|LTR|LITRE|LITER|L)\b", t)
    if not m:
        return None
    ml = float(m.group(1)) * (1 if m.group(2) == "ML" else 1000)
    if re.search(r"\bRB\b", t):
        cont = "RB"
    elif re.search(r"\bPET\b", t):
        cont = "PET"
    elif re.search(r"\bCAN\b", t):
        cont = "CAN"
    elif re.search(r"T\.?\s*PACK|TETRA", t):
        cont = "TETRA"
    else:
        return None
    bt = re.search(r"(\d+)\s*BT\b", t)
    return {"ml": ml, "cont": cont, "bt": int(bt.group(1)) if bt else None}


def find_row(rows, nb, pk):
    nk = brand_key(nb)
    c = [x for x in rows if brand_key(x["brand"]) == nk and x["ml"] is not None
         and abs(x["ml"] - pk["ml"]) < 1e-6 and x["cont"] == pk["cont"]]
    if pk["bt"] is not None:
        f = [x for x in c if x["bpc"] == pk["bt"]]
        if f:
            c = f
    if len(c) > 1:
        pref = CFG["pack_pref"].get(f"{nb}|{int(math.floor(pk['ml'] + 0.5))}|{pk['cont']}")
        f = [x for x in c if pref is not None and x["bpc"] == pref]
        if f:
            return {"row": f[0]}
        return {"row": c[0], "note": f"had {len(c)} possible packs; used {c[0]['pack']}"}
    return {"row": c[0] if c else None}


# --------------------------------------------------------------------------
# Patching the NKA sheet XML (only the PSO cells are touched)
# --------------------------------------------------------------------------
def patch_sheet(xml, pso_col, rows, placed):
    L = col_name(pso_col)
    cleared = 0
    attrs = r'((?:\s+[\w:]+="[^"]*")*)'
    strip_t = lambda a: re.sub(r'\s+t="[^"]*"', "", a, count=1)

    for x in rows:
        ref = L + str(x["row"])
        val = placed.get(x["row"])

        row_re = re.compile(r'(<row r="%d"(?=[\s>/])[^>]*?(?<!/)>)([\s\S]*?)(</row>)' % x["row"])
        m = row_re.search(xml)
        if not m:
            # self-closing empty row?  expand it so a value can be placed
            sc = re.search(r'<row r="%d"(?=[\s/])([^>]*?)\s*/>' % x["row"], xml)
            if sc and val is not None:
                xml = xml[:sc.start()] + '<row r="%d"%s></row>' % (x["row"], sc.group(1)) + xml[sc.end():]
                m = row_re.search(xml)
            if not m:
                if val:
                    raise ValueError(f"Row {x['row']} not found in sheet XML")
                continue

        inner = m.group(2)
        if val is None:
            num = None
        elif abs(val - round(val)) < 1e-9:
            num = str(int(round(val)))
        else:
            num = f"{val:.6f}".rstrip("0").rstrip(".")

        self_re = re.compile(r'<c r="%s"%s\s*/>' % (ref, attrs))
        full_re = re.compile(r'<c r="%s"%s\s*>([\s\S]*?)</c>' % (ref, attrs))
        state = {"hit": False}

        def full_cb(mm):
            nonlocal cleared
            state["hit"] = True
            a, body = mm.group(1), mm.group(2)
            if re.search(r"<f[ >/]", body):          # never overwrite a formula
                return mm.group(0)
            vm = re.search(r"<v>([^<]*)</v>", body)
            if vm:
                try:
                    nonzero = float(vm.group(1)) != 0
                except ValueError:
                    nonzero = True
                if nonzero and num is None:
                    cleared += 1
            if num is None:
                return '<c r="%s"%s/>' % (ref, strip_t(a))
            return '<c r="%s"%s><v>%s</v></c>' % (ref, strip_t(a), num)

        inner = full_re.sub(full_cb, inner, count=1)

        if not state["hit"]:
            def self_cb(mm):
                state["hit"] = True
                if num is None:
                    return mm.group(0)
                return '<c r="%s"%s><v>%s</v></c>' % (ref, strip_t(mm.group(1)), num)
            inner = self_re.sub(self_cb, inner, count=1)

        if not state["hit"] and num is not None:     # cell missing: insert in column order
            pos = len(inner)
            for mm in re.finditer(r'<c r="([A-Z]+)\d+"', inner):
                if col_num(mm.group(1)) > pso_col:
                    pos = mm.start()
                    break
            inner = inner[:pos] + '<c r="%s"><v>%s</v></c>' % (ref, num) + inner[pos:]

        xml = xml[:m.start()] + m.group(1) + inner + m.group(3) + xml[m.end():]
    return xml, cleared


def set_full_calc_on_load(wb_xml):
    """Ask Excel to recalculate all totals when the file is opened."""
    if re.search(r"<calcPr[^>]*/>", wb_xml):
        def cb(mm):
            rest = re.sub(r'\s+fullCalcOnLoad="[^"]*"', "", mm.group(1))
            return "<calcPr" + rest + ' fullCalcOnLoad="1"/>'
        return re.sub(r"<calcPr([^>]*?)\s*/>", cb, wb_xml, count=1)
    # calcPr must come before these elements in the workbook schema
    after = re.search(r"<(?:oleSize|customWorkbookViews|pivotCaches|smartTagPr|smartTagTypes|"
                      r"webPublishing|fileRecoveryPr|webPublishObjects|extLst)\b", wb_xml)
    tag = '<calcPr fullCalcOnLoad="1"/>'
    if after:
        return wb_xml[:after.start()] + tag + wb_xml[after.start():]
    return wb_xml.replace("</workbook>", tag + "</workbook>")


# --------------------------------------------------------------------------
# Main job
# --------------------------------------------------------------------------
def fill(sale_path, nka_path, out_path, accounts=None, month=None, year=None):
    accounts = accounts or [CFG["channel"]]
    with zipfile.ZipFile(sale_path) as sale_zip:
        sale = read_sale(sale_zip, month, year)

    with zipfile.ZipFile(nka_path) as nka_zip:
        nka = scan_nka(nka_zip, month, year)
        acc_map = {norm(n): (n, c) for n, c in nka["accounts"]}
        selected = []
        for a in accounts:
            if norm(a) not in acc_map:
                raise ValueError(f'Account "{a}" was not found in row {CFG["header_row"]} of the NKA tab '
                                 f'"{nka["info"]["name"]}". Available: ' + ", ".join(n for n, _ in nka["accounts"]))
            selected.append(acc_map[norm(a)])

        placed, lines, unmatched = {}, [], []
        for (sb, sp), qty in sale["totals"].items():
            nb, pk = map_brand(sb), parse_pack(sp)
            hit = find_row(nka["rows"], nb, pk) if (nb and pk) else {"row": None}
            if not hit["row"]:
                why = ("brand not recognised" if not nb
                       else "pack not understood" if not pk
                       else f"no {nb} row for this pack in NKA")
                unmatched.append({"sb": sb, "sp": sp, "qty": qty, "why": why})
                continue
            r = hit["row"]["row"]
            placed[r] = placed.get(r, 0) + qty
            lines.append({"sb": sb, "sp": sp, "qty": qty, "row": hit["row"], "note": hit.get("note")})

        new_xml, cleared = nka["S"].text, 0
        for _name, col in selected:                      # only the ticked columns are touched
            new_xml, c = patch_sheet(new_xml, col, nka["rows"], placed)
            cleared += c
        sheet_path = nka["info"]["path"]
        wb_xml = set_full_calc_on_load(nka_zip.read("xl/workbook.xml").decode("utf-8"))

        tmp = out_path + ".tmp"
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
            for item in nka_zip.infolist():
                data = nka_zip.read(item.filename)
                if item.filename == sheet_path:
                    data = new_xml.encode("utf-8")
                elif item.filename == "xl/workbook.xml":
                    data = wb_xml.encode("utf-8")
                out.writestr(item, data, compress_type=zipfile.ZIP_DEFLATED)
    os.replace(tmp, out_path)

    notes = []
    if nka["note"]:
        notes.append(nka["note"])
    pr = sale["period"]
    if pr["cols"] and month and year:
        t = f'Sale file has a {"/".join(pr["cols"])} column: used {pr["used"]} rows for {period_label(month, year)}'
        if pr["other"]:
            t += f', ignored {pr["other"]} rows from other periods'
        if pr["unparsed"]:
            t += f', skipped {pr["unparsed"]} rows whose month/year could not be read'
        notes.append(t + ".")
    elif month and year:
        notes.append(f'The sale file has no Month/Year/Date column, so all its rows were used as '
                     f'{period_label(month, year)} sales.')

    return {
        "lines": sorted(lines, key=lambda l: l["row"]["row"]),
        "unmatched": unmatched,
        "bad": sale["bad"],
        "sale_tab": sale["tab"],
        "nka_tab": nka["info"]["name"],
        "accounts": [n for n, _ in selected],
        "month": month, "year": year,
        "notes": notes,
        "sale_total": sum(sale["totals"].values()),
        "placed_total": sum(placed.values()),
        "cleared": cleared,
    }


def default_output_path(nka_path, month=None, year=None):
    base, ext = os.path.splitext(nka_path)
    ext = ext if ext.lower() in (".xlsx", ".xlsm") else ".xlsx"
    tag = f" - {period_label(month, year)}" if month and year else ""
    return f"{base}{tag} - filled{ext}"


def summary_message(r, out_path):
    ok = not r["unmatched"] and abs(r["sale_total"] - r["placed_total"]) < 1e-6
    name = os.path.basename(out_path)
    cols = ", ".join(r["accounts"])
    when = f' for {period_label(r["month"], r["year"])}' if r["month"] and r["year"] else ""
    if ok:
        return True, (f'Done{when}. All {fmt_num(r["sale_total"])} cases from the sale file are in the '
                      f'{cols} column{"s" if len(r["accounts"]) > 1 else ""}. '
                      f'Saved as "{name}" next to the NKA file.')
    return False, (f'Saved "{name}"{when}, but {fmt_num(r["sale_total"] - r["placed_total"])} cases '
                   f'could NOT be placed - see the red lines below.')


def previous_month():
    t = datetime.date.today().replace(day=1) - datetime.timedelta(days=1)
    return t.month, t.year


def parse_month_arg(s):
    s = str(s).strip()
    if s.isdigit() and 1 <= int(s) <= 12:
        return int(s)
    for i, n in enumerate(MONTH_NAMES):
        if n.lower().startswith(s.lower()[:3]) and len(s) >= 3:
            return i + 1
    raise ValueError(f'Cannot read month "{s}" (use 1-12 or a name such as Oct).')


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------
def run_cli(args):
    pm, py = previous_month()
    month = parse_month_arg(args.month) if args.month else pm
    year = args.year or py
    accounts = [a.strip() for a in args.accounts.split(",") if a.strip()] if args.accounts else [CFG["channel"]]
    out = args.output or default_output_path(args.nka, month, year)
    r = fill(args.sale, args.nka, out, accounts, month, year)
    ok, msg = summary_message(r, out)
    print(msg)
    for n in r["notes"]:
        print("Note:", n)
    print()
    hdr = f'{"Sale brand":<22}{"Sale pack":<28}{"Cases":>8}  {"NKA row":<8}Placed in NKA as'
    print(hdr)
    print("-" * len(hdr))
    for l in r["lines"]:
        note = f' ({l["note"]})' if l["note"] else ""
        print(f'{l["sb"]:<22}{l["sp"]:<28}{fmt_num(l["qty"]):>8}  {l["row"]["row"]:<8}'
              f'{l["row"]["brand"]} / {l["row"]["pack"]}{note}')
    for u in r["unmatched"]:
        print(f'{u["sb"]:<22}{u["sp"]:<28}{fmt_num(u["qty"]):>8}  {"-":<8}NOT PLACED: {u["why"]}')
    print(f'\nTotal: {fmt_num(r["placed_total"])} of {fmt_num(r["sale_total"])}')
    if r["bad"]:
        print("Sale rows with a non-numeric quantity were ignored: row " + ", ".join(map(str, r["bad"])))
    return 0 if ok else 1


# --------------------------------------------------------------------------
# In-memory wrappers for the web app (added; logic above is untouched)
# --------------------------------------------------------------------------
def inspect_nka_bytes(data, month=None, year=None):
    """Same as inspect_nka() but for uploaded bytes."""
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        n = scan_nka(z, month, year)
    return {"tab": n["info"]["name"], "accounts": [a for a, _c in n["accounts"]], "note": n["note"]}


def fill_bytes(sale_bytes, nka_bytes, accounts=None, month=None, year=None, ext=".xlsx"):
    """Run fill() on uploaded bytes. Returns (result dict, filled workbook bytes)."""
    ext = ext if ext.lower() in (".xlsx", ".xlsm") else ".xlsx"
    with tempfile.TemporaryDirectory() as tmp:
        out_path = os.path.join(tmp, "filled" + ext)
        r = fill(io.BytesIO(sale_bytes), io.BytesIO(nka_bytes), out_path, accounts, month, year)
        with open(out_path, "rb") as fh:
            return r, fh.read()


def main(argv=None):
    ap = argparse.ArgumentParser(description="Fill the key-account columns of the NKA file from the sale file.")
    ap.add_argument("sale", help="Sale file (.xlsx/.xlsm)")
    ap.add_argument("nka", help="NKA Sale file (.xlsx/.xlsm)")
    ap.add_argument("-m", "--month", help="Month (1-12 or name). Default: last month")
    ap.add_argument("-y", "--year", type=int, help="Year, e.g. 2026. Default: year of last month")
    ap.add_argument("-a", "--accounts", help="Comma-separated key accounts to fill (default: PSO)")
    ap.add_argument("-o", "--output", help='Output path (default: next to the NKA file, "... - filled.xlsx")')
    try:
        return run_cli(ap.parse_args(argv))
    except (ValueError, OSError, zipfile.BadZipFile) as e:
        print("Error:", e)
        return 1


if __name__ == "__main__":
    sys.exit(main())
