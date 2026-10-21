"""
toll_core.py - shared engine (no GUI): PDF -> trips -> filters -> grouped sets.
Used by both app.py (Streamlit web app) and toll_calculator.py (desktop).
"""
import re
import datetime as dt
from statistics import median
from difflib import SequenceMatcher
from collections import Counter

import numpy as np
from PIL import Image

try:
    import pymupdf as fitz
except ImportError:                      # older PyMuPDF
    import fitz

_ocr_engine = None


def _get_ocr():
    global _ocr_engine
    if _ocr_engine is None:
        from rapidocr_onnxruntime import RapidOCR
        _ocr_engine = RapidOCR()
    return _ocr_engine


def _run_ocr(img):
    result, _ = _get_ocr()(np.array(img.convert("RGB")))
    words = []
    for box, text, _conf in (result or []):
        xs = [p[0] for p in box]
        ys = [p[1] for p in box]
        words.append(dict(text=text.strip(), x0=min(xs), x1=max(xs),
                          cx=sum(xs) / 4, cy=sum(ys) / 4))
    return words


# --------------------------------------------------------------------------
def _find(words, pattern):
    rx = re.compile(pattern, re.I)
    hits = [w for w in words if rx.search(w["text"])]
    return min(hits, key=lambda w: w["cy"]) if hits else None


def find_header(words):
    """Locate header cells. Returns dict or None if this isn't an upright table."""
    h = {k: _find(words, p) for k, p in dict(
        entry=r"\bentry\b", exit=r"\bexit\b", rem=r"remaining",
        mot=r"motorway", date=r"date", sr=r"^sr\b").items()}
    if not (h["entry"] and h["exit"] and h["rem"] and h["date"]):
        return None
    # all header cells must sit on (roughly) the same horizontal band
    if abs(h["entry"]["cy"] - h["exit"]["cy"]) > 40:
        return None
    return h


def detect_rotation(page, dpi=110):
    """Try the 4 orientations on a low-res render; return the angle (PIL, CCW)
    that makes the header readable."""
    base = _render(page, dpi)
    for angle in (-90, 0, 90, 180):          # most common first
        img = base.rotate(angle, expand=True) if angle else base
        if find_header(_run_ocr(img)):
            return angle
    raise RuntimeError("Could not find the table header (Entry Station / Exit "
                       "Station ...) on this page in any orientation.")


def _render(page, dpi):
    pix = page.get_pixmap(dpi=dpi)
    return Image.frombytes("RGB", (pix.width, pix.height), pix.samples)


# --------------------------------------------------------------------------
def parse_page(words, width, hdr=None):
    """Rebuild table rows from positioned words of one upright page."""
    hdr = hdr or find_header(words)
    if hdr is None:
        return [], None
    W = width
    srC = hdr["sr"]["cx"] if hdr["sr"] else min(w["cx"] for w in words)
    dateC = hdr["date"]["cx"]
    entryC, exitC = hdr["entry"]["cx"], hdr["exit"]["cx"]
    d = exitC - entryC
    st_lo, st_mid, st_hi = entryC - d / 2, entryC + d / 2, exitC + d / 2
    remC = hdr["rem"]["cx"]
    num_hi = (remC + hdr["mot"]["cx"]) / 2 if hdr["mot"] else remC + d / 2
    top = hdr["entry"]["cy"] + 0.012 * W

    body = [w for w in words if w["cy"] > top]
    sr_tokens = sorted((w for w in body if w["text"].isdigit()
                        and abs(w["cx"] - srC) < 0.02 * W), key=lambda w: w["cy"])

    # Row grid.  Serial numbers give the first estimate of where rows are, but a
    # missed / garbled serial number must never make a whole row disappear, so:
    #   (a) fill gaps that are a whole number of row-heights wide, and
    #   (b) create rows for any text that sits away from every known row.
    centers = [[w["cy"], int(w["text"])] for w in sr_tokens]
    gaps = [q[0] - p[0] for p, q in zip(centers, centers[1:]) if q[0] - p[0] > 0.012 * W]
    if gaps:
        mg = median(gaps)
        pitch = median([g for g in gaps if g < 1.4 * mg])
    else:
        pitch = 0.023 * W
    for p, q in zip(list(centers), list(centers)[1:]):
        k = round((q[0] - p[0]) / pitch)
        for j in range(1, k):
            centers.append([p[0] + (q[0] - p[0]) * j / k, None])
    near = lambda w: min((abs(c[0] - w["cy"]) for c in centers), default=1e9)
    orphans = sorted((w for w in body if near(w) > 0.6 * pitch), key=lambda w: w["cy"])
    group = []
    for w in orphans + [None]:
        if group and (w is None or w["cy"] - group[-1]["cy"] > 0.5 * pitch):
            if len(group) >= 2:                      # ignore single stray marks
                centers.append([sum(g["cy"] for g in group) / len(group), None])
            group = []
        if w is not None:
            group.append(w)
    centers.sort()
    if not centers:
        return [], hdr
    ys = [c[0] for c in centers]
    rows = {i: [] for i in range(len(centers))}
    for w in body:
        i = min(range(len(ys)), key=lambda k: abs(ys[k] - w["cy"]))
        if abs(ys[i] - w["cy"]) < pitch * 0.6:
            rows[i].append(w)

    trips = []
    for i, (_cy, sr_no) in enumerate(centers):
        ws = sorted(rows[i], key=lambda w: (round(w["cy"] / 8), w["cx"]))
        cell = lambda lo, hi: " ".join(w["text"] for w in ws if lo <= w["cx"] < hi).strip()

        dtxt = cell(dateC - 0.045 * W, dateC + 0.045 * W)
        md = re.search(r"(20\d\d)\D?(\d\d)\D?(\d\d)", dtxt)
        mt = re.search(r"(\d\d):(\d\d):(\d\d)", dtxt)
        try:
            date = dt.date(*map(int, md.groups())) if md else None
        except ValueError:
            date = None

        entry = re.sub(r"\s+", " ", cell(st_lo, st_mid))
        exit_ = re.sub(r"\s+", " ", cell(st_mid, st_hi))

        zone = sorted((w for w in ws if st_hi <= w["cx"] < num_hi), key=lambda w: w["cx"])
        nums = [int(w["text"]) for w in zone if w["text"].isdigit()]
        fare = cur = rem = None
        if len(nums) >= 3:
            fare, cur, rem = nums[-3:]
        elif len(nums) == 2:
            cur, rem = nums
        # lane + fare glued together, e.g. "X3140" / "X3 140"
        if fare is None:
            for w in zone:
                m = re.fullmatch(r"[A-Za-z]\s?\d\s?(\d{2,4})", w["text"])
                if m:
                    fare = int(m.group(1))
        if fare is None and cur is not None and rem is not None and cur > rem:
            fare = cur - rem                          # derive from balances

        flag = ""
        if fare is not None and cur is not None and rem is not None and cur - rem != fare:
            flag = "balance mismatch"
        if fare is None and (entry or exit_):
            flag = "fare unreadable"
        if date is None:
            flag = "unreadable date"
        if not (date or entry or exit_ or fare or cur or rem):
            continue                                  # empty / noise row
        trips.append(dict(sr=sr_no, date=date,
                          time=mt.group(0) if mt else "", entry=entry,
                          exit=exit_, fare=fare, cur=cur, rem=rem, flag=flag))
    return trips, hdr


def ocr_pdf(path, dpi=200, progress=None):
    """OCR every page -> [(words, image_width), ...] (the slow part)."""
    doc = fitz.open(path)
    pages = []
    for n, page in enumerate(doc, 1):
        if progress:
            progress(f"Reading page {n} of {len(doc)} \u2026")
        angle = detect_rotation(page)
        img = _render(page, dpi)
        img = img.rotate(angle, expand=True) if angle else img
        pages.append((_run_ocr(img), img.width))
    return pages


def trips_from_ocr(pages):
    """OCR words -> verified list of trip rows (newest first, like the PDF)."""
    all_trips, seen = [], set()
    for words, width in pages:
        trips, _ = parse_page(words, width)
        for t in trips:
            key = (t["date"], t["time"], t["cur"], t["rem"])
            if t["time"] and t["cur"] is not None and key in seen:
                continue                       # same transaction read twice
            seen.add(key)
            all_trips.append(t)
    repair_dates(all_trips)
    repair_fares(all_trips)
    check_balance_chain(all_trips)
    return all_trips


def read_pdf(path, dpi=200, progress=None):
    """Full pipeline: PDF file -> list of trip dicts."""
    return trips_from_ocr(ocr_pdf(path, dpi, progress))


def _flag(t, msg):
    if msg not in t["flag"]:
        t["flag"] = (t["flag"] + "; " if t["flag"] else "") + msg


def _longest_nonincreasing(dates):
    n = len(dates)
    best, prev = [1] * n, [-1] * n
    for i in range(n):
        for j in range(i):
            if dates[j] >= dates[i] and best[j] + 1 > best[i]:
                best[i], prev[i] = best[j] + 1, j
    i = max(range(n), key=lambda k: best[k])
    out = []
    while i != -1:
        out.append(i)
        i = prev[i]
    return out[::-1]


def repair_dates(trips):
    """The report is sorted newest -> oldest.  A date that breaks that order is an
    OCR slip: fix it when exactly one single-digit change fits, else flag it."""
    idx = [i for i, t in enumerate(trips) if t["date"]]
    if len(idx) < 3:
        return
    keep = {idx[k] for k in _longest_nonincreasing([trips[i]["date"] for i in idx])}
    anchors = sorted(keep)
    for i, t in enumerate(trips):
        if i in keep:
            continue
        before = [a for a in anchors if a < i]
        after = [a for a in anchors if a > i]
        hi = (trips[before[-1]]["date"], trips[before[-1]]["time"]) if before else (dt.date.max, "")
        lo = (trips[after[0]]["date"], trips[after[0]]["time"]) if after else (dt.date.min, "")
        if t["date"] is None:
            if hi[0] == lo[0]:
                t["date"] = hi[0]
                t["flag"] = t["flag"].replace("unreadable date", "").strip("; ")
                _flag(t, "date inferred from neighbouring rows")
            else:
                _flag(t, "unreadable date")
            continue

        def not_after(a, b):          # a <= b, comparing times only when both known
            if a[0] != b[0]:
                return a[0] < b[0]
            return not (a[1] and b[1]) or a[1] <= b[1]

        old, text, cands = t["date"], t["date"].strftime("%Y%m%d"), set()
        for p in range(8):
            for ch in "0123456789":
                if ch != text[p]:
                    try:
                        d = dt.datetime.strptime(text[:p] + ch + text[p + 1:], "%Y%m%d").date()
                    except ValueError:
                        continue
                    if not_after(lo, (d, t["time"])) and not_after((d, t["time"]), hi):
                        cands.add(d)
        if len(cands) == 1:
            t["date"] = cands.pop()
            _flag(t, f"date corrected from {old:%d-%m-%Y} (was out of order)")
        else:
            _flag(t, "date out of order - please check")


def repair_fares(trips):
    """fare = current balance - remaining balance.  If the fare that was read
    disagrees with that, but the balances are confirmed by the neighbouring rows
    (newer.cur == this.rem and this.cur == older.rem), the fare was misread:
    correct it.  If the balances themselves look wrong, leave it flagged."""
    for i, t in enumerate(trips):
        if t["cur"] is None or t["rem"] is None or not t["fare"]:
            continue
        if t["cur"] - t["rem"] == t["fare"] or t["cur"] <= t["rem"]:
            continue
        agree = clash = 0
        newer = trips[i - 1] if i > 0 else None
        older = trips[i + 1] if i + 1 < len(trips) else None
        if newer and newer["cur"] is not None:
            agree, clash = (agree + 1, clash) if newer["cur"] == t["rem"] else (agree, clash + 1)
        if older and older["rem"] is not None:
            agree, clash = (agree + 1, clash) if older["rem"] == t["cur"] else (agree, clash + 1)
        if agree >= 1 and clash == 0:
            _flag(t, f"fare corrected from balances (read {t['fare']}, now {t['cur'] - t['rem']})")
            t["fare"] = t["cur"] - t["rem"]
            t["flag"] = t["flag"].replace("balance mismatch; ", "").replace("balance mismatch", "").strip("; ")


def check_balance_chain(trips):
    """Between two consecutive trips (no recharge in between) the newer trip's
    'current balance' must equal the older trip's 'remaining balance'.  If not,
    a trip is probably missing or misread there - flag both neighbours."""
    newer, gap = None, False
    for t in trips:
        if t["cur"] is not None and t["rem"] is not None:
            if newer is not None and not gap and newer["cur"] != t["rem"]:
                msg = "balance chain broken - a trip may be missing or misread near here"
                _flag(newer, msg)
                _flag(t, msg)
            newer, gap = t, False
        else:                      # recharge or unreadable row: chain can't be checked
            newer, gap = (None, False) if is_toll_trip(t) else (newer, True)


# --------------------------------------------------------------------------
def canonicalise_stations(trips, threshold=0.86):
    """Merge near-identical OCR spellings of the same station name."""
    names = Counter(n for t in trips for n in (t["entry"], t["exit"]) if n)
    canon = {}
    for name, _ in names.most_common():
        for known in set(canon.values()):
            if SequenceMatcher(None, name.lower(), known.lower()).ratio() >= threshold:
                canon[name] = known
                break
        else:
            canon[name] = name
    for t in trips:
        t["entry"] = canon.get(t["entry"], t["entry"])
        t["exit"] = canon.get(t["exit"], t["exit"])
    return trips


# ==========================================================================
#                      filtering, leave dates and grouping
# ==========================================================================
import datetime as _dt

COUNTED = "Counted"
ST_RANGE = "Outside date range"
ST_SUNDAY = "Sunday"
ST_LEAVE = "Leave / excluded date"
ST_BADDATE = "Unreadable date"
ST_NOFARE = "Fare unreadable"


def is_toll_trip(t):
    """Real toll trips only (recharges etc. have neither fare nor stations)."""
    return bool(t["entry"] or t["exit"] or t["fare"])


def date_span(a, b):
    """Every date from a to b inclusive."""
    if b < a:
        raise ValueError(f"Range end {fmt_date(b)} is before its start {fmt_date(a)}.")
    return {a + _dt.timedelta(days=i) for i in range((b - a).days + 1)}


DATE_FMT = "%d-%m-%Y"                      # day-month-year everywhere
_DATE_RX = re.compile(r"\d{4}-\d{1,2}-\d{1,2}|\d{1,2}[-/.]\d{1,2}[-/.]\d{4}")


def fmt_date(d):
    return d.strftime(DATE_FMT) if d else "?"


def parse_date(text):
    """'07-09-2026' (also 7/9/2026, 7.9.2026, or ISO 2026-09-07) -> date."""
    text = text.strip()
    if re.fullmatch(r"\d{4}-\d{1,2}-\d{1,2}", text):
        return _dt.datetime.strptime(text, "%Y-%m-%d").date()
    return _dt.datetime.strptime(re.sub(r"[/.]", "-", text), DATE_FMT).date()


def describe_trip(t):
    """One readable line for a warning list."""
    stations = f"{t['entry'] or '—'} → {t['exit'] or '—'}"
    fare = f" (fare {t['fare']})" if t["fare"] else ""
    return f"{fmt_date(t['date'])} {t['time']}  {stations}{fare}".replace("  ", " ")


def parse_leave_text(text):
    """'01-09-2026 to 07-09-2026, 12-09-2026' -> set of dates (day-month-year).
    Separate items with commas / semicolons / new lines; a range is two dates in
    one item (joined by 'to', '-', ':' ... anything)."""
    out = set()
    for part in re.split(r"[,;\n]+", text or ""):
        part = part.strip()
        if not part:
            continue
        found = _DATE_RX.findall(part)
        try:
            ds = [parse_date(x) for x in found]
        except ValueError:
            raise ValueError(f"'{part}' contains an invalid date. Use DD-MM-YYYY.")
        if len(ds) == 1:
            out.add(ds[0])
        elif len(ds) == 2:
            out |= date_span(ds[0], ds[1])
        else:
            raise ValueError(f"Could not understand '{part}'. Use e.g. "
                             "01-09-2026 to 07-09-2026, 12-09-2026")
    return out


def classify_trips(trips, lo, hi, exclude_sundays=False, leave_dates=()):
    """[(trip, status)] for every real toll trip; status is COUNTED or the
    reason it was left out.  This is the single source of truth for filtering."""
    leave = set(leave_dates)
    rows = []
    for t in trips:
        if not is_toll_trip(t):
            continue
        d = t["date"]
        if not t["fare"]:
            st = ST_NOFARE
        elif d is None:
            st = ST_BADDATE
        elif not (lo <= d <= hi):
            st = ST_RANGE
        elif exclude_sundays and d.weekday() == 6:
            st = ST_SUNDAY
        elif d in leave:
            st = ST_LEAVE
        else:
            st = COUNTED
        rows.append((t, st))
    return rows


def group_counted(classified):
    """{(entry, exit, fare): frequency} for the counted trips."""
    counts = {}
    for t, st in classified:
        if st == COUNTED:
            key = (t["entry"] or "\u2014", t["exit"] or "\u2014", t["fare"])
            counts[key] = counts.get(key, 0) + 1
    return dict(sorted(counts.items()))
