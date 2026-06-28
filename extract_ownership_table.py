"""Extract KSEI/BEI 5% ownership tables from PDF disclosures.

The extractor groups positioned PDF words into rows and columns, then normalizes
wrapped fields, duplicated glyphs, and inconsistent numeric formatting. It writes
one ownership record per Excel row and can emit word-level debug data.
"""

from __future__ import annotations

import argparse
import json
import os
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import pandas as pd
import pdfplumber


TICKER_RE = re.compile(r"(?<![A-Z])([A-Z]{4})(?![A-Z])")
ROWNO_RE = re.compile(r"^\s*(\d{1,5})\b")
INT_THOUSANDS_RE = re.compile(r"\d{1,3}(?:,\d{3})+")
# Values may be comma-grouped or plain integers.
INT_ANY_RE = re.compile(r"(?<!\d)(\d{1,3}(?:,\d{3})+|\d+)(?!\d)")
PCT_RE = re.compile(r"(?<!\d)(\d{1,3})\.(\d{2})(?!\d)")
REPORT_DATE_RE = re.compile(
    r"per\s+(?:tanggal\s+)?(\d{1,2})[-\s]+([A-Za-z]{3,12})[-\s]+(\d{4})",
    re.IGNORECASE,
)
GENERIC_DATE_RE = re.compile(r"(\d{1,2})[-\s]+([A-Za-z]{3,12})[-\s]+(\d{4})", re.IGNORECASE)
TABLE_DATE_RE = re.compile(r"(\d{1,2})-([A-Za-z]{3})-(\d{4})", re.IGNORECASE)
FILENAME_DATE_RE = re.compile(r"^(20\d{2})(\d{2})(\d{2})_")
BLUE_RGB = (0.0, 0.0, 1.0)

COUNTRY_HINTS = {
    "INDONESIA",
    "SINGAPORE",
    "MALAYSIA",
    "UNITED",
    "KINGDOM",
    "UNITEDSTATES",
    "AMERICA",
    "HONG",
    "KONG",
    "CHINA",
    "JAPAN",
    "KOREA",
    "THAILAND",
    "PHILIPPINES",
    "VIETNAM",
    "TAIWAN",
    "AUSTRALIA",
    "NETHERLANDS",
    "GERMANY",
    "FRANCE",
    "SWITZERLAND",
    "LUXEMBOURG",
}

MONTHS = {
    "jan": 1,
    "january": 1,
    "januari": 1,
    "feb": 2,
    "february": 2,
    "februari": 2,
    "mar": 3,
    "march": 3,
    "maret": 3,
    "apr": 4,
    "april": 4,
    "may": 5,
    "mei": 5,
    "jun": 6,
    "june": 6,
    "juni": 6,
    "jul": 7,
    "july": 7,
    "juli": 7,
    "aug": 8,
    "august": 8,
    "agustus": 8,
    "sep": 9,
    "sept": 9,
    "september": 9,
    "oct": 10,
    "october": 10,
    "okt": 10,
    "oktober": 10,
    "nov": 11,
    "november": 11,
    "dec": 12,
    "december": 12,
    "des": 12,
    "desember": 12,
}



def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", str(s or "").strip())


def parse_pdf_date_token(value: str) -> str:
    """Parse KSEI date tokens such as 05-JUN-2026 or 5 Jun 2026."""
    value = _norm(value)
    m = TABLE_DATE_RE.search(value) or REPORT_DATE_RE.search(value) or GENERIC_DATE_RE.search(value)
    if not m:
        return ""
    day, month_name, year = m.groups()
    month = MONTHS.get(month_name.lower())
    if not month:
        return ""
    try:
        return f"{int(year):04d}-{month:02d}-{int(day):02d}"
    except ValueError:
        return ""


def extract_report_dates_from_text(text: str) -> Dict[str, str]:
    """Extract report/as-of dates from the PDF cover and table header text."""
    text = _compact_spaced_text(text or "")
    report_date = ""
    m = REPORT_DATE_RE.search(text)
    if m:
        report_date = parse_pdf_date_token(" ".join(m.groups()))

    table_dates = [parse_pdf_date_token("-".join(m.groups())) for m in TABLE_DATE_RE.finditer(text)]
    table_dates = [x for x in table_dates if x]

    previous_report_date = ""
    if table_dates:
        if report_date and report_date in table_dates:
            idx = table_dates.index(report_date)
            if idx > 0:
                previous_report_date = table_dates[idx - 1]
        if not previous_report_date and len(table_dates) >= 2:
            previous_report_date = table_dates[-2]
        if not report_date:
            report_date = table_dates[-1]

    return {
        "report_date": report_date,
        "previous_report_date": previous_report_date,
    }


def filename_date(filename: str) -> str:
    m = FILENAME_DATE_RE.search(Path(filename).name)
    if not m:
        return ""
    y, month, day = m.groups()
    try:
        return f"{int(y):04d}-{int(month):02d}-{int(day):02d}"
    except ValueError:
        return ""


def _color_tuple(value: Any) -> Tuple[float, ...]:
    if value is None:
        return tuple()
    if isinstance(value, (int, float)):
        return (float(value),)
    try:
        return tuple(float(x) for x in value)
    except TypeError:
        return tuple()


def is_blue_font_color(value: Any) -> bool:
    color = _color_tuple(value)
    if len(color) < 3:
        return False
    red, green, blue = color[:3]
    return red <= 0.25 and green <= 0.35 and blue >= 0.75


def word_is_highlighted(word: Dict[str, Any]) -> bool:
    return is_blue_font_color(word.get("non_stroking_color"))


def highlighted_text(words: List[Dict[str, Any]]) -> str:
    return join_bucket([w for w in words if word_is_highlighted(w)])


def highlighted_columns(buckets: Dict[str, List[Dict[str, Any]]]) -> str:
    names = [name for name, words in buckets.items() if any(word_is_highlighted(w) for w in words)]
    return ",".join(names)


def merge_csv_values(left: str, right: str) -> str:
    values: list[str] = []
    seen: set[str] = set()
    for part in (left or "", right or ""):
        for value in str(part).split(","):
            value = value.strip()
            if value and value not in seen:
                seen.add(value)
                values.append(value)
    return ",".join(values)


def parse_delta_token(value: str) -> Optional[int]:
    value = _compact_spaced_text(value or "").strip()
    if not value or value == "-":
        return None
    m = re.search(r"-?\d{1,3}(?:,\d{3})+|-?\d+", value)
    if not m:
        return None
    return int(m.group(0).replace(",", ""))


def compute_change_fields(record: Dict[str, Any]) -> Dict[str, Any]:
    shares_before = record.get("shares_before")
    shares_after = record.get("shares_after")
    pct_before = record.get("pct_before")
    pct_after = record.get("pct_after")
    warnings = str(record.get("parse_warnings") or "")
    clean_numeric_parse = "missing_pct" not in warnings and "out_of_range" not in warnings

    delta_from_field = parse_delta_token(str(record.get("delta_shares_or_delta_field") or ""))
    delta_shares = None
    if delta_from_field is not None:
        delta_shares = delta_from_field
    elif clean_numeric_parse and shares_before is not None and shares_after is not None:
        delta_shares = int(shares_after) - int(shares_before)

    delta_pct = None
    if pct_before is not None and pct_after is not None:
        delta_pct = round(float(pct_after) - float(pct_before), 4)

    highlighted = bool(record.get("is_highlighted_change"))
    has_reliable_delta = highlighted or clean_numeric_parse
    has_numeric_change = bool(
        has_reliable_delta
        and (
            (delta_shares is not None and delta_shares != 0)
            or (delta_pct is not None and abs(delta_pct) > 1e-9)
        )
    )

    record["delta_shares"] = delta_shares
    record["delta_pct"] = delta_pct
    record["has_numeric_change"] = has_numeric_change
    record["change_reason"] = ",".join(
        reason
        for reason, active in (
            ("blue_text", highlighted),
            ("numeric_delta", has_numeric_change),
        )
        if active
    )
    return record


def load_manifest_metadata(manifest_path: str | Path | None, pdf_path: Path) -> Dict[str, Any]:
    if not manifest_path:
        return {}
    path = Path(manifest_path)
    if not path.exists():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {}

    pdf_name = pdf_path.name
    for row in data if isinstance(data, list) else []:
        local_name = Path(str(row.get("local_pdf_path") or "")).name
        attachment_name = str(row.get("attachment_filename") or "")
        if pdf_name in {local_name, attachment_name}:
            return dict(row)
    return {}


def _compact_spaced_text(s: str) -> str:
    """Collapse spaced-out glyph extraction artifacts.

    Examples:
      - "I N D O N E S I A" -> "INDONESIA"
      - "5 9 9 , 4 9 1 , 2 2 8" -> "599,491,228"
      - "4 1 . 1 0" -> "41.10"
    """
    if not s:
        return ""
    s = str(s)

    def _collapse_letters(m: re.Match) -> str:
        return m.group(0).replace(" ", "")

    s = re.sub(r"\b(?:[A-Z]\s+){2,}[A-Z]\b", _collapse_letters, s)

    # Join glyph-by-glyph numbers without merging separately formatted values.
    toks = s.split()
    out = []
    i = 0
    while i < len(toks):
        tok = toks[i]
        def is_glyph(t: str) -> bool:
            return (len(t) == 1) and (t.isdigit() or t in {",", "."})

        if is_glyph(tok):
            j = i
            run = []
            while j < len(toks) and is_glyph(toks[j]):
                run.append(toks[j])
                j += 1
            # Short runs such as "1 . 2" are too ambiguous to join safely.
            digit_count = sum(1 for t in run if t.isdigit())
            if digit_count >= 3:
                out.append("".join(run))
            else:
                out.extend(run)
            i = j
            continue

        out.append(tok)
        i += 1

    s = " ".join(out)

    s = re.sub(r"\s+", " ", s).strip()
    return s


def _is_country_token(t: str) -> bool:
    t = _compact_spaced_text(t or "").strip().upper()
    if not t:
        return False
    if t in COUNTRY_HINTS:
        return True
    # Handle country names joined to neighboring extracted text.
    if "INDONESIA" in t or "SINGAPORE" in t or "MALAYSIA" in t:
        return True
    return False


def _refine_nat_dom_status_from_positions(
    row_words: List[Dict[str, Any]],
    page_width: float,
    bucket_text: Dict[str, str],
) -> None:
    """Recover nationality, domicile, and status from word positions.

    This fallback is used when graphical or merged headers make the inferred
    column boundaries unreliable. Existing values are replaced only when the
    positional candidates are unambiguous.
    """
    cands: List[Tuple[str, float]] = []
    for w in row_words:
        txt = _compact_spaced_text(w.get("text", "")).strip()
        if not txt:
            continue
        xmid = _word_mid_x(w)
        cands.append((txt, xmid))

    # The first large number marks the right-side numeric block.
    num_xs = []
    for txt, x in cands:
        if INT_THOUSANDS_RE.search(_compact_spaced_text(txt)):
            num_xs.append(x)
        elif PCT_RE.search(_compact_spaced_text(txt)):
            num_xs.append(x)
    numeric_start_x = min(num_xs) if num_xs else (0.70 * page_width)

    status_cands = [
        (txt.upper(), x)
        for txt, x in cands
        if txt.strip().upper() in {"L", "A"} and (0.45 * page_width) <= x <= (numeric_start_x - 2)
    ]
    # Status is the rightmost L/A token before the numeric block.
    status_val = ""
    status_x = None
    if status_cands:
        status_val, status_x = sorted(status_cands, key=lambda p: p[1])[-1]

    right_edge = status_x if status_x is not None else numeric_start_x
    country_cands = [(txt, x) for txt, x in cands if _is_country_token(txt) and x < right_edge]
    domicile_val = ""
    domicile_x = None
    if country_cands:
        domicile_val, domicile_x = sorted(country_cands, key=lambda p: p[1])[-1]

    nationality_val = ""
    if domicile_x is not None:
        left_country = [(txt, x) for txt, x in country_cands if x < (domicile_x - 5)]
        if left_country:
            nationality_val = sorted(left_country, key=lambda p: p[1])[-1][0]

    if status_val:
        bucket_text["status"] = status_val

    if domicile_val:
        # Long non-country values are usually address text from a shifted column.
        cur = (bucket_text.get("domicile", "") or "").strip()
        if (not cur) or (len(cur) > 30) or (not _is_country_token(cur)):
            bucket_text["domicile"] = domicile_val

    if nationality_val:
        cur = (bucket_text.get("nationality", "") or "").strip()
        if (not cur) or (len(cur) > 30) or (not _is_country_token(cur)):
            bucket_text["nationality"] = nationality_val


def dedupe_words(words: List[Dict[str, Any]], ndigits: int = 2) -> List[Dict[str, Any]]:
    """De-dupe duplicated words caused by layered text.

    We key by (text, rounded bbox). This is similar to `dedupe_chars` in extract_blue_fast.py
    but at word-level.
    """
    seen = set()
    out = []
    for w in words:
        txt = w.get("text")
        if not txt:
            continue
        key = (
            txt,
            round(float(w.get("x0", 0.0)), ndigits),
            round(float(w.get("x1", 0.0)), ndigits),
            round(float(w.get("top", 0.0)), ndigits),
            round(float(w.get("bottom", w.get("top", 0.0))), ndigits),
        )
        if key in seen:
            continue
        seen.add(key)
        out.append(w)
    return out


def _word_mid_x(w: Dict[str, Any]) -> float:
    return (float(w["x0"]) + float(w["x1"])) / 2.0


def _word_mid_y(w: Dict[str, Any]) -> float:
    return (float(w["top"]) + float(w.get("bottom", w["top"]))) / 2.0


def cluster_rows(words: List[Dict[str, Any]], y_tol: float) -> List[List[Dict[str, Any]]]:
    """Cluster words into visual rows using y-mid with a tight tolerance."""
    words = sorted(words, key=lambda w: (_word_mid_y(w), float(w["x0"])))
    rows: List[List[Dict[str, Any]]] = []
    refs: List[float] = []
    for w in words:
        y = _word_mid_y(w)
        if not rows:
            rows.append([w])
            refs.append(y)
            continue
        if abs(y - refs[-1]) <= y_tol:
            rows[-1].append(w)
            # Smooth minor baseline differences within a visual row.
            refs[-1] = (refs[-1] * 0.7) + (y * 0.3)
        else:
            rows.append([w])
            refs.append(y)
    return rows


@dataclass
class ColumnBands:
    # Ordered left edges; each band's right edge is the next cut.
    cuts: List[Tuple[str, float]]

    def bucket(self, x_mid: float, page_width: float) -> str:
        cuts = self.cuts
        for i in range(len(cuts)):
            name, x0 = cuts[i]
            x1 = cuts[i + 1][1] if i + 1 < len(cuts) else page_width + 1
            if x_mid >= x0 and x_mid < x1:
                return name
        return cuts[-1][0]


def infer_bands_from_header(page: pdfplumber.page.Page) -> ColumnBands:
    """Infer column cutpoints from the header word positions.

    The header is stable across pages in these KSEI exports.
    We only need approximate x cutpoints.
    """
    # Header labels occupy the top portion of the page.
    hdr_words = [w for w in page.extract_words(use_text_flow=True) if float(w["top"]) < 60]
    hdr_words = dedupe_words(hdr_words)

    def x_of(label: str, default: float) -> float:
        for w in hdr_words:
            if w["text"].strip().upper() == label:
                return float(w["x0"])
        return default

    # Defaults match the standard landscape BEI disclosure layout.
    cuts = [
        ("row_no", x_of("NOKODE", 25.0)),
        ("ticker", x_of("EFEK", 35.0)),
        ("emiten", x_of("NAMA", 70.0)),  # first NAMA after kode efek
        ("broker", x_of("PEMEGANG", 130.0)),
        ("shareholder", 205.0),
        ("account_name", 280.0),
        ("address", x_of("ALAMAT", 360.0)),
        ("address2", 425.0),
        ("nationality", x_of("KEBANGSAAN", 475.0)),
        ("domicile", x_of("DOMISILI", 510.0)),
        ("status", x_of("STATUS", 535.0)),
        ("shares_before", 556.0),
        ("combined_before", 578.0),
        ("pct_before", 612.0),
        ("shares_after", 657.0),
        ("combined_after", 676.0),
        ("pct_after", 710.0),
        ("delta", 752.0),
    ]
    # Missing or duplicated labels must not produce overlapping bands.
    fixed: List[Tuple[str, float]] = []
    last = -1e9
    for name, x in cuts:
        x = float(x)
        if x <= last:
            x = last + 5.0
        fixed.append((name, x))
        last = x
    return ColumnBands(fixed)


def join_bucket(words: List[Dict[str, Any]]) -> str:
    if not words:
        return ""
    words = sorted(words, key=lambda w: float(w["x0"]))
    joined = " ".join(w["text"] for w in words)
    joined = _compact_spaced_text(joined)
    return _norm(joined)


def parse_int_token(s: str) -> Optional[int]:
    if not s:
        return None
    s = _compact_spaced_text(s)
    m = INT_ANY_RE.search(s)
    if not m:
        return None
    return int(m.group(0).replace(",", ""))


def parse_pct_token(s: str) -> Optional[float]:
    if not s:
        return None
    s = _compact_spaced_text(s)
    m = PCT_RE.search(s)
    if not m:
        return None
    return float(f"{m.group(1)}.{m.group(2)}")


# Match percentages first so values such as 41.10 remain a single token.
NUM_TOKEN_RE = re.compile(r"(?<!\w)(\d{1,3}\.\d{2}|-?\d{1,3}(?:,\d{3})+|-?\d+)(?!\w)")

def extract_numeric_tail(all_txt: str) -> Dict[str, Any]:
    """Extract the right-side numeric block by taking the last 6-7 numeric-like tokens.

    Returns a dict with keys:
      shares_before, combined_before, pct_before,
      shares_after, combined_after, pct_after,
      delta_shares_or_delta_field
    """
    out: Dict[str, Any] = {}
    s = _compact_spaced_text(all_txt or "")
    toks = NUM_TOKEN_RE.findall(s)
    if len(toks) < 6:
        return out

    tail = toks[-7:] if len(toks) >= 7 else toks[-6:]
    if len(tail) == 6:
        tail = [""] + tail

    sb, cb, pb, sa, ca, pa, d = tail

    def _pi(x: str):
        return parse_int_token(x)

    def _pp(x: str):
        return parse_pct_token(x)

    out["shares_before"] = _pi(sb)
    out["combined_before"] = _pi(cb)
    out["pct_before"] = _pp(pb)
    out["shares_after"] = _pi(sa)
    out["combined_after"] = _pi(ca)
    out["pct_after"] = _pp(pa)
    out["delta_shares_or_delta_field"] = _norm(d)
    return out


def looks_like_new_record(bucket_text: Dict[str, str]) -> bool:
    """Heuristic: decide if a clustered row begins a new logical record."""
    row_no = bucket_text.get("row_no", "")
    ticker = bucket_text.get("ticker", "")
    status = bucket_text.get("status", "")
    right = " ".join(
        bucket_text.get(k, "")
        for k in ("shares_before", "combined_before", "pct_before", "shares_after", "combined_after", "pct_after", "delta")
    )
    right = _compact_spaced_text(right)

    has_rowno = bool(ROWNO_RE.match(row_no))
    has_ticker = bool(TICKER_RE.search(ticker))
    has_numeric = bool(INT_THOUSANDS_RE.search(right) or PCT_RE.search(right))
    has_status = status.strip().upper() in {"L", "A"}

    if has_rowno and has_ticker:
        return True

    if has_ticker and (has_numeric or has_status):
        return True

    # Country text helps identify records whose numeric columns extracted poorly.
    combined = " ".join(bucket_text.values()).upper()
    if has_ticker and any(c in combined for c in COUNTRY_HINTS) and (has_numeric or has_status):
        return True

    return False


def looks_like_continuation(bucket_text: Dict[str, str], prev_ticker: str) -> bool:
    """Heuristic: row is continuation of previous row (address wrap)."""
    row_no = bucket_text.get("row_no", "")
    status = bucket_text.get("status", "")
    right = " ".join(
        bucket_text.get(k, "")
        for k in ("shares_before", "combined_before", "pct_before", "shares_after", "combined_after", "pct_after", "delta")
    )
    right = _compact_spaced_text(right)
    has_rowno = bool(ROWNO_RE.match(row_no))
    has_numeric = bool(INT_THOUSANDS_RE.search(right) or PCT_RE.search(right))
    has_status = status.strip().upper() in {"L", "A"}

    if (not has_rowno) and (not has_numeric) and (not has_status):
        return True

    ticker = bucket_text.get("ticker", "").strip().upper()
    if ticker and prev_ticker and ticker == prev_ticker and (not has_numeric):
        return True

    return False


def extract_table_from_page(
    page: pdfplumber.page.Page,
    bands: ColumnBands,
    page_no: int,
    source_file: str,
    document_info: Optional[Dict[str, Any]] = None,
    debug_words_dir: Optional[Path] = None,
) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    """Return (raw_row_debug, parsed_records) for a page."""
    document_info = document_info or {}
    # Exclude page margins and footers before positional extraction.
    crop_bbox = (15, 55, page.width - 15, page.height - 18)
    cropped = page.crop(crop_bbox)

    # Positional bucketing does not require pdfplumber's text-flow inference.
    words = cropped.extract_words(
        use_text_flow=False,
        keep_blank_chars=False,
        extra_attrs=["non_stroking_color"],
    )
    words = dedupe_words(words)

    data_words = words

    # Scale row tolerance to the document's font size.
    heights = [float(w.get("bottom", w["top"])) - float(w["top"]) for w in data_words[:500] if w.get("bottom")]
    med_h = sorted(heights)[len(heights) // 2] if heights else 6.0
    y_tol = max(1.8, min(3.2, med_h * 0.45))

    rows = cluster_rows(data_words, y_tol=y_tol)

    if debug_words_dir:
        debug_words_dir.mkdir(parents=True, exist_ok=True)
        dfw = pd.DataFrame([
            {
                "page": page_no,
                "text": w.get("text"),
                "x0": w.get("x0"),
                "x1": w.get("x1"),
                "top": w.get("top"),
                "bottom": w.get("bottom"),
                "non_stroking_color": w.get("non_stroking_color"),
                "is_blue": word_is_highlighted(w),
            }
            for w in data_words
        ])
        dfw.to_csv(debug_words_dir / f"words_p{page_no:03d}.csv", index=False)

    raw_rows_debug: List[Dict[str, Any]] = []
    parsed: List[Dict[str, Any]] = []

    current: Optional[Dict[str, Any]] = None
    prev_ticker = ""

    for ridx, row_words in enumerate(rows):
        buckets: Dict[str, List[Dict[str, Any]]] = {name: [] for name, _ in bands.cuts}
        for w in row_words:
            b = bands.bucket(_word_mid_x(w), page.width)
            buckets.setdefault(b, []).append(w)

        bucket_text = {k: join_bucket(v) for k, v in buckets.items()}
        _refine_nat_dom_status_from_positions(row_words, page.width, bucket_text)

        all_txt = join_bucket(row_words)
        all_txt = _compact_spaced_text(all_txt)
        row_highlighted_columns = highlighted_columns(buckets)
        row_highlighted_text = highlighted_text(row_words)
        raw_rows_debug.append(
            {
                "source_pdf": source_file,
                "source_file": source_file,
                "page": page_no,
                "row_index": ridx,
                "raw_text": all_txt,
                "is_highlighted_change": bool(row_highlighted_columns),
                "highlighted_columns": row_highlighted_columns,
                "highlighted_text": row_highlighted_text,
                **{f"col_{k}": bucket_text.get(k, "") for k, _ in bands.cuts},
            }
        )

        # Multi-page exports may repeat the table header.
        if not all_txt or "NOKODE" in all_txt.replace(" ", "").upper():
            continue

        t = bucket_text.get("ticker", "")
        # Tickers can be joined to row numbers during PDF text extraction.
        if not TICKER_RE.search(t):
            glued = (bucket_text.get("row_no", "") + " " + t).upper()
            m = TICKER_RE.search(glued)
            if m:
                t = m.group(1)
        else:
            t = TICKER_RE.search(t).group(1)

        bucket_text["ticker"] = t

        if current and looks_like_continuation(bucket_text, prev_ticker=prev_ticker):
            cont = " ".join(
                s for s in [bucket_text.get("address", ""), bucket_text.get("address2", ""), bucket_text.get("emiten", ""), bucket_text.get("broker", ""), bucket_text.get("shareholder", ""), bucket_text.get("account_name", "")] if s
            ).strip()
            if cont:
                if current.get("address2"):
                    current["address2"] = _norm(current["address2"] + " " + cont)
                elif current.get("address"):
                    current["address"] = _norm(current["address"] + " " + cont)
                else:
                    current["address"] = _norm(cont)
                current["raw_text"] = _norm(current.get("raw_text", "") + " | " + all_txt)
                current["parse_warnings"] = _norm((current.get("parse_warnings", "") + " continuation").strip())
                if row_highlighted_columns:
                    current["is_highlighted_change"] = True
                    current["highlighted_columns"] = merge_csv_values(
                        current.get("highlighted_columns", ""),
                        row_highlighted_columns,
                    )
                    current["highlighted_text"] = _norm(
                        " | ".join(
                            x
                            for x in [current.get("highlighted_text", ""), row_highlighted_text]
                            if x
                        )
                    )
            continue

        is_new = looks_like_new_record(bucket_text)
        if not is_new:
            # Remaining unclassified rows are typically footer content.
            continue

        if current:
            parsed.append(current)

        row_no = bucket_text.get("row_no", "")
        m = ROWNO_RE.match(row_no)
        row_no_val = m.group(1) if m else ""

        status = bucket_text.get("status", "").strip().upper()
        if status not in {"L", "A"}:
            status = ""

        shares_before = parse_int_token(bucket_text.get("shares_before", ""))
        combined_before = parse_int_token(bucket_text.get("combined_before", ""))
        pct_before = parse_pct_token(bucket_text.get("pct_before", ""))
        shares_after = parse_int_token(bucket_text.get("shares_after", ""))
        combined_after = parse_int_token(bucket_text.get("combined_after", ""))
        pct_after = parse_pct_token(bucket_text.get("pct_after", ""))

        # Tail parsing recovers values when column boundaries drift.
        tail = extract_numeric_tail(all_txt)

        if pct_before is None and tail.get("pct_before") is not None:
            pct_before = tail.get("pct_before")
        if pct_after is None and tail.get("pct_after") is not None:
            pct_after = tail.get("pct_after")

        # Small band-derived values are often postal codes from a shifted address.
        def _prefer_tail_int(cur, t):
            if t is None:
                return cur
            if cur is None:
                return t
            if cur < 10000 and t >= 10000:
                return t
            return cur

        shares_before = _prefer_tail_int(shares_before, tail.get("shares_before"))
        combined_before = _prefer_tail_int(combined_before, tail.get("combined_before"))
        shares_after = _prefer_tail_int(shares_after, tail.get("shares_after"))
        combined_after = _prefer_tail_int(combined_after, tail.get("combined_after"))

        if shares_before is None and tail.get("shares_before") is not None:
            shares_before = tail.get("shares_before")
        if combined_before is None and tail.get("combined_before") is not None:
            combined_before = tail.get("combined_before")
        if shares_after is None and tail.get("shares_after") is not None:
            shares_after = tail.get("shares_after")
        if combined_after is None and tail.get("combined_after") is not None:
            combined_after = tail.get("combined_after")

        warnings: List[str] = []
        if not t:
            warnings.append("missing_ticker")
        if pct_after is not None and not (0 <= pct_after <= 100):
            warnings.append("pct_after_out_of_range")
        if pct_before is not None and not (0 <= pct_before <= 100):
            warnings.append("pct_before_out_of_range")
        if pct_after is None:
            warnings.append("missing_pct_after")
        if pct_before is None:
            warnings.append("missing_pct_before")

        confidence = 1.0
        if not bucket_text.get("broker"):
            confidence -= 0.15
        if not bucket_text.get("shareholder"):
            confidence -= 0.15
        if shares_after is None or shares_before is None:
            confidence -= 0.2
        if pct_after is None or pct_before is None:
            confidence -= 0.2
        confidence = max(0.0, min(1.0, confidence))

        current = {
            "source_pdf": source_file,
            "source_file": source_file,
            "disclosure_datetime": document_info.get("disclosure_datetime", ""),
            "announcement_date": document_info.get("announcement_date", ""),
            "report_date": document_info.get("report_date", ""),
            "previous_report_date": document_info.get("previous_report_date", ""),
            "attachment_url": document_info.get("attachment_url", ""),
            "page": page_no,
            "row_no": row_no_val,
            "ticker": t,
            "emiten": bucket_text.get("emiten", ""),
            "broker": bucket_text.get("broker", ""),
            "shareholder": bucket_text.get("shareholder", ""),
            "account_name": bucket_text.get("account_name", ""),
            "address": bucket_text.get("address", ""),
            "address2": bucket_text.get("address2", ""),
            "nationality": bucket_text.get("nationality", ""),
            "domicile": bucket_text.get("domicile", ""),
            "status": status,
            "shares_after": shares_after,
            "combined_after": combined_after,
            "pct_after": pct_after,
            "shares_before": shares_before,
            "combined_before": combined_before,
            "pct_before": pct_before,
            "delta_shares_or_delta_field": _norm(bucket_text.get("delta", "")),
            "is_highlighted_change": bool(row_highlighted_columns),
            "highlighted_columns": row_highlighted_columns,
            "highlighted_text": row_highlighted_text,
            "raw_text": all_txt,
            "parse_warnings": ";".join(warnings),
            "confidence": confidence,
        }

        # Some layouts expose the delta only in the combined numeric tail.
        if not current.get("delta_shares_or_delta_field"):
            tail = extract_numeric_tail(all_txt)
            if tail.get("delta_shares_or_delta_field"):
                current["delta_shares_or_delta_field"] = tail.get("delta_shares_or_delta_field")

        prev_ticker = t or prev_ticker

    if current:
        parsed.append(current)

    for r in parsed:
        addr = r.get("address", "")
        addr2 = r.pop("address2", "")
        r["address"] = _norm(" ".join([a for a in [addr, addr2] if a]).strip())
        compute_change_fields(r)

    return raw_rows_debug, parsed


def build_argparser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Extract KSEI/BEI ownership table into a clean Excel")
    p.add_argument("--pdf", required=True, help="Path to PDF")
    p.add_argument("--out", default="outputs/extracted/ownership_table.xlsx", help="Output Excel path")
    p.add_argument("--manifest", default="", help="Optional scraper manifest JSON for disclosure metadata")
    p.add_argument("--debug-dir", default="", help="If set, write debug artifacts to this folder")
    p.add_argument(
        "--include-raw-debug",
        action="store_true",
        help="Also include the raw clustered-row sheet inside the output Excel (can be large/slow).",
    )
    p.add_argument("--max-pages", type=int, default=0, help="If >0, only process first N pages")
    p.add_argument("--page-from", type=int, default=0, help="1-based start page to process (0=auto)")
    p.add_argument("--page-to", type=int, default=0, help="1-based end page to process (0=auto)")
    return p


def _empty_frames() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()


def extract_pdf_to_frames(
    pdf_path: Path,
    manifest_path: str | Path | None = None,
    debug_dir: Path | None = None,
    include_raw_debug: bool = False,
    max_pages: int = 0,
    page_from: int = 0,
    page_to: int = 0,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    raw_debug_all: List[Dict[str, Any]] = []
    rows_all: List[Dict[str, Any]] = []
    manifest_meta = load_manifest_metadata(manifest_path, pdf_path)

    with pdfplumber.open(str(pdf_path)) as pdf:
        # Front matter varies, so locate the first page containing the table header.
        table_page_idx = None
        first_pages_text: list[str] = []
        for i, p in enumerate(pdf.pages):
            t = (p.extract_text() or "")
            if i < 3:
                first_pages_text.append(t)
            if "NoKode" in t or "NoKode Efek" in t:
                table_page_idx = i
                break
        if table_page_idx is None:
            raise RuntimeError("Could not locate table header in PDF")

        header_text = (pdf.pages[table_page_idx].extract_text() or "")
        document_dates = extract_report_dates_from_text("\n".join(first_pages_text + [header_text]))
        document_info = {
            "source_pdf": pdf_path.name,
            "disclosure_datetime": manifest_meta.get("disclosure_datetime", ""),
            "announcement_date": manifest_meta.get("announcement_date", ""),
            "attachment_url": manifest_meta.get("attachment_url", ""),
            **document_dates,
        }
        if not document_info["announcement_date"]:
            document_info["announcement_date"] = manifest_meta.get("filename_date", "") or filename_date(pdf_path.name)

        bands = infer_bands_from_header(pdf.pages[table_page_idx])

        print("[header layout] column cutpoints:")
        for name, x0 in bands.cuts:
            print(f"  {name:16s} x0={x0:.1f}")

        max_pages = max_pages if max_pages and max_pages > 0 else len(pdf.pages)
        auto_start = table_page_idx + 1
        start_page = page_from if page_from and page_from > 0 else auto_start
        end_page = page_to if page_to and page_to > 0 else min(len(pdf.pages), max_pages)

        start_idx = max(0, start_page - 1)
        end_idx_excl = min(len(pdf.pages), end_page)

        for pidx in range(start_idx, end_idx_excl):
            if (pidx - table_page_idx) % 5 == 0:
                print(f"[progress] processing page {pidx+1}/{end_page}")
            page = pdf.pages[pidx]
            raw_dbg, parsed = extract_table_from_page(
                page,
                bands=bands,
                page_no=pidx + 1,
                source_file=str(pdf_path.name),
                document_info=document_info,
                debug_words_dir=(debug_dir / "words") if debug_dir else None,
            )
            if include_raw_debug:
                raw_debug_all.extend(raw_dbg)
            rows_all.extend(parsed)

    print(f"[progress] finished page parsing, total parsed rows so far: {len(rows_all)}")

    df = pd.DataFrame(rows_all)
    # Keep a stable output schema even when no records provide an optional field.
    wanted = [
        "source_pdf",
        "source_file",
        "disclosure_datetime",
        "announcement_date",
        "report_date",
        "previous_report_date",
        "attachment_url",
        "page",
        "row_no",
        "ticker",
        "emiten",
        "broker",
        "shareholder",
        "account_name",
        "address",
        "nationality",
        "domicile",
        "status",
        "shares_after",
        "combined_after",
        "pct_after",
        "shares_before",
        "combined_before",
        "pct_before",
        "delta_shares_or_delta_field",
        "delta_shares",
        "delta_pct",
        "is_highlighted_change",
        "has_numeric_change",
        "change_reason",
        "highlighted_columns",
        "highlighted_text",
        "raw_text",
        "parse_warnings",
        "confidence",
    ]
    for c in wanted:
        if c not in df.columns:
            df[c] = "" if c in {
                "source_pdf",
                "source_file",
                "disclosure_datetime",
                "announcement_date",
                "report_date",
                "previous_report_date",
                "attachment_url",
                "raw_text",
                "parse_warnings",
                "highlighted_columns",
                "highlighted_text",
                "change_reason",
            } else None
    df = df[wanted]

    for col in ["page", "shares_after", "combined_after", "shares_before", "combined_before", "delta_shares"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").astype("Int64")
    for col in ["pct_after", "pct_before", "delta_pct", "confidence"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")
    for col in ["is_highlighted_change", "has_numeric_change"]:
        df[col] = df[col].fillna(False).astype(bool)

    n_rows = len(df)
    pct_bad = int(((df["pct_after"].notna()) & ((df["pct_after"] < 0) | (df["pct_after"] > 100))).sum()) if n_rows else 0
    pct_bad2 = int(((df["pct_before"].notna()) & ((df["pct_before"] < 0) | (df["pct_before"] > 100))).sum()) if n_rows else 0
    print(f"[summary] extracted rows: {n_rows}")
    print(f"[summary] pct_after out-of-range: {pct_bad} | pct_before out-of-range: {pct_bad2}")

    changes = df[(df["is_highlighted_change"]) | (df["has_numeric_change"])].copy() if n_rows else pd.DataFrame(columns=df.columns)
    warnings = df[df["parse_warnings"].fillna("").astype(str).str.len() > 0].copy() if n_rows else pd.DataFrame(columns=df.columns)
    raw_debug = pd.DataFrame(raw_debug_all)
    return df, changes, warnings, raw_debug


def manifest_downloads_frame(manifest_path: str | Path | None, pdf_path: Path) -> pd.DataFrame:
    if manifest_path:
        path = Path(manifest_path)
        if path.exists():
            try:
                data = json.loads(path.read_text(encoding="utf-8"))
                if isinstance(data, list):
                    return pd.DataFrame(data)
            except json.JSONDecodeError:
                pass
    return pd.DataFrame(
        [
            {
                "attachment_filename": pdf_path.name,
                "local_pdf_path": str(pdf_path),
                "download_status": "local",
            }
        ]
    )


def write_workbook(
    out_path: Path,
    ownership_all: pd.DataFrame,
    ownership_changes: pd.DataFrame,
    downloads: pd.DataFrame,
    parse_warnings: pd.DataFrame,
    raw_debug: pd.DataFrame | None = None,
) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out_path, engine="openpyxl") as xw:
        ownership_changes.to_excel(xw, index=False, sheet_name="ownership_changes")
        ownership_all.to_excel(xw, index=False, sheet_name="ownership_all")
        downloads.to_excel(xw, index=False, sheet_name="downloads")
        parse_warnings.to_excel(xw, index=False, sheet_name="parse_warnings")
        if raw_debug is not None and not raw_debug.empty:
            raw_debug.to_excel(xw, index=False, sheet_name="raw_rows_debug")


def main() -> None:
    args = build_argparser().parse_args()
    pdf_path = Path(args.pdf)
    out_path = Path(args.out)
    debug_dir = Path(args.debug_dir) if args.debug_dir else None

    ownership_all, ownership_changes, parse_warnings, raw_debug = extract_pdf_to_frames(
        pdf_path,
        manifest_path=args.manifest,
        debug_dir=debug_dir,
        include_raw_debug=args.include_raw_debug,
        max_pages=args.max_pages,
        page_from=args.page_from,
        page_to=args.page_to,
    )

    downloads = manifest_downloads_frame(args.manifest, pdf_path)
    write_workbook(
        out_path,
        ownership_all,
        ownership_changes,
        downloads,
        parse_warnings,
        raw_debug=raw_debug if args.include_raw_debug else None,
    )
    print(f"[done] wrote: {out_path}")


if __name__ == "__main__":
    main()
