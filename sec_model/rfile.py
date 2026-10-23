"""Parse one SEC R-file page (``R<n>.htm``) into a ParsedStatement.

Columns are keyed by period end date and duration; rows by XBRL tag. Values
are converted to millions (money), millions of shares/units (shares) or left
as-is (per-share and other).
"""
import math
import re
from datetime import date
from decimal import Decimal, InvalidOperation
from typing import Callable

from bs4 import BeautifulSoup

from .types import Column, ParsedStatement, Row

_MONTHS = {m: i for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"], 1)}
_DATE = re.compile(r"\b([A-Za-z]{3})[a-z]*\.?\s+(\d{1,2}),\s*(\d{4})")
_DURATION = re.compile(r"(\d+)\s+Months?\s+Ended", re.I)
_DEFREF = re.compile(r"defref_([\w\-]+)(?:=([\w\-]+))?")
_FOOTNOTE = re.compile(r"\[\d+\]")
_PAD_EM = re.compile(r"padding-left:\s*([\d.]+)\s*em", re.I)
_SCALES = {"Millions": 1.0, "Thousands": 0.001, "Billions": 1000.0}
_EMPTY = {"", "—", "–", "-"}
# Currency: "INCOME - CAD ($)" in the title or "Dec. 31, 2022 CAD ($)" in a column header;
# some titles carry only the scale ("... - CAD CAD in Millions").
_CUR_PAREN = re.compile(r"\b([A-Z]{3}) \(\$\)")
_CUR_SCALE = re.compile(r"\b([A-Z]{3}) in (?:Millions|Thousands|Billions)\b")
_MONEY_SCALE = r"(?:\$|\b[A-Z]{3}) in (Millions|Thousands|Billions)"

_PERUNIT_TAG = re.compile(r"PerShare|PerUnit")
_PERUNIT_LABEL = re.compile(r"per\s+(share|unit)", re.I)
_SHARES_TAG = re.compile(r"WeightedAverage")
_SHARES_WORD = re.compile(r"\b(shares|units)\b", re.I)
_SHARES_QUAL = re.compile(r"\b(weighted|outstanding|issued)\b", re.I)
_PARENS = re.compile(r"\([^)]*\)")
_DATA_TYPES = {"monetaryItemType": "money", "sharesItemType": "shares",
               "perShareItemType": "perunit", "perUnitItemType": "perunit",
               "pureItemType": "other", "percentItemType": "other", "stringItemType": "other"}


def parse_rfile(html: str, stype: str, source_url: str,
                unit_hint: Callable[[str], str | None] | None = None) -> ParsedStatement:
    soup = BeautifulSoup(html, "html.parser")
    table = soup.find("table", class_="report")
    if table is None:
        raise ValueError(f"no report table in {source_url}")
    trs = table.find_all("tr", recursive=False)
    n_header = sum(1 for _ in _leading_header_rows(trs))
    title, label_cols, columns, keep = _parse_header(trs[:n_header])

    currency = _currency(title, columns)
    money_scale = _SCALES.get(_search(_MONEY_SCALE, title, ""), 1e-6)
    share_scale = _SCALES.get(_search(r"shares in (Millions|Thousands)", title, ""), 1e-6)
    scales = {"money": money_scale, "shares": share_scale}
    page_types = _page_data_types(soup)

    rows: list[Row] = []
    used: set[str] = set()
    member = None            # set once a dimension (axis=member) header row is seen
    seen_abstract = False
    for tr in trs[n_header:]:
        cells = tr.find_all(["td", "th"], recursive=False)
        anchor = cells[0].find("a", onclick=_DEFREF) if cells else None
        if anchor is None:
            continue         # footnote text rows, spacer rows
        tag, mem = _DEFREF.search(anchor["onclick"]).groups()
        label = " ".join(_FOOTNOTE.sub(" ", cells[0].get_text(" ")).split())
        if mem:              # dimension header: "Commodity sales" for Axis=Member
            member, seen_abstract = mem, False
            key = f"{tag}={mem}"
            rows.append(Row(key, key, label, 0, True, "other", [None] * len(columns)))
            used.add(key)
            continue
        is_abstract = tag.endswith("Abstract")
        base = 1 if member else 0
        explicit = _indent_of(cells[0])
        if explicit is not None:
            indent = explicit
        else:
            indent = base if is_abstract else base + (1 if seen_abstract else 0)
        seen_abstract = seen_abstract or is_abstract
        unit = _unit_of(tag, label, unit_hint, page_types.get(tag))
        scale = scales.get(unit, 1.0)
        raw = cells[label_cols:]
        values = [_parse_value(raw[i].get_text(" ") if i < len(raw) else "", scale) for i in keep]
        key = _unique_key(tag, label, member, used)
        rows.append(Row(key, tag, label, indent, is_abstract, unit, values))

    return ParsedStatement(stype=stype, title=title, currency=currency, money_scale=money_scale,
                           share_scale=share_scale, columns=columns, rows=rows, source_url=source_url)


def _leading_header_rows(trs):
    for tr in trs:
        if tr.find("th", recursive=False) is None:
            return
        yield tr


def _parse_header(header_rows) -> tuple[str, int, list[Column], list[int]]:
    """Return (title, number of label cells per data row, columns, data-cell indexes kept)."""
    if not header_rows:
        raise ValueError("R-file has no header row")
    first = header_rows[0].find_all("th", recursive=False)
    title_cell = first[0]
    title = " ".join(title_cell.get_text(" ").split())
    label_cols = int(title_cell.get("colspan", 1))
    if len(header_rows) >= 2:     # duration groups in row 1, dates in row 2
        groups = _expand(first[1:])
        dates = _expand(header_rows[1].find_all("th", recursive=False))
    else:                          # instant statement: dates in row 1
        groups, dates = [], _expand(first[1:])
    columns, keep = [], []
    for i, date_text in enumerate(dates):
        end = _parse_date(date_text)
        if end is None:
            continue
        group = groups[i] if i < len(groups) else ""
        m = _DURATION.search(group)
        columns.append(Column(end=end, months=int(m.group(1)) if m else None,
                              label=f"{group} {date_text}".strip()))
        keep.append(i)
    return title, label_cols, columns, keep


def _expand(ths) -> list[str]:
    out = []
    for th in ths:
        out.extend([" ".join(th.get_text(" ").split())] * int(th.get("colspan", 1)))
    return out


def _parse_date(text: str) -> date | None:
    m = _DATE.search(text)
    if not m or m.group(1).lower() not in _MONTHS:
        return None
    try:
        return date(int(m.group(3)), _MONTHS[m.group(1).lower()], int(m.group(2)))
    except ValueError:
        return None


def _parse_value(text: str, scale: float) -> float | None:
    s = _FOOTNOTE.sub("", text).replace("$", "").replace(",", "")
    s = "".join(s.split())
    if s in _EMPTY:
        return None
    neg = s.startswith("(") and s.endswith(")")
    if neg:
        s = s[1:-1]
    try:
        d = Decimal(s)
    except InvalidOperation:
        return None
    if not d.is_finite():
        return None
    v = float((-d if neg else d) * Decimal(repr(scale)))
    return v if math.isfinite(v) else None   # "1e400" overflows to inf


def _indent_of(cell) -> int | None:
    for el in [cell, *cell.find_all(["div", "span"])]:
        m = _PAD_EM.search(el.get("style", ""))
        if m:
            return round(float(m.group(1)))
    nested = cell.find_all(["div", "span"], class_="pl")
    return len(nested) if nested else None


def _page_data_types(soup) -> dict[str, str]:
    """tag -> unit from the page's hidden authRefData tables ("Data Type: xbrli:monetaryItemType")."""
    out = {}
    for t in soup.find_all("table", class_="authRefData"):
        strong = t.find("strong", string=re.compile("Data Type"))
        label_td = strong.find_parent("td") if strong else None
        td = label_td.find_next_sibling("td") if label_td else None   # malformed block: skip
        unit = _DATA_TYPES.get(td.get_text().strip().split(":")[-1]) if td else None
        if unit and t.get("id", "").startswith("defref_"):
            out[t["id"][len("defref_"):]] = unit
    return out


def _unit_of(tag: str, label: str, unit_hint, page_type: str | None) -> str:
    """unit_hint (companyfacts) wins, then the page's XBRL data type, then tag/label rules."""
    hint = unit_hint(tag) if unit_hint else None
    if hint is not None:
        return hint
    if page_type is not None:
        return page_type
    if _PERUNIT_TAG.search(tag) or _PERUNIT_LABEL.search(label):
        return "perunit"
    bare = _PARENS.sub("", label)    # "(2,161,760,683 units issued ...)" describes a $ line
    if _SHARES_TAG.search(tag) or (_SHARES_WORD.search(bare) and _SHARES_QUAL.search(bare)):
        return "shares"
    return "money"


def _unique_key(tag: str, label: str, member: str | None, used: set[str]) -> str:
    if member:
        key = f"{tag}|{label}|{member}"
    elif tag not in used:
        key = tag
    else:
        key = f"{tag}|{label}"
    base, n = key, 2
    while key in used:
        key, n = f"{base}|{n}", n + 1
    used.add(key)
    return key


def _currency(title: str, columns: list[Column]) -> str:
    """A currency code in a kept column header wins (the title then carries only "$ in
    Millions"), then "XXX ($)" in the title, then "XXX in Millions", else USD."""
    for c in columns:
        m = _CUR_PAREN.search(c.label)
        if m:
            return m.group(1)
    for pattern in (_CUR_PAREN, _CUR_SCALE):
        m = pattern.search(title)
        if m:
            return m.group(1)
    return "USD"


def _search(pattern: str, text: str, default: str) -> str:
    m = re.search(pattern, text)
    return m.group(1) if m else default
