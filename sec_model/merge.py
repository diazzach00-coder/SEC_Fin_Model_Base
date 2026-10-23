"""Stack the same statement type from several 10-K filings into one MergedStatement.

One row per key, one value per fiscal year; for each year the newest filing
that reports it supplies the value (so restatements win).
"""
from datetime import date

from .types import (Filing, MergedRow, MergedStatement, ParsedStatement,
                    SourceRec, UserFacingError)


def _column_index(parsed: ParsedStatement, end: date) -> int | None:
    """Index of the column for a fiscal year ending ``end``, or None.

    Instant statements (BS) have months=None; duration statements (IS/CF)
    must be a full 12-month period -- a None duration means the parser found
    no duration header, so it is not trusted.
    """
    want = None if parsed.stype == "BS" else 12
    for i, col in enumerate(parsed.columns):
        if col.end == end and col.months == want:
            return i
    return None


def _insert_dropped_keys(order: list[str], older: ParsedStatement) -> None:
    """Place keys of ``older`` missing from ``order`` right after their predecessor."""
    prev = None
    for r in older.rows:
        if r.key not in order:
            if prev is not None and prev in order:
                order.insert(order.index(prev) + 1, r.key)
            else:
                order.append(r.key)
        prev = r.key


def merge_statements(parsed: list[tuple[Filing, ParsedStatement]],
                     fy_labels: list[str],
                     fy_ends: dict[str, date]) -> MergedStatement:
    if not parsed:
        raise UserFacingError("No parsed statements to merge.")
    newest_first = sorted(parsed, key=lambda fp: (fp[0].filing_date, fp[0].accession),
                          reverse=True)
    newest = newest_first[0][1]

    for _, ps in reversed(newest_first):
        if ps.currency != newest.currency:
            raise UserFacingError(
                f"Currency changes across filings ({ps.currency}→{newest.currency}); not supported.")

    # Row order and per-key metadata (newest filing containing the key wins).
    order = [r.key for r in newest.rows]
    for _, ps in newest_first[1:]:
        _insert_dropped_keys(order, ps)
    meta = {}
    for _, ps in newest_first:
        for r in ps.rows:
            meta.setdefault(r.key, r)
    newest_keys = {r.key for r in newest.rows}

    # Per (fiscal year, key): the newest filing that has a column for the year
    # AND contains the key supplies the value (even if the cell is None).
    # sources[fy] is the newest filing with a column for the year.
    values: dict[str, dict[str, float | None]] = {k: {} for k in order}
    sources: dict[str, SourceRec] = {}
    for fy in fy_labels:
        candidates = []
        end = fy_ends.get(fy)
        if end is not None:
            for filing, ps in newest_first:
                idx = _column_index(ps, end)
                if idx is not None:
                    cells = {r.key: (r.values[idx] if idx < len(r.values) else None)
                             for r in ps.rows}
                    candidates.append((filing, ps, cells))
        if candidates:
            filing, ps, _ = candidates[0]
            sources[fy] = SourceRec(accession=filing.accession, form=filing.form,
                                    filing_date=filing.filing_date, r_url=ps.source_url,
                                    scale_text=ps.title)
        for k in order:
            values[k][fy] = next((cells[k] for _, _, cells in candidates if k in cells), None)

    rows = [MergedRow(key=k, tag=meta[k].tag, label=meta[k].label, indent=meta[k].indent,
                      is_abstract=meta[k].is_abstract, unit=meta[k].unit,
                      values=values[k], still_reported=k in newest_keys)
            for k in order]
    return MergedStatement(stype=newest.stype, currency=newest.currency,
                           fy_labels=list(fy_labels), fy_ends=dict(fy_ends),
                           rows=rows, sources=sources)
