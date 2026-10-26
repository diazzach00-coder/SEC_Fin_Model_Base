"""Drivers and Revenue Streams sheets: the analyst's input surfaces behind the Model sheet.

Both sheets use the Model's column letters for the same fiscal year (history from D,
forecast right after), so `Driver(key)` renders as ``Drivers!<same column><row>`` and
`Block(key)` as the Revenue Streams total cell in the same column. History columns are
left empty here.

Year-1 driver cells and Expr scalar defaults are formulas, not pasted numbers (spec 7.2:
nothing is a magic number): each is ``to_excel(default, ctx)`` with ``ctx`` the Model's
last-history-column context, sheet-qualified (`model_sheet.render_context(...,
qualified=True)`), so `Hist` renders ``AVERAGE(Model!$..$)`` over the Model ratio row and
`Line` the Model's last history cell. They keep the input style (blue on yellow) because
the analyst may overwrite them. Constant defaults (OTHER_OP, JV_TOGGLE, ...) are values.
"""
from openpyxl.styles import Alignment, Font
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName

from sec_model.formula import Const, Expr, RenderContext, to_excel

from .asreported import unit_note
from . import styles as S

DRIVERS_SHEET = "Drivers"
STREAMS_SHEET = "Revenue Streams"
N_STREAMS = 25

NUMBER_FORMATS = {"money": S.FMT_MONEY, "pct": S.FMT_PCT, "units": S.FMT_UNITS,
                  "perunit": S.FMT_PERUNIT, "days": S.FMT_DAYS, "multiple": S.FMT_X}
_UNITS = {"pct": "%", "days": "days", "multiple": "x"}
# 0/1 switches: the cell gets a 1/0 dropdown so a typo cannot silently change the model.
TOGGLE_SCALARS = ("JV_TOGGLE", "WC_HOLD")
FONT_NOTE = Font(name=S.FONT_NAME, size=9, italic=True, color="808080")
FONT_FORMULA_BOLD = Font(name=S.FONT_NAME, size=10, bold=True, color="000000")


def unit_for(fmt: str, currency: str, entity_type: str = "corporation") -> str:
    """Unit-column text for a Model / Drivers format: money, share counts and per-share
    data share `asreported.unit_note` (the Model's "units" fmt is its "shares" unit)."""
    if fmt in ("money", "units", "perunit"):
        return unit_note("shares" if fmt == "units" else fmt, currency, entity_type)
    return _UNITS.get(fmt, "")


def frame(ws, title: str, subtitle: str, first_labels: list[str], fy_labels, fcst_labels,
          extra: dict[int, str] | None = None) -> int:
    """Title rows 2-3, header row 4 (B.. labels, then years from D); freeze D5.
    ``extra``: {column: header} written after the year columns. Returns the last year column."""
    last_col = S.DSC + len(fy_labels) + len(fcst_labels) - 1
    bar_end = max([last_col] + list((extra or {}).keys()))
    S.title_bar(ws, 2, title, bar_end, subtitle=subtitle)
    S.header_row(ws, 4, first_labels, 2)
    S.header_row(ws, 4, list(fy_labels) + list(fcst_labels), S.DSC,
                 [False] * len(fy_labels) + [True] * len(fcst_labels))
    for c, text in (extra or {}).items():
        S.header_row(ws, 4, [text], c)
        ws.cell(4, c).alignment = Alignment(horizontal="left")
    S.set_widths(ws, {"A": 2})
    for c in range(S.DSC, last_col + 1):
        S.set_widths(ws, {c: 12})
    ws.freeze_panes = "D5"
    return last_col


def section_header(ws, row: int, text: str, last_col: int) -> None:
    ws.cell(row, 2, text).font = S.FONT_BOLD
    for c in range(2, last_col + 1):
        ws.cell(row, c).fill = S.FILL_SECTION


def add_toggle_validation(ws, coord: str) -> None:
    """Restrict ``coord`` on ``ws`` to 1 or 0 (in-cell dropdown, no blank)."""
    dv = DataValidation(type="list", formula1='"1,0"', allow_blank=False)
    dv.error, dv.errorTitle = "Enter 1 (on) or 0 (off).", "Switch"
    dv.add(coord)
    ws.add_data_validation(dv)


def _input(cell, value, fmt: str | None, formula: bool = False) -> None:
    """Yellow input cell: blue font for a typed value, black for a formula."""
    cell.value = value
    cell.font = S.FONT_FORMULA if formula else S.FONT_INPUT
    cell.fill = S.FILL_INPUT
    if fmt is not None:
        cell.number_format = NUMBER_FORMATS.get(fmt, S.FMT_MONEY)


def plan_drivers(spec) -> tuple[dict, dict]:
    """(driver_row {key: row}, scalar_cell {key: "Drivers!$C$<row>"}) as `write_drivers`
    lays them out, so the Model layout exists before the sheet is written."""
    driver_row = {d.key: 5 + i for i, d in enumerate(spec.drivers)}
    first = 5 + len(spec.drivers) + 2
    scalar_cell = {s.key: f"{DRIVERS_SHEET}!$C${first + i}" for i, s in enumerate(spec.scalars)}
    return driver_row, scalar_cell


def _default_cell(default, ctx: RenderContext, value):
    """A Const default is written as its value; any other Expr as its formula."""
    if isinstance(default, Const):
        return default.value
    if isinstance(default, Expr):
        return to_excel(default, ctx)
    return value


def write_drivers(wb, ws, spec, fy_labels, fcst_labels, driver_defaults: dict,
                  scalar_defaults: dict, currency: str,
                  ctx: RenderContext) -> tuple[dict, dict]:
    """One row per driver (year 1 = the default as a formula over the Model, later years
    ``=prev``), then scalars. ``ctx`` renders defaults in the last history column (see the
    module docstring); ``driver_defaults`` / ``scalar_defaults`` supply the values of
    non-Expr defaults. Returns (driver_row {key: row}, scalar_cell {key:
    "Drivers!$C$<row>"}), equal to `plan_drivers(spec)`."""
    note_col = S.DSC + len(fy_labels) + len(fcst_labels)
    last_col = frame(ws, "Drivers",
                     "Yellow = input. Year 1 defaults come from the company's own history; "
                     "later years copy the prior year - overwrite any year.",
                     ["Driver", "Unit"], fy_labels, fcst_labels, {note_col: "Note"})
    S.set_widths(ws, {"B": 46, "C": 12, note_col: 60})
    first_fcst = S.DSC + len(fy_labels)
    driver_row: dict[str, int] = {}
    r = 5
    for d in spec.drivers:
        driver_row[d.key] = r
        ws.cell(r, 2, d.label).font = S.FONT_BASE
        ws.cell(r, 3, unit_for(d.fmt, currency, spec.entity_type)).font = S.FONT_BASE
        for c in range(first_fcst, last_col + 1):
            if c == first_fcst:
                _input(ws.cell(r, c), _default_cell(d.default, ctx, driver_defaults.get(d.key)),
                       d.fmt)
            else:
                _input(ws.cell(r, c), f"={S.col(c - 1)}{r}", d.fmt, formula=True)
        ws.cell(r, note_col, d.note).font = FONT_NOTE
        r += 1

    r += 1
    section_header(ws, r, "MODEL SWITCHES & SCALARS", last_col)
    scalar_cell: dict[str, str] = {}
    for s in spec.scalars:
        r += 1
        ref = f"{DRIVERS_SHEET}!$C${r}"
        scalar_cell[s.key] = ref
        ws.cell(r, 2, s.label).font = S.FONT_BASE
        _input(ws.cell(r, 3), _default_cell(s.default, ctx, scalar_defaults.get(s.key)), s.fmt)
        if s.key in TOGGLE_SCALARS:
            add_toggle_validation(ws, f"C{r}")
        ws.cell(r, 4, s.note).font = FONT_NOTE
        wb.defined_names[s.key] = DefinedName(s.key, attr_text=ref)
    if (driver_row, scalar_cell) != plan_drivers(spec):
        raise RuntimeError("plan_drivers is out of sync with write_drivers")
    return driver_row, scalar_cell


def _stream_block(ws, top: int, title: str, total_label: str, first_fcst: int, last_col: int,
                  consolidated) -> int:
    """Section header at ``top``, 25 stream rows, totals row. ``consolidated(c)`` is the
    row-1 formula for column c. Returns the totals row."""
    section_header(ws, top, title, last_col)
    first = top + 1
    for i in range(N_STREAMS):
        r = first + i
        name = ("Total company (placeholder)" if i == 0
                else "Other / unallocated" if i == N_STREAMS - 1 else None)
        _input(ws.cell(r, 2), name, None)
        _input(ws.cell(r, 3), "All segments" if i == 0 else None, None)
        for c in range(first_fcst, last_col + 1):
            _input(ws.cell(r, c), consolidated(c, r) if i == 0 else None, "money",
                   formula=i == 0)
    total = first + N_STREAMS
    ws.cell(total, 2, total_label).font = S.FONT_BOLD
    for c in range(2, last_col + 1):
        cell = ws.cell(total, c)
        cell.fill = S.FILL_SUBTOTAL
        if c >= first_fcst:
            L = S.col(c)
            cell.value = f"=SUM({L}{first}:{L}{total - 1})"
            cell.font = FONT_FORMULA_BOLD
            cell.number_format = S.FMT_MONEY
            cell.border = S.BORDER_TOP
    return total


def write_streams(ws, fy_labels, fcst_labels, model_sheet: str, model_row: dict,
                  driver_row: dict) -> dict[str, int]:
    """Revenue and operating-cost blocks of 25 streams, totals and reconciliation.

    Row 1 of each block is a ``Total company (placeholder)`` row: revenue = last reported
    revenue grown each year by that year's REV_GROWTH (chained, as `forecast.compute` does),
    cost = revenue x COST_PCT. Returns {"ASSET_REV": total row, "ASSET_COST": total row}."""
    last_col = frame(ws, STREAMS_SHEET,
                     "Enter revenue and operating cost by stream (segment, product line, "
                     "region, asset \u2026) for each forecast year. Row 1 grows the "
                     "consolidated figures off the Drivers sheet; overwrite it with real "
                     "streams.",
                     ["Stream", "Segment"], fy_labels, fcst_labels)
    S.set_widths(ws, {"B": 40, "C": 18})
    first_fcst = S.DSC + len(fy_labels)
    last_hist = S.col(first_fcst - 1)
    g_row, cost_row = driver_row["REV_GROWTH"], driver_row["COST_PCT"]

    def rev_formula(c, r):
        base = (f"N({model_sheet}!{last_hist}{model_row['REV']})" if c == first_fcst
                else f"{S.col(c - 1)}{r}")
        return f"={base}*(1+{DRIVERS_SHEET}!{S.col(c)}{g_row})"

    rev_total = _stream_block(ws, 5, "REVENUE BY STREAM", "Total revenue", first_fcst, last_col,
                              rev_formula)
    rev_first = rev_total - N_STREAMS
    cost_total = _stream_block(
        ws, rev_total + 2, "OPERATING COST BY STREAM", "Total operating cost", first_fcst,
        last_col, lambda c, r: f"={S.col(c)}{rev_first}*{DRIVERS_SHEET}!{S.col(c)}{cost_row}")

    r = cost_total + 2
    section_header(ws, r, "RECONCILIATION", last_col)
    for label, key in (("Last reported revenue (Model)", "REV"),
                       ("Last reported operating cost (Model)", "OPCOST")):
        r += 1
        ws.cell(r, 2, label).font = S.FONT_BASE
        cell = ws.cell(r, first_fcst - 1, f"={model_sheet}!{last_hist}{model_row[key]}")
        cell.font = S.FONT_FORMULA
        cell.number_format = S.FMT_MONEY
    r += 1
    ws.cell(r, 2, "EBITDA before equity earnings").font = S.FONT_BOLD
    for c in range(first_fcst, last_col + 1):
        L = S.col(c)
        cell = ws.cell(r, c, f"={L}{rev_total}-{L}{cost_total}")
        cell.font = FONT_FORMULA_BOLD
        cell.number_format = S.FMT_MONEY
        cell.fill = S.FILL_SUBTOTAL
        cell.border = S.BORDER_TOP
    return {"ASSET_REV": rev_total, "ASSET_COST": cost_total}
