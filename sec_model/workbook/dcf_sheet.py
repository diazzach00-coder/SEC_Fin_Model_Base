"""DCF sheet: UFCF build from the Model, discounting, terminal value both ways,
equity bridge, multiples, value-per-unit sensitivity grids and a football field.

Mirrors the DCF tab of ``DCFWacc Templates.xlsx``. Sensitivities are ordinary formulas
(no Excel Data Tables): each grid cell recomputes value per unit in closed form from the
helper rows ``t`` (years from the valuation date) and stub-adjusted UFCF:
``(SUMPRODUCT(UFCF, 1/(1+w)^t) + TV(w, g|m)/(1+w)^t_N - net debt) / units``.
Price, units and net debt come from the WACC sheet (one source for both sheets).
"""
from dataclasses import dataclass, field
from datetime import date

from openpyxl.chart import BarChart, Reference

from sec_model.types import MarketInputs
from . import styles as S
from .inputs_sheets import FONT_NOTE, _input, section_header
from .model_sheet import ModelLayout
from .wacc_sheet import (
    STEPS, WaccLayout, absref, currency_note, formula, headers, label, offset,
)

DCF_SHEET = "DCF"
FMT_DATE = "mmm d, yyyy"
FC = 5                              # first forecast column (E); D = last actual year
EXIT_COL = 8                        # exit-multiple method values (H); labels from E
WACC_C, PRICE_C, MID_C = "$H$6", "$H$7", "$H$8"
VAL_DATE, STUB = "$D$10", "$D$11"


@dataclass
class DcfLayout:
    header_row: int
    row: dict[str, int]                         # block key -> DCF row
    ev_cells: dict[str, str]                    # "perpetuity" | "exit" -> "DCF!$C$43"
    equity_value_cells: dict[str, str]
    value_per_unit_cells: dict[str, str]
    upside_cells: dict[str, str]
    growth_cell: str
    exit_multiple_cell: str
    grids: dict[str, dict] = field(default_factory=dict)   # name -> {"corner": (r, c), "range"}
    grid_centre_cells: dict[str, str] = field(default_factory=dict)
    football_rows: list[int] = field(default_factory=list)  # perpetuity, exit, 52-week
    first_fcst_col: int = FC
    last_fcst_col: int = FC
    sheet: str = DCF_SHEET


# label, key, number format, bold
_UFCF_ROWS = (("Revenue", "REV", S.FMT_MONEY, False),
              ("EBITDA", "EBITDA", S.FMT_MONEY, True),
              ("EBITDA margin", "MARGIN", S.FMT_PCT, False),
              ("EBIT", "EBIT", S.FMT_MONEY, False),
              ("Tax rate", "TAX_RATE", S.FMT_PCT, False),
              ("Tax on EBIT", "TAX", S.FMT_MONEY, False),
              ("NOPAT", "NOPAT", S.FMT_MONEY, True),
              ("Depreciation & amortization", "DA", S.FMT_MONEY, False),
              ("Change in working capital", "DWC", S.FMT_MONEY, False),
              ("Capital expenditures", "CAPEX", S.FMT_MONEY, False),
              ("Unlevered free cash flow (UFCF)", "UFCF", S.FMT_MONEY, True))
_MODEL_LINE = {"REV": "REV", "EBITDA": "EBITDA", "EBIT": "EBIT", "DA": "DA", "DWC": "CF_DWC",
               "CAPEX": "CAPEX"}


def _ufcf_cell(key: str, c: int, row: dict, model_col: str, layout: ModelLayout) -> str:
    L, r = S.col(c), row
    m = layout.sheet
    if key in _MODEL_LINE:
        return f"={m}!{model_col}{layout.line_row[_MODEL_LINE[key]]}"
    actual = c == FC - 1
    if key == "MARGIN":
        return (f"=IFERROR(N({L}{r['EBITDA']})/N({L}{r['REV']}),0)" if actual
                else f"=IFERROR({L}{r['EBITDA']}/{L}{r['REV']},0)")
    if key == "TAX_RATE":
        if actual:
            lr = layout.line_row
            return (f"=IFERROR(N({m}!{model_col}{lr['TAX']})/"
                    f"N({m}!{model_col}{lr['PRETAX']}),0)")
        return f"={layout.drivers_sheet}!{model_col}{layout.driver_row['TAX_RATE']}"
    ebit = f"N({L}{r['EBIT']})" if actual else f"{L}{r['EBIT']}"
    if key == "TAX":
        return f"={ebit}*{L}{r['TAX_RATE']}"
    if key == "NOPAT":
        return f"={ebit}-{L}{r['TAX']}"
    parts = [f"{L}{r['NOPAT']}"] + [f"N({L}{r[k]})" if actual else f"{L}{r[k]}"
                                    for k in ("DA", "DWC", "CAPEX")]
    return "=" + "+".join(parts)


def _header(ws, wacc: WaccLayout, market: MarketInputs, last_fye: date, company: str,
            ticker: str) -> None:
    for r, text, value in ((6, "Company name", company), (7, "Ticker", ticker)):
        label(ws, r, 2, text)
        label(ws, r, 4, value)
    label(ws, 8, 2, "Most recent fiscal year end")
    formula(ws, 8, 4, last_fye, FMT_DATE).font = S.FONT_INPUT
    label(ws, 9, 2, "End of first forecast year")
    formula(ws, 9, 4, "=EDATE(D8,12)", FMT_DATE)
    label(ws, 10, 2, "Valuation date")
    formula(ws, 10, 4, "=TODAY()", FMT_DATE)
    label(ws, 11, 2, "Year-1 fraction remaining (stub)")
    formula(ws, 11, 4, "=IF(D10>=D9,0,MIN(1,YEARFRAC(D10,D9)))", "0.000")
    label(ws, 6, 6, "Discount rate (WACC)")
    formula(ws, 6, 8, f"={wacc.wacc_cell}", S.FMT_PCT, bold=True)
    label(ws, 7, 6, "Share price")
    p = wacc.price_cell
    formula(ws, 7, 8, f'=IF(ISNUMBER({p}),{p},"")', S.FMT_PERUNIT)
    if market.price is None:
        ws.cell(7, 9, "enter price on WACC sheet").font = FONT_NOTE
    label(ws, 8, 6, "Mid-year convention (1 = on)")
    _input(ws.cell(8, 8), 0, None)
    ws.cell(8, 8).number_format = "0"


def _terminal_value(ws, row: dict, last: str, wacc: WaccLayout) -> None:
    """Perpetuity method in B:C, exit-multiple method in E:H (rows 37-46)."""
    r = row
    pv_rng = f"{S.col(FC)}{r['PV']}:{last}{r['PV']}"
    df_n = f"{last}{r['DF']}"
    label(ws, 37, 2, "Perpetuity growth method", bold=True)
    label(ws, 37, 5, "EBITDA exit multiple method", bold=True)
    label(ws, 38, 2, "Long-term growth rate (g)")
    _input(ws.cell(38, 3), 0.025, "pct")
    label(ws, 38, 5, "Terminal-year EBITDA")
    formula(ws, 38, 8, f"={last}{r['EBITDA']}", S.FMT_MONEY)
    label(ws, 39, 2, "Terminal-year UFCF x (1+g)")
    formula(ws, 39, 3, f"={last}{r['UFCF']}*(1+C38)", S.FMT_MONEY)
    label(ws, 39, 5, "EBITDA exit multiple")
    p = wacc.price_cell
    _input(ws.cell(39, 8), f'=IFERROR(IF({p}="",10,({wacc.market_cap_cell}+'
                           f'{wacc.net_debt_cell})/N($D${r["EBITDA"]})),10)', None, formula=True)
    ws.cell(39, 8).number_format = S.FMT_X
    ws.cell(39, 9, "default: current EV / last-year EBITDA (10.0x without a price)").font = \
        FONT_NOTE
    rows = (("Terminal value", "=C39/($H$6-C38)", "=H38*H39", S.FMT_MONEY, False),
            ("PV of terminal value", f"=C40*{df_n}", f"=H40*{df_n}", S.FMT_MONEY, False),
            ("PV of stage-1 cash flows", f"=SUM({pv_rng})", f"=SUM({pv_rng})", S.FMT_MONEY,
             False),
            ("Enterprise value", "=C41+C42", "=H41+H42", S.FMT_MONEY, True),
            ("Terminal value % of EV", "=IFERROR(C41/C43,0)", "=IFERROR(H41/H43,0)", S.FMT_PCT,
             False),
            ("Stage-1 cash flows % of EV", "=IFERROR(C42/C43,0)", "=IFERROR(H42/H43,0)",
             S.FMT_PCT, False))
    for i, (text, perp, exit_, fmt, out) in enumerate(rows):
        rr = 40 + i
        label(ws, rr, 2, text, bold=out)
        label(ws, rr, 5, text, bold=out)
        formula(ws, rr, 3, perp, fmt, output=out)
        formula(ws, rr, 8, exit_, fmt, output=out)
    label(ws, 46, 2, "Implied exit EBITDA multiple")
    formula(ws, 46, 3, f"=IFERROR(C40/{last}{r['EBITDA']},0)", S.FMT_X)
    label(ws, 46, 5, "Implied perpetuity growth rate")
    formula(ws, 46, 8, f"=IFERROR(($H$6*H40-{last}{r['UFCF']})/(H40+{last}{r['UFCF']}),0)",
            S.FMT_PCT)


def _grid(ws, top: int, title: str, corner_text: str, col_base: str, col_step: float,
          col_fmt: str, tv: str, row: dict, last: str, wacc: WaccLayout) -> dict:
    """5x5 value-per-unit grid; WACC down column C, ``col_base`` offsets across D:H.
    ``tv`` is the terminal-value expression with ``{w}`` and ``{x}`` placeholders."""
    label(ws, top - 1, 2, title, bold=True)
    headers(ws, top, {3: corner_text})
    first = S.col(FC)
    ufcf = f"${first}${row['UFCF_ADJ']}:${last}${row['UFCF_ADJ']}"
    t = f"${first}${row['T']}:${last}${row['T']}"
    t_n = f"${last}${row['T']}"
    for j, k in enumerate(STEPS):
        formula(ws, top, 4 + j, offset(col_base, k, col_step), col_fmt, bold=True)
        r = top + 1 + j
        formula(ws, r, 3, offset(WACC_C, k, 0.005), S.FMT_PCT, bold=True)
        for i in range(len(STEPS)):
            w, x = f"$C{r}", f"{S.col(4 + i)}${top}"
            f = (f"=IFERROR((SUMPRODUCT({ufcf},1/(1+{w})^{t})+{tv.format(w=w, x=x)}"
                 f"/(1+{w})^{t_n}-{wacc.net_debt_cell})/{wacc.diluted_units_cell},\"n/a\")")
            formula(ws, r, 4 + i, f, S.FMT_PERUNIT, output=(i == 2 and k == 0))
    return {"corner": (top, 3), "range": f"D{top + 1}:H{top + 5}"}


def football_chart(ws, rows: list[int], anchor: str) -> None:
    """Stacked horizontal bar: invisible 'Low' series + visible 'Spread' series."""
    chart = BarChart()
    chart.type, chart.grouping, chart.overlap = "bar", "stacked", 100
    chart.title = "Value per unit"
    chart.add_data(Reference(ws, min_col=3, max_col=4, min_row=rows[0] - 1, max_row=rows[-1]),
                   titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=2, min_row=rows[0], max_row=rows[-1]))
    chart.series[0].graphicalProperties.noFill = True
    chart.series[0].graphicalProperties.line.noFill = True
    chart.legend = None
    chart.height, chart.width = 6.5, 16
    ws.add_chart(chart, anchor)


def write_dcf(wb, layout: ModelLayout, wacc: WaccLayout, market: MarketInputs,
              fcst_labels: list[str], currency: str, last_fye: date, *,
              company: str = "", ticker: str = "") -> DcfLayout:
    """Write the DCF sheet off the Model, Drivers and WACC sheets."""
    n = layout.last_col - layout.first_fcst_col + 1
    if n < 1 or n != len(fcst_labels):
        raise ValueError(f"{len(fcst_labels)} forecast labels for {n} Model forecast columns")
    ws = S.get_sheet(wb, DCF_SHEET)
    lastc = FC + n - 1
    last = S.col(lastc)
    bar_end = max(lastc, 9)
    S.title_bar(ws, 2, "Discounted Cash Flow Valuation", bar_end,
                subtitle=f"All figures in {currency} MM except per-unit data; "
                         "yellow = input, green = result")
    currency_note(ws, currency)
    S.set_widths(ws, {"A": 2, "B": 40, "C": 13})
    for c in range(4, bar_end + 1):
        S.set_widths(ws, {c: 13})
    _header(ws, wacc, market, last_fye, company, ticker)

    # --- UFCF build ------------------------------------------------------------------
    section_header(ws, 13, "UNLEVERED FREE CASH FLOW", lastc)
    hist_label = wb[layout.sheet].cell(4, layout.first_fcst_col - 1).value or "Actual"
    S.header_row(ws, 14, ["Fiscal year"], 2)
    S.header_row(ws, 14, [hist_label] + list(fcst_labels), 4, [False] + [True] * n)
    row = {"FYE": 15}
    label(ws, 15, 2, "Fiscal year ended")
    formula(ws, 15, 4, "=D8", FMT_DATE)
    for c in range(FC, lastc + 1):
        formula(ws, 15, c, f"=EDATE({S.col(c - 1)}15,12)", FMT_DATE)
    for i, (_, key, _, _) in enumerate(_UFCF_ROWS):
        row[key] = 16 + i
    for text, key, fmt, bold in _UFCF_ROWS:
        r = row[key]
        label(ws, r, 2, text, bold=bold)
        for c in range(FC - 1, lastc + 1):
            model_col = S.col(layout.first_fcst_col + c - FC)   # DCF D = last Model history
            cell = formula(ws, r, c, _ufcf_cell(key, c, row, model_col, layout), fmt, bold=bold)
            if c >= FC:
                cell.fill = S.FILL_FORECAST
            if key == "UFCF":
                cell.border = S.BORDER_TOP

    # --- discounting -----------------------------------------------------------------
    section_header(ws, 28, "PRESENT VALUE OF UFCF", lastc)
    headers(ws, 29, {4: "Valuation date", **{FC + i: ("Year 1 (stub)" if i == 0 else
                                                      f"Year {i + 1}") for i in range(n)}})
    row.update(DISC_DATE=30, T=31, UFCF_ADJ=32, DF=33, PV=34)
    for r, text in ((30, "Discount date"), (31, "Years from valuation date (t)"),
                    (32, "UFCF (year 1 stub-adjusted)"), (33, "Discount factor"),
                    (34, "PV of UFCF")):
        label(ws, r, 2, text, bold=(r == 34))
    formula(ws, 30, 4, "=D10", FMT_DATE)
    for c in range(FC, lastc + 1):
        L = S.col(c)
        prev = VAL_DATE if c == FC else f"{S.col(c - 1)}15"
        formula(ws, 30, c, f"=IF({MID_C}=0,{L}15,AVERAGE({prev},{L}15))", FMT_DATE)
        formula(ws, 31, c, f"=({L}30-{VAL_DATE})/365", "0.000")
        formula(ws, 32, c, f"={L}{row['UFCF']}*{STUB}" if c == FC else f"={L}{row['UFCF']}",
                S.FMT_MONEY)
        formula(ws, 33, c, f"=1/(1+{WACC_C})^{L}31", "0.0000")
        formula(ws, 34, c, f"={L}32*{L}33", S.FMT_MONEY, bold=True)

    # --- terminal value --------------------------------------------------------------
    section_header(ws, 36, "TERMINAL VALUE", lastc)
    _terminal_value(ws, row, last, wacc)
    row.update(TV=40, TV_PV=41, STAGE1=42, EV_TV=43)

    # --- equity value ----------------------------------------------------------------
    section_header(ws, 48, "EQUITY VALUE & VALUE PER UNIT", lastc)
    headers(ws, 49, {3: "Perpetuity", 4: "Exit multiple"})
    row.update(EV=50, NET_DEBT=51, EQUITY=52, UNITS=53, VPU=54, PRICE=55, UPSIDE=56)
    p = wacc.price_cell
    for r, text, perp, exit_, fmt, out in (
            (50, "Enterprise value", "=C43", "=H43", S.FMT_MONEY, False),
            (51, "Less: net debt", f"={wacc.net_debt_cell}", f"={wacc.net_debt_cell}",
             S.FMT_MONEY, False),
            (52, "Equity value", "=C50-C51", "=D50-D51", S.FMT_MONEY, False),
            (53, "Diluted units (MM)", f"={wacc.diluted_units_cell}",
             f"={wacc.diluted_units_cell}", S.FMT_UNITS, False),
            (54, "Value per unit", '=IFERROR(C52/C53,"n/a")', '=IFERROR(D52/D53,"n/a")',
             S.FMT_PERUNIT, True),
            (55, "Current price", f"={PRICE_C}", f"={PRICE_C}", S.FMT_PERUNIT, False),
            (56, "% over / (under) valued",
             f'=IF({p}="","enter price",IFERROR({p}/C54-1,"n/a"))',
             f'=IF({p}="","enter price",IFERROR({p}/D54-1,"n/a"))', S.FMT_PCT, False)):
        label(ws, r, 2, text, bold=out)
        formula(ws, r, 3, perp, fmt, output=out)
        formula(ws, r, 4, exit_, fmt, output=out)

    # --- multiples -------------------------------------------------------------------
    section_header(ws, 58, "IMPLIED MULTIPLES", lastc)
    headers(ws, 59, {3: "Current", 4: "Current", 5: "Year 1", 6: "Year 1",
                     7: "Terminal year", 8: "Terminal year"})
    headers(ws, 60, {c: ("Perpetuity" if c % 2 else "Exit") for c in range(3, 9)})
    row.update(MULT_REV=61, MULT_EBITDA=62, MULT_EBIT=63)
    for r, text, key in ((61, "EV / Revenue", "REV"), (62, "EV / EBITDA", "EBITDA"),
                         (63, "EV / EBIT", "EBIT")):
        label(ws, r, 2, text)
        rr = row[key]
        for c, ev, denom in ((3, "$C$50", f"N($D${rr})"), (4, "$D$50", f"N($D${rr})"),
                             (5, "$C$50", f"E{rr}"), (6, "$D$50", f"E{rr}"),
                             (7, "$C$50", f"{last}{rr}"), (8, "$D$50", f"{last}{rr}")):
            formula(ws, r, c, f'=IFERROR({ev}/{denom},"n/a")', S.FMT_X)

    # --- sensitivity grids -----------------------------------------------------------
    section_header(ws, 65, "SENSITIVITY: VALUE PER UNIT", lastc)
    last_ufcf, last_ebitda = f"${last}${row['UFCF']}", f"${last}${row['EBITDA']}"
    grids = {
        "perpetuity": _grid(ws, 67, "WACC (rows) x long-term growth g (columns)", "WACC \\ g",
                            "$C$38", 0.005, S.FMT_PCT,
                            last_ufcf + "*(1+{x})/({w}-{x})", row, last, wacc),
        "exit": _grid(ws, 75, "WACC (rows) x exit EBITDA multiple (columns)",
                      "WACC \\ multiple", "$H$39", 0.5, S.FMT_X,
                      last_ebitda + "*{x}", row, last, wacc),
    }

    # --- football field --------------------------------------------------------------
    section_header(ws, 82, "FOOTBALL FIELD: VALUE PER UNIT", lastc)
    headers(ws, 83, {3: "Low", 4: "Spread", 5: "High"})
    ff = [84, 85, 86]
    gp, ge = grids["perpetuity"]["range"], grids["exit"]["range"]
    label(ws, 84, 2, '="DCF perpetuity: g "&TEXT($D$67,"0.0%")&" to "&TEXT($H$67,"0.0%")'
                     '&", WACC "&TEXT($C$68,"0.0%")&" to "&TEXT($C$72,"0.0%")')
    label(ws, 85, 2, '="DCF exit multiple: "&TEXT($D$75,"0.0x")&" to "&TEXT($H$75,"0.0x")'
                     '&", WACC "&TEXT($C$76,"0.0%")&" to "&TEXT($C$80,"0.0%")')
    for r, rng in ((84, gp), (85, ge)):
        formula(ws, r, 3, f"=MIN({rng})", S.FMT_PERUNIT)
        formula(ws, r, 4, f"=E{r}-C{r}", S.FMT_PERUNIT)
        formula(ws, r, 5, f"=MAX({rng})", S.FMT_PERUNIT)
    label(ws, 86, 2, "52-week trading range (enter low / high)")
    _input(ws.cell(86, 3), None, "perunit")
    formula(ws, 86, 4, "=IF(AND(ISNUMBER(C86),ISNUMBER(E86)),E86-C86,0)", S.FMT_PERUNIT)
    _input(ws.cell(86, 5), None, "perunit")
    football_chart(ws, ff, "B88")

    ws.freeze_panes = "C5"
    a = lambda c, r: absref(DCF_SHEET, c, r)  # noqa: E731
    return DcfLayout(
        header_row=14, row=row,
        ev_cells={"perpetuity": a(3, 43), "exit": a(8, 43)},
        equity_value_cells={"perpetuity": a(3, 52), "exit": a(4, 52)},
        value_per_unit_cells={"perpetuity": a(3, 54), "exit": a(4, 54)},
        upside_cells={"perpetuity": a(3, 56), "exit": a(4, 56)},
        growth_cell=a(3, 38), exit_multiple_cell=a(8, 39), grids=grids,
        grid_centre_cells={k: a(6, g["corner"][0] + 3) for k, g in grids.items()},
        football_rows=ff, first_fcst_col=FC, last_fcst_col=lastc,
    )
