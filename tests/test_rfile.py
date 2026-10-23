from pathlib import Path

from sec_model.rfile import parse_rfile

FIXTURES = Path(__file__).parent / "fixtures"


def FIX(rel: str) -> str:
    return (FIXTURES / rel).read_text(encoding="utf-8")


def _row(label: str, tag: str, cells: str, style: str = "border-bottom: 0px;") -> str:
    return (
        f'<tr class="ro"><td class="pl" style="{style}" valign="top">'
        f'<a class="a" href="javascript:void(0);" '
        f"onclick=\"Show.showAR( this, 'defref_{tag}', window );\">{label}</a></td>"
        f"{cells}</tr>"
    )


def _table(*rows: str, title: str = "INCOME STATEMENT - USD ($)", extra: str = "") -> str:
    return (
        '<html><body><table class="report">'
        '<tr><th class="tl" colspan="1" rowspan="2"><div><strong>'
        f"{title}<br> $ in Millions</strong></div></th>"
        '<th class="th" colspan="2">12 Months Ended</th></tr>'
        '<tr><th class="th"><div>Dec. 31, 2025</div></th>'
        '<th class="th"><div>Dec. 31, 2024</div></th></tr>'
        + "".join(rows)
        + "</table>" + extra + "</body></html>"
    )


TWO_REVENUE_ROWS_HTML = _table(
    _row("Product", "us-gaap_Revenues", '<td class="nump">$ 10</td><td class="nump">$ 9</td>'),
    _row("Service", "us-gaap_Revenues", '<td class="nump">5</td><td class="nump">4</td>'),
)


def test_oke_cf_columns_and_scale():
    s = parse_rfile(FIX("OKE/CF.htm"), "CF", "u")
    assert s.currency == "USD" and s.money_scale == 1
    assert [c.end.year for c in s.columns] == [2025, 2024, 2023] and s.columns[0].months == 12


def test_oke_cf_first_rows():
    s = parse_rfile(FIX("OKE/CF.htm"), "CF", "u")
    ni = next(r for r in s.rows if r.tag == "us-gaap_ProfitLoss")
    assert ni.values == [3462.0, 3112.0, 2659.0] and ni.indent >= 1
    assert s.rows[0].is_abstract


def test_et_blank_first_column_and_footnote():
    # ET's title cell spans 2 columns: every row carries a row-footnote cell
    # (<td class="th"><sup>[1]</sup></td>) between the label and the values.
    # It is not a period column; values must stay aligned to the dated columns.
    s = parse_rfile(FIX("ET/CF.htm"), "CF", "u")
    assert [c.end.year for c in s.columns] == [2025, 2024, 2023]
    ni = next(r for r in s.rows if r.tag == "us-gaap_ProfitLoss")
    assert ni.values == [5708.0, 6565.0, 5294.0]
    eq = next(r for r in s.rows if r.tag == "us-gaap_IncomeLossFromEquityMethodInvestments")
    assert eq.values == [-419.0, -379.0, -383.0]   # "[1]" is not a number
    assert eq.label == "Income (Loss) from Equity Method Investments"


def test_epd_custom_tag_kept():
    s = parse_rfile(FIX("EPD/CF.htm"), "CF", "u")
    assert any(r.tag == "epd_DepreciationAndAccretion" and r.values[0] == 2107.0 for r in s.rows)


def test_thousands_filer_scaled_to_millions():
    s = parse_rfile(FIX("KNTK/BS.htm"), "BS", "u")
    assets = next(r for r in s.rows if r.tag == "us-gaap_Assets")
    assert s.money_scale == 0.001 and assets.values[0] < 1e6   # i.e. in $MM


def test_thousands_scaling_is_exact():
    s = parse_rfile(FIX("KNTK/IS.htm"), "IS", "u")
    rev = next(r for r in s.rows if r.tag == "us-gaap_RevenueFromContractWithCustomerExcludingAssessedTax")
    assert rev.values == [1764.389, 1482.929, 1256.412]
    assert s.share_scale == 0.001


def test_bs_is_instant():
    s = parse_rfile(FIX("OKE/BS.htm"), "BS", "u")
    assert all(c.months is None for c in s.columns) and len(s.columns) == 2


def test_per_share_not_scaled():
    s = parse_rfile(FIX("OKE/IS.htm"), "IS", "u")
    eps = next(r for r in s.rows if "EarningsPerShareDiluted" in r.tag)
    assert eps.unit == "perunit" and 0 < eps.values[0] < 100


def test_non_december_fye():
    s = parse_rfile(FIX("ORCL/IS.htm"), "IS", "u")
    assert s.columns[0].end.month == 5


def test_bank_layout_parses():
    s = parse_rfile(FIX("JPM/IS.htm"), "IS", "u")
    assert len([r for r in s.rows if not r.is_abstract]) >= 20
    assert not any(r.tag == "us-gaap_OperatingIncomeLoss" for r in s.rows)   # parser assumes nothing about layout


def test_duplicate_tag_gets_label_key():
    html = TWO_REVENUE_ROWS_HTML   # minimal table with two us-gaap_Revenues rows, labels "Product" and "Service"
    s = parse_rfile(html, "IS", "u")
    assert [r.key for r in s.rows] == ["us-gaap_Revenues", "us-gaap_Revenues|Service"]


def test_dimension_member_rows_have_distinct_keys():
    # OKE IS repeats "Total revenues" under two ProductOrServiceAxis members.
    s = parse_rfile(FIX("OKE/IS.htm"), "IS", "u")
    keys = [r.key for r in s.rows]
    assert len(keys) == len(set(keys))
    tag = "us-gaap_RevenueFromContractWithCustomerExcludingAssessedTax"
    total = next(r for r in s.rows if r.key == tag)
    assert total.values == [33629.0, 21698.0, 17677.0]   # bare tag = consolidated row
    members = [r for r in s.rows if r.tag == tag and r.key != tag]
    assert [m.values[0] for m in members] == [28878.0, 4751.0]
    header = next(r for r in s.rows if "Axis=" in r.tag)
    assert header.is_abstract and header.label == "Commodity sales"


def test_padding_left_em_sets_indent():
    html = _table(_row("Cost", "us-gaap_CostOfRevenue", '<td class="nump">1</td><td class="nump">2</td>',
                       style="padding-left: 2em;"))
    assert parse_rfile(html, "IS", "u").rows[0].indent == 2


def test_equity_lines_mentioning_units_issued_stay_money():
    # Labels like "Common limited partner interests (2,161,760,683 units issued and
    # outstanding ...)" are dollar lines; the page's own XBRL data type says so.
    epd = parse_rfile(FIX("EPD/BS.htm"), "BS", "u")
    lp = next(r for r in epd.rows if r.tag == "us-gaap_LimitedPartnersCapitalAccount")
    assert lp.unit == "money" and lp.values[0] > 1000
    jpm = parse_rfile(FIX("JPM/BS.htm"), "BS", "u")
    assert next(r for r in jpm.rows if r.tag == "us-gaap_CommonStockValue").unit == "money"
    et = parse_rfile(FIX("ET/CF.htm"), "CF", "u")
    assert next(r for r in et.rows
                if r.tag == "us-gaap_ProceedsFromIssuanceOfPreferredLimitedPartnersUnits").unit == "money"


def test_unit_hint_overrides_label_rules():
    s = parse_rfile(TWO_REVENUE_ROWS_HTML, "IS", "u", unit_hint=lambda tag: "perunit")
    assert s.rows[0].unit == "perunit"


def test_currency_from_title():
    rev = _row("Revenue", "us-gaap_Revenues", '<td class="nump">1</td><td class="nump">2</td>')
    loss = parse_rfile(_table(rev, title="CONSOLIDATED STATEMENTS OF INCOME (LOSS) - USD ($)"), "IS", "u")
    assert loss.currency == "USD"
    assert parse_rfile(_table(rev, title="STATEMENTS OF INCOME - CAD ($)"), "IS", "u").currency == "CAD"


def test_malformed_authref_block_falls_back_to_label_rules():
    broken = ('<div style="display: none;"><table class="authRefData" id="defref_us-gaap_Revenues">'
              "<tr><th><strong> Data Type:</strong></th></tr></table></div>")
    rev = _row("Revenue", "us-gaap_Revenues", '<td class="nump">$ 10</td><td class="nump">$ 9</td>')
    s = parse_rfile(_table(rev, extra=broken), "IS", "u")
    assert s.rows[0].unit == "money" and s.rows[0].values == [10.0, 9.0]


def test_non_numeric_cells_are_none():
    cells = ['<td class="text">[1]</td>', '<td class="num">—</td>', '<td class="nump">1e400</td>',
             '<td class="text">abc</td>', '<td class="nump">nan</td>', '<td class="nump">-</td>']
    rows = [_row(f"R{i}", f"us-gaap_X{i}", c + '<td class="nump">1</td>') for i, c in enumerate(cells)]
    s = parse_rfile(_table(*rows), "IS", "u")
    assert [r.values for r in s.rows] == [[None, 1.0]] * len(cells)
