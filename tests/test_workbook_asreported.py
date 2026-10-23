from datetime import date
from io import BytesIO

from openpyxl import Workbook, load_workbook

from sec_model.concepts import CONCEPTS
from sec_model.types import (CheckResult, ConceptValue, MergedRow, MergedStatement,
                             SourceRec)
from sec_model.workbook import styles as S
from sec_model.workbook.asreported import SHEET_NAMES, write_as_reported
from sec_model.workbook.mapping_sheet import write_mapping
from sec_model.workbook.meta import CoverInfo, write_checks, write_cover, write_sources

FYS = ["FY2023", "FY2024"]


def _src(fy):
    return SourceRec(f"0001-24-{fy[-2:]}", "10-K", date(2024, 2, 20), f"https://sec.gov/R2_{fy}.htm",
                     "In Millions")


def _row(key, label, vals, indent=0, abstract=False, unit="money", still=True):
    return MergedRow(key, "us-gaap_" + key, label, indent, abstract, unit,
                     dict(zip(FYS, vals)), still)


def _is():
    rows = [
        _row("Hdr", "Revenues [Abstract]", [None, None], abstract=True),
        _row("Revenues", "Revenues", [100.0, 120.0], indent=1),
        _row("Legacy", "Legacy item", [5.0, None], indent=1, still=False),
        _row("Total", "Total costs and expenses", [60.0, None], indent=1),
        _row("EPS", "Earnings per unit", [1.5, 1.75], unit="perunit"),
        _row("Units", "Units outstanding", [10.0, 11.0], unit="shares"),
    ]
    return MergedStatement("IS", "USD", FYS, {fy: date(int(fy[2:]), 12, 31) for fy in FYS}, rows,
                           {fy: _src(fy) for fy in FYS})


def _cv(concept, fy, value, key, status="OK", tag=None):
    return ConceptValue(concept, fy, value, key, tag or ("us-gaap_" + key if key else None),
                        "lbl", status)


def test_as_reported_layout_and_hidden_key_column():
    wb = Workbook()
    rows = write_as_reported(wb, _is(), SHEET_NAMES["IS"])
    buf = BytesIO()
    wb.save(buf)
    ws = load_workbook(BytesIO(buf.getvalue()))["IS (As Reported)"]
    assert rows["Revenues"] == 6 and rows["Hdr"] == 5
    assert ws.column_dimensions["A"].hidden is True
    assert ws["A6"].value == "Revenues" and ws["B6"].value == "Revenues"
    assert ws["B6"].alignment.indent == 1
    assert ws["C6"].value == "$MM" and ws["C9"].value == "per unit" and ws["C10"].value == "MM units"
    assert ws["D6"].value == 100.0 and ws["E6"].value == 120.0
    assert ws["D6"].font.color.rgb.endswith("0000FF") and ws["D6"].font.name == "Calibri"
    assert ws["D6"].number_format == S.FMT_MONEY and ws["D9"].number_format == S.FMT_PERUNIT
    assert ws["D4"].value == "FY2023" and ws["E4"].value == "FY2024"
    assert ws.freeze_panes == "D5"
    assert ws.column_dimensions["B"].width == 48 and ws.column_dimensions["D"].width == 13
    assert ws["B2"].fill.start_color.rgb.endswith("1F3864") and ws["B2"].font.size == 14
    # None -> blank, never 0; abstract row has no values
    assert ws["E8"].value is None and ws["E7"].value is None and ws["D5"].value is None


def test_non_usd_unit_note():
    wb = Workbook()
    stmt = _is()
    stmt.currency = "CAD"
    write_as_reported(wb, stmt, SHEET_NAMES["IS"])
    assert wb["IS (As Reported)"]["C6"].value == "CAD MM"


def test_dropped_row_label_and_style():
    wb = Workbook()
    write_as_reported(wb, _is(), SHEET_NAMES["IS"])
    ws = wb["IS (As Reported)"]
    assert ws["B7"].value == "Legacy item (no longer reported)"
    assert ws["B7"].font.italic and ws["B7"].font.color.rgb.endswith("808080")
    assert ws["D7"].font.italic
    assert ws["B5"].font.bold and ws["B5"].fill.start_color.rgb.endswith("F2F2F2")
    assert ws["B8"].font.bold and ws["B8"].border.top.style == "thin"


def test_mapping_value_formula_points_at_key():
    wb = Workbook()
    stmt = _is()
    write_as_reported(wb, stmt, SHEET_NAMES["IS"])
    mapping = {c.name: {} for c in CONCEPTS}
    mapping["Revenue"] = {fy: _cv("Revenue", fy, v, "Revenues") for fy, v in zip(FYS, (100.0, 120.0))}
    mapping["Gains"] = {fy: _cv("Gains", fy, None, None, "UNMAPPED") for fy in FYS}
    mapping["OperatingIncome"] = {
        fy: _cv("OperatingIncome", fy, 7.0, None, "DERIVED", tag="SUM(a,b)") for fy in FYS}
    info = write_mapping(wb, mapping, dict(SHEET_NAMES),
                         {"IS": stmt, "BS": stmt, "CF": stmt}, FYS)
    ws = wb["Mapping"]
    r = info["rows"]["Revenue"]
    assert r == 5 + [c.name for c in CONCEPTS].index("Revenue")
    vc, kc = info["value_cols"]["FY2024"], info["key_cols"]["FY2024"]
    f = ws.cell(r, vc).value
    keyref = ws.cell(r, kc).coordinate
    assert "MATCH(" in f and keyref in f and "'IS (As Reported)'!E:E" in f and "$A:$A" in f
    assert f.startswith("=IFERROR(INDEX(")
    assert ws.cell(r, kc).value == "Revenues" and ws.cell(r, kc).fill.start_color.rgb.endswith("FFF2CC")
    assert ws.cell(r, 1).value == "Revenue" and ws.cell(r, 2).value == "OK"
    assert ws.cell(r, 3).value == "IS (As Reported)"
    assert info["key_cols"]["FY2023"] == info["value_cols"]["FY2024"] + 2   # one spacer column
    # unmapped: blank yellow key, still a formula; derived: hard-coded number + tag text
    g = info["rows"]["Gains"]
    assert ws.cell(g, info["key_cols"]["FY2024"]).value is None
    assert ws.cell(g, info["value_cols"]["FY2024"]).value.startswith("=IFERROR(")
    d = info["rows"]["OperatingIncome"]
    assert ws.cell(d, info["value_cols"]["FY2024"]).value == 7.0
    assert ws.cell(d, info["value_cols"]["FY2024"]).font.color.rgb.endswith("0000FF")
    assert ws.cell(d, info["key_cols"]["FY2024"]).value == "SUM(a,b)"


def test_cover_lists_filing_per_year_and_check_counts():
    wb = Workbook()
    names = ["Cover", "Checks", "Sources"]
    info = CoverInfo("Acme Corp", "ACM", 123, "December 31", "USD", FYS, ["FY2025E"],
                     date(2026, 10, 22), sheet_names=names,
                     check_counts={"PASS": 5, "WARN": 1, "FAIL": 2, "SKIP": 0},
                     filings=[("FY2023", "10-K", "0001-24-23", date(2024, 2, 20)),
                              ("FY2024", "10-K/A", "0001-25-24", date(2025, 2, 20))])
    write_cover(wb, info)
    write_checks(wb, [])
    write_sources(wb, [])
    ws = wb["Cover"]
    assert wb.sheetnames[0] == "Cover" and "Sheet" not in wb.sheetnames
    assert ws["B2"].value == "Acme Corp  |  SEC 3-Statement Model"
    cells = {c.value for row in ws.iter_rows() for c in row if c.value is not None}
    assert {"FY2023", "10-K/A", "0001-25-24", "2025-02-20", "Acme Corp", "ACM", 123} <= cells
    counts = {ws.cell(c.row, 2).value: c.value for row in ws.iter_rows() for c in row
              if c.column == 3 and ws.cell(c.row, 2).value in ("PASS", "WARN", "FAIL", "SKIP")
              and isinstance(c.value, int)}
    assert counts == {"PASS": 5, "WARN": 1, "FAIL": 2, "SKIP": 0}
    fail_lines = [c for row in ws.iter_rows() for c in row
                  if c.column == 2 and c.value == "FAIL" and c.font.bold
                  and c.font.color.rgb.endswith("C00000")]
    assert fail_lines
    links = [c for row in ws.iter_rows() for c in row if c.hyperlink is not None]
    assert [c.value for c in links] == names
    assert links[1].hyperlink.location == "'Checks'!A1"


def test_workbook_reopens_without_error():
    wb = Workbook()
    stmt = _is()
    write_as_reported(wb, stmt, SHEET_NAMES["IS"])
    write_checks(wb, [CheckResult("BS balances", "FY2024", "PASS", 10.0, 10.0, 0.5, ""),
                      CheckResult("x", "FY2024", "FAIL", 1.0, None, 0.5, "bad"),
                      CheckResult("y", "FY2024", "SKIP", None, None, 0.0, "no data")])
    write_sources(wb, [stmt])
    buf = BytesIO()
    wb.save(buf)
    wb2 = load_workbook(BytesIO(buf.getvalue()))
    ck = wb2["Checks"]
    assert ck.freeze_panes == "A2" and ck.auto_filter.ref == "A1:G4"
    assert ck["C2"].fill.start_color.rgb.endswith("D9EAD3")
    assert ck["C3"].fill.start_color.rgb.endswith("F4CCCC")
    assert ck["C4"].fill.start_color.rgb.endswith("EEEEEE")
    assert ck["D2"].number_format == S.FMT_MONEY and ck["E4"].value is None
    sc = wb2["Sources"]
    assert sc.max_row == 3 and sc["A2"].value == "Income Statement" and sc["B3"].value == "FY2024"
    assert sc["F2"].hyperlink.target == "https://sec.gov/R2_FY2023.htm"
    assert sc["G2"].value == "In Millions"
