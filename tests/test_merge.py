from datetime import date

import pytest

from sec_model.merge import merge_statements
from sec_model.types import (Column, Filing, ParsedStatement, Row,
                             UserFacingError)

FY_ENDS = {"FY2023": date(2023, 12, 31), "FY2024": date(2024, 12, 31),
           "FY2025": date(2025, 12, 31)}


def filing(year, accession=None, filing_date=None, form="10-K"):
    return Filing(cik=1, accession=accession or f"0000000001-{year % 100}-000001",
                  form=form, filing_date=filing_date or date(year + 1, 2, 20),
                  report_date=date(year, 12, 31), primary_doc="x.htm")


def row(key, label, values, **kw):
    return Row(key=key, tag=key.split("|")[0], label=label, indent=kw.get("indent", 0),
               is_abstract=kw.get("is_abstract", False), unit=kw.get("unit", "money"),
               values=values)


def stmt(columns, rows, stype="IS", currency="USD", url="http://r/R4.htm"):
    return ParsedStatement(stype=stype, title="CONSOLIDATED STATEMENTS - USD ($) $ in Millions",
                           currency=currency, money_scale=1, share_scale=1,
                           columns=[Column(end=e, months=m, label="") for e, m in columns],
                           rows=rows, source_url=url)


def cols(*years, months=12):
    return [(date(y, 12, 31), months) for y in years]


def test_restated_value_from_newest_filing_wins():
    old = stmt(cols(2024, 2023), [row("us-gaap_Revenues", "Revenues", [100.0, 90.0])])
    new = stmt(cols(2025, 2024, 2023), [row("us-gaap_Revenues", "Revenues", [120.0, 105.0, 95.0])])
    m = merge_statements([(filing(2024), old), (filing(2025), new)],
                         ["FY2023", "FY2024", "FY2025"], FY_ENDS)
    assert len(m.rows) == 1
    assert m.rows[0].values == {"FY2023": 95.0, "FY2024": 105.0, "FY2025": 120.0}
    assert m.rows[0].still_reported is True


def test_label_change_same_tag_is_one_row():
    old = stmt(cols(2024), [row("us-gaap_Revenues", "Revenues", [100.0])])
    new = stmt(cols(2025), [row("us-gaap_Revenues", "Total revenues", [120.0])])
    m = merge_statements([(filing(2024), old), (filing(2025), new)],
                         ["FY2024", "FY2025"], FY_ENDS)
    assert len(m.rows) == 1
    assert m.rows[0].label == "Total revenues"
    assert m.rows[0].values == {"FY2024": 100.0, "FY2025": 120.0}


def test_dropped_row_inserted_after_predecessor_and_flagged():
    old = stmt(cols(2024), [row("A", "A", [1.0]), row("B", "B", [2.0]), row("C", "C", [3.0])])
    new = stmt(cols(2025), [row("A", "A", [10.0]), row("C", "C", [30.0])])
    m = merge_statements([(filing(2024), old), (filing(2025), new)],
                         ["FY2024", "FY2025"], FY_ENDS)
    assert [r.key for r in m.rows] == ["A", "B", "C"]
    b = m.rows[1]
    assert b.still_reported is False
    assert b.values == {"FY2024": 2.0, "FY2025": None}
    assert m.rows[0].still_reported and m.rows[2].still_reported


def test_non_12_month_column_ignored():
    # 1-month stub shares the fiscal year end; only the 12-month column may supply the value
    s = stmt([(date(2025, 12, 31), 1), (date(2025, 12, 31), 12)],
             [row("us-gaap_Revenues", "Revenues", [5.0, 120.0])])
    m = merge_statements([(filing(2025), s)], ["FY2025"], FY_ENDS)
    assert m.rows[0].values == {"FY2025": 120.0}


def test_sources_record_filing_per_year():
    old = stmt(cols(2024, 2023), [row("A", "A", [1.0, 2.0])], url="http://r/old")
    new = stmt(cols(2025, 2024), [row("A", "A", [3.0, 4.0])], url="http://r/new")
    f_old, f_new = filing(2024), filing(2025)
    m = merge_statements([(f_old, old), (f_new, new)],
                         ["FY2023", "FY2024", "FY2025"], FY_ENDS)
    assert m.sources["FY2023"].accession == f_old.accession
    assert m.sources["FY2023"].r_url == "http://r/old"
    assert m.sources["FY2024"].accession == f_new.accession
    assert m.sources["FY2025"].filing_date == f_new.filing_date
    assert m.sources["FY2025"].form == "10-K"
    assert "$ in Millions" in m.sources["FY2025"].scale_text
    assert m.fy_labels == ["FY2023", "FY2024", "FY2025"]


def test_balance_sheet_instant_columns_match():
    s = stmt([(date(2025, 12, 31), None), (date(2024, 12, 31), None)],
             [row("us-gaap_Assets", "Assets", [9.0, 8.0])], stype="BS")
    m = merge_statements([(filing(2025), s)], ["FY2024", "FY2025"], FY_ENDS)
    assert m.stype == "BS"
    assert m.rows[0].values == {"FY2024": 8.0, "FY2025": 9.0}


def test_duration_column_without_months_does_not_match():
    s = stmt([(date(2025, 12, 31), None)], [row("A", "A", [7.0])])
    m = merge_statements([(filing(2025), s)], ["FY2025"], FY_ENDS)
    assert m.rows[0].values == {"FY2025": None}
    assert "FY2025" not in m.sources


def test_currency_change_raises():
    old = stmt(cols(2024), [row("A", "A", [1.0])], currency="USD")
    new = stmt(cols(2025), [row("A", "A", [1.0])], currency="CAD")
    with pytest.raises(UserFacingError, match="USD→CAD"):
        merge_statements([(filing(2024), old), (filing(2025), new)],
                         ["FY2024", "FY2025"], FY_ENDS)


def test_same_filing_date_tie_breaks_by_accession():
    a = stmt(cols(2025), [row("A", "A", [1.0])])
    b = stmt(cols(2025), [row("A", "A", [2.0])])
    d = date(2026, 2, 20)
    fa = filing(2025, accession="0000000001-26-000001", filing_date=d)
    fb = filing(2025, accession="0000000001-26-000002", filing_date=d, form="10-K/A")
    m = merge_statements([(fb, b), (fa, a)], ["FY2025"], FY_ENDS)
    assert m.rows[0].values == {"FY2025": 2.0}
    assert m.sources["FY2025"].form == "10-K/A"


def test_dropped_row_value_comes_from_older_filing_for_its_year():
    old = stmt(cols(2024), [row("A", "A", [1.0]), row("B", "B", [2.0]), row("C", "C", [3.0])])
    new = stmt(cols(2025, 2024), [row("A", "A", [10.0, 1.5]), row("C", "C", [30.0, 3.5])])
    m = merge_statements([(filing(2024), old), (filing(2025), new)],
                         ["FY2024", "FY2025"], FY_ENDS)
    b = next(r for r in m.rows if r.key == "B")
    assert b.values == {"FY2024": 2.0, "FY2025": None}
    # primary supplier for FY2024 is still the newest filing with that column
    assert m.sources["FY2024"].accession == filing(2025).accession
    assert next(r for r in m.rows if r.key == "A").values["FY2024"] == 1.5


def test_overlapping_year_newer_value_wins_when_both_have_key():
    old = stmt(cols(2024), [row("A", "A", [1.0])])
    new = stmt(cols(2025, 2024), [row("A", "A", [10.0, 1.5])])
    m = merge_statements([(filing(2024), old), (filing(2025), new)],
                         ["FY2024", "FY2025"], FY_ENDS)
    assert m.rows[0].values == {"FY2024": 1.5, "FY2025": 10.0}


def test_none_cell_in_newer_filing_does_not_fall_through():
    old = stmt(cols(2024), [row("A", "A", [1.0])])
    new = stmt(cols(2025, 2024), [row("A", "A", [10.0, None])])
    m = merge_statements([(filing(2024), old), (filing(2025), new)],
                         ["FY2024", "FY2025"], FY_ENDS)
    assert m.rows[0].values["FY2024"] is None


def test_empty_parsed_raises():
    with pytest.raises(UserFacingError, match="No parsed statements to merge."):
        merge_statements([], ["FY2025"], FY_ENDS)


def test_label_missing_from_fy_ends_yields_none_and_no_source():
    s = stmt(cols(2025), [row("A", "A", [1.0])])
    m = merge_statements([(filing(2025), s)], ["FY2025", "FY2026"], FY_ENDS)
    assert m.rows[0].values == {"FY2025": 1.0, "FY2026": None}
    assert "FY2026" not in m.sources


def test_short_row_values_yield_none():
    s = stmt(cols(2025, 2024), [row("A", "A", [1.0])])
    m = merge_statements([(filing(2025), s)], ["FY2024", "FY2025"], FY_ENDS)
    assert m.rows[0].values == {"FY2024": None, "FY2025": 1.0}


# ---- one designated supplier filing per fiscal year (KNTK predecessor recast) ----
def test_designated_supplier_is_newest_filing_covering_all_three_statements():
    from sec_model.merge import designated_suppliers
    f24, f25 = filing(2024), filing(2025)
    s24 = {"IS": stmt(cols(2024, 2023), []), "CF": stmt(cols(2024, 2023), [], stype="CF"),
           "BS": stmt(cols(2024, 2023, months=None), [], stype="BS")}
    # The FY2025 10-K has FY2023 IS/CF columns but its BS stops at FY2024.
    s25 = {"IS": stmt(cols(2025, 2024, 2023), []), "CF": stmt(cols(2025, 2024, 2023), [], stype="CF"),
           "BS": stmt(cols(2025, 2024, months=None), [], stype="BS")}
    sup = designated_suppliers([(f24, s24), (f25, s25)], FY_ENDS)
    assert sup == {"FY2023": f24.accession, "FY2024": f25.accession, "FY2025": f25.accession}


def test_supplier_filing_values_first_and_lacking_keys_from_others():
    old = stmt(cols(2024, 2023), [row("us-gaap_Revenues", "Revenues", [148.0, 136.0]),
                                  row("Legacy", "Legacy", [5.0, 4.0])])
    new = stmt(cols(2025, 2024, 2023), [row("us-gaap_Revenues", "Revenues", [1213.0, 662.0, 410.0])])
    f24, f25 = filing(2024), filing(2025)
    m = merge_statements([(f24, old), (f25, new)], ["FY2023", "FY2024", "FY2025"], FY_ENDS,
                         suppliers={"FY2023": f24.accession})
    rev = next(r for r in m.rows if r.key == "us-gaap_Revenues")
    assert rev.values == {"FY2023": 136.0, "FY2024": 662.0, "FY2025": 1213.0}
    assert rev.value_src == {"FY2023": f24.accession, "FY2024": f25.accession,
                             "FY2025": f25.accession}
    assert m.sources["FY2023"].accession == f24.accession
    assert m.sources["FY2024"].accession == f25.accession
    legacy = next(r for r in m.rows if r.key == "Legacy")   # key fallback still works
    assert legacy.values["FY2024"] == 5.0 and legacy.value_src["FY2024"] == f24.accession
