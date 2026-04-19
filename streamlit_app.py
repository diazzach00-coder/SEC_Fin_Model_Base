import time
import warnings
from datetime import datetime
from io import BytesIO

import pandas as pd
import requests
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

warnings.filterwarnings("ignore")

st.set_page_config(page_title="SEC 10-K Model Builder", layout="wide")


IS_SPEC = [
    ("Revenue", [
        "RevenueFromContractWithCustomerExcludingAssessedTax",
        "Revenues",
        "SalesRevenueNet",
        "RevenueFromContractWithCustomerIncludingAssessedTax",
    ], False, False),
    ("Cost of Revenue", ["CostOfGoodsSold", "CostOfRevenue", "CostOfGoodsAndServicesSold"], False, False),
    ("Gross Profit", ["GrossProfit"], False, "Revenue - Cost of Revenue"),
    ("R&D Expense", ["ResearchAndDevelopmentExpense", "ResearchAndDevelopmentExpenseExcludingAcquiredInProcessCost"], False, False),
    ("SG&A Expense", ["SellingGeneralAndAdministrativeExpense", "GeneralAndAdministrativeExpense"], False, False),
    ("Operating Income", ["OperatingIncomeLoss", "IncomeLossFromContinuingOperationsBeforeIncomeTaxes"], False, False),
    ("Interest Expense", ["InterestExpense", "InterestAndDebtExpense"], False, False),
    ("Pretax Income", [
        "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
        "IncomeLossFromContinuingOperationsBeforeIncomeTaxesDomestic",
    ], False, False),
    ("Income Tax Expense", ["IncomeTaxExpenseBenefit", "CurrentIncomeTaxExpenseBenefit"], False, False),
    ("Net Income", ["NetIncomeLoss", "NetIncomeLossAvailableToCommonStockholdersBasic"], False, False),
    ("EPS Basic (per share)", ["EarningsPerShareBasic"], False, False),
    ("EPS Diluted (per share)", ["EarningsPerShareDiluted"], False, False),
    ("Diluted Shares", ["WeightedAverageNumberOfDilutedSharesOutstanding", "CommonStockSharesOutstanding"], False, False),
    ("D&A", ["DepreciationDepletionAndAmortization", "Depreciation", "DepreciationAndAmortization"], False, False),
]

BS_SPEC = [
    ("Cash & Equivalents", ["CashAndCashEquivalentsAtCarryingValue", "CashCashEquivalentsAndShortTermInvestments"], False),
    ("Short-Term Investments", ["ShortTermInvestments", "AvailableForSaleSecuritiesCurrent"], False),
    ("Accounts Receivable", ["AccountsReceivableNetCurrent", "ReceivablesNetCurrent"], False),
    ("Inventory", ["InventoryNet", "InventoryGross"], False),
    ("Total Current Assets", ["AssetsCurrent"], False),
    ("PP&E Net", [
        "PropertyPlantAndEquipmentNet",
        "PropertyPlantAndEquipmentAndFinanceLeaseRightOfUseAssetAfterAccumulatedDepreciationAndAmortization",
    ], False),
    ("Goodwill", ["Goodwill"], False),
    ("Intangible Assets", ["IntangibleAssetsNetExcludingGoodwill", "FiniteLivedIntangibleAssetsNet"], False),
    ("Total Assets", ["Assets"], False),
    ("Accounts Payable", ["AccountsPayableCurrent", "AccountsPayableAndAccruedLiabilitiesCurrent"], False),
    ("Short-Term Debt", ["ShortTermBorrowings", "LongTermDebtCurrent"], False),
    ("Total Current Liabilities", ["LiabilitiesCurrent"], False),
    ("Long-Term Debt", ["LongTermDebt", "LongTermDebtNoncurrent"], False),
    ("Total Liabilities", ["Liabilities"], False),
    ("Total Equity", ["StockholdersEquity", "StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest"], False),
    ("Total Liabilities & Equity", ["LiabilitiesAndStockholdersEquity"], False),
]

CFS_SPEC = [
    ("Net Income (CFS)", ["NetIncomeLoss"], False),
    ("D&A", ["DepreciationDepletionAndAmortization", "DepreciationAndAmortization"], False),
    ("Stock-Based Compensation", ["ShareBasedCompensation", "AllocatedShareBasedCompensationExpense"], False),
    ("Change in Working Capital", ["IncreaseDecreaseInOperatingCapital"], False),
    ("Changes in Receivables", ["IncreaseDecreaseInAccountsReceivable", "IncreaseDecreaseInReceivables"], False),
    ("Changes in Inventory", ["IncreaseDecreaseInInventories"], False),
    ("Changes in Payables", ["IncreaseDecreaseInAccountsPayable", "IncreaseDecreaseInAccountsPayableAndAccruedLiabilities"], False),
    ("Operating Cash Flow", ["NetCashProvidedByUsedInOperatingActivities"], False),
    ("Capital Expenditures", ["PaymentsToAcquirePropertyPlantAndEquipment", "PaymentsForCapitalImprovements"], True),
    ("Acquisitions", ["PaymentsToAcquireBusinessesNetOfCashAcquired", "PaymentsToAcquireBusinessesAndInterestInAffiliates"], True),
    ("Investing Cash Flow", ["NetCashProvidedByUsedInInvestingActivities"], False),
    ("Debt Issuance", ["ProceedsFromIssuanceOfLongTermDebt", "ProceedsFromIssuanceOfDebt"], False),
    ("Debt Repayment", ["RepaymentsOfLongTermDebt", "RepaymentsOfDebt"], True),
    ("Dividends Paid", ["PaymentsOfDividends", "PaymentsOfDividendsCommonStock"], True),
    ("Share Repurchases", ["PaymentsForRepurchaseOfCommonStock"], True),
    ("Financing Cash Flow", ["NetCashProvidedByUsedInFinancingActivities"], False),
]


@st.cache_data(show_spinner=False)
def fetch_json(url: str, user_agent: str, retries: int = 3):
    headers_data = {"User-Agent": user_agent, "Accept-Encoding": "gzip, deflate", "Host": "data.sec.gov"}
    headers_sec = {"User-Agent": user_agent, "Accept-Encoding": "gzip, deflate", "Host": "www.sec.gov"}
    headers = headers_data if "data.sec.gov" in url else headers_sec
    last_error = None
    for attempt in range(retries):
        try:
            time.sleep(0.15)
            response = requests.get(url, headers=headers, timeout=30)
            response.raise_for_status()
            return response.json()
        except Exception as exc:
            last_error = exc
            if attempt < retries - 1:
                time.sleep(2 ** attempt)
    raise last_error


def to_millions(value, unit):
    if value is None:
        return None
    return round(value / 1_000_000, 2) if unit in ("USD", "shares") else value


def resolve_tag(facts_gaap, tag_candidates, fy_end_set, negate=False):
    for tag in tag_candidates:
        if tag not in facts_gaap:
            continue
        units_dict = facts_gaap[tag].get("units", {})
        unit_key = next((k for k in units_dict if k in ("USD", "shares", "pure")), None)
        if unit_key is None:
            continue
        annual = [e for e in units_dict[unit_key] if e.get("form") in ("10-K", "10-K/A")]
        fy_map = {}
        for entry in annual:
            end = entry.get("end")
            if end not in fy_end_set:
                continue
            filed = entry.get("filed", "")
            if end not in fy_map or filed > fy_map[end]["filed"]:
                fy_map[end] = {"val": entry.get("val"), "unit": unit_key, "filed": filed}
        if not fy_map:
            continue
        result = {}
        for end, info in fy_map.items():
            value = to_millions(info["val"], info["unit"])
            result[end] = -value if negate and value is not None else value
        return result, tag
    return {}, None


def build_row(facts_gaap, tag_candidates, fy_ends, negate=False):
    data, tag_used = resolve_tag(facts_gaap, tag_candidates, set(fy_ends), negate=negate)
    return [data.get(fy) for fy in fy_ends], tag_used


def build_financial_data(ticker: str, user_agent: str):
    ticker_upper = ticker.upper().strip()
    tickers_data = fetch_json("https://www.sec.gov/files/company_tickers.json", user_agent)

    cik_raw = None
    company_name = None
    for entry in tickers_data.values():
        if entry["ticker"].upper() == ticker_upper:
            cik_raw = entry["cik_str"]
            company_name = entry["title"]
            break

    if cik_raw is None:
        raise ValueError(f"Ticker '{ticker_upper}' not found in SEC ticker list.")

    cik = str(cik_raw).zfill(10)
    submissions = fetch_json(f"https://data.sec.gov/submissions/CIK{cik}.json", user_agent)
    filings = submissions.get("filings", {}).get("recent", {})

    tenk_filings = [
        {
            "accessionNumber": filings["accessionNumber"][i],
            "filingDate": filings["filingDate"][i],
            "reportDate": filings["reportDate"][i],
            "primaryDocument": filings["primaryDocument"][i],
        }
        for i, form in enumerate(filings.get("form", []))
        if form in ("10-K", "10-K/A")
    ]

    if not tenk_filings:
        raise ValueError(f"No 10-K filings found for {ticker_upper}.")

    tenk_filings.sort(key=lambda x: x["filingDate"], reverse=True)
    tenk_filings = tenk_filings[:5]
    tenk_filings.sort(key=lambda x: x["reportDate"])

    fy_ends = [f["reportDate"] for f in tenk_filings]
    fy_labels = [f"FY{f['reportDate'][:4]}" for f in tenk_filings]
    most_recent_year = int(fy_ends[-1][:4])
    proj_labels = [f"FY{most_recent_year + i}E" for i in range(1, 4)]

    facts_data = fetch_json(f"https://data.sec.gov/api/xbrl/companyfacts/CIK{cik}.json", user_agent)
    facts_gaap = facts_data.get("facts", {}).get("us-gaap", {})

    is_rows, is_tags, missing_is = {}, {}, []
    for label, tags, negate, derived in IS_SPEC:
        if derived:
            vals, tag_used = build_row(facts_gaap, tags, fy_ends, negate)
            if tag_used is None:
                a_lbl, b_lbl = derived.split(" - ")
                a = is_rows.get(a_lbl, [None] * len(fy_ends))
                b = is_rows.get(b_lbl, [None] * len(fy_ends))
                vals = [(aa - bb if aa is not None and bb is not None else None) for aa, bb in zip(a, b)]
                tag_used = "derived"
        else:
            vals, tag_used = build_row(facts_gaap, tags, fy_ends, negate)
        if tag_used is None:
            missing_is.append(label)
        is_rows[label] = vals
        is_tags[label] = tag_used
    df_is = pd.DataFrame(is_rows, index=fy_labels).T
    df_is.index.name = "Line Item"

    bs_rows, bs_tags, missing_bs = {}, {}, []
    for label, tags, negate in BS_SPEC:
        vals, tag_used = build_row(facts_gaap, tags, fy_ends, negate)
        bs_rows[label] = vals
        bs_tags[label] = tag_used
        if tag_used is None:
            missing_bs.append(label)
    df_bs = pd.DataFrame(bs_rows, index=fy_labels).T
    df_bs.index.name = "Line Item"

    cfs_rows, cfs_tags, missing_cfs = {}, {}, []
    for label, tags, negate in CFS_SPEC:
        vals, tag_used = build_row(facts_gaap, tags, fy_ends, negate)
        cfs_rows[label] = vals
        cfs_tags[label] = tag_used
        if tag_used is None:
            missing_cfs.append(label)
    ocf = cfs_rows.get("Operating Cash Flow", [None] * len(fy_ends))
    capex = cfs_rows.get("Capital Expenditures", [None] * len(fy_ends))
    cfs_rows["Free Cash Flow"] = [(o + c if o is not None and c is not None else None) for o, c in zip(ocf, capex)]
    cfs_tags["Free Cash Flow"] = "derived"
    df_cfs = pd.DataFrame(cfs_rows, index=fy_labels).T
    df_cfs.index.name = "Line Item"

    return {
        "ticker": ticker_upper,
        "company_name": company_name,
        "cik": cik,
        "fy_ends": fy_ends,
        "fy_labels": fy_labels,
        "proj_labels": proj_labels,
        "facts_gaap": facts_gaap,
        "df_is": df_is,
        "df_bs": df_bs,
        "df_cfs": df_cfs,
        "is_tags": is_tags,
        "bs_tags": bs_tags,
        "cfs_tags": cfs_tags,
        "missing_is": missing_is,
        "missing_bs": missing_bs,
        "missing_cfs": missing_cfs,
    }


def build_workbook(financials):
    company_name = financials["company_name"]
    ticker = financials["ticker"]
    cik = financials["cik"]
    fy_labels = financials["fy_labels"]
    proj_labels = financials["proj_labels"]
    df_is = financials["df_is"]
    df_bs = financials["df_bs"]
    df_cfs = financials["df_cfs"]
    is_tags = financials["is_tags"]
    bs_tags = financials["bs_tags"]
    cfs_tags = financials["cfs_tags"]

    today = datetime.today().strftime("%Y-%m-%d")
    n_fy = len(fy_labels)
    n_proj = len(proj_labels)
    data_start_col = 3
    data_end_col = data_start_col + n_fy - 1
    proj_start_col = data_end_col + 1

    navy = "1F3864"
    med_gray = "BFBFBF"
    light_gray = "F2F2F2"
    subtotal_c = "E8E8E8"
    alt_row = "FAFAFA"
    proj_fill = "EBF3FB"
    wacc_input = "FFF2CC"
    wacc_rslt = "D9EAD3"
    thin_top = Border(top=Side(style="thin"))

    fmt_dollar = '$#,##0.0;($#,##0.0);"-"'
    fmt_pct = '0.0%;(0.0%);"-"'
    fmt_eps = '$#,##0.00;($#,##0.00);"-"'
    fmt_shares = '#,##0.0;(#,##0.0);"-"'
    fmt_2dp = '#,##0.00'

    subtotals_is = {"Gross Profit", "Operating Income", "Net Income"}
    subtotals_bs = {"Total Current Assets", "Total Assets", "Total Liabilities", "Total Equity", "Total Liabilities & Equity"}
    subtotals_cfs = {"Operating Cash Flow", "Investing Cash Flow", "Financing Cash Flow", "Free Cash Flow"}
    eps_rows = {"EPS Basic (per share)", "EPS Diluted (per share)"}
    shares_rows = {"Diluted Shares"}

    def col(n):
        return get_column_letter(n)

    def set_header(ws, stmt_name):
        last_col = col(proj_start_col + n_proj - 1)
        ws.row_dimensions[1].height = 8
        ws.merge_cells(f"B2:{last_col}2")
        cell = ws["B2"]
        cell.value = f"{company_name} | {stmt_name}"
        cell.font = Font(bold=True, size=14, color="FFFFFF", name="Calibri")
        cell.fill = PatternFill("solid", fgColor=navy)
        cell.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[2].height = 24
        ws["B3"].value = "All figures in $MM unless noted"
        ws["B3"].font = Font(italic=True, size=9, color="808080", name="Calibri")
        ws.row_dimensions[3].height = 14
        hdr_fill = PatternFill("solid", fgColor=med_gray)
        hdr_font = Font(bold=True, size=10, name="Calibri")
        hdr_align = Alignment(horizontal="center")
        for i, fy in enumerate(fy_labels):
            c = ws.cell(row=4, column=data_start_col + i, value=fy)
            c.font = hdr_font
            c.fill = hdr_fill
            c.alignment = hdr_align
        proj_hdr_fill = PatternFill("solid", fgColor="D9D9D9")
        for i, pl in enumerate(proj_labels):
            c = ws.cell(row=4, column=proj_start_col + i, value=pl)
            c.font = Font(bold=True, size=10, name="Calibri", color="595959")
            c.fill = proj_hdr_fill
            c.alignment = hdr_align

    def set_col_widths(ws):
        ws.column_dimensions["A"].width = 3
        ws.column_dimensions["B"].width = 35
        for i in range(n_fy):
            ws.column_dimensions[col(data_start_col + i)].width = 13
        for i in range(n_proj):
            ws.column_dimensions[col(proj_start_col + i)].width = 13

    def write_data_rows(ws, df, start_row, subtotals, section_inserts=None):
        inserts = {pos: lbl for lbl, pos in (section_inserts or [])}
        row_map = {}
        r = start_row
        alt = False
        for idx, label in enumerate(df.index):
            if idx in inserts:
                ws.cell(row=r, column=2, value=inserts[idx]).font = Font(bold=True, size=10, name="Calibri")
                for cn in range(2, proj_start_col + n_proj + 1):
                    ws.cell(row=r, column=cn).fill = PatternFill("solid", fgColor=light_gray)
                ws.row_dimensions[r].height = 16
                r += 1
                alt = False
            is_sub = label in subtotals
            bg = subtotal_c if is_sub else (alt_row if alt else "FFFFFF")
            rfill = PatternFill("solid", fgColor=bg)
            num_fmt = fmt_eps if label in eps_rows else fmt_shares if label in shares_rows else fmt_dollar
            lc = ws.cell(row=r, column=2, value=label)
            lc.font = Font(bold=is_sub, size=10, name="Calibri")
            lc.fill = rfill
            for i, fy in enumerate(fy_labels):
                value = df.loc[label, fy]
                dc = ws.cell(row=r, column=data_start_col + i, value=(float(value) if pd.notna(value) and value is not None else None))
                dc.number_format = num_fmt
                dc.font = Font(color="0000FF", name="Calibri", size=10, bold=is_sub)
                dc.fill = rfill
                dc.alignment = Alignment(horizontal="right")
                if is_sub:
                    dc.border = thin_top
            pfill = PatternFill("solid", fgColor=proj_fill)
            for i in range(n_proj):
                pc = ws.cell(row=r, column=proj_start_col + i)
                pc.number_format = num_fmt
                pc.font = Font(color="000000", name="Calibri", size=10)
                pc.fill = pfill
                pc.alignment = Alignment(horizontal="right")
            row_map[label] = r
            ws.row_dimensions[r].height = 15
            r += 1
            alt = not alt
        return r, row_map

    def write_section_hdr(ws, r, label, fg=light_gray):
        ws.cell(row=r, column=2, value=label).font = Font(bold=True, size=10, name="Calibri")
        for cn in range(2, proj_start_col + n_proj + 1):
            ws.cell(row=r, column=cn).fill = PatternFill("solid", fgColor=fg)
        ws.row_dimensions[r].height = 16

    def write_formula_row(ws, r, label, fn, num_fmt=fmt_pct):
        ws.cell(row=r, column=2, value=label).font = Font(size=10, name="Calibri")
        for i in range(n_fy):
            dc = ws.cell(row=r, column=data_start_col + i, value=fn(r, data_start_col + i))
            dc.number_format = num_fmt
            dc.font = Font(color="000000", name="Calibri", size=10)
            dc.alignment = Alignment(horizontal="right")
        ws.row_dimensions[r].height = 15

    def write_stats_block(ws, start_r, row_map, stat_labels, fmt_map=None):
        write_section_hdr(ws, start_r, "SUMMARY STATISTICS", fg="DEE6EF")
        hdr_r = start_r + 1
        for cn, txt in [(2, "Line Item"), (3, "5-Yr CAGR"), (4, f"5-Yr Avg ({fy_labels[0]}–{fy_labels[-1]})")]:
            c = ws.cell(row=hdr_r, column=cn, value=txt)
            c.font = Font(bold=True, size=9, name="Calibri", color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="2E5F8A")
            c.alignment = Alignment(horizontal="center")
        ws.row_dimensions[hdr_r].height = 14
        c_s = col(data_start_col)
        c_e = col(data_end_col)
        n_p = n_fy - 1 if n_fy > 1 else 1
        r = hdr_r + 1
        alt = False
        for label in stat_labels:
            dr = row_map.get(label)
            if dr is None:
                continue
            fill = PatternFill("solid", fgColor=alt_row if alt else "FFFFFF")
            num_fmt = (fmt_map or {}).get(label, fmt_dollar)
            lc = ws.cell(row=r, column=2, value=label)
            lc.font = Font(size=9, name="Calibri")
            lc.fill = fill
            cc = ws.cell(row=r, column=3, value=f'=IFERROR(({c_e}{dr}/{c_s}{dr})^(1/{n_p})-1,"-")')
            cc.number_format = fmt_pct
            cc.font = Font(bold=True, size=9, name="Calibri")
            cc.fill = fill
            cc.alignment = Alignment(horizontal="center")
            ac = ws.cell(row=r, column=4, value=f'=IFERROR(AVERAGE({c_s}{dr}:{c_e}{dr}),"-")')
            ac.number_format = num_fmt
            ac.font = Font(bold=True, size=9, name="Calibri")
            ac.fill = fill
            ac.alignment = Alignment(horizontal="right")
            ws.row_dimensions[r].height = 14
            r += 1
            alt = not alt
        ws.column_dimensions[col(3)].width = max(ws.column_dimensions[col(3)].width or 0, 12)
        ws.column_dimensions[col(4)].width = max(ws.column_dimensions[col(4)].width or 0, 18)
        return r

    wb = Workbook()
    wb.remove(wb.active)

    cover = wb.create_sheet("Cover")
    cover.column_dimensions["A"].width = 4
    cover.column_dimensions["B"].width = 30
    cover.column_dimensions["C"].width = 44
    cover.merge_cells("B2:D2")
    t = cover["B2"]
    t.value = f"{company_name} | SEC 10-K Financial Model"
    t.font = Font(bold=True, size=16, color="FFFFFF", name="Calibri")
    t.fill = PatternFill("solid", fgColor=navy)
    t.alignment = Alignment(horizontal="left", vertical="center")
    cover.row_dimensions[2].height = 30
    for i, (k, v) in enumerate([
        ("Company", company_name),
        ("Ticker", ticker),
        ("CIK", cik),
        ("Date Pulled", today),
        ("Fiscal Years", ", ".join(fy_labels)),
        ("Forecast Cols", ", ".join(proj_labels)),
        ("Units", "All values in $MM unless per-share"),
        ("Data Source", "https://data.sec.gov (XBRL company facts)"),
    ], start=4):
        cover.cell(row=i, column=2, value=k).font = Font(bold=True, name="Calibri", size=10)
        cover.cell(row=i, column=3, value=v).font = Font(name="Calibri", size=10)
    cover.cell(row=13, column=2, value="Contents").font = Font(bold=True, size=11, name="Calibri")
    for i, sname in enumerate(["Income Statement", "Balance Sheet", "Cash Flow Statement", "WACC", "Raw SEC Facts"], start=14):
        c = cover.cell(row=i, column=2, value=sname)
        c.hyperlink = f"#{sname}!A1"
        c.font = Font(color="0563C1", underline="single", name="Calibri", size=10)

    ws_is = wb.create_sheet("Income Statement")
    set_col_widths(ws_is)
    set_header(ws_is, "Income Statement")
    next_r, is_map = write_data_rows(ws_is, df_is, 5, subtotals_is, [("REVENUE", 0), ("OPERATING EXPENSES", 2), ("PROFITABILITY", 5), ("PER SHARE & OTHER", 10)])
    next_r += 1
    write_section_hdr(ws_is, next_r, "MARGINS & RATIOS")
    next_r += 1
    rv = is_map.get("Revenue")
    gp = is_map.get("Gross Profit")
    oi = is_map.get("Operating Income")
    ni = is_map.get("Net Income")
    rd = is_map.get("R&D Expense")
    sg = is_map.get("SG&A Expense")
    tx = is_map.get("Income Tax Expense")
    pt = is_map.get("Pretax Income")

    def mfn(nr, dr):
        return lambda r, c: f'=IFERROR({col(c)}{nr}/{col(c)}{dr},"-")'

    for lbl, fn in [
        ("Gross Margin %", mfn(gp, rv)),
        ("Operating Margin %", mfn(oi, rv)),
        ("Net Margin %", mfn(ni, rv)),
        ("R&D % of Revenue", mfn(rd, rv)),
        ("SG&A % of Revenue", mfn(sg, rv)),
        ("Effective Tax Rate %", mfn(tx, pt)),
    ]:
        write_formula_row(ws_is, next_r, lbl, fn, fmt_pct)
        next_r += 1
    ws_is.cell(row=next_r, column=2, value="YoY Revenue Growth %").font = Font(size=10, name="Calibri")
    for i in range(n_fy):
        if i == 0:
            continue
        dc = ws_is.cell(row=next_r, column=data_start_col + i, value=f'=IFERROR({col(data_start_col + i)}{rv}/{col(data_start_col + i - 1)}{rv}-1,"-")')
        dc.number_format = fmt_pct
        dc.font = Font(color="000000", name="Calibri", size=10)
        dc.alignment = Alignment(horizontal="right")
    ws_is.row_dimensions[next_r].height = 15
    next_r += 2
    write_stats_block(ws_is, next_r, is_map, ["Revenue", "Gross Profit", "Operating Income", "Net Income", "EPS Diluted (per share)"], {"EPS Diluted (per share)": fmt_eps})

    ws_bs = wb.create_sheet("Balance Sheet")
    set_col_widths(ws_bs)
    set_header(ws_bs, "Balance Sheet")
    next_r_bs, bs_map = write_data_rows(ws_bs, df_bs, 5, subtotals_bs, [("CURRENT ASSETS", 0), ("NON-CURRENT ASSETS", 5), ("CURRENT LIABILITIES", 9), ("NON-CURRENT LIABILITIES", 12), ("EQUITY", 14)])
    next_r_bs += 1
    write_section_hdr(ws_bs, next_r_bs, "KEY RATIOS")
    next_r_bs += 1
    ca = bs_map.get("Total Current Assets")
    cl = bs_map.get("Total Current Liabilities")
    te = bs_map.get("Total Equity")
    std = bs_map.get("Short-Term Debt")
    ltd = bs_map.get("Long-Term Debt")
    csh = bs_map.get("Cash & Equivalents")
    for lbl, fn, fmt in [
        ("Current Ratio", lambda r, c: f'=IFERROR({col(c)}{ca}/{col(c)}{cl},"-")', fmt_2dp),
        ("Debt-to-Equity", lambda r, c: f'=IFERROR(({col(c)}{std}+{col(c)}{ltd})/{col(c)}{te},"-")', fmt_2dp),
        ("Net Debt ($MM)", lambda r, c: f'=IFERROR(({col(c)}{std}+{col(c)}{ltd})-{col(c)}{csh},"-")', fmt_dollar),
    ]:
        write_formula_row(ws_bs, next_r_bs, lbl, fn, fmt)
        next_r_bs += 1
    next_r_bs += 1
    write_stats_block(ws_bs, next_r_bs, bs_map, ["Total Current Assets", "Total Assets", "Total Current Liabilities", "Total Liabilities", "Total Equity"])

    ws_cfs = wb.create_sheet("Cash Flow Statement")
    set_col_widths(ws_cfs)
    set_header(ws_cfs, "Cash Flow Statement")
    next_r_cfs, cfs_map = write_data_rows(ws_cfs, df_cfs, 5, subtotals_cfs, [("OPERATING ACTIVITIES", 0), ("INVESTING ACTIVITIES", 8), ("FINANCING ACTIVITIES", 11)])
    next_r_cfs += 1
    write_section_hdr(ws_cfs, next_r_cfs, "KEY METRICS")
    next_r_cfs += 1
    fcf_r = cfs_map.get("Free Cash Flow")
    cpx_r = cfs_map.get("Capital Expenditures")
    ni_c = cfs_map.get("Net Income (CFS)")
    is_sh = "'Income Statement'"
    for lbl, fn, fmt in [
        ("FCF Margin %", lambda r, c: f'=IFERROR({col(c)}{fcf_r}/{is_sh}!{col(c)}{rv},"-")', fmt_pct),
        ("CapEx % of Revenue", lambda r, c: f'=IFERROR(ABS({col(c)}{cpx_r})/{is_sh}!{col(c)}{rv},"-")', fmt_pct),
        ("FCF Conversion %", lambda r, c: f'=IFERROR({col(c)}{fcf_r}/{col(c)}{ni_c},"-")', fmt_pct),
    ]:
        write_formula_row(ws_cfs, next_r_cfs, lbl, fn, fmt)
        next_r_cfs += 1
    next_r_cfs += 1
    write_stats_block(ws_cfs, next_r_cfs, cfs_map, ["Operating Cash Flow", "Capital Expenditures", "Investing Cash Flow", "Financing Cash Flow", "Free Cash Flow"])

    ws_w = wb.create_sheet("WACC")
    ws_w.column_dimensions["A"].width = 4
    ws_w.column_dimensions["B"].width = 38
    ws_w.column_dimensions["C"].width = 16
    ws_w.column_dimensions["D"].width = 16
    ws_w.column_dimensions["E"].width = 13
    ws_w.column_dimensions["F"].width = 13
    ws_w.column_dimensions["G"].width = 13

    def w_title(r, text):
        ws_w.merge_cells(f"B{r}:G{r}")
        c = ws_w[f"B{r}"]
        c.value = text
        c.font = Font(bold=True, size=14, color="FFFFFF", name="Calibri")
        c.fill = PatternFill("solid", fgColor=navy)
        c.alignment = Alignment(horizontal="left", vertical="center")
        ws_w.row_dimensions[r].height = 26

    def w_sec(r, text):
        ws_w.merge_cells(f"B{r}:G{r}")
        c = ws_w[f"B{r}"]
        c.value = text
        c.font = Font(bold=True, size=10, color="FFFFFF", name="Calibri")
        c.fill = PatternFill("solid", fgColor="2E5F8A")
        ws_w.row_dimensions[r].height = 16

    def w_row(r, label, value=None, formula=None, fmt=fmt_pct, is_input=False, is_result=False):
        lc = ws_w.cell(row=r, column=2, value=label)
        lc.font = Font(bold=is_result, size=10, name="Calibri")
        if is_result:
            lc.fill = PatternFill("solid", fgColor=wacc_rslt)
        vc = ws_w.cell(row=r, column=3, value=formula if formula else value)
        vc.number_format = fmt
        vc.alignment = Alignment(horizontal="center")
        if is_input:
            vc.fill = PatternFill("solid", fgColor=wacc_input)
            vc.font = Font(color="0000FF", size=10, name="Calibri")
        elif is_result:
            vc.fill = PatternFill("solid", fgColor=wacc_rslt)
            vc.font = Font(bold=True, color="375623", size=11, name="Calibri")
        else:
            vc.font = Font(color="000000", size=10, name="Calibri")
        ws_w.row_dimensions[r].height = 15

    w_title(2, f"{company_name} | WACC Buildup")
    ws_w["B3"].value = "Yellow cells = user inputs. Green cell = WACC result (referenced by DCF model)."
    ws_w["B3"].font = Font(italic=True, size=9, color="808080", name="Calibri")
    w_sec(5, "COST OF DEBT")
    w_row(6, "Pre-tax cost of debt (Kd)", value=0.045, fmt=fmt_pct, is_input=True)
    w_row(7, "Marginal tax rate (T)", value=0.21, fmt=fmt_pct, is_input=True)
    w_row(8, "After-tax cost of debt Kd × (1 − T)", formula="=C6*(1-C7)", fmt=fmt_pct)
    w_sec(10, "COST OF EQUITY — CAPM")
    w_row(11, "Risk-free rate (Rf) — 10-yr UST", value=0.043, fmt=fmt_pct, is_input=True)
    w_row(12, "Equity beta (β)", value=1.20, fmt=fmt_2dp, is_input=True)
    w_row(13, "Equity risk premium (ERP)", value=0.055, fmt=fmt_pct, is_input=True)
    w_row(14, "Cost of equity Ke = Rf + β × ERP", formula="=C11+C12*C13", fmt=fmt_pct)
    w_sec(16, "CAPITAL STRUCTURE")
    for cn, txt in [(2, "Component"), (3, "Market Value ($MM)"), (4, "Target Override"), (5, "Weight")]:
        c = ws_w.cell(row=17, column=cn, value=txt)
        c.font = Font(bold=True, size=9, name="Calibri")
        c.fill = PatternFill("solid", fgColor=med_gray)
        c.alignment = Alignment(horizontal="center")
    ws_w.row_dimensions[17].height = 14
    for r_, lbl, we_formula in [
        (18, "Equity (Market Capitalization)", '=IF(D18<>"",D18,C18/(C18+C19))'),
        (19, "Net Debt (Total Debt − Cash)", '=1-E18'),
    ]:
        ws_w.cell(row=r_, column=2, value=lbl).font = Font(size=10, name="Calibri")
        inp = ws_w.cell(row=r_, column=3)
        inp.fill = PatternFill("solid", fgColor=wacc_input)
        inp.font = Font(color="0000FF", size=10, name="Calibri")
        inp.number_format = fmt_dollar
        inp.alignment = Alignment(horizontal="center")
        we = ws_w.cell(row=r_, column=5, value=we_formula)
        we.number_format = fmt_pct
        we.font = Font(color="000000", size=10, name="Calibri")
        we.alignment = Alignment(horizontal="center")
        ws_w.row_dimensions[r_].height = 15
    w_sec(21, "WACC")
    ws_w.cell(row=22, column=2, value="WACC = Ke × We + Kd(1−T) × Wd").font = Font(bold=True, size=11, name="Calibri")
    ws_w.cell(row=22, column=2).fill = PatternFill("solid", fgColor=wacc_rslt)
    wr = ws_w.cell(row=22, column=3, value="=C14*E18+C8*E19")
    wr.number_format = fmt_pct
    wr.font = Font(bold=True, size=13, color="375623", name="Calibri")
    wr.fill = PatternFill("solid", fgColor=wacc_rslt)
    wr.alignment = Alignment(horizontal="center", vertical="center")
    ws_w.row_dimensions[22].height = 22

    w_sec(25, "SENSITIVITY — WACC vs Beta & Equity Risk Premium")
    ws_w.cell(row=26, column=2, value="WACC").font = Font(bold=True, size=9, name="Calibri")
    ws_w.cell(row=26, column=3, value="Beta →").font = Font(bold=True, size=9, name="Calibri", color="595959")
    ws_w.cell(row=27, column=2, value="ERP ↓").font = Font(bold=True, size=9, name="Calibri", color="595959")
    beta_offsets = [-0.30, -0.15, 0.00, 0.15, 0.30]
    erp_offsets = [-0.010, -0.005, 0.000, 0.005, 0.010]
    for j, bo in enumerate(beta_offsets):
        beta_f = "=C12" if bo == 0 else f"={bo:+.2f}+C12"
        c = ws_w.cell(row=26, column=3 + j, value=beta_f)
        c.number_format = fmt_2dp
        c.font = Font(bold=True, size=9, name="Calibri")
        c.fill = PatternFill("solid", fgColor=med_gray)
        c.alignment = Alignment(horizontal="center")
    for i, eo in enumerate(erp_offsets):
        erp_f = "=C13" if eo == 0 else f"={eo:+.3f}+C13"
        er = ws_w.cell(row=27 + i, column=2, value=erp_f)
        er.number_format = fmt_pct
        er.font = Font(bold=True, size=9, name="Calibri")
        er.fill = PatternFill("solid", fgColor=med_gray)
        er.alignment = Alignment(horizontal="center")
        for j, bo in enumerate(beta_offsets):
            beta_expr = "C12" if bo == 0 else f"({bo:+.2f}+C12)"
            erp_expr = "C13" if eo == 0 else f"({eo:+.3f}+C13)"
            formula = f"=(C11+{beta_expr}*{erp_expr})*E18+C8*E19"
            vc = ws_w.cell(row=27 + i, column=3 + j, value=formula)
            vc.number_format = fmt_pct
            vc.font = Font(size=9, name="Calibri")
            vc.alignment = Alignment(horizontal="center")
            if i == 2 and j == 2:
                vc.fill = PatternFill("solid", fgColor=wacc_rslt)
                vc.font = Font(bold=True, size=9, color="375623", name="Calibri")
        ws_w.row_dimensions[27 + i].height = 14

    ws_w.cell(row=34, column=2, value="Notes").font = Font(bold=True, size=9, name="Calibri")
    for i, note in enumerate([
        "• Yellow cells are user inputs. Update Rf, beta, ERP, market cap, and net debt for this company.",
        "• Risk-free rate: current 10-year US Treasury yield.",
        "• ERP: replace with your preferred market assumption.",
        "• Beta: replace with observed or source-verified beta.",
        "• Market cap and net debt: pull from the Balance Sheet tab or live market data.",
        "• Cell C22 (WACC result) can be referenced directly in a DCF model.",
    ], start=35):
        ws_w.cell(row=i, column=2, value=note).font = Font(size=9, name="Calibri", color="595959")
        ws_w.row_dimensions[i].height = 13

    ws_raw = wb.create_sheet("Raw SEC Facts")
    ws_raw.freeze_panes = "A2"
    ws_raw.append(["Statement", "Line Item", "Tag Used"] + fy_labels)
    raw_rows = []
    for stmt, df_map, tag_map in [("Income Statement", df_is, is_tags), ("Balance Sheet", df_bs, bs_tags), ("Cash Flow Statement", df_cfs, cfs_tags)]:
        for label in df_map.index:
            row = [stmt, label, tag_map.get(label, "")]
            for fy in fy_labels:
                v = df_map.loc[label, fy]
                row.append(float(v) if pd.notna(v) and v is not None else None)
            ws_raw.append(row)
            raw_rows.append(row)
    n_raw = len(raw_rows) + 1
    tbl = Table(displayName="RawSECFacts", ref=f"A1:{col(3 + n_fy)}{n_raw}")
    tbl.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    ws_raw.add_table(tbl)
    ws_raw.column_dimensions["A"].width = 20
    ws_raw.column_dimensions["B"].width = 30
    ws_raw.column_dimensions["C"].width = 55
    for i in range(n_fy):
        ws_raw.column_dimensions[col(4 + i)].width = 13

    return wb


def workbook_to_bytes(workbook):
    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    return output


st.title("SEC 10-K Financial Model Builder")
st.write("Convert SEC XBRL company facts into a formatted Excel model workbook.")

with st.sidebar:
    st.header("Inputs")
    ticker = st.text_input("Ticker", value="AAPL").strip().upper()
    user_agent = st.text_input("SEC User-Agent", value="Your Name your@email.com")
    st.caption("Use a real name/email in the User-Agent to follow SEC fair-access guidance.")

if st.button("Build Excel Model", type="primary"):
    if not ticker or not user_agent:
        st.error("Please enter both a ticker and a SEC-compliant User-Agent.")
    else:
        try:
            with st.spinner("Pulling SEC data and building workbook..."):
                financials = build_financial_data(ticker, user_agent)
                workbook = build_workbook(financials)
                excel_bytes = workbook_to_bytes(workbook)
                output_name = f"{financials['ticker']}_10K_Historical_{datetime.today().strftime('%Y-%m-%d')}.xlsx"

            st.success("Workbook created successfully.")

            c1, c2, c3 = st.columns(3)
            c1.metric("Company", financials["company_name"])
            c2.metric("Ticker", financials["ticker"])
            c3.metric("CIK", financials["cik"])

            st.write(f"Fiscal years pulled: {', '.join(financials['fy_labels'])}")
            st.write(f"Forecast columns created: {', '.join(financials['proj_labels'])}")

            missing_summary = pd.DataFrame({
                "Statement": ["Income Statement", "Balance Sheet", "Cash Flow Statement"],
                "Missing line items": [
                    ", ".join(financials["missing_is"]) if financials["missing_is"] else "None",
                    ", ".join(financials["missing_bs"]) if financials["missing_bs"] else "None",
                    ", ".join(financials["missing_cfs"]) if financials["missing_cfs"] else "None",
                ],
            })
            st.subheader("Mapping summary")
            st.dataframe(missing_summary, use_container_width=True)

            tab1, tab2, tab3 = st.tabs(["Income Statement", "Balance Sheet", "Cash Flow Statement"])
            with tab1:
                st.dataframe(financials["df_is"], use_container_width=True)
            with tab2:
                st.dataframe(financials["df_bs"], use_container_width=True)
            with tab3:
                st.dataframe(financials["df_cfs"], use_container_width=True)

            st.download_button(
                label="Download Excel Workbook",
                data=excel_bytes,
                file_name=output_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as exc:
            st.error(f"Error: {exc}")
else:
    st.info("Enter a ticker and SEC User-Agent, then click 'Build Excel Model'.")
