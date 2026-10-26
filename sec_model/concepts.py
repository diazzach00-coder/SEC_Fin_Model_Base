"""Model concepts and the XBRL tags / label patterns that may supply them.

Data only. ``tags`` are written without the ``us-gaap_`` prefix and are tried in
order; ``label_re`` (case-insensitive ``re.search``) is the fallback for custom
tags. ``sign`` is -1 only where the as-reported sign is opposite to the model's
convention. ``sum_groups`` means the concept is the sum of several rows. ``stypes`` lists
the statements searched, in order. ``gate`` names a concept that must be mapped in the same
year before ``gated_tags`` (and any label match outside ``ungated_label_re``) may be used;
a gated-tag row whose label matches ``gated_label_exclude`` is never used.
"""
from dataclasses import dataclass


@dataclass(frozen=True)
class SumGroup:
    """One component of a derived (summed) concept.

    mode "first": use the first tag present (so overlapping tags are not double
    counted); mode "all": add every tag present. The group is skipped entirely
    when any tag in ``skip_if`` supplied an earlier group (it already includes
    this group's items).
    """
    tags: tuple[str, ...]
    mode: str = "first"
    skip_if: tuple[str, ...] = ()


@dataclass(frozen=True)
class Concept:
    name: str
    stypes: tuple[str, ...]          # statements searched, in order: "IS" | "BS" | "CF"
    tags: tuple[str, ...]
    label_re: str | None
    sign: int = 1
    sum_groups: tuple[SumGroup, ...] = ()
    # (trigger tag, tags to add): when the concept resolved via the trigger tag and
    # any of the add-on rows is present in the same statement, the value is their sum.
    add_when: tuple[str, tuple[str, ...]] | None = None
    # (tag, sign) for tags whose XBRL sign is opposite to the concept's convention, e.g. an
    # income-positive "interest, net" element mapped to InterestExpense (expense positive).
    tag_signs: tuple[tuple[str, int], ...] = ()
    # Broad tags that mean this concept only for some filers: accepted only when the
    # ``gate`` concept is mapped (non-None) in the same year; otherwise the concept is
    # re-resolved without them and with ``ungated_label_re`` (mapping._apply_gates).
    gate: str | None = None
    gated_tags: tuple[str, ...] = ()
    ungated_label_re: str | None = None
    # A row under a gated tag whose label matches this is skipped even when the gate is open
    # (LongTermInvestments also holds securities portfolios).
    gated_label_exclude: str | None = None

    def tag_sign(self, tag: str | None) -> int:
        bare = (tag or "").removeprefix("us-gaap_")
        return next((s for t, s in self.tag_signs if t == bare), 1)


# Drawn revolver / line-of-credit tags: the RevolverDrawn concept and TotalDebt's revolver
# group use the same tags in the same first-present order, so both pick the same row.
REVOLVER_TAGS = ("LineOfCredit", "LongTermLineOfCredit", "LinesOfCreditCurrent",
                 "RevolvingCreditFacility")

CONCEPTS: list[Concept] = [
    # ---- Income statement ----
    Concept("Revenue", ("IS",),
            ("Revenues", "RevenueFromContractWithCustomerExcludingAssessedTax",
             "RevenueFromContractWithCustomerIncludingAssessedTax", "SalesRevenueNet"),
            r"^total revenues?"),
    Concept("TotalCostsAndExpenses", ("IS",),
            ("CostsAndExpenses", "OperatingExpenses", "OperatingCostsAndExpenses"),
            r"total costs and expenses"),
    Concept("OperatingIncome", ("IS",), ("OperatingIncomeLoss",), r"operating income"),
    # D&A is on the IS for some filers, only on the CF for others (EPD, ORCL).
    Concept("DA", ("IS", "CF"),
            ("DepreciationDepletionAndAmortization", "DepreciationAndAmortization",
             "DepreciationAmortizationAndAccretionNet", "Depreciation"),
            r"depreciation",
            add_when=("Depreciation", ("AmortizationOfIntangibleAssets",))),
    Concept("Impairments", ("IS",),
            ("AssetImpairmentCharges", "OtherAssetImpairmentCharges", "GoodwillImpairmentLoss",
             "ImpairmentOfLongLivedAssetsHeldForUse"),
            r"impairment"),
    # IS-reported gains are already positive = gain.
    Concept("Gains", ("IS",),
            ("GainLossOnDispositionOfAssets", "GainLossOnDispositionOfAssets1",
             "GainLossOnSaleOfBusiness"),
            r"gain.*(sale|disposition)"),
    Concept("EquityEarnings", ("IS",), ("IncomeLossFromEquityMethodInvestments",),
            r"equity (in )?(net )?(earnings|income)"),
    # KMI reports only "Interest, net" (InterestIncomeExpenseNonoperatingNet, income-positive).
    # MPLX: "Net interest and other financial costs" is InterestAndDebtExpense.
    # JPM's FY2024+ "Interest expense" is InterestExpenseOperating; without it the bank's
    # "Net interest income" (InterestIncomeExpenseNet) was taken as negative interest expense.
    Concept("InterestExpense", ("IS",),
            ("InterestExpense", "InterestExpenseNonoperating", "InterestExpenseDebt",
             "InterestAndDebtExpense", "InterestExpenseOperating",
             "InterestIncomeExpenseNonoperatingNet", "InterestIncomeExpenseNet"),
            r"interest expense",
            tag_signs=(("InterestIncomeExpenseNonoperatingNet", -1),
                       ("InterestIncomeExpenseNet", -1))),
    Concept("PretaxIncome", ("IS",),
            ("IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
             "IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments",
             # newer element ("..., before Tax"); ET uses it from FY2022
             "IncomeLossIncludingPortionAttributableToNoncontrollingInterest"),
            r"before income tax"),
    # The lookahead keeps "Income before income taxes" from matching.
    Concept("IncomeTax", ("IS",), ("IncomeTaxExpenseBenefit",), r"^(?!.*before).*income tax"),
    # Consolidated net income. Filers without ProfitLoss but with NCI (ENB to FY2019) tag
    # it as continuing operations including NCI; NetIncomeLoss excludes NCI, so it is last.
    Concept("NetIncome", ("IS",),
            ("ProfitLoss", "IncomeLossFromContinuingOperationsIncludingPortionAttributableToNoncontrollingInterest",
             "NetIncomeLoss"),
            None),
    # ET (FY2025) and EPD tag NCI income as the "nonredeemable" element. The label fallback
    # skips "Income (Loss), Including Portion Attributable to Noncontrolling Interest" rows
    # (consolidated totals, not the NCI share).
    Concept("NIToNCI", ("IS",),
            ("NetIncomeLossAttributableToNoncontrollingInterest",
             "NetIncomeLossAttributableToNonredeemableNoncontrollingInterest"),
            r"^(?!.*including portion).*noncontrolling"),
    Concept("DilutedUnits", ("IS",),
            ("WeightedAverageNumberOfDilutedSharesOutstanding",
             "WeightedAverageNumberOfSharesOutstandingDiluted",
             "WeightedAverageLimitedPartnershipUnitsOutstandingDiluted"),
            # "per" excluded so "Diluted earnings per common unit" is not taken.
            r"^(?!.*\bper\b).*diluted.*(units?|shares?)"),

    # ---- Cash flow ----
    Concept("OperatingCashFlow", ("CF",),
            ("NetCashProvidedByUsedInOperatingActivities",
             "NetCashProvidedByUsedInOperatingActivitiesContinuingOperations"),
            None),
    Concept("CFNetIncome", ("CF",),
            ("ProfitLoss", "IncomeLossFromContinuingOperationsIncludingPortionAttributableToNoncontrollingInterest",
             "NetIncomeLoss"),
            None),
    Concept("DeferredTax", ("CF",), ("DeferredIncomeTaxExpenseBenefit",), r"deferred (income )?tax"),
    Concept("DistributionsFromAffiliates", ("CF",), ("EquityMethodInvestmentDividendsOrDistributions",),
            r"distributions received"),
    Concept("Capex", ("CF",),
            ("PaymentsToAcquirePropertyPlantAndEquipment", "PaymentsToAcquireProductiveAssets",
             "PaymentsForCapitalImprovements"),
            r"capital expenditures"),
    Concept("InvestingCF", ("CF",),
            ("NetCashProvidedByUsedInInvestingActivities",
             "NetCashProvidedByUsedInInvestingActivitiesContinuingOperations"),
            None),
    # Label fallback excludes NCI rows (TRGP's custom "Distributions to noncontrolling
    # interests" came before "Dividends paid to common shareholders") and admits
    # "dividends paid". mapping._separate_distributions guards the two sharing one row.
    Concept("Distributions", ("CF",),
            ("PaymentsOfDividendsCommonStock", "PaymentsOfDividends",
             "DistributionMadeToLimitedPartnerCashDistributionsPaid",
             "DistributionMadeToMemberOrLimitedPartnerCashDistributionsPaid",
             "PaymentsOfCapitalDistribution"),
            r"^(?!.*(noncontrolling|minority)).*(distributions?|dividends?) (paid|to)"),
    # The label must name a payout: TRGP FY2024 "Contributions from noncontrolling
    # interests" precedes the distributions row.
    Concept("DistributionsToNCI", ("CF",),
            ("PaymentsOfDividendsMinorityInterest", "PaymentsOfDistributionsToAffiliates"),
            r"(distribution|dividend|payment).*(noncontrolling|minority)"),
    Concept("FinancingCF", ("CF",),
            ("NetCashProvidedByUsedInFinancingActivities",
             "NetCashProvidedByUsedInFinancingActivitiesContinuingOperations"),
            None),
    Concept("NetChangeInCash", ("CF",),
            ("CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalentsPeriodIncreaseDecreaseIncludingExchangeRateEffect",
             # filers with no FX line (MPLX) use the "excluding" element for the total
             "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalentsPeriodIncreaseDecreaseExcludingExchangeRateEffect",
             "CashAndCashEquivalentsPeriodIncreaseDecrease"),
            None),
    # Beginning and ending cash share tags; the label regex tells them apart.
    Concept("CFBeginningCash", ("CF",),
            ("CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents",
             "CashAndCashEquivalentsAtCarryingValue", "Cash"),
            r"beginning"),
    Concept("CFEndingCash", ("CF",),
            ("CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents",
             "CashAndCashEquivalentsAtCarryingValue", "Cash"),
            r"end of|ending"),

    # ---- Balance sheet ----
    Concept("Cash", ("BS",),
            ("CashAndCashEquivalentsAtCarryingValue",
             "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents", "Cash",
             "CashAndDueFromBanks"),
            None),
    # ENB: "Trade receivables and unbilled revenues" / "Trade payables and accrued
    # liabilities" use the broader elements; its affiliate rows matched the bare label
    # fallback. Tags reported only by member (EPD, ET) resolve via mapping's member sum.
    # CVX: "Accounts and notes receivable". The label fallback admits a custom "Trade
    # receivables" row, which beats a lone related-party member row (mapping._resolve_one).
    Concept("Receivables", ("BS",),
            ("AccountsReceivableNetCurrent", "ReceivablesNetCurrent",
             "AccountsAndOtherReceivablesNetCurrent", "AccountsNotesAndLoansReceivableNetCurrent"),
            r"^(?!.*(affiliat|related))(accounts |trade )?receivables?\b"),
    Concept("Inventory", ("BS",), ("InventoryNet",), r"^inventor(y|ies)"),
    Concept("NetPPE", ("BS",), ("PropertyPlantAndEquipmentNet",), r"property, plant"),
    # KMI and ENB tag equity investments as LongTermInvestments ("Investments", "Long-term
    # investments"), WMB as Investments. Those elements also hold securities portfolios,
    # so they count only for a filer that reports equity earnings that year, and never on
    # a row labelled as securities.
    Concept("EquityInvestments", ("BS",),
            ("EquityMethodInvestments",
             "InvestmentsInAffiliatesSubsidiariesAssociatesAndJointVentures",
             "LongTermInvestments", "Investments"),
            r"investments in unconsolidated|^investments$",
            gate="EquityEarnings", gated_tags=("LongTermInvestments", "Investments"),
            ungated_label_re=r"investments in unconsolidated",
            gated_label_exclude=r"marketable|securities|available.for.sale"),
    Concept("TotalAssets", ("BS",), ("Assets",), r"^total assets"),
    Concept("Payables", ("BS",),
            ("AccountsPayableCurrent", "AccountsPayableTradeCurrent",
             "AccountsPayableAndOtherAccruedLiabilitiesCurrent",
             "AccountsPayableAndAccruedLiabilitiesCurrent"),
            r"^(?!.*(affiliat|related))accounts payable"),
    Concept("DebtCurrent", ("BS",), ("DebtCurrent", "LongTermDebtCurrent"),
            r"current (portion|maturities)"),
    Concept("LongTermDebt", ("BS",),
            ("LongTermDebtNoncurrent", "LongTermDebt", "LongTermDebtAndCapitalLeaseObligations"),
            r"^long-term debt"),
    # Long-term + current + short-term + drawn revolver, first-present within the "first"
    # groups so overlapping tags are not double counted. A long-term total that already
    # includes current maturities (LongTermDebt, ...IncludingCurrentMaturities) is a
    # long-term component too, so short-term borrowings are still added (CVX, JPM); it then
    # skips the current-maturities group. A separately presented revolver / line of credit
    # (the RevolverDrawn tags) never overlaps the component lines, so it is summed with
    # them; when a total-including tag supplied the long-term group the revolver is assumed
    # to be inside it and is not added.
    Concept("TotalDebt", ("BS",),
            ("LongTermDebt", "LongTermDebtAndCapitalLeaseObligationsIncludingCurrentMaturities"),
            None,
            sum_groups=(
                SumGroup(("LongTermDebtNoncurrent", "LongTermDebtAndCapitalLeaseObligations",
                          "LongTermNotesAndLoans", "OtherLongTermDebtNoncurrent",
                          "LongTermDebt",
                          "LongTermDebtAndCapitalLeaseObligationsIncludingCurrentMaturities")),
                SumGroup(("DebtCurrent", "LongTermDebtCurrent",
                          "LongTermDebtAndCapitalLeaseObligationsCurrent", "NotesPayableCurrent"),
                         skip_if=("LongTermDebt",
                                  "LongTermDebtAndCapitalLeaseObligationsIncludingCurrentMaturities")),
                # DebtCurrent already includes short-term borrowings and commercial paper.
                SumGroup(("ShortTermBorrowings", "CommercialPaper"), mode="all",
                         skip_if=("DebtCurrent",)),
                SumGroup(REVOLVER_TAGS,
                         skip_if=("LongTermDebt",
                                  "LongTermDebtAndCapitalLeaseObligationsIncludingCurrentMaturities")),
            )),
    # Drawn revolver / line of credit; part of TotalDebt. Unmapped -> 0 (no opening revolver).
    # The label fallback skips total / term-debt rows that merely mention the revolver;
    # mapping._guard_revolver drops a year at or above TotalDebt, without TotalDebt, or
    # outside a component-built (SUM) TotalDebt.
    Concept("RevolverDrawn", ("BS",), REVOLVER_TAGS,
            r"^(?!.*(long-term debt|total|notes|senior)).*(revolv|line of credit)"),
    Concept("LiabilitiesAndEquity", ("BS",),
            ("LiabilitiesAndStockholdersEquity", "LiabilitiesAndPartnersCapital"),
            r"^total liabilities and"),
    Concept("EquityCommon", ("BS",), ("StockholdersEquity", "PartnersCapital"), None),
    Concept("NCI", ("BS",), ("MinorityInterest",), None),
    Concept("TotalEquityInclNCI", ("BS",),
            ("StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest",
             "PartnersCapitalIncludingPortionAttributableToNoncontrollingInterest"),
            None),
    Concept("UnitsOutstanding", ("BS",),
            ("CommonStockSharesOutstanding", "LimitedPartnersCapitalAccountUnitsOutstanding"),
            None),
]
