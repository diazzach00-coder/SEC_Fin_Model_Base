"""Model concepts and the XBRL tags / label patterns that may supply them.

Data only. ``tags`` are written without the ``us-gaap_`` prefix and are tried in
order; ``label_re`` (case-insensitive ``re.search``) is the fallback for custom
tags. ``sign`` is -1 only where the as-reported sign is opposite to the model's
convention. ``sum_groups`` means the concept is the sum of several rows. ``stypes`` lists
the statements searched, in order.
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
    Concept("InterestExpense", ("IS",),
            ("InterestExpense", "InterestExpenseNonoperating", "InterestExpenseDebt"),
            r"interest expense"),
    Concept("PretaxIncome", ("IS",),
            ("IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
             "IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments"),
            r"before income tax"),
    # The lookahead keeps "Income before income taxes" from matching.
    Concept("IncomeTax", ("IS",), ("IncomeTaxExpenseBenefit",), r"^(?!.*before).*income tax"),
    # Consolidated net income. Filers without ProfitLoss but with NCI (ENB to FY2019) tag
    # it as continuing operations including NCI; NetIncomeLoss excludes NCI, so it is last.
    Concept("NetIncome", ("IS",),
            ("ProfitLoss", "IncomeLossFromContinuingOperationsIncludingPortionAttributableToNoncontrollingInterest",
             "NetIncomeLoss"),
            None),
    Concept("NIToNCI", ("IS",), ("NetIncomeLossAttributableToNoncontrollingInterest",),
            r"noncontrolling"),
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
    Concept("Distributions", ("CF",),
            ("PaymentsOfDividendsCommonStock", "PaymentsOfDividends",
             "DistributionMadeToLimitedPartnerCashDistributionsPaid",
             "DistributionMadeToMemberOrLimitedPartnerCashDistributionsPaid",
             "PaymentsOfCapitalDistribution"),
            r"distributions? (paid|to)"),
    Concept("DistributionsToNCI", ("CF",),
            ("PaymentsOfDividendsMinorityInterest", "PaymentsOfDistributionsToAffiliates"),
            r"noncontrolling"),
    Concept("FinancingCF", ("CF",),
            ("NetCashProvidedByUsedInFinancingActivities",
             "NetCashProvidedByUsedInFinancingActivitiesContinuingOperations"),
            None),
    Concept("NetChangeInCash", ("CF",),
            ("CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalentsPeriodIncreaseDecreaseIncludingExchangeRateEffect",
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
    Concept("Receivables", ("BS",), ("AccountsReceivableNetCurrent", "ReceivablesNetCurrent"),
            r"^(accounts )?receivables?\b"),
    Concept("Inventory", ("BS",), ("InventoryNet",), r"^inventor(y|ies)"),
    Concept("NetPPE", ("BS",), ("PropertyPlantAndEquipmentNet",), r"property, plant"),
    Concept("EquityInvestments", ("BS",),
            ("EquityMethodInvestments",
             "InvestmentsInAffiliatesSubsidiariesAssociatesAndJointVentures"),
            r"investments in unconsolidated"),
    Concept("TotalAssets", ("BS",), ("Assets",), r"^total assets"),
    Concept("Payables", ("BS",), ("AccountsPayableCurrent", "AccountsPayableTradeCurrent"),
            r"^accounts payable"),
    Concept("DebtCurrent", ("BS",), ("DebtCurrent", "LongTermDebtCurrent"),
            r"current (portion|maturities)"),
    Concept("LongTermDebt", ("BS",),
            ("LongTermDebtNoncurrent", "LongTermDebt", "LongTermDebtAndCapitalLeaseObligations"),
            r"^long-term debt"),
    # Noncurrent + current + short-term, first-present within the first two groups so
    # overlapping tags are not double counted; else the first of ``tags`` (total
    # debt reported as one line).
    Concept("TotalDebt", ("BS",),
            ("LongTermDebt", "LongTermDebtAndCapitalLeaseObligationsIncludingCurrentMaturities"),
            None,
            sum_groups=(
                SumGroup(("LongTermDebtNoncurrent", "LongTermDebtAndCapitalLeaseObligations",
                          "LongTermNotesAndLoans", "OtherLongTermDebtNoncurrent")),
                SumGroup(("DebtCurrent", "LongTermDebtCurrent",
                          "LongTermDebtAndCapitalLeaseObligationsCurrent", "NotesPayableCurrent")),
                # DebtCurrent already includes short-term borrowings and commercial paper.
                SumGroup(("ShortTermBorrowings", "CommercialPaper"), mode="all",
                         skip_if=("DebtCurrent",)),
            )),
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
