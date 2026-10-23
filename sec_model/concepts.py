"""Model concepts and the XBRL tags / label patterns that may supply them.

Data only. ``tags`` are written without the ``us-gaap_`` prefix and are tried in
order; ``label_re`` (case-insensitive ``re.search``) is the fallback for custom
tags. ``sign`` is -1 only where the as-reported sign is opposite to the model's
convention. ``sum_tags`` means the concept is the sum of several rows.
"""
from dataclasses import dataclass


@dataclass(frozen=True)
class Concept:
    name: str
    stype: str                       # "IS" | "BS" | "CF": the only statement searched
    tags: tuple[str, ...]
    label_re: str | None
    sign: int = 1
    sum_tags: tuple[str, ...] = ()


CONCEPTS: list[Concept] = [
    # ---- Income statement ----
    Concept("Revenue", "IS",
            ("Revenues", "RevenueFromContractWithCustomerExcludingAssessedTax",
             "RevenueFromContractWithCustomerIncludingAssessedTax", "SalesRevenueNet"),
            r"^total revenues?"),
    Concept("TotalCostsAndExpenses", "IS",
            ("CostsAndExpenses", "OperatingExpenses", "OperatingCostsAndExpenses"),
            r"total costs and expenses"),
    Concept("OperatingIncome", "IS", ("OperatingIncomeLoss",), r"operating income"),
    Concept("DA", "IS",
            ("DepreciationDepletionAndAmortization", "DepreciationAndAmortization",
             "DepreciationAmortizationAndAccretionNet", "Depreciation"),
            r"depreciation"),
    Concept("Impairments", "IS",
            ("AssetImpairmentCharges", "OtherAssetImpairmentCharges", "GoodwillImpairmentLoss",
             "ImpairmentOfLongLivedAssetsHeldForUse"),
            r"impairment"),
    # IS-reported gains are already positive = gain.
    Concept("Gains", "IS",
            ("GainLossOnDispositionOfAssets", "GainLossOnDispositionOfAssets1",
             "GainLossOnSaleOfBusiness"),
            r"gain.*(sale|disposition)"),
    Concept("EquityEarnings", "IS", ("IncomeLossFromEquityMethodInvestments",),
            r"equity (in )?(net )?(earnings|income)"),
    Concept("InterestExpense", "IS",
            ("InterestExpense", "InterestExpenseNonoperating", "InterestExpenseDebt"),
            r"interest expense"),
    Concept("PretaxIncome", "IS",
            ("IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
             "IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments"),
            r"before income tax"),
    # The lookahead keeps "Income before income taxes" from matching.
    Concept("IncomeTax", "IS", ("IncomeTaxExpenseBenefit",), r"^(?!.*before).*income tax"),
    Concept("NetIncome", "IS", ("ProfitLoss", "NetIncomeLoss"), None),
    Concept("NIToNCI", "IS", ("NetIncomeLossAttributableToNoncontrollingInterest",),
            r"noncontrolling"),
    Concept("DilutedUnits", "IS",
            ("WeightedAverageNumberOfDilutedSharesOutstanding",
             "WeightedAverageLimitedPartnershipUnitsOutstandingDiluted"),
            None),

    # ---- Cash flow ----
    Concept("OperatingCashFlow", "CF",
            ("NetCashProvidedByUsedInOperatingActivities",
             "NetCashProvidedByUsedInOperatingActivitiesContinuingOperations"),
            None),
    Concept("CFNetIncome", "CF", ("ProfitLoss", "NetIncomeLoss"), None),
    Concept("DeferredTax", "CF", ("DeferredIncomeTaxExpenseBenefit",), r"deferred (income )?tax"),
    Concept("DistributionsFromAffiliates", "CF", ("EquityMethodInvestmentDividendsOrDistributions",),
            r"distributions received"),
    Concept("Capex", "CF",
            ("PaymentsToAcquirePropertyPlantAndEquipment", "PaymentsToAcquireProductiveAssets",
             "PaymentsForCapitalImprovements"),
            r"capital expenditures"),
    Concept("InvestingCF", "CF",
            ("NetCashProvidedByUsedInInvestingActivities",
             "NetCashProvidedByUsedInInvestingActivitiesContinuingOperations"),
            None),
    Concept("Distributions", "CF",
            ("PaymentsOfDividendsCommonStock", "PaymentsOfDividends",
             "DistributionMadeToLimitedPartnerCashDistributionsPaid",
             "DistributionMadeToMemberOrLimitedPartnerCashDistributionsPaid",
             "PaymentsOfCapitalDistribution"),
            r"distributions? (paid|to)"),
    Concept("DistributionsToNCI", "CF",
            ("PaymentsOfDividendsMinorityInterest", "PaymentsOfDistributionsToAffiliates"),
            r"noncontrolling"),
    Concept("FinancingCF", "CF",
            ("NetCashProvidedByUsedInFinancingActivities",
             "NetCashProvidedByUsedInFinancingActivitiesContinuingOperations"),
            None),
    Concept("NetChangeInCash", "CF",
            ("CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalentsPeriodIncreaseDecreaseIncludingExchangeRateEffect",
             "CashAndCashEquivalentsPeriodIncreaseDecrease"),
            None),
    # Beginning and ending cash share tags; the label regex tells them apart.
    Concept("CFBeginningCash", "CF",
            ("CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents",
             "CashAndCashEquivalentsAtCarryingValue", "Cash"),
            r"beginning"),
    Concept("CFEndingCash", "CF",
            ("CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents",
             "CashAndCashEquivalentsAtCarryingValue", "Cash"),
            r"end of|ending"),

    # ---- Balance sheet ----
    Concept("Cash", "BS",
            ("CashAndCashEquivalentsAtCarryingValue",
             "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents", "Cash"),
            None),
    Concept("Receivables", "BS", ("AccountsReceivableNetCurrent", "ReceivablesNetCurrent"),
            r"^(accounts )?receivables?\b"),
    Concept("Inventory", "BS", ("InventoryNet",), r"^inventor(y|ies)"),
    Concept("NetPPE", "BS", ("PropertyPlantAndEquipmentNet",), r"property, plant"),
    Concept("EquityInvestments", "BS",
            ("EquityMethodInvestments",
             "InvestmentsInAffiliatesSubsidiariesAssociatesAndJointVentures"),
            r"investments in unconsolidated"),
    Concept("TotalAssets", "BS", ("Assets",), r"^total assets"),
    Concept("Payables", "BS", ("AccountsPayableCurrent", "AccountsPayableTradeCurrent"),
            r"^accounts payable"),
    Concept("DebtCurrent", "BS", ("DebtCurrent", "LongTermDebtCurrent"),
            r"current (portion|maturities)"),
    Concept("LongTermDebt", "BS",
            ("LongTermDebtNoncurrent", "LongTermDebt", "LongTermDebtAndCapitalLeaseObligations"),
            r"^long-term debt"),
    # Sum of components when any is present; otherwise the first of ``tags``.
    Concept("TotalDebt", "BS",
            ("LongTermDebt", "LongTermDebtAndCapitalLeaseObligations"), None,
            sum_tags=("LongTermDebtNoncurrent", "LongTermDebtCurrent", "DebtCurrent",
                      "CommercialPaper", "ShortTermBorrowings")),
    Concept("LiabilitiesAndEquity", "BS",
            ("LiabilitiesAndStockholdersEquity", "LiabilitiesAndPartnersCapital"),
            r"^total liabilities and"),
    Concept("EquityCommon", "BS", ("StockholdersEquity", "PartnersCapital"), None),
    Concept("NCI", "BS", ("MinorityInterest",), None),
    Concept("TotalEquityInclNCI", "BS",
            ("StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest",
             "PartnersCapitalIncludingPortionAttributableToNoncontrollingInterest"),
            None),
    Concept("UnitsOutstanding", "BS",
            ("CommonStockSharesOutstanding", "LimitedPartnersCapitalAccountUnitsOutstanding"),
            None),
]
