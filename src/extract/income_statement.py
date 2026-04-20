"""Extract quarterly income statement from SEC EDGAR.

Philosophy: discover every IS concept the company actually reports in XBRL,
classify it into a section (Revenue / Costs / Non-op / Taxes / Net income),
and present it in SEC statement-of-operations order. The Excel writer then
renders exactly what SEC has — no hand-mapping, no erroneous data.
"""
from __future__ import annotations
from dataclasses import dataclass, field
from datetime import date
from typing import Dict, List, Optional, Tuple

from .periods import Period, parse_date, classify_duration


# ---------- Section classification ----------

# Concepts we ALWAYS want, in priority order, to establish fixed skeleton rows.
CORE_CONCEPTS = {
    "total_revenue": ["Revenues", "RevenueFromContractWithCustomerExcludingAssessedTax",
                      "RevenueFromContractWithCustomerIncludingAssessedTax", "SalesRevenueNet"],
    "total_opex": ["CostsAndExpenses", "OperatingExpenses"],
    "operating_income": ["OperatingIncomeLoss"],
    "pre_tax_income": ["IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
                       "IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments"],
    "tax_provision": ["IncomeTaxExpenseBenefit"],
    "net_income": ["NetIncomeLoss", "ProfitLoss"],
    "eps_basic": ["EarningsPerShareBasic"],
    "eps_diluted": ["EarningsPerShareDiluted"],
    "shares_basic": ["WeightedAverageNumberOfSharesOutstandingBasic"],
    "shares_diluted": ["WeightedAverageNumberOfDilutedSharesOutstanding"],
}

# Balance sheet / cash flow for the DCF and company-financials block.
BS_CF_CONCEPTS = {
    "cash_and_equivalents": ["CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents",
                             "CashAndCashEquivalentsAtCarryingValue"],
    "long_term_debt": ["LongTermDebt", "LongTermDebtNoncurrent"],
    "short_term_debt": ["ShortTermBorrowings", "LongTermDebtCurrent"],
    "d_and_a": ["DepreciationDepletionAndAmortization", "DepreciationAndAmortization", "Depreciation"],
    "capex": ["PaymentsToAcquirePropertyPlantAndEquipment", "PaymentsToAcquireProductiveAssets"],
    "operating_cash_flow": ["NetCashProvidedByUsedInOperatingActivities"],
    "shares_outstanding": ["CommonStockSharesOutstanding", "EntityCommonStockSharesOutstanding"],
}

# Patterns to EXCLUDE from IS auto-discovery (these aren't statement-of-operations items).
EXCLUDE_PATTERNS = [
    # OCI / comprehensive income
    "Comprehensive", "OtherComprehensive",
    "Oci", "Aoci", "ReclassificationFromAoci", "OciBeforeReclassifications",
    "ForeignCurrencyTransaction",  # OCI-only unless ending at face of IS
    # Stock-based comp detail (EPS calc, not IS face)
    "StockIssuedDuring", "StockRepurchase",
    "ShareBasedCompensationAllocation", "AllocatedShareBasedCompensation",
    "SharebasedCompensationArrangement", "EmployeeServiceShareBased",
    "ShareBasedCompensation",
    "StockBased", "ShareBasedPayment",
    # Share count adjustments and EPS mechanics
    "Adjustments", "AntidilutiveSecurities",
    "Participating", "IncrementalCommon",
    "DilutiveSecurities", "DilutedEarningsPer",
    "NetIncomeLossAttributableToNoncontrolling",
    "NetIncomeLossAttributableToParent",
    "NetIncomeLossAvailableToCommonStockholders",
    # Deferred tax / BS items
    "DeferredTaxAssets", "DeferredTaxLiabilities",
    "FiniteLivedIntangible", "GoodwillImpairment",
    "PriorPeriod", "ChangeInUnrealized",
    "EquityComponent", "StockholdersEquity",
    "AccruedLiabilities", "AccountsPayable",
    # Cash flow statement items
    "IncreaseDecrease", "PaymentsTo", "PaymentsFor", "PaymentsRelatedTo",
    "ProceedsFrom",
    "NetCashProvidedByUsedIn", "RepaymentsOf",
    "NetCashProvidedByUsed",
    "CashCashEquivalents",
    "OtherNoncashIncomeExpense", "OtherNoncash",
    "OperatingLease", "Lease",
    # Depreciation normally reported on CF, not IS face
    "Depreciation",
]

# Hand-curated display labels for common us-gaap concepts. For unknowns we
# auto-format (CamelCase → "Title case").
LABEL_OVERRIDES = {
    "Revenues": "Total revenue",
    "RevenueFromContractWithCustomerExcludingAssessedTax": "Total revenue",
    "RevenueFromContractWithCustomerIncludingAssessedTax": "Total revenue",
    "SalesRevenueNet": "Total revenue",
    "RevenueNotFromContractWithCustomer": "Other revenue",
    "RevenueFromRelatedParties": "Related-party revenue",
    "CostOfRevenue": "Cost of revenue",
    "CostOfGoodsAndServicesSold": "Cost of revenue",
    "CostOfGoodsSold": "Cost of goods sold",
    "CostOfServices": "Cost of services",
    "GrossProfit": "Gross profit",
    "ResearchAndDevelopmentExpense": "Technology and development",
    "SellingAndMarketingExpense": "Sales and marketing",
    "MarketingExpense": "Marketing",
    "SellingExpense": "Selling",
    "GeneralAndAdministrativeExpense": "General and administrative",
    "SellingGeneralAndAdministrativeExpense": "Selling, general and administrative",
    "InformationTechnologyAndDataProcessing": "Transaction expense",  # COIN-style
    "TechnologyServicesCosts": "Technology services costs",
    "RestructuringCharges": "Restructuring",
    "RestructuringAndRelatedCostIncurredCost": "Restructuring",
    "AssetImpairmentCharges": "Asset impairment",
    "OtherOperatingIncomeExpenseNet": "Other operating (income) expense, net",
    "CryptoAssetRealizedAndUnrealizedGainLossOperating": "Gains on crypto assets held for operations, net",
    "CryptoAssetRealizedAndUnrealizedGainLossNonoperating": "Gains on crypto assets held for investment, net",
    "CryptoAssetImpairment": "Crypto asset impairment, net",
    "CryptoAssetImpairmentLoss": "Crypto asset impairment, net",
    "OperatingExpenses": "Total costs and expenses",
    "CostsAndExpenses": "Total costs and expenses",
    "OperatingIncomeLoss": "Operating income (loss)",
    "InterestExpense": "Interest expense",
    "InterestExpenseDebt": "Interest expense",
    "InterestExpenseNonoperating": "Interest expense",
    "InterestIncomeOperating": "Interest income",
    "InterestIncomeNonoperating": "Interest income",
    "OtherNonoperatingIncomeExpense": "Other expense, net",
    "NonoperatingIncomeExpense": "Other non-operating income (expense)",
    "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest": "Income (loss) before income taxes",
    "IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments": "Income (loss) before income taxes",
    "IncomeTaxExpenseBenefit": "Provision for (benefit from) income taxes",
    "NetIncomeLoss": "Net income (loss)",
    "ProfitLoss": "Net income (loss)",
    "EarningsPerShareBasic": "EPS - Basic",
    "EarningsPerShareDiluted": "EPS - Diluted",
    "WeightedAverageNumberOfSharesOutstandingBasic": "Shares Outstanding - Basic",
    "WeightedAverageNumberOfDilutedSharesOutstanding": "Shares Outstanding - Diluted",
    # Balance sheet / cash flow
    "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents": "Cash and cash equivalents",
    "CashAndCashEquivalentsAtCarryingValue": "Cash and cash equivalents",
    "LongTermDebtNoncurrent": "Long-term debt",
    "LongTermDebt": "Long-term debt",
    "DepreciationDepletionAndAmortization": "Depreciation and amortization",
    "PaymentsToAcquirePropertyPlantAndEquipment": "CapEx",
    "NetCashProvidedByUsedInOperatingActivities": "Cash from operations",
}


def _display_label(concept: str) -> str:
    if concept in LABEL_OVERRIDES:
        return LABEL_OVERRIDES[concept]
    # Fallback: split CamelCase and title-case
    import re
    spaced = re.sub(r"(?<!^)(?=[A-Z][a-z])|(?<=[a-z])(?=[A-Z])", " ", concept)
    return spaced.capitalize()


def _looks_like_is_concept(concept: str) -> bool:
    """Return True if the concept looks like an income-statement item.

    Strategy: exclude anything clearly BS/CF/equity/comprehensive,
    accept the rest. Companies tag industry-specific cost lines with
    names that don't contain 'Expense' (e.g. InformationTechnologyAndDataProcessing
    for COIN's transaction costs), so a positive whitelist is too narrow.
    """
    for bad in EXCLUDE_PATTERNS:
        if bad in concept:
            return False
    # Always-excluded specific concepts
    hard_exclude = {
        "Goodwill", "Inventory", "Inventories",
        "AccountsReceivable", "AccountsReceivableNet",
        "PropertyPlantAndEquipment", "IntangibleAssets",
        "DerivativeCollateral", "SafeguardingAsset", "SafeguardingLiability",
        "CryptoAssetFairValue", "CryptoAssetCost", "CryptoAssetAddition",
        "CryptoAssetSale", "CryptoAssetDisposition",
        "TemporaryEquity", "CommonStockShares", "PreferredStock",
        "RetainedEarnings", "AdditionalPaidInCapital",
        # EPS computation detail (already captured by EarningsPerShare)
        "NetIncomeLossAvailableToCommonStockholders",
        "InterestOnConvertibleDebtNetOfTax",
        "ParticipatingSecurities",
        # Hedging / fair-value items that live in OCI rather than IS face
        "FairValueHedge", "CashFlowHedge",
        "GainLossOnFairValueHedges",
    }
    for bad in hard_exclude:
        if bad in concept:
            return False
    return True


# ---------- Data containers ----------

@dataclass
class ISLine:
    """One row on the income statement."""
    concept: str                    # us-gaap concept tag
    section: str                    # "revenue" | "cost" | "opex" | "operating" | "nonop" | "tax" | "net_income" | "eps" | "shares"
    label: str                      # display label
    values: Dict[str, float]        # period_label → value (raw, not divided)
    is_total: bool = False          # bold / separator


@dataclass
class IncomeStatement:
    ticker: str
    cik: str
    title: str
    latest_fy: int
    latest_reported_quarter: Optional[Period]
    periods: List[Period]
    lines: List[ISLine] = field(default_factory=list)
    balance_sheet: Dict[str, Dict[str, float]] = field(default_factory=dict)
    cash_flow: Dict[str, Dict[str, float]] = field(default_factory=dict)

    def line(self, concept_or_section: str) -> Optional[ISLine]:
        for l in self.lines:
            if l.concept == concept_or_section or l.section == concept_or_section:
                return l
        return None

    def core_value(self, section: str, period_label: str) -> Optional[float]:
        """Get the main value for a 'core' section (e.g. total_revenue → first revenue total)."""
        for l in self.lines:
            if l.section == section and l.is_total:
                return l.values.get(period_label)
        return None


# ---------- Fact matching ----------

def _period_for_fact(fact: dict, instant: bool = False) -> Optional[Period]:
    if "end" not in fact:
        return None
    end = parse_date(fact["end"])
    if instant or not fact.get("start"):
        fp = fact.get("fp", "")
        # Fall back to end.year when `fy` is missing OR explicitly None.
        # SEC data sometimes has fy=null for older balance-sheet facts.
        fy = fact.get("fy") or end.year
        frame = fact.get("frame") or ""
        if fp == "FY" or frame.endswith("Q4I") or (end.month == 12 and end.day >= 28):
            return Period(fiscal_year=fy, quarter=None, start=date(fy, 1, 1), end=end)
        q_month = end.month
        q = ((q_month - 1) // 3) + 1
        if q_month == 1 and end.day <= 15:
            q = 4
            return Period(fiscal_year=end.year - 1, quarter=q,
                          start=date(end.year - 1, 10, 1), end=end)
        q_start_month = (q - 1) * 3 + 1
        return Period(fiscal_year=end.year, quarter=q,
                      start=date(end.year, q_start_month, 1), end=end)

    start = parse_date(fact["start"])
    dur = classify_duration(start, end)
    if dur is None:
        return None
    if dur == "annual":
        return Period(fiscal_year=end.year, quarter=None, start=start, end=end)
    q_month = end.month
    q = ((q_month - 1) // 3) + 1
    if q_month == 1 and end.day <= 15:
        q = 4
        return Period(fiscal_year=end.year - 1, quarter=q, start=start, end=end)
    return Period(fiscal_year=end.year, quarter=q, start=start, end=end)


def _pick_best_fact(facts: List[dict], period: Period, instant: bool = False) -> Optional[float]:
    target = period
    best = None
    best_priority = float("inf")
    for f in facts:
        p = _period_for_fact(f, instant=instant)
        if p is None:
            continue
        if p.is_annual != target.is_annual:
            continue
        match = (p.fiscal_year == target.fiscal_year and
                 (target.is_annual or p.quarter == target.quarter))
        if not match:
            continue
        priority = 0
        if "frame" in f:
            priority -= 2
        if target.is_annual and f.get("form") == "10-K":
            priority -= 1
        if not target.is_annual and f.get("form") == "10-Q":
            priority -= 1
        if priority < best_priority:
            best_priority = priority
            best = f["val"]
    return best


def _extract_values(us_gaap: dict, concept: str, periods: List[Period],
                    units: List[str], instant: bool = False) -> Dict[str, float]:
    if concept not in us_gaap:
        return {}
    available_units = us_gaap[concept].get("units", {})
    facts = None
    for u in units:
        if u in available_units:
            facts = available_units[u]
            break
    if facts is None and available_units:
        facts = next(iter(available_units.values()))
    if not facts:
        return {}
    values = {}
    for p in periods:
        v = _pick_best_fact(facts, p, instant=instant)
        if v is not None:
            values[p.label] = v
    return values


def _extract_first_matching(us_gaap: dict, concepts: List[str], periods: List[Period],
                            units: List[str], instant: bool = False) -> Tuple[Optional[str], Dict[str, float]]:
    """Try each concept in order, return the first with recent data."""
    recent_labels = {p.label for p in periods[-4:]}
    best_name, best_values, best_recent = None, {}, -1
    for concept in concepts:
        values = _extract_values(us_gaap, concept, periods, units, instant)
        if not values:
            continue
        recent_count = sum(1 for k in values if k in recent_labels)
        if recent_count > best_recent:
            best_name, best_values, best_recent = concept, values, recent_count
            if recent_count >= 3:
                break
    return best_name, best_values


def _derive_q4(values: Dict[str, float], fy: int) -> None:
    fy_key = str(fy)
    q_labels = [f"{q}Q{str(fy)[-2:]}" for q in (1, 2, 3, 4)]
    q4 = q_labels[3]
    if q4 in values or fy_key not in values:
        return
    parts = [values.get(q) for q in q_labels[:3]]
    if any(p is None for p in parts):
        return
    values[q4] = values[fy_key] - sum(parts)  # type: ignore


def _latest_fy(us_gaap: dict) -> int:
    """Find the most recent FY with reported revenue across any revenue concept.

    Companies migrate XBRL tags (e.g. MP switched from `Revenues` to
    `RevenueFromContractWithCustomerExcludingAssessedTax` in 2022), so we
    scan all candidate revenue concepts.
    """
    candidates = [
        "Revenues",
        "RevenueFromContractWithCustomerExcludingAssessedTax",
        "RevenueFromContractWithCustomerIncludingAssessedTax",
        "SalesRevenueNet",
    ]
    fys: List[int] = []
    for concept in candidates:
        if concept not in us_gaap:
            continue
        usd = us_gaap[concept].get("units", {}).get("USD", [])
        fys.extend(f["fy"] for f in usd if f.get("fp") == "FY" and f.get("fy"))
    return max(fys) if fys else date.today().year - 1


def _classify_concept(concept: str) -> str:
    """Return section label for a concept."""
    c = concept
    # Order matters — most specific first
    if c in ("Revenues", "RevenueFromContractWithCustomerExcludingAssessedTax",
             "RevenueFromContractWithCustomerIncludingAssessedTax", "SalesRevenueNet"):
        return "revenue_total"
    if c.startswith("Revenue") or c.endswith("Revenue") or c.endswith("Revenues"):
        return "revenue"
    # Insurance-specific revenue components
    if "PremiumsEarned" in c or c.startswith("DirectPremium") or c.startswith("AssumedPremium") \
            or c == "NetInvestmentIncome":
        return "revenue"
    if c in ("CostsAndExpenses", "OperatingExpenses"):
        return "opex_total"
    if c in ("CostOfRevenue", "CostOfGoodsAndServicesSold", "CostOfServices", "CostOfGoodsSold"):
        return "cost_of_revenue"
    if c == "GrossProfit":
        return "gross_profit"
    if c == "OperatingIncomeLoss":
        return "operating_income"
    if c in ("IncomeTaxExpenseBenefit",):
        return "tax"
    if "IncomeLossFromContinuingOperationsBeforeIncomeTaxes" in c:
        return "pre_tax"
    if c in ("NetIncomeLoss", "ProfitLoss"):
        return "net_income"
    if "EarningsPerShare" in c:
        return "eps"
    if "WeightedAverageNumber" in c and "Shares" in c:
        return "shares"
    # Nonop detection
    if "Nonoperating" in c or c.startswith("InterestExpense") or c.startswith("InterestIncome") \
            or (("GainLoss" in c or "Gain" in c or "Loss" in c) and "Nonoperating" in c):
        return "nonop"
    # Equity / marketable securities gains/losses → non-operating
    if c.startswith("EquitySecurities") or c.startswith("DebtSecurities") \
            or c.startswith("MarketableSecurities") \
            or c.startswith("DerivativeInstrumentsNotDesignated") \
            or "FairValueHedg" in c:
        return "nonop"
    # Operating adjustments (gains/losses operating, other operating)
    if "OperatingIncomeExpenseNet" in c or "OtherOperating" in c \
            or ("GainLoss" in c and "Operating" in c):
        return "opex_adjustment"
    # Default: treat unknown IS concepts as operating expense items
    # (e.g. COIN's InformationTechnologyAndDataProcessing = Transaction expense)
    return "opex"


SECTION_ORDER = [
    "revenue",            # sub-revenue lines
    "revenue_total",      # Total revenue row
    "cost_of_revenue",    # COGS
    "gross_profit",
    "opex",               # individual opex lines
    "opex_adjustment",    # gains/losses operating, other operating
    "opex_total",         # Total costs and expenses
    "operating_income",
    "nonop",              # interest, other non-operating
    "pre_tax",
    "tax",
    "net_income",
    "eps",
    "shares",
]


# ---------- Main extraction ----------

def extract_income_statement(ticker: str, cik: str, title: str, company_facts: dict,
                             quarters_back: int = 12) -> IncomeStatement:
    us_gaap = company_facts.get("facts", {}).get("us-gaap", {})
    latest_fy = _latest_fy(us_gaap)

    from .periods import build_period_axis
    periods = build_period_axis(latest_fy, quarters_back)
    recent_labels = {p.label for p in periods[-4:]}

    stmt = IncomeStatement(
        ticker=ticker.upper(),
        cik=cik,
        title=title,
        latest_fy=latest_fy,
        latest_reported_quarter=Period(latest_fy, None,
                                       date(latest_fy, 1, 1), date(latest_fy, 12, 31)),
        periods=periods,
    )

    # 1. Extract CORE concepts (always shown)
    core_extracted: Dict[str, Tuple[str, Dict[str, float]]] = {}
    for key, candidates in CORE_CONCEPTS.items():
        if key in ("eps_basic", "eps_diluted"):
            units = ["USD/shares"]
        elif key in ("shares_basic", "shares_diluted"):
            units = ["shares"]
        else:
            units = ["USD"]
        concept_name, values = _extract_first_matching(us_gaap, candidates, periods, units)
        if values:
            # Q4 derivation for duration items only
            if key not in ("eps_basic", "eps_diluted", "shares_basic", "shares_diluted"):
                for p in periods:
                    if p.is_annual:
                        _derive_q4(values, p.fiscal_year)
            core_extracted[key] = (concept_name, values)

    # 2. Derive total_opex if missing (= revenue - operating income)
    if "total_opex" not in core_extracted and "total_revenue" in core_extracted \
            and "operating_income" in core_extracted:
        rev_vals = core_extracted["total_revenue"][1]
        op_vals = core_extracted["operating_income"][1]
        derived = {k: rev_vals[k] - op_vals[k] for k in rev_vals if k in op_vals}
        if derived:
            core_extracted["total_opex"] = ("_derived_total_opex", derived)

    # 3. Auto-discover ALL other IS concepts with decent quarterly coverage
    core_concept_names = {cn for cn, _ in core_extracted.values()}
    discovered: List[Tuple[str, Dict[str, float]]] = []
    for concept, data in us_gaap.items():
        if concept in core_concept_names:
            continue
        if not _looks_like_is_concept(concept):
            continue
        # Must be USD duration
        units = data.get("units", {})
        if "USD" not in units:
            continue
        # Only keep concepts with ≥ 2 recent quarter facts (active, not legacy)
        usd_facts = units["USD"]
        recent_q_facts = 0
        for f in usd_facts:
            if f.get("start") and f.get("end", "").startswith(("2024", "2025", "2026")):
                s, e = parse_date(f["start"]), parse_date(f["end"])
                if classify_duration(s, e) == "quarter":
                    recent_q_facts += 1
        if recent_q_facts < 2:
            continue
        values = _extract_values(us_gaap, concept, periods, ["USD"])
        for p in periods:
            if p.is_annual:
                _derive_q4(values, p.fiscal_year)
        if values and sum(1 for k in values if k in recent_labels) >= 2:
            discovered.append((concept, values))

    # 4. Dedupe discovered against identical value sets
    def _is_dup(vals, existing):
        for _, ex_vals in existing:
            shared = set(vals) & set(ex_vals)
            if shared and all(abs(vals[k] - ex_vals[k]) < 1 for k in shared):
                return True
        return False

    deduped: List[Tuple[str, Dict[str, float]]] = []
    for c, v in discovered:
        if not _is_dup(v, deduped):
            deduped.append((c, v))

    # 5. Build ISLine list in SECTION_ORDER
    all_lines: Dict[str, List[ISLine]] = {s: [] for s in SECTION_ORDER}

    # Place core rows
    CORE_TO_SECTION = {
        "total_revenue": "revenue_total",
        "total_opex": "opex_total",
        "operating_income": "operating_income",
        "pre_tax_income": "pre_tax",
        "tax_provision": "tax",
        "net_income": "net_income",
        "eps_basic": "eps",
        "eps_diluted": "eps",
        "shares_basic": "shares",
        "shares_diluted": "shares",
    }
    for key, (concept_name, values) in core_extracted.items():
        section = CORE_TO_SECTION[key]
        is_total = key in ("total_revenue", "total_opex", "net_income")
        all_lines[section].append(ISLine(
            concept=concept_name,
            section=section,
            label=_display_label(concept_name) if not concept_name.startswith("_") else "Total costs and expenses",
            values=values,
            is_total=is_total,
        ))

    # Place discovered rows
    for concept, values in deduped:
        section = _classify_concept(concept)
        # Skip rows that fall into a section we don't render or that would be a dupe of core
        if section not in SECTION_ORDER:
            continue
        # Skip section=revenue_total etc if already filled by core
        if section in ("revenue_total", "opex_total", "operating_income", "pre_tax",
                       "tax", "net_income", "gross_profit"):
            if all_lines.get(section):
                continue
        all_lines[section].append(ISLine(
            concept=concept,
            section=section,
            label=_display_label(concept),
            values=values,
            is_total=False,
        ))

    # Flatten in section order
    for sec in SECTION_ORDER:
        stmt.lines.extend(all_lines[sec])

    # 6. Extract balance sheet & cash flow (unchanged)
    instant_keys = {"cash_and_equivalents", "long_term_debt", "short_term_debt", "shares_outstanding"}
    for key, candidates in BS_CF_CONCEPTS.items():
        units_pref = ["shares"] if key == "shares_outstanding" else ["USD"]
        is_instant = key in instant_keys
        _, values = _extract_first_matching(us_gaap, candidates, periods, units_pref, instant=is_instant)
        if key in ("d_and_a", "capex", "operating_cash_flow"):
            for p in periods:
                if p.is_annual:
                    _derive_q4(values, p.fiscal_year)
            if values:
                stmt.cash_flow[key] = values
        else:
            if values:
                stmt.balance_sheet[key] = values

    return stmt


if __name__ == "__main__":
    from src.data.edgar import EdgarClient
    c = EdgarClient()
    cik = c.ticker_to_cik("COIN")
    facts = c.company_facts(cik)
    stmt = extract_income_statement("COIN", cik, c.company_title("COIN"), facts)
    print(f"Latest FY: {stmt.latest_fy}")
    print(f"Periods: {[p.label for p in stmt.periods]}\n")
    print(f"IS lines ({len(stmt.lines)}):")
    for line in stmt.lines:
        latest = stmt.periods[-1].label
        v = line.values.get("2025", 0)
        tot = "★" if line.is_total else " "
        print(f"  [{line.section:17s}]{tot} {line.label[:50]:50s}  2025={v/1e3:>12,.0f}")
