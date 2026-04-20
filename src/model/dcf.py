"""Bank-style DCF with Bull/Base/Bear scenarios, explicit + fade forecast,
dual terminal value (Gordon Growth + Exit Multiple), SBC as cash expense,
working-capital modeling, and full sensitivity tables.

Output is consumed by `output/excel_financials.py` (Valuation sheet),
`output/word_report.py` (report context) and `output/pptx_onepager.py`.
"""
from __future__ import annotations
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from ..extract.income_statement import IncomeStatement
from .wacc import WACCInputs, WACCResult, compute_wacc


# ------------------------------------------------------------------
# Assumptions
# ------------------------------------------------------------------

@dataclass
class DCFAssumptions:
    # Horizon
    explicit_years: int = 5
    fade_years: int = 5
    terminal_growth: float = 0.030
    terminal_growth_megacap: float = 0.035
    megacap_threshold: float = 100_000_000_000.0   # $100B market cap
    tax_rate: float = 0.21

    # Revenue growth
    revenue_growth_cap: float = 0.30
    revenue_growth_floor: float = -0.05
    terminal_revenue_multiple_cap: float = 20.0   # Y10 revenue ≤ base × 20

    # Margin glide (premiums ADDED to current-FY EBIT margin)
    ebit_margin_premium_bull: float = 0.03
    ebit_margin_premium_base: float = 0.01
    ebit_margin_premium_bear: float = -0.02

    # NWC: ΔNWC = ratio × ΔRevenue  (ratio derived from history, fallback shown)
    nwc_pct_delta_revenue: float = 0.05

    # Stock-based comp treatment
    sbc_as_cash_expense: bool = True

    # CapEx and SBC ratio fades
    capex_fade_years: int = 5
    terminal_capex_pct_revenue: float = 0.08
    sbc_fade_years: int = 5
    terminal_sbc_pct_revenue: float = 0.03

    # Buyback modeling — banks factor this in, we do too
    model_buybacks: bool = True
    buyback_yield_cap: float = 0.06

    # Scenario WACC spreads
    wacc_adjustment_bull: float = -0.010
    wacc_adjustment_bear: float = 0.010

    # Terminal value blend — weighted toward Exit Multiple for quality capture
    tv_weight_gordon: float = 0.4
    tv_weight_exit: float = 0.6


@dataclass
class DCFScenarioSpec:
    name: str
    revenue_growth_starting_adj: float  # added to trailing CAGR starting growth
    ebit_margin_premium: float
    wacc_adjustment: float


# ------------------------------------------------------------------
# Projections
# ------------------------------------------------------------------

@dataclass
class DCFProjectionRow:
    year: int                # absolute fiscal year (e.g. 2026)
    phase: str               # "explicit" | "fade" | "terminal"
    year_count: float        # 1.0, 2.0, ... (for discounting)
    revenue_growth: float
    revenue: float
    ebit_margin: float
    ebit: float
    tax: float
    nopat: float
    d_and_a: float
    capex: float
    nwc_change: float
    sbc: float
    fcff: float              # NOPAT + D&A - CapEx - ΔNWC - SBC (if treated as cash)
    discount_factor: float
    pv_fcff: float
    ebitda: float            # for exit multiple


@dataclass
class DCFScenarioResult:
    scenario: DCFScenarioSpec
    wacc: float
    projections: List[DCFProjectionRow]
    sum_pv_fcff: float
    # Terminal value — two methods
    terminal_value_gordon: float
    pv_terminal_gordon: float
    terminal_value_exit: float
    pv_terminal_exit: float
    pv_terminal_blended: float
    # Equity bridge
    enterprise_value: float      # sum_pv_fcff + pv_terminal_blended
    cash: float
    debt: float
    equity_value: float
    shares_outstanding: float
    price_per_share: float
    # Gordon-only and Exit-only price-per-share (for football field)
    price_per_share_gordon_only: float
    price_per_share_exit_only: float


@dataclass
class DCFResult:
    """Top-level bank-style DCF output."""
    ticker: str
    base_fy: int
    base_revenue: float
    shares_outstanding: float
    cash: float
    debt: float
    assumptions: DCFAssumptions
    wacc_base: WACCResult
    # Trailing operating ratios used for forecasts (all % of revenue)
    da_pct_revenue: float
    capex_pct_revenue: float
    sbc_pct_revenue: float
    nwc_pct_delta_revenue: float
    trailing_revenue_cagr: float
    current_ebit_margin: float
    # Peer metrics for exit multiple
    peer_median_ev_ebitda: Optional[float]
    peer_median_pe: Optional[float]
    peer_low_ev_ebitda: Optional[float]
    peer_high_ev_ebitda: Optional[float]
    peer_low_pe: Optional[float]
    peer_high_pe: Optional[float]
    # Annual buyback yield (dollars repurchased / market cap) — drives
    # per-share accretion. 0 for companies that don't buy back shares.
    buyback_yield: float = 0.0
    # The effective terminal growth we used (may differ from assumptions if megacap)
    effective_terminal_growth: float = 0.03
    # Scenarios
    scenarios: Dict[str, DCFScenarioResult] = field(default_factory=dict)
    # Sensitivity grids (base case)
    sensitivity_wacc_g: Dict[Tuple[float, float], float] = field(default_factory=dict)
    sensitivity_wacc_exit: Dict[Tuple[float, float], float] = field(default_factory=dict)
    # Market context for football field
    current_price: Optional[float] = None
    market_cap: Optional[float] = None
    week52_low: Optional[float] = None
    week52_high: Optional[float] = None
    # Analyst consensus — drives forward revenue projection when available
    analyst_growth_y1: Optional[float] = None   # YoY growth into the first forecast year
    analyst_growth_y2: Optional[float] = None   # YoY growth into the second forecast year
    analyst_rev_y1: Optional[float] = None      # absolute consensus revenue Y1
    analyst_rev_y2: Optional[float] = None      # absolute consensus revenue Y2
    analyst_target_mean: Optional[float] = None
    analyst_target_high: Optional[float] = None
    analyst_target_low: Optional[float] = None
    analyst_rec_key: Optional[str] = None
    analyst_count: Optional[int] = None
    analyst_eps_y1: Optional[float] = None
    analyst_eps_y2: Optional[float] = None
    # Derived target margins from consensus EPS × shares / revenue.
    # Y1 seeds the first forecast year, Y2 the second.
    # `target_ebit_margin` aliases Y2 for backwards compatibility.
    target_ebit_margin_y1: Optional[float] = None
    target_ebit_margin_y2: Optional[float] = None
    target_ebit_margin: Optional[float] = None

    @property
    def base(self) -> DCFScenarioResult:
        return self.scenarios["Base"]


# ------------------------------------------------------------------
# Trailing ratios / helpers
# ------------------------------------------------------------------

def _section_value(stmt: IncomeStatement, section: str, label: str) -> Optional[float]:
    for line in stmt.lines:
        if line.section == section:
            v = line.values.get(label)
            if v is not None:
                return v
    return None


def _latest_annual_section(stmt: IncomeStatement, section: str) -> Optional[float]:
    for p in sorted(stmt.periods, key=lambda x: x.fiscal_year, reverse=True):
        if p.is_annual:
            v = _section_value(stmt, section, p.label)
            if v is not None:
                return v
    return None


def _latest_annual_bucket(stmt: IncomeStatement, bucket_name: str, key: str) -> Optional[float]:
    bucket = getattr(stmt, bucket_name)
    for p in sorted(stmt.periods, key=lambda x: x.fiscal_year, reverse=True):
        if p.is_annual:
            v = bucket.get(key, {}).get(p.label)
            if v is not None:
                return v
    return None


def _trailing_cagr(stmt: IncomeStatement, years: int = 3) -> float:
    annuals = {}
    for p in stmt.periods:
        if p.is_annual:
            v = _section_value(stmt, "revenue_total", p.label)
            if v is not None and v > 0:
                annuals[p.fiscal_year] = v
    if len(annuals) < 2:
        return 0.10
    yrs = sorted(annuals.keys())
    start_year = yrs[max(0, len(yrs) - years - 1)]
    end_year = yrs[-1]
    n = end_year - start_year
    if n <= 0:
        return 0.10
    return (annuals[end_year] / annuals[start_year]) ** (1 / n) - 1


def _current_ebit_margin(stmt: IncomeStatement) -> float:
    for p in sorted(stmt.periods, key=lambda x: x.fiscal_year, reverse=True):
        if p.is_annual:
            rev = _section_value(stmt, "revenue_total", p.label)
            op = _section_value(stmt, "operating_income", p.label)
            if rev and op is not None and rev > 0:
                return op / rev
    return 0.15


def _trailing_ratio_cf(stmt: IncomeStatement, key: str) -> float:
    for p in sorted(stmt.periods, key=lambda x: x.fiscal_year, reverse=True):
        if p.is_annual:
            rev = _section_value(stmt, "revenue_total", p.label)
            v = stmt.cash_flow.get(key, {}).get(p.label)
            if rev and v is not None and rev > 0:
                return abs(v) / rev
    return 0.0


def _latest_ebitda(stmt: IncomeStatement) -> Optional[float]:
    op = _latest_annual_section(stmt, "operating_income")
    da = _latest_annual_bucket(stmt, "cash_flow", "d_and_a")
    if op is not None and da is not None:
        return op + da
    return None


def _extract_sbc_ratio(stmt: IncomeStatement, company_facts: dict) -> float:
    """SBC as % of revenue (latest FY). Returns 0.0 if not tagged."""
    us = company_facts.get("facts", {}).get("us-gaap", {})
    sbc_concepts = [
        "ShareBasedCompensation",
        "AllocatedShareBasedCompensationExpense",
        "StockBasedCompensation",
    ]
    for c in sbc_concepts:
        if c not in us:
            continue
        facts = us[c].get("units", {}).get("USD", [])
        fy_facts = [f for f in facts if f.get("fp") == "FY"]
        if not fy_facts:
            continue
        latest = max(fy_facts, key=lambda f: f.get("fy", 0))
        rev = _latest_annual_section(stmt, "revenue_total") or 0
        if rev > 0:
            return abs(latest["val"]) / rev
    return 0.0


def _extract_nwc_ratio(stmt: IncomeStatement, company_facts: dict) -> Optional[float]:
    """ΔNWC / ΔRevenue from trailing FY. Returns None if not derivable."""
    us = company_facts.get("facts", {}).get("us-gaap", {})
    nwc_concept = "IncreaseDecreaseInOperatingCapital"
    if nwc_concept not in us:
        return None
    facts = us[nwc_concept].get("units", {}).get("USD", [])
    fy_facts = sorted([f for f in facts if f.get("fp") == "FY"], key=lambda f: f.get("fy", 0))
    if not fy_facts:
        return None
    latest = fy_facts[-1]
    nwc_change = latest["val"]  # negative usually = NWC increase
    # Revenue Δ in same year
    latest_rev = _latest_annual_section(stmt, "revenue_total")
    prev_year = latest.get("fy", 0) - 1
    prev_rev = _section_value(stmt, "revenue_total", str(prev_year))
    if not latest_rev or not prev_rev:
        return None
    d_rev = latest_rev - prev_rev
    if abs(d_rev) < 1e6:
        return None
    # NWC increases (use of cash) are reported as negative; flip sign so ratio is positive.
    return max(-nwc_change / d_rev, 0.0)


# ------------------------------------------------------------------
# Core projection engine
# ------------------------------------------------------------------

def _extract_buyback_yield(company_facts: dict, market_cap: float, cap: float = 0.06) -> float:
    """Annualized share-buyback yield = last-FY buybacks / current market cap.

    Uses SEC concept `PaymentsForRepurchaseOfCommonStock` (value usually
    reported positive in the cash-flow-from-financing section; sign doesn't
    matter — we take absolute value). Returns 0.0 if the concept isn't
    tagged or there's no data. Capped at `cap` to avoid distortion from
    one-time special repurchases.
    """
    if not market_cap or market_cap <= 0:
        return 0.0
    us = company_facts.get("facts", {}).get("us-gaap", {})
    concepts = [
        "PaymentsForRepurchaseOfCommonStock",
        "PaymentsForRepurchaseOfEquity",
        "StockRepurchasedDuringPeriodValue",
    ]
    for c in concepts:
        if c not in us:
            continue
        units = us[c].get("units", {}).get("USD", [])
        fy_facts = [f for f in units if f.get("fp") == "FY"]
        if not fy_facts:
            continue
        latest = max(fy_facts, key=lambda f: f.get("fy", 0))
        annual_buyback = abs(latest.get("val", 0))
        y = annual_buyback / market_cap
        return min(y, cap)
    return 0.0


def _fade_ratio(trailing: float, terminal: float, year_count: int,
                fade_start: int = 2, fade_years: int = 5) -> float:
    """Ratio with a linear fade from `trailing` to `terminal`.

    Years ≤ fade_start      → trailing (construction / transition phase)
    Years > fade_start+N    → terminal
    In between              → linear interpolation

    If trailing is already below terminal, keep trailing (no reason to fade up).
    """
    if trailing <= terminal:
        return trailing
    if year_count <= fade_start:
        return trailing
    if year_count >= fade_start + fade_years:
        return terminal
    frac = (year_count - fade_start) / fade_years
    return trailing + (terminal - trailing) * frac

def _project_one_scenario(
    result: DCFResult,
    spec: DCFScenarioSpec,
    assumptions: DCFAssumptions,
    base_revenue: float,
    base_fy: int,
    wacc: float,
    exit_multiple_ev_ebitda: float,
) -> DCFScenarioResult:
    """Project a single scenario and compute its equity value.

    Growth logic (sell-side bank style):
      • If analyst consensus growth is available (Y1 and Y2), seed the first two
        forecast years from consensus + scenario adjustment.
      • From Y3 onwards, decay linearly from the Y2 growth rate to terminal.
      • If no consensus is available, fall back to trailing 3Y CAGR as the seed.
      • Growth is capped/floored to avoid absurd projections.
    """
    terminal_growth = assumptions.terminal_growth
    total_years = assumptions.explicit_years + assumptions.fade_years

    # Margin trajectory — analyst-consensus driven (same pattern as revenue):
    #   Y1 margin = Y1 target from consensus EPS (if available)
    #              else midpoint between current margin and Y2 target
    #              else current margin + premium (legacy)
    #   Y2 margin = Y2 target from consensus EPS (if available) + scenario premium
    #   Y3+       = modest linear expansion from Y2 to (Y2 + terminal margin expansion)
    m_y1_consensus = result.target_ebit_margin_y1
    m_y2_consensus = result.target_ebit_margin_y2
    margin_end = (m_y2_consensus or result.current_ebit_margin) + spec.ebit_margin_premium
    # Long-run margin: hold Y2 margin with a small expansion (+300bps by terminal year).
    # Miners and commodity cos typically don't expand margins much post-scale.
    m_terminal = (m_y2_consensus + 0.03) if m_y2_consensus is not None else margin_end
    m_terminal += spec.ebit_margin_premium

    # Determine Y1 and Y2 seed growth rates.
    # When sell-side consensus is available, trust it — skip the config growth cap for Y1/Y2.
    # Config cap still applies if we're falling back to trailing CAGR.
    use_consensus = (result.analyst_growth_y1 is not None and
                     result.analyst_growth_y2 is not None)
    if use_consensus:
        # Apply only the floor, not the cap — sell-side can legitimately project >30% for ramping companies.
        g_y1 = max(result.analyst_growth_y1 + spec.revenue_growth_starting_adj,
                   assumptions.revenue_growth_floor)
        g_y2 = max(result.analyst_growth_y2 + spec.revenue_growth_starting_adj,
                   assumptions.revenue_growth_floor)
    else:
        seed = max(min(result.trailing_revenue_cagr + spec.revenue_growth_starting_adj,
                       assumptions.revenue_growth_cap),
                   assumptions.revenue_growth_floor)
        g_y1 = g_y2 = seed

    projections: List[DCFProjectionRow] = []
    revenue = base_revenue
    sum_pv = 0.0

    for i in range(total_years):
        year_count = i + 1
        fy = base_fy + year_count
        phase = "explicit" if year_count <= assumptions.explicit_years else "fade"

        # Growth schedule
        if year_count == 1:
            growth = g_y1
        elif year_count == 2:
            growth = g_y2
        else:
            # Linear decay from Y2 growth to terminal over remaining years
            frac = (year_count - 2) / max(1, total_years - 2)
            growth = g_y2 - (g_y2 - terminal_growth) * frac

        revenue_prev = revenue
        revenue = revenue_prev * (1 + growth)

        # Revenue growth guardrail: cap Y10 revenue at base × multiple_cap.
        # Prevents hyper-growth extrapolation (e.g. MP going 55× in 10 years).
        # Historic outlier growth: NFLX ~22×, NVDA ~15×, TSLA ~24× per decade.
        # Above the cap, scale back the projected revenue (effectively lowers
        # growth for that year and subsequent years proportionally).
        max_terminal_revenue = result.base_revenue * assumptions.terminal_revenue_multiple_cap
        if revenue > max_terminal_revenue:
            revenue = max_terminal_revenue
            growth = (revenue / revenue_prev) - 1 if revenue_prev > 0 else 0

        # Margin schedule — mirrors growth: Y1/Y2 from consensus if available, then expand.
        if year_count == 1:
            if m_y1_consensus is not None:
                margin = m_y1_consensus + spec.ebit_margin_premium
            elif m_y2_consensus is not None:
                # No Y1 EPS estimate, but we have Y2 → bridge halfway from current to Y2 target
                margin = (result.current_ebit_margin + m_y2_consensus) / 2 + spec.ebit_margin_premium
            else:
                margin = result.current_ebit_margin + spec.ebit_margin_premium * (1 / total_years)
        elif year_count == 2:
            if m_y2_consensus is not None:
                margin = m_y2_consensus + spec.ebit_margin_premium
            else:
                margin = result.current_ebit_margin + spec.ebit_margin_premium * (2 / total_years)
        else:
            # Years 3-10: linear expansion from Y2 margin to terminal (Y2 + 300bps)
            anchor = (m_y2_consensus or margin_end)
            frac = (year_count - 2) / max(1, total_years - 2)
            margin = anchor + (m_terminal - anchor) * frac

        ebit = revenue * margin
        tax = ebit * assumptions.tax_rate
        nopat = ebit - tax
        d_and_a = revenue * result.da_pct_revenue
        # CapEx and SBC fade from elevated trailing ratios to mature steady-state,
        # reflecting the reality that construction-phase capex and early-stage SBC
        # don't persist at their current % of revenue as the company scales.
        capex_ratio = _fade_ratio(result.capex_pct_revenue,
                                  assumptions.terminal_capex_pct_revenue,
                                  year_count, fade_start=2,
                                  fade_years=assumptions.capex_fade_years)
        sbc_ratio = _fade_ratio(result.sbc_pct_revenue,
                                assumptions.terminal_sbc_pct_revenue,
                                year_count, fade_start=2,
                                fade_years=assumptions.sbc_fade_years)
        capex = revenue * capex_ratio
        nwc_change = max(revenue - revenue_prev, 0.0) * result.nwc_pct_delta_revenue
        sbc = revenue * sbc_ratio if assumptions.sbc_as_cash_expense else 0.0

        fcff = nopat + d_and_a - capex - nwc_change - sbc
        discount = 1 / (1 + wacc) ** year_count
        pv = fcff * discount
        sum_pv += pv

        projections.append(DCFProjectionRow(
            year=fy, phase=phase, year_count=year_count,
            revenue_growth=growth, revenue=revenue,
            ebit_margin=margin, ebit=ebit, tax=tax, nopat=nopat,
            d_and_a=d_and_a, capex=capex, nwc_change=nwc_change, sbc=sbc,
            fcff=fcff, discount_factor=discount, pv_fcff=pv,
            ebitda=ebit + d_and_a,
        ))

    # Terminal value — two methods
    last = projections[-1]
    # Gordon Growth: normalize FCF to year-(n+1)
    fcff_terminal = last.fcff * (1 + terminal_growth)
    tv_gordon = fcff_terminal / (wacc - terminal_growth) if wacc > terminal_growth else 0.0
    pv_tv_gordon = tv_gordon * last.discount_factor

    # Exit Multiple: Year-n EBITDA × peer median EV/EBITDA
    tv_exit = last.ebitda * exit_multiple_ev_ebitda if exit_multiple_ev_ebitda else 0.0
    pv_tv_exit = tv_exit * last.discount_factor

    # Blend
    w_g, w_e = assumptions.tv_weight_gordon, assumptions.tv_weight_exit
    if tv_exit == 0:
        # fallback to 100% Gordon when no peer multiple
        w_g, w_e = 1.0, 0.0
    pv_tv_blended = w_g * pv_tv_gordon + w_e * pv_tv_exit

    ev = sum_pv + pv_tv_blended
    equity = ev + result.cash - result.debt

    # Buyback accretion: shares compound down by buyback_yield per year over
    # the forecast horizon. Use the TERMINAL year's share count (fully
    # compounded) because that's what matters for the TV portion, which
    # dominates most DCF valuations for mature companies.
    # Banks universally factor this in (Apple's ~2.25%/yr yield = ~25% cumulative
    # share reduction over 10 years = 33% per-share accretion on TV).
    buyback_shares = result.shares_outstanding
    if result.buyback_yield > 0:
        buyback_shares = result.shares_outstanding * (1 - result.buyback_yield) ** total_years
    pps = equity / buyback_shares if buyback_shares > 0 else 0.0

    # Gordon-only and Exit-only PPS for football field
    equity_gordon = sum_pv + pv_tv_gordon + result.cash - result.debt
    equity_exit = sum_pv + pv_tv_exit + result.cash - result.debt
    pps_gordon = equity_gordon / buyback_shares if buyback_shares > 0 else 0.0
    pps_exit = equity_exit / buyback_shares if buyback_shares > 0 else 0.0

    return DCFScenarioResult(
        scenario=spec, wacc=wacc, projections=projections,
        sum_pv_fcff=sum_pv,
        terminal_value_gordon=tv_gordon, pv_terminal_gordon=pv_tv_gordon,
        terminal_value_exit=tv_exit, pv_terminal_exit=pv_tv_exit,
        pv_terminal_blended=pv_tv_blended,
        enterprise_value=ev, cash=result.cash, debt=result.debt,
        equity_value=equity, shares_outstanding=result.shares_outstanding,
        price_per_share=pps,
        price_per_share_gordon_only=pps_gordon,
        price_per_share_exit_only=pps_exit,
    )


# ------------------------------------------------------------------
# Public entry
# ------------------------------------------------------------------

def run_dcf(stmt: IncomeStatement, assumptions: DCFAssumptions,
            wacc_config: dict, company_facts: dict,
            quote=None, peer_quotes: Optional[List] = None) -> DCFResult:
    """Run the full bank-style DCF with Bull/Base/Bear and sensitivity tables."""

    base_fy = stmt.latest_fy
    base_revenue = _latest_annual_section(stmt, "revenue_total") or 0.0
    if base_revenue <= 0:
        raise ValueError(f"No base-year revenue for {stmt.ticker}")

    # Trailing operating ratios
    da_pct = _trailing_ratio_cf(stmt, "d_and_a") or 0.02
    capex_pct = _trailing_ratio_cf(stmt, "capex") or 0.01
    sbc_pct = _extract_sbc_ratio(stmt, company_facts)
    nwc_pct = _extract_nwc_ratio(stmt, company_facts)
    if nwc_pct is None:
        nwc_pct = assumptions.nwc_pct_delta_revenue
    cagr = _trailing_cagr(stmt, 3)
    margin = _current_ebit_margin(stmt)

    # Balance sheet
    cash = _latest_annual_bucket(stmt, "balance_sheet", "cash_and_equivalents") or 0.0
    debt = (_latest_annual_bucket(stmt, "balance_sheet", "long_term_debt") or 0.0) + \
           (_latest_annual_bucket(stmt, "balance_sheet", "short_term_debt") or 0.0)

    # Shares (diluted, weighted-average, latest FY)
    shares = 0.0
    for line in stmt.lines:
        if line.section == "shares" and "Diluted" in line.label:
            for p in sorted(stmt.periods, key=lambda x: x.fiscal_year, reverse=True):
                if p.is_annual:
                    v = line.values.get(p.label)
                    if v is not None:
                        shares = v
                        break
            if shares:
                break

    # Market inputs for WACC — from yfinance quote if available
    beta = wacc_config.get("beta_fallback", 1.10)
    market_cap = 0.0
    ebitda_ttm = _latest_ebitda(stmt)
    if quote is not None:
        beta = getattr(quote, "beta", None) or beta
        market_cap = getattr(quote, "market_cap", 0.0) or 0.0
        if getattr(quote, "ebitda_ttm", None):
            ebitda_ttm = quote.ebitda_ttm

    # Beta sanity floor for small-caps: yfinance's 3Y regression collapses
    # toward zero for thinly-traded stocks. Floor small-cap betas at the
    # configured minimum so WACC doesn't get artificially low.
    small_cap_threshold = wacc_config.get("small_cap_threshold", 10.0e9)
    beta_min_smallcap = wacc_config.get("beta_min_smallcap", 1.00)
    if market_cap and market_cap < small_cap_threshold and beta < beta_min_smallcap:
        beta = beta_min_smallcap

    # Mega-cap quality premium: for companies above the megacap threshold
    # (default $100B), use the higher terminal growth assumption. This
    # reflects durable moats, services/recurring revenue, emerging market
    # runway, etc. — all the things that justify banks pricing AAPL/MSFT
    # with higher implied terminal growth than a no-name industrial.
    effective_terminal_growth = assumptions.terminal_growth
    if market_cap and market_cap >= assumptions.megacap_threshold:
        effective_terminal_growth = assumptions.terminal_growth_megacap

    # Share buyback yield — drives per-share accretion over the forecast
    # period. Banks always include this for companies like AAPL/MSFT that
    # return ~2-5% of market cap per year via buybacks.
    buyback_yield = 0.0
    if assumptions.model_buybacks and market_cap:
        buyback_yield = _extract_buyback_yield(
            company_facts, market_cap, cap=assumptions.buyback_yield_cap
        )

    wacc_inputs = WACCInputs(
        risk_free_rate=wacc_config["risk_free_rate"],
        equity_risk_premium=wacc_config["equity_risk_premium"],
        beta=beta,
        tax_rate=assumptions.tax_rate,
        market_cap=market_cap or (shares * (quote.current_price if quote and quote.current_price else 0)),
        total_debt=debt,
        cash=cash,
        ebitda_ttm=ebitda_ttm,
        credit_spread_net_cash=wacc_config.get("credit_spread_net_cash", 0.010),
        credit_spread_ig_low=wacc_config.get("credit_spread_ig_low", 0.015),
        credit_spread_ig_mid=wacc_config.get("credit_spread_ig_mid", 0.025),
        credit_spread_crossover=wacc_config.get("credit_spread_crossover", 0.040),
        credit_spread_high_yield=wacc_config.get("credit_spread_high_yield", 0.060),
        credit_spread_distressed=wacc_config.get("credit_spread_distressed", 0.090),
    )
    wacc_base = compute_wacc(wacc_inputs)

    # Peer multiples for exit multiple and football field
    peer_median_ev_ebitda = None
    peer_median_pe = None
    peer_low_ev_ebitda = peer_high_ev_ebitda = None
    peer_low_pe = peer_high_pe = None
    if peer_quotes:
        ev_ebitdas = [q.ev_ebitda for q in peer_quotes if q.ev_ebitda and q.ev_ebitda > 0]
        pes = [q.pe_ratio for q in peer_quotes if q.pe_ratio and q.pe_ratio > 0]
        if ev_ebitdas:
            s = sorted(ev_ebitdas)
            peer_median_ev_ebitda = s[len(s) // 2]
            peer_low_ev_ebitda = s[0]
            peer_high_ev_ebitda = s[-1]
        if pes:
            s = sorted(pes)
            peer_median_pe = s[len(s) // 2]
            peer_low_pe = s[0]
            peer_high_pe = s[-1]

    result = DCFResult(
        ticker=stmt.ticker, base_fy=base_fy, base_revenue=base_revenue,
        shares_outstanding=shares, cash=cash, debt=debt,
        assumptions=assumptions, wacc_base=wacc_base,
        da_pct_revenue=da_pct, capex_pct_revenue=capex_pct,
        sbc_pct_revenue=sbc_pct, nwc_pct_delta_revenue=nwc_pct,
        trailing_revenue_cagr=cagr, current_ebit_margin=margin,
        peer_median_ev_ebitda=peer_median_ev_ebitda,
        peer_median_pe=peer_median_pe,
        peer_low_ev_ebitda=peer_low_ev_ebitda,
        peer_high_ev_ebitda=peer_high_ev_ebitda,
        peer_low_pe=peer_low_pe, peer_high_pe=peer_high_pe,
        current_price=getattr(quote, "current_price", None) if quote else None,
        market_cap=market_cap if market_cap else None,
        week52_low=getattr(quote, "week52_low", None) if quote else None,
        week52_high=getattr(quote, "week52_high", None) if quote else None,
        analyst_growth_y1=getattr(quote, "analyst_rev_growth_current_year", None) if quote else None,
        analyst_growth_y2=getattr(quote, "analyst_rev_growth_next_year", None) if quote else None,
        analyst_rev_y1=getattr(quote, "analyst_rev_current_year", None) if quote else None,
        analyst_rev_y2=getattr(quote, "analyst_rev_next_year", None) if quote else None,
        analyst_target_mean=getattr(quote, "analyst_target_mean", None) if quote else None,
        analyst_target_high=getattr(quote, "analyst_target_high", None) if quote else None,
        analyst_target_low=getattr(quote, "analyst_target_low", None) if quote else None,
        analyst_rec_key=getattr(quote, "analyst_rec_key", None) if quote else None,
        analyst_count=getattr(quote, "analyst_count", None) if quote else None,
        analyst_eps_y1=getattr(quote, "analyst_eps_current_year", None) if quote else None,
        analyst_eps_y2=getattr(quote, "analyst_eps_next_year", None) if quote else None,
        buyback_yield=buyback_yield,
        effective_terminal_growth=effective_terminal_growth,
    )

    # Override the assumptions copy used for projections with the effective
    # terminal growth (which may be the megacap 3.5% value).
    assumptions.terminal_growth = effective_terminal_growth

    # Derive target EBIT margins for Y1 and Y2 from consensus EPS × shares / revenue.
    # Implied net margin; gross up by (1 - tax) to get EBIT margin proxy.
    if result.analyst_eps_y1 and result.analyst_rev_y1 and shares > 0:
        implied_net_margin_y1 = (result.analyst_eps_y1 * shares) / result.analyst_rev_y1
        result.target_ebit_margin_y1 = implied_net_margin_y1 / (1 - assumptions.tax_rate)
    if result.analyst_eps_y2 and result.analyst_rev_y2 and shares > 0:
        implied_net_margin_y2 = (result.analyst_eps_y2 * shares) / result.analyst_rev_y2
        result.target_ebit_margin_y2 = implied_net_margin_y2 / (1 - assumptions.tax_rate)
    # Legacy alias — used by Excel formula (G15)
    result.target_ebit_margin = result.target_ebit_margin_y2 or result.target_ebit_margin_y1

    # Build the three scenarios
    scenario_specs = [
        DCFScenarioSpec("Bear", revenue_growth_starting_adj=-0.03,
                        ebit_margin_premium=assumptions.ebit_margin_premium_bear,
                        wacc_adjustment=assumptions.wacc_adjustment_bear),
        DCFScenarioSpec("Base", revenue_growth_starting_adj=0.0,
                        ebit_margin_premium=assumptions.ebit_margin_premium_base,
                        wacc_adjustment=0.0),
        DCFScenarioSpec("Bull", revenue_growth_starting_adj=0.03,
                        ebit_margin_premium=assumptions.ebit_margin_premium_bull,
                        wacc_adjustment=assumptions.wacc_adjustment_bull),
    ]

    # Base-case exit multiple: use peer median when available, else 0 (which
    # causes the TV blend to fall back to 100% Gordon Growth). We don't use a
    # fake fallback here — that would inflate the valuation.
    exit_multiple = peer_median_ev_ebitda or 0.0

    for spec in scenario_specs:
        scenario_wacc = wacc_base.wacc + spec.wacc_adjustment
        scenario_wacc = max(scenario_wacc, wacc_inputs.risk_free_rate + 0.005)
        res = _project_one_scenario(result, spec, assumptions,
                                     base_revenue, base_fy,
                                     scenario_wacc, exit_multiple)
        result.scenarios[spec.name] = res

    # Sensitivity grids — base case only
    base_spec = next(s for s in scenario_specs if s.name == "Base")
    wacc_grid = [wacc_base.wacc + d for d in (-0.02, -0.01, -0.005, 0.0, 0.005, 0.01, 0.02)]
    g_grid = [assumptions.terminal_growth + d for d in (-0.010, -0.005, 0.0, 0.005, 0.010)]
    # Sensitivity 2 only renders if we have valid peer EV/EBITDA data.
    # If all peers are pre-profit (e.g. UAMY, early-stage miners), we leave
    # the grid empty and the Excel writer shows an explanatory note instead.
    exit_grid = [exit_multiple * m for m in (0.7, 0.85, 1.0, 1.15, 1.3)] if exit_multiple else []

    for w in wacc_grid:
        for g in g_grid:
            if w <= g:
                continue
            alt = DCFAssumptions(**{**assumptions.__dict__, "terminal_growth": g})
            alt.tv_weight_gordon = 1.0  # sensitivity 1 = pure Gordon Growth
            alt.tv_weight_exit = 0.0
            res = _project_one_scenario(result, base_spec, alt, base_revenue, base_fy,
                                         max(w, wacc_inputs.risk_free_rate + 0.005),
                                         exit_multiple)
            result.sensitivity_wacc_g[(round(w, 4), round(g, 4))] = res.price_per_share

    for w in wacc_grid:
        for em in exit_grid:
            alt = DCFAssumptions(**{**assumptions.__dict__})
            alt.tv_weight_gordon = 0.0  # sensitivity 2 = pure Exit Multiple
            alt.tv_weight_exit = 1.0
            res = _project_one_scenario(result, base_spec, alt, base_revenue, base_fy,
                                         max(w, wacc_inputs.risk_free_rate + 0.005), em)
            result.sensitivity_wacc_exit[(round(w, 4), round(em, 4))] = res.price_per_share

    return result


# Legacy alias for compatibility while callers migrate
DCFInputs = DCFAssumptions
