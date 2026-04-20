"""Thin yfinance wrapper for price, market-cap, EV, and peer fundamentals.

Used for the Comps table and the one-pager's Financial Highlights block.
"""
from __future__ import annotations
from dataclasses import dataclass
from typing import Optional

import yfinance as yf


@dataclass
class QuoteSnapshot:
    ticker: str
    company_name: str
    current_price: Optional[float]
    week52_high: Optional[float]
    market_cap: Optional[float]
    enterprise_value: Optional[float]
    revenue_ttm: Optional[float]
    gross_profit_ttm: Optional[float]
    ebitda_ttm: Optional[float]
    cash: Optional[float]
    total_debt: Optional[float]
    free_cash_flow: Optional[float]
    capex: Optional[float]
    operating_cash_flow: Optional[float]
    net_leverage: Optional[float]   # (debt - cash) / EBITDA
    pe_ratio: Optional[float]
    forward_pe: Optional[float]
    ev_ebitda: Optional[float]
    ev_sales: Optional[float]
    fcf_yield: Optional[float]      # FCF / market_cap
    hq_country: Optional[str]
    sector: Optional[str]
    industry: Optional[str]
    beta: Optional[float] = None
    week52_low: Optional[float] = None
    # Analyst consensus (sell-side)
    analyst_rev_current_year: Optional[float] = None   # 0y avg
    analyst_rev_next_year: Optional[float] = None      # +1y avg
    analyst_rev_growth_current_year: Optional[float] = None  # 0y YoY growth
    analyst_rev_growth_next_year: Optional[float] = None     # +1y YoY growth
    analyst_target_mean: Optional[float] = None
    analyst_target_high: Optional[float] = None
    analyst_target_low: Optional[float] = None
    analyst_target_median: Optional[float] = None
    analyst_rec_mean: Optional[float] = None           # 1=Strong Buy, 5=Sell
    analyst_rec_key: Optional[str] = None              # "strong_buy" / "buy" / "hold" / ...
    analyst_count: Optional[int] = None
    # EPS estimates — used to derive implied margin trajectory
    analyst_eps_current_year: Optional[float] = None
    analyst_eps_next_year: Optional[float] = None

    @property
    def pct_off_52w_high(self) -> Optional[float]:
        if self.current_price and self.week52_high and self.week52_high > 0:
            return (self.current_price - self.week52_high) / self.week52_high
        return None


def _safe(d: dict, *keys, default=None):
    for k in keys:
        v = d.get(k)
        if v is not None:
            return v
    return default


def get_quote(ticker: str) -> QuoteSnapshot:
    t = yf.Ticker(ticker)
    info = {}
    try:
        info = t.info or {}
    except Exception:
        info = {}

    market_cap = info.get("marketCap")
    ev = info.get("enterpriseValue")
    ebitda = info.get("ebitda")
    total_debt = info.get("totalDebt")
    cash = info.get("totalCash") or info.get("cash")
    fcf = info.get("freeCashflow")
    ocf = info.get("operatingCashflow")
    revenue_ttm = info.get("totalRevenue")
    gross = info.get("grossProfits")
    current_price = info.get("currentPrice") or info.get("regularMarketPrice")
    week52_high = info.get("fiftyTwoWeekHigh")
    week52_low = info.get("fiftyTwoWeekLow")
    pe = info.get("trailingPE")
    fwd_pe = info.get("forwardPE")
    sector = info.get("sector")
    industry = info.get("industry")
    country = info.get("country")
    name = info.get("longName") or info.get("shortName") or ticker
    beta = info.get("beta") or info.get("beta3Year")

    # Derive
    ev_ebitda = (ev / ebitda) if (ev and ebitda and ebitda > 0) else None
    ev_sales = (ev / revenue_ttm) if (ev and revenue_ttm and revenue_ttm > 0) else None
    fcf_yield = (fcf / market_cap) if (fcf and market_cap and market_cap > 0) else None
    net_leverage = ((total_debt - (cash or 0)) / ebitda) if (total_debt and ebitda and ebitda > 0) else None

    # CapEx — yfinance doesn't expose it directly in .info; derive from cashflow statement
    capex = None
    try:
        cf = t.cashflow
        if cf is not None and not cf.empty:
            for row_name in ("Capital Expenditure", "Capital Expenditures"):
                if row_name in cf.index:
                    capex = float(cf.loc[row_name].iloc[0])
                    break
    except Exception:
        pass

    # Analyst consensus — sell-side revenue estimates and price targets
    analyst_rev_cy = analyst_rev_ny = None
    analyst_growth_cy = analyst_growth_ny = None
    try:
        re = t.revenue_estimate
        if re is not None and not re.empty:
            if "0y" in re.index:
                analyst_rev_cy = float(re.loc["0y", "avg"])
                analyst_growth_cy = float(re.loc["0y", "growth"])
            if "+1y" in re.index:
                analyst_rev_ny = float(re.loc["+1y", "avg"])
                analyst_growth_ny = float(re.loc["+1y", "growth"])
    except Exception:
        pass

    # EPS estimates
    analyst_eps_cy = analyst_eps_ny = None
    try:
        ee = t.earnings_estimate
        if ee is not None and not ee.empty:
            if "0y" in ee.index:
                analyst_eps_cy = float(ee.loc["0y", "avg"])
            if "+1y" in ee.index:
                analyst_eps_ny = float(ee.loc["+1y", "avg"])
    except Exception:
        pass

    analyst_target_mean = info.get("targetMeanPrice")
    analyst_target_high = info.get("targetHighPrice")
    analyst_target_low = info.get("targetLowPrice")
    analyst_target_median = info.get("targetMedianPrice")
    analyst_rec_mean = info.get("recommendationMean")
    analyst_rec_key = info.get("recommendationKey")
    analyst_count = info.get("numberOfAnalystOpinions")

    return QuoteSnapshot(
        ticker=ticker.upper(),
        company_name=name,
        current_price=current_price,
        week52_high=week52_high,
        week52_low=week52_low,
        market_cap=market_cap,
        enterprise_value=ev,
        revenue_ttm=revenue_ttm,
        gross_profit_ttm=gross,
        ebitda_ttm=ebitda,
        cash=cash,
        total_debt=total_debt,
        free_cash_flow=fcf,
        capex=capex,
        operating_cash_flow=ocf,
        net_leverage=net_leverage,
        pe_ratio=pe,
        forward_pe=fwd_pe,
        ev_ebitda=ev_ebitda,
        ev_sales=ev_sales,
        fcf_yield=fcf_yield,
        hq_country=country,
        sector=sector,
        industry=industry,
        beta=beta,
        analyst_rev_current_year=analyst_rev_cy,
        analyst_rev_next_year=analyst_rev_ny,
        analyst_rev_growth_current_year=analyst_growth_cy,
        analyst_rev_growth_next_year=analyst_growth_ny,
        analyst_target_mean=analyst_target_mean,
        analyst_target_high=analyst_target_high,
        analyst_target_low=analyst_target_low,
        analyst_target_median=analyst_target_median,
        analyst_rec_mean=analyst_rec_mean,
        analyst_rec_key=analyst_rec_key,
        analyst_count=analyst_count,
        analyst_eps_current_year=analyst_eps_cy,
        analyst_eps_next_year=analyst_eps_ny,
    )


if __name__ == "__main__":
    for tkr in ("COIN", "HOOD", "CME"):
        q = get_quote(tkr)
        print(f"{q.ticker:6s} {q.company_name[:30]:30s} "
              f"price=${q.current_price}  mcap={q.market_cap/1e9 if q.market_cap else 0:.1f}B  "
              f"EV/EBITDA={q.ev_ebitda}")
