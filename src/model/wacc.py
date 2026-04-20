"""CAPM-derived WACC (weighted average cost of capital).

WACC = (E/V) × Ke + (D/V) × Kd × (1 − t)
  Ke = rf + β × ERP
  Kd = rf + credit_spread   (credit spread bucketed by net_debt / EBITDA)

This is the standard sell-side-research methodology. Market inputs (rf, ERP,
spreads) live in config.yaml and should be refreshed quarterly.
"""
from __future__ import annotations
from dataclasses import dataclass
from typing import Optional


@dataclass
class WACCInputs:
    risk_free_rate: float
    equity_risk_premium: float
    beta: float
    tax_rate: float
    market_cap: float               # E (market value of equity)
    total_debt: float               # D (book value of debt — market value ≈ book for IG)
    cash: float                     # used for leverage calc
    ebitda_ttm: Optional[float]     # used for leverage calc
    credit_spread_net_cash: float = 0.010
    credit_spread_ig_low: float = 0.015
    credit_spread_ig_mid: float = 0.025
    credit_spread_crossover: float = 0.040
    credit_spread_high_yield: float = 0.060
    credit_spread_distressed: float = 0.090


@dataclass
class WACCResult:
    inputs: WACCInputs
    leverage_ratio: Optional[float]          # net debt / EBITDA
    credit_spread: float
    cost_of_equity: float                    # Ke
    pretax_cost_of_debt: float               # Kd
    aftertax_cost_of_debt: float             # Kd × (1 − t)
    equity_weight: float                     # E / (E+D)
    debt_weight: float                       # D / (E+D)
    wacc: float

    def __str__(self) -> str:
        return (f"WACC = {self.wacc:.2%}\n"
                f"  Ke = {self.cost_of_equity:.2%} (rf {self.inputs.risk_free_rate:.2%} + β {self.inputs.beta:.2f} × ERP {self.inputs.equity_risk_premium:.2%})\n"
                f"  Kd(after-tax) = {self.aftertax_cost_of_debt:.2%} (pretax {self.pretax_cost_of_debt:.2%}, spread {self.credit_spread:.2%})\n"
                f"  E/V = {self.equity_weight:.1%}, D/V = {self.debt_weight:.1%}\n"
                f"  Leverage = {self.leverage_ratio:.2f}x" if self.leverage_ratio is not None
                else f"  Leverage = n/a (no EBITDA)")


def _bucket_credit_spread(inputs: WACCInputs) -> tuple[Optional[float], float]:
    """Return (leverage_ratio, credit_spread) based on net_debt / EBITDA bucket."""
    net_debt = inputs.total_debt - inputs.cash
    if inputs.ebitda_ttm is None or inputs.ebitda_ttm <= 0:
        # Distressed / unprofitable — penalize heavily
        return None, inputs.credit_spread_distressed
    leverage = net_debt / inputs.ebitda_ttm
    if leverage < 0:
        return leverage, inputs.credit_spread_net_cash
    if leverage < 1.0:
        return leverage, inputs.credit_spread_ig_low
    if leverage < 3.0:
        return leverage, inputs.credit_spread_ig_mid
    if leverage < 4.5:
        return leverage, inputs.credit_spread_crossover
    return leverage, inputs.credit_spread_high_yield


def compute_wacc(inputs: WACCInputs) -> WACCResult:
    leverage, spread = _bucket_credit_spread(inputs)

    ke = inputs.risk_free_rate + inputs.beta * inputs.equity_risk_premium
    kd_pretax = inputs.risk_free_rate + spread
    kd_aftertax = kd_pretax * (1 - inputs.tax_rate)

    # Market-value weights
    total_capital = inputs.market_cap + inputs.total_debt
    if total_capital <= 0:
        we, wd = 1.0, 0.0
    else:
        we = inputs.market_cap / total_capital
        wd = inputs.total_debt / total_capital

    wacc = we * ke + wd * kd_aftertax

    # Floor WACC at rf + 1% (avoids silly results for super-low-beta cases)
    wacc = max(wacc, inputs.risk_free_rate + 0.01)

    return WACCResult(
        inputs=inputs,
        leverage_ratio=leverage,
        credit_spread=spread,
        cost_of_equity=ke,
        pretax_cost_of_debt=kd_pretax,
        aftertax_cost_of_debt=kd_aftertax,
        equity_weight=we,
        debt_weight=wd,
        wacc=wacc,
    )
