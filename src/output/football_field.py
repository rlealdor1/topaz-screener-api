"""Football-field valuation chart: horizontal bars showing the price ranges
implied by DCF (multiple cases), trading comps, and market context."""
from __future__ import annotations
from pathlib import Path
from typing import List, Optional, Tuple

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle


def build_football_field(bands: List[Tuple[str, float, float]],
                          current_price: Optional[float],
                          ticker: str,
                          output_path: Path) -> Path:
    """Render a football-field chart.

    bands = list of (label, low, high) tuples, top-to-bottom in the plot.
    current_price draws a vertical marker line.
    """
    if not bands:
        raise ValueError("No valuation bands supplied.")

    bands = [(lbl, lo, hi) for lbl, lo, hi in bands if lo is not None and hi is not None]
    fig, ax = plt.subplots(figsize=(9, max(3, 0.6 * len(bands) + 1.5)), dpi=150)

    colors = ["#1F3864", "#2E75B6", "#5B9BD5", "#8BBC4B", "#A9D08E",
              "#FFC000", "#ED7D31", "#C00000"]
    y_positions = list(range(len(bands)))
    for i, (label, lo, hi) in enumerate(bands):
        if lo > hi:
            lo, hi = hi, lo
        width = hi - lo
        y = len(bands) - 1 - i
        color = colors[i % len(colors)]
        ax.add_patch(Rectangle((lo, y - 0.3), width, 0.6, color=color, alpha=0.75,
                                edgecolor="black", linewidth=0.5))
        # Midpoint label
        mid = (lo + hi) / 2
        ax.text(mid, y, f"${lo:.0f} – ${hi:.0f}",
                ha="center", va="center", fontsize=9, color="white", fontweight="bold")

    # Current price vertical line
    if current_price:
        ax.axvline(current_price, color="black", linestyle="--", linewidth=1.2, label="Current")
        ax.text(current_price, len(bands) - 0.3, f"  ${current_price:.2f}",
                fontsize=9, color="black", fontweight="bold", va="bottom")

    ax.set_yticks(y_positions)
    ax.set_yticklabels([lbl for lbl, _, _ in reversed(bands)], fontsize=10)
    ax.set_xlabel("Implied price per share ($)", fontsize=10)
    ax.set_title(f"{ticker} — Football Field Valuation", fontsize=12, fontweight="bold", loc="left")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(axis="x", linestyle="--", alpha=0.3)
    ax.set_ylim(-0.6, len(bands) - 0.4)

    # Nice x-axis range
    all_x = []
    for _, lo, hi in bands:
        all_x.extend([lo, hi])
    if current_price:
        all_x.append(current_price)
    if all_x:
        pad = (max(all_x) - min(all_x)) * 0.1 or 5
        ax.set_xlim(max(0, min(all_x) - pad), max(all_x) + pad)

    plt.tight_layout()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    plt.savefig(output_path, format="png", bbox_inches="tight")
    plt.close(fig)
    return output_path


def build_bands_from_dcf(dcf_result) -> List[Tuple[str, float, float]]:
    """Derive football-field bands from a DCFResult."""
    bands: List[Tuple[str, float, float]] = []

    scenarios = dcf_result.scenarios
    if "Bull" in scenarios and "Bear" in scenarios:
        bear = scenarios["Bear"].price_per_share
        base = scenarios["Base"].price_per_share
        bull = scenarios["Bull"].price_per_share
        bands.append(("DCF — Bear to Bull (blended TV)", min(bear, base, bull), max(bear, base, bull)))

    base_scen = scenarios.get("Base")
    if base_scen:
        g_vals = list(dcf_result.sensitivity_wacc_g.values())
        if g_vals:
            bands.append(("DCF — Gordon Growth sensitivity", min(g_vals), max(g_vals)))
        e_vals = list(dcf_result.sensitivity_wacc_exit.values())
        if e_vals:
            bands.append(("DCF — Exit Multiple sensitivity", min(e_vals), max(e_vals)))

    shares = dcf_result.shares_outstanding or 0.0
    if dcf_result.peer_low_pe and dcf_result.peer_high_pe and base_scen:
        # Use latest FY EPS to imply price: P = P/E × EPS
        # We approximate EPS as NOPAT / shares from base year
        eps_current = (dcf_result.base_revenue * dcf_result.current_ebit_margin *
                       (1 - dcf_result.assumptions.tax_rate)) / shares if shares else 0
        if eps_current > 0:
            bands.append(("Trading Comps — Peer P/E range",
                          dcf_result.peer_low_pe * eps_current,
                          dcf_result.peer_high_pe * eps_current))

    if dcf_result.peer_low_ev_ebitda and dcf_result.peer_high_ev_ebitda and shares:
        # EBITDA ≈ EBIT + D&A from base FY
        ebit = dcf_result.base_revenue * dcf_result.current_ebit_margin
        da = dcf_result.base_revenue * dcf_result.da_pct_revenue
        ebitda = ebit + da
        if ebitda > 0:
            net_cash = dcf_result.cash - dcf_result.debt
            low_price = (dcf_result.peer_low_ev_ebitda * ebitda + net_cash) / shares
            high_price = (dcf_result.peer_high_ev_ebitda * ebitda + net_cash) / shares
            bands.append(("Trading Comps — Peer EV/EBITDA range", low_price, high_price))

    # Sell-side analyst price target range
    if dcf_result.analyst_target_low and dcf_result.analyst_target_high:
        label = "Analyst price targets"
        if dcf_result.analyst_count:
            label += f" ({dcf_result.analyst_count} analysts)"
        bands.append((label, dcf_result.analyst_target_low, dcf_result.analyst_target_high))

    if dcf_result.week52_low and dcf_result.week52_high:
        bands.append(("52-Week trading range", dcf_result.week52_low, dcf_result.week52_high))

    return bands
