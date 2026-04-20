"""Dual-axis chart helper for the one-pager."""
from __future__ import annotations
from pathlib import Path
from typing import Dict, Optional

import matplotlib
matplotlib.use("Agg")  # no display needed
import matplotlib.pyplot as plt


def revenue_kpi_chart(revenue_by_period: Dict[str, float],
                      kpi_by_period: Optional[Dict[str, float]],
                      kpi_label: str,
                      output_path: Path,
                      color_rev: str = "#A9D08E",   # Topaz green-ish
                      color_kpi: str = "#1F3864") -> Path:
    """Create a dual-axis chart: revenue bars + KPI line. Returns path to PNG."""
    periods = [p for p in revenue_by_period.keys() if not p.isdigit() and "E" not in p]
    # Keep only quarterly labels (format like 1Q23); drop annual totals for chart clarity
    periods = [p for p in periods if "Q" in p]
    periods = sorted(periods, key=_period_sort_key)
    rev_values = [revenue_by_period[p] / 1000 for p in periods]  # thousands → millions

    fig, ax1 = plt.subplots(figsize=(7, 3.6), dpi=150)
    ax1.plot(periods, rev_values, marker="o", linewidth=2.5, color=color_rev, label="Revenue")
    ax1.fill_between(periods, rev_values, alpha=0.15, color=color_rev)
    ax1.set_ylabel("Revenue ($M)", fontsize=10, color=color_rev)
    ax1.tick_params(axis="y", labelcolor=color_rev, labelsize=9)
    ax1.tick_params(axis="x", labelsize=9, rotation=45)
    ax1.grid(axis="y", linestyle="--", alpha=0.3)
    ax1.spines["top"].set_visible(False)
    ax1.spines["right"].set_visible(False)

    if kpi_by_period:
        kpi_periods = [p for p in periods if p in kpi_by_period]
        kpi_values = [kpi_by_period[p] for p in kpi_periods]
        ax2 = ax1.twinx()
        ax2.plot(kpi_periods, kpi_values, marker="s", linewidth=2.5, color=color_kpi,
                 label=kpi_label)
        ax2.set_ylabel(kpi_label, fontsize=10, color=color_kpi)
        ax2.tick_params(axis="y", labelcolor=color_kpi, labelsize=9)
        ax2.spines["top"].set_visible(False)

    plt.title("Revenue Growth" + (f" vs {kpi_label}" if kpi_by_period else ""),
              fontsize=11, loc="left", fontweight="bold")

    plt.tight_layout()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    plt.savefig(output_path, format="png", bbox_inches="tight")
    plt.close(fig)
    return output_path


def _period_sort_key(label: str) -> tuple:
    """Sort 1Q23, 2Q23, ..., 1Q25 in chronological order."""
    if "Q" in label:
        q, y = label.split("Q")
        return (2000 + int(y), int(q))
    return (int(label), 5)
