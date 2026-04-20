"""Peer selection with sector-based fallback.

Resolution order:
1. CLI --peers override (highest priority, explicit user choice)
2. peers.yaml entry for the ticker (hand-curated)
3. Sector default peers (derived from yfinance `sector` field on the quote)
4. Generic broad-market peers (last-resort S&P 500 bellwethers)

This means every ticker always gets a peer set, so the Comps table always
produces and the DCF always has peer EV/EBITDA data for its Exit Multiple
terminal-value branch.
"""
from __future__ import annotations
from pathlib import Path
from typing import List, Optional

import yaml


# Default peer sets per yfinance sector. Each list is 7 mega/large-cap names
# that broadly represent the sector's valuation multiples. The generator
# automatically filters out the target ticker from its own peer list.
SECTOR_DEFAULT_PEERS: dict[str, list[str]] = {
    "Technology":             ["MSFT", "AAPL", "GOOGL", "META", "NVDA", "ORCL", "CRM"],
    "Communication Services": ["GOOGL", "META", "NFLX", "DIS", "T", "VZ", "CMCSA"],
    "Healthcare":             ["JNJ", "UNH", "LLY", "PFE", "ABBV", "MRK", "TMO"],
    "Financial Services":     ["JPM", "BAC", "WFC", "GS", "MS", "BX", "SCHW"],
    "Consumer Cyclical":      ["AMZN", "TSLA", "HD", "MCD", "NKE", "SBUX", "LOW"],
    "Consumer Defensive":     ["WMT", "COST", "PG", "KO", "PEP", "MO", "CL"],
    "Industrials":            ["HON", "UNP", "CAT", "DE", "GE", "LMT", "RTX"],
    "Energy":                 ["XOM", "CVX", "COP", "EOG", "SLB", "PSX", "MPC"],
    "Basic Materials":        ["LIN", "SHW", "FCX", "NEM", "APD", "ECL", "CTVA"],
    "Real Estate":            ["PLD", "AMT", "EQIX", "SPG", "CCI", "O",    "WELL"],
    "Utilities":              ["NEE", "DUK", "SO",  "D",   "AEP", "SRE",  "EXC"],
}

# Last-resort bellwethers if sector is unknown or empty.
GENERIC_PEERS: list[str] = ["AAPL", "MSFT", "JPM", "JNJ", "XOM", "WMT", "HD"]


def load_peer_overrides(peers_yaml_path: Path) -> dict:
    if not peers_yaml_path.exists():
        return {}
    with open(peers_yaml_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def get_peers(
    ticker: str,
    peers_yaml_path: Path,
    cli_override: list[str] | None = None,
    sector: Optional[str] = None,
) -> List[str]:
    """Return a peer ticker list, guaranteed non-empty.

    Resolution order:
      1. `cli_override`      — explicit override passed via --peers on the CLI
      2. peers.yaml          — hand-curated per-ticker list (preferred when present)
      3. SECTOR_DEFAULT_PEERS — if yfinance reported a sector for this ticker
      4. GENERIC_PEERS       — last-resort bellwethers

    The target ticker is always filtered out of the returned list.
    """
    target = ticker.upper()

    # 1. CLI override
    if cli_override:
        resolved = [t.strip().upper() for t in cli_override if t.strip()]
    else:
        # 2. peers.yaml
        overrides = load_peer_overrides(peers_yaml_path)
        curated = overrides.get(target, [])
        if curated:
            resolved = [t.upper() for t in curated]
        elif sector and sector in SECTOR_DEFAULT_PEERS:
            # 3. Sector defaults
            resolved = SECTOR_DEFAULT_PEERS[sector][:]
        else:
            # 4. Last resort
            resolved = GENERIC_PEERS[:]

    # Always filter out the target itself (in case it ended up in its own sector list)
    return [t for t in resolved if t != target]
