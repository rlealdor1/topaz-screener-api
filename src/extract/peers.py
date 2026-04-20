"""Peer selection: yaml override → sector-based auto-pick fallback."""
from __future__ import annotations
from pathlib import Path
from typing import List

import yaml


def load_peer_overrides(peers_yaml_path: Path) -> dict:
    if not peers_yaml_path.exists():
        return {}
    with open(peers_yaml_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def get_peers(ticker: str, peers_yaml_path: Path,
              cli_override: list[str] | None = None) -> List[str]:
    """Return peer ticker list. CLI override wins; else yaml; else empty (caller handles fallback)."""
    if cli_override:
        return [t.strip().upper() for t in cli_override if t.strip()]
    overrides = load_peer_overrides(peers_yaml_path)
    return [t.upper() for t in overrides.get(ticker.upper(), [])]
