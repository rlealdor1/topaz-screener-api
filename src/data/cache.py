"""On-disk JSON cache keyed by URL. SEC rate-limits, so we aggressively cache."""
from __future__ import annotations
import hashlib
import json
from pathlib import Path
from typing import Any, Optional


class JsonCache:
    def __init__(self, root: Path):
        self.root = Path(root)
        self.root.mkdir(parents=True, exist_ok=True)

    def _key_path(self, key: str) -> Path:
        digest = hashlib.sha1(key.encode("utf-8")).hexdigest()[:16]
        return self.root / f"{digest}.json"

    def get(self, key: str) -> Optional[Any]:
        path = self._key_path(key)
        if not path.exists():
            return None
        try:
            return json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            return None

    def set(self, key: str, value: Any) -> None:
        path = self._key_path(key)
        path.write_text(json.dumps(value), encoding="utf-8")

    def clear(self) -> None:
        for p in self.root.glob("*.json"):
            p.unlink()
