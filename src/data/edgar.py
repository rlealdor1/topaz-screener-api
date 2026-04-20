"""SEC EDGAR client — ticker lookup, company facts (XBRL), filings submissions.

Uses the free, official SEC JSON endpoints. Requires User-Agent header.
Rate limits to 10 req/s (SEC hard cap).

`company_tickers.json` is bundled locally at src/data/company_tickers.json
so we never have to hit SEC's `/files/company_tickers.json` endpoint
(which is aggressively rate-limited on shared-IP cloud hosts like Render).
Refresh locally by running `curl -H "User-Agent: $SEC_USER_AGENT"
https://www.sec.gov/files/company_tickers.json -o src/data/company_tickers.json`
whenever a newly-IPO'd ticker needs to be screenable.
"""
from __future__ import annotations
import json
import os
import time
from pathlib import Path
from typing import Optional

import requests
from dotenv import load_dotenv

from .cache import JsonCache

load_dotenv(override=True)

_UA = os.environ.get("SEC_USER_AGENT", "Ramon Leal Doris rleal@topaz.com.mx")
_HEADERS = {"User-Agent": _UA, "Accept-Encoding": "gzip, deflate"}

# Bundled local copy of SEC's full ticker list. Avoids hitting the
# /files/company_tickers.json endpoint which is heavily rate-limited.
_BUNDLED_TICKERS_PATH = Path(__file__).resolve().parent / "company_tickers.json"


class EdgarClient:
    TICKERS_URL = "https://www.sec.gov/files/company_tickers.json"
    FACTS_URL = "https://data.sec.gov/api/xbrl/companyfacts/CIK{cik}.json"
    SUBMISSIONS_URL = "https://data.sec.gov/submissions/CIK{cik}.json"
    CONCEPT_URL = "https://data.sec.gov/api/xbrl/companyconcept/CIK{cik}/us-gaap/{concept}.json"

    def __init__(self, cache_dir: Optional[Path] = None):
        if cache_dir is None:
            cache_dir = Path(__file__).resolve().parents[2] / ".cache" / "edgar"
        self.cache = JsonCache(cache_dir)
        self._last_request_ts = 0.0

    def _throttle(self) -> None:
        # SEC: 10 req/s max. Keep 120ms between requests for safety.
        elapsed = time.time() - self._last_request_ts
        if elapsed < 0.12:
            time.sleep(0.12 - elapsed)
        self._last_request_ts = time.time()

    def _get(self, url: str, use_cache: bool = True) -> dict:
        if use_cache:
            cached = self.cache.get(url)
            if cached is not None:
                return cached

        # SEC rate-limits by IP. On shared-IP infra (Render, AWS, etc.) we
        # sometimes hit 429 even when well under the documented 10 req/s,
        # because other tenants on the same IP are burning quota. Retry with
        # exponential backoff up to 4 times before giving up.
        max_attempts = 4
        backoff_seconds = [2, 5, 15, 45]
        last_err: Exception | None = None
        for attempt in range(max_attempts):
            self._throttle()
            try:
                resp = requests.get(url, headers=_HEADERS, timeout=30)
                if resp.status_code == 429:
                    # Too many requests — wait and retry
                    wait = backoff_seconds[min(attempt, len(backoff_seconds) - 1)]
                    time.sleep(wait)
                    last_err = requests.HTTPError(
                        f"429 Too Many Requests from SEC (attempt {attempt + 1}/{max_attempts}); "
                        f"waited {wait}s before retry"
                    )
                    continue
                resp.raise_for_status()
                data = resp.json()
                if use_cache:
                    self.cache.set(url, data)
                return data
            except requests.HTTPError as e:
                if resp is not None and resp.status_code in (500, 502, 503, 504):
                    wait = backoff_seconds[min(attempt, len(backoff_seconds) - 1)]
                    time.sleep(wait)
                    last_err = e
                    continue
                raise
        # Exhausted all retries
        raise last_err or RuntimeError(f"SEC request failed after {max_attempts} attempts: {url}")

    # ------------------------------------------------------------------
    # Ticker → CIK
    # ------------------------------------------------------------------
    def _load_tickers(self) -> dict:
        """Load the ticker→CIK map. Prefers the bundled local file to avoid
        hitting SEC's rate-limited /files/company_tickers.json endpoint."""
        # 1. Try bundled file (shipped with the repo)
        if _BUNDLED_TICKERS_PATH.exists():
            try:
                with open(_BUNDLED_TICKERS_PATH, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        # 2. Try on-disk cache (set from a previous successful fetch)
        cached = self.cache.get(self.TICKERS_URL)
        if cached is not None:
            return cached
        # 3. Fall back to fetching from SEC (last resort; may hit 429)
        return self._get(self.TICKERS_URL)

    def ticker_to_cik(self, ticker: str) -> str:
        ticker = ticker.upper().strip()
        data = self._load_tickers()
        # data is a dict of {"0": {"cik_str": 320193, "ticker": "AAPL", "title": "Apple Inc."}, ...}
        for entry in data.values():
            if entry.get("ticker", "").upper() == ticker:
                return str(entry["cik_str"]).zfill(10)
        raise ValueError(f"Ticker {ticker} not found in SEC EDGAR company_tickers.json")

    def company_title(self, ticker: str) -> str:
        ticker = ticker.upper().strip()
        data = self._load_tickers()
        for entry in data.values():
            if entry.get("ticker", "").upper() == ticker:
                return entry["title"]
        return ticker

    # ------------------------------------------------------------------
    # XBRL company facts (all reported concepts)
    # ------------------------------------------------------------------
    def company_facts(self, cik: str) -> dict:
        cik = cik.zfill(10)
        url = self.FACTS_URL.format(cik=cik)
        return self._get(url)

    def submissions(self, cik: str) -> dict:
        cik = cik.zfill(10)
        url = self.SUBMISSIONS_URL.format(cik=cik)
        return self._get(url)

    def clear_cache(self) -> None:
        self.cache.clear()


if __name__ == "__main__":
    # Smoke test
    client = EdgarClient()
    cik = client.ticker_to_cik("COIN")
    print(f"COIN CIK: {cik}")
    title = client.company_title("COIN")
    print(f"Title: {title}")
    facts = client.company_facts(cik)
    concepts = list(facts.get("facts", {}).get("us-gaap", {}).keys())
    print(f"Total us-gaap concepts reported: {len(concepts)}")
    revenue_keys = [k for k in concepts if "revenue" in k.lower()]
    print(f"Revenue-related concepts: {revenue_keys[:10]}")
