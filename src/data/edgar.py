"""SEC EDGAR client — ticker lookup, company facts (XBRL), filings submissions.

Uses the free, official SEC JSON endpoints. Requires User-Agent header.
Rate limits to 10 req/s (SEC hard cap).
"""
from __future__ import annotations
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
        self._throttle()
        resp = requests.get(url, headers=_HEADERS, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        if use_cache:
            self.cache.set(url, data)
        return data

    # ------------------------------------------------------------------
    # Ticker → CIK
    # ------------------------------------------------------------------
    def ticker_to_cik(self, ticker: str) -> str:
        ticker = ticker.upper().strip()
        data = self._get(self.TICKERS_URL)
        # data is a dict of {"0": {"cik_str": 320193, "ticker": "AAPL", "title": "Apple Inc."}, ...}
        for entry in data.values():
            if entry.get("ticker", "").upper() == ticker:
                return str(entry["cik_str"]).zfill(10)
        raise ValueError(f"Ticker {ticker} not found in SEC EDGAR company_tickers.json")

    def company_title(self, ticker: str) -> str:
        ticker = ticker.upper().strip()
        data = self._get(self.TICKERS_URL)
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
