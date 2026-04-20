"""In-memory job store + background worker that runs the screener pipeline.

One job_id → one dict in JOBS. The worker updates `status` and `step` as it
progresses so the polling endpoint can surface live progress.

For team-scale usage (a few jobs per hour) this in-memory store is fine.
If we ever need persistence across Render restarts, swap to a Supabase
`research_jobs` table.
"""
from __future__ import annotations
import logging
import threading
import traceback
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, Optional

import yaml
from dotenv import load_dotenv

load_dotenv(override=True)

# Screener imports — reuse the exact same pipeline the CLI uses.
from src.data.edgar import EdgarClient
from src.data.yfinance_client import get_quote
from src.extract.income_statement import extract_income_statement
from src.extract.peers import get_peers
from src.model.dcf import DCFAssumptions, run_dcf
from src.output.excel_comps import build_comps_workbook
from src.output.excel_financials import build_workbook
from src.output.football_field import build_bands_from_dcf, build_football_field
from src.output.pptx_onepager import build_onepager
from src.output.word_report import (
    ReportPayload,
    call_claude,
    write_docx,
    write_prompt_file,
)
from .supabase_storage import upload_deliverables


log = logging.getLogger("screener.jobs")

# Thread-safe job store. Single-instance Render service → plain dict is fine.
JOBS: Dict[str, Dict] = {}
_LOCK = threading.Lock()

# Project root (repo root) — files get generated into a per-job subdir.
_HERE = Path(__file__).resolve().parent.parent
_CONFIG_PATH = _HERE / "config.yaml"
_PEERS_PATH = _HERE / "peers.yaml"
_PROMPT_PATH = _HERE / "templates" / "prompt_consensus.txt"
_WORK_ROOT = _HERE / "_jobs"   # ephemeral local scratch area


def _load_config() -> dict:
    with open(_CONFIG_PATH, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def _update(job_id: str, **changes) -> None:
    with _LOCK:
        if job_id in JOBS:
            JOBS[job_id].update(changes)


def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


# ---------- Worker ----------


def run_job(job_id: str, ticker: str) -> None:
    """Run the full screening pipeline for one ticker. Updates JOBS as it goes."""
    ticker = ticker.upper().strip()
    work_dir = _WORK_ROOT / job_id
    work_dir.mkdir(parents=True, exist_ok=True)

    try:
        _update(job_id, status="running", step="Fetching SEC data", started_at=_now_iso())
        config = _load_config()

        # ---- 1. SEC data ----
        client = EdgarClient()
        cik = client.ticker_to_cik(ticker)
        title = client.company_title(ticker)
        facts = client.company_facts(cik)
        stmt = extract_income_statement(
            ticker, cik, title, facts,
            quarters_back=config["history"]["quarters_back"],
        )

        _update(job_id, step="Fetching market data")
        quote = get_quote(ticker)

        peer_tickers = get_peers(ticker, _PEERS_PATH, None)
        peer_quotes = [get_quote(t) for t in peer_tickers] if peer_tickers else []

        # ---- 2. DCF ----
        _update(job_id, step="Running bank-style DCF")
        cfg_dcf = config["dcf"]
        assumptions = DCFAssumptions(**{
            k: v for k, v in cfg_dcf.items() if k in DCFAssumptions.__dataclass_fields__
        })
        dcf = run_dcf(stmt, assumptions, config["wacc"], facts,
                      quote=quote, peer_quotes=peer_quotes)

        # ---- 3. Football field ----
        _update(job_id, step="Building football field")
        football_png = work_dir / f"{ticker}_football_field.png"
        bands = build_bands_from_dcf(dcf)
        if bands:
            build_football_field(bands, dcf.current_price, ticker, football_png)
        else:
            football_png = None  # type: ignore

        # ---- 4. Excel (IS + Valuation) ----
        _update(job_id, step="Generating Excel workbook")
        excel_path = work_dir / f"{ticker}.xlsx"
        build_workbook(
            stmt, dcf, excel_path,
            current_share_price=quote.current_price,
            current_market_cap=quote.market_cap,
            football_field_png=football_png,
        )

        # ---- 5. Comps ----
        comps_path: Optional[Path] = None
        if peer_quotes:
            _update(job_id, step="Building comps table")
            comps_path = work_dir / f"{ticker} Comps.xlsx"
            build_comps_workbook(quote, peer_quotes, comps_path)

        # ---- 6. Word report (Claude) ----
        _update(job_id, step="Calling Claude for research report")
        prompt_template = _PROMPT_PATH.read_text(encoding="utf-8")
        prompt_txt_path = work_dir / f"{ticker}_prompt.txt"
        write_prompt_file(stmt, dcf, quote, prompt_template, prompt_txt_path)

        word_path: Optional[Path] = None
        payload: Optional[ReportPayload] = None
        try:
            payload = call_claude(
                stmt, dcf, quote, prompt_template,
                model=config["claude"]["model"],
                max_tokens=config["claude"]["max_tokens"],
                temperature=config["claude"]["temperature"],
            )
            word_path = work_dir / f"{ticker}.docx"
            write_docx(payload, stmt, quote, word_path)
        except Exception as e:
            log.exception("Claude call failed; prompt still written for manual fallback")
            _update(job_id, claude_error=str(e))

        # ---- 7. PPTX one-pager ----
        _update(job_id, step="Building one-pager")
        one_pager_content = {}
        if payload:
            if payload.one_pager:
                one_pager_content.update(payload.one_pager)
            if payload.executive_summary:
                one_pager_content["executive_summary"] = payload.executive_summary
        pptx_path = work_dir / f"{ticker}.pptx"
        build_onepager(
            stmt, quote, one_pager_content, pptx_path,
            dcf=dcf, football_field_png=football_png,
        )

        # ---- 8. Upload to Supabase ----
        _update(job_id, step="Uploading deliverables")
        local_files: Dict[str, Path] = {"excel": excel_path, "pptx": pptx_path}
        if comps_path:
            local_files["comps"] = comps_path
        if word_path:
            local_files["word"] = word_path
        if football_png:
            local_files["football"] = football_png
        if prompt_txt_path.exists():
            local_files["prompt"] = prompt_txt_path

        urls = upload_deliverables(ticker, local_files)

        _update(
            job_id,
            status="complete",
            step="Complete",
            deliverables=urls,
            completed_at=_now_iso(),
            company_name=title,
            current_price=quote.current_price,
            analyst_target_mean=quote.analyst_target_mean,
            dcf_base_price=dcf.base.price_per_share if dcf.base else None,
            recommendation=(payload.executive_summary.get("recommendation")
                            if payload and payload.executive_summary else None),
        )

    except Exception as e:
        log.exception("Job %s failed", job_id)
        _update(
            job_id,
            status="failed",
            step="Failed",
            error=str(e),
            error_type=type(e).__name__,
            traceback=traceback.format_exc()[-1500:],
            completed_at=_now_iso(),
        )
