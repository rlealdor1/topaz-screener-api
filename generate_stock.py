#!/usr/bin/env python
"""Stock research generator CLI.

Usage:
    python generate_stock.py COIN
    python generate_stock.py COIN --only excel
    python generate_stock.py COIN --peers HOOD,MARA,RIOT,CME
    python generate_stock.py COIN --refresh
"""
from __future__ import annotations
import argparse
import sys
import time
import traceback
from pathlib import Path

import yaml
from dotenv import load_dotenv

load_dotenv(override=True)


HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))


def _load_config() -> dict:
    with open(HERE / "config.yaml", "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def _load_prompt_template() -> str:
    return (HERE / "templates" / "prompt_consensus.txt").read_text(encoding="utf-8")


def _timed(msg: str):
    class T:
        def __enter__(self):
            print(f"  → {msg}...", flush=True)
            self.t0 = time.time()
            return self
        def __exit__(self, *a):
            dt = time.time() - self.t0
            print(f"    done in {dt:.1f}s", flush=True)
    return T()


def main():
    parser = argparse.ArgumentParser(description="Generate stock research deliverables.")
    parser.add_argument("ticker", help="Stock ticker (e.g. COIN, NFLX, OSCR)")
    parser.add_argument("--only", choices=["excel", "comps", "word", "onepager", "all"],
                        default="all", help="Generate only this deliverable")
    parser.add_argument("--peers", help="Comma-separated peer tickers to override peers.yaml")
    parser.add_argument("--refresh", action="store_true", help="Clear SEC cache and re-fetch")
    parser.add_argument("--output-dir", help="Override output root (default: Auto-generated/{TICKER})")
    args = parser.parse_args()

    ticker = args.ticker.upper().strip()
    config = _load_config()
    out_root = Path(args.output_dir) if args.output_dir else HERE / config["output"]["root"] / ticker
    out_root.mkdir(parents=True, exist_ok=True)

    print(f"\n=== Generating research for {ticker} → {out_root} ===\n")

    # Step 1: SEC data
    from src.data.edgar import EdgarClient
    from src.extract.income_statement import extract_income_statement
    from src.model.dcf import run_dcf, DCFAssumptions

    client = EdgarClient()
    if args.refresh:
        client.clear_cache()

    with _timed(f"Looking up CIK for {ticker}"):
        cik = client.ticker_to_cik(ticker)
        title = client.company_title(ticker)
    with _timed("Fetching SEC company facts"):
        facts = client.company_facts(cik)
    with _timed("Extracting quarterly income statement"):
        stmt = extract_income_statement(ticker, cik, title, facts,
                                        quarters_back=config["history"]["quarters_back"])

    # Step 2: Market data + peers (always needed for WACC / exit multiple)
    from src.data.yfinance_client import get_quote
    from src.extract.peers import get_peers
    with _timed("Fetching current quote + fundamentals (yfinance)"):
        quote = get_quote(ticker)

    peers_cli = args.peers.split(",") if args.peers else None
    peer_tickers = get_peers(ticker, HERE / "peers.yaml", peers_cli,
                             sector=quote.sector)
    peer_quotes = []
    if peer_tickers:
        with _timed(f"Fetching peer quotes ({len(peer_tickers)} tickers)"):
            peer_quotes = [get_quote(t) for t in peer_tickers]
        # Filter out peers yfinance couldn't resolve (missing market cap)
        peer_quotes = [p for p in peer_quotes if p and p.market_cap]

    # Step 3: Bank-style DCF (Bull/Base/Bear)
    cfg_dcf = config["dcf"]
    cfg_wacc = config["wacc"]
    assumptions = DCFAssumptions(
        explicit_years=cfg_dcf.get("explicit_years", 5),
        fade_years=cfg_dcf.get("fade_years", 5),
        terminal_growth=cfg_dcf.get("terminal_growth", 0.025),
        tax_rate=cfg_dcf.get("tax_rate", 0.21),
        revenue_growth_cap=cfg_dcf.get("revenue_growth_cap", 0.30),
        revenue_growth_floor=cfg_dcf.get("revenue_growth_floor", -0.05),
        ebit_margin_premium_bull=cfg_dcf.get("ebit_margin_premium_bull", 0.03),
        ebit_margin_premium_base=cfg_dcf.get("ebit_margin_premium_base", 0.01),
        ebit_margin_premium_bear=cfg_dcf.get("ebit_margin_premium_bear", -0.02),
        nwc_pct_delta_revenue=cfg_dcf.get("nwc_pct_delta_revenue", 0.05),
        sbc_as_cash_expense=cfg_dcf.get("sbc_as_cash_expense", True),
        capex_fade_years=cfg_dcf.get("capex_fade_years", 5),
        terminal_capex_pct_revenue=cfg_dcf.get("terminal_capex_pct_revenue", 0.08),
        sbc_fade_years=cfg_dcf.get("sbc_fade_years", 5),
        terminal_sbc_pct_revenue=cfg_dcf.get("terminal_sbc_pct_revenue", 0.03),
        model_buybacks=cfg_dcf.get("model_buybacks", True),
        buyback_yield_cap=cfg_dcf.get("buyback_yield_cap", 0.06),
        terminal_growth_megacap=cfg_dcf.get("terminal_growth_megacap", 0.035),
        megacap_threshold=cfg_dcf.get("megacap_threshold", 100_000_000_000.0),
        terminal_revenue_multiple_cap=cfg_dcf.get("terminal_revenue_multiple_cap", 20.0),
        wacc_adjustment_bull=cfg_dcf.get("wacc_adjustment_bull", -0.010),
        wacc_adjustment_bear=cfg_dcf.get("wacc_adjustment_bear", 0.010),
        tv_weight_gordon=cfg_dcf.get("tv_weight_gordon", 0.4),
        tv_weight_exit=cfg_dcf.get("tv_weight_exit", 0.6),
    )
    with _timed("Running bank-style DCF (Bull/Base/Bear)"):
        dcf = run_dcf(stmt, assumptions, cfg_wacc, facts, quote=quote, peer_quotes=peer_quotes)

    # Step 4: Football field chart — generated whenever Excel or one-pager is needed,
    # since both embed it.
    football_png = None
    if args.only in ("excel", "onepager", "all"):
        from src.output.football_field import build_football_field, build_bands_from_dcf
        bands = build_bands_from_dcf(dcf)
        if bands:
            football_png = out_root / f"{ticker}_football_field.png"
            with _timed("Building football field chart"):
                build_football_field(bands, dcf.current_price, ticker, football_png)

    # Step 5: Excel financials (IS + Valuation with embedded football field)
    if args.only in ("excel", "all"):
        from src.output.excel_financials import build_workbook
        with _timed(f"Building {ticker}.xlsx"):
            build_workbook(stmt, dcf, out_root / f"{ticker}.xlsx",
                           current_share_price=quote.current_price,
                           current_market_cap=quote.market_cap,
                           football_field_png=football_png)

    # Step 6: Comps
    if args.only in ("comps", "all"):
        from src.output.excel_comps import build_comps_workbook
        if not peer_tickers:
            print(f"  ⚠ No peers configured for {ticker} in peers.yaml and none provided via --peers")
        else:
            with _timed(f"Building {ticker} Comps.xlsx"):
                build_comps_workbook(quote, peer_quotes, out_root / f"{ticker} Comps.xlsx")

    # Step 7: Claude-generated Word report
    payload = None
    if args.only in ("word", "onepager", "all"):
        from src.output.word_report import call_claude, write_docx, write_prompt_file, ReportPayload
        prompt_template = _load_prompt_template()

        # Always write the prompt.txt for manual iteration
        with _timed(f"Writing {ticker}_prompt.txt"):
            write_prompt_file(stmt, dcf, quote, prompt_template,
                              out_root / f"{ticker}_prompt.txt")

        try:
            with _timed(f"Calling Claude ({config['claude']['model']}) for research report"):
                payload = call_claude(stmt, dcf, quote, prompt_template,
                                      model=config["claude"]["model"],
                                      max_tokens=config["claude"]["max_tokens"],
                                      temperature=config["claude"]["temperature"])
            if args.only in ("word", "all"):
                with _timed(f"Writing {ticker}.docx"):
                    write_docx(payload, stmt, quote, out_root / f"{ticker}.docx")
        except Exception as e:
            print(f"  ⚠ Claude generation failed: {e}")
            print(f"     The prompt.txt was still written — you can paste it into ChatGPT manually.")
            payload = ReportPayload()

    # Step 7: PPTX one-pager (bank-style sell-side layout)
    if args.only in ("onepager", "all"):
        from src.output.pptx_onepager import build_onepager
        one_pager_content = payload.one_pager if payload else {}
        # Merge executive summary into the one-pager content so the recommendation
        # chip and thesis-level bullets can be sourced from Claude's full response.
        if payload and payload.executive_summary:
            one_pager_content.setdefault("executive_summary", payload.executive_summary)
        with _timed(f"Building {ticker}.pptx one-pager"):
            build_onepager(stmt, quote, one_pager_content,
                           out_root / f"{ticker}.pptx",
                           dcf=dcf, football_field_png=football_png)

    print(f"\n✔ All deliverables written to: {out_root}\n")
    for f in sorted(out_root.glob("*")):
        print(f"    {f.name}  ({f.stat().st_size:,} bytes)")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\nError: {e}", file=sys.stderr)
        traceback.print_exc()
        sys.exit(1)
