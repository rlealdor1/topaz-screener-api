"""Claude-generated equity research report (.docx).

Calls Anthropic's Claude API with the user's existing prompt (prompt_consensus.txt)
plus structured financial data. Writes the result as a formatted Word document.

Also returns a structured dict of "one_pager_blocks" that the PPTX generator consumes.
"""
from __future__ import annotations
import json
import os
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional

from dotenv import load_dotenv
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from ..extract.income_statement import IncomeStatement
from ..model.dcf import DCFResult
from ..data.yfinance_client import QuoteSnapshot

load_dotenv(override=True)


@dataclass
class ReportPayload:
    """Structured output from Claude, parsed from JSON envelope."""
    executive_summary: Dict[str, str] = field(default_factory=dict)
    strategic_rationale: str = ""
    business_overview: str = ""
    opportunity: str = ""
    key_risks: str = ""
    watchpoints: str = ""
    recommendations: Dict[str, str] = field(default_factory=dict)  # 1yr/2yr/3yr
    consensus: str = ""
    # One-pager components (consumed by pptx generator)
    one_pager: Dict[str, object] = field(default_factory=dict)
    # Raw markdown in case parsing failed
    raw_markdown: str = ""


def _financial_context(stmt: IncomeStatement, dcf: DCFResult, quote: QuoteSnapshot) -> str:
    """Render compact financial context for Claude."""
    lines: List[str] = []
    lines.append(f"# {stmt.title} ({stmt.ticker})")
    lines.append(f"Current price: ${quote.current_price}")
    if quote.market_cap:
        lines.append(f"Market cap: ${quote.market_cap/1e9:.2f}B")
    if quote.enterprise_value:
        lines.append(f"Enterprise value: ${quote.enterprise_value/1e9:.2f}B")
    lines.append(f"Sector: {quote.sector} / Industry: {quote.industry}")
    lines.append("")
    lines.append("## Quarterly Income Statement (USD thousands)")
    period_labels = [p.label for p in stmt.periods]
    lines.append("Period | " + " | ".join(period_labels))
    for line in stmt.lines:
        row = [line.label]
        is_eps = line.section == "eps"
        for p in period_labels:
            v = line.values.get(p)
            if v is None:
                row.append("-")
            elif is_eps:
                row.append(f"{v:.2f}")
            else:
                row.append(f"{v/1000:,.0f}")
        lines.append(" | ".join(row))
    lines.append("")

    lines.append("## DCF Valuation Output (Bank-Style)")
    lines.append(f"Base year: {dcf.base_fy}, base revenue: ${dcf.base_revenue/1e9:.2f}B")
    lines.append(f"Methodology: {dcf.assumptions.explicit_years}Y explicit + {dcf.assumptions.fade_years}Y fade + dual terminal value (Gordon Growth + Exit Multiple)")
    w = dcf.wacc_base
    lines.append(f"WACC (CAPM): {w.wacc:.2%}  (Ke={w.cost_of_equity:.2%}, Kd={w.aftertax_cost_of_debt:.2%} after-tax, E/V={w.equity_weight:.0%}, β={w.inputs.beta:.2f})")
    lines.append(f"Terminal growth: {dcf.assumptions.terminal_growth:.1%},  Peer median EV/EBITDA: {dcf.peer_median_ev_ebitda:.1f}x" if dcf.peer_median_ev_ebitda else "")
    lines.append("")
    lines.append("Scenario outputs (blended Gordon Growth + Exit Multiple):")
    for name in ("Bear", "Base", "Bull"):
        s = dcf.scenarios[name]
        upside = (s.price_per_share / dcf.current_price - 1) if dcf.current_price else None
        up_str = f" ({upside:+.0%} vs current)" if upside is not None else ""
        lines.append(f"  {name}: WACC {s.wacc:.2%}, EV ${s.enterprise_value/1e9:.2f}B, Equity ${s.equity_value/1e9:.2f}B, Price ${s.price_per_share:.2f}{up_str}")
    lines.append("")

    # Analyst consensus context
    if dcf.analyst_count or dcf.analyst_target_mean:
        lines.append("## Analyst Consensus (sell-side)")
        if dcf.analyst_count:
            lines.append(f"Analyst coverage: {dcf.analyst_count} analysts, rating = {dcf.analyst_rec_key or 'n/a'}")
        if dcf.analyst_target_mean:
            lines.append(f"Price targets: low ${dcf.analyst_target_low:.2f} / mean ${dcf.analyst_target_mean:.2f} / high ${dcf.analyst_target_high:.2f}")
        if dcf.analyst_growth_y1 is not None:
            lines.append(f"Consensus revenue growth: Y1 {dcf.analyst_growth_y1:+.1%}, Y2 {dcf.analyst_growth_y2:+.1%}")
            lines.append(f"(DCF projections seeded from these consensus estimates, then fade to terminal growth.)")
        lines.append("")

    lines.append("## TTM Market Fundamentals")
    for attr in ("revenue_ttm", "ebitda_ttm", "free_cash_flow", "pe_ratio", "forward_pe",
                 "ev_ebitda", "ev_sales", "net_leverage"):
        v = getattr(quote, attr, None)
        if v is not None:
            if isinstance(v, float) and abs(v) > 1e6:
                lines.append(f"{attr}: ${v/1e9:.2f}B")
            else:
                lines.append(f"{attr}: {v}")

    return "\n".join(lines)


def _build_system_prompt(prompt_template: str) -> str:
    """Combine user's ChatGPT prompt with structured-output instructions."""
    return f"""{prompt_template}

---

OUTPUT FORMAT REQUIREMENTS:

Return your response as a single JSON object wrapped in <report> tags. Use this exact schema:

<report>
{{
  "executive_summary": {{
    "recommendation": "INVEST NOW" | "DO NOT INVEST" | "WAIT FOR [CONDITIONS]",
    "thesis": "2-3 sentence thesis statement",
    "current_price": "$XX.XX",
    "fair_value": "$XX.XX",
    "upside_downside_pct": "+XX%" | "-XX%",
    "price_target_6m": "$XX.XX (XX%)",
    "price_target_12m": "$XX.XX (XX%)",
    "price_target_24m": "$XX.XX (XX%)",
    "conviction": "High" | "Medium" | "Low",
    "key_strengths": ["bullet 1", "bullet 2", "bullet 3"],
    "key_risks": ["bullet 1", "bullet 2", "bullet 3"],
    "technical_view": "Bullish/Neutral/Bearish: brief explanation",
    "action_today": "buy at current / wait for pullback to $X / avoid"
  }},
  "strategic_rationale": "2-3 paragraphs on why this company matters strategically, market position, structural moats",
  "business_overview": "2-3 paragraphs on business model, revenue streams, recent earnings, key financial metrics",
  "opportunity": "2-3 paragraphs on growth vectors, TAM expansion, catalysts over 12-24 months",
  "key_risks": "2-3 paragraphs detailing 3-5 material risks with probability assessment",
  "watchpoints": "bullet list of specific metrics/events to monitor quarterly",
  "recommendations": {{
    "one_year": "Recommendation with target, 2-3 sentences",
    "two_year": "Recommendation with target, 2-3 sentences",
    "three_year": "Recommendation with target, 2-3 sentences"
  }},
  "consensus": "Final consensus paragraph — buy / hold / sell with reasoning",
  "one_pager": {{
    "tagline": "One-line company descriptor, e.g. 'AI Global Learning Platform'",
    "intro_paragraph": "2-3 sentences summarizing the company for the one-pager (150 words max)",
    "operations_footprint": ["bullet 1", "bullet 2", "bullet 3", "bullet 4", "bullet 5"],
    "key_strengths": ["bullet 1", "bullet 2", "bullet 3", "bullet 4", "bullet 5"],
    "key_risks": ["bullet 1", "bullet 2", "bullet 3", "bullet 4"],
    "primary_kpi_label": "Paid Subscribers" | "MAU" | "Transacting Users" | etc,
    "primary_kpi_series": {{"1Q23": value, "2Q23": value, ...}}  // optional, empty dict OK
  }}
}}
</report>

Write in the tone of an institutional equity research report — analytical, thesis-driven, data-heavy.
Use specific numbers from the provided financial context. Do not use emoji.
Base all conclusions on the data provided. Mention uncertainty where data is incomplete.
"""


def call_claude(stmt: IncomeStatement, dcf: DCFResult, quote: QuoteSnapshot,
                prompt_template: str, model: str = "claude-sonnet-4-6",
                max_tokens: int = 8000, temperature: float = 0.4) -> ReportPayload:
    """Call Claude and parse the structured response."""
    try:
        import anthropic
    except ImportError as e:
        raise RuntimeError("anthropic SDK not installed") from e

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise RuntimeError("ANTHROPIC_API_KEY not set in environment or .env file")

    client = anthropic.Anthropic(api_key=api_key)

    context = _financial_context(stmt, dcf, quote)
    system = _build_system_prompt(prompt_template)
    user_msg = (f"Analyze the following company using the framework above.\n\n"
                f"{context}\n\n"
                f"Generate the complete report in the specified JSON schema.")

    msg = client.messages.create(
        model=model,
        max_tokens=max_tokens,
        temperature=temperature,
        system=system,
        messages=[{"role": "user", "content": user_msg}],
    )

    text = "".join(block.text for block in msg.content if hasattr(block, "text"))
    return _parse_claude_response(text)


def _parse_claude_response(text: str) -> ReportPayload:
    """Extract the JSON envelope from Claude's response."""
    payload = ReportPayload(raw_markdown=text)
    match = re.search(r"<report>\s*(\{.*?\})\s*</report>", text, re.DOTALL)
    if not match:
        # Try to find a raw JSON block
        match = re.search(r"\{.*\"executive_summary\".*\}", text, re.DOTALL)
    if not match:
        return payload
    json_str = match.group(1) if match.re.pattern.startswith("<report>") else match.group(0)
    try:
        data = json.loads(json_str)
    except json.JSONDecodeError:
        return payload
    payload.executive_summary = data.get("executive_summary", {})
    payload.strategic_rationale = data.get("strategic_rationale", "")
    payload.business_overview = data.get("business_overview", "")
    payload.opportunity = data.get("opportunity", "")
    payload.key_risks = data.get("key_risks", "")
    payload.watchpoints = data.get("watchpoints", "")
    payload.recommendations = data.get("recommendations", {})
    payload.consensus = data.get("consensus", "")
    payload.one_pager = data.get("one_pager", {})
    return payload


# ---------- Word document writer ----------

def _set_paragraph_font(para, size=11, bold=False, color=None, italic=False):
    for run in para.runs:
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.italic = italic
        if color:
            run.font.color.rgb = RGBColor.from_string(color)


def _add_hr(doc):
    """Add a thin horizontal rule (paragraph with bottom border)."""
    p = doc.add_paragraph()
    pPr = p._element.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "A6A6A6")
    pBdr.append(bottom)
    pPr.append(pBdr)


def _add_section_heading(doc, text: str, size: int = 14):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.bold = True
    run.font.color.rgb = RGBColor.from_string("1F3864")
    _add_hr(doc)


def _add_body_paragraph(doc, text: str):
    if not text:
        return
    # Split by single newline to preserve paragraph structure
    for para_text in text.split("\n\n"):
        if not para_text.strip():
            continue
        p = doc.add_paragraph()
        run = p.add_run(para_text.strip())
        run.font.size = Pt(11)


def _add_bullet_list(doc, items: List[str]):
    for item in items:
        p = doc.add_paragraph(style="List Bullet")
        run = p.add_run(item)
        run.font.size = Pt(11)


def _add_kv_block(doc, pairs: List[tuple[str, str]]):
    """Add a key-value block for executive summary metrics."""
    for k, v in pairs:
        p = doc.add_paragraph()
        bold = p.add_run(f"{k}: ")
        bold.font.bold = True
        bold.font.size = Pt(11)
        val = p.add_run(str(v))
        val.font.size = Pt(11)


def write_docx(payload: ReportPayload, stmt: IncomeStatement, quote: QuoteSnapshot,
               output_path: Path) -> Path:
    doc = Document()

    # Title
    title = doc.add_paragraph()
    run = title.add_run(f"{stmt.ticker}  —  {stmt.title}")
    run.font.size = Pt(20)
    run.font.bold = True
    run.font.color.rgb = RGBColor.from_string("1F3864")

    subtitle = doc.add_paragraph()
    run = subtitle.add_run(f"Equity Research Report")
    run.font.size = Pt(12)
    run.font.italic = True
    run.font.color.rgb = RGBColor.from_string("595959")

    _add_hr(doc)

    # ===== Executive Summary =====
    _add_section_heading(doc, "Executive Summary")
    es = payload.executive_summary
    if es:
        _add_kv_block(doc, [
            ("Recommendation", es.get("recommendation", "-")),
            ("Thesis", es.get("thesis", "-")),
            ("Current Price", es.get("current_price", f"${quote.current_price or '-'}")),
            ("Fair Value", es.get("fair_value", "-")),
            ("Upside / Downside", es.get("upside_downside_pct", "-")),
            ("Price Target (6mo)", es.get("price_target_6m", "-")),
            ("Price Target (12mo)", es.get("price_target_12m", "-")),
            ("Price Target (24mo)", es.get("price_target_24m", "-")),
            ("Conviction", es.get("conviction", "-")),
            ("Technical View", es.get("technical_view", "-")),
            ("Action Today", es.get("action_today", "-")),
        ])
        doc.add_paragraph().add_run("Key Strengths").font.bold = True
        _add_bullet_list(doc, es.get("key_strengths", []))
        doc.add_paragraph().add_run("Key Risks").font.bold = True
        _add_bullet_list(doc, es.get("key_risks", []))

    # ===== Strategic Rationale =====
    _add_section_heading(doc, "Strategic Rationale")
    _add_body_paragraph(doc, payload.strategic_rationale)

    # ===== Business Overview =====
    _add_section_heading(doc, "Business Overview & Fundamentals")
    _add_body_paragraph(doc, payload.business_overview)

    # ===== Opportunity =====
    _add_section_heading(doc, "Opportunity & Catalysts")
    _add_body_paragraph(doc, payload.opportunity)

    # ===== Key Risks =====
    _add_section_heading(doc, "Key Risks & Headwinds")
    _add_body_paragraph(doc, payload.key_risks)

    # ===== Watchpoints =====
    _add_section_heading(doc, "Watchpoints")
    _add_body_paragraph(doc, payload.watchpoints)

    # ===== Recommendations 1/2/3-year =====
    _add_section_heading(doc, "Recommendations (1-3 Year View)")
    recs = payload.recommendations
    if recs:
        _add_kv_block(doc, [
            ("1-year", recs.get("one_year", "-")),
            ("2-year", recs.get("two_year", "-")),
            ("3-year", recs.get("three_year", "-")),
        ])

    # ===== Consensus =====
    _add_section_heading(doc, "Consensus")
    _add_body_paragraph(doc, payload.consensus)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(output_path)
    return output_path


def write_fallback_docx(stmt: IncomeStatement, quote: QuoteSnapshot, note: str,
                         output_path: Path) -> Path:
    """If Claude API is unavailable, write a skeleton doc with the data context and a note."""
    doc = Document()
    title = doc.add_paragraph()
    run = title.add_run(f"{stmt.ticker}  —  {stmt.title}")
    run.font.size = Pt(20); run.font.bold = True
    run.font.color.rgb = RGBColor.from_string("1F3864")

    p = doc.add_paragraph()
    r = p.add_run(note)
    r.font.italic = True
    r.font.color.rgb = RGBColor.from_string("C00000")

    doc.add_paragraph()
    doc.add_paragraph("Financial context (to paste into ChatGPT if you prefer manual generation):")
    from src.model.dcf import DCFResult  # type: ignore
    # Just the context string, no DCF passed
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(output_path)
    return output_path


def write_prompt_file(stmt: IncomeStatement, dcf: DCFResult, quote: QuoteSnapshot,
                      prompt_template: str, output_path: Path) -> Path:
    """Write the exact prompt + context to a .txt for manual iteration."""
    context = _financial_context(stmt, dcf, quote)
    system = _build_system_prompt(prompt_template)
    content = f"=== SYSTEM PROMPT ===\n\n{system}\n\n=== USER MESSAGE ===\n\nAnalyze the following company using the framework above.\n\n{context}\n"
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(content, encoding="utf-8")
    return output_path
