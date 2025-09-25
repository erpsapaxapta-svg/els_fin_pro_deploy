# -*- coding: utf-8 -*-
from __future__ import annotations
from pathlib import Path
from typing import List, Optional
from datetime import datetime
import yaml

from .engine.renderer import render_table
from .bs_generators import (
    build_excel_ui,
    list_periods as bs_list_periods,
    list_companies_and_sectors as bs_list_companies,
)

BASE_DIR = Path(__file__).resolve().parents[1]
OUTPUTS_DIR = BASE_DIR / "outputs"
CONFIG_DIR = BASE_DIR / "config" / "models"
INPUT_FILE = BASE_DIR / "consolidation_inputs" / "Consolidated_TBs.xlsx"

OUTPUTS_DIR.mkdir(exist_ok=True)
CONFIG_DIR.mkdir(parents=True, exist_ok=True)

# ---------- Helpers ----------
def _load_template(model: str) -> dict:
    cfg_path = CONFIG_DIR / f"{model}.yml"
    if not cfg_path.exists():
        raise FileNotFoundError(f"Template not found: {cfg_path.name}")
    return yaml.safe_load(cfg_path.read_text(encoding="utf-8"))

def get_periods_list() -> List[str]:
    try:
        return [p.upper() for p in bs_list_periods(str(INPUT_FILE))]
    except Exception:
        # fallback Ø«Ø§Ø¨ØªØ© Ù„Ùˆ Ø§Ù„Ù…Ù„Ù ØºÙŠØ± Ù…ØªØ§Ø­
        return ["DEC 2025","SEP 2025","JUN 2025","MAR 2025","DEC 2024","SEP 2024","JUN 2024","MAR 2024","DEC 2023"]

def get_companies_list() -> List[str]:
    try:
        comps, _ = bs_list_companies(str(INPUT_FILE))
        return comps
    except Exception:
        return ["HO", "Cairo Kh", "Alex", "Dubai"]

# ---------- Main ----------
def generate_multi(*, model: str, periods: List[str], company: Optional[str]) -> List[Path]:
    """
    - BS: ÙŠØ³ØªØ®Ø¯Ù… Ù…ÙˆÙ„Ù‘Ø¯ Ø§Ù„Ù…ÙŠØ²Ø§Ù†ÙŠØ© Ø§Ù„Ø­Ù‚ÙŠÙ‚ÙŠ (Ù…Ù† Excel).
    - PL/CF: ÙŠØ¨Ù†ÙŠ Ù…Ù† Ø§Ù„Ù‚Ø§Ù„Ø¨ (placeholder Ø­Ø§Ù„ÙŠÙ‹Ø§).
    """
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_company = (company or "All").replace(" ", "_")

    if model == "bs":
        out_xlsx = OUTPUTS_DIR / f"build_bs_{safe_company}_{len(periods)}_{ts}.xlsx"
        # Ù‡Ù†Ù„ØªØ²Ù… Ø¨Ø§Ù„ÙØªØ±Ø§Øª Ø§Ù„Ù…Ø®ØªØ§Ø±Ø© â€“ Ù„Ùˆ ØºÙŠØ± Ù…ØªØ§Ø­Ø© Ù‡ÙŠØ·Ù„Ø¹ Ø®Ø·Ø£ ÙˆØ§Ø¶Ø­ Ù…Ù† Ø§Ù„Ù…ÙˆÙ„Ù‘Ø¯
        build_excel_ui(
            input_file=str(INPUT_FILE),
            output_file=str(out_xlsx),
            periods=periods,
            company=company,
            label_mode="PERIOD",
        )
        return [out_xlsx]

    # --- Ø¨Ø§Ù‚ÙŠ Ø§Ù„Ù…ÙˆØ¯ÙŠÙ„Ø§Øª (Ù‚ÙˆØ§Ù„Ø¨) ---
    template = _load_template(model)
    out_xlsx = OUTPUTS_DIR / f"build_{model}_{safe_company}_{len(periods)}_{ts}.xlsx"
    render_table(template=template, periods=periods, company=company, out_path=out_xlsx)
    return [out_xlsx]

