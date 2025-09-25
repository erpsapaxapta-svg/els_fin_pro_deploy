from __future__ import annotations

from datetime import datetime
from pathlib import Path
from urllib.parse import quote

from fastapi import APIRouter, HTTPException, Depends
from pydantic import BaseModel

from backend.deps import require_api_key
from backend.modules.bs_generators import build_excel_ui  # دالتك كما هي

# توحيد المسارات على /api/v1
router = APIRouter(
    prefix="/api/v1/build",
    tags=["build"],
    dependencies=[Depends(require_api_key)],
)

# مسارات الملفات داخل صورة الدوكر/المشروع
BASE_DIR = Path(__file__).resolve().parents[1]          # backend/
INPUT_FILE = BASE_DIR / "consolidation_inputs" / "Consolidated_TBs.xlsx"
OUTPUTS_DIR = BASE_DIR / "outputs"
OUTPUTS_DIR.mkdir(parents=True, exist_ok=True)


class BuildPayload(BaseModel):
    model: str
    periods: list[str] | None = None
    company: str | None = None


@router.post("")
def build_report(payload: BuildPayload):
    """
    يبني تقرير الـ Balance Sheet ويُرجِع JSON فيه رابط تحميل نسبي.
    POST /api/v1/build
    body: { "model": "bs", "periods": [...], "company": "HO"|"All"|None }
    """
    if (payload.model or "").lower() not in {"bs", "balance sheet", "balance_sheet"}:
        raise HTTPException(status_code=400, detail="Unsupported model (only BS supported)")

    # "All" = None
    company = None if (payload.company or "").strip().lower() == "all" else payload.company

    # اسم ملف ديناميكي
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_comp = (company or "ALL").replace(" ", "_")
    filename = f"build_bs_{safe_comp}_{stamp}.xlsx"
    out_path = OUTPUTS_DIR / filename

    # التوليد
    res = build_excel_ui(
        input_file=str(INPUT_FILE),
        output_file=str(out_path),
        periods=payload.periods,
        label_mode="PERIOD",
        company=company,
    )

    final_file = Path(res.get("output_file", out_path))
    if not final_file.exists():
        raise HTTPException(status_code=500, detail="Build succeeded but file not found")

    # هنرجّع لينك التحميل على الراوتر الموحد /api/v1/download
    return {
        "ok": True,
        "filename": final_file.name,
        "download_url": f"/api/v1/download?file={quote(final_file.name)}",
        "meta": {
            "company": res.get("company"),
            "periods": res.get("periods"),
        },
    }
