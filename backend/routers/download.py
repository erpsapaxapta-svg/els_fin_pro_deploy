# backend/routers/download.py
from pathlib import Path

from fastapi import APIRouter, HTTPException, Query, Depends
from fastapi.responses import FileResponse

from backend.deps import require_api_key

router = APIRouter(
    prefix="/api/v1",
    tags=["download"],
    dependencies=[Depends(require_api_key)],  # حماية بالرأس/البارام X-API-KEY
)

BASE_DIR = Path(__file__).resolve().parents[1]  # backend/
OUTPUTS_DIR = (BASE_DIR / "outputs").resolve()


@router.get("/download")
def download(file: str = Query(..., description="filename inside outputs/")):
    """
    يحمّل ملفًا من مجلد outputs/ بشكل آمن.
    GET /api/v1/download?file=<filename>
    """
    # منع الـ path traversal ومحاولة الخروج من outputs/
    p = (OUTPUTS_DIR / file).resolve()
    if OUTPUTS_DIR not in p.parents or not p.is_file():
        raise HTTPException(status_code=404, detail="File not found")

    return FileResponse(
        path=p,
        media_type=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            if p.suffix.lower() in {".xlsx", ".xlsm", ".xls"} else "application/octet-stream"
        ),
        filename=p.name,
        headers={
            "Cache-Control": "no-store",
            "Content-Disposition": f'attachment; filename="{p.name}"',
        },
    )
