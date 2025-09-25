from fastapi import APIRouter, Depends, Query
from fastapi.responses import StreamingResponse, JSONResponse
from typing import Optional, List
import io, csv

from backend.deps import require_api_key

# ملاحظات:
# - غيّر أسماء/مسارات الدوال تحت لو مختلفة في مشروعك
try:
    from backend.services.data_access import fetch_balance_sheet
except Exception:  # fallback لو الموديول مش متاح أثناء التطوير
    def fetch_balance_sheet(period, entity, offset, limit):
        rows = [
            {"ref":"1000","description":"Cash","entity":"HO","period":period,"amount":12345.67},
            {"ref":"1100","description":"AR","entity":"HO","period":period,"amount":890.12},
        ]
        return rows[offset:offset+limit], len(rows)

router = APIRouter(
    prefix="/api/v1/bi",
    tags=["bi"],
    dependencies=[Depends(require_api_key)],
)

@router.get("/balance-sheet.json")
async def bs_json(
    period: str = Query(..., description="e.g. 2024-12 or 2025-Q2 or 2025-06-30"),
    entity: Optional[List[str]] = Query(None, description="Repeatable: entity=HO&entity=Training"),
    offset: int = 0,
    limit: int = 1000,
):
    rows, total = fetch_balance_sheet(period, entity, offset, limit)
    return JSONResponse({"data": rows, "pagination": {"offset": offset, "limit": limit, "total": total}})

@router.get("/balance-sheet.csv")
async def bs_csv(
    period: str = Query(...),
    entity: Optional[List[str]] = Query(None),
    offset: int = 0,
    limit: int = 100000,
):
    rows, total = fetch_balance_sheet(period, entity, offset, limit)
    output = io.StringIO()
    writer = csv.DictWriter(output, fieldnames=["ref","description","entity","period","amount"])
    writer.writeheader()
    for r in rows:
        writer.writerow(r)
    output.seek(0)
    headers = {
        "X-Total-Count": str(total),
        "Content-Disposition": f"attachment; filename=balance_sheet_{period}.csv",
    }
    return StreamingResponse(iter([output.getvalue()]), media_type="text/csv", headers=headers)
