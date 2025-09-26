# backend/routers/meta.py
from fastapi import APIRouter, Depends
from backend.deps import require_api_key
from backend.modules.bs_generators import (
    list_periods,
    list_companies_and_sectors,
    INPUT_FILE,
)

router = APIRouter(
    prefix="/api/v1/meta",
    tags=["meta"],
    # لو عايزها عامة احذف السطر التالي
    dependencies=[Depends(require_api_key)],
)


@router.get("/companies")
def get_companies():
    """يرجع قائمة الشركات (كما كانت من قبل)."""
    comps, _ = list_companies_and_sectors(INPUT_FILE)
    return comps


@router.get("/periods")
def get_periods():
    """يرجع قائمة الفترات المتاحة (من ملف Consolidated_TBs.xlsx)."""
    return list_periods(INPUT_FILE)
