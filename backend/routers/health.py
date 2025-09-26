from fastapi import APIRouter
import time

router = APIRouter(prefix="/api/v1", tags=["health"])

@router.get("/health")
def health():
    return {"status": "ok", "ts": int(time.time())}
@router.get("/ready")
def ready():
    return {"ready": True}
