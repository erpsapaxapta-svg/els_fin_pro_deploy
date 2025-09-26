import os, json, time
from typing import List
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

# احتفظ بروتراتك كما هي
from backend.routers import health, bi_exports, build, download, meta

APP_TITLE = "ELS Finance API"
APP_VER = os.getenv("APP_VER", "1.0.0")

# CORS: يقبل JSON Array أو قائمة مفصولة بفواصل
_raw = os.getenv("ALLOWED_ORIGINS", "")
try:
    allowed: List[str] = (json.loads(_raw) if _raw.strip().startswith("[")
                          else [o.strip() for o in _raw.split(",") if o.strip()])
except Exception:
    allowed = []
if not allowed:
    allowed = ["http://localhost:3000", "http://127.0.0.1:3000", "https://app.powerbi.com"]

app = FastAPI(
    title=APP_TITLE,
    version=APP_VER,
    description="Public API for BI tools (JSON/CSV) — versioned on /api/v1. Use X-API-KEY header or api_key query.",
    docs_url="/docs",
    redoc_url="/redoc",
    openapi_url="/openapi.json",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=allowed,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# -------- Health endpoints --------
@app.get("/api/v1/live")
def live():
    return {"status": "live", "ts": int(time.time())}

@app.get("/api/v1/ready")
async def ready():
    # (بسيطة الآن؛ ممكن نضيف فحص DB لاحقًا)
    return {"status": "ready"}

# -------- باقي الروترات كما هي --------
app.include_router(health.router)
app.include_router(bi_exports.router)
app.include_router(build.router)
app.include_router(download.router)
app.include_router(meta.router)

@app.get("/")
def root():
    return {"service": "els-finance-api", "docs": "/docs", "version": "v1"}
