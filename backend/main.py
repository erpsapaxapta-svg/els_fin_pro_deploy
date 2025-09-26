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
# ضيف الاستيراد ده مع بقية الاستيرادات في أعلى الملف
from fastapi.routing import APIRoute

# حطّ المسار التشخيصي تحت تعريف التطبيق مباشرة
@app.get("/__routes")
def list_routes():
    """تشخيص: يرجّع كل المسارات المسجّلة في التطبيق."""
    return [
        {"path": r.path, "name": r.name, "methods": list(r.methods)}
        for r in app.routes
        if isinstance(r, APIRoute)
    ]
    from fastapi.responses import JSONResponse
from fastapi.openapi.utils import get_openapi
from fastapi.openapi.docs import get_swagger_ui_html, get_redoc_html

# لو كان FastAPI ما سجّلش الـ docs تلقائيًا، نعملهم يدويًا:
@app.get("/openapi.json", include_in_schema=False)
def openapi():
    return JSONResponse(
        get_openapi(
            title=APP_TITLE,
            version=APP_VER,
            routes=app.routes,
            description="ELS Finance API",
        )
    )

@app.get("/docs", include_in_schema=False)
def swagger_docs():
    return get_swagger_ui_html(openapi_url="/openapi.json", title="API Docs")

@app.get("/redoc", include_in_schema=False)
def redoc_docs():
    return get_redoc_html(openapi_url="/openapi.json", title="API ReDoc")


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
