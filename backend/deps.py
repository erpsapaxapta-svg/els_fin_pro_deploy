# backend/deps.py
from fastapi import Header, HTTPException
import os

EXPECTED = os.getenv("EXT_API_KEY", "change-me-please")

def require_api_key(x_api_key: str | None = Header(default=None, alias="X-API-KEY")):
    if x_api_key != EXPECTED:
        raise HTTPException(status_code=401, detail="invalid api key")
    return True
