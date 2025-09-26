# ===== Base image (مثبّتة) =====
FROM python:3.11-slim-bullseye

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1

# تثبيت libpq وغيرها لو بتحتاجها (لو هتتكامل مع Postgres)
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential curl && \
    rm -rf /var/lib/apt/lists/*

WORKDIR /app

# نسخ المتطلبات أولاً عشان الـlayer يتخزّن
COPY backend/requirements.txt /app/requirements.txt
RUN pip install --upgrade pip && pip install -r /app/requirements.txt

# نسخ الكود
COPY backend /app/backend

# ===== تشغيل Uvicorn في وضع إنتاج =====
EXPOSE 8000
CMD ["python", "-m", "uvicorn", "backend.main:app", "--host", "0.0.0.0", "--port", "8000"]
