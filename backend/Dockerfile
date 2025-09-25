# ---------- build stage ----------
FROM python:3.11-slim AS build
WORKDIR /app

ENV PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1

# context is ./backend, so paths are relative to it
COPY requirements.txt /app/requirements.txt

RUN python -m venv /opt/venv \
 && /opt/venv/bin/pip install --upgrade pip \
 && /opt/venv/bin/pip install --no-cache-dir -r /app/requirements.txt

# copy backend source (relative to ./backend)
COPY . /app/backend

# ---------- run stage (non-root) ----------
FROM python:3.11-slim
WORKDIR /app

RUN adduser --disabled-password --gecos "" app && chown -R app:app /app
USER app

ENV PATH="/opt/venv/bin:$PATH"

COPY --from=build /opt/venv /opt/venv
COPY --from=build /app/backend /app/backend

EXPOSE 8000

# production server (compose can override if needed)
CMD ["gunicorn","-k","uvicorn.workers.UvicornWorker","backend.main:app","--bind","0.0.0.0:8000","--workers","4","--timeout","60","--keep-alive","5"]
