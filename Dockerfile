# Topaz Screener API — Docker image for Render / any container host.
# Python 3.12 slim base; includes build-essential for any C extensions
# (numpy/pandas wheels are typically prebuilt, so this is a small image).

FROM python:3.12-slim AS base

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1

WORKDIR /app

# System libs matplotlib and openpyxl need on slim images
RUN apt-get update && apt-get install -y --no-install-recommends \
        libfreetype6 \
        libpng16-16 \
        fonts-liberation \
        ca-certificates \
    && rm -rf /var/lib/apt/lists/*

# Install Python deps first (better layer caching)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy the rest of the app (src/, api/, templates/, config files)
COPY . .

# Render sets $PORT; default to 8000 locally.
ENV PORT=8000

EXPOSE 8000

# Uvicorn with a single worker — jobs run in BackgroundTasks so we don't
# need multiprocess. Bind to 0.0.0.0 so Render can route to the container.
CMD ["sh", "-c", "uvicorn api.main:app --host 0.0.0.0 --port ${PORT:-8000}"]
