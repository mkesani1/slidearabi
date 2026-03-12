# SlideArabi 1.1.0 — Railway deployment (optimized for memory)
# ────────────────────────────────────────────────────────────
FROM python:3.12-slim

# Install LibreOffice (headless) for slide preview rendering,
# poppler-utils (pdftoppm) for per-page PDF→JPEG extraction,
# curl for health checks / geo-pricing lookups.
# NOTE: fonts-noto-cjk REMOVED — not needed for Arabic, and adds 200-400MB
#       that causes OOM kills on Railway's 512MB Hobby tier.
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
        libreoffice-impress \
        libreoffice-common \
        poppler-utils \
        curl \
        fontconfig \
        fonts-noto-core \
        fonts-hosny-amiri \
    && fc-cache -fv \
    && rm -rf /var/lib/apt/lists/* /usr/share/doc/* /usr/share/man/*

# Set working directory
WORKDIR /app

# Copy requirements first (layer caching)
COPY requirements.txt .
# v1.1.3: force fastmcp==3.1.0 for MCP server support
RUN pip install --no-cache-dir -r requirements.txt && pip show fastmcp | head -2

# Copy ONLY Python source files into the package (see .dockerignore)
COPY . /app/slidearabi/
# Ensure __init__.py exists
RUN test -f /app/slidearabi/__init__.py || echo "" > /app/slidearabi/__init__.py

# Create tmp dir for job files
RUN mkdir -p /tmp/slideshift_jobs

# Expose port (Railway injects $PORT)
EXPOSE 8000

# Railway handles health checks via railway.toml — no Docker HEALTHCHECK needed
# Print memory at startup for diagnostics, then launch uvicorn
CMD ["sh", "-c", "echo '[BOOT] Memory:' && cat /proc/meminfo | head -3 && echo '[BOOT] PORT='$PORT && exec uvicorn slidearabi.server:app --host 0.0.0.0 --port ${PORT:-8000} --workers 1"]
