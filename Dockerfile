# SlideArabi 1.0.0 — Railway deployment
# ──────────────────────────────────────
FROM python:3.12-slim

# Install LibreOffice (headless) for slide preview rendering + curl for API calls
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
        libreoffice-impress \
        libreoffice-common \
        curl \
        fonts-noto-core \
        fonts-noto-extra \
        fonts-noto-cjk \
    && rm -rf /var/lib/apt/lists/*

# Set working directory
WORKDIR /app

# Copy requirements first (layer caching)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy the full slidearabi package
COPY . /app/slidearabi/
# Ensure __init__.py exists
RUN echo "" > /app/slidearabi/__init__.py

# Create tmp dir for job files
RUN mkdir -p /tmp/slideshift_jobs

# Expose port (Railway injects $PORT)
EXPOSE 8000

# Health check
HEALTHCHECK --interval=30s --timeout=10s --start-period=30s \
    CMD curl -f http://localhost:${PORT:-8000}/health || exit 1

# Run with uvicorn; Railway sets PORT env var
CMD ["sh", "-c", "uvicorn slidearabi.server:app --host 0.0.0.0 --port ${PORT:-8000} --workers 1"]
