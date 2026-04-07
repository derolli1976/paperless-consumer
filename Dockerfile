FROM python:3.12-slim

# Set working directory
WORKDIR /app

# Install LibreOffice for Word-to-PDF conversion
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-writer \
    && rm -rf /var/lib/apt/lists/*

# Copy requirements first for better caching
COPY requirements.txt .

# Install Python dependencies
RUN python -m pip install --upgrade pip --retries 5 --timeout 60 && \
    pip install --no-cache-dir --retries 5 --timeout 60 -r requirements.txt

# Copy application
COPY consumer.py .

# Create non-root user
RUN useradd -m -u 1000 consumer

# Create data directory for import log and set ownership
RUN mkdir -p /app/data && \
    chown -R consumer:consumer /app

# Switch to non-root user
USER consumer

# Health check
HEALTHCHECK --interval=60s --timeout=10s --start-period=30s --retries=3 \
    CMD pgrep -f consumer.py || exit 1

# Run application (unbuffered output for logs)
CMD ["python", "-u", "consumer.py"]
