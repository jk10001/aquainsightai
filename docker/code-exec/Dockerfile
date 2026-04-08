# Use Python 3 slim as base (Debian-based)
FROM python:3-slim

# Avoid interactive apt prompts and speed up pip
ENV DEBIAN_FRONTEND=noninteractive \
    PIP_NO_CACHE_DIR=1 \
    # Help Kaleido find Chromium explicitly (usually auto-detected, but this is robust)
    CHROME_PATH=/usr/bin/chromium \
    # Ensure ~/.local/bin is available for pip --user installs
    PATH="/home/appuser/.local/bin:${PATH}"

# Install Chromium and a sensible font set for Plotly rendering
RUN apt-get update && apt-get install -y --no-install-recommends \
    chromium \
    fonts-dejavu \
    fonts-liberation \
    fonts-noto-color-emoji \
    && rm -rf /var/lib/apt/lists/*

# Create a non-root user so headless Chromium doesn't need --no-sandbox
RUN useradd -ms /bin/bash appuser
USER appuser

# Set working directory inside container
WORKDIR /app

# Copy requirements file
COPY --chown=appuser:appuser requirements-docker.txt .

# Install Python dependencies
RUN pip install -r requirements-docker.txt
