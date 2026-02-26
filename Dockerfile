FROM python:3.11-slim

# Install system dependencies for Playwright/Chromium
RUN apt-get update && apt-get install -y --no-install-recommends \
    libnss3 libnspr4 libatk1.0-0 libatk-bridge2.0-0 libcups2 \
    libdrm2 libxkbcommon0 libxcomposite1 libxdamage1 libxrandr2 \
    libgbm1 libpango-1.0-0 libcairo2 libasound2 libatspi2.0-0 \
    libxshmfence1 fonts-liberation libappindicator3-1 \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install Python dependencies
COPY back/requirements.txt ./back/requirements.txt
RUN pip install --no-cache-dir -r back/requirements.txt

# Install Playwright Chromium browser
RUN playwright install chromium

# Copy project files
COPY back/ ./back/
COPY front/ ./front/

# Create data directories
RUN mkdir -p back/uploads back/data back/exports

# Set environment variables
ENV FLASK_ENV=production
ENV PORT=8080

EXPOSE 8080

# Run with Gunicorn
CMD gunicorn --chdir back --bind 0.0.0.0:$PORT --workers 2 --timeout 120 app:app
