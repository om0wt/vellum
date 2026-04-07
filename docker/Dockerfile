# Build context is the PROJECT ROOT (set by docker/docker-compose.yml's
# `context: ..`). All COPY paths are relative to that root, so they
# reference src/... and requirements.txt at the project root.

FROM python:3.11-slim

WORKDIR /app

# System packages:
# * tesseract-ocr + language packs — required by the OCR (scanned PDF)
#   path. Without these, the web app's OCR option is hidden because
#   list_tesseract_languages() returns None at startup. To add more
#   languages, append additional `tesseract-ocr-XXX` packages here
#   (see https://packages.debian.org/search?keywords=tesseract-ocr).
#
# pdf2docx and PyMuPDF Python wheels are self-contained on linux/amd64
# and arm64, so we don't need any apt packages for those.
RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        tesseract-ocr \
        tesseract-ocr-eng \
        tesseract-ocr-slk \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Application source. Copied from src/ at the project root into /app
# (the WORKDIR), so the in-container layout looks like:
#   /app/pdf_to_docx.py
#   /app/app.py
#   /app/templates/index.html
COPY src/pdf_to_docx.py src/app.py ./
COPY src/templates/ ./templates/

# Pre-create the access log directory so the app can mount the same path
# as a volume from the host without permission issues.
RUN mkdir -p /app/logs

# EXPOSE is metadata only — the actual listen port is decided by the
# PORT env var passed in by docker-compose.yml (which sets it to 5001).
# This number is purely for documentation / `docker inspect`.
EXPOSE 5001

CMD ["python", "app.py"]
