#!/usr/bin/env bash
# Build a standalone Windows distribution of gui.py via PyInstaller-on-Wine
# in Docker.
#
# Run from the project root or from build-windows/ — both work.
# Output: ./dist/PDF-to-DOCX/    (folder containing PDF-to-DOCX.exe + runtime)
# Distribute the whole folder, e.g. zipped.
set -euo pipefail

# Resolve project root regardless of where this script is invoked from
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
cd "$PROJECT_ROOT"

IMAGE_TAG="pdf2docx-windows-builder"

mkdir -p dist

echo ">> Building builder image ($IMAGE_TAG)…"
docker build --platform linux/amd64 -t "$IMAGE_TAG" -f build-windows/Dockerfile .

echo
echo ">> Running PyInstaller via Wine…"
# Mount the project's `src/` directory as /src in the container so the
# Dockerfile CMD's `/src/gui.py` and `/src/pdf_to_docx.py` paths resolve.
docker run --rm --platform linux/amd64 -v "$PROJECT_ROOT/src:/src" "$IMAGE_TAG"

echo
DIST_DIR="$PROJECT_ROOT/dist/PDF-to-DOCX"
if [ -d "$DIST_DIR" ] && [ -f "$DIST_DIR/PDF-to-DOCX.exe" ]; then
    folder_size=$(du -sh "$DIST_DIR" | cut -f1)
    exe_size=$(du -h "$DIST_DIR/PDF-to-DOCX.exe" | cut -f1)
    echo ">> Built: $DIST_DIR/"
    echo "   PDF-to-DOCX.exe: $exe_size  |  total folder: $folder_size"
    echo "   Distribute the whole PDF-to-DOCX/ folder (zip it)."
else
    echo ">> ERROR: dist/PDF-to-DOCX/ was not produced — check the output above." >&2
    exit 1
fi
