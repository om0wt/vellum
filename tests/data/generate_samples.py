#!/usr/bin/env python3
"""Generate small public sample PDFs for tests/data/.

These fixtures are intentionally synthetic — placeholder content only —
so they can be committed to a public repo without revealing any real
documents. They cover the two main conversion code paths the project
supports:

* ``text_with_tables.pdf`` — text-based PDF with a heading, body
  paragraph, bullet list, and a small bordered table. Exercises the
  pdf2docx + lattice-table-detection path plus the first-table-header
  re-injection logic.
* ``scanned_text.pdf`` — image-only PDF (a PNG embedded as the page
  background). pdf2docx sees zero extractable words on this; it's the
  canonical example for the tesseract OCR fallback.

Run from anywhere with the project venv active:

    python tests/data/generate_samples.py

Both files are deterministic, well under 100 KB, and safe to regenerate
at any time. The committed copies in this directory are sufficient for
running the test suite — only re-run the generator if you've changed
the script and want fresh output.

Dependencies: PyMuPDF (already a runtime dep) and Pillow. Pillow is
**not** a runtime dep of the converter, so install it separately if
you want to regenerate the fixtures:

    pip install Pillow
"""
from __future__ import annotations

import io
from pathlib import Path

import fitz  # PyMuPDF
from PIL import Image, ImageDraw, ImageFont

OUT_DIR = Path(__file__).parent


def make_text_with_tables(path: Path) -> None:
    """Build a text-based A4 PDF with heading, paragraph, bullets, and a
    bordered table — all using PyMuPDF's text + draw primitives so the
    output has a real text layer (no OCR needed)."""
    doc = fitz.open()
    page = doc.new_page(width=595, height=842)  # A4 in points

    # Title (24pt) — exercises the heading-detection path.
    page.insert_text((50, 80), "Sample Document", fontsize=24, fontname="helv")

    # Subtitle / smaller heading.
    page.insert_text(
        (50, 115), "A synthetic fixture for the test suite",
        fontsize=14, fontname="helv",
    )

    # Body paragraph.
    body = (
        "This is a public sample PDF used by the project's tests. It "
        "contains a heading, this paragraph, a short bullet list, and "
        "a small bordered table below. No OCR is required to extract "
        "the text — the lattice-table converter handles it directly."
    )
    page.insert_textbox(
        fitz.Rect(50, 140, 545, 220),
        body,
        fontsize=11,
        fontname="helv",
    )

    # Bullet list.
    bullets = [
        "First bullet item",
        "Second bullet item with slightly longer text",
        "Third bullet item",
    ]
    by = 230
    for line in bullets:
        page.insert_text((60, by), "•", fontsize=11, fontname="helv")
        page.insert_text((75, by), line, fontsize=11, fontname="helv")
        by += 18

    # Bordered 4-column × 4-row table (1 header row + 3 data rows).
    table_top = 310
    cell_w = 120
    cell_h = 26
    headers = ["Row", "Col A", "Col B", "Col C"]
    rows = [
        ["1", "alpha", "beta", "gamma"],
        ["2", "delta", "epsilon", "zeta"],
        ["3", "eta", "theta", "iota"],
    ]

    # Header row — gray fill.
    for col, text in enumerate(headers):
        rect = fitz.Rect(
            50 + col * cell_w, table_top,
            50 + (col + 1) * cell_w, table_top + cell_h,
        )
        page.draw_rect(rect, fill=(0.85, 0.85, 0.85), color=(0, 0, 0), width=0.7)
        page.insert_text(
            (rect.x0 + 6, rect.y0 + 18),
            text,
            fontsize=11,
            fontname="hebo",  # helvetica-bold
        )

    # Data rows — outlined only.
    for row_i, row in enumerate(rows):
        y = table_top + (row_i + 1) * cell_h
        for col, text in enumerate(row):
            rect = fitz.Rect(
                50 + col * cell_w, y,
                50 + (col + 1) * cell_w, y + cell_h,
            )
            page.draw_rect(rect, color=(0, 0, 0), width=0.7)
            page.insert_text(
                (rect.x0 + 6, rect.y0 + 18),
                text,
                fontsize=11,
                fontname="helv",
            )

    doc.save(str(path), garbage=4, deflate=True)
    doc.close()


def make_scanned(path: Path) -> None:
    """Build an image-only PDF: render text into a PNG with Pillow,
    embed the PNG as the entire page. pdf2docx will see zero text on
    this; it's the canonical example for the OCR fallback."""
    # A4 at 150 DPI = ~1240 × 1754 px.
    img = Image.new("RGB", (1240, 1754), "white")
    draw = ImageDraw.Draw(img)

    # Try a real TrueType font for crisp glyphs; fall back to PIL's
    # bitmap default if no system font is available.
    font_paths = [
        "/System/Library/Fonts/Helvetica.ttc",          # macOS
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",  # Linux
        "C:\\Windows\\Fonts\\arial.ttf",                # Windows
    ]
    font = None
    for fp in font_paths:
        try:
            font = ImageFont.truetype(fp, 36)
            break
        except OSError:
            continue
    if font is None:
        font = ImageFont.load_default()

    lines = [
        "Scanned Sample Document",
        "",
        "This page exists only as an image — there is no",
        "extractable text layer in the PDF. Tools that rely",
        "on text extraction (like pdf2docx) will see zero",
        "words on this document.",
        "",
        "The OCR conversion path uses tesseract to recognize",
        "the rendered glyphs and emit text directly into a",
        "Word document via python-docx, with line-based",
        "merging and font-size-driven heading classification.",
    ]
    y = 120
    for line in lines:
        draw.text((100, y), line, fill="black", font=font)
        y += 60

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)

    doc = fitz.open()
    page = doc.new_page(width=595, height=842)
    page.insert_image(page.rect, stream=buf.read())
    doc.save(str(path), garbage=4, deflate=True)
    doc.close()


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    text_path = OUT_DIR / "text_with_tables.pdf"
    scan_path = OUT_DIR / "scanned_text.pdf"
    make_text_with_tables(text_path)
    make_scanned(scan_path)
    print(f"created {text_path}  ({text_path.stat().st_size} bytes)")
    print(f"created {scan_path}  ({scan_path.stat().st_size} bytes)")


if __name__ == "__main__":
    main()
