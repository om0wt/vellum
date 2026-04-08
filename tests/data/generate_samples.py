#!/usr/bin/env python3
"""Generate small public sample PDFs for tests/data/.

These fixtures are intentionally synthetic — placeholder content only —
so they can be committed to a public repo without revealing any real
documents. They cover the main conversion code paths the project
supports:

* ``text_with_tables.pdf`` — text-based PDF with a heading, body
  paragraph, bullet list, and a small bordered table. Exercises the
  pdf2docx + lattice-table-detection path plus the first-table-header
  re-injection logic.
* ``scanned_text.pdf`` — image-only PDF (a PNG embedded as the page
  background). pdf2docx sees zero extractable words on this; it's the
  canonical example for the tesseract OCR fallback.
* ``scanned-demo.pdf`` — a more realistic 3-page mixed-layout fixture
  (one-column intro + table, two-column body + figure, figure + table
  + closing) that has been rasterized at 200 DPI per page so the
  result has no extractable text layer. Used to exercise the OCR
  pipeline against a document that has multi-column layout, multiple
  tables, and embedded figures — closer to what a real flatbed
  scanner produces from a printed report.

Run from anywhere with the project venv active:

    python tests/data/generate_samples.py

All files are deterministic and safe to regenerate at any time. The
committed copies in this directory are sufficient for running the test
suite — only re-run the generator if you've changed the script and
want fresh output.

Dependencies: PyMuPDF (already a runtime dep) and Pillow. Pillow is
**not** a runtime dep of the converter, so install it separately if
you want to regenerate the fixtures:

    pip install Pillow

The ``scanned-demo.pdf`` generator additionally tries to download two
stock photos from picsum.photos (Lorem Picsum) so the figures look
like real photographs rather than abstract gradients. If the host is
offline, it transparently falls back to a generated placeholder so
the script still completes.
"""
from __future__ import annotations

import io
import random
import urllib.error
import urllib.request
from pathlib import Path

import fitz  # PyMuPDF
from PIL import Image, ImageDraw, ImageEnhance, ImageFilter, ImageFont

OUT_DIR = Path(__file__).parent

# A4 in PDF points (1 pt = 1/72 inch)
A4_W = 595
A4_H = 842
MARGIN = 50


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


# ---------------------------------------------------------------------------
# scanned-demo.pdf — multi-page mixed-layout fixture, scan-degraded
# ---------------------------------------------------------------------------


def _load_font(size: int) -> ImageFont.ImageFont:
    """Best-effort TrueType font loader, falls back to PIL's bitmap."""
    for fp in (
        "/System/Library/Fonts/Helvetica.ttc",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "C:\\Windows\\Fonts\\arial.ttf",
    ):
        try:
            return ImageFont.truetype(fp, size)
        except OSError:
            continue
    return ImageFont.load_default()


def _fetch_image_bytes(url: str, timeout: int = 10) -> bytes | None:
    """Try to fetch an image. Returns ``None`` on any failure (offline,
    DNS error, HTTP error, timeout) so the caller can fall back to a
    locally generated placeholder."""
    try:
        req = urllib.request.Request(
            url,
            headers={"User-Agent": "vellum-fixture-generator/1.0"},
        )
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return resp.read()
    except (urllib.error.URLError, TimeoutError, OSError):
        return None


def _placeholder_image(width: int, height: int, label: str, seed: int) -> bytes:
    """Generate a deterministic placeholder JPEG. Used when there is no
    network access at fixture-generation time. Looks like a faux photo
    so the resulting page still has obvious figure regions for the
    layout test, even when offline."""
    rng = random.Random(seed)
    # Build a soft diagonal gradient + a few large translucent disks for
    # visual weight.
    img = Image.new("RGB", (width, height), "#888888")
    px = img.load()
    for y in range(height):
        for x in range(width):
            t = (x + y) / (width + height)
            r = int(80 + 120 * t)
            g = int(110 + 90 * (1 - t))
            b = int(140 + 60 * t)
            px[x, y] = (r, g, b)
    draw = ImageDraw.Draw(img)
    for _ in range(5):
        cx = rng.randint(0, width)
        cy = rng.randint(0, height)
        r = rng.randint(width // 8, width // 4)
        draw.ellipse((cx - r, cy - r, cx + r, cy + r),
                     fill=(rng.randint(160, 240),
                           rng.randint(160, 240),
                           rng.randint(160, 240)))
    font = _load_font(36)
    bbox = draw.textbbox((0, 0), label, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    draw.text(((width - tw) / 2, (height - th) / 2),
              label, fill="white", font=font)
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=85)
    return buf.getvalue()


def _get_picture(width: int, height: int, seed: str, label: str) -> bytes:
    """Fetch a Lorem Picsum stock photo by seed (deterministic) so the
    resulting fixture has actual photographic content. Falls back to a
    locally generated placeholder when offline."""
    url = f"https://picsum.photos/seed/{seed}/{width}/{height}"
    data = _fetch_image_bytes(url)
    if data:
        return data
    return _placeholder_image(width, height, label, seed=hash(seed) & 0xFFFF)


def _simulate_scan(pil_img: Image.Image, *, seed: int = 0) -> bytes:
    """Apply degradations that mimic a flatbed scanner output:
    grayscale, slight blur, sensor noise, mild skew, lowered contrast,
    JPEG compression. Result is still legible to tesseract (the OCR
    pipeline this fixture is built to test) but defeats most
    client-side OCR overlays such as macOS Preview's Live Text and the
    OCR features in Adobe Acrobat Reader.

    Returns JPEG bytes ready to embed via ``page.insert_image``.
    """
    rng = random.Random(seed)

    # 1. Mild skew (real scans are never perfectly axis-aligned).
    angle = rng.uniform(-0.6, 0.6)
    img = pil_img.rotate(
        angle,
        resample=Image.BICUBIC,
        fillcolor=(255, 255, 255),
        expand=False,
    )

    # 2. Slight Gaussian blur (scanner optics + lens MTF rolloff).
    img = img.filter(ImageFilter.GaussianBlur(radius=0.45))

    # 3. Desaturate to grayscale (typical "B&W document" scan), then
    #    back to RGB so we can JPEG-encode without chroma surprises.
    img = img.convert("L").convert("RGB")

    # 4. Sensor noise — sprinkle a small number of dark/light pixels.
    px = img.load()
    w, h = img.size
    n_noise = (w * h) // 700
    for _ in range(n_noise):
        x = rng.randint(0, w - 1)
        y = rng.randint(0, h - 1)
        v = rng.randint(0, 255)
        px[x, y] = (v, v, v)

    # 5. Lower the contrast slightly — scanned pages are never pure
    #    black on pure white.
    img = ImageEnhance.Contrast(img).enhance(0.9)
    img = ImageEnhance.Brightness(img).enhance(0.97)

    # 6. JPEG encode at moderate quality (real scanners save jpeg).
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=72, optimize=True)
    return buf.getvalue()


def _draw_table(
    page: fitz.Page,
    *,
    top: float,
    cols: list[tuple[float, str, float]],
    rows: list[list[str]],
    cell_h: float = 24,
) -> float:
    """Draw a bordered table at ``top`` on ``page``. ``cols`` is a list
    of (x, header_label, width) tuples. Returns the y coordinate just
    below the last row so the caller can place follow-up content."""
    # Header row — gray fill.
    for x, label, w in cols:
        rect = fitz.Rect(x, top, x + w, top + cell_h)
        page.draw_rect(rect, fill=(0.85, 0.85, 0.85),
                       color=(0, 0, 0), width=0.7)
        page.insert_text((x + 6, top + 17), label,
                         fontsize=11, fontname="hebo")
    # Data rows — outlined only.
    for r_i, row in enumerate(rows):
        y = top + (r_i + 1) * cell_h
        for c_i, (x, _label, w) in enumerate(cols):
            rect = fitz.Rect(x, y, x + w, y + cell_h)
            page.draw_rect(rect, color=(0, 0, 0), width=0.7)
            page.insert_text((x + 6, y + 17), row[c_i],
                             fontsize=11, fontname="helv")
    return top + (len(rows) + 1) * cell_h


def make_scanned_demo(path: Path) -> None:
    """Build a 3-page mixed-layout fixture that exercises the OCR
    pipeline against a realistic document:

    * Page 1 — title, intro paragraph (one column), a 4-column table
    * Page 2 — section heading, two-column body text, a figure with caption
    * Page 3 — figure with caption, a status table, closing paragraph

    The text-based version is then rasterized at 200 DPI and each page
    is replaced with the rasterized image after going through
    ``_simulate_scan`` (skew, blur, noise, grayscale, JPEG). The final
    PDF has zero text layer (forensically verifiable) and the rendered
    pages look like real flatbed-scanner output.
    """
    src = fitz.open()

    # ---------- Page 1 — title + intro + table ----------
    p1 = src.new_page(width=A4_W, height=A4_H)
    p1.insert_text((MARGIN, 90), "Vellum Scanned Demo",
                   fontsize=24, fontname="hebo")
    p1.insert_text((MARGIN, 118),
                   "A synthetic mixed-layout fixture for OCR testing",
                   fontsize=12, fontname="helv", color=(0.35, 0.35, 0.35))

    intro = (
        "This document is a deliberately mixed-layout sample used to test "
        "the OCR conversion path of Vellum. It contains one-column body "
        "text, a multi-column section, two tables, and embedded figures. "
        "The document was first typeset as a normal text-based PDF; every "
        "page was then rasterized at 200 DPI and the resulting bitmap was "
        "re-embedded as the only content of each page after going through "
        "a scanner-style degradation pipeline (slight skew, blur, sensor "
        "noise, grayscale, JPEG compression). The end result has no "
        "extractable text layer at all, which mirrors what a flatbed "
        "scanner would produce when scanning a printed copy."
    )
    p1.insert_textbox(
        fitz.Rect(MARGIN, 145, A4_W - MARGIN, 290),
        intro, fontsize=11, fontname="helv",
    )

    # Table 1 — invoice-style line items.
    table1_cols = [
        (MARGIN,        "Item",        180),
        (MARGIN + 180,  "Quantity",     90),
        (MARGIN + 270,  "Unit price",  100),
        (MARGIN + 370,  "Subtotal",    125),
    ]
    table1_rows = [
        ["Notebook, A5",     "12", "  3.50",  " 42.00"],
        ["Pen, blue",        "48", "  0.80",  " 38.40"],
        ["Stapler",          " 4", " 12.00",  " 48.00"],
        ["Paper, A4 ream",   " 6", "  4.25",  " 25.50"],
    ]
    _draw_table(p1, top=310, cols=table1_cols, rows=table1_rows)

    # ---------- Page 2 — two-column body + figure ----------
    p2 = src.new_page(width=A4_W, height=A4_H)
    p2.insert_text((MARGIN, 90), "Multi-column section",
                   fontsize=18, fontname="hebo")

    col_top = 120
    col_h = 360
    gap = 18
    col_w = (A4_W - 2 * MARGIN - gap) / 2
    left_rect = fitz.Rect(MARGIN, col_top,
                          MARGIN + col_w, col_top + col_h)
    right_rect = fitz.Rect(MARGIN + col_w + gap, col_top,
                           A4_W - MARGIN, col_top + col_h)

    left_text = (
        "Lorem ipsum dolor sit amet, consectetur adipiscing elit. "
        "Sed do eiusmod tempor incididunt ut labore et dolore magna "
        "aliqua. Ut enim ad minim veniam, quis nostrud exercitation "
        "ullamco laboris nisi ut aliquip ex ea commodo consequat. "
        "Duis aute irure dolor in reprehenderit in voluptate velit "
        "esse cillum dolore eu fugiat nulla pariatur.\n\n"
        "Excepteur sint occaecat cupidatat non proident, sunt in "
        "culpa qui officia deserunt mollit anim id est laborum. "
        "Curabitur pretium tincidunt lacus, at hendrerit nulla "
        "feugiat eu. Vivamus ac leo pretium, faucibus dui in, "
        "elementum nisl."
    )
    right_text = (
        "Sed ut perspiciatis unde omnis iste natus error sit "
        "voluptatem accusantium doloremque laudantium, totam rem "
        "aperiam, eaque ipsa quae ab illo inventore veritatis et "
        "quasi architecto beatae vitae dicta sunt explicabo.\n\n"
        "Nemo enim ipsam voluptatem quia voluptas sit aspernatur "
        "aut odit aut fugit, sed quia consequuntur magni dolores "
        "eos qui ratione voluptatem sequi nesciunt. Neque porro "
        "quisquam est qui dolorem ipsum quia dolor sit amet, "
        "consectetur, adipisci velit."
    )
    p2.insert_textbox(left_rect, left_text, fontsize=10, fontname="helv")
    p2.insert_textbox(right_rect, right_text, fontsize=10, fontname="helv")

    # Figure 1 — placed below the columns.
    pic1 = _get_picture(800, 450, "vellum-pic-1", "Figure 1")
    fig1_rect = fitz.Rect(MARGIN, col_top + col_h + 20,
                          A4_W - MARGIN, col_top + col_h + 220)
    p2.insert_image(fig1_rect, stream=pic1)
    p2.insert_text(
        (MARGIN, col_top + col_h + 235),
        "Figure 1 — A scenic placeholder image fetched from picsum.photos.",
        fontsize=9, fontname="helv", color=(0.3, 0.3, 0.3),
    )

    # ---------- Page 3 — figure + table + closing ----------
    p3 = src.new_page(width=A4_W, height=A4_H)
    p3.insert_text((MARGIN, 90), "Figures and summary",
                   fontsize=18, fontname="hebo")

    pic2 = _get_picture(800, 450, "vellum-pic-2", "Figure 2")
    fig2_rect = fitz.Rect(MARGIN, 115, A4_W - MARGIN, 320)
    p3.insert_image(fig2_rect, stream=pic2)
    p3.insert_text(
        (MARGIN, 335),
        "Figure 2 — Another scenic placeholder image fetched from picsum.photos.",
        fontsize=9, fontname="helv", color=(0.3, 0.3, 0.3),
    )

    # Table 2 — section status.
    table2_cols = [
        (MARGIN,         "Section",          200),
        (MARGIN + 200,   "Status",           120),
        (MARGIN + 320,   "Last updated",     175),
    ]
    table2_rows = [
        ["Introduction",     "Done",        "2026-04-08"],
        ["Two-column body",  "Done",        "2026-04-08"],
        ["Figures",          "Done",        "2026-04-08"],
        ["Summary",          "In progress", "2026-04-08"],
    ]
    bottom_y = _draw_table(p3, top=365, cols=table2_cols, rows=table2_rows)

    closing = (
        "This concludes the Vellum scanned demo fixture. The original PDF "
        "(before scanning) had three pages with mixed layouts, two tables, "
        "two figures, and both single- and multi-column body text. After "
        "the rasterization and scanner-degradation step that produces the "
        "final PDF, every text glyph has been baked into a slightly noisy "
        "bitmap and the only way to recover the content is via OCR."
    )
    p3.insert_textbox(
        fitz.Rect(MARGIN, bottom_y + 20, A4_W - MARGIN, A4_H - MARGIN),
        closing, fontsize=11, fontname="helv",
    )

    # ---------- Rasterize every page at 200 DPI and re-embed ----------
    out = fitz.open()
    DPI = 200
    matrix = fitz.Matrix(DPI / 72, DPI / 72)
    for i, src_page in enumerate(src):
        pix = src_page.get_pixmap(matrix=matrix, alpha=False)
        # Convert pixmap → PIL via PNG round-trip (bulletproof, no
        # stride/alpha edge cases) → scan-degrade → JPEG bytes.
        pil_img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
        scanned_jpeg = _simulate_scan(pil_img, seed=42 + i)
        out_page = out.new_page(width=A4_W, height=A4_H)
        out_page.insert_image(out_page.rect, stream=scanned_jpeg)

    out.save(str(path), garbage=4, deflate=True)
    out.close()
    src.close()


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    text_path = OUT_DIR / "text_with_tables.pdf"
    scan_path = OUT_DIR / "scanned_text.pdf"
    demo_path = OUT_DIR / "scanned-demo.pdf"
    make_text_with_tables(text_path)
    make_scanned(scan_path)
    make_scanned_demo(demo_path)
    print(f"created {text_path}  ({text_path.stat().st_size} bytes)")
    print(f"created {scan_path}  ({scan_path.stat().st_size} bytes)")
    print(f"created {demo_path}  ({demo_path.stat().st_size} bytes)")


if __name__ == "__main__":
    main()
