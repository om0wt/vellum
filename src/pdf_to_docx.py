#!/usr/bin/env python3
"""
Convert a PDF to an editable DOCX while preserving formatting.

Uses pdf2docx for layout/table extraction, then post-processes the result:
  * Replaces the legacy "Symbol" font (which pdf2docx assigns to bullet glyphs)
    with the actual font used in the PDF (e.g. "SymbolMT") so that bullets
    render correctly in Word instead of as missing-glyph squares.
  * Re-injects the header row of the first table on page 1. pdf2docx
    sometimes drops a table's header row when its column structure differs
    from the data rows below it (e.g. an empty leftmost cell + merged cells).
    The header is re-extracted directly from the PDF via PyMuPDF and
    reinserted with matching shading + bold text.

Usage:
    python pdf_to_docx.py input.pdf [output.docx]

If output is omitted, the .docx file is written next to the PDF with the same
basename. The script is safe to re-run; it overwrites the output.

Tip: pdf2docx's "stream table" heuristic sometimes invents tables out of
aligned text (e.g. label/value blocks). Disable it via --no-stream-tables when
your PDF has key-value lists that should stay as paragraphs.
"""

from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
import tempfile
from collections import Counter
from pathlib import Path
from typing import Callable, Optional

import fitz  # PyMuPDF, already a dependency of pdf2docx
from pdf2docx import Converter
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

# Fonts whose names cause Word to render U+2022 (bullet) as a missing-glyph
# square. The legacy "Symbol" PostScript font remaps Unicode code points, so
# the actual bullet character ends up unmapped. Replace with the OpenType
# variant that keeps Unicode mappings intact.
BAD_BULLET_FONTS = {"Symbol"}
REPLACEMENT_BULLET_FONT = "SymbolMT"


# ----- OCR support ---------------------------------------------------------
#
# pdf2docx can't process scanned PDFs at all — its layout analyzer relies on
# real font glyphs with measurable widths, and skips image-only pages with
# the warning "Words count: 0. It might be a scanned pdf, which is not
# supported yet". So for scanned PDFs we bypass pdf2docx entirely: render
# each page, run tesseract directly to get word-level text + positions
# (TSV output), and build a DOCX from that text via python-docx.
#
# Tesseract is NOT bundled — the user must install it themselves (e.g.
# `brew install tesseract tesseract-lang` on macOS, the UB Mannheim
# installer on Windows).


def list_tesseract_languages() -> Optional[list[str]]:
    """Return a sorted list of language codes installed in the local
    tesseract, or ``None`` if tesseract is not on PATH (or otherwise fails
    to respond).

    The output of ``tesseract --list-langs`` looks like::

        List of available languages (3):
        eng
        osd
        slk

    ``osd`` is the orientation/script detection model, not a real language,
    so it is filtered out.
    """
    if shutil.which("tesseract") is None:
        return None
    try:
        result = subprocess.run(
            ["tesseract", "--list-langs"],
            capture_output=True,
            text=True,
            timeout=5,
        )
    except (subprocess.TimeoutExpired, OSError):
        return None

    # Different tesseract versions print to stdout vs stderr; combine both.
    text = (result.stdout or "") + "\n" + (result.stderr or "")
    langs: set[str] = set()
    for line in text.splitlines():
        line = line.strip()
        if not line or ":" in line or " " in line:
            continue
        if all(c.isalnum() or c == "_" for c in line):
            langs.add(line)
    langs.discard("osd")
    return sorted(langs) if langs else None


ProgressCallback = Callable[[int, int], None]


def _parse_tesseract_tsv(tsv_path: Path, dpi: int = 300) -> list[dict]:
    """Parse tesseract's TSV output into an ordered list of paragraph dicts.

    Returns a list of dicts, one per visually-coherent paragraph in the
    OCR output, each with ``text``, ``font_size_pt``, ``alignment``
    ("left" or "center"), and ``space_before_pt`` so the caller can
    render visual hierarchy *and* whitespace in the DOCX.

    Layout reconstruction works in two phases:

    1. **Build per-line records** from level-5 (word) rows, computing the
       median word height per line (≈ font size at this DPI).

    2. **Merge consecutive lines into paragraphs** only when all of these
       hold:
         - Same tesseract block (so columns don't bleed across).
         - Similar line height (height ratio between 0.75 and 1.33) — so
           a heading line and a body line below it stay as separate
           paragraphs even if tesseract grouped them in the same block.
         - Small vertical gap (< ~0.7 line heights between bottom-of-prev
           and top-of-next) — so visually separated sections don't merge.

       When the gap to the previous paragraph is large (≥ 1× line
       height), we record ``space_before_pt`` so the renderer can add
       extra spacing in the DOCX.

    This is a major change from grouping by ``(block_num, par_num)``,
    which collapsed adjacent headings + body into one mushy paragraph.
    The line-then-merge approach preserves both heading hierarchy and
    whitespace between sections.
    """
    import csv

    # (page, block, par, line) → list of word dicts
    words_by_line: dict[tuple[int, int, int, int], list[dict]] = {}
    page_widths: dict[int, int] = {}

    with open(tsv_path, encoding="utf-8") as f:
        reader = csv.reader(f, delimiter="\t", quoting=csv.QUOTE_NONE)
        header = next(reader, None)
        if header is None:
            return []
        col = {name: i for i, name in enumerate(header)}
        required = {
            "level", "page_num", "block_num", "par_num", "line_num",
            "word_num", "left", "top", "width", "height", "text",
        }
        if not required.issubset(col):
            return []

        for row in reader:
            if len(row) < len(header):
                continue
            try:
                level = int(row[col["level"]])
            except ValueError:
                continue

            if level == 1:
                try:
                    page_widths[int(row[col["page_num"]])] = int(row[col["width"]])
                except ValueError:
                    pass
                continue

            if level != 5:  # words only
                continue

            try:
                page_num = int(row[col["page_num"]])
                block = int(row[col["block_num"]])
                par = int(row[col["par_num"]])
                line = int(row[col["line_num"]])
                word_num = int(row[col["word_num"]])
                left = int(row[col["left"]])
                top = int(row[col["top"]])
                width = int(row[col["width"]])
                height = int(row[col["height"]])
            except ValueError:
                continue

            text = row[col["text"]]
            if not text or not text.strip():
                continue

            words_by_line.setdefault((page_num, block, par, line), []).append({
                "word_num": word_num,
                "left": left,
                "top": top,
                "width": width,
                "height": height,
                "text": text,
            })

    page_width = next(iter(page_widths.values()), 0)

    # ---- Phase 1: build line records, sorted in reading order ----
    line_records: list[dict] = []
    for key in sorted(words_by_line.keys()):
        page_num, block, par, line_num = key
        words = sorted(words_by_line[key], key=lambda w: w["word_num"])
        if not words:
            continue
        text = " ".join(w["text"] for w in words)
        heights = sorted(w["height"] for w in words)
        median_h = heights[len(heights) // 2]
        line_records.append({
            "block": block,
            "par": par,
            "line_num": line_num,
            "text": text,
            "top": min(w["top"] for w in words),
            "bottom": max(w["top"] + w["height"] for w in words),
            "left": min(w["left"] for w in words),
            "right": max(w["left"] + w["width"] for w in words),
            "height": median_h,
        })

    # ---- Phase 2: greedy merge consecutive lines into paragraphs ----
    paragraphs: list[dict] = []
    current: list[dict] | None = None
    prev_para_bottom: int | None = None

    def _finalize(group: list[dict], gap_to_prev: int | None) -> dict:
        text = " ".join(ln["text"] for ln in group)
        heights = sorted(ln["height"] for ln in group)
        median_h = heights[len(heights) // 2]
        font_size_pt = max(6.0, min(round(median_h * 72.0 / dpi, 1), 72.0))

        alignment = "left"
        if page_width > 0:
            min_left = min(ln["left"] for ln in group)
            max_right = max(ln["right"] for ln in group)
            left_margin = min_left
            right_margin = page_width - max_right
            if left_margin > page_width * 0.10:
                imbalance = abs(left_margin - right_margin) / page_width
                if imbalance < 0.05:
                    alignment = "center"

        # Convert vertical gap to extra "space before" in points. We only
        # add explicit spacing when the gap is bigger than ~1 line height
        # (otherwise normal paragraph spacing already covers it).
        space_before_pt = 0.0
        if gap_to_prev is not None and gap_to_prev > median_h * 1.0:
            # Cap so a giant whitespace block doesn't blow out the page.
            extra_px = min(gap_to_prev - int(median_h), int(median_h * 4))
            space_before_pt = round(extra_px * 72.0 / dpi, 1)

        return {
            "text": text,
            "font_size_pt": font_size_pt,
            "alignment": alignment,
            "space_before_pt": space_before_pt,
        }

    for line in line_records:
        if current is None:
            current = [line]
            continue

        prev = current[-1]
        same_block = line["block"] == prev["block"]
        # Heights: ratio close to 1 means similar text size.
        ratio = line["height"] / max(prev["height"], 1)
        similar_height = 0.75 <= ratio <= 1.33
        gap = line["top"] - prev["bottom"]
        small_gap = gap < line["height"] * 0.7

        if same_block and similar_height and small_gap:
            current.append(line)
        else:
            paragraphs.append(_finalize(
                current,
                gap_to_prev=(current[0]["top"] - prev_para_bottom)
                            if prev_para_bottom is not None else None,
            ))
            prev_para_bottom = current[-1]["bottom"]
            current = [line]

    if current is not None:
        paragraphs.append(_finalize(
            current,
            gap_to_prev=(current[0]["top"] - prev_para_bottom)
                        if prev_para_bottom is not None else None,
        ))

    # ---- Phase 3: detect code lines and merge consecutive code blocks ----
    paragraphs = _merge_code_paragraphs(paragraphs)

    return paragraphs


def _looks_like_code(text: str) -> bool:
    """Heuristic: does this line look like code, CLI invocation, or JSON?

    Score-based detector keyed on markers that almost never appear in
    natural prose:

    * ``://`` (URL — strong signal)
    * `` -- `` (CLI flag separator)
    * leading/trailing ``{`` or ``}`` (JSON brace on its own line)
    * ``": "`` or ``":"`` (JSON key/value pair)
    * 4+ double-quote characters (densely quoted strings)
    * ratio of structural punctuation ``{}[]()<>=;:|&"'`` > 10%

    Threshold of score ≥ 2 catches CLI commands, single-line JSON, and
    multi-line JSON fragments while leaving body text alone (which
    typically scores 0 or 1 from incidental punctuation).
    """
    if not text:
        return False

    score = 0
    if "://" in text:
        score += 2
    if " -- " in text:
        score += 2
    stripped = text.strip()
    if stripped.startswith("{") or stripped.endswith("{"):
        score += 2
    if stripped.startswith("}") or stripped.endswith("}"):
        score += 2
    if '": "' in text or '":"' in text:
        score += 2
    if text.count('"') >= 4:
        score += 1

    structural = sum(1 for c in text if c in '{}[]()<>=;:|&"\'`')
    if len(text) > 0 and structural / len(text) > 0.10:
        score += 1

    return score >= 2


def _merge_code_paragraphs(paragraphs: list[dict]) -> list[dict]:
    """Mark code-style paragraphs and merge runs of consecutive code
    paragraphs into a single multi-line block.

    Tesseract often splits a JSON code block across several
    "paragraphs" because braces on their own lines have very small
    measured heights compared to the lines with text content, and the
    line-merge step rejects them as too dissimilar. This pass undoes
    that fragmentation by re-merging anything the code detector
    classifies as code.

    Each merged code block:
    * Has ``is_code = True`` so the renderer can switch to monospace.
    * Has ``text`` joined with ``\\n`` so the renderer can preserve line
      breaks via ``run.add_break()``.
    * Inherits ``space_before_pt`` from the first paragraph in the run.
    * Uses the median font size of its constituent lines (smaller and
      more representative than the max).
    """
    for p in paragraphs:
        p["is_code"] = _looks_like_code(p.get("text", ""))

    merged: list[dict] = []
    i = 0
    while i < len(paragraphs):
        p = paragraphs[i]
        if not p.get("is_code"):
            merged.append(p)
            i += 1
            continue

        run_start = i
        while i < len(paragraphs) and paragraphs[i].get("is_code"):
            i += 1
        run_end = i

        block = paragraphs[run_start:run_end]
        if len(block) == 1:
            merged.append(block[0])
            continue

        joined_text = "\n".join(b["text"] for b in block)
        sizes = sorted(
            b["font_size_pt"] for b in block if b.get("font_size_pt")
        )
        font_size = sizes[len(sizes) // 2] if sizes else 0
        merged.append({
            "text": joined_text,
            "font_size_pt": font_size,
            "alignment": "left",
            "space_before_pt": block[0].get("space_before_pt", 0),
            "is_code": True,
        })

    return merged


def _classify_headings(paragraphs: list[dict]) -> None:
    """Tag each paragraph dict with a ``heading_level`` (0..6).

    Heuristic: the body font size is the **statistical mode** (most
    common) across all paragraphs. Anything significantly larger than
    that — by ≥ 30% — is treated as a heading. Heading sizes are then
    clustered (within 1.5pt of each other) and assigned levels in
    descending size order: largest cluster → Heading 1, next → Heading 2,
    capped at H6.

    The function mutates the input list in place. Page-break markers and
    blank-page placeholders are skipped.

    Why heading levels matter: applying Word's built-in ``Heading 1``…
    ``Heading 6`` styles isn't just visual — it tells Word the paragraph
    is *semantically* a heading, which is what populates the navigation
    pane, drives table-of-contents generation, and exposes structure to
    accessibility tools. Direct font-size formatting alone produces a
    document that *looks* structured but is one flat ``Normal`` blob in
    Word's eyes.
    """
    sizes = [
        round(p["font_size_pt"], 1)
        for p in paragraphs
        if p.get("font_size_pt") and not p.get("_page_break")
    ]
    if not sizes:
        return

    # Body size = most common (rounded to 0.5pt for stable mode detection)
    rounded = [round(s * 2) / 2 for s in sizes]
    body_size = Counter(rounded).most_common(1)[0][0]
    threshold = body_size * 1.3

    # Cluster heading-candidate sizes by ≤ 1.5pt proximity, descending.
    distinct = sorted({s for s in sizes if s > threshold}, reverse=True)
    clusters: list[list[float]] = []
    for s in distinct:
        if clusters and clusters[-1][-1] - s <= 1.5:
            clusters[-1].append(s)
        else:
            clusters.append([s])

    size_to_level: dict[float, int] = {}
    for level, cluster in enumerate(clusters[:6], start=1):
        for s in cluster:
            size_to_level[s] = level

    for p in paragraphs:
        if p.get("_page_break") or not p.get("font_size_pt"):
            continue
        s = round(p["font_size_pt"], 1)
        p["heading_level"] = size_to_level.get(s, 0)


def _promote_section_headings(paragraphs: list[dict], gap_threshold_pt: float = 20.0) -> None:
    """Promote any heading paragraph with a large vertical gap above it to
    Heading 1.

    Tesseract's word-height measurements are noisy enough that two
    visually equivalent section headings can end up classified as
    different heading levels (e.g. 23.5pt → H2 and 18pt → H3 even though
    both are top-level sections in the source). The structural signal
    that distinguishes a top-level section heading from a sub-heading is
    *whitespace*: top-level sections have a large gap before them, while
    sub-headings sit close under their parent. We use the
    ``space_before_pt`` already computed by ``_parse_tesseract_tsv`` as
    that signal: any heading with > ``gap_threshold_pt`` of whitespace
    above it is promoted to Heading 1.

    Mutates ``paragraphs`` in place.
    """
    for p in paragraphs:
        if p.get("_page_break"):
            continue
        if not p.get("heading_level"):
            continue
        sb = p.get("space_before_pt", 0) or 0
        if sb > gap_threshold_pt:
            p["heading_level"] = 1


def ocr_to_docx(
    input_pdf: Path,
    output_docx: Path,
    language: str = "eng",
    dpi: int = 300,
    progress_callback: Optional[ProgressCallback] = None,
) -> None:
    """Build a DOCX directly from a scanned PDF via tesseract OCR.

    Renders each page to PNG via PyMuPDF, runs
    ``tesseract page.png stem -l LANG tsv`` to get word-level text with
    layout grouping, parses the TSV into paragraphs, classifies heading
    levels globally based on font sizes, and writes the result to a new
    DOCX via python-docx with Word's built-in ``Heading 1..6`` styles
    applied where appropriate. Pages are separated by hard page breaks.

    Bypasses pdf2docx entirely — pdf2docx can't process scanned PDFs
    (its layout analyzer requires real glyph metrics, which OCR text
    doesn't have). Layout is preserved at the *paragraph* level: column
    structure and tables are flattened to a linear flow, but heading
    hierarchy and section spacing are recovered from word geometry.

    Parameters
    ----------
    input_pdf, output_docx
        Source PDF and destination DOCX.
    language
        Tesseract language code (e.g. ``"eng"``, ``"slk"``, or
        ``"slk+eng"`` for mixed-language documents).
    dpi
        Render resolution; 300 is the tesseract sweet spot for printed
        text. Raise to 400 for fine print, lower to 200 for speed.
    progress_callback
        Optional ``fn(current_page, total_pages)`` invoked once per page
        before OCR is performed.

    Raises
    ------
    RuntimeError
        If tesseract is not installed.
    subprocess.CalledProcessError
        If a tesseract invocation fails (e.g. unknown language code).
    """
    if shutil.which("tesseract") is None:
        raise RuntimeError(
            "tesseract is not installed or not on PATH. "
            "Install it locally first (e.g. `brew install tesseract tesseract-lang`)."
        )

    src = fitz.open(str(input_pdf))
    doc = Document()
    try:
        total = len(src)

        # ---- Pass 1: render + OCR + parse, accumulate across pages ----
        # Heading classification needs the whole document's font-size
        # distribution, so we collect everything first and render second.
        all_paragraphs: list[dict] = []

        with tempfile.TemporaryDirectory(prefix="pdf2docx-ocr-") as tmp:
            tmp_dir = Path(tmp)
            for i, page in enumerate(src):
                if progress_callback is not None:
                    progress_callback(i + 1, total)

                pix = page.get_pixmap(dpi=dpi)
                png_path = tmp_dir / f"page-{i:04d}.png"
                pix.save(str(png_path))

                stem = tmp_dir / f"page-{i:04d}"
                subprocess.run(
                    [
                        "tesseract",
                        str(png_path),
                        str(stem),
                        "-l",
                        language,
                        "tsv",
                    ],
                    check=True,
                    capture_output=True,
                )

                tsv_path = stem.with_suffix(".tsv")
                page_paragraphs = _parse_tesseract_tsv(tsv_path, dpi=dpi)

                if i > 0:
                    all_paragraphs.append({"_page_break": True})

                if page_paragraphs:
                    all_paragraphs.extend(page_paragraphs)
                else:
                    all_paragraphs.append({"_blank_page": True})

        # ---- Classify heading levels globally ----
        _classify_headings(all_paragraphs)
        # Promote any heading with a large gap above it to H1 — top-level
        # section headings have visible whitespace above them in the
        # source, so this overrides OCR's noisy size measurements.
        _promote_section_headings(all_paragraphs)

        # ---- Pass 2: render to DOCX ----
        for p in all_paragraphs:
            if p.get("_page_break"):
                doc.add_page_break()
                continue
            if p.get("_blank_page"):
                doc.add_paragraph()
                continue
            if not p.get("text", "").strip():
                continue

            para = doc.add_paragraph()
            level = p.get("heading_level", 0)
            if level > 0:
                # Word's built-in heading styles are named "Heading 1"..6.
                # python-docx looks them up by name in doc.styles.
                para.style = doc.styles[f"Heading {level}"]
            if p.get("alignment") == "center":
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if p.get("space_before_pt"):
                para.paragraph_format.space_before = Pt(p["space_before_pt"])

            is_code = p.get("is_code", False)
            # Split on '\n' for multi-line code blocks; non-code paragraphs
            # have no '\n' so this becomes a single-line loop.
            lines = p["text"].split("\n")
            prev_run = None
            for line_idx, line_text in enumerate(lines):
                if line_idx > 0 and prev_run is not None:
                    # Soft line break within the same paragraph (Shift+Enter
                    # in Word), preserves the paragraph's style/alignment
                    # but starts a new visual line.
                    prev_run.add_break()
                run = para.add_run(line_text)
                if is_code:
                    # Courier New is universally available on Windows,
                    # macOS, and Linux LibreOffice — safest cross-platform
                    # monospace.
                    run.font.name = "Courier New"
                # Direct font-size override on top of the heading style:
                # keeps the visual size from the source PDF instead of the
                # style's default. The heading style still gives Word the
                # semantic tag for navigation/TOC/accessibility.
                if p.get("font_size_pt"):
                    run.font.size = Pt(p["font_size_pt"])
                prev_run = run

        doc.save(str(output_docx))
    finally:
        src.close()


def convert_pdf(pdf_path: Path, docx_path: Path, parse_stream_table: bool = True) -> None:
    """Run pdf2docx with sensible defaults."""
    cv = Converter(str(pdf_path))
    try:
        cv.convert(
            str(docx_path),
            start=0,
            end=None,
            # pdf2docx kwargs (see pdf2docx.common.Settings):
            parse_stream_table=parse_stream_table,
        )
    finally:
        cv.close()


def _fix_run_font(run_element) -> bool:
    """If the run uses a 'bad' bullet font, swap it for the OpenType variant.

    Returns True if any change was made.
    """
    rpr = run_element.find(qn("w:rPr"))
    if rpr is None:
        return False
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        return False
    changed = False
    for attr in ("ascii", "hAnsi", "cs", "eastAsia"):
        key = qn(f"w:{attr}")
        val = rfonts.get(key)
        if val in BAD_BULLET_FONTS:
            rfonts.set(key, REPLACEMENT_BULLET_FONT)
            changed = True
    return changed


def fix_bullet_fonts(docx_path: Path) -> int:
    """Walk every run in the document and fix Symbol → SymbolMT.

    Returns the number of runs that were patched.
    """
    doc = Document(str(docx_path))
    patched = 0

    def walk_paragraphs(paragraphs):
        nonlocal patched
        for p in paragraphs:
            for run in p.runs:
                if _fix_run_font(run._element):
                    patched += 1

    walk_paragraphs(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                walk_paragraphs(cell.paragraphs)
                # nested tables
                for nested in cell.tables:
                    for nrow in nested.rows:
                        for ncell in nrow.cells:
                            walk_paragraphs(ncell.paragraphs)

    if patched:
        doc.save(str(docx_path))
    return patched


def _reconstruct_first_table_header(pdf_path: Path) -> tuple[list[str] | None, str | None]:
    """Re-extract the header row of page 1's first table directly from the PDF.

    pdf2docx occasionally drops a table's header row when the header's cell
    structure differs from the data rows (e.g. an empty leftmost cell or
    merged cells below). PyMuPDF's table detector still sees the full table,
    so we use it to recover the header text per logical column plus the
    header background fill color.

    Returns ``(header_texts, fill_hex)`` where ``header_texts`` has one entry
    per logical column. Returns ``(None, None)`` when no usable header band
    can be found.
    """
    doc = fitz.open(str(pdf_path))
    try:
        page = doc[0]
        tabs = page.find_tables()
        if not tabs.tables:
            return None, None
        table = tabs.tables[0]
        table_left, table_top, table_right, _ = table.bbox

        # First data row = the first full-width row that lies below table_top.
        # PyMuPDF over-segments columns when header cells span multiple data
        # columns, but the data row's non-None cells give us the canonical
        # logical column boundaries.
        data_row = None
        for row in table.rows:
            rl, rt, rr, _ = row.bbox
            if (abs(rl - table_left) < 1
                    and abs(rr - table_right) < 1
                    and rt > table_top + 1):
                data_row = row
                break
        if data_row is None:
            return None, None

        # Collect ALL x-edges (both lefts and rights) from non-None data
        # cells, then sort and merge boundaries that are too close
        # together to be real columns.
        #
        # Why this matters: PyMuPDF's `find_tables()` sometimes represents
        # column-divider lines as their own narrow "cells" sitting between
        # the real cells (e.g. 5pt-wide phantom cells between 110pt-wide
        # real columns). The naive "boundary at every cell edge" approach
        # then over-counts columns and the header reconstruction returns
        # more columns than the real table has, which makes the caller
        # refuse to inject the header due to the column-count mismatch.
        #
        # Adaptive tolerance: any pair of boundaries closer than
        # ``max(5pt, 2% of table width)`` is treated as one. 2% of table
        # width scales naturally with the document — for a 700pt-wide
        # table that's 14pt; for a 200pt-wide table that's still at
        # least the 5pt floor. This is below any practical text column
        # width (a 14pt-wide column couldn't fit a single character of
        # 11pt text) and well above the typical phantom-cell width.
        all_xs: list[float] = []
        for cell in data_row.cells:
            if cell is None:
                continue
            all_xs.append(cell[0])
            all_xs.append(cell[2])
        all_xs.sort()

        table_width = table_right - table_left
        merge_tolerance = max(5.0, 0.02 * table_width)

        col_xs: list[float] = []
        for x in all_xs:
            if not col_xs or x - col_xs[-1] > merge_tolerance:
                col_xs.append(x)
        if len(col_xs) < 2:
            return None, None

        header_y_top = table_top
        header_y_bottom = data_row.bbox[1]
        if header_y_bottom - header_y_top < 1:
            return None, None  # no header band above the first data row

        header_texts: list[str] = []
        for i in range(len(col_xs) - 1):
            bbox = (col_xs[i], header_y_top, col_xs[i + 1], header_y_bottom)
            text = page.get_text("text", clip=bbox).strip()
            header_texts.append(" ".join(text.split()))

        if not any(header_texts):
            return None, None  # nothing to inject

        # Sample a filled rectangle inside the header band for the fill color.
        fill_hex: str | None = None
        for d in page.get_drawings():
            fill = d.get("fill")
            rect = d.get("rect")
            if fill is None or rect is None:
                continue
            cx = (rect[0] + rect[2]) / 2
            cy = (rect[1] + rect[3]) / 2
            if (table_left - 1 <= cx <= table_right + 1
                    and header_y_top - 1 <= cy <= header_y_bottom + 1):
                r, g, b = fill[:3]
                fill_hex = (
                    f"{int(round(r * 255)):02X}"
                    f"{int(round(g * 255)):02X}"
                    f"{int(round(b * 255)):02X}"
                )
                break

        return header_texts, fill_hex
    finally:
        doc.close()


def _set_cell_shading(cell, hex_color: str) -> None:
    """Apply a solid background fill to a python-docx cell."""
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:shd")):
        tcPr.remove(old)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def _mark_as_header_row(tr_element) -> None:
    """Set <w:tblHeader/> on a row so Word treats it as a repeating header."""
    trPr = tr_element.find(qn("w:trPr"))
    if trPr is None:
        trPr = OxmlElement("w:trPr")
        tr_element.insert(0, trPr)
    if trPr.find(qn("w:tblHeader")) is None:
        trPr.append(OxmlElement("w:tblHeader"))


def fix_first_table_header(pdf_path: Path, docx_path: Path) -> bool:
    """Re-inject a missing header row into the first table of the DOCX.

    Detection is conservative: the inserted row is only added when the
    reconstructed header column count matches the existing table's column
    count, and only when the table's current first row is not already that
    header (so re-running the script is safe).

    Also removes the orphaned paragraphs immediately above the table that
    pdf2docx left behind containing the header text fragments.
    """
    header_texts, fill_hex = _reconstruct_first_table_header(pdf_path)
    if header_texts is None:
        return False

    doc = Document(str(docx_path))
    if not doc.tables:
        return False
    table = doc.tables[0]

    if len(table.columns) != len(header_texts):
        return False  # column mismatch — refuse to mangle the table

    def _norm(s: str) -> str:
        return " ".join((s or "").split()).lower()

    existing_first = [_norm(c.text) for c in table.rows[0].cells]
    if existing_first == [_norm(h) for h in header_texts]:
        return False  # already fixed; nothing to do

    # Build the new row by appending then moving its <w:tr> to the top.
    new_row = table.add_row()
    for cell, text in zip(new_row.cells, header_texts):
        cell.text = text
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        for para in cell.paragraphs:
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in para.runs:
                run.bold = True
        if fill_hex:
            _set_cell_shading(cell, fill_hex)

    tbl_el = table._tbl
    tr_el = new_row._tr
    tbl_el.remove(tr_el)
    first_existing_tr = tbl_el.find(qn("w:tr"))
    first_existing_tr.addprevious(tr_el)
    _mark_as_header_row(tr_el)

    # Remove orphaned paragraphs above the table that contain header text.
    # We walk backwards from the table and only delete paragraphs whose
    # tokens are a strict subset of the reconstructed header tokens — this
    # avoids touching unrelated content above.
    body = doc.element.body
    children = list(body.iterchildren())
    tbl_index = children.index(tbl_el)

    header_tokens: set[str] = set()
    for h in header_texts:
        header_tokens.update(h.lower().split())

    if header_tokens:
        for prev in reversed(children[:tbl_index]):
            if prev.tag != qn("w:p"):
                break
            text = "".join(t.text or "" for t in prev.iter(qn("w:t"))).strip()
            if not text:
                continue  # blank paragraph — skip without breaking the walk
            tokens = set(text.lower().split())
            if tokens and tokens.issubset(header_tokens):
                body.remove(prev)
            else:
                break

    doc.save(str(docx_path))
    return True


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Convert PDF to editable DOCX with formatting preserved.")
    parser.add_argument("pdf", type=Path, help="Input PDF file")
    parser.add_argument("docx", type=Path, nargs="?", help="Output DOCX file (default: same basename as PDF)")
    # Stream-table detection is OFF by default — for school-curriculum-
    # style PDFs (the dominant use case here), pdf2docx's stream-table
    # heuristic invents tables out of label/value blocks and produces
    # garbage. Opt back in with --stream-tables when you have a PDF that
    # genuinely needs it (e.g. label-less ASCII tables). The GUI uses
    # the same default (the "Disable stream-table detection" checkbox is
    # checked on launch).
    parser.add_argument(
        "--stream-tables",
        action="store_true",
        help="Enable pdf2docx stream-table detection. Off by default — only "
             "use this when you have a PDF that genuinely needs stream-mode "
             "table detection. For most curriculum-style PDFs, leaving it "
             "off gives much cleaner output.",
    )
    parser.add_argument(
        "--ocr",
        action="store_true",
        help="Treat the PDF as scanned: OCR each page with tesseract and "
             "build the DOCX directly from the recognized text. Bypasses "
             "pdf2docx (which can't handle scanned PDFs).",
    )
    parser.add_argument(
        "--ocr-language",
        default="eng",
        help="Tesseract language code for --ocr (default: eng). Use a code "
             "from `tesseract --list-langs`, or combine with '+' (e.g. "
             "'slk+eng') for mixed-language documents.",
    )
    args = parser.parse_args(argv)

    pdf_path: Path = args.pdf
    if not pdf_path.exists():
        print(f"error: {pdf_path} does not exist", file=sys.stderr)
        return 1

    docx_path: Path = args.docx or pdf_path.with_suffix(".docx")

    if args.ocr:
        print(f"OCR-converting {pdf_path} -> {docx_path} (lang={args.ocr_language})")
        def _progress(cur, total):
            print(f"  page {cur}/{total}", flush=True)
        ocr_to_docx(
            pdf_path,
            docx_path,
            language=args.ocr_language,
            progress_callback=_progress,
        )
        print("Done.")
        return 0

    print(f"Converting {pdf_path} -> {docx_path}")
    convert_pdf(pdf_path, docx_path, parse_stream_table=args.stream_tables)
    patched = fix_bullet_fonts(docx_path)
    print(f"Patched {patched} bullet runs (Symbol -> {REPLACEMENT_BULLET_FONT})")
    header_fixed = fix_first_table_header(pdf_path, docx_path)
    print(f"First table header: {'re-injected from PDF' if header_fixed else 'unchanged'}")
    print("Done.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
