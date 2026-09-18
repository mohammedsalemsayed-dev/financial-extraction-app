# -*- coding: utf-8 -*-
"""Synthesize a single-page PDF from one FinTabNet.c annotation JSON entry.

FinTabNet.c's publicly distributed "PDF_Annotations" archive ships ONLY the
ground-truth JSON (per-cell text + PDF-coordinate bounding boxes) -- not the
original source PDFs (no mirror in active use redistributes those; even the
docling-eval project's own FinTabNet harness works from rendered table
IMAGES, not PDFs). Our extractor is pdfplumber-based and needs a real text
layer, not an image, so neither the annotation-only archive nor an
image-based mirror can be fed to it directly.

Workaround: the annotation JSON already records exactly what a real PDF's
text layer would have contained for this table -- each cell's precise
PDF-coordinate bounding box (`pdf_text_tight_bbox`) and its text
(`pdf_text_content`). Re-drawing that same text at those same positions on a
correctly-sized blank page produces a PDF whose text layer is
geometrically faithful to the original, without needing the original file
itself. pdfplumber reads the result exactly as it would the source PDF.

Known, accepted limitation: this schema has no vector-graphics data (ruled
lines, cell borders), so a synthesized page never has any. That specifically
exercises the borderless/text-position detection path -- which happens to be
the path this whole tool has been built and tuned around (real financial
statements are usually borderless), so it is a fair, if narrower, test of
the tool's actual specialty, not an artificial handicap invented for this
harness.

Coordinate note: FinTabNet's `pdf_bbox`/`pdf_text_tight_bbox` fields use a
TOP-left origin (y grows downward, matching pdfplumber's own `top`/`bottom`
convention) -- confirmed empirically (header rows sit at low y, near the top
of the page). reportlab's canvas uses PDF-native BOTTOM-left origin, so y is
flipped on every draw call: `y_reportlab = page_height - y_top_left`.
"""
from __future__ import annotations

from reportlab.pdfgen import canvas
from reportlab.pdfbase.pdfmetrics import stringWidth

FONT = "Helvetica"


def synthesize_pdf(table, out_path):
    """`table`: one entry from a FinTabNet.c `*_tables.json` file (a dict
    with `pdf_full_page_bbox`, `cells`, etc -- see module docstring).
    Writes a single-page PDF to `out_path`."""
    page_w, page_h = table["pdf_full_page_bbox"][2], table["pdf_full_page_bbox"][3]
    c = canvas.Canvas(str(out_path), pagesize=(page_w, page_h))
    for cell in table["cells"]:
        text = (cell.get("pdf_text_content") or "").strip()
        if not text:
            continue
        bbox = cell.get("pdf_text_tight_bbox") or cell.get("pdf_bbox")
        if not bbox:
            continue
        x0, y0, x1, y1 = bbox
        box_h = max(y1 - y0, 1.0)
        box_w = max(x1 - x0, 1.0)
        # Fit the real font size to the box the ground truth says this text
        # occupied, capped to a plausible statement-table range -- exact
        # font metrics will never match the original's real typeface, but
        # position (which is what our extractor's word-clustering actually
        # keys on) is what matters here, not pixel-perfect rendering.
        #
        # NO overflow tolerance here -- an earlier version allowed the
        # rendered text to run up to 15% wider than the ground truth's own
        # bbox, on the theory that a slightly-too-small font reads worse
        # than a slightly-too-wide one. Found live: CAH_2018 p77's
        # "International" row, 2016 column -- ground truth text "4,682"
        # tight-bbox ends at x1=545.8; Helvetica at the size that 15%
        # tolerance allowed rendered to x1=547.7, ~2pt past pdfplumber's own
        # auto-detected column boundary (inherited from OTHER rows' shorter
        # numbers in the same column) -- just enough for pdfplumber's table
        # cell-cropping to silently drop the trailing "2", scoring as if our
        # EXTRACTOR corrupted a digit when the real cause was our own
        # synthesis rendering wider than the source. Fitting strictly
        # inside the recorded box (down to a 4pt floor rather than ever
        # overflowing) keeps a table's own internal column alignment
        # faithful to the ground truth, which matters far more here than
        # matching the original's exact font size.
        size = min(max(box_h * 0.8, 5.0), 12.0)
        while size > 4.0 and stringWidth(text, FONT, size) > box_w:
            size -= 0.5
        c.setFont(FONT, size)
        baseline_y = page_h - y1 + (box_h - size) * 0.3
        c.drawString(x0, baseline_y, text)
    c.showPage()
    c.save()


def build_ground_truth_rows(table):
    """Reconstruct a plain row-major grid (list of row-lists, matching this
    project's own `rows` convention) from FinTabNet.c's cell/row/column
    structure -- a spanning cell's text is repeated into every (row, col)
    position it covers, same as how a human transcribing the table would
    read it column by column."""
    n_rows = len(table["rows"])
    n_cols = len(table["columns"])
    grid = [[None] * n_cols for _ in range(n_rows)]
    for cell in table["cells"]:
        text = (cell.get("pdf_text_content") or "").strip() or None
        for r in cell["row_nums"]:
            for col in cell["column_nums"]:
                if r < n_rows and col < n_cols:
                    grid[r][col] = text
    return grid
