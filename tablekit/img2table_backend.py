"""
Optional second geometry strategy for borderless / landscape statements.

img2table (OpenCV-based, MIT licensed) infers table structure from whitespace
gaps in a *rendered* page image, then -- because pdf_text_extraction=True --
reads the actual TEXT for each cell from the PDF's own text layer, not OCR.

Tested against this project's corpus (2026-09): it clearly beats the
hand-written regex reconstruction on pages where the auditor's report is
interleaved with the statement in the same physical column band (du's
2010-2014 landscape spreads) -- exactly the failure mode that pass after pass
of regex patching in extract_all_tables.py never fully solved. It is
noticeably WORSE than our own reconstruction when a statement's value columns
are narrow and tightly packed (e.g. Etisalat's older 3-year balance sheet),
because its column boundaries come from whitespace-gap size, and a narrow gap
reads the same as no gap at all -- several years' figures get merged into one
cell as a literal string like "432,541 221,711 205,270".

Two correctness traps found and fixed during integration (see git history /
CHANGELOG for the concrete before/after):
  1. Pre-cropping a landscape page to a fixed half BEFORE calling img2table
     throws away its own (more precise) region-boundary detection -- run it
     ONCE on the whole page, uncropped, then match its regions to a heading
     by position (`rows_under_heading`).
  2. When img2table splits one logical table into a left "labels" region and
     a right "figures-only" region, naively zipping their rows together BY
     POSITION IN THE LIST silently mispairs a label with the wrong figure the
     moment the two regions disagree on row count internally (e.g. the labels
     region merges several physical lines into one row while the numbers
     region -- driven by isolated numeric tokens -- doesn't). Fixed by
     reading each cell's own Y-coordinate (img2table's `ExtractedTable.content`
     keeps per-cell bboxes; the flattened `.df` throws that away) and pairing
     rows by actual vertical overlap, with the finer-grained side driving so a
     merged label is reused across the rows it truly spans rather than
     silently attached to the wrong one.

So this is wired in as a CHALLENGER, not a replacement: for every statement
heading `_recon_tables` finds, extract_all_tables.py matches the img2table
region(s) under that heading, scores the result with the real analyze()
reconciliation check against the regex-based candidate, and keeps whichever
one actually foots (see `_recon_tables` in extract_all_tables.py). If
img2table isn't installed, `img2table_page_tables` returns [] and nothing
downstream changes.

Optional dependencies: img2table, opencv-python-headless, pandas.
"""
from __future__ import annotations

import logging
import re

LOG = logging.getLogger("tablekit.img2table")

try:
    from img2table.document import PDF as _Img2TablePDF
    HAVE_IMG2TABLE = True
except Exception:
    HAVE_IMG2TABLE = False

# OCR is a separate, optional capability on top of img2table: it needs both
# the `img2table.ocr.TesseractOCR` reader AND the actual Tesseract binary
# installed on the machine (pip alone only gets you the Python wrapper) --
# `get_tesseract_version()` is what actually proves the binary is reachable,
# same spirit as the HAVE_IMG2TABLE probe above.
import os as _os
import shutil as _shutil

if _shutil.which("tesseract") is None:
    # a package-manager install (winget's UB-Mannheim build, the Windows
    # .exe installer) doesn't reliably land tesseract.exe on PATH for an
    # already-running process -- prepend the well-known install directory to
    # THIS process's PATH before img2table.ocr.TesseractOCR ever shells out
    # to it (it always invokes the bare "tesseract" command, inheriting
    # os.environ as-is -- it has no way to point at a specific binary path).
    for _dir in (r"C:\Program Files\Tesseract-OCR",
                r"C:\Program Files (x86)\Tesseract-OCR"):
        if _os.path.exists(_os.path.join(_dir, "tesseract.exe")):
            _os.environ["PATH"] = _dir + _os.pathsep + _os.environ.get("PATH", "")
            break

try:
    from img2table.document import Image as _Img2TableImage
    from img2table.ocr import TesseractOCR as _TesseractOCR
    import pytesseract as _pytesseract
    _pytesseract.get_tesseract_version()
    HAVE_OCR = True
except Exception:
    HAVE_OCR = False

# img2table rasterises PDF pages at this many pixels per PDF point (see
# img2table/document/pdf.py: `page.render(scale=200 / 72)`) -- multiply its
# pixel bboxes by PT_PER_PX to get back to PDF points (pdfplumber's units).
PT_PER_PX = 72.0 / 200.0

# resolution OCR renders the drawn box at (pixels per PDF point). Only the
# box itself is rendered -- never the whole page -- which is what keeps a
# manual OCR extraction a sub-second click instead of a multi-minute scan.
OCR_PX_PER_PT = 4.0

_cache = {}   # (pdf_path, page_index0) -> [_Region]  (per-process)

_LEAK_RE = re.compile(r"[a-z]{3,}\s+[a-z]{3,}\s+[a-z]{3,}\s+[a-z]{3,}\s+[a-z]{3,}", re.I)


class _Region:
    """One img2table-detected table region, keeping each row's own Y-extent
    (row_top, row_bottom, values) instead of the flattened dataframe -- the
    Y-extent is what makes correct row alignment across split regions
    possible (see module docstring, correctness trap #2)."""
    __slots__ = ("x0", "y0", "x1", "y1", "row_rows")

    def __init__(self, x0, y0, x1, y1, row_rows):
        self.x0, self.y0, self.x1, self.y1 = x0, y0, x1, y1
        self.row_rows = row_rows   # [(top, bottom, [values])], PDF points, sorted by top

    @property
    def ncols(self):
        return max((len(v) for _, _, v in self.row_rows), default=0)

    @property
    def nrows(self):
        return len(self.row_rows)


def img2table_page_tables(pdf_path, page_index0, min_confidence=50):
    """Run img2table ONCE on the whole (uncropped) page.  Returns a list of
    `_Region`, with all coordinates converted to PDF points, top-down from
    the page's own top-left (pdfplumber's convention) so they're directly
    comparable to a heading's `ln["top"]`/`ln["x0"]`. Cached per (path, page)
    since a page can have several headings that each need to consult it.
    """
    if not HAVE_IMG2TABLE:
        return []
    key = (str(pdf_path), page_index0)
    if key in _cache:
        return _cache[key]
    out = []
    try:
        doc = _Img2TablePDF(src=str(pdf_path), pages=[page_index0], pdf_text_extraction=True)
        result = doc.extract_tables(
            borderless_tables=True, implicit_rows=True, implicit_columns=True,
            min_confidence=min_confidence)
        for _, tbls in result.items():
            for t in tbls:
                row_rows = []
                for _, cells in t.content.items():
                    if not cells:
                        continue
                    top = min(c.bbox.y1 for c in cells) * PT_PER_PX
                    bot = max(c.bbox.y2 for c in cells) * PT_PER_PX
                    vals = [c.value for c in cells]
                    row_rows.append((top, bot, vals))
                row_rows.sort(key=lambda r: r[0])
                if not row_rows:
                    continue
                x0 = t.bbox.x1 * PT_PER_PX
                y0 = t.bbox.y1 * PT_PER_PX
                x1 = t.bbox.x2 * PT_PER_PX
                y1 = t.bbox.y2 * PT_PER_PX
                out.append(_Region(x0, y0, x1, y1, row_rows))
    except Exception:
        LOG.debug("img2table failed on page %s", page_index0, exc_info=True)
        out = []
    _cache[key] = out
    return out


def _y_overlap(a_top, a_bot, b_top, b_bot):
    return max(0.0, min(a_bot, b_bot) - max(a_top, b_top))


def _join_side_by_side(left: _Region, right: _Region) -> _Region:
    """Pair each row of the FINER (more rows) region with its best Y-overlap
    match in the COARSER region, so a label merged across several physical
    lines is reused for each figure row it actually spans, rather than
    index-zipped into a silently wrong pairing."""
    driver, other, driver_is_left = (
        (left, right, True) if left.nrows >= right.nrows else (right, left, False))
    out_rows = []
    for d_top, d_bot, d_vals in driver.row_rows:
        best, best_ov = None, 0.0
        for o_top, o_bot, o_vals in other.row_rows:
            ov = _y_overlap(d_top, d_bot, o_top, o_bot)
            if ov > best_ov:
                best, best_ov = o_vals, ov
        other_vals = best if best is not None else [None] * other.ncols
        vals = (list(d_vals) + list(other_vals)) if driver_is_left else (list(other_vals) + list(d_vals))
        out_rows.append((d_top, d_bot, vals))
    x0, x1 = min(left.x0, right.x0), max(left.x1, right.x1)
    y0, y1 = min(left.y0, right.y0), max(left.y1, right.y1)
    return _Region(x0, y0, x1, y1, out_rows)


def _merge_side_by_side(regions, x_gap=40, y_overlap_frac=0.5):
    """Detect img2table regions that are really ONE table split left/right
    (e.g. a labels block and a separate figures-only block) and re-join them.
    Two regions qualify when they're horizontally adjacent (small x-gap) and
    their overall Y-spans substantially overlap."""
    regions = list(regions)
    changed = True
    while changed and len(regions) > 1:
        changed = False
        for i, a in enumerate(regions):
            for j, b in enumerate(regions):
                if i == j or b.x0 < a.x0:
                    continue
                y_ov = _y_overlap(a.y0, a.y1, b.y0, b.y1)
                shorter = min(a.y1 - a.y0, b.y1 - b.y0) or 1.0
                if 0 <= b.x0 - a.x1 <= x_gap and y_ov > y_overlap_frac * shorter:
                    merged = _join_side_by_side(a, b)
                    regions[i] = merged
                    del regions[j]
                    changed = True
                    break
            if changed:
                break
    return regions


_TOKEN_RE_STR = r"\(?-?[\d,]+(?:\.\d+)?\)?%?|[-–—�]"
_GLUED_CELL_RE = re.compile(rf"^(?:{_TOKEN_RE_STR})(?:\s+(?:{_TOKEN_RE_STR}))+$")
_TOKEN_FIND_RE = re.compile(_TOKEN_RE_STR)


def _split_glued_cell(v):
    """A cell whose text is two (or more) number-like tokens (or the '�'
    nil-placeholder glyph) separated by whitespace, and nothing else, is
    img2table having merged two adjacent value columns into one because the
    gap between them was too narrow on THIS particular row. Split it back
    into its parts; anything else is returned unchanged."""
    if not isinstance(v, str):
        return [v]
    s = v.strip()
    if _GLUED_CELL_RE.match(s):
        parts = _TOKEN_FIND_RE.findall(s)
        if len(parts) >= 2:
            return parts
    return [v]


def _normalize_row_width(vals, ncols):
    """Pad/expand a row to exactly `ncols` values, splitting a glued cell
    (see `_split_glued_cell`) starting from the rightmost cell when the row
    is short -- this is what a narrower stacked region's rows need before
    they can be safely concatenated onto a wider region's rows (see
    `_merge_stacked`); without it, two real columns silently collapse into
    one glued string for every row in the narrower region."""
    vals = list(vals)
    while len(vals) < ncols:
        expanded = False
        for i in range(len(vals) - 1, -1, -1):
            parts = _split_glued_cell(vals[i])
            if len(parts) > 1:
                vals[i:i + 1] = parts
                expanded = True
                break
        if not expanded:
            vals = vals + [None] * (ncols - len(vals))
            break
    return vals[:ncols]


def _merge_stacked(regions, max_gap=60, col_tol=1, x_tol=25):
    """Concatenate vertically-stacked pieces of the same column count that
    img2table split apart (typically at a section-header blank-row gap).
    When the two pieces' column counts differ by up to `col_tol`, the
    narrower piece's rows are widened to match (splitting any glued cell)
    rather than concatenated as-is -- otherwise every row from the narrower
    piece ends up misaligned by one column against the wider piece.

    Column count and Y-adjacency alone aren't enough of a test: two
    unrelated notes that both happen to render as e.g. a 4-column table and
    sit close together vertically (common on a dense notes page, or when
    `rows_in_box`'s candidate pool ends up wider than one column because the
    user's box was drawn generously) would otherwise get glued into one
    corrupted table. Require the pieces' X-ranges to actually line up too."""
    regions = sorted(regions, key=lambda r: r.y0)
    out = []
    for r in regions:
        if (out and abs(r.ncols - out[-1].ncols) <= col_tol
                and r.y0 - out[-1].y1 <= max_gap
                and abs(r.x0 - out[-1].x0) <= x_tol):
            prev = out[-1]
            target_ncols = max(prev.ncols, r.ncols)
            fixed_rows = [(top, bot, _normalize_row_width(v, target_ncols))
                         for top, bot, v in r.row_rows]
            out[-1] = _Region(prev.x0, prev.y0, max(prev.x1, r.x1), r.y1,
                              prev.row_rows + fixed_rows)
        else:
            out.append(_Region(r.x0, r.y0, r.x1, r.y1, list(r.row_rows)))
    return out


def _looks_like_prose(s):
    return bool(_LEAK_RE.search(s)) or len(s.split()) >= 8


def _drop_prose_columns(rows):
    """img2table sometimes keeps a whole column of interleaved prose (an
    unrelated auditor's-report paragraph physically beside the statement) as
    its own column in the same region as the real label column. Drop any
    column that reads as prose on most of its non-empty cells, UNLESS
    dropping it would leave no text column at all."""
    if not rows:
        return rows
    ncols = max(len(r) for r in rows)
    prose_frac = []
    for c in range(ncols):
        cells = [r[c] for r in rows if c < len(r) and isinstance(r[c], str) and r[c].strip()]
        if not cells:
            prose_frac.append(0.0)
            continue
        prose_frac.append(sum(1 for s in cells if _looks_like_prose(s)) / len(cells))
    drop = {c for c, f in enumerate(prose_frac) if f >= 0.5}
    text_cols_left = sum(1 for c in range(ncols)
                         if c not in drop and any(
                             c < len(r) and isinstance(r[c], str) and r[c].strip()
                             for r in rows))
    if not drop or text_cols_left == 0:
        return rows
    keep = [c for c in range(ncols) if c not in drop]
    return [[r[c] if c < len(r) else None for c in keep] for r in rows]


def _has_label_column(rows):
    if not rows:
        return False
    hits = sum(1 for r in rows if r and isinstance(r[0], str)
              and re.search(r"[A-Za-z]{3,}", r[0]))
    return hits >= max(3, 0.3 * len(rows))


def rows_in_box(page_tables, x0, y0, x1, y1, min_overlap_frac=0.5, row_margin=20):
    """Select and merge the img2table region(s) that substantially overlap a
    user-drawn box (manual selection mode) -- keyed by geometric overlap
    alone, unlike `rows_under_heading`, since the user pointed directly at
    the table and there's no other heading to disambiguate against.

    Two safeguards against a box that's drawn a bit generously (common --
    users round outward "to be safe"), which on a dense notes page can
    otherwise drag in a neighbouring note's heading or an unrelated column
    of prose that happens to sit close by:
      1. `min_overlap_frac` requires a candidate region to be MOSTLY (by
         default >=50%) covered by the drawn box before it's even considered
         -- a region only clipped at the edge doesn't qualify on its own.
      2. After merging, rows are trimmed to the drawn box's own Y-range
         (plus a small `row_margin` of slack for an edge that's a few points
         short) -- this is what actually stops a stacked-merge from pulling
         the next section's header in just because it was within the
         stacking gap tolerance.

    Returns (rows, bbox, row_bands) or None. `bbox` is derived ONLY from the
    ROWS THAT SURVIVED trimming (never unioned with the user's raw drawn
    box) -- so the highlight shown back is a tight, faithful reflection of
    exactly what was extracted, not inflated by an imprecisely-drawn input
    box. `row_bands` is `[(top, bot), ...]`, one per entry in `rows`, in the
    SAME order -- lets a caller reattach a label column img2table missed
    (see module docstring: a ruled box that only wraps the numbers, with
    labels sitting outside it, is invisible to img2table's own structure
    detection -- it only ever sees the ruled numeric grid)."""
    def _overlap_area(r):
        ox = max(0.0, min(r.x1, x1) - max(r.x0, x0))
        oy = max(0.0, min(r.y1, y1) - max(r.y0, y0))
        return ox * oy

    def _overlap_frac(r):
        area = _overlap_area(r)
        r_area = max(1.0, (r.x1 - r.x0) * (r.y1 - r.y0))
        return area / r_area

    cands = [r for r in page_tables if _overlap_frac(r) >= min_overlap_frac]
    if not cands:
        return None
    cands = _merge_side_by_side(cands)
    cands = _merge_stacked(cands)
    if not cands:
        return None
    cands.sort(key=_overlap_area, reverse=True)
    best = cands[0]

    kept = [(top, bot, vals) for top, bot, vals in best.row_rows
            if bot >= y0 - row_margin and top <= y1 + row_margin]
    if not kept:
        kept = best.row_rows   # trimming would leave nothing -- don't blank the result

    rows_of = lambda rr: [list(v) for _, _, v in rr]
    raw_rows = rows_of(kept)
    rows = _drop_prose_columns(raw_rows)
    if not rows:
        return None
    # _drop_prose_columns only ever drops COLUMNS, never rows, so `kept` and
    # `rows` still line up 1:1 regardless of whether any column got dropped
    row_bands = [(top, bot) for top, bot, _ in kept]
    tight_y0 = min(top for top, _, _ in kept)
    tight_y1 = max(bot for _, bot, _ in kept)
    return rows, (best.x0, tight_y0, best.x1, tight_y1), row_bands


def ocr_rows_in_box(pdf_path, page_index0, x0, y0, x1, y1, min_confidence=50, pad=6):
    """OCR failsafe for a manually-drawn box that has NO extractable text at
    all (a scanned page, or a scanned figure pasted into an otherwise-digital
    report). Only ever called after the real-text path (`rows_in_box` /
    pdfplumber's own `.extract_table()`) has already been tried and found
    nothing AND the box's own text layer came back empty -- this is a
    failsafe, not a second attempt at parsing text that's already there.

    Renders and OCRs ONLY the drawn box (never the whole page) so this stays
    a sub-second click instead of a multi-minute whole-document scan -- the
    same "opt-in, page/box at a time" shape as manual mode's real-text path,
    just with Tesseract standing in for the PDF's own text layer.

    Returns (rows, bbox, row_bands) shaped exactly like `rows_in_box`'s hit,
    or None. `bbox`/`row_bands` are in the same PDF-point space as every
    other coordinate in this module, converted back from the OCR image's own
    pixel space via `OCR_PX_PER_PT`.
    """
    if not HAVE_OCR:
        return None
    import io
    import pdfplumber

    try:
        with pdfplumber.open(pdf_path) as pdf:
            page = pdf.pages[page_index0]
            cx0, cy0 = max(0.0, x0 - pad), max(0.0, y0 - pad)
            cx1, cy1 = min(page.width, x1 + pad), min(page.height, y1 + pad)
            crop_im = page.crop((cx0, cy0, cx1, cy1)).to_image(
                resolution=int(72 * OCR_PX_PER_PT))
        buf = io.BytesIO()
        crop_im.original.save(buf, format="PNG")
        doc = _Img2TableImage(src=buf.getvalue())
        tables = doc.extract_tables(
            ocr=_TesseractOCR(lang="eng"), borderless_tables=True,
            implicit_rows=True, implicit_columns=True,
            min_confidence=min_confidence)
    except Exception:
        LOG.debug("OCR failed on %s page %s box %s", pdf_path, page_index0,
                  (x0, y0, x1, y1), exc_info=True)
        return None
    if not tables:
        return None
    # the box was hand-drawn around one table, so pick whichever OCR region
    # has the most cells rather than trying to merge/disambiguate several
    best = max(tables, key=lambda t: sum(len(c) for c in t.content.values()))
    row_rows = []
    for _, cells in sorted(best.content.items()):
        if not cells:
            continue
        top = cy0 + min(c.bbox.y1 for c in cells) / OCR_PX_PER_PT
        bot = cy0 + max(c.bbox.y2 for c in cells) / OCR_PX_PER_PT
        row_rows.append((top, bot, [c.value for c in cells]))
    if not row_rows:
        return None
    rows = [list(v) for _, _, v in row_rows]
    row_bands = [(top, bot) for top, bot, _ in row_rows]
    bbox = (cx0 + best.bbox.x1 / OCR_PX_PER_PT, cy0 + best.bbox.y1 / OCR_PX_PER_PT,
            cx0 + best.bbox.x2 / OCR_PX_PER_PT, cy0 + best.bbox.y2 / OCR_PX_PER_PT)
    return rows, bbox, row_bands


def rows_under_heading(page_tables, heading_top, heading_x0, x_lo, x_hi,
                       other_head_x0s=(), max_span=2000):
    """From the whole-page img2table result, recover the ONE logical table
    that belongs to a heading: starts at/just below `heading_top`, overlaps
    the heading's half of the page [x_lo, x_hi], and -- critically, on a 2-up
    page with another statement heading nearby -- is not actually closer to
    that OTHER heading (a region can graze this heading's half by a few
    points while its bulk belongs to a wholly separate, adjacent statement;
    on du's 2025 layout the changes-in-equity and cash-flow statements sit
    side by side and a naive overlap test pulled the cash flow's own region
    into the equity statement's candidate pool). Region center distance to
    each heading's x0 is the tie-breaker; a fixed width-fraction threshold
    was tried first and rejected -- it also excluded du 2010's legitimate
    (barely-overlapping, but uncontested) labels region."""
    def _closer_to_other(r):
        if not other_head_x0s:
            return False
        center = (r.x0 + r.x1) / 2
        d_self = abs(center - heading_x0)
        return any(abs(center - ox) < d_self for ox in other_head_x0s)

    cands = [r for r in page_tables
             if r.y0 >= heading_top - 10 and r.y0 <= heading_top + max_span
             and r.x1 > x_lo and r.x0 < x_hi
             and not _closer_to_other(r)]
    if not cands:
        return None
    cands = _merge_side_by_side(cands)
    cands = _merge_stacked(cands)
    if not cands:
        return None
    rows_of = lambda r: [list(v) for _, _, v in r.row_rows]
    cands.sort(key=lambda r: (_has_label_column(rows_of(r)), r.nrows), reverse=True)
    return _drop_prose_columns(rows_of(cands[0])) or None
