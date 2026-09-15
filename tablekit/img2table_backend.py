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

_cache: "dict[tuple, list[_Region]]" = {}   # (pdf_path, page_index0) -> [_Region]  (per-process)

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
                    vals = [_unscramble_reversed_paren_wrap(c.value) for c in cells]
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
    # Refuse (rather than serve) a region that looks like two unrelated
    # statements fused into one -- see _looks_like_two_fused_statements.
    # There's no safe way to split it back apart here (that would need
    # knowing exactly where one table ends and the other begins), so this
    # is the same "detect and refuse" choice already made for a merged-rows
    # cell (see _looks_garbled in extract_all_tables.py): better to fall
    # through to "no table found" for this region than silently attach one
    # statement's columns to another's rows.
    out = [r for r in out if not _looks_like_two_fused_statements(
        [v for _, _, v in r.row_rows])]
    _cache[key] = out
    return out


_REV_PAREN_LINE_RE = re.compile(
    r"^\)\s*([\d,]+(?:\.\d+)?)\s*\n\s*([\d,]+(?:\.\d+)?)\s*\(\s*$")


def _unscramble_reversed_paren_wrap(v):
    """A negative number whose parens came out reversed (")1,234(", see
    parse_number's docstring for why) can ALSO have its digits split across
    two physical lines within the same img2table cell, with the LINE ORDER
    itself reversed by the same bidi artifact -- e.g. ")69,040\\n1,1("
    for what the PDF actually printed as "(1,169,040)". Put the lines back
    in the right order (and drop the newline) before this value ever
    reaches parse_number: by the time a cell gets there it has already been
    through normspace, which collapses "\\n" to an ordinary space -- at that
    point there is no way to distinguish a genuine two-line split from an
    unrelated space-separated cell, so this has to happen here, on the raw
    value, while the newline is still a distinct signal. Deliberately
    narrow (anchored on the reversed-paren wrapper, exactly one newline,
    and both halves must look like bare digit groups) rather than a general
    "try reversing any two tokens" rule, which would risk mis-firing on
    ordinary two-number cells that have nothing to do with this artifact.
    Confirmed against two real occurrences by reconciling the surrounding
    row's own arithmetic, not just by inspection."""
    if not isinstance(v, str) or "\n" not in v:
        return v
    m = _REV_PAREN_LINE_RE.match(v.strip())
    if not m:
        return v
    line1, line2 = m.group(1), m.group(2)
    return f"){line2}{line1}("


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


def _region_has_label_column(region):
    return _has_label_column([v for _, _, v in region.row_rows])


def _merge_side_by_side(regions, x_gap=40, y_overlap_frac=0.5):
    """Detect img2table regions that are really ONE table split left/right
    (e.g. a labels block and a separate figures-only block) and re-join them.
    Two regions qualify when they're horizontally adjacent (small x-gap) and
    their overall Y-spans substantially overlap.

    That geometry alone isn't enough on a 2-column landscape page, though:
    two ENTIRELY UNRELATED full-page tables (e.g. a balance sheet on the
    left, a completely different statement of changes in equity on the
    right) can satisfy the exact same small-x-gap / big-y-overlap test,
    since both happen to run nearly the full page height. The distinguishing
    signal is the one already in this module's own docstring for what a
    legitimate split looks like: a real "figures-only block" has no label
    column of its own. If BOTH sides already have their own label column,
    they're two independently complete tables, not one table's labels half
    and figures half -- joining them would zip unrelated rows together."""
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
                if 0 <= b.x0 - a.x1 <= x_gap and y_ov > y_overlap_frac * shorter \
                        and not (_region_has_label_column(a) and _region_has_label_column(b)):
                    merged = _join_side_by_side(a, b)
                    regions[i] = merged
                    del regions[j]
                    changed = True
                    break
            if changed:
                break
    return regions


# The reversed-paren alternative goes FIRST and requires both ")" and "("
# (neither is optional there) -- some reports come out of PDF text
# extraction with a negative number's parens reversed, ")1,234(", a bidi-
# reordering artifact of the source PDF (see parse_number). When img2table
# also glues two adjacent value columns into one cell on a given row, the
# glued result is two of these reversed tokens back to back, ")87,579(
# )11,915(", not two normal ones -- ordering the reversed form first keeps
# the (both-optional) normal-parens alternative from matching only the
# digits and leaving stray "(" / ")" characters unconsumed.
_TOKEN_RE_STR = r"\)-?[\d,]+(?:\.\d+)?\(%?|\(?-?[\d,]+(?:\.\d+)?\)?%?|[-–—�]"
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


# Deliberately narrow anchors: each phrase only makes sense as the PRIMARY
# marker of one specific statement kind, not vocabulary that could plausibly
# turn up as a passing mention inside a different one.
_BS_ANCHOR_RE = re.compile(r"total assets\b|total liabilities\b|net assets\b", re.I)
_SOCE_ANCHOR_RE = re.compile(
    r"balance at \d|transactions with (the )?owners|"
    r"total comprehensive income for the year", re.I)


def _looks_like_two_fused_statements(rows):
    """True when a single img2table region carries primary-anchor vocabulary
    from two DIFFERENT statement kinds (so far: balance sheet + statement of
    changes in equity) -- the signature of img2table's own borderless-table
    clustering having fused two unrelated, independently-complete tables on
    a 2-column landscape page (found live: a balance sheet and a completely
    different equity statement on en-2021-etisalat-group-annual-report.pdf
    p62, glueing "Share capital"/"Reserves" columns onto balance-sheet rows
    like "Goodwill and other intangible assets"). That fusion happens
    geometrically -- small x-gap, large y-overlap between the two source
    regions -- which is indistinguishable from a genuine labels-block /
    figures-block split of ONE table (see _merge_side_by_side); vocabulary
    is the signal that actually is specific to two independent statements.

    Only counts a cell as a match when it doesn't read as prose: an
    accounting-policy note can mention "transactions with the owners" in a
    sentence without being anywhere near a real equity statement, and a
    first pass at this check (a 24-file, 2428-page sweep) found exactly
    that -- two accounting-policy pages that happened to use this phrase in
    a sentence, both eliminated once prose cells were excluded. With that
    filter, the same sweep found exactly one true positive and zero false
    positives across the entire real corpus."""
    bs_hit = soce_hit = False
    for row in rows:
        for c in row:
            if not (isinstance(c, str) and c.strip() and not _looks_like_prose(c)):
                continue
            if _BS_ANCHOR_RE.search(c):
                bs_hit = True
            if _SOCE_ANCHOR_RE.search(c):
                soce_hit = True
        if bs_hit and soce_hit:
            return True
    return False


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
      1. `min_overlap_frac` requires a candidate region to substantially
         overlap the drawn box (by default >=50%) before it's even
         considered -- checked BOTH ways (overlap as a fraction of the
         REGION's own area, and as a fraction of the BOX's own area), taking
         whichever is more generous. Checking region-coverage alone would
         reject a box the user drew ON PURPOSE to grab only PART of a bigger
         auto-detected region (e.g. just a table's "Assets" half, leaving
         "Liabilities" out) -- such a box can legitimately cover under 50%
         of the region while still being 100% inside it, which is exactly
         what box-coverage catches. Checking box-coverage alone would accept
         a tiny sliver of a neighbouring note that happens to fall entirely
         within a much larger drawn box -- region-coverage catches that.
         Either check passing is enough; a region only clipped at the edge
         (failing BOTH) still doesn't qualify on its own.
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
    box_area = max(1.0, (x1 - x0) * (y1 - y0))

    def _overlap_area(r):
        ox = max(0.0, min(r.x1, x1) - max(r.x0, x0))
        oy = max(0.0, min(r.y1, y1) - max(r.y0, y0))
        return ox * oy

    def _overlap_frac(r):
        area = _overlap_area(r)
        r_area = max(1.0, (r.x1 - r.x0) * (r.y1 - r.y0))
        return max(area / r_area, area / box_area)

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


def ocr_rows_in_box(pdf_path, page_index0, x0, y0, x1, y1, min_confidence=1, pad=20):
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

    `pad` matters a lot more here than it would for real-text extraction: a
    box drawn tight against the table's true edge (which is exactly what
    users do when trying to trim out extra whitespace) crops straight through
    the PIXELS of an edge character -- pdfplumber's word boxes tolerate that
    fine, but Tesseract reads a clipped digit as a different digit or drops
    it outright (confirmed live: a box ending 1-2pt past the last column
    misread "$20,565,087" as "$20,565,C" or dropped the row's tail entirely).
    20pt of padding gives OCR room to see the whole glyph regardless of how
    tightly the box itself was drawn.

    `min_confidence` defaults to 1, not img2table's usual 50: Tesseract's
    per-word confidence on this kind of image is noisy in a way that isn't
    correlated with correctness -- confirmed live on the same table, same
    PSM, same run: "2,333,277,000" scored 90 while "$55,483,771,000" right
    next to it scored 17, both equally correct. A threshold anywhere near 50
    silently drops cells Tesseract actually read right, and a silently
    MISSING cell is worse than a low-confidence one -- OCR results already
    carry the "verify by eye" badge, so a shaky-but-present read a user can
    glance at and correct beats a gap they might not notice at all.

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
        # img2table's TesseractOCR defaults to PSM 11 ("sparse text, no
        # particular order"), which on this kind of tightly-tracked numeric
        # column splits a single figure into separate low-confidence word
        # fragments ("2,299,638," + "166", each conf ~6) that then fall
        # below min_confidence and vanish from the cell entirely -- confirmed
        # by hand: a figure Tesseract reads fine as ONE token under PSM 4
        # (conf ~90) reads as two worthless fragments under PSM 11. PSM 4
        # ("a single column of text of variable sizes") matches this app's
        # actual documents -- one drawn box, one table -- and doesn't
        # fragment numbers the same way.
        tables = doc.extract_tables(
            ocr=_TesseractOCR(lang="eng", psm=4), borderless_tables=True,
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
    col_ranges = []   # [(x1_min, x2_max), ...] per column index, crop-pixel space
    for _, cells in sorted(best.content.items()):
        if not cells:
            continue
        top = cy0 + min(c.bbox.y1 for c in cells) / OCR_PX_PER_PT
        bot = cy0 + max(c.bbox.y2 for c in cells) / OCR_PX_PER_PT
        row_rows.append((top, bot, [c.value for c in cells]))
        for i, c in enumerate(cells):
            if i >= len(col_ranges):
                col_ranges.append((c.bbox.x1, c.bbox.x2))
            else:
                lo, hi = col_ranges[i]
                col_ranges[i] = (min(lo, c.bbox.x1), max(hi, c.bbox.x2))
    if not row_rows:
        return None
    rows = [list(v) for _, _, v in row_rows]
    row_bands = [(top, bot) for top, bot, _ in row_rows]
    bbox = (cx0 + best.bbox.x1 / OCR_PX_PER_PT, cy0 + best.bbox.y1 / OCR_PX_PER_PT,
            cx0 + best.bbox.x2 / OCR_PX_PER_PT, cy0 + best.bbox.y2 / OCR_PX_PER_PT)
    # a ruled/whitespace-boundaried box that only wraps the NUMBER columns --
    # with the row labels sitting to its left, inside the user's drawn box
    # but outside what img2table's own structure detection considered the
    # table -- is invisible to it, same failure mode as the real-text path
    # (see extract_all_tables._attach_left_labels' docstring). Recover them
    # here the OCR way: re-run Tesseract over just the strip to the left of
    # the detected numeric grid, within the SAME already-rendered crop (no
    # extra page render needed), and match each word span to its row by Y.
    label_hits = sum(1 for r in rows if r and isinstance(r[0], str)
                     and re.search(r"[A-Za-z]{3,}", r[0]))
    if label_hits < max(2, 0.3 * len(rows)) and best.bbox.x1 > 4:
        try:
            labels = _ocr_left_labels(crop_im.original, best.bbox.x1, cx0, cy0,
                                      [(top, bot) for top, bot, _ in row_rows])
            rows = [([lb] if lb else [None]) + r for lb, r in zip(labels, rows)]
        except Exception:
            LOG.debug("OCR label recovery failed on %s page %s box %s",
                      pdf_path, page_index0, (x0, y0, x1, y1), exc_info=True)
    # same story one more time, vertically: a caption/year header row sitting
    # ABOVE img2table's detected numeric grid (its own top edge -- best.bbox.y1
    # -- starts at the FIRST DATA row, not the header above it) is invisible
    # to it for the same reason the label column was -- it only ever sees the
    # grid it inferred from whitespace gaps between DATA rows, and a header
    # separated from the body by a ruled line reads as "outside" that grid.
    # If there's a gap above the first detected row that's roughly one line
    # tall (not the page title or three rows of unrelated prose above it),
    # OCR just that strip and, if it yields text in each of the same column
    # bands the data rows use, prepend it as a header row.
    gap_pt = row_rows[0][0] - cy0   # first detected row's top minus crop top, in PDF points
    gap_px = gap_pt * OCR_PX_PER_PT
    if 6 <= gap_pt <= 55 and col_ranges:
        try:
            header = _ocr_header_row(crop_im.original, gap_px, col_ranges, cx0)
            if header and any(header):
                rows = [header] + rows
                row_bands = [(cy0, row_rows[0][0])] + row_bands
                bbox = (bbox[0], cy0, bbox[2], bbox[3])
        except Exception:
            LOG.debug("OCR header recovery failed on %s page %s box %s",
                      pdf_path, page_index0, (x0, y0, x1, y1), exc_info=True)
    return rows, bbox, row_bands


def _ocr_header_row(crop_pil, gap_px, col_ranges, cx0):
    """OCR the horizontal strip above img2table's detected numeric grid (a
    year/column-caption row like "1994  1993" living above the ruled line
    that separates it from the data) and bucket each recovered word into
    whichever of `col_ranges` (crop-pixel x1/x2 spans, one per data column)
    it falls under, joining multiple words in the same bucket. Returns a row
    shaped like the data rows -- [None, col0_text, col1_text, ...] -- or None
    if nothing usable was found."""
    strip = crop_pil.crop((0, 0, crop_pil.width, max(1, int(round(gap_px)))))
    data = _pytesseract.image_to_data(strip, output_type=_pytesseract.Output.DICT)
    buckets = [[] for _ in col_ranges]
    for i, txt in enumerate(data["text"]):
        txt = (txt or "").strip()
        if not txt:
            continue
        cx = data["left"][i] + data["width"][i] / 2
        for ci, (lo, hi) in enumerate(col_ranges):
            if lo - 10 <= cx <= hi + 10:
                buckets[ci].append((data["left"][i], txt))
                break
    if not any(buckets):
        return None
    cells = [" ".join(t for _, t in sorted(b)) or None for b in buckets]
    return [None] + cells


def _ocr_left_labels(crop_pil, table_x1_px, cx0, cy0, row_bands):
    """OCR the strip to the left of the detected numeric columns, within the
    crop image already rendered by the caller (no extra page render), and
    return one recovered label string (or None) per entry in `row_bands` --
    the OCR equivalent of extract_all_tables._attach_left_labels, which does
    the same thing against the PDF's own real text layer."""
    strip = crop_pil.crop((0, 0, max(1, int(round(table_x1_px))), crop_pil.height))
    data = _pytesseract.image_to_data(strip, output_type=_pytesseract.Output.DICT)
    words = []
    for i, txt in enumerate(data["text"]):
        txt = (txt or "").strip()
        if not txt:
            continue
        left, top, _, h = (data["left"][i], data["top"][i],
                           data["width"][i], data["height"][i])
        words.append({
            "text": txt,
            "x0": cx0 + left / OCR_PX_PER_PT,
            "top": cy0 + top / OCR_PX_PER_PT,
            "bottom": cy0 + (top + h) / OCR_PX_PER_PT,
        })
    out = []
    for top, bot in row_bands:
        band = sorted((w for w in words if top - 2 <= w["top"] <= bot + 2),
                      key=lambda w: w["x0"])
        out.append(" ".join(w["text"] for w in band).strip() or None)
    return out


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
