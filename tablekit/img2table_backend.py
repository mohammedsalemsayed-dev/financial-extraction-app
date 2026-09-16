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


class _RowRow(tuple):
    """A (top, bottom, values) row that also carries each of THIS row's own
    cells' raw (pre-collapse-artifact, pre-repair-pipeline) (x1, x2, value)
    geometry, for callers that want to filter a row by column position (see
    grid_in_box) -- riding the data inside the tuple itself, rather than a
    second list on _Region kept in lockstep by position, is deliberate: a
    separate list would need every one of the handful of places row_rows
    gets rebuilt (the repair pipeline, both merge functions, rows_in_box's
    own trim) to also rebuild it in parallel, which is exactly the
    "index-based silent mispairing" failure class this module's own
    docstring already names as a previously-fixed real bug, for `vals`
    itself. `top, bot, vals = row` keeps working exactly as it always has --
    this is a plain 3-tuple to every existing caller; only `.raw_cells`
    (None when not known/applicable) is new.

    `raw_cells`: [(x1, x2, value), ...] in PDF points, one per cell BEFORE
    the repair pipeline (_normalize_row_width / _redistribute_glued_cells /
    _split_glued_row) can split or merge cells -- so its own element count
    can differ from `vals`'s. That pipeline is never re-run against
    fabricated sub-cell positions; a caller that wants geometry-filtered
    values re-runs the SAME, unmodified repair functions fresh on a
    filtered subset of raw_cells instead (see grid_in_box).

    No `__slots__` here -- CPython doesn't support adding non-empty slots to
    a `tuple` subclass (its variable-length layout leaves no room), so this
    carries a normal `__dict__` instead; a per-row object is not a hot
    enough path for the extra pointer to matter."""

    def __new__(cls, top, bot, vals, raw_cells=None):
        obj = tuple.__new__(cls, (top, bot, vals))
        obj._raw_cells = raw_cells
        return obj

    @property
    def raw_cells(self):
        return self._raw_cells


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
                    # img2table represents a row it couldn't find real column
                    # boundaries for (sparse/no vertical ruling between cells
                    # on THIS row -- common for a header line spanning every
                    # column, or a whole data row with tight number spacing)
                    # as N cells that all carry an IDENTICAL COPY of the full
                    # text, not one real cell plus blanks. Left as-is, every
                    # value shows up duplicated in every column instead of
                    # once -- found live: a real "Property, plant and
                    # equipment 6 8,063,422 7,741,119" data row came back
                    # with that entire glued string repeated 4 times, once
                    # per column, on du annual 2020's balance sheet. Collapse
                    # it to a single cell now (keyed on 2+ words so a
                    # genuinely repeated short token, e.g. "-" in several
                    # real empty columns, is left alone); _split_glued_row
                    # below (once the table's real column count is known)
                    # pulls the trailing figures back out into their own
                    # cells.
                    if (len(vals) > 1 and vals[0] is not None
                            and all(v == vals[0] for v in vals[1:])
                            and len(str(vals[0]).split()) >= 2):
                        vals = [vals[0]] + [None] * (len(vals) - 1)
                    # Captured AFTER the collapse above, not before: a row
                    # collapsed to one real cell should carry ONE raw_cells
                    # entry reflecting that (matching `vals`), not N entries
                    # all claiming the same duplicated text under N
                    # different x-positions -- narrowing a box to one of
                    # those "columns" later would otherwise resurrect the
                    # whole duplicated line instead of correctly showing
                    # nothing. Same length as `vals`/`cells` always (the
                    # collapse above only replaces elements with None, never
                    # removes any), so this zip is safe unconditionally.
                    raw_cells = [(c.bbox.x1 * PT_PER_PX, c.bbox.x2 * PT_PER_PX, v)
                                for c, v in zip(cells, vals)]
                    row_rows.append((top, bot, vals, raw_cells))
                row_rows.sort(key=lambda r: r[0])
                if not row_rows:
                    continue
                # the collapse above can leave a row as one wide cell
                # ("Property, plant and equipment 6 8,063,422 7,741,119")
                # where every OTHER row in this same region cleanly filled
                # every column -- pull the trailing figures (as many as the
                # table's own column count calls for) back into their own
                # cells now that a real width is known; whatever's left
                # (the label, possibly with its own trailing note-ref
                # number) stays in column 0, where _strip_note_refs
                # downstream already knows how to clean up a trailing note
                # number the normal way.
                #
                # Three DIFFERENT partial-failure shapes, run in sequence:
                # _normalize_row_width first for a row that's genuinely
                # SHORTER than the table's real width with its glued cell
                # at the end (found live on Etisalat's balance sheet --
                # ["Contract assets", 22, "432,541 221,711 205,270"] for a
                # 5-wide, 3-year table -- previously only ever applied
                # during _merge_stacked, never as a general per-row
                # cleanup); then _redistribute_glued_cells for a row
                # that's ALREADY the full width but has a glued cell mid-
                # row with unused None padding after it (_normalize_row_
                # width can't reach this -- it only ever EXPANDS a short
                # row); then _split_glued_row for a row collapsed to one
                # wide cell above.
                ncols = max((len(v) for _, _, v, _ in row_rows), default=0)
                row_rows = [_RowRow(top, bot,
                                    _split_glued_row(
                                        _redistribute_glued_cells(
                                            _normalize_row_width(v, ncols)), ncols),
                                    raw_cells=rc)
                            for top, bot, v, rc in row_rows]
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
    index-zipped into a silently wrong pairing.

    Reuses this SAME match for each row's raw_cells (see _RowRow) rather
    than running a separate Y-overlap pass for it -- a separate pass could
    disagree with the one `vals` already used and silently desync the two,
    which is exactly the class of bug this function's own Y-overlap
    matching was originally built to avoid for `vals` itself. If either
    side's matched row has no raw_cells (e.g. a synthetic region built
    without it, or no Y-overlap match found at all), the merged row's
    raw_cells is None rather than a partial/fabricated guess -- a caller
    filtering by column position on a None-geometry row should pass it
    through unfiltered, not silently drop it."""
    driver, other, driver_is_left = (
        (left, right, True) if left.nrows >= right.nrows else (right, left, False))
    out_rows = []
    for d_row in driver.row_rows:
        d_top, d_bot, d_vals = d_row
        best, best_ov = None, 0.0
        for o_row in other.row_rows:
            o_top, o_bot, _ = o_row
            ov = _y_overlap(d_top, d_bot, o_top, o_bot)
            if ov > best_ov:
                best, best_ov = o_row, ov
        other_vals = best[2] if best is not None else [None] * other.ncols
        vals = (list(d_vals) + list(other_vals)) if driver_is_left else (list(other_vals) + list(d_vals))
        d_raw = getattr(d_row, "raw_cells", None)
        o_raw = getattr(best, "raw_cells", None) if best is not None else None
        raw_cells = None
        if d_raw is not None and o_raw is not None:
            raw_cells = (list(d_raw) + list(o_raw)) if driver_is_left else (list(o_raw) + list(d_raw))
        out_rows.append(_RowRow(d_top, d_bot, vals, raw_cells=raw_cells))
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


def _redistribute_glued_cells(vals):
    """A row already at the table's real width can still have ONE cell
    holding several glued numeric tokens ("432,541 221,711 205,270")
    while the cells right after it are untouched None padding -- found
    live on Etisalat's balance sheet: ["Contract assets", 22,
    "432,541 221,711 205,270", None, None] for a 5-column, 3-year table.
    img2table found the right number of real columns overall but merged
    two or three of them into one on THIS particular row (a narrower gap
    between them here than on other rows). _normalize_row_width doesn't
    reach this -- it only ever EXPANDS a row that's shorter than ncols,
    and this row already reports the full width. Split the glued cell
    and shift its extra pieces into the trailing None run right after
    it, only when there's room for every piece -- never into a slot that
    already holds something real, which would mean this isn't actually
    the shape it looks like."""
    vals = list(vals)
    for i in range(len(vals)):
        parts = _split_glued_cell(vals[i])
        if len(parts) <= 1:
            continue
        end = i + len(parts)
        if end <= len(vals) and all(vals[k] is None for k in range(i + 1, end)):
            vals[i:end] = parts
    return vals


def _split_glued_row(vals, ncols):
    """A row img2table couldn't find column boundaries for at all comes
    back (after the identical-cell-per-column repair in
    img2table_page_tables) as one wide cell in column 0 and blanks
    elsewhere -- "Property, plant and equipment 6 8,063,422 7,741,119"
    for a 4-column table. Pull number-like tokens off the END into their
    own cells, greedily, stopping at the first non-number-like token or
    once (ncols - 1) cells are filled -- NOT a fixed count: a "Total ..."
    row legitimately has fewer real values than the table's widest row
    (no note-ref number of its own), so demanding exactly (ncols - 1)
    trailing tokens would wrongly refuse to split it. Whatever's left
    (the label, possibly with its own trailing note-ref number, "...
    equipment 6") stays in column 0 for _strip_note_refs to clean up
    downstream, same as any other row's label. A row whose numbers sit
    mid-label instead of at the very end (a wrapped label split across
    the figures, e.g. "...through other 11 18,368 18,368 comprehensive
    income") has no trailing numeric tokens to find at all and is
    correctly left alone -- a narrower, rarer shape not handled here.
    Not the same job as _split_glued_cell/_normalize_row_width above --
    those split a cell that's ENTIRELY number tokens; this one has a
    real label glued onto the front."""
    if ncols <= 1 or not vals or vals[0] is None or any(v is not None for v in vals[1:]):
        return vals
    tokens = str(vals[0]).split()
    if len(tokens) < 3:      # need at least label + 2 real trailing figures
        return vals
    max_want = ncols - 1
    tail = []
    i = len(tokens) - 1
    while i >= 0 and len(tail) < max_want and _TOKEN_FIND_RE.fullmatch(tokens[i]):
        tail.insert(0, tokens[i])
        i -= 1
    if len(tail) < 2:
        return vals
    label = " ".join(tokens[:i + 1])
    if not label:
        return vals
    return [label] + tail


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
            # raw_cells rides through unchanged (see _RowRow) -- widening a
            # short row's VALUES (splitting a glued cell purely textually)
            # doesn't invent any real sub-cell x-position to assign the new
            # piece, so there's nothing correct to update it to; a caller
            # that wants geometry-filtered values re-derives them from the
            # region's own real cells instead (see grid_in_box).
            fixed_rows = [_RowRow(row[0], row[1], _normalize_row_width(row[2], target_ncols),
                                  raw_cells=getattr(row, "raw_cells", None))
                         for row in r.row_rows]
            out[-1] = _Region(prev.x0, prev.y0, max(prev.x1, r.x1), r.y1,
                              prev.row_rows + fixed_rows)
        else:
            out.append(_Region(r.x0, r.y0, r.x1, r.y1, list(r.row_rows)))
    return out


def _looks_like_prose(s):
    return bool(_LEAK_RE.search(s)) or len(s.split()) >= 8


# Deliberately narrow anchors: each phrase only makes sense as the PRIMARY
# marker of one specific statement kind, not vocabulary that could plausibly
# turn up as a passing mention inside a different one. In particular, a
# generic line like "profit for the year" is NOT here even though it reads
# as P&L-shaped -- it routinely shows up as a perfectly normal ROW inside a
# real, unfused statement of changes in equity (the year's profit moving
# into retained earnings) and often as the opening reconciling line of a
# real cash flow statement, so using it as a P&L anchor would flag plenty
# of genuine, single-statement tables. "Cost of sales" / "gross profit" are
# P&L-INTERNAL subtotal concepts that don't appear as a row in any other
# statement kind, which is what makes them safe here.
_BS_ANCHOR_RE = re.compile(r"total assets\b|total liabilities\b|net assets\b", re.I)
_SOCE_ANCHOR_RE = re.compile(
    r"balance at \d|transactions with (the )?owners|"
    r"total comprehensive income for the year", re.I)
_PL_ANCHOR_RE = re.compile(r"cost of sales\b|gross profit\b", re.I)
_CF_ANCHOR_RE = re.compile(
    r"cash flows? from operating activities|"
    r"cash and cash equivalents at (the )?end", re.I)
_STATEMENT_ANCHORS = {
    "bs": _BS_ANCHOR_RE, "soce": _SOCE_ANCHOR_RE,
    "pl": _PL_ANCHOR_RE, "cf": _CF_ANCHOR_RE,
}


def _looks_like_two_fused_statements(rows):
    """True when a single img2table region carries primary-anchor vocabulary
    from two DIFFERENT statement kinds -- the signature of img2table's own
    borderless-table clustering having fused two unrelated, independently-
    complete tables on a 2-column landscape page. That fusion happens
    geometrically -- small x-gap, large y-overlap between the two source
    regions -- which is indistinguishable from a genuine labels-block /
    figures-block split of ONE table (see _merge_side_by_side); vocabulary
    is the signal that actually is specific to two independent statements.

    Originally balance-sheet + statement-of-changes-in-equity only (found
    live: en-2021-etisalat-group-annual-report.pdf p62, glueing "Share
    capital"/"Reserves" columns onto balance-sheet rows like "Goodwill and
    other intangible assets"). Extended to income statement + cash flow
    statement after the SAME fusion was found live gluing "Revenue"/"Cost
    of sales" together with "Cash flows from operating activities" row-by-
    row on du annual 2011.pdf p23 -- a real user manually marking just the
    visible income statement got the cash flow statement's rows zipped in
    beside it, because img2table's own region detection had already fused
    them into one native region before rows_in_box ever runs. Structured
    as a general "which kinds have a hit" check (not a hardcoded pair) so
    adding a further statement kind later is just one more anchor entry.

    Only counts a cell as a match when the anchor phrase sits at (or right
    at) the START of the cell, not buried inside a longer sentence: an
    accounting-policy note can mention "transactions with the owners" in a
    sentence without being anywhere near a real equity statement, and a
    first pass at this check (a 24-file, 2428-page sweep) found exactly
    that -- two accounting-policy pages that happened to use this phrase
    mid-sentence, both eliminated once such mid-string matches were
    excluded. A real statement row-label/section-header IS (up to leading
    whitespace) the anchor phrase -- "Total assets", "Balance at 1 January
    2020", "Cash flows from operating activities" -- while a sentence that
    merely references one has real words before the match.

    This used to gate on _looks_like_prose(c) instead of match position,
    which happened to work for the original BS/SOCE anchors (short enough,
    or broken up by a number, to duck _looks_like_prose's own word-count
    rule) but wrongly excluded the cash-flow anchors added alongside this
    docstring: "Cash flows from operating activities" is a perfectly
    normal 5-consecutive-word section-header cell, and 5 consecutive
    plain-English words of 3+ letters is exactly what _looks_like_prose's
    own LEAK_RE flags as prose -- found live keeping du annual 2011.pdf
    p23's income-statement/cash-flow fusion (see above) from being caught
    at all."""
    hits = set()
    for row in rows:
        for c in row:
            if not (isinstance(c, str) and c.strip()):
                continue
            stripped = c.strip()
            for kind, rx in _STATEMENT_ANCHORS.items():
                m = rx.search(stripped)
                if m and m.start() <= 2:
                    hits.add(kind)
        if len(hits) >= 2:
            return True
    return False


def _drop_prose_columns(rows):
    """img2table sometimes keeps a whole column of interleaved prose (an
    unrelated auditor's-report paragraph physically beside the statement) as
    its own column in the same region as the real label column. Drop any
    column that reads as prose on most of its non-empty cells, UNLESS
    dropping it would leave no LABEL column at all.

    That safeguard has to check for a real LABEL column, not just "any
    column with alphabetic content anywhere" -- two earlier, narrower
    versions of this check each let a real label column get dropped
    anyway, both found live:

    1. A still-unparsed figure cell is a Python str too ("12,951,414") at
       this point in the pipeline, so a naive "any non-empty string" check
       sees the numeric columns as perfectly good "text columns" and never
       trips (en-2022-1-eand-group-annual-report.pdf p50's cash flow
       statement: more than half the genuine, wrapped-across-several-lines
       labels are long enough to trip _looks_like_prose's own word-count
       rule, so the whole label column got marked for dropping, and the
       two figure columns sitting right next to it -- still strings --
       kept the safeguard from ever tripping).
    2. Even switching that check to "any ALPHABETIC content" isn't enough:
       a purely-numeric column's own HEADER cell often carries letters too
       ("Notes", "AED'000 2019") -- one such cell is enough to make an
       otherwise all-numeric column "count" as a surviving label column
       under a bare `any(...)`, even once the real label column (with
       alphabetic text on most of its OWN non-empty cells, not just its
       header) was the one just dropped (etisalat-group-annual-report-
       english-2019.pdf p46's right-half statement: labels correctly
       captured by img2table were wiped out this way, leaving a table of
       bare figures with a stray "Notes" cell as its only surviving text).

    See _column_reads_as_labels for the fraction-of-THIS-COLUMN'S-OWN-
    content check that replaces both -- deliberately NOT reusing
    _has_label_column's fraction-of-ALL-ROWS threshold (calibrated for
    THAT function's own job of ranking whole candidate regions, with an
    absolute "at least 3 hits" floor that a small table can never clear
    even when every single one of its rows has a real label)."""
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
    label_cols_left = sum(1 for c in range(ncols)
                          if c not in drop and _column_reads_as_labels(rows, c))
    if not drop or label_cols_left == 0:
        return rows
    keep = [c for c in range(ncols) if c not in drop]
    return [[r[c] if c < len(r) else None for c in keep] for r in rows]


def _column_reads_as_labels(rows, c):
    """True when column `c` carries real label text on a healthy fraction
    of ITS OWN non-empty cells -- not a numeric column whose only
    alphabetic content is a single header cell ("Notes", "AED'000 2019")
    sitting above many purely-numeric data cells. Needs at least 2 non-
    empty cells to judge at all (a single cell, even a real label, isn't
    enough signal to call this a label column with any confidence)."""
    cells = [r[c] for r in rows if c < len(r) and r[c] is not None and str(r[c]).strip()]
    if len(cells) < 2:
        return False
    hits = sum(1 for v in cells if isinstance(v, str) and re.search(r"[A-Za-z]{3,}", v))
    return hits >= max(2, 0.3 * len(cells))


def _has_label_column(rows):
    if not rows:
        return False
    hits = sum(1 for r in rows if r and isinstance(r[0], str)
              and re.search(r"[A-Za-z]{3,}", r[0]))
    return hits >= max(3, 0.3 * len(rows))


def _best_matching_region(page_tables, x0, y0, x1, y1, min_overlap_frac=0.5):
    """The candidate-matching/merge logic shared by `rows_in_box` (the
    user-drawn-box extraction path) and `grid_in_box` (the detect-only grid
    preview path) -- pulled out so both pick the SAME region for the same
    box, rather than two independently-written selection rules drifting
    apart. See `rows_in_box`'s own docstring for why the overlap check goes
    both ways (region-coverage and box-coverage, whichever is more
    generous). Returns the best-matching, already fully merged `_Region`,
    or None if nothing clears the overlap bar."""
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
    return cands[0]


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
    best = _best_matching_region(page_tables, x0, y0, x1, y1, min_overlap_frac)
    if best is None:
        return None

    # Index-based, not "for top, bot, vals in ... : kept.append((top, bot,
    # vals))" -- the latter would silently rebuild each row as a plain
    # 3-tuple, throwing away a _RowRow's own .raw_cells (see _RowRow) even
    # though nothing here actually needs to change per-row content, only
    # decide which rows survive.
    kept = [row for row in best.row_rows
            if row[1] >= y0 - row_margin and row[0] <= y1 + row_margin]
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


def _cluster_edges_into_bands(edges, tol):
    """Sort a flat list of cell-edge x-positions and merge any two within
    `tol` points of each other into one boundary -- the same "sort, then
    merge adjacent-within-tolerance" idiom already used for
    `_merge_stacked`'s own `x_tol`. Consecutive boundaries become column
    bands: N boundaries -> N-1 bands. Adjacent real cells share almost the
    same edge value (one cell's x2 sits right next to the next cell's x1),
    which is exactly what collapses into a single shared boundary here."""
    if not edges:
        return []
    edges = sorted(edges)
    clusters = [[edges[0]]]
    for e in edges[1:]:
        if e - clusters[-1][-1] <= tol:
            clusters[-1].append(e)
        else:
            clusters.append([e])
    boundaries = [sum(c) / len(c) for c in clusters]
    return [(boundaries[i], boundaries[i + 1]) for i in range(len(boundaries) - 1)]


def grid_in_box(page_tables, x0, y0, x1, y1, min_overlap_frac=0.5, row_margin=20,
                col_margin=20, col_tol=8):
    """Detect-only counterpart to `rows_in_box`, for the grid-preview
    feature: returns the DETECTED row/column line positions for whichever
    region matches the box -- no repair pipeline, no note-ref stripping, no
    prose-column dropping, no `_clean`/`analyze`/health pass. Shares
    `_best_matching_region` with `rows_in_box` so the preview and the
    eventual extraction always agree on which region is "the" match for a
    given box, rather than two independently-written selection rules
    drifting apart.

    Returns (row_bands, col_bands, bbox) or None. `row_bands` is exactly
    `rows_in_box`'s own row_bands -- already computed there today, just
    never returned to a caller that only wants detection. `col_bands` is
    derived by clustering every KEPT row's own raw_cells (see _RowRow)
    x-edges together across the whole matched region, not just one
    "template" row -- more robust to any single row's boundaries being
    slightly off, in the same spirit as `_has_label_column`'s own "a few
    hits, not just one" tolerance -- then trimmed to the box's OWN X-range
    (plus `col_margin` slack, mirroring `row_margin`), the same way
    `rows_in_box` already trims rows to the box's Y-range: a column
    entirely outside the drawn box shouldn't appear in the preview, which
    is what makes narrowing the box narrow the detected columns -- the
    concrete gap this whole feature exists to close (confirmed live: before
    this trim, `col_bands` came back identical regardless of how narrow the
    box was drawn). Can come back EMPTY (a graceful partial result: still
    get row lines) when no kept row ever carried raw_cells at all -- a
    region matched via a code path that doesn't have per-cell geometry (a
    synthetic region in a test, or a stale in-process cache entry from
    before this capability existed), not a failure to detect the table
    itself."""
    best = _best_matching_region(page_tables, x0, y0, x1, y1, min_overlap_frac)
    if best is None:
        return None

    kept = [row for row in best.row_rows
            if row[1] >= y0 - row_margin and row[0] <= y1 + row_margin]
    if not kept:
        kept = best.row_rows

    row_bands = [(row[0], row[1]) for row in kept]
    tight_y0 = min(top for top, _ in row_bands)
    tight_y1 = max(bot for _, bot in row_bands)

    edges = []
    for row in kept:
        raw = getattr(row, "raw_cells", None)
        if not raw:
            continue
        for cx1, cx2, _ in raw:
            edges.append(cx1)
            edges.append(cx2)
    all_col_bands = _cluster_edges_into_bands(edges, col_tol)
    # Mirror row_bands' own Y-trim: keep only column bands that actually
    # overlap the user's drawn box X-range (plus col_margin slack for a box
    # drawn a few points short of a column's true edge, same reasoning as
    # row_margin). A column entirely outside the box -- a neighbouring
    # table's own columns, or a year the user's box deliberately didn't
    # reach -- shouldn't show up in the preview grid at all; this is what
    # actually makes narrowing the box narrow the detected columns, the
    # concrete gap this whole feature exists to close.
    col_bands = [(bx1, bx2) for bx1, bx2 in all_col_bands
                if bx2 >= x0 - col_margin and bx1 <= x1 + col_margin]
    if not col_bands:
        col_bands = all_col_bands   # trimming would leave nothing -- don't blank the result

    return row_bands, col_bands, (best.x0, tight_y0, best.x1, tight_y1)


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
