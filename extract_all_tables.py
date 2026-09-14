"""
extract_all_tables.py  --  generic "dump every table in a PDF to Excel"

Not targeted at any company or statement.  For each page it detects every
table-like region it can (ruled tables first; then text-alignment tables on
pages where ruling finds little), cleans it lightly, guesses a title from the
nearest heading above it, and writes one worksheet per table plus a Contents
index.

Digital PDFs only (needs a real text layer -- no OCR).

    python extract_all_tables.py report.pdf
    python extract_all_tables.py  a.pdf  b.pdf  folder/
    python extract_all_tables.py report.pdf --out tables.xlsx --min-rows 2 --min-cols 2
"""
from __future__ import annotations

import argparse
import logging
import re
import sys
from pathlib import Path

LOG = logging.getLogger("tablekit.extract")   # quiet by default; -v / --debug turns it on

try:  # make console output robust on code-page-limited Windows terminals
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[union-attr]
except Exception:
    pass

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------- config ---
# CONFIG + its TOML override loader live in tablekit/config.py; parse_number /
# cell coercion in tablekit/parse.py.  Imported here so `import extract_all_tables`
# keeps exposing all of them -- parse_number itself isn't called from
# anywhere else IN this file (hence the noqa: F401 below), but
# tablekit_tests/test_golden.py calls it as X.parse_number(), so it's a
# deliberate re-export, not dead weight.
sys.path.insert(0, str(Path(__file__).resolve().parent))
from tablekit.config import CONFIG, load_config_overrides   # noqa: E402
from tablekit.parse import (parse_number, coerce_cell as _cell,  # noqa: E402,F401
                            normspace as _txt, NUM_RE as _NUM_RE, FormattedNumber)

load_config_overrides()

# The geometry reconstruction from the targeted extractor -- pure layout logic,
# no company vocabulary. Used to recover borderless / landscape-2-up statements
# that pdfplumber's own table finder cannot see. Optional: the script still
# runs (with weaker coverage on unruled statements) if it isn't importable.
try:
    import telecom_extract as _te
    HAVE_RECON = True
except Exception:
    _te = None  # type: ignore[assignment]  # optional dependency: absent-module sentinel
    HAVE_RECON = False

# Second, independent geometry strategy for the same borderless/2-up statements
# (OpenCV-based instead of regex-based) -- see tablekit/img2table_backend.py
# for why it exists and its known strengths/weaknesses. Optional: falls back
# to regex-only reconstruction if img2table isn't installed.
try:
    from tablekit.img2table_backend import (
        img2table_page_tables, rows_under_heading, rows_in_box, HAVE_IMG2TABLE,
        ocr_rows_in_box, HAVE_OCR)
except Exception:
    HAVE_IMG2TABLE = False
    HAVE_OCR = False

    # `*a, **k` deliberately -- these are inert stand-ins, never meant to
    # validate call shape the way the real functions do (mypy's "conditional
    # function variants must have identical signatures" is about exactly
    # that mismatch, which is the point here, not a bug).
    def img2table_page_tables(*a, **k):  # type: ignore[misc]
        return []

    def ocr_rows_in_box(*a, **k):  # type: ignore[misc]
        return None

    def rows_under_heading(*a, **k):  # type: ignore[misc]
        return None

    def rows_in_box(*a, **k):  # type: ignore[misc]
        return None

# ----------------------------------------------------------------- detection ---
LINES = {"vertical_strategy": "lines", "horizontal_strategy": "lines",
         "snap_tolerance": 4, "join_tolerance": 4, "intersection_tolerance": 4}
LINES_H = {"vertical_strategy": "text", "horizontal_strategy": "lines",
           "snap_tolerance": 4}
TEXT = {"vertical_strategy": "text", "horizontal_strategy": "text",
        "min_words_vertical": 2, "min_words_horizontal": 1,
        "snap_tolerance": 6, "keep_blank_chars": False}


def _overlap(a, b) -> float:
    """area(intersection) / area(smaller)  for two (x0,y0,x1,y1) boxes."""
    ix0, iy0 = max(a[0], b[0]), max(a[1], b[1])
    ix1, iy1 = min(a[2], b[2]), min(a[3], b[3])
    if ix1 <= ix0 or iy1 <= iy0:
        return 0.0
    inter = (ix1 - ix0) * (iy1 - iy0)
    sa = (a[2] - a[0]) * (a[3] - a[1])
    sb = (b[2] - b[0]) * (b[3] - b[1])
    return inter / max(1.0, min(sa, sb))


_RECON_HEAD_RE = re.compile(
    r"statement of (profit or loss|comprehensive income|financial position|"
    r"cash flows?|changes in equity)|consolidated income statement|"
    r"\bincome statement\b", re.I)

_ANCHOR_WORDS = (
    "revenue", "turnover", "total income", "operating income", "gross profit",
    "cost of sales", "operating expenses", "operating profit", "finance income",
    "finance cost", "profit before", "profit for the year", "net profit",
    "total assets", "total equity", "total liabilities", "net assets",
    "total non-current assets", "total current assets", "share capital",
    "retained earnings", "cash flows from operating", "net cash",
    "cash and cash equivalents", "financing activities", "investing activities",
    "adjustments for", "balance at", "as at 1 january", "as at 31 december",
    "total comprehensive income", "other comprehensive income")


def _title_from_content(labs):
    """Best statement name from the row labels themselves."""
    if ("profit for the year" in labs or "net profit for the year" in labs) and \
       ("revenue" in labs or "total income" in labs or "operating expenses" in labs
        or "cost of sales" in labs):
        return ("Statement of comprehensive income"
                if "other comprehensive income" in labs
                else "Statement of profit or loss")
    if "total assets" in labs or ("net assets" in labs and "total equity" in labs) \
       or ("total non-current assets" in labs and "total current assets" in labs):
        return "Statement of financial position"
    if "cash flows from operating" in labs or "net cash" in labs:
        return "Statement of cash flows"
    if "balance at" in labs or "as at 1 january" in labs:
        return "Statement of changes in equity"
    return None


def _candidate_quality(rows, title):
    """Score a raw candidate reconstruction (before it's committed to) by
    running the REAL semantic/reconciliation check on it.  Used to pick a
    winner between two competing geometry strategies for the same heading.
    Higher is better: (foots_rank, label+figure health, row count)."""
    try:
        t = {"rows": rows, "title": title, "file": "", "page_label": 0,
             "shape": classify(rows)}
        analyze(t)
        _attach_health(t)
    except Exception:
        return (-1, 0.0, 0)
    foots_rank = {True: 2, None: 1, False: 0}.get(t.get("foots"), 1)
    health = (t.get("health") or {}).get("score", 0.0)
    return (foots_rank, health, len(rows))


def _recon_tables(page, pdf_path=None):
    """Recover the primary financial statements that pdfplumber's own table
    finder can't see (borderless, or two side-by-side on a landscape page),
    one reconstruction per statement heading. Returns [(bbox, rows, title)].

    Two independent geometry strategies compete for each heading -- the
    regex-based reconstruction (`_te.reconstruct_page_statement`) and, when
    available, img2table's OpenCV-based borderless-table detector cropped to
    the same region. Whichever one actually reconciles (or, if neither does,
    has cleaner labels) wins; see tablekit/img2table_backend.py for why both
    are kept rather than picking one permanently."""
    if not HAVE_RECON:
        return []
    out = []
    mid = page.width / 2
    two_up = page.width > CONFIG["two_up_page_width"]
    page_tables = (img2table_page_tables(pdf_path, page.page_number - 1)
                   if HAVE_IMG2TABLE and pdf_path is not None else [])
    try:
        heads = []
        _lines = _te.build_lines(page.extract_words())
        for i, ln in enumerate(_lines):
            txt = ln["text"]
            # stitch a heading that wrapped onto a second line, e.g.
            # "Consolidated statement" + "of cash flows" (du's old landscape 2-up)
            if re.search(r"\bstatement\b", txt, re.I) and not _RECON_HEAD_RE.search(txt):
                for nx in _lines[i + 1:i + 4]:
                    if (0 < nx["top"] - ln["top"] <= 30
                            and abs(nx["x0"] - ln["x0"]) < 60
                            and re.match(r"^(of|and|&)\b", nx["text"].strip(), re.I)):
                        txt = txt + " " + nx["text"]
                        break
            m = _RECON_HEAD_RE.search(txt)
            if not m or m.start() > CONFIG["recon_head_max_start"] or len(txt) > CONFIG["recon_head_max_len"]:
                continue
            if any(abs(ln["top"] - h["top"]) < 18 and abs(ln["x0"] - h["x0"]) < 40
                   for h in heads):
                continue
            heads.append({**ln, "text": txt})
        for ln in heads:
            raw = _te.reconstruct_page_statement(page, ln["top"], ln["x0"]) or []
            try:                    # split "6 1,569,189" -> note "6" + 1569189, etc.
                raw = _te._strip_note_refs(raw)
            except Exception:
                pass
            pr = _clean(raw)
            if not pr or len(pr) < 4:
                continue
            # the reconstruction must actually look like a statement, not a
            # policy paragraph that merely mentions one
            labs = " || ".join(_row_label(r).lower() for r in pr)
            if sum(1 for a in _ANCHOR_WORDS if a in labs) < CONFIG["recon_min_anchor_words"]:
                continue
            # name it from its own content (2-up heading grabs are unreliable)
            title = _title_from_content(labs) or _txt(
                ln["text"][_RECON_HEAD_RE.search(ln["text"]).start():])[:120]
            if two_up:
                # confine the bbox to the heading's own half of the page so a
                # side-by-side statement isn't deduped away
                if ln["x0"] < mid:
                    bx0, bx1 = 0.0, mid + 20
                else:
                    bx0, bx1 = mid - 20, page.width
            else:
                bx0, bx1 = 0.0, page.width
            bbox = (bx0, ln["top"], bx1, page.height)

            # challenger: img2table's own table region(s) under this heading,
            # matched by position (not a fixed crop -- see img2table_backend).
            #
            # SAFETY: only even consider it when the regex candidate does NOT
            # already reconcile.  Letting it compete unconditionally was tried
            # and reverted -- "foots" passing is necessary but not sufficient
            # evidence of correctness (a truncated/wrong table can coincide
            # with a balance identity), so on reports that already worked it
            # occasionally picked a *worse* but coincidentally-passing
            # img2table candidate over an already-correct one (verified
            # against the hand-checked anchors -- see CHANGELOG). Restricting
            # it to "pr doesn't reconcile" makes this strictly additive: every
            # report already proven correct is untouched; only genuinely
            # broken candidates (du's old 2-up auditor-interleaved pages) are
            # ever replaced.
            best_rows, best_title = pr, title
            q1 = _candidate_quality(pr, title)
            if page_tables and q1[0] < 2:
                other_x0s = [h["x0"] for h in heads if h is not ln]
                raw2 = rows_under_heading(page_tables, ln["top"], ln["x0"],
                                          bx0, bx1, other_x0s)
                if raw2:
                    pr2 = _clean(raw2)
                    if pr2 and len(pr2) >= 4:
                        labs2 = " || ".join(_row_label(r).lower() for r in pr2)
                        if sum(1 for a in _ANCHOR_WORDS if a in labs2) >= CONFIG["recon_min_anchor_words"]:
                            title2 = _title_from_content(labs2) or title
                            q2 = _candidate_quality(pr2, title2)
                            # require the challenger to actually reconcile --
                            # not just "score higher" -- before trusting it
                            # over a candidate that doesn't foot at all
                            if q2[0] == 2 and q2 > q1:
                                best_rows, best_title = pr2, title2
            out.append((bbox, best_rows, best_title))
    except Exception:
        LOG.debug("reconstruction failed on p%s", getattr(page, 'page_number', '?'), exc_info=True)
    return out


def find_all_tables(page, pdf_path=None):
    """Every distinct table on the page, as (bbox, rows).  Runs pdfplumber's
    strategies plus the geometry reconstruction, and drops near-duplicates
    (ruled / reconstructed results outrank text-alignment guesses)."""
    found = []  # (bbox, rows, rank, title_hint)  lower rank = more trustworthy
    for bbox, rows, thint in _recon_tables(page, pdf_path):
        found.append((tuple(bbox), rows, 0, thint))
    for rank, settings in enumerate((LINES, LINES_H, TEXT), start=1):
        try:
            tables = page.find_tables(table_settings=settings)
        except Exception:
            tables = []
        for t in tables:
            try:
                rows = t.extract()
            except Exception:
                continue
            rows = _clean(rows)
            if not rows:
                continue
            found.append((tuple(t.bbox), rows, rank, None))

    found.sort(key=lambda f: (f[2], -(f[0][2] - f[0][0]) * (f[0][3] - f[0][1])))
    kept = []
    for bbox, rows, rank, thint in found:
        if any(_overlap(bbox, k[0]) > CONFIG["overlap_dedup"] for k in kept):
            continue
        # two heading anchors on a landscape 2-up page can reconstruct the same
        # rows from opposite sides -- drop the exact-content duplicate too
        if any(rows == k[1] for k in kept):
            continue
        kept.append((bbox, rows, thint))
    # a single reconstruction that swallowed two stacked statements
    # (du's old landscape 2-up) -- cut it at the second statement heading
    kept = [seg for (bbox, rows, thint) in kept
            for seg in _split_stacked_statements(bbox, rows, thint)]
    kept.sort(key=lambda k: (round(k[0][1]), k[0][0]))   # top-to-bottom, left-to-right
    return kept


def _looks_labelless(rows):
    """Mirrors img2table_backend._has_label_column's own threshold: true
    when column 0 is NOT carrying real label text on a healthy fraction of
    rows -- i.e. this "table" is just bare figures with nothing saying what
    they are."""
    if not rows:
        return True
    hits = sum(1 for r in rows if r and isinstance(r[0], str)
              and re.search(r"[A-Za-z]{3,}", r[0]))
    return hits < max(2, 0.3 * len(rows))


def _attach_left_labels(page, rows, row_bands, search_x0, region_x0):
    """Prepend a label column recovered from the page's own words, one row
    at a time: for each row's known Y-band (from the img2table region that
    supplied `rows`), take whatever text sits to the left of the matched
    numeric region but still inside the user's drawn box. `rows` and
    `row_bands` are the same length and in the same order (see
    `rows_in_box`'s docstring) -- if they aren't (e.g. a pdfplumber-fallback
    table with no row_bands at all), the caller never reaches here."""
    words = page.extract_words()
    out = []
    for i, row in enumerate(rows):
        if i >= len(row_bands):
            out.append(row)
            continue
        top, bot = row_bands[i]
        band = [w for w in words
               if top - 1 <= w["top"] <= bot + 1
               and search_x0 - 5 <= w["x0"] < region_x0]
        label = _txt(" ".join(w["text"] for w in sorted(band, key=lambda w: w["x0"])))
        out.append(([label] if label else [None]) + list(row))
    return out


def _finish_manual_table(pdf_path, page_index0, rows, shown_bbox, title):
    """Shared tail for extract_region / extract_region_ocr: same
    classify/analyze/health pipeline every automatically-detected table goes
    through, so foots verification, year detection and the editable-grid UI
    all work identically regardless of where the rows came from."""
    t = {
        "file": Path(pdf_path).name, "page_label": page_index0 + 1,
        "title": title, "rows": rows, "bbox": list(shown_bbox),
        "shape": classify(rows), "_manual": True,
    }
    analyze(t)
    _attach_health(t)
    return t


def box_has_text(pdf_path, page_index0, bbox):
    """True if the drawn box has ANY extractable text at all. Used to decide
    whether the OCR failsafe should even be offered -- OCR is for a box
    that's genuinely a scanned image, not a second attempt at a box whose
    text just didn't form a parseable table."""
    x0, y0, x1, y1 = bbox
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[page_index0]
        cx0, cy0 = max(0.0, x0), max(0.0, y0)
        cx1, cy1 = min(page.width, x1), min(page.height, y1)
        if cx1 <= cx0 or cy1 <= cy0:
            return False
        return bool((page.crop((cx0, cy0, cx1, cy1)).extract_text() or "").strip())


def extract_region(pdf_path, page_index0, bbox, title=None):
    """Manual-selection counterpart to the automatic heading-based scan: the
    user has already drawn a box around the exact table they want, so there's
    no heading to find or disambiguate against -- just pull whatever's under
    the box and run it through the SAME classify/analyze/health pipeline
    every automatically-detected table goes through, so foots verification,
    year detection and the editable-grid UI all work identically either way.
    Returns a table dict, or None if nothing table-like was found in the box
    (the caller can then check `box_has_text` to decide whether to offer the
    OCR failsafe -- see `extract_region_ocr`).
    """
    x0, y0, x1, y1 = bbox
    raw = None
    matched_bbox = None
    row_bands = None
    if HAVE_IMG2TABLE:
        page_tables = img2table_page_tables(pdf_path, page_index0)
        hit = rows_in_box(page_tables, x0, y0, x1, y1)
        if hit:
            raw, matched_bbox, row_bands = hit
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[page_index0]
        if not raw:
            try:
                raw = page.crop((x0, y0, x1, y1)).extract_table()
            except Exception:
                raw = None
        if not raw:
            return None
        # a ruled box that wraps only the number columns -- with row labels
        # sitting outside it, unruled -- is invisible to img2table's own
        # structure detection; it only ever sees the ruled numeric grid, so
        # the "table" it hands back has real figures and no idea what any
        # of them are. Recover the labels the same way telecom_extract's
        # attach_left_labels() does: search each row's own vertical band,
        # to the left of the matched region, for the text that belongs to it.
        if row_bands and matched_bbox and _looks_labelless(raw):
            raw = _attach_left_labels(page, raw, row_bands, x0, matched_bbox[0])
        rows = _clean(raw)
        if not rows:
            return None
        if not title:
            guessed = guess_title(page, (x0, y0, x1, y1), page.extract_words()) or ""
            # a box drawn right at (or just below) the table's own heading
            # can leave no room above it for guess_title to find -- it then
            # falls back to whatever's above that, typically a fragment of
            # the page's running header or subtitle (NOT necessarily one
            # guess_title itself flags as a run-header -- x-span narrowing
            # can crop a run-header phrase down to an innocuous-looking
            # leftover word). When the table's own first row already reads
            # like a real heading (a numbered note, or a statement title),
            # prefer that UNLESS guessed is itself already a confident,
            # unambiguous heading match.
            first_label = _row_label(rows[0]) if rows else ""
            first_is_heading = bool(_NOTE_HEAD_RE.match(first_label) or _STMT_RE.search(first_label))
            guessed_is_confident = bool(_NOTE_HEAD_RE.match(guessed) or _STMT_RE.search(guessed))
            if first_is_heading and not guessed_is_confident:
                title = first_label
            else:
                title = guessed
    # show the box that actually matches what was extracted -- rows_in_box
    # already trims its match down to the rows that survive (see its own
    # docstring for why: a generously-drawn box shouldn't drag a neighbouring
    # note's heading into the result OR the highlight), so its bbox can be
    # trusted directly. Only the pdfplumber fallback (no per-row geometry
    # available) falls back to the drawn box itself.
    shown_bbox = list(matched_bbox) if matched_bbox else [x0, y0, x1, y1]
    return _finish_manual_table(pdf_path, page_index0, rows, shown_bbox, title)


def extract_region_ocr(pdf_path, page_index0, bbox, title=None):
    """OCR failsafe for extract_region: only ever called explicitly (the
    user clicking a dedicated "OCR this region" action) after a normal
    extraction attempt on this same box came back empty AND the box's own
    text layer is genuinely empty (see `box_has_text`) -- this never runs
    automatically and never re-tries a box that already has real text.
    Renders and OCRs only the drawn box, not the whole page, so this stays a
    single click, not a background scan. Returns a table dict, or None."""
    if not HAVE_OCR:
        return None
    x0, y0, x1, y1 = bbox
    hit = ocr_rows_in_box(pdf_path, page_index0, x0, y0, x1, y1)
    if not hit:
        return None
    raw, matched_bbox, _row_bands_unused = hit
    rows = _clean(raw)
    if not rows:
        return None
    if not title:
        first_label = _row_label(rows[0]) if rows else ""
        title = first_label if _NOTE_HEAD_RE.match(first_label) or _STMT_RE.search(first_label) \
            else "OCR region"
    t = _finish_manual_table(pdf_path, page_index0, rows, matched_bbox, title)
    # flag OCR-derived tables distinctly -- digit misreads (0/8/6, 1/7...) are
    # a real risk, worth a visible "verify this by eye" cue the UI can show
    # that a real-text extraction doesn't need.
    t["_ocr"] = True
    return t


def _split_stacked_statements(bbox, rows, thint):
    """If `rows` contains a second statement starting partway down (a
    'Cash flows from operating activities' / 'statement of ...' row well below
    the top), split into two tables so each can be typed and footed alone."""
    cut = None
    for i, r in enumerate(rows):
        if i < 5 or i > len(rows) - 4:
            continue
        lb = _row_label(r).lower().strip()
        # distinctive start of a cash-flow statement, or a row that is itself
        # just a statement heading -- never a mid-statement line item
        if (lb.startswith("cash flows from operating activities")
                or lb.startswith("cash flow from operating activities")
                or (len(lb) <= 60 and re.match(
                    r"^(consolidated |group )?statement of "
                    r"(cash flows?|comprehensive income|profit or loss|"
                    r"financial position|changes in equity)\b", lb))):
            cut = i
            break
    if cut is None:
        return [(bbox, rows, thint)]
    x0, y0, x1, y1 = bbox
    mid = y0 + (y1 - y0) * cut / max(len(rows), 1)
    top = [r for r in rows[:cut] if _row_label(r).strip()]
    bot = [r for r in rows[cut:] if _row_label(r).strip()]
    out = []
    if len(top) >= 4:
        t_labs = " || ".join(_row_label(r).lower() for r in top)
        out.append(((x0, y0, x1, mid), top, _title_from_content(t_labs) or thint))
    if len(bot) >= 4:
        b_labs = " || ".join(_row_label(r).lower() for r in bot)
        out.append(((x0, mid, x1, y1), bot, _title_from_content(b_labs)))
    return out or [(bbox, rows, thint)]


# ------------------------------------------------------------------- cleanup ---
# _txt / parse_number / _cell / _NUM_RE come from tablekit.parse (imported at
# the top).  Referenced by those names throughout this file and the tests.


_GLUED_SECT_RE = re.compile(
    r"^(?P<keep>(total\b[a-z\s/-]*?(assets|liabilities|equity)|net (current )?assets))\s+"
    r"(?P<tail>(current|non[- ]?current|equity|liabilities|assets|represented)\b.*)$", re.I)


def _desect_labels(rows):
    """A borderless reconstruction often welds the next section header onto a
    'Total ... assets/liabilities/equity' line ('Total non-current assets
    Current a...').  Trim the trailing header fragment; the figures belong to
    the total and are kept."""
    out = []
    for r in rows:
        if r and isinstance(r[0], str):
            m = _GLUED_SECT_RE.match(r[0].strip())
            if m and len(m.group("tail").split()) <= 4:
                r = [m.group("keep").strip()] + list(r[1:])
        out.append(r)
    return out


# a sentence out of the auditor's / directors' report, interleaved column-wise
# with an old landscape 2-up statement (du 2010-2014)
_AUDITOR_PROSE_RE = re.compile(
    r"we have audited|in our opinion|the consolidated financial statements|"
    r"material misstatement|reasonable assurance|internal control|"
    r"responsibilit|going concern|board of directors|signed on|"
    r"pricewaterhousecoopers|\bkpmg\b|\bdeloitte\b|ernst\s*&\s*young|\bey\b|"
    r"registered auditor|audit (evidence|opinion|procedures)|"
    r"united arab emirates federal law|articles of association|"
    r"whose report dated|expressed an unqualified", re.I)


def _deprose_labels(rows):
    """When a label carries a chunk of interleaved auditor's-report prose
    *and* a recognisable statement line at its head or tail, keep only the
    statement line.  Pure-prose labels (real line lost) are left for
    label_health to flag."""
    # multi-word balance-sheet / P&L line items, to spot one buried mid-prose
    _MIDLINE = re.compile(
        r"\b((?:total )?(?:non-current|current) (?:assets|liabilities)"
        r"|trade (?:and other )?(?:receivables|payables)"
        r"|due (?:from|to) (?:a )?related part(?:y|ies)"
        r"|property,? plant and equipment|cash and (?:cash equivalents|bank balances)"
        r"|share (?:capital|premium)|retained earnings|other reserves"
        r"|(?:total )?(?:non-current )?(?:current )?"
        r"(?:borrowings|provisions|inventories|deferred (?:tax|fees|revenue))"
        r"|employee benefits|net current assets)\b", re.I)
    out = []
    for r in rows:
        lb = r[0] if r and isinstance(r[0], str) else None
        has_fig = any(isinstance(c, (int, float)) and abs(c) >= 100 for c in r[1:])
        if lb and _AUDITOR_PROSE_RE.search(lb) and len(lb.split()) >= 6:
            keep = None
            # (a) a "Total ... / Net ..." statement line at the very end
            m = re.search(r"(?:^|\s)((?:total|net)\b[\w\s'’/()-]{2,45}"
                          r"(?:assets|liabilities|equity|current assets|"
                          r"current liabilities|profit|income|expenses))\s*$", lb, re.I)
            if m:
                keep = m.group(1)
            # (b) a recognised line item as the trailing Title-Case phrase
            if keep is None:
                m = re.search(r"([A-Z][A-Za-z][\w\s'’/&()-]{2,40})\s*$", lb)
                if m and _KNOWN_LINE_RE.match(m.group(1).strip()):
                    keep = m.group(1)
            # (c) a "Total ... / Share ... / Retained ..." line at the very start
            if keep is None:
                m = re.match(r"^((?:total|net|gross|share|retained|current|non-current)"
                             r"[\w\s'’/()-]{2,45}?)(?:\s+[A-Z][a-z]+.*|$)", lb)
                if m:
                    keep = m.group(1)
            # (d) the row carries figures and a known multi-word line item sits
            #     somewhere in the middle of the sentence -- pull it out
            if keep is None and has_fig:
                m = _MIDLINE.search(lb)
                if m:
                    keep = m.group(1)
            if keep and 1 <= len(keep.split()) <= 7:
                r = [keep.strip()] + list(r[1:])
        out.append(r)
    return out


def _clean(rows):
    if not rows:
        return []
    rows = [[_cell(c) for c in r] for r in rows]
    # drop fully-empty rows
    rows = [r for r in rows if any(c is not None and str(c).strip() != "" for c in r)]
    if not rows:
        return []
    width = max(len(r) for r in rows)
    rows = [r + [None] * (width - len(r)) for r in rows]
    # drop fully-empty columns
    keep = [c for c in range(width)
            if any(rows[r][c] is not None and str(rows[r][c]).strip() != ""
                   for r in range(len(rows)))]
    rows = [[r[c] for c in keep] for r in rows]
    rows = _desect_labels(rows)
    rows = _deprose_labels(rows)
    return rows if (rows and rows[0]) else []


# --------------------------------------------------------------------- title ---
_HEAD_RE = re.compile(r"[A-Za-z]")
_SUBTITLE_RE = re.compile(
    r"^(for the (year|period)|as at|as of|31 december|note[s]?\b|aed|usd|"
    r"\(?continued\)?|in thousands|the accompanying|\d{4}\b)", re.I)


def _looks_heading(text):
    t = text.strip()
    if len(t) < 3 or len(t) > 110:
        return False
    if not _HEAD_RE.search(t):
        return False
    if _SUBTITLE_RE.match(t):
        return False
    # a heading isn't mostly digits / a full sentence of prose
    words = t.split()
    if len(words) > 14:
        return False
    digits = sum(c.isdigit() for c in t)
    return digits < len(t) * 0.4


_RUNHDR_RE = re.compile(
    r"strategic report|corporate governance|financial statements|"
    r"annual report|integrated report|contents|overview|sustainability", re.I)
_STMT_RE = re.compile(
    r"((consolidated|group)\s+)?statement of (profit or loss|comprehensive income|"
    r"financial position|cash flows?|changes in equity)"
    r"|consolidated income statement|\bincome statement\b|\bbalance sheet\b", re.I)
# a numbered note heading ("8 Intangible assets", "8.2 Property, plant and
# equipment") -- kept in sync with _statement_kind's own check of the title
_NOTE_HEAD_RE = re.compile(r"^\s*\d{1,2}(\.\d{1,2})*[.\)]?\s+[A-Z]")


def guess_title(page, bbox, words_cache):
    """Walk upward from the table for the closest 1-2 lines that read like a
    heading (skip subtitles like 'for the year ended ...', number rows, prose).
    Prefer a 'statement of ...' phrase anywhere in the band over a running
    page header."""
    x0, y0, x1, y1 = bbox
    # explicit statement heading over this table's x-span, searched a bit
    # higher (these titles often sit well above a borderless statement)
    band = [w for w in words_cache if y0 - 170 <= w["bottom"] <= y0 + 6
            and x0 - 50 <= (w["x0"] + w["x1"]) / 2 <= x1 + 50]
    if band:
        blines = {}
        for w in band:
            blines.setdefault(round(w["top"]), []).append(w)
        tops = sorted(blines)
        for k, top in enumerate(tops):
            txt = _txt(" ".join(
                w["text"] for w in sorted(blines[top], key=lambda w: w["x0"])))
            m = _STMT_RE.search(txt)
            if m:
                # stitch a wrapped heading ("Consolidated statement of" / "cash flows")
                nxt = ""
                if k + 1 < len(tops):
                    nxt = _txt(" ".join(
                        w["text"] for w in sorted(blines[tops[k + 1]], key=lambda w: w["x0"])))
                cand = (m.group(0) if m.start() > 0 else txt)
                if nxt and not _STMT_RE.search(nxt) and len(nxt.split()) <= 5 \
                        and not _NUM_RE.match((nxt.split() or ["x"])[0]) \
                        and not _RUNHDR_RE.search(nxt):
                    cand = (cand + " " + nxt).strip()
                return cand[:120]
        # no statement heading in the band -- try a numbered note heading
        # instead ("8 Intangible assets"). Searched top-to-bottom (furthest
        # from the table first) so the actual SECTION heading wins over
        # something that happens to start with a number but sits right next
        # to the table -- e.g. the table's own wrapped column-header row
        # ("Software  Capital work  ...  Total") passes the generic
        # heading-looks-like check below and, being closest, used to win.
        for top in tops:
            txt = _txt(" ".join(
                w["text"] for w in sorted(blines[top], key=lambda w: w["x0"])))
            if _NOTE_HEAD_RE.match(txt) and not _SUBTITLE_RE.match(txt) \
                    and not _RUNHDR_RE.search(txt):
                return txt[:120]
    cands = [w for w in words_cache
             if y0 - 130 <= w["bottom"] <= y0 + 2
             and w["x1"] > x0 - 60 and w["x0"] < x1 + 60]
    if not cands:
        return ""
    lines = {}
    for w in cands:
        lines.setdefault(round(w["top"]), []).append(w)
    ordered = []
    for top in sorted(lines):
        # On a 2-up page the heading line is shared by two statements; keep
        # only the words that sit over THIS table's own x-span (a bit wider).
        lw = [w for w in lines[top] if x0 - 45 <= (w["x0"] + w["x1"]) / 2 <= x1 + 45]
        lw = lw or lines[top]
        txt = _txt(" ".join(w["text"] for w in sorted(lw, key=lambda w: w["x0"])))
        ordered.append(txt)
    # nearest heading-looking line, optionally prefixed by the line above it
    # when that also looks like a heading (wrapped title)
    for i in range(len(ordered) - 1, -1, -1):
        if _looks_heading(ordered[i]) and not _RUNHDR_RE.search(ordered[i].strip()):
            title = ordered[i]
            if i > 0 and _looks_heading(ordered[i - 1]) and len(ordered[i - 1].split()) <= 6:
                title = ordered[i - 1] + " " + title
            return title[:120]
    return (ordered[-1] if ordered else "")[:120]


def _trim_trailing_prose(rows):
    """Drop trailing rows that are a sentence of running text with no figures
    (auditor sign-off blocks, 'The notes on pages ... form an integral part')."""
    last_num = -1
    for i, r in enumerate(rows):
        if any(isinstance(c, (int, float)) for c in r):
            last_num = i
    if last_num < 0:
        return rows
    out = rows[:last_num + 1]
    # keep an immediately-following short unlabelled total row if present
    for r in rows[last_num + 1:]:
        txt = " ".join(str(c) for c in r if c is not None)
        if len(txt) <= 40 and not re.search(r"[a-z]{4,}\s+[a-z]{4,}\s+[a-z]{4,}", txt.lower()):
            out.append(r)
        else:
            break
    return out


def classify(rows):
    """Rough kind hint for the inventory list."""
    body = rows[1:] if len(rows) > 1 else rows
    ncells = sum(len(r) for r in body) or 1
    nums = sum(1 for r in body for c in r if isinstance(c, (int, float)))
    longtext = sum(1 for r in body for c in r
                   if isinstance(c, str) and len(c) > 40)
    dens = nums / ncells
    labels = " ".join(str(c) for r in rows for c in r if isinstance(c, str)).lower()
    if any(k in labels for k in (
            "profit for the year", "total assets", "total equity",
            "cash flows from operating", "total comprehensive income",
            "net profit for the year", "total income")):
        return "statement"
    if longtext > ncells * 0.35:
        return "mostly-text"
    if dens >= 0.25:
        return "table"
    return "small/other"


# ---------------------------------------------------------------- understand ---
# Give each detected table a light semantic reading: what statement it is,
# which columns are which year, which rows are totals, and -- where the shape
# allows an arithmetic check -- whether it foots.

_YEAR_RE = re.compile(r"(?:^|\D)(19|20)\d{2}(?:\D|$)")
_TOTAL_RE = re.compile(
    r"^\s*(total\b|net\b|gross (profit|margin)|operating (profit|income|expenses)\b"
    r"|(net\s+)?profit (for the|before|attributable)|loss (for the|before)"
    r"|profit/\(loss\)|earnings before|ebitda|comprehensive income for the year"
    r"|cash (generated|used|and cash equivalents)|net cash)", re.I)
# stricter: rows that are UNAMBIGUOUSLY a total (used by the footing check,
# where a plain "Operating expenses" line may itself be a leaf, e.g. du's P&L)
_HARD_TOTAL_RE = re.compile(
    r"^\s*(total\b|gross (profit|margin)|(net\s+)?(profit|loss) (for the|before)"
    r"|profit/\(loss\) before|operating profit\b|earnings before|ebitda"
    r"|comprehensive income for the year|net cash (generated|used|flows?))", re.I)
_LEAK_RE = re.compile(r"[a-z]{3,}\s+[a-z]{3,}\s+[a-z]{3,}\s+[a-z]{3,}\s+[a-z]{3,}", re.I)  # 5+ words = prose


def _row_label(r):
    for c in r:
        if isinstance(c, str) and re.search(r"[A-Za-z]{2,}", c):
            return c.strip()
    return ""


def _row_nums(r):
    return [c for c in r if isinstance(c, (int, float)) and not isinstance(c, bool)]


def _statement_kind(title, rows):
    t = (title or "").lower()
    labs = " || ".join(_row_label(r).lower() for r in rows)
    pl_content = ("revenue" in labs or "total income" in labs
                  or "profit for the year" in labs or "operating expenses" in labs
                  or "operating income" in labs)
    bs_content = ("total assets" in labs or "total equity" in labs
                  or "total liabilities" in labs or "net assets" in labs)
    cf_content = ("operating activities" in labs or "financing activities" in labs
                  or "investing activities" in labs)
    # a title alone isn't enough -- a mistitled restatement / transition grid
    # borrows a nearby statement heading, so require corroborating line items
    if (re.search(r"profit or loss|income statement|comprehensive income", t) and pl_content
            and not (bs_content and not pl_content)) or \
       ("profit for the year" in labs and ("revenue" in labs or "total income" in labs
                                           or "operating expenses" in labs)):
        return "income statement"
    if (re.search(r"financial position|balance sheet", t) and bs_content) or \
       ("total assets" in labs and "total equity" in labs):
        return "statement of financial position"
    if (re.search(r"cash flow", t) and cf_content) \
       or "cash flows from operating activities" in labs:
        return "statement of cash flows"
    if re.search(r"changes in equity", t) or "balance at 1 january" in labs \
       or "as at 1 january" in labs:
        return "statement of changes in equity"
    if re.match(r"^\s*\d{1,2}(\.\d{1,2})*[.\)]?\s+[A-Z]", title or ""):
        return "note"
    return "table"


_PLAUS_YEAR = lambda y: 1990 <= y <= 2035


def _years_in(text):
    out = []
    for m in re.finditer(r"\b(19|20)\d{2}\b", str(text)):
        y = int(m.group(0))
        if _PLAUS_YEAR(y) and y not in out:
            out.append(y)
    return out


def _add_note(t, key, en_text, **note_vars):
    """Append one note in TWO parallel forms, always kept in lockstep
    (same length, same order): `t["notes"]` -- plain English prose, exactly
    as before this existed, unchanged for the CLI/`--audit`/Excel export --
    and `t["notes_i18n"]` -- a {key, vars} pair the web UI can look up in
    its own translation table (see the comment above I18N in webui.html)
    and render in the active language, falling back to the English string
    when the key isn't one it recognises. Two lists instead of restructuring
    "notes" itself so nothing downstream that already reads plain English
    strings out of it has to change."""
    t.setdefault("notes", []).append(en_text)
    t.setdefault("notes_i18n", []).append({"key": key, "vars": note_vars})


def analyze(t, page_years=None, doc_years=None):
    """Attach a semantic reading to a table dict (in place); also return it."""
    rows = t["rows"]
    kind = _statement_kind(t.get("title"), rows)

    # header row = the first of the top rows that carries year / unit / "Notes"
    # tokens but no real figures.  Many reconstructed tables have NO such row
    # (data starts immediately) -- then there is no header to skip.
    header_idx = -1
    for i, r in enumerate(rows[:4]):
        joined = " ".join(str(c) for c in r if c is not None)
        nums = _row_nums(r)
        big = [v for v in nums if abs(v) >= 100]
        # a row whose only figures are plausible years and that carries no
        # label is itself the year header (['', 2024, 2023])
        year_only = (nums and all(_PLAUS_YEAR(int(v)) for v in nums)
                     and not _row_label(r))
        looks_hdr = (year_only or
                     (not big and (_years_in(joined)
                                   or re.search(r"AED|USD|EGP|SAR|['’]000|\bnotes?\b",
                                                joined, re.I))))
        if looks_hdr:
            header_idx = i
        elif nums:
            break
    data_start = header_idx + 1 if header_idx >= 0 else 0
    header_idx = max(header_idx, 0)
    years = _years_in(" ".join(str(c) for c in rows[header_idx])) if rows else []
    if not years:  # sometimes the year sits one row lower, or in cell text
        for r in rows[:header_idx + 3]:
            years = _years_in(" ".join(str(c) for c in r if c is not None))
            if years:
                break
    if not years and doc_years:                  # fall back to the file's fiscal year
        years = list(doc_years)
    elif not years and page_years:               # then to the page's own dates
        years = list(page_years)
    # detected only the comparative year (prose bleed ate the current one) --
    # promote to the file's own fiscal pair
    if doc_years and len(years) == 1 and years[0] == doc_years[1]:
        years = list(doc_years)
    years = sorted(set(years), reverse=True)[:4]

    # value columns = columns that hold real figures on most data rows
    ncols = max((len(r) for r in rows), default=1)
    data = rows[data_start:]
    valcols, notecol = [], None
    for c in range(1, ncols):
        vals = [r[c] for r in data if c < len(r) and isinstance(r[c], (int, float))]
        if len(vals) < max(3, 0.3 * len(data)):
            continue
        small = [v for v in vals if abs(v) < 1000]
        if len(small) == len(vals) and len(vals) < 0.7 * len(data):
            notecol = notecol if notecol is not None else c   # a note-ref column
        else:
            valcols.append(c)
    label_col = 0

    # year-column sanity: the header's year tokens, read left-to-right, should
    # line up with `years` (most-recent first).  If the header years are in the
    # opposite order, every figure is being attributed to the wrong year.
    if len(years) >= 2 and len(valcols) >= 2:
        hdr_years = []
        for r in rows[:max(1, data_start)]:
            for c in valcols:
                if c < len(r):
                    yy = _years_in(str(r[c]))
                    if yy:
                        hdr_years.append(yy[0])
        if len(hdr_years) >= 2 and hdr_years == sorted(hdr_years):
            # ascending in the header, but we mapped descending
            _add_note(t, "yearsReversed",
                "year columns may be reversed — the page header lists years "
                f"oldest-first ({hdr_years[0]}…{hdr_years[-1]}); figures could be "
                "attributed to the wrong year. Use 'Swap year columns' if so.",
                first=hdr_years[0], last=hdr_years[-1])

    # segmental / multi-entity: value-column headers carry entity names, not
    # years -- the columns are not comparable as a time series
    if not years and len(valcols) >= 2:
        htext = " ".join(
            str(rows[i][c]) for i in range(max(1, data_start))
            for c in valcols if c < len(rows[i]) and rows[i][c] is not None)
        wordy = len(re.findall(r"[A-Za-z]{4,}", htext))
        if wordy >= 2 and not re.search(r"AED|USD|['’]000|note", htext, re.I):
            _add_note(t, "segmentalColumns",
                "columns look like segments / entities, not reporting years — "
                "the figures across columns may not be a comparable time series.")

    # total / subtotal rows
    totals = []
    for i, r in enumerate(data):
        lbl = _row_label(r)
        if lbl and _TOTAL_RE.match(lbl):
            totals.append(data_start + i)

    # does it foot?  (only where a check is well-defined)
    foots = None
    try:
        if kind == "income statement" and valcols:
            # cut at 'profit/net profit for the year' so EPS / 'attributable
            # to' lines below it don't become the reconcile target
            cut = len(data)
            # 1. an explicit "profit/loss for the year" row -- possibly with a
            #    section header ("Other comprehensive ...", "attributable to:")
            #    welded onto it by the reconstruction
            for i, r in enumerate(data):
                if re.match(r"^\s*(net\s+)?(profit|loss) for the (year|period)\b"
                            r"(\s+(profit\s+)?attributable\s+to\s*:?"
                            r"|\s+other comprehensive.*"
                            r"|\s+from continuing operations)?\s*$",
                            _row_label(r), re.I):
                    cut = i + 1
            # 2. no such label (reconstruction lost it) -- stop just before the
            #    OCI / EPS tail instead
            if cut == len(data):
                for i, r in enumerate(data):
                    if re.search(r"other comprehensive (income|loss|\()"
                                 r"|total comprehensive income"
                                 r"|earnings per share|^\s*basic\b.*diluted",
                                 _row_label(r), re.I):
                        cut = i
                        break
            # 3. trim a trailing per-share / EPS tail (values < 1000 thousand)
            while cut >= 3:
                lastvals = [data[cut - 1][c] for c in valcols
                            if c < len(data[cut - 1])
                            and isinstance(data[cut - 1][c], (int, float))]
                if lastvals and all(abs(v) < 1000 for v in lastvals):
                    cut -= 1
                else:
                    break
            oks = []
            for c in valcols:
                seq = [(_row_label(r), r[c]) for r in data[:cut]
                       if c < len(r) and isinstance(r[c], (int, float))]
                oks.append(_reconciles(seq))
            foots = any(o is True for o in oks) if any(o is not None for o in oks) else None
        elif kind == "statement of financial position" and valcols:
            ok_any = None
            for c in valcols:
                ta = _find_val(data, r"total assets\b", c)
                te_ = _find_val(data, r"(shareholders['’]?|owners['’]?)\s+equity\b(?!\s+and)"
                                      r"|\btotal equity\b(?!\s+and)"
                                      r"|shareholders['’]?\s+funds\b", c)
                tl = _find_val(data, r"total liabilities\b(?!\s+and)", c)
                teq = _find_val(data, r"total equity and liabilities\b", c)
                na = _find_val(data, r"net assets\b", c)
                tnca = _find_val(data, r"total non[- ]?current assets\b", c)
                tca = _find_val(data, r"total current assets\b", c)
                _btol = CONFIG["tol_bs"]
                checks = []
                # assets side is internally consistent (works even when the
                # equity/liabilities half was truncated by the reconstruction)
                if ta is not None and tnca is not None and tca is not None:
                    checks.append(abs(ta - (tnca + tca)) <= _btol)
                if ta is not None and teq is not None:
                    checks.append(abs(ta - teq) <= _btol)
                if ta is not None and te_ is not None and tl is not None:
                    checks.append(abs(ta - (te_ + tl)) <= _btol)
                if na is not None and te_ is not None:
                    checks.append(abs(na - te_) <= _btol)      # du's "net assets = total equity"
                if ta is None and tnca is not None and tca is not None and te_ is not None:
                    # assets side only reaches sub-totals; equity/liab net to equity
                    tncl = _find_val(data, r"total non[- ]?current liabilities\b", c) or 0
                    tcl = _find_val(data, r"total current liabilities\b", c) or 0
                    checks.append(abs((tnca + tca) - (tncl + tcl) - te_) <= _btol)
                if checks:
                    ok_any = (ok_any or False) or any(checks)
            foots = ok_any
            # BS that footed on the assets side only, with no liabilities total
            # in view -- the equity/liabilities half didn't make it into the table
            _all = " || ".join(_row_label(r).lower() for r in data)
            if (foots and "total equity" in _all
                    and "total liabilities" not in _all
                    and "total equity and liabilities" not in _all
                    and "current liabilities" not in _all):
                _add_note(t, "assetsOnlyIncomplete",
                    "equity / liabilities side incomplete — only the assets side "
                    "was captured (check the following PDF page for the rest)")
        elif kind == "note" and valcols:
            oks = []
            for c in valcols:
                seq = [(_row_label(r), r[c]) for r in data
                       if c < len(r) and isinstance(r[c], (int, float))]
                oks.append(_reconciles(seq))
            foots = any(o is True for o in oks) if any(o is not None for o in oks) else None
        elif kind == "statement of cash flows" and valcols:
            oks = [_cashflow_foots(data, c) for c in valcols]
            # only claim NO FOOT if EVERY checkable column fails
            if any(o is True for o in oks):
                foots = True
            elif all(o is False for o in oks) and oks:
                foots = False
            else:
                foots = None
        elif kind == "statement of changes in equity" and valcols:
            # check the widest (usually rightmost "Total") column: closing
            # balance == opening balance + sum of the movement rows
            oks = [_equity_foots(data, c) for c in valcols]
            if any(o is True for o in oks):
                foots = True
            elif all(o is False for o in oks) and oks:
                foots = False
            else:
                foots = None
    except Exception:
        foots = None

    # record the worked arithmetic behind the verdict -- for EVERY value column,
    # so a broken prior-year column is not hidden by a good current-year one
    foot_detail, foot_by_col = "", []
    try:
        yrs = years or []
        for k, c in enumerate(valcols):
            yr = yrs[k] if k < len(yrs) else f"col {k + 1}"
            line = ""
            if kind in ("income statement", "note"):
                seq = [(_row_label(r), r[c])
                       for r in (data[:cut] if kind == "income statement" else data)
                       if c < len(r) and isinstance(r[c], (int, float))]
                ex = reconcile_explain(seq)
                line = ex.get("worked", "")
                foot_by_col.append({"year": yr, "ok": ex.get("ok"), "worked": line})
            elif kind == "statement of financial position":
                ta = _find_val(data, r"total assets\b", c)
                te2 = _find_val(data, r"(shareholders['’]?|owners['’]?)\s+equity\b(?!\s+and)"
                                      r"|\btotal equity\b(?!\s+and)", c)
                tl2 = _find_val(data, r"total liabilities\b(?!\s+and)", c)
                tnca = _find_val(data, r"total non[- ]?current assets\b", c)
                tca = _find_val(data, r"total current assets\b", c)
                if ta is not None and te2 is not None and tl2 is not None:
                    diff = ta - te2 - tl2
                    line = (f"total assets {ta:,.0f} vs equity {te2:,.0f} + "
                            f"liabilities {tl2:,.0f} = {te2 + tl2:,.0f} (diff {diff:,.0f})")
                    foot_by_col.append({"year": yr, "ok": abs(diff) <= CONFIG["tol_bs"], "worked": line})
                elif ta is not None and tnca is not None and tca is not None:
                    diff = tnca + tca - ta
                    line = (f"non-current {tnca:,.0f} + current {tca:,.0f} = "
                            f"{tnca + tca:,.0f} vs total assets {ta:,.0f} (diff {diff:,.0f})")
                    foot_by_col.append({"year": yr, "ok": abs(diff) <= CONFIG["tol_bs"], "worked": line})
            elif kind == "statement of cash flows":
                op = _find_val(data, r"net cash.{0,4}(generated|used|from|flows? from)\s+operating"
                                     r"|generated from operating activities", c)
                net = _find_val(data, r"(net\s+)?\(?(increase|decrease|decline)\)?[\s/()a-z]{0,20}"
                                      r"in cash", c)
                clo = _find_val(data, r"(cash and (cash equivalents|bank balances)|equivalents) "
                                      r"at (the )?(end|31 dec|close)", c)
                if op is not None and net is not None:
                    line = f"net cash from operating {op:,.0f}; net change in cash {net:,.0f}"
                    if clo is not None:
                        line += f"; closing cash {clo:,.0f}"
                    foot_by_col.append({"year": yr, "ok": _cashflow_foots(data, c), "worked": line})
            elif kind == "statement of changes in equity":
                ok = _equity_foots(data, c)
                if ok is not None:
                    line = "closing balance = opening + Σ movements"
                    # equity columns are Share capital / Retained earnings /
                    # Total etc, not years -- use the real header text once
                    # there's no year left to label a column with
                    if k >= len(yrs):
                        yr = _col_header_label(rows, data_start, c) or yr
                    foot_by_col.append({"year": yr, "ok": ok, "worked": line})
            if line:
                foot_detail += (("\n" if foot_detail else "") + f"{yr}: {line}")
    except Exception:
        foot_detail = ""

    # a "statement" candidate that doesn't look like one -- downgrade it.
    if kind in ("income statement", "statement of financial position",
                "statement of cash flows", "statement of changes in equity"):
        # structural giveaways (transition bridge / MD&A highlights / segmental
        # grid) demote the table even if some column happens to reconcile
        if _is_structural_non_statement(kind, data, valcols):
            kind = "table"
        elif foots is not True and _looks_like_not_a_statement(
                kind, data, valcols, years, totals):
            kind = "table"
    elif kind != "table":
        if not valcols and not years and t.get("shape") in ("mostly-text", "small/other"):
            kind = "table"
    if kind == "table":
        foots = None       # arithmetic verdict only meaningful for a real statement/note
        foot_detail, foot_by_col = "", []

    t.update(kind=kind, years=years, value_cols=valcols, note_col=notecol,
             label_col=label_col, header_idx=header_idx, data_start=data_start,
             total_rows=totals,
             foots=foots, foot_detail=foot_detail, foot_by_col=foot_by_col,
             notes=t.get("notes", []))
    return t


_DOWNGRADE_ANCHORS = (
    "revenue", "turnover", "total income", "operating income",
    "cost of sales", "operating expenses", "operating profit",
    "profit before", "profit for the year", "net profit",
    "total assets", "total equity", "total liabilities", "net assets",
    "non-current assets", "current assets", "share capital",
    "cash flows from operating", "net cash", "cash and cash equivalents",
    "financing activities", "investing activities",
    "balance at", "as at 1 january", "total comprehensive income")
_HIGHLIGHTS_RE = re.compile(
    r"\bebitda\b|\bhighlights?\b|profit and loss summary"
    r"|balance sheet summary|cash flow summary", re.I)
# a "... margin" METRIC row (short, ends in margin) -- not "Margin on guarantees"
_MARGIN_ROW_RE = re.compile(r"^[\w /()-]{0,24}\bmargin\b\s*%?$", re.I)


def _is_structural_non_statement(kind, data, valcols):
    """Shape-level giveaways that this is NOT a face statement, regardless of
    whether a column reconciles:
      · an MD&A highlights block (EBITDA / margin / "... Summary" rows)
      · a segmental / multi-entity income statement (>= 4 figure columns)
      · a "previously reported + adjustment + restated" transition grid
        (3 columns, col_a + col_b == col_c on most rows)
    """
    # a "previously reported + adjustment + restated" / transaction-impact grid:
    # 3 figure columns where col_a + col_b == col_c on almost every row.  A real
    # 3-year statement (2024/2023/2022) essentially never satisfies this.
    if len(valcols) == 3:
        a, b, c = valcols
        rowvals = [[r[i] if i < len(r) else None for i in (a, b, c)] for r in data]
        full = [v for v in rowvals if all(isinstance(x, (int, float)) for x in v)]
        hits = sum(1 for v in full
                   if abs(v[0] + v[1] - v[2]) <= max(2.0, abs(v[2]) * 1e-4))
        if len(full) >= 5 and hits >= 0.85 * len(full):
            return True
    return False


def _looks_like_not_a_statement(kind, data, valcols, years, totals):
    """A candidate typed as a face statement but that reads like a note / a
    policy page.  Applied only when the figures do NOT already reconcile."""
    labs = " || ".join(_row_label(r).lower() for r in data)

    n_anchor = sum(1 for a in _DOWNGRADE_ANCHORS if a in labs)

    def _is_prose(r):
        if [v for v in _row_nums(r) if abs(v) >= 100]:
            return False
        lb = _row_label(r)
        return bool(_LEAK_RE.search(lb)) or len(lb.split()) >= CONFIG["prose_min_words"]
    prose = sum(1 for r in data if _is_prose(r))

    has_total = bool(totals) or any(
        re.search(r"total|net (assets|cash|profit)|profit for the year", _row_label(r), re.I)
        for r in data)

    yr_ok = len(years) < 2 or abs(years[0] - years[1]) <= CONFIG["max_year_gap"]

    # EBITDA / "... margin" rows point at an MD&A highlights block -- but du
    # genuinely prints EBITDA on the face of its P&L, so this only counts when
    # the figures also fail to reconcile (which is why it lives here, gated)
    highlights = bool(_HIGHLIGHTS_RE.search(labs)) or \
        any(_MARGIN_ROW_RE.match(_row_label(r)) for r in data)

    # >= 4 figure columns on an "income statement" -> segmental / multi-entity
    many_cols = kind == "income statement" and len(valcols) >= 4

    return (n_anchor < 3
            or prose > len(data) * CONFIG["prose_row_ratio"]
            or not has_total or not yr_ok
            or highlights or many_cols)


def _find_val(data, label_re, col):
    """value in `col` on the LAST row whose label contains `label_re`."""
    rx = re.compile(label_re, re.I)
    hit = None
    for r in data:
        if rx.search(_row_label(r)) and col < len(r) and isinstance(r[col], (int, float)):
            hit = r[col]
    return hit


def _cashflow_foots(data, col, tol=None):
    """net change in cash == operating + investing + financing subtotals, OR
    opening + net change (+ FX) == closing, OR -- for statements that print no
    investing sub-total (du's format) -- operating sub-total + (sum of the
    investing lines) + financing sub-total == net change.  Cash-flow rounding
    drift is wider, so the tolerance is looser."""
    if tol is None:
        tol = CONFIG["tol_cashflow"]
    def _v(r):
        return r[col] if col < len(r) and isinstance(r[col], (int, float)) else None
    labs = [_row_label(r) for r in data]

    def _idx(rx):
        rx = re.compile(rx, re.I)
        hit = None
        for i, l in enumerate(labs):
            if rx.search(l):
                hit = i
        return hit

    # leading words ("Net cash", "Cash") are sometimes clipped by the label
    # reconstruction, so anchor on the distinctive tail of each line
    op_i  = _idx(r"(net cash|cash).{0,4}(generated (from|by)|used (in|for)|(in|from|flows? from))\s+operating"
                 r"|generated from operating activities")
    inv_i = _idx(r"(net )?cash (used (in|for)|generated (from|by)|(in|from))\s+investing"
                 r"|used in investing activities")
    fin_i = _idx(r"(net\s+)?cash\s+.{0,30}financing activities\s*$")
    net_i = _idx(r"(net\s+)?\(?(increase|decrease|decline)\)?[\s/()a-z]{0,20}in cash( and cash equivalents)?")
    # bare section header that starts the financing block (value usually blank)
    fs_i  = _idx(r"^\s*(cash flows? (from|used (in|for))\s+)?financing activities\s*$")

    op  = _v(data[op_i])  if op_i  is not None else None
    inv = _v(data[inv_i]) if inv_i is not None else None
    fin = _v(data[fin_i]) if fin_i is not None else None
    net = _v(data[net_i]) if net_i is not None else None
    opening = _find_val(data, r"(cash and (cash equivalents|bank balances)|equivalents) at (the )?(beginning|start|1 jan)", col)
    closing = _find_val(data, r"(cash and (cash equivalents|bank balances)|equivalents) at (the )?(end|31 dec|close)", col)
    # optional FX-retranslation line that sits between net change and closing
    fx = _find_val(data, r"effect of (foreign )?(exchange|currency)|exchange rate changes"
                         r"|currency (translation|retranslation)", col) or 0.0

    # scale the tolerance to the size of the flows -- a fixed 200 is too tight
    # for a AED-billions statement and too loose for a small one
    _scale = max(abs(v) for v in (op, fin, net, closing) if v) if any(
        v for v in (op, fin, net, closing)) else 0
    tol = max(tol, _scale * 5e-4)

    if op is not None and inv is not None and fin is not None and net is not None:
        return abs((op + inv + fin) - net) <= tol
    if opening is not None and closing is not None and net is not None:
        return abs((opening + net + fx) - closing) <= tol
    if op is not None and inv is not None and fin is not None \
            and opening is not None and closing is not None:
        return abs((opening + op + inv + fin + fx) - closing) <= tol

    # du-style: no investing sub-total, but a "financing activities" header
    # delimits the block -- sum the investing lines directly
    if (op_i is not None and net_i is not None and fs_i is not None
            and op is not None and net is not None
            and op_i < fs_i < net_i):
        inv_sum = sum(v for r in data[op_i + 1:fs_i]
                      if (v := _v(r)) is not None)
        if fin is not None and fin_i is not None and fin_i > fs_i:
            fin_val = fin
        else:
            fin_val = sum(v for r in data[fs_i + 1:net_i]
                          if (v := _v(r)) is not None)
        return abs((op + inv_sum + fin_val) - net) <= tol

    return None          # not enough anchor rows to judge -> leave it blank


def _col_header_label(rows, data_start, col):
    """Join whatever header-row text sits in this column (e.g. 'Share' /
    'capital AED 000' stacked across two header lines) into one short label.
    Used where a value column doesn't correspond to a single "year" the way
    most statements' columns do -- a changes-in-equity statement's value
    columns are Share capital / Share premium / ... / Total, not years, so
    the generic "col N" placeholder used elsewhere is meaningless there."""
    parts = []
    for r in rows[:data_start]:
        if col < len(r) and isinstance(r[col], str) and r[col].strip():
            parts.append(r[col].strip())
    label = re.sub(r"\s+", " ", " ".join(parts)).strip()
    if len(label) > 28:
        label = label[:28].rsplit(" ", 1)[0]
    return label or None


def _equity_foots(data, col, tol=None):
    """for the TOTAL column of a changes-in-equity statement: last balance ==
    first balance + sum of the movement rows."""
    if tol is None:
        tol = CONFIG["tol_equity"]
    bal_rows = [i for i, r in enumerate(data)
                if re.search(r"(as )?(at|balance).{0,4}(1 january|31 december|beginning|end)",
                             _row_label(r), re.I)
                and col < len(r) and isinstance(r[col], (int, float))]
    if len(bal_rows) < 2:
        return None
    # a two-year SOCE stacks two opening->closing blocks; check the LAST one
    a, b = bal_rows[-2], bal_rows[-1]
    opening = data[a][col]
    closing = data[b][col]
    moves = [data[i][col] for i in range(a + 1, b)
             if col < len(data[i]) and isinstance(data[i][col], (int, float))
             and not re.search(r"(as )?(at|balance)", _row_label(data[i]), re.I)]
    if len(moves) < 2:            # too thin to be a real check -> don't claim
        return None
    return abs(opening + sum(moves) - closing) <= tol


def _reconciles(seq, tol=None):
    """seq: [(label, value), ...] for ONE value column of a statement / note.
    Walk it keeping a running total of the LEAF lines; a line is not a leaf if
    its label is a total/subtotal phrase, or its value equals the running
    total (a rolling subtotal), or its value equals the sum of a contiguous
    run of the immediately-preceding leaves (a section subtotal like DFM's
    'Operating expenses' = G&A + Amortisation + Interest).  The final value
    must equal the running leaf total."""
    if tol is None:
        tol = CONFIG["tol_pl"]
    seq = [(l, float(v)) for l, v in seq if isinstance(v, (int, float))]
    if len(seq) < 3:
        return None
    target = seq[-1][1]
    running = 0.0
    leaves = []           # leaf values, in order
    for lbl, v in seq[:-1]:
        if _HARD_TOTAL_RE.match(lbl or ""):
            continue
        if leaves and abs(v - running) <= tol:
            continue
        # section subtotal: v == sum of the last k leaves (k >= 2)
        acc = 0.0
        is_sub = False
        for k in range(1, min(len(leaves), 12) + 1):
            acc += leaves[-k]
            if k >= 2 and abs(v - acc) <= tol:
                is_sub = True
                break
        if is_sub:
            continue
        running += v
        leaves.append(v)
    return len(leaves) >= 2 and abs(running - target) <= tol


def reconcile_explain(seq, tol=None):
    """Same walk as `_reconciles`, but return a dict that SHOWS the work:
        {ok, target, running, gap, break_label, worked}
    `worked` is a human string; `break_label` is the first line after which the
    running total diverges from where it should be (best-effort localisation)."""
    if tol is None:
        tol = CONFIG["tol_pl"]
    s = [(l, float(v)) for l, v in seq if isinstance(v, (int, float))]
    if len(s) < 3:
        return {"ok": None, "worked": "too few numeric rows to check"}
    target = s[-1][1]
    running, leaves = 0.0, []          # leaves: list of (label, value)
    for lbl, v in s[:-1]:
        if _HARD_TOTAL_RE.match(lbl or ""):
            continue
        if leaves and abs(v - running) <= tol:
            continue
        acc, is_sub = 0.0, False
        for k in range(1, min(len(leaves), 12) + 1):
            acc += leaves[-k][1]
            if k >= 2 and abs(v - acc) <= tol:
                is_sub = True
                break
        if is_sub:
            continue
        running += v
        leaves.append((lbl, v))
    gap = running - target
    ok = len(leaves) >= 2 and abs(gap) <= tol
    # localisation -- only where it can be done honestly:
    #  A. a leaf that shouldn't be counted (doubled row / stray subtotal):
    #     running - leaf == target
    #  B. gap is small vs the total and one leaf's magnitude ~= the gap:
    #     a row of about that size is missing, or that leaf is mis-read
    break_label, break_note = None, None
    if not ok:
        for lbl, v in leaves:
            if abs((running - v) - target) <= tol:
                break_label = lbl
                break_note = "counted twice / is a stray subtotal"
                break
        if break_label is None and abs(gap) <= abs(target or 1) * 0.15:
            cand = min(((abs(abs(v) - abs(gap)), lbl) for lbl, v in leaves),
                       default=(None, None))
            if cand[0] is not None and cand[0] <= abs(gap) * 0.1:
                break_label = cand[1]
                break_note = "a row about this size is missing, or this one is mis-read"
    tail = ""
    if break_label:
        tail = f"  (check row: “{break_label}” — {break_note})"
    return {
        "ok": ok, "target": target, "running": running, "gap": gap,
        "break_label": break_label, "break_note": break_note,
        "worked": (f"sum of {len(leaves)} line items = {running:,.0f}; "
                   f"printed total = {target:,.0f}; difference = {gap:,.0f}{tail}"),
    }


_KNOWN_LINE_RE = re.compile(
    r"^(revenue|turnover|cost of sales|gross profit|operating (profit|expenses|income)"
    r"|finance (income|costs?|expense)|profit (before|for the|attributable)"
    r"|(loss|profit)/?\(?(loss|profit)?\)?|income tax|taxation|depreciation|amortisation"
    r"|impairment|share of (results|profit|loss)|dividends?|federal royalty|royalty"
    r"|total\b|net\b|non-current|current|inventories|trade (and other )?(receivables|payables)"
    r"|cash (and )?(cash equivalents|bank balances|generated|used)|borrowings|lease"
    r"|property, plant|intangible|goodwill|provisions?|retained earnings|share (capital|premium)"
    r"|other (comprehensive|reserves|income|assets|liabilities)|changes in|adjustments for"
    r"|purchase of|proceeds from|acquisition of|repayment of|payment of|interest (paid|received)"
    r"|(basic|diluted).*(earnings|per share)|earnings per share|cash flows? (from|used)"
    r"|equity( and liabilities)?|liabilities|assets|represented by|attributable to)",
    re.I)


def label_health(t):
    """How trustworthy are this table's row LABELS (not its figures)?
    Returns {score 0-1, suspect: [{row, label, why}]}.  Catches the failure
    modes the reconstruction has: prose bled in from an adjacent column, a
    heading welded onto a line, a first word clipped off, an empty label."""
    rows = t["rows"]
    hi = t.get("data_start", t.get("header_idx", 0) + 1)
    data = rows[hi:]
    suspect, scored = [], 0
    for i, r in enumerate(data):
        lb = _row_label(r)
        has_fig = bool([v for v in _row_nums(r) if abs(v) >= 100])
        if not has_fig and not lb:
            continue                       # blank spacer row -- not counted
        scored += 1
        known = bool(_KNOWN_LINE_RE.match(lb))
        why = None
        if has_fig and not lb:
            why = "figures with no label"
        elif not has_fig and _LEAK_RE.search(lb) and not known and len(lb.split()) >= 8:
            why = "reads as a prose sentence"
        elif len(lb.split()) >= 14:
            why = "very long label (probably merged text)"
        elif re.match(r"^[a-z]", lb) and not re.match(r"^[a-z]+\)", lb) and not known:
            why = "starts lower-case (leading word clipped?)"
        elif re.search(r"[A-Z]{5,}", lb) and re.search(r"[a-z]", lb):
            why = "embedded ALL-CAPS run (section header welded in)"
        if why:
            suspect.append({"row": i, "label": lb, "why": why})
    score = 1.0 if scored == 0 else max(0.0, 1.0 - len(suspect) / scored)
    return {"score": round(score, 3), "suspect": suspect, "scored": scored}


def _median(xs):
    xs = sorted(xs)
    n = len(xs)
    return 0 if not n else (xs[n // 2] if n % 2 else (xs[n // 2 - 1] + xs[n // 2]) / 2)


def _mean_sd(xs):
    xs = list(xs)
    if len(xs) < 2:
        return (xs[0] if xs else 0.0), 0.0
    m = sum(xs) / len(xs)
    var = sum((x - m) ** 2 for x in xs) / (len(xs) - 1)
    return m, var ** 0.5


def figure_health(t):
    """Sanity-check the FIGURES (label_health covers the labels).

    Thresholds are DERIVED from each column's own distribution rather than
    magic constants: a value is an outlier if its log10-magnitude sits more
    than `figure_outlier_sd` (default 4) standard deviations above the column
    mean AND is at least 8x the column's second-largest value.  Also flags a
    row whose two year figures differ by >= 100x (one probably mis-read).
    """
    import math
    hi = t.get("data_start", t.get("header_idx", 0) + 1)
    data = t["rows"][hi:]
    vcols = t.get("value_cols") or []
    totals = set(i - hi for i in t.get("total_rows", []))
    if not vcols:
        return {"score": 1.0, "suspect": [], "scored": 0}
    k = CONFIG.get("figure_outlier_sd", 4.0)
    logmean, logsd, second = {}, {}, {}
    for c in vcols:
        mags = sorted((abs(r[c]) for r in data
                       if c < len(r) and isinstance(r[c], (int, float)) and r[c]),
                      reverse=True)
        logs = [math.log10(m) for m in mags if m > 0]
        logmean[c], logsd[c] = _mean_sd(logs)
        second[c] = mags[1] if len(mags) > 1 else (mags[0] if mags else 0)
    suspect, scored = [], 0
    for i, r in enumerate(data):
        vals = [(c, r[c]) for c in vcols
                if c < len(r) and isinstance(r[c], (int, float))]
        if not vals:
            continue
        scored += 1
        lb = _row_label(r)
        why = None
        for c, v in vals:
            if (v and logsd[c] > 0 and i not in totals
                    and not _HARD_TOTAL_RE.match(lb or "")
                    and math.log10(abs(v)) > logmean[c] + k * logsd[c]
                    and abs(v) > 8 * second[c]):
                why = f"{v:,.0f} is a magnitude outlier for its column (extra digit?)"
        if why is None and len(vals) >= 2:
            nz = [abs(v) for _, v in vals if v]
            if len(nz) >= 2 and min(nz) and max(nz) / min(nz) >= 100:
                why = ("one year is ~100x the other ("
                       + " vs ".join(f"{v:,.0f}" for _, v in vals) + ") — mis-read?")
        if why:
            suspect.append({"row": i, "label": lb, "why": why,
                            "values": [v for _, v in vals]})
    score = 1.0 if scored == 0 else max(0.0, 1.0 - len(suspect) / scored)
    return {"score": round(score, 3), "suspect": suspect, "scored": scored}


def cross_year_check(tables):
    """When several reports are scanned together, verify that report A's
    prior-year column equals report B's current-year column for the same
    statement and the same year.  Attaches t['consistency'] to each table it
    could check.  This is the strongest signal available -- previously only
    `--compare` used it."""
    STMT_ = ("income statement", "statement of financial position",
             "statement of cash flows", "statement of changes in equity")
    by_kind = {}
    for t in tables:
        if t["kind"] in STMT_ and (t.get("years") or []):
            by_kind.setdefault(t["kind"], []).append(t)
    for kind, group in by_kind.items():
        for i, ta in enumerate(group):
            best = None
            for tb in group[i + 1:]:
                if ta["file"] == tb["file"]:
                    continue
                shared = sorted(set(ta["years"]) & set(tb["years"]))
                if not shared:
                    continue
                y = shared[-1]
                sa, sb = _series(ta), _series(tb)

                def at(series, yr):
                    return {k: v.get(yr) for k, v in series.items()
                            if isinstance(v.get(yr), (int, float))}
                da, db = at(sa, y), at(sb, y)
                common = set(da) & set(db)
                if len(common) < 4:
                    continue
                mism = [(k, da[k], db[k]) for k in common
                        if abs(da[k] - db[k]) > max(2.0, abs(da[k]) * 1e-4)]
                # a handful of differing lines reads as a genuine restatement of
                # specific items; a large fraction reads as a column being
                # mis-aligned / mis-scaled in one of the two extractions
                frac = len(mism) / max(1, len(common))
                verdict = ("agree" if not mism
                           else "restated" if (len(mism) <= 3 or frac <= 0.25)
                           else "columns likely misaligned")
                rec = {"vs": tb["file"], "year": y, "checked": len(common),
                       "mismatch": len(mism), "verdict": verdict,
                       "worst": sorted(mism, key=lambda m: -abs(m[1] - m[2]))[:3]}
                if best is None or rec["checked"] > best[0]["checked"]:
                    best = (rec, tb)
            if best:
                rec, tb = best
                ta["consistency"] = rec
                mirror = dict(rec); mirror["vs"] = ta["file"]
                tb.setdefault("consistency", mirror)


def _looks_scanned(pdf):
    """Digital-text check: near-empty text layer across the sampled pages."""
    try:
        pages = pdf.pages
        sample = pages[:12] if len(pages) > 12 else pages
        chars = sum(len(p.extract_text() or "") for p in sample) / max(1, len(sample))
        return chars < CONFIG["scanned_chars_per_page"]
    except Exception:
        return False


# ------------------------------------------------------------------ compare ---
# Line items that different reports (or the same issuer year to year) word
# differently but mean the same thing.  Maps a normalised label -> canonical.
_LABEL_ALIASES = {
    "turnover": "revenue",
    "total revenue": "revenue",
    "revenues": "revenue",
    "cost of revenue": "cost of sales",
    "profit before taxation": "profit before tax",
    "profit before income tax": "profit before tax",
    "income tax expense": "income tax",
    "income tax expenses": "income tax",
    "taxation": "income tax",
    "net profit for the year": "profit for the year",
    "profit for the period": "profit for the year",
    "depreciation and amortisation": "depreciation and amortization",
    "property plant and equipment": "property plant equipment",
    "trade and other receivables": "trade receivables",
    "trade and other payables": "trade payables",
    "cash and bank balances": "cash and cash equivalents",
    "cash and short term deposits": "cash and cash equivalents",
    "total shareholders equity": "total equity",
    "shareholders equity": "total equity",
    "shareholders funds": "total equity",
    "net cash from operating activities": "net cash generated from operating activities",
    "net cash used in operating activities": "net cash generated from operating activities",
    "net cash flows from operating activities": "net cash generated from operating activities",
}


def _norm_label(s):
    s = re.sub(r"\(note[s]?\s*[\d.,\s]+\)", "", s, flags=re.I)
    s = re.sub(r"[^a-z0-9 ]", " ", s.lower())
    s = re.sub(r"\s+", " ", s).strip()
    return _LABEL_ALIASES.get(s, s)


def _match_tables(a_tables, b_tables):
    """Pair up the same logical table across two scans (by kind, then title)."""
    from difflib import SequenceMatcher
    pairs, used_b = [], set()
    for ta in a_tables:
        best, best_s = None, 0.0
        for j, tb in enumerate(b_tables):
            if j in used_b or ta["kind"] != tb["kind"]:
                continue
            s = SequenceMatcher(None, _norm_label(ta.get("title", "")),
                                _norm_label(tb.get("title", ""))).ratio()
            if ta["kind"] in ("income statement", "statement of financial position",
                              "statement of cash flows", "statement of changes in equity"):
                s += 0.5  # the primary statements: kind alone is a strong signal
            if s > best_s:
                best, best_s = j, s
        if best is not None and best_s >= 0.45:
            used_b.add(best)
            pairs.append((ta, b_tables[best]))
    return pairs


def _series(t):
    """{normalised label: {year: value}} for a table, using its value columns."""
    out = {}
    data = t["rows"][t.get("data_start", t["header_idx"] + 1):]
    if t["kind"] == "income statement":
        # stop at 'profit for the year' so the OCI tail (often fragmented in a
        # reconstruction) doesn't pollute the diff
        for i, r in enumerate(data):
            if re.match(r"^\s*(net\s+)?(profit|loss) for the (year|period)\b"
                        r"(\s+(profit\s+)?attributable\s+to\s*:?)?\s*$",
                        _row_label(r), re.I):
                data = data[:i + 1]
                break
    yrs = t["years"] or list(range(9000, 9000 + len(t["value_cols"])))
    for r in data:
        lbl = _row_label(r)
        if not lbl or _LEAK_RE.search(lbl):
            continue
        key = _norm_label(lbl)
        if not key:
            continue
        rec = out.setdefault(key, {"label": lbl})
        for k, c in enumerate(t["value_cols"]):
            if c < len(r) and isinstance(r[c], (int, float)):
                rec[yrs[k] if k < len(yrs) else f"c{k}"] = r[c]
    return out


def diff_tables(ta, tb):
    """Row-level diff of the same table in two reports.  Returns rows for a
    worksheet plus a short verdict."""
    sa, sb = _series(ta), _series(tb)
    # fuzzy reconciliation: a line item whose wording drifted between the two
    # reports ("Federal royalty" vs "Federal royalty on regulated profit")
    # should not show up as one NEW + one removed
    from difflib import SequenceMatcher
    only_a = [k for k in sa if k not in sb]
    only_b = [k for k in sb if k not in sa]
    for ka in list(only_a):
        best, best_s = None, 0.0
        for kb in only_b:
            s = SequenceMatcher(None, ka, kb).ratio()
            # a strong containment (one label is the other plus a qualifier)
            if ka in kb or kb in ka:
                s = max(s, 0.9)
            if s > best_s:
                best, best_s = kb, s
        if best is not None and best_s >= 0.82:
            sb[ka] = sb.pop(best)
            only_b.remove(best)

    ya = sorted(y for y in (ta["years"] or []) if isinstance(y, int))
    yb = sorted(y for y in (tb["years"] or []) if isinstance(y, int))
    a_cur, a_prev = (ya[-1] if ya else None), (ya[0] if len(ya) > 1 else None)
    b_cur, b_prev = (yb[-1] if yb else None), (yb[0] if len(yb) > 1 else None)

    hdr = ["Line item",
           f"A {a_cur}" if a_cur else "A cur", f"A {a_prev}" if a_prev else "A prev",
           f"B {b_cur}" if b_cur else "B cur", f"B {b_prev}" if b_prev else "B prev",
           "Δ A(cur−prev)", "Δ% A", "restated?", "status"]
    out_rows = [hdr]
    keys = list(dict.fromkeys(list(sa) + list(sb)))
    n_new = n_gone = n_restated = n_changed = 0
    for k in keys:
        ra, rb = sa.get(k), sb.get(k)
        lbl = (ra or rb)["label"]
        av_c = ra.get(a_cur) if ra else None
        av_p = ra.get(a_prev) if ra else None
        bv_c = rb.get(b_cur) if rb else None
        bv_p = rb.get(b_prev) if rb else None
        d = (av_c - av_p) if isinstance(av_c, (int, float)) and isinstance(av_p, (int, float)) else None
        dp = (100.0 * d / abs(av_p)) if d is not None and av_p else None
        restated = ""
        # A's prior-year column should equal B's current-year column (same year)
        if a_prev and b_cur and a_prev == b_cur \
           and isinstance(av_p, (int, float)) and isinstance(bv_c, (int, float)):
            if abs(av_p - bv_c) > 2:
                restated = f"{av_p:,.0f} vs {bv_c:,.0f}"; n_restated += 1
        if ra and not rb:
            status = "NEW in A"; n_new += 1
        elif rb and not ra:
            status = "removed (only in B)"; n_gone += 1
        elif d not in (None, 0):
            status = "changed"; n_changed += 1
        else:
            status = ""
        out_rows.append([lbl, av_c, av_p, bv_c, bv_p, d,
                         round(dp, 1) if dp is not None else None, restated, status])
    verdict = (f"{n_changed} changed, {n_new} new, {n_gone} removed"
               + (f", {n_restated} RESTATED" if n_restated else ", prior-year columns agree"))
    # `verdict` above is the CLI/xlsx-facing English sentence; `counts` is the
    # same result as plain numbers so a caller that wants translated UI text
    # (see webui.html's renderComparePanel) can build its own sentence from
    # them instead of hard-coding English -- see the note above I18N in
    # webui.html for why the verdict PROSE itself isn't machine-translated.
    counts = {"changed": n_changed, "new": n_new, "removed": n_gone, "restated": n_restated}
    return out_rows, verdict, counts


# --------------------------------------------------------------------- write ---
TITLE_FILL = PatternFill("solid", fgColor="1F3864")
HEAD_FILL = PatternFill("solid", fgColor="D9E2F3")
IDX_FILL = PatternFill("solid", fgColor="1F3864")


def safe_sheet_name(name, used):
    name = re.sub(r"[\[\]:*?/\\]", "_", name).strip() or "table"
    name = name[:28]
    base, i = name, 2
    while name in used:
        name = f"{base[:25]}_{i}"
        i += 1
    used.add(name)
    return name


_KIND_SHORT = {
    "income statement": "P&L",
    "statement of financial position": "Balance sheet",
    "statement of cash flows": "Cash flow",
    "statement of changes in equity": "Equity",
    "note": "Note",
}


def short_sheet_name(t, n, used):
    """A readable tab name: an explicit user-set name wins, else 'P&L 2024' /
    'Balance sheet 2024', else the guessed title."""
    if t.get("_sheet_name"):
        base = str(t["_sheet_name"])[:28]
    else:
        base = _KIND_SHORT.get(t["kind"])
        yr = (t.get("years") or [None])[0]
        if base and yr:
            base = f"{base} {yr}"
        elif not base:
            base = (t.get("title") or f"Table p{t['page_label']}")[:24]
    return safe_sheet_name(f"{base}", used) if base not in used else \
        safe_sheet_name(f"{base} ({n})", used)


def autofit(ws, max_col, cap=60):
    widths = {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None or cell.column > max_col:
                continue
            widths[cell.column] = max(widths.get(cell.column, 8),
                                      min(len(str(cell.value)) + 2, cap))
    for c, w in widths.items():
        ws.column_dimensions[get_column_letter(c)].width = w


def build_workbook(all_tables):
    """One sheet per table (title, source, data, and a change column where two
    year columns are present) + a Contents index carrying the semantic reading."""
    wb = Workbook()
    idx = wb.active
    idx.title = "Contents"
    idx.append(["#", "File", "PDF page", "What", "Years", "Foots?", "Title",
                "Rows", "Cols", "Sheet"])
    for c in range(1, 11):
        cell = idx.cell(row=1, column=c)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = IDX_FILL
    idx.freeze_panes = "A2"

    used = set()
    for n, t in enumerate(all_tables, 1):
        rows = t["rows"]
        ncols = max((len(r) for r in rows), default=1)
        vcols = t.get("value_cols") or []
        add_delta = len(vcols) >= 2
        outw = ncols + (2 if add_delta else 0)

        sn = short_sheet_name(t, n, used)
        ws = wb.create_sheet(title=sn)

        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(outw, 4))
        c = ws.cell(row=1, column=1, value=t["title"] or "(untitled table)")
        c.font = Font(bold=True, color="FFFFFF", size=12)
        c.fill = TITLE_FILL
        c.alignment = Alignment(vertical="center", indent=1)
        ws.row_dimensions[1].height = 22

        yrs = "/".join(str(y) for y in (t.get("years") or [])) or "?"
        foot = {True: "figures foot ✓", False: "figures do NOT foot ✗", None: ""}[t.get("foots")]
        c = ws.cell(row=2, column=1, value=(
            f"{t['file']}  ·  PDF page {t['page_label']}  ·  {t['kind']}  ·  "
            f"years {yrs}  ·  {len(rows)}×{ncols}   {foot}"))
        c.font = Font(italic=True, color="595959", size=9)

        r0 = 4
        hi = t.get("header_idx", 0)
        for i, row in enumerate(rows):
            for cx in range(ncols):
                val = row[cx] if cx < len(row) else None
                # a cell that carried a currency symbol / '%' the value
                # itself can't keep (needed as a plain number for footing,
                # Δ, health) -- write what was actually printed instead of
                # the bare number, same principle as the browser's own
                # "fmt" side-channel (see serve.py's _fmt_row).
                cell_val = val.formatted() if isinstance(val, FormattedNumber) and (val.prefix or val.suffix) else val
                cell = ws.cell(row=r0 + i, column=cx + 1, value=cell_val)
                if i <= hi:
                    cell.font = Font(bold=True); cell.fill = HEAD_FILL
                elif (r0 + i - r0) and (hi + 1 + (i - hi - 1)) in ():
                    pass
                if isinstance(cell_val, (int, float)):
                    cell.number_format = "#,##0.00" if isinstance(cell_val, float) else "#,##0"
                    cell.alignment = Alignment(horizontal="right")
                elif isinstance(val, FormattedNumber):
                    cell.alignment = Alignment(horizontal="right")
            if add_delta and i > hi:
                a = row[vcols[0]] if vcols[0] < len(row) else None
                b = row[vcols[1]] if vcols[1] < len(row) else None
                if isinstance(a, (int, float)) and isinstance(b, (int, float)):
                    dcell = ws.cell(row=r0 + i, column=ncols + 1, value=a - b)
                    dcell.number_format = "#,##0;(#,##0)"
                    dcell.alignment = Alignment(horizontal="right")
                    if b:
                        pc = ws.cell(row=r0 + i, column=ncols + 2, value=round(100.0 * (a - b) / abs(b), 1))
                        pc.number_format = '0.0"%"'
                        pc.alignment = Alignment(horizontal="right")
        if add_delta:
            for off, lab in ((1, "Δ (change)"), (2, "Δ %")):
                hc = ws.cell(row=r0 + hi, column=ncols + off, value=lab)
                hc.font = Font(bold=True); hc.fill = HEAD_FILL
        # mark total / subtotal rows bold
        for tr in t.get("total_rows", []):
            for cx in range(outw):
                ws.cell(row=r0 + tr, column=cx + 1).font = Font(bold=True)

        ws.freeze_panes = ws.cell(row=r0 + hi + 1, column=2)
        autofit(ws, outw)

        idx.append([n, t["file"], t["page_label"], t["kind"], yrs,
                    _foot_mark(t), t["title"], len(rows), ncols, sn])
    autofit(idx, 10)
    return wb


def build_compare_workbook(pairs, meta):
    """pairs: list of (table_A, table_B); meta: (fileA, fileB).
    One sheet per matched table with a row-level diff, plus a summary."""
    wb = Workbook()
    summ = wb.active
    summ.title = "Summary"
    summ.append(["Comparing", meta[0], "against", meta[1]])
    summ.append([])
    summ.append(["#", "What", "Title (A)", "Rows compared", "Verdict", "Sheet"])
    for c in range(1, 7):
        summ.cell(row=3, column=c).font = Font(bold=True, color="FFFFFF")
        summ.cell(row=3, column=c).fill = IDX_FILL

    used = set()
    for n, (ta, tb) in enumerate(pairs, 1):
        drows, verdict, _counts = diff_tables(ta, tb)
        sn = safe_sheet_name((ta.get("title") or ta["kind"]) + f" diff ({n})", used)
        ws = wb.create_sheet(title=sn)
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=9)
        c = ws.cell(row=1, column=1,
                    value=f"{ta['kind']}  —  {ta.get('title','')}   [{verdict}]")
        c.font = Font(bold=True, color="FFFFFF", size=11); c.fill = TITLE_FILL
        c.alignment = Alignment(vertical="center", indent=1)
        r0 = 3
        for i, row in enumerate(drows):
            for cx, val in enumerate(row):
                cell = ws.cell(row=r0 + i, column=cx + 1, value=val)
                if i == 0:
                    cell.font = Font(bold=True); cell.fill = HEAD_FILL
                elif isinstance(val, (int, float)):
                    cell.number_format = "#,##0;(#,##0)"
                    cell.alignment = Alignment(horizontal="right")
                elif cx == 8 and val:
                    cell.font = Font(color="C00000" if "RESTAT" in verdict.upper()
                                     and val == "changed" else "1F3864")
                elif cx == 7 and val:  # restated? column
                    cell.font = Font(bold=True, color="C00000")
        ws.freeze_panes = ws.cell(row=r0 + 1, column=2)
        autofit(ws, 9)
        summ.append([n, ta["kind"], ta.get("title", ""), len(drows) - 1, verdict, sn])
    autofit(summ, 6)
    return wb


# ---------------------------------------------------------------------- main ---
def gather(inputs, recursive):
    out, seen = [], set()
    for a in inputs:
        p = Path(a)
        if p.is_dir():
            for f in sorted(p.glob("**/*.pdf" if recursive else "*.pdf")):
                if f.resolve() not in seen:
                    seen.add(f.resolve()); out.append(f)
        elif p.exists() and p.suffix.lower() == ".pdf":
            if p.resolve() not in seen:
                seen.add(p.resolve()); out.append(p)
        else:
            print(f"  SKIP: {p}")
    return out



def _parse_selection(spec, n):
    """'1,3,5-8'  ->  [1,3,5,6,7,8]  (1-based, clamped to 1..n)"""
    out = set()
    for part in re.split(r"[,\s]+", spec.strip()):
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            out.update(range(int(a), int(b) + 1))
        else:
            out.add(int(part))
    return sorted(x for x in out if 1 <= x <= n)


_BS_CLOSE_RE = re.compile(r"total equity and liabilities|total liabilities and equity", re.I)
_PL_CLOSE_RE = re.compile(r"(profit|loss) for the (year|period)|total comprehensive income", re.I)
_CONT_LINE_RE = re.compile(
    r"^(share (capital|premium)|retained earnings|other reserves|accumulated"
    r"|non-controlling|total equity|current liabilities|non-current liabilities"
    r"|trade and other payables|borrowings|lease liabilit|provisions?|deferred"
    r"|total liabilities|income tax|contract liabilit|total (current|non-current) liabilities)", re.I)


def _has_close(t, rx):
    return any(rx.search(_row_label(r)) for r in t["rows"])


def _stitch_page_breaks(tables):
    """Merge a statement that runs onto the next page.  Only fires when the
    first fragment is missing its closing total AND the next table on the
    following page is an un-headed continuation of balance-sheet / P&L lines."""
    drop = set()
    for i, t in enumerate(tables):
        if id(t) in drop or t["kind"] not in (
                "statement of financial position", "income statement"):
            continue
        rx = _BS_CLOSE_RE if t["kind"] == "statement of financial position" else _PL_CLOSE_RE
        if _has_close(t, rx):
            continue                       # already complete
        for j in range(i + 1, min(i + 4, len(tables))):
            u = tables[j]
            if id(u) in drop or u["file"] != t["file"]:
                continue
            if u["page_label"] not in (t["page_label"], t["page_label"] + 1):
                break
            labs = [_row_label(r) for r in u["rows"][:12] if _row_label(r)]
            if not labs:
                continue
            cont = sum(1 for l in labs if _CONT_LINE_RE.search(l))
            if u.get("header_idx", 0) == 0 and cont >= max(2, len(labs) * 0.5):
                saved = {k: t.get(k) for k in
                         ("rows", "kind", "years", "value_cols", "foots",
                          "foot_detail", "foot_by_col", "health", "notes",
                          "header_idx", "data_start", "total_rows")}
                try:
                    width = max(max(len(r) for r in t["rows"]),
                                max(len(r) for r in u["rows"]))
                    t["rows"] = ([r + [None] * (width - len(r)) for r in t["rows"]]
                                 + [r + [None] * (width - len(r)) for r in u["rows"]])
                    analyze(t, doc_years=([max(t["years"]), max(t["years"]) - 1]
                                          if t.get("years") else None))
                    _attach_health(t)
                    if t["kind"] not in ("statement of financial position",
                                         "income statement"):
                        raise ValueError("merge changed the statement kind")
                    t["_stitched_from"] = u["page_label"]
                    drop.add(id(u))
                except Exception:
                    t.update(saved)          # roll back a bad merge
                break
    if drop:
        tables[:] = [t for t in tables if id(t) not in drop]


def scan(pdfs, pr, min_rows, min_cols, warn=print, progress=None):
    """`progress(page_index0, total_pages)`, called before each page is
    processed -- lets a caller (serve.py) show real "page X of Y" feedback
    during a scan that can otherwise run for a minute or more with no
    visible sign of life. Return a truthy value from it to stop the scan
    early (cooperative cancel) -- whatever's been found so far is returned."""
    all_tables = []
    for pdf_path in pdfs:
        # fiscal year from the file name ('en-2020-...', 'du annual 2013.pdf') --
        # the most reliable period signal for tables that carry no year header
        _fn_yrs = [int(y) for y in re.findall(r"\b(?:19|20)\d{2}\b", pdf_path.stem)
                   if 1990 <= int(y) <= 2035]
        doc_years = [max(_fn_yrs), max(_fn_yrs) - 1] if _fn_yrs else None
        with pdfplumber.open(pdf_path) as pdf:
            if _looks_scanned(pdf):
                warn(f"  !! {pdf_path.name}: little or no extractable text — this "
                     f"looks like a SCANNED PDF. This tool needs a real text layer "
                     f"(no OCR); results will be empty or unreliable.")
            total_pages = len(pdf.pages)
            for i, page in enumerate(pdf.pages):
                if pr is not None and i not in pr:
                    continue
                if progress is not None and progress(i, total_pages):
                    _stitch_page_breaks(all_tables)
                    cross_year_check(all_tables)
                    return all_tables
                words = page.extract_words()
                ptext = page.extract_text() or ""
                # years the page itself talks about ('for the year ended 2021'
                # / column headers) -- used only when a table has no header row
                from collections import Counter
                pcnt = Counter(re.findall(r"\b(?:19|20)\d{2}\b", ptext))
                page_years = [int(y) for y, _ in pcnt.most_common(3)
                              if 1990 <= int(y) <= 2035][:2]
                for bbox, rows, thint in find_all_tables(page, pdf_path):
                    if len(rows) < min_rows or max(len(r) for r in rows) < min_cols:
                        continue
                    rows = _trim_trailing_prose(rows)
                    title = thint or guess_title(page, bbox, words)
                    t = {
                        "file": pdf_path.name, "page": i + 1, "page_label": i + 1,
                        "title": title, "rows": rows, "shape": classify(rows),
                        "bbox": [round(v, 1) for v in bbox],
                        "page_size": [round(page.width, 1), round(page.height, 1)],
                    }
                    analyze(t, page_years=page_years, doc_years=doc_years)
                    _attach_health(t)
                    all_tables.append(t)
    _stitch_page_breaks(all_tables)
    cross_year_check(all_tables)
    return all_tables


def _attach_health(t):
    """Combine label + figure health onto the table."""
    lab = label_health(t)
    fig = figure_health(t)
    t["health_labels"] = lab
    t["health_figures"] = fig
    t["health"] = {
        "score": round(min(lab["score"], fig["score"]), 3),
        "labels": lab["score"], "figures": fig["score"],
        "suspect": ([{**s, "kind": "label"} for s in lab["suspect"]]
                    + [{**s, "kind": "figure"} for s in fig["suspect"]]),
        "scored": max(lab.get("scored", 0), fig.get("scored", 0)),
    }


def _foot_mark(t):
    return {True: "foots", False: "NO FOOT", None: ""}[t.get("foots")]


def print_inventory(tables):
    print(f"\n  {'#':>3}  {'page':>4}  {'size':>9}  {'what':<28}  {'yrs':<9}  {'foots':<8}  title")
    print("  " + "-" * 118)
    for n, t in enumerate(tables, 1):
        ncols = max((len(r) for r in t["rows"]), default=1)
        size = f"{len(t['rows'])}r x {ncols}c"
        yrs = "/".join(str(y) for y in (t.get("years") or [])) or "-"
        print(f"  {n:>3}  {t['page_label']:>4}  {size:>9}  {t['kind']:<28}  {yrs:<9}  "
              f"{_foot_mark(t):<8}  {t['title'][:60]}")
    print()


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pdfs", nargs="+", help="PDF files and/or folders")
    ap.add_argument("--out", default=None, help="output .xlsx (default: <first-pdf>_tables.xlsx)")
    ap.add_argument("--recursive", action="store_true")
    ap.add_argument("--min-rows", type=int, default=2)
    ap.add_argument("--min-cols", type=int, default=2)
    ap.add_argument("--pages", default=None,
                    help="page range, e.g. 40-55 (1-based physical pages)")
    ap.add_argument("--list", action="store_true",
                    help="just print the numbered inventory of detected tables, write nothing")
    ap.add_argument("--only", default=None,
                    help="export only these inventory numbers, e.g.  1,3,5-8  (default: all)")
    ap.add_argument("--all", action="store_true",
                    help="keep tiny / mostly-text fragments too (default: drop them)")
    ap.add_argument("--compare", action="store_true",
                    help="two PDFs: match the same tables across them and write a row-level diff "
                         "(new / removed / changed lines, and a prior-year-column consistency check)")
    ap.add_argument("--audit", action="store_true",
                    help="print the worked arithmetic behind every foots / NO FOOT verdict "
                         "and every label-health warning, then exit")
    ap.add_argument("--json", dest="as_json", default=None, metavar="FILE",
                    help="also write the full inventory (kinds, years, verdicts, health, rows) "
                         "as machine-readable JSON")
    ap.add_argument("-v", "--debug", action="store_true",
                    help="log detection/reconciliation internals (diagnostics)")
    ap.add_argument("--serve", action="store_true",
                    help="open the local preview UI in a browser instead of writing a file")
    args = ap.parse_args()
    if args.debug:
        logging.basicConfig(level=logging.DEBUG, format="  [%(name)s] %(message)s")

    if not HAVE_IMG2TABLE:
        print("  note: img2table not installed -- borderless/2-up statement detection is "
              "running on the regex reconstructor alone (no second-opinion challenger).")
        print("        pip install img2table opencv-python-headless pandas\n")

    if args.serve:
        import serve as _srv
        _srv.run([str(p) for p in gather(args.pdfs, args.recursive)], debug=args.debug)
        return

    pdfs = gather(args.pdfs, args.recursive)
    if not pdfs:
        print("No PDFs found."); sys.exit(1)

    pr = None
    if args.pages:
        a, _, b = args.pages.partition("-")
        pr = range(int(a) - 1, (int(b) if b else int(a)))

    # ---- compare mode --------------------------------------------------------
    if args.compare:
        if len(pdfs) != 2:
            print("--compare needs exactly two PDFs (newer first, older second)."); sys.exit(1)
        print(f"Scanning {pdfs[0].name} ...")
        a_tabs = [t for t in scan([pdfs[0]], pr, args.min_rows, args.min_cols)]
        print(f"Scanning {pdfs[1].name} ...")
        b_tabs = [t for t in scan([pdfs[1]], pr, args.min_rows, args.min_cols)]
        pairs = _match_tables(a_tabs, b_tabs)
        if not pairs:
            print("No comparable tables matched between the two files."); sys.exit(0)
        print(f"\n  matched {len(pairs)} table(s):")
        for ta, tb in pairs:
            drows, verdict, _counts = diff_tables(ta, tb)
            print(f"    {ta['kind']:<32} {ta.get('title','')[:40]:<42} -> {verdict}")
        out = Path(args.out) if args.out else pdfs[0].with_name(
            pdfs[0].stem + "_vs_" + pdfs[1].stem + "_diff.xlsx")
        build_compare_workbook(pairs, (pdfs[0].name, pdfs[1].name)).save(out)
        print(f"\n  diff -> {out}")
        return

    print(f"Scanning {len(pdfs)} PDF(s)...")
    tables = scan(pdfs, pr, args.min_rows, args.min_cols)
    if not tables:
        print("No tables detected."); sys.exit(0)

    if not args.all:
        def _junk(t):
            if t["kind"] in ("income statement", "statement of financial position",
                             "statement of cash flows", "statement of changes in equity",
                             "note"):
                return False
            return t.get("shape") == "mostly-text" or len(t["rows"]) < 3
        drop = [t for t in tables if _junk(t)]
        tables = [t for t in tables if t not in drop]
        if drop:
            print(f"  ({len(drop)} tiny/mostly-text fragment(s) hidden - use --all to keep them)")
        if not tables:
            print("Only fragments detected. Re-run with --all to see them."); sys.exit(0)

    print_inventory(tables)

    if args.audit:
        _print_audit(tables)
        return

    if args.as_json:
        _write_json(tables, Path(args.as_json))
        print(f"  inventory JSON -> {args.as_json}")

    if args.list:
        print(f"  {len(tables)} table(s).  Re-run with  --only <numbers>  to export a subset,")
        print("  or without --list to export all.")
        return

    picked = tables
    if args.only:
        sel = _parse_selection(args.only, len(tables))
        if not sel:
            print("  --only matched nothing."); sys.exit(1)
        picked = [tables[i - 1] for i in sel]
        print(f"  exporting {len(picked)} of {len(tables)}: {sel}")

    out = Path(args.out) if args.out else pdfs[0].with_name(pdfs[0].stem + "_tables.xlsx")
    build_workbook(picked).save(out)
    print(f"\n  {len(picked)} table(s) -> {out}")


def _print_audit(tables):
    print("\n  AUDIT — the arithmetic behind each verdict\n  " + "=" * 70)
    for n, t in enumerate(tables, 1):
        if t["kind"] not in ("income statement", "statement of financial position",
                             "statement of cash flows", "statement of changes in equity",
                             "note"):
            continue
        v = {True: "FOOTS ✓", False: "NO FOOT ✗", None: "no check"}[t.get("foots")]
        print(f"\n  [{n}] p{t['page_label']}  {t['kind']}  —  {t.get('title','')[:56]}")
        print(f"       verdict : {v}")
        for nt in t.get("notes") or []:
            print(f"       ! note  : {nt}")
        for fc in t.get("foot_by_col") or []:
            mark = {True: "✓", False: "✗", None: "?"}[fc.get("ok")]
            print(f"       {str(fc['year'])[:8]:>8} {mark} : {fc['worked']}")
        if not t.get("foot_by_col") and t.get("foot_detail"):
            print(f"       working : {t['foot_detail']}")
        cc = t.get("consistency")
        if cc:
            v = cc.get("verdict") or ("agree" if cc["mismatch"] == 0 else "differs")
            tag = "agree ✓" if cc["mismatch"] == 0 else f"{cc['mismatch']} of {cc['checked']} differ — {v}"
            print(f"       x-year  : {cc['year']} column vs {cc['vs']}: {tag}")
            for k, a, b in cc.get("worst", []):
                print(f"                 · {k}: {a:,.0f} here vs {b:,.0f} there")
        h = t.get("health") or {}
        if h.get("suspect"):
            print(f"       health  : labels {h.get('labels', 1):.0%} / figures "
                  f"{h.get('figures', 1):.0%} — {len(h['suspect'])} suspect row(s):")
            for s in h["suspect"][:8]:
                print(f"                 · [{s.get('kind','?')}] “{s['label'][:56]}” — {s['why']}")
    print()


def _write_json(tables, path):
    import json
    out = []
    for n, t in enumerate(tables, 1):
        out.append({
            "n": n, "file": t["file"], "page": t["page_label"], "kind": t["kind"],
            "title": t.get("title"), "years": t.get("years"),
            "foots": t.get("foots"), "foot_detail": t.get("foot_detail", ""),
            "foot_by_col": t.get("foot_by_col"),
            "consistency": t.get("consistency"),
            "notes": t.get("notes"),
            "health": t.get("health"),
            "value_cols": t.get("value_cols"), "header_idx": t.get("header_idx"),
            "data_start": t.get("data_start"), "total_rows": t.get("total_rows"),
            "stitched_from": t.get("_stitched_from"),
            "rows": t["rows"],
        })
    path.write_text(json.dumps(out, indent=1, ensure_ascii=False, default=str),
                    encoding="utf-8")


if __name__ == "__main__":
    main()
