"""
Pull exactly TWO things out of large annual-report PDFs into one Excel file:

    1. "CONSOLIDATED STATEMENT OF PROFIT OR LOSS" -> the whole statement,
       from the top (Revenue, etc.) down through and including the
       "Profit for the year" row. Everything after that row (attributable-to
       splits, other comprehensive income, EPS, etc.) is left out.
    2. "General and administrative expenses"      -> the whole table

Built for the case where the report calls these things slightly different
names in different years (e.g. "CONSOLIDATED STATEMENT OF PROFIT OR LOSS"
vs "CONSOLIDATED INCOME STATEMENT", "Profit for the year" vs "Net profit
for the year" vs "Profit for the financial year"). Matching is done with
regex/keyword patterns, not exact strings -- see TARGETS below.

It reuses the same table-extraction engine as your pdf_to_excel_v2.py
(de-fragmenting split tables, filling gaps where ruling lines are missing,
filtering out sidebar/chart junk) -- only the "which page / which table"
logic is new, so it only ever processes the handful of pages that actually
matter instead of the whole report.

-----------------------------------------------------------------------
SETUP (run once):
    pip install pdfplumber openpyxl

USAGE:
    python extract_two_tables.py "report_2025.pdf" "report_2010.pdf"

    Optional page hints (just speeds up the search / helps it pick the
    right occurrence if a phrase appears more than once -- it still scans
    the WHOLE file regardless, so an unknown or wrong hint is harmless):
    python extract_two_tables.py report_2025.pdf --pl-hint 57 --admin-hint 60
    python extract_two_tables.py report_2010.pdf --pl-hint 70 --admin-hint 75

    Each PDF becomes its own sheet in one workbook:
        <output folder>/two_tables_extracted.xlsx
    Optional output path as the last thing:
    python extract_two_tables.py a.pdf b.pdf --out combined.xlsx

If a target can't be found in a file, the script prints the closest partial
matches it *did* see (console) and writes a "NOT FOUND" note in the sheet,
so you can see what the report actually calls it and tighten TARGETS below.
-----------------------------------------------------------------------
"""

import sys
import re
import argparse
import difflib
import tempfile
from collections import Counter
from pathlib import Path

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    from pypdf import PdfReader, PdfWriter
    HAVE_PYPDF = True
except ImportError:
    HAVE_PYPDF = False

# ---------------------------------------------------------------- styling ---
# (same look as pdf_to_excel_v2.py so output is consistent with what you had)

FONT_NAME = "Calibri"
TITLE_FILL = PatternFill("solid", fgColor="1F3864")
SECTION_FILL = PatternFill("solid", fgColor="D9E2F3")
HEADER_FILL = PatternFill("solid", fgColor="2E5395")
ROW_FILL_ALT = PatternFill("solid", fgColor="F2F5FB")
NOT_FOUND_FILL = PatternFill("solid", fgColor="FCE4E4")

TITLE_FONT = Font(name=FONT_NAME, size=16, bold=True, color="FFFFFF")
SUBTITLE_FONT = Font(name=FONT_NAME, size=10, italic=True, color="595959")
SECTION_FONT = Font(name=FONT_NAME, size=11, bold=True, color="1F3864")
HEADER_FONT = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
BODY_FONT = Font(name=FONT_NAME, size=10)
NOTE_FONT = Font(name=FONT_NAME, size=9, italic=True, color="808080")
NOT_FOUND_FONT = Font(name=FONT_NAME, size=10, italic=True, color="C00000")

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

MIN_TABLE_WIDTH = 45
MIN_TABLE_HEIGHT = 8
FULL_PAGE_COVERAGE = 0.9
MIN_FILL_RATIO = 0.15
BOX_CAPTURE_ROWS = 3
BOX_CAPTURE_RATIO = 0.6
BBOX_TOLERANCE = 6
CAPTION_MAX_GAP = 26
LINE_TOLERANCE = 3      # points; words within this of each other's "top" are one line

# ----------------------------------------------------------- WHAT TO FIND ---
# Edit these patterns if a report uses wording that doesn't match. Patterns
# are tried in order; the first one that matches a page/row wins.

# Minimum fraction of a target's reference_row_labels that must fuzzy-match
# a candidate table's row labels before that table is trusted. Tune down if
# a genuinely correct table is being rejected (check the console score),
# tune up if the wrong table is being picked.
MIN_CONTENT_SCORE = 0.35
FUZZY_ROW_MATCH_THRESHOLD = 0.72   # how close two row labels must be to count as "the same row"

TARGETS = [
    {
        "name": "Consolidated Statement of Profit or Loss (through Net profit for the year)",
        "mode": "table_until_row",
        "heading_patterns": [
            re.compile(r"CONSOLIDATED\s+STATEMENT\s+OF\s+PROFIT\s+(OR|AND)\s+LOSS", re.I),
            re.compile(r"STATEMENT\s+OF\s+PROFIT\s+(OR|AND)\s+LOSS", re.I),
            re.compile(r"CONSOLIDATED\s+INCOME\s+STATEMENT", re.I),
        ],
        # tried in order (exact-ish); first row whose label matches ends the table
        "row_patterns": [
            re.compile(r"^NET\s+PROFIT\s+FOR\s+THE\s+(YEAR|PERIOD)\s*$", re.I),
            re.compile(r"^PROFIT\s+FOR\s+THE\s+(YEAR|PERIOD)\s*$", re.I),
            re.compile(r"^(NET\s+)?PROFIT\s+FOR\s+THE\s+(FINANCIAL\s+)?(YEAR|PERIOD)\b", re.I),
        ],
        # fuzzy fallback if none of the row_patterns above match exactly
        "cutoff_reference_labels": ["Net profit for the year", "Profit for the year"],
        # confirmed row labels from a known-good year (2023/2022 report) --
        # used to score candidate tables when the heading wording differs
        "reference_row_labels": [
            "Trading commission fees", "Brokerage fees",
            "Clearing, settlement and depository fees", "Listing and market data fees",
            "Other fees", "Investment income", "Dividend income", "Other income",
            "Total income", "General and administrative expenses",
            "Amortisation of other intangible assets", "Interest expense",
            "Total expenses", "Net profit for the year", "Owners of the Company",
            "Non-controlling interest",
        ],
        "hint_key": "pl_hint",
    },
    {
        "name": "General and administrative expenses",
        "mode": "full_table",
        "heading_patterns": [
            re.compile(r"GENERAL\s+AND\s+ADMINISTRATIVE\s+EXPENSES", re.I),
            re.compile(r"GENERAL\s*&\s*ADMINISTRATIVE\s+EXPENSES", re.I),
            re.compile(r"ADMINISTRATIVE\s+(AND\s+GENERAL\s+)?EXPENSES", re.I),
        ],
        "reference_row_labels": [
            "Payroll and other benefits", "Depreciation", "Maintenance expenses",
            "Telecommunication expenses", "Professional expenses",
            "Board of Directors remuneration and expenses", "Other expenses",
        ],
        "hint_key": "admin_hint",
    },
]

# --------------------------------------------------- extraction (reused) ---
# Identical logic to pdf_to_excel_v2.py's engine.

# ---------------------------------------------------- printed page labels ---
# Many reports restart their page numbering partway through the PDF (front
# matter numbered separately from the report body), so the physical position
# of a page in the PDF file can be quite different from the page number
# actually printed on it. When the PDF defines proper /PageLabels metadata,
# we use that to report the number a human reading the document would see,
# rather than just the raw physical page index.

def load_page_labels(pdf_path):
    """List of printed page-label strings, one per physical page (index 0 =
    physical page 1), or None if unavailable/unreadable."""
    if not HAVE_PYPDF:
        return None
    try:
        labels = PdfReader(str(pdf_path)).page_labels
        return list(labels) if labels else None
    except Exception:
        return None


def display_page(labels, physical_page_1_based):
    """Human-facing page string for a 1-based physical page index. Shows the
    printed label when it's known and differs from the physical index (e.g.
    "54 (PDF page 61)"), otherwise just the physical page number."""
    if labels and 0 <= physical_page_1_based - 1 < len(labels):
        label = labels[physical_page_1_based - 1]
        if label and label != str(physical_page_1_based):
            return f"{label} (PDF page {physical_page_1_based})"
    return str(physical_page_1_based)


def clean_cell(value):
    if value is None:
        return None
    text = str(value).strip().replace("\n", " ")
    if text == "":
        return None
    stripped = text.replace(",", "").replace("$", "").replace("%", "").strip()
    negative = stripped.startswith("(") and stripped.endswith(")")
    if negative:
        stripped = stripped[1:-1]
    if re.fullmatch(r"-?\d+(\.\d+)?", stripped) and stripped not in ("", "-"):
        num = float(stripped)
        if negative:
            num = -num
        return int(num) if num.is_integer() else num
    return text


def _reference_column_bounds(table):
    best_row = max(table.rows, key=lambda r: sum(1 for c in r.cells if c is not None))
    return [(c[0], c[2]) for c in best_row.cells if c is not None]


def robust_extract(page, table, all_words=None):
    rows = table.extract()
    if all_words is None:
        all_words = page.extract_words()
    try:
        col_bounds = _reference_column_bounds(table)
    except ValueError:
        return rows

    for ri, trow in enumerate(table.rows):
        if ri >= len(rows):
            break
        row_cells = trow.cells
        r_top, r_bot = trow.bbox[1], trow.bbox[3]
        for ci in range(len(rows[ri])):
            if rows[ri][ci] not in (None, ""):
                continue
            if ci >= len(col_bounds):
                continue
            c0, c1 = col_bounds[ci]
            covered = False
            for other in row_cells:
                if other is None:
                    continue
                ox0, otop, ox1, obot = other
                if ox0 < c1 - 1 and ox1 > c0 + 1 and not (abs(ox0 - c0) < 2 and abs(ox1 - c1) < 2):
                    covered = True
                    break
            if covered:
                continue
            words = [
                w for w in all_words
                if c0 <= (w["x0"] + w["x1"]) / 2 <= c1 and r_top <= (w["top"] + w["bottom"]) / 2 <= r_bot
            ]
            if not words:
                continue
            words.sort(key=lambda w: (round(w["top"]), w["x0"]))
            rows[ri][ci] = " ".join(w["text"] for w in words)
    return attach_left_labels(page, table, rows, all_words)


MAX_LABEL_LOOKUP_DIST = 300     # pts to search left of a numbers-only table for its row labels
MAX_TRAILING_LOOKUP_DIST = 90   # pts to search right for a missed adjacent numeric column (kept
                                 # tight -- unlike labels, a same-statement number column sits close
                                 # by; anything farther is more likely an unrelated neighboring box)
MAX_TRAILING_LOOKUP_DIST_WIDE = 250  # fallback if nothing found within the tight distance -- some
                                      # wide landscape layouts leave a much bigger gap between a
                                      # ruled label box and its numbers; still safe since only the
                                      # nearest contiguous cluster found is ever used


def attach_left_labels(page, table, rows, all_words):
    """
    Some reports rule a box around only the numbers in a statement/note
    (the row-label text sits further left, outside any ruled box, so
    line-based table detection never captures it). If a table's rows are
    mostly unlabeled, look for text immediately to the left of the table
    -- within each row's own vertical band, capped at MAX_LABEL_LOOKUP_DIST
    so we don't reach across into an adjacent column on a multi-column
    page -- and prepend it as a label column.
    """
    if not rows:
        return rows

    def has_text_label(row):
        # row cells here are still raw (uncleaned) strings from table.extract();
        # use clean_cell so a numeric-looking string ("255,254") doesn't count
        # as a label the way a real row label ("Trading commission fees") does.
        return any(isinstance(clean_cell(c), str) for c in row)

    def has_real_label(row):
        # Stricter than has_text_label: a short note-reference token like
        # "15, 16" or "20" is technically a string too, but it isn't a real
        # descriptive label -- nor is header-ish content like "2019
        # AED'000". A row whose only "text" is one of these still needs
        # (and should get) a left-label search; a row that already carries
        # genuine descriptive words (e.g. "Interest expense") shouldn't
        # have one attempted at all.
        for c in row:
            v = clean_cell(c)
            if not isinstance(v, str):
                continue
            s = v.strip()
            if s == "-" or re.fullmatch(r"[\d,.\s]+", s):
                continue
            if looks_like_column_header([s]):
                continue
            return True
        return False

    labeled = sum(1 for r in rows if has_text_label(r))
    needs_labels = labeled / len(rows) < 0.5

    # Independently of whether labels are missing: if most rows have no
    # values at all (just a label and nothing else -- e.g. a ruled box that
    # only encloses the label column, with the actual number columns sitting
    # entirely outside its bbox), values need to be recovered too. Both
    # checks run independently since a table can be missing one, the other,
    # or both.
    def numeric_count(row):
        return sum(1 for c in row if clean_cell(c) is not None and not isinstance(clean_cell(c), str))

    with_values = [r for r in rows if numeric_count(r) > 0]
    avg_numeric = (sum(numeric_count(r) for r in with_values) / len(with_values)) if with_values else 0
    # Most financial statement rows carry the same number of value columns
    # (e.g. 2: current + prior year). Establish that "typical" count from
    # the rows that already have values, then later check each row
    # INDIVIDUALLY against it -- a table-wide average can look fine even
    # when one specific row (e.g. a text header row like "2018 AED'000",
    # which itself carries no numeric cells) is short a column that every
    # data row already has.
    from collections import Counter
    typical_cols = Counter(numeric_count(r) for r in with_values).most_common(1)[0][0] if with_values else 0
    needs_values = avg_numeric < 1.5 or typical_cols > 1

    if not needs_labels and not needs_values:
        return rows

    x_limit_left = max(0, table.bbox[0] - MAX_LABEL_LOOKUP_DIST)
    x_limit_right_wide = min(page.width, table.bbox[2] + MAX_TRAILING_LOOKUP_DIST_WIDE)
    out = []
    for ri, trow in enumerate(table.rows):
        if ri >= len(rows):
            break
        r_top, r_bot = trow.bbox[1], trow.bbox[3]
        row_needs_label = (
            needs_labels and not has_real_label(rows[ri]) and not looks_like_column_header(rows[ri])
        )
        label_words = []
        if row_needs_label:
            label_words = [
                w for w in all_words
                if x_limit_left <= w["x0"] < table.bbox[0] - 1
                and r_top <= (w["top"] + w["bottom"]) / 2 <= r_bot
            ]
        # Some reports rule a box tight around just ONE numeric column
        # (e.g. only the current year), leaving a second column (e.g. the
        # prior year, for comparison) sitting just outside the box to the
        # right, uncaptured. Pick up any such trailing text/numbers too,
        # grouping by x-gaps in case there's more than one extra column.
        # Only bother for a row that's individually short of the table's
        # typical column count (see comment above) -- an already-complete
        # row doesn't need (and shouldn't risk) a trailing-content search.
        row_needs_values = needs_values and (
            avg_numeric < 1.5 or typical_cols == 0 or numeric_count(rows[ri]) < typical_cols
        )
        trailing_words = []
        if row_needs_values:
            trailing_words = [
                w for w in all_words
                if table.bbox[2] + 1 < w["x1"] <= x_limit_right_wide
                and r_top <= (w["top"] + w["bottom"]) / 2 <= r_bot
            ]
        row = list(rows[ri])
        label_x0 = None
        # Always reserve the label slot -- even when this particular row has
        # no label (e.g. an unlabeled totals row), so its numbers stay in
        # the same columns as every other row's numbers instead of shifting
        # left into where the label would have been. But only actually
        # SEARCH for one when this row doesn't already carry its own text
        # (row_needs_label) -- otherwise (e.g. a header-ish row that
        # already reads "2019 AED'000") the search has nothing genuine to
        # find and can instead grab unrelated text from a nearby column
        # that merely shares this row's vertical band.
        if needs_labels:
            if row_needs_label and label_words:
                label_words.sort(key=lambda w: w["x0"])
                # Only the contiguous run immediately adjacent to the table
                # (reading left-to-right, ending right at the table edge) --
                # on a densely packed multi-column page, words further left
                # within the search window can belong to an entirely
                # different column/note that happens to share this row's
                # vertical band, not to this label.
                clusters, cur = [], [label_words[0]]
                for w in label_words[1:]:
                    if w["x0"] - cur[-1]["x1"] > 15:
                        clusters.append(cur)
                        cur = []
                    cur.append(w)
                clusters.append(cur)
                last_text = " ".join(w["text"] for w in clusters[-1]).strip()
                if len(clusters) > 1 and re.fullmatch(r"\d{1,3}", last_text):
                    # The gap-separated cluster right next to the table is
                    # just a short note-reference number (e.g. "...income
                    # <gap> 16"), not a real second label. Keep it as its
                    # OWN cell (a genuine Notes-column value, consistent
                    # with how every other row's note reference is stored)
                    # rather than gluing it onto the label text -- any
                    # resulting column-count mismatch against rows with no
                    # note of their own is straightened out afterward by
                    # normalize_missing_notes_column.
                    label_cell = " ".join(w["text"] for w in clusters[0])
                    label_x0 = clusters[0][0]["x0"]
                    row = [label_cell, last_text] + row
                else:
                    nearest = clusters[-1]
                    label_x0 = nearest[0]["x0"]
                    row = [" ".join(w["text"] for w in nearest)] + row
            else:
                row = [None] + row
        if trailing_words:
            # First separate into visual lines (a header can stack "2017"
            # above "AED'000" on two lines within the same row band) --
            # sorting everything by x0 alone, ignoring top, would wrongly
            # merge overlapping-x words from different lines into one
            # nonsensical cluster. Then within each line, cluster by x-gap
            # as before.
            tw_lines = build_lines(trailing_words, x_gap_split=15)
            segments = sorted(
                ({"top": l["top"], "x0": l["x0"], "x1": l["x1"], "text": l["text"]} for l in tw_lines),
                key=lambda s: s["x0"],
            )
            # Group segments into columns by x-overlap (regardless of which
            # line/top they're on), then join same-column segments top to
            # bottom -- this is what lets "2017" (line 1) and "AED'000"
            # (line 2) combine into one header cell instead of colliding
            # with each other or with unrelated same-top text.
            clusters = []
            for seg in segments:
                placed = False
                for c in clusters:
                    if not (seg["x1"] < c["x0"] - 5 or seg["x0"] > c["x1"] + 5):
                        c["parts"].append(seg)
                        c["x0"] = min(c["x0"], seg["x0"])
                        c["x1"] = max(c["x1"], seg["x1"])
                        placed = True
                        break
                if not placed:
                    clusters.append({"x0": seg["x0"], "x1": seg["x1"], "parts": [seg]})
            clusters.sort(key=lambda c: c["x0"])
            for c in clusters:
                c["parts"].sort(key=lambda s: s["top"])
                c["text"] = " ".join(p["text"] for p in c["parts"])

            def looks_like_value(cluster):
                text = cluster["text"].strip()
                if not text or len(text) > 24:
                    return False
                if re.fullmatch(r"[\d,.\-\(\)%\s]+", text):
                    return True
                # A standalone currency/unit marker (e.g. a lone "AED'000"
                # cell -- its paired year sits on the header row above, not
                # this row), or one stacking a year with it (e.g. "2017" /
                # "AED'000" combined), is also a legitimate value cell.
                return bool(re.search(r"AED|USD|EGP|SAR|['\u2019]000", text))

            kept = [clusters[0]] if looks_like_value(clusters[0]) else []
            for c in clusters[1:]:
                if len(kept) >= 3:  # cap: e.g. note-ref + current year + prior year
                    break
                if not kept:
                    break  # first cluster wasn't value-like -- don't guess at later ones
                if c["x0"] - kept[-1]["x1"] > 95:
                    break
                if not looks_like_value(c):
                    break
                kept.append(c)
            row += [c["text"] for c in kept]
        out.append((row, label_x0))

    # Consistency check: most rows' labels should start at roughly the same
    # x0 (a table's label column is left-aligned). A row whose label starts
    # far from that common position is very likely contamination from an
    # unrelated column on a densely packed page -- drop just that label
    # (keep its values) rather than risk showing wrong text.
    if needs_labels:
        x0s = sorted(lx for _, lx in out if lx is not None)
        if len(x0s) >= 2:
            median_x0 = x0s[len(x0s) // 2]
            cleaned = []
            for row, lx in out:
                if lx is not None and abs(lx - median_x0) > 80:
                    row = [None] + row[1:]
                cleaned.append(row)
            return cleaned
    return [row for row, _ in out]


def is_junk_table(table, page):
    x0, top, x1, bottom = table.bbox
    width = x1 - x0
    height = bottom - top
    if width < MIN_TABLE_WIDTH:
        return True
    if height < MIN_TABLE_HEIGHT:
        return True
    if width >= page.width * FULL_PAGE_COVERAGE and height >= page.height * FULL_PAGE_COVERAGE:
        return True

    rows = table.extract()
    total_cells = sum(len(r) for r in rows)
    filled_cells = sum(1 for r in rows for c in r if c and str(c).strip())
    fill_ratio = filled_cells / total_cells if total_cells else 0
    if fill_ratio < MIN_FILL_RATIO:
        return True

    if len(rows) <= BOX_CAPTURE_ROWS:
        words_in_bbox = [
            w for w in page.extract_words()
            if x0 <= (w["x0"] + w["x1"]) / 2 <= x1 and top <= (w["top"] + w["bottom"]) / 2 <= bottom
        ]
        if words_in_bbox:
            captured_words = sum(len(str(c).split()) for r in rows for c in r if c)
            capture_ratio = captured_words / len(words_in_bbox)
            if capture_ratio < BOX_CAPTURE_RATIO:
                return True
    return False


def is_title_row(row):
    filled = [c for c in row if c is not None and str(c).strip() != ""]
    if len(filled) != 1:
        return False
    value = filled[0]
    if isinstance(value, (int, float)):
        return False
    text = str(value).strip()
    if not (1 <= len(text.split()) <= 10):
        return False
    if text.endswith((".", ",", ";")):
        return False
    return True


def get_heading_line(words_pool, x0, x1, top, max_gap=CAPTION_MAX_GAP, x_tolerance=15):
    candidates = [
        w for w in words_pool
        # small tolerance: a heading sitting immediately flush against the
        # table below it can have bottom fractionally > top due to font
        # metric rounding, which would otherwise wrongly exclude it
        if w["bottom"] <= top + 2 and (top - w["bottom"]) <= max_gap
        and w["x0"] >= x0 - x_tolerance and w["x1"] <= x1 + x_tolerance
    ]
    if not candidates:
        return None
    line_top = max(w["top"] for w in candidates)
    line_words = [w for w in candidates if abs(w["top"] - line_top) < 2]
    line_words.sort(key=lambda w: w["x0"])
    return " ".join(w["text"] for w in line_words).strip()


def find_actual_heading(page, table_top, x0=None, x1=None, look_up_limit=350):
    """
    Best-effort guess at what the report itself calls the section
    immediately above a matched table -- useful for reporting to the user
    when a target was located by content match rather than an exact
    heading-pattern match (e.g. reference wording says "Statement of
    Profit or Loss" but the report actually prints "Income Statement").

    Heuristic: within a window above the table (restricted to roughly the
    table's own column, if x0/x1 given, so a heading in an adjacent
    column on a multi-column page isn't picked up), the heading text is
    whatever line(s) use the single largest font size in that window
    (headings are reliably bigger than the surrounding body text, even
    though the exact size varies by report and by heading level). Lines
    sharing that top size are merged in reading order to catch wrapped
    two-line titles.
    """
    try:
        words = page.extract_words(extra_attrs=["size"])
    except Exception:
        return None
    if not words:
        return None

    # Pass 1 -- tight column: a "<N>. Title" note heading normally starts
    # right at (or very near) the table's own left edge, not far off to the
    # side, so search a narrow band around the table's actual x-range. This
    # avoids grabbing a *different* note's heading that happens to sit at a
    # similar height in a neighboring column on a multi-column notes page.
    if x0 is not None and x1 is not None:
        tight_lo, tight_hi = x0 - 40, x1 + 40
        tight_words = [
            w for w in words
            if table_top - look_up_limit <= w["top"] < table_top - 2
            and tight_lo <= w["x0"] <= tight_hi
        ]
        tight_lines = build_lines(tight_words, x_gap_split=18)
        note_heading_lines = [
            l for l in tight_lines if re.match(r"^\d{1,2}\.\s+[A-Z]", normalize(l["text"]))
        ]
        if note_heading_lines:
            note_heading_lines.sort(key=lambda l: -l["top"])
            best = note_heading_lines[0]
            # A long heading can overflow slightly past the tight window on
            # its OWN line (not a wrapped second line) -- rebuild that first
            # line using a moderately wider, same-top-only word pool so a
            # trailing word isn't cut off by the tight boundary.
            same_top_words = [
                w for w in words
                if abs(w["top"] - best["top"]) <= LINE_TOLERANCE
                and best["x0"] - 2 <= w["x0"] <= tight_hi + 100
            ]
            if same_top_words:
                same_top_words.sort(key=lambda w: w["x0"])
                extended = [same_top_words[0]]
                for w in same_top_words[1:]:
                    if w["x0"] - extended[-1]["x1"] > 18:
                        break
                    extended.append(w)
                if len(extended) > len(best["words"]):
                    best = {"top": best["top"], "words": extended,
                             "text": " ".join(w["text"] for w in extended), "x0": extended[0]["x0"]}
            # Merge a wrapped continuation line (heading text that wraps onto
            # a second line, e.g. "18. General and" / "administrative
            # expenses"). A slightly wider (but still x0-aligned) word pool
            # is used here only -- a long title can wrap past the table's
            # own right edge, but a genuine continuation still starts at
            # essentially the same x0 as the heading itself, which is what
            # actually guards against pulling in a neighboring column.
            continuation_words = [
                w for w in words
                if table_top - look_up_limit <= w["top"] < table_top - 2
                and tight_lo <= w["x0"] <= tight_hi + 120
            ]
            continuation_lines = build_lines(continuation_words, x_gap_split=18)
            merged_text = best["text"]
            cursor = best
            while True:
                below = [
                    l for l in continuation_lines
                    if l is not cursor and 0 < l["top"] - cursor["top"] <= 20
                    and abs(l["x0"] - cursor["x0"]) <= 20
                ]
                if not below:
                    break
                below.sort(key=lambda l: l["top"])
                nxt = below[0]
                if re.match(r"^\d{1,2}\.\s+[A-Z]", normalize(nxt["text"])):
                    break
                merged_text += " " + nxt["text"]
                cursor = nxt
            return normalize(merged_text) or None

    # Pass 2 -- fallback for headings without note numbering (e.g. a full
    # financial statement's own title, which can be indented well to the
    # left of just its numbers column) -- widen the search and go by
    # largest font size in the window instead.
    x_lo = max(0, (x0 if x0 is not None else 0) - MAX_LABEL_LOOKUP_DIST)
    x_hi = (x1 if x1 is not None else page.width) + 50
    window_words = [
        w for w in words
        if table_top - look_up_limit <= w["top"] < table_top - 2
        and x_lo <= w["x0"] <= x_hi
    ]
    if not window_words:
        return None
    lines = build_lines(window_words)
    if not lines:
        return None

    for line in lines:
        line["size"] = max(w["size"] for w in line["words"])
    max_size = max(line["size"] for line in lines)
    same_size = [l for l in lines if l["size"] >= max_size - 0.5]
    # There can be several same-sized headings/lines in the window on a
    # multi-column page (e.g. another statement's title sitting at the same
    # vertical position in a different column) -- keep only lines that are
    # BOTH vertically close to the table AND horizontally overlapping/near
    # the nearest such line, so an unrelated column's text at a coincidentally
    # similar height doesn't get stitched in. When lines tie on vertical
    # position (same top, different columns), prefer whichever is actually
    # x-aligned with our own table over an arbitrary list-order tie-break.
    table_x0 = x0 if x0 is not None else 0
    same_size.sort(key=lambda l: (-l["top"], abs(l["x0"] - table_x0)))
    anchor = same_size[0]
    same_column = [
        l for l in same_size
        if not (l["x1"] < anchor["x0"] - 30 or l["x0"] > anchor["x1"] + 30)
    ]
    same_column.sort(key=lambda l: -l["top"])
    cluster = [same_column[0]]
    for line in same_column[1:]:
        if cluster[-1]["top"] - line["top"] > 40:
            break
        cluster.append(line)
    cluster.sort(key=lambda l: l["top"])
    text = normalize(" ".join(l["text"] for l in cluster))
    return text or None


def is_note_boundary(page, words_pool, x0, x1, top, max_gap=150):
    """
    True if a "<N>. <Title>" note heading appears anywhere in the gap
    directly above this fragment (not just on the single nearest line --
    a header row like "2017 / AED'000" often sits between the note
    heading and the fragment's own first data row). x0/x1 should be the
    table's OWN bounds (not pre-widened) -- a tight column-restricted
    check is tried first, widening only if that finds nothing, so a
    neighboring column's heading on a multi-column page doesn't bleed in.
    """
    def check(lo, hi):
        candidates = [
            w for w in words_pool
            # same small tolerance as get_heading_line -- see comment there
            if w["bottom"] <= top + 2 and (top - w["bottom"]) <= max_gap
            and w["x0"] >= lo - 15 and w["x1"] <= hi + 15
        ]
        if not candidates:
            return False
        for line in build_lines(candidates, x_gap_split=18):
            if re.match(r"^\d{1,2}\.\s+[A-Z]", normalize(line["text"])):
                return True
        return False

    if check(x0, x1):
        return True
    return check(max(0, x0 - MAX_LABEL_LOOKUP_DIST), x1)


def group_and_merge_fragments(page, tables):
    groups = []
    for t in tables:
        x0, top, x1, bottom = t.bbox
        placed = False
        for g in groups:
            if abs(g["x0"] - x0) <= BBOX_TOLERANCE and abs(g["x1"] - x1) <= BBOX_TOLERANCE:
                g["tables"].append(t)
                placed = True
                break
        if not placed:
            groups.append({"x0": x0, "x1": x1, "tables": [t]})

    all_words = page.extract_words()
    merged = []
    for g in groups:
        g["tables"].sort(key=lambda t: t.bbox[1])
        subgroups = [[]]
        for t in g["tables"]:
            extracted = robust_extract(page, t, all_words)
            first_row = [clean_cell(c) for c in extracted[0]] if extracted else []
            starts_new = subgroups[-1] and (
                is_title_row(first_row)
                # search a wide left margin, not just the narrow fragment's own
                # x-range -- some reports put the note-number heading far to
                # the left of a box that only rules around its numbers column
                or is_note_boundary(page, all_words, g["x0"], g["x1"], t.bbox[1])
            )
            if starts_new:
                subgroups.append([])
            subgroups[-1].append(t)

        for sub in subgroups:
            fragment_count = len(sub)
            # Some reports rule a box around every OTHER line, leaving the
            # line in between completely unboxed -- that row is invisible
            # to every table-detection path, not just this one. Detect any
            # vertical gap between two consecutive ruled fragments that
            # still has real words in it, and reconstruct that row from
            # word positions the same way an entirely-unruled table would
            # be, using a nearby fragment's own label/value split as the
            # column boundary.
            label_x_end = None
            for t in sub:
                cells = t.rows[0].cells if t.rows else None
                if cells and cells[0]:
                    label_x_end = cells[0][2]
                    break
            if label_x_end is None:
                label_x_end = g["x0"] + 0.55 * (g["x1"] - g["x0"])

            entries = [(t.bbox[1], t.bbox[3], t) for t in sub]  # (top, bottom, table)
            gap_rows = []  # (top, extracted_rows)
            for (top_a, bot_a, _), (top_b, bot_b, _) in zip(entries, entries[1:]):
                if top_b - bot_a < 3:
                    continue
                gap_words = [
                    w for w in all_words
                    if bot_a <= w["top"] < top_b and g["x0"] - 5 <= w["x0"] <= g["x1"] + 5
                ]
                if not gap_words:
                    continue
                recon = reconstruct_unruled_rows(page, g["x0"] - 5, g["x1"] + 5, bot_a, top_b, label_x_end)
                if recon:
                    gap_rows.append((bot_a, recon))

            rows = []
            extracted_by_table = [(t.bbox[1], robust_extract(page, t, all_words)) for t in sub]
            combined = [(top, ext) for top, ext in extracted_by_table] + gap_rows
            combined.sort(key=lambda item: item[0])
            all_extracted = [ext for _, ext in combined]
            max_cols = max((len(row) for ext in all_extracted for row in ext), default=0)
            for ext in all_extracted:
                for row in ext:
                    row = list(row)
                    if len(row) < max_cols:
                        row += [None] * (max_cols - len(row))
                    cleaned = [clean_cell(c) for c in row]
                    if any(c is not None and str(c).strip() != "" for c in cleaned):
                        rows.append(cleaned)
            if not rows:
                continue
            top = min(t.bbox[1] for t in sub)
            bottom = max(t.bbox[3] for t in sub)
            merged.append({
                "rows": rows, "fragments": fragment_count,
                "x0": g["x0"], "x1": g["x1"], "top": top, "bottom": bottom,
            })
    return merged


def looks_like_column_header(row):
    """True if a row plausibly IS a column header (contains a year, "Notes",
    or a currency/unit marker) rather than just happening to be all-text --
    a section label like "Operating income" or "Income", or a descriptive
    subtitle like "for the year ended 31 December 2023", is also all-text
    but is not a column header and must not be mistaken for one. A real
    column header is short and made of label/year/unit tokens, not prose."""
    if not row:
        return False
    text = " ".join(str(c) for c in row if c)
    if not text:
        return False
    words = text.split()
    if len(words) > 8:
        return False
    prose_words = {"for", "the", "ended", "as", "at", "on", "of", "to", "and", "in", "year"}
    if any(w.lower().strip(",.") in prose_words for w in words):
        return False
    if re.search(r"\b(19|20)\d{2}\b", text):
        return True
    if re.search(r"\bNotes?\b", text, re.I):
        return True
    if re.search(r"AED|USD|EGP|SAR|['\u2019]000", text):
        return True
    return False


def split_glued_year_unit_row(rows):
    """
    A native table row can span two visual PDF lines within one row band
    (e.g. "2018" / "AED'000" stacked), which the ordinary same-row text
    join fuses into one cell ("2018 AED'000") -- unlike a header
    reconstructed from above the table, which already keeps such lines
    separate. Detect that fused pattern in a would-be header row and split
    it back into two rows so it matches the report's own two-line layout.
    """
    if not rows:
        return rows
    pattern = re.compile(r"^\s*\(?((?:19|20)\d{2})\)?\s+\(?((?:AED|USD|EGP|SAR)\s*['\u2019]?\s*000\)?)\s*$", re.I)
    r = rows[0]
    matches = [(c, pattern.match(str(c))) for c in r if c is not None]
    if not any(m for _, m in matches):
        return rows
    years_row = [pattern.match(str(c)).group(1) if pattern.match(str(c)) else c for c in r]
    units_row = [pattern.match(str(c)).group(2) if pattern.match(str(c)) else None for c in r]
    return [years_row, units_row] + rows[1:]


def rows_to_header_and_body(rows):
    """
    Returns (header_rows, body) where header_rows is a list of one or more
    rows (kept SEPARATE, matching however many lines the report's own
    column header actually spans -- e.g. "2019 | 2018" on one line,
    "AED'000 | AED'000" on the line below -- rather than glued into one
    combined row). Each row is independently aligned to the right columns.
    """
    if not rows:
        return [], []
    rows = split_glued_year_unit_row(rows)
    idx = 0
    header_rows = []
    reference = None  # first header row: establishes which columns are bare years
    while idx < len(rows):
        r = rows[idx]
        if not (r and any(c for c in r) and looks_like_column_header(r)):
            break
        if reference is None:
            reference = list(r)
            header_rows.append(list(r))
        else:
            width = max(len(reference), len(r))
            if len(reference) < width:
                reference += [None] * (width - len(reference))
            non_none = [c for c in r if c is not None]
            is_bare_unit_row = bool(non_none) and all(
                re.fullmatch(r"(AED|USD|EGP|SAR)['\u2019]?\s*000", str(c).strip(), re.I) for c in non_none
            )
            aligned = [None] * width
            if is_bare_unit_row:
                # A row that's nothing but currency/unit markers (e.g. one
                # or two "AED'000" cells) isn't tied to whatever raw column
                # index it happened to land in -- such a row can be one
                # giant unruled cell internally, so the index is
                # unreliable. Line it up under whichever columns the first
                # header row already put a bare year in.
                units = iter(non_none)
                for c in range(width):
                    if reference[c] is not None and re.fullmatch(r"(19|20)\d{2}", str(reference[c]).strip()):
                        unit = next(units, None)
                        if unit is None:
                            break
                        aligned[c] = unit
            else:
                for c in range(min(width, len(r))):
                    aligned[c] = r[c]
            header_rows.append(aligned)
        idx += 1
    return header_rows, rows[idx:]


def reconstruct_column_header(page, table_top, x0, x1, max_gap=45):
    """
    The "Notes / 2023 / 2022 / AED'000 / AED'000" style column-header row
    sometimes sits entirely above a table's own ruled box (not one of its
    rows at all), so normal extraction never sees it. Look in the gap
    directly above the table for text that plausibly is that header, and
    rebuild it into 1-2 header-style rows ([None, col1, col2, ...]) if found.
    Tries a tight column-restricted search first (avoiding a neighboring
    column's header on a multi-statement page), widening only if that
    finds nothing.
    """
    try:
        words = page.extract_words()
    except Exception:
        return []

    def search(x_lo, x_hi):
        band = [
            w for w in words
            if table_top - max_gap <= w["top"] < table_top
            and x_lo <= w["x0"] <= x_hi
        ]
        if not band:
            return []
        lines = build_lines(band, x_gap_split=10_000)  # keep each full top-row together;
                                                         # we do our own column-splitting below
        lines.sort(key=lambda l: l["top"])
        found = []
        for line in lines:
            if not looks_like_column_header([line["text"]]):
                continue
            ws = sorted(line["words"], key=lambda w: w["x0"])
            clusters, cur = [], [ws[0]]
            for w in ws[1:]:
                if w["x0"] - cur[-1]["x1"] > 15:
                    clusters.append(cur)
                    cur = []
                cur.append(w)
            clusters.append(cur)
            cell_texts = [" ".join(w["text"] for w in c) for c in clusters]
            # Two adjacent unit markers (e.g. "AED'000" for this year right
            # next to "AED'000" for last year) can sit close enough that
            # the general x-gap clustering above merges them into one cell
            # ("AED'000 AED'000") instead of two -- split that back apart
            # so it lines up one-per-value-column like everything else.
            split_texts = []
            for t in cell_texts:
                parts = t.split()
                if len(parts) == 2 and parts[0] == parts[1] and re.fullmatch(
                    r"(AED|USD|EGP|SAR)['\u2019]?\s*000", parts[0], re.I
                ):
                    split_texts.extend(parts)
                else:
                    split_texts.append(t)
            cell_texts = split_texts
            # A widened search can occasionally reach far enough right to
            # pick up a NEIGHBORING table's own header on a page with
            # several tables side by side, which shows up as the same
            # value (e.g. "2024") repeating within one line -- a strong
            # signal of cross-table contamination. Cut the line at the
            # first repeat rather than keep the duplicated tail. A repeated
            # currency/unit marker is exempt -- "AED'000" legitimately
            # repeats once per value column and was just deliberately
            # split back into separate cells above.
            seen, cut = set(), None
            for i, t in enumerate(cell_texts):
                if re.fullmatch(r"(AED|USD|EGP|SAR)['\u2019]?\s*000", t.strip(), re.I):
                    continue
                key = t.strip().lower()
                if key in seen:
                    cut = i
                    break
                seen.add(key)
            if cut is not None:
                cell_texts = cell_texts[:cut]
            found.append([None] + cell_texts)
        return found

    tight = search(x0 - 40, x1 + 250)
    return tight if tight else search(max(0, x0 - 350), x1 + 300)


def reconstruct_unruled_rows(page, x0, x1, top, bottom, label_x_end):
    """
    Manually rebuild rows from raw word positions in a region with NO
    ruling lines at all around its data (some reports don't rule their
    statements at all, so pdfplumber's table detection only ever sees a
    stray header fragment, never the body). Words left of label_x_end on a
    line are joined into that row's label; words at/right of it are
    clustered into value columns by horizontal gap. This bypasses
    pdfplumber's own table-shape inference, which tends to fragment
    wrapped label text into spurious extra columns in cases like this.
    """
    words = [w for w in page.extract_words() if top <= w["top"] < bottom and x0 <= w["x0"] <= x1]
    if not words:
        return []
    lines = build_lines(words, x_gap_split=10_000)  # keep each full text line together
    raw_rows = []
    for line in sorted(lines, key=lambda l: l["top"]):
        label_words = sorted((w for w in line["words"] if w["x0"] < label_x_end), key=lambda w: w["x0"])
        value_words = sorted((w for w in line["words"] if w["x0"] >= label_x_end), key=lambda w: w["x0"])
        label_text = " ".join(w["text"] for w in label_words) or None
        label_x0 = label_words[0]["x0"] if label_words else None
        cols, cur = [], []
        for w in value_words:
            if cur and w["x0"] - cur[-1]["x1"] > 15:
                cols.append(cur)
                cur = []
            cur.append(w)
        if cur:
            cols.append(cur)
        # Keep each column's own x0 alongside its text for now -- needed
        # below to detect a Notes column that only SOME rows populate.
        col_cells = [{"text": " ".join(w["text"] for w in c), "x0": c[0]["x0"]} for c in cols]
        if label_text or col_cells:
            raw_rows.append({
                "label": label_text, "cols": col_cells,
                "top": line["top"], "label_x0": label_x0,
            })

    # A label that wraps onto a second line (with the values only printed
    # alongside its FIRST line) otherwise shows up as a spurious extra row
    # containing just the wrapped-over words. Fold such a row back into the
    # previous one when it's a tight continuation: no values of its own,
    # sitting close underneath, left-aligned with the row above.
    rows = []
    for entry in raw_rows:
        prev = rows[-1] if rows else None
        is_label_only = entry["label"] and not entry["cols"]
        if (
            is_label_only and prev is not None and prev["cols"]
            and entry["top"] - prev["top"] <= 12
            and entry["label_x0"] is not None and prev["label_x0"] is not None
            and abs(entry["label_x0"] - prev["label_x0"]) <= 10
        ):
            prev["label"] = f"{prev['label']} {entry['label']}"
            prev["top"] = entry["top"]
            continue
        rows.append(dict(entry))

    # Some rows carry a short note-reference number as their first value
    # column (e.g. "20  168,808  79,989"), but rows with no note of their
    # own just skip straight to the amount (e.g. "226,064  200,493") --
    # without a placeholder, that leaves the two kinds of rows with a
    # different number of columns, so a shared header ends up misaligned
    # against about half the rows. Detect the notes column's x-position
    # from rows that do have one, then insert a blank for rows that don't.
    note_x0s = [
        r["cols"][0]["x0"] for r in rows
        if r["cols"] and re.fullmatch(r"\d{1,3}(,\s*\d{1,3})*", r["cols"][0]["text"].strip())
    ]
    if note_x0s and len(note_x0s) < len(rows):
        notes_x0 = min(note_x0s)
        for r in rows:
            if not r["cols"]:
                continue
            first_x0 = r["cols"][0]["x0"]
            if first_x0 > notes_x0 + 30:
                r["cols"] = [{"text": None, "x0": notes_x0}] + r["cols"]

    return [[r["label"]] + [c["text"] for c in r["cols"]] for r in rows]


def normalize_missing_notes_column(rows):
    """
    Some rows carry a short note-reference number right after the label
    (e.g. "Investment income | 20 | 168,808 | 79,989"), but a row with no
    note of its own just skips straight to the amount (e.g. "Trading
    commission fees | 226,064 | 200,493"). By the time rows reach here they
    may already be padded to a uniform length elsewhere (a trailing blank
    filling out the row), which hides that mismatch from a simple length
    check -- so this looks at content instead: does position 1 hold a
    plausible note reference for SOME rows but a clearly-too-big value (or
    nothing) for others, with room (a trailing blank) to shift into.
    """
    if not rows:
        return rows

    def is_note_ref(cell):
        if cell is None:
            return False
        s = str(cell).strip()
        return bool(re.fullmatch(r"\d{1,2}(,\s*\d{1,2})*", s))

    has_notes_col = any(len(r) > 1 and is_note_ref(r[1]) for r in rows)
    if not has_notes_col:
        return rows

    def is_value_like(cell):
        return isinstance(cell, (int, float)) or (isinstance(cell, str) and cell.strip() == "-")

    out = []
    for r in rows:
        # Structural signal, not magnitude: both value slots are actually
        # filled with numbers but the row still ends in a padding blank --
        # that only happens when this row's notes slot was skipped and
        # everything after it shifted one column left. Works regardless of
        # how big or small the values themselves are.
        if (
            len(r) >= 3 and r[-1] is None and not is_note_ref(r[1])
            and is_value_like(r[1]) and is_value_like(r[2])
        ):
            r = [r[0], None] + list(r[1:-1])
        out.append(r)
    return out


def get_real_merged_tables(page):
    """All non-junk, de-fragmented tables on a page, top to bottom."""
    raw_tables = page.find_tables()
    real_tables = [t for t in raw_tables if not is_junk_table(t, page)]
    merged = group_and_merge_fragments(page, real_tables) if real_tables else []
    merged.sort(key=lambda m: m["top"])
    for g in merged:
        g["rows"] = normalize_missing_notes_column(g["rows"])

    # Fallback: a group that's just a tiny stub (e.g. only a column-header
    # row, with no ruled box at all around the actual data rows below it --
    # some reports don't rule their statements at all) gets its rows rebuilt
    # directly from word positions, cropped tightly to the stub's own
    # x-range. Because that's already confined to one column, a much wider
    # vertical search is safe here in a way it wouldn't be scanning the
    # whole (possibly multi-column) page at once.
    expanded = []
    for i, g in enumerate(merged):
        if len(g["rows"]) > 2:
            expanded.append(g)
            continue
        # Find the next group's top edge in roughly the SAME column (x-range
        # overlaps), not just the next item in top-sorted order -- two stub
        # tables can sit side by side at nearly the same vertical position
        # on a multi-column page, which would otherwise collapse the crop
        # height to near zero.
        same_column_below = [
            g2 for g2 in merged
            if g2 is not g and g2["top"] > g["top"] + 5
            and not (g2["x1"] < g["x0"] - 20 or g2["x0"] > g["x1"] + 20)
        ]
        bottom_limit = min((g2["top"] for g2 in same_column_below), default=page.height - 20) - 3
        if bottom_limit - g["top"] < 20:
            expanded.append(g)
            continue

        # The stub's own (ruled) header row tells us where the label column
        # ends and the value block begins -- reuse that boundary so wrapped
        # label text doesn't get mis-split into extra columns.
        orig_tables = [
            t for t in real_tables
            if abs(t.bbox[0] - g["x0"]) <= BBOX_TOLERANCE and abs(t.bbox[2] - g["x1"]) <= BBOX_TOLERANCE
        ]
        label_x_end = None
        if orig_tables:
            cells = orig_tables[0].rows[0].cells
            if cells and cells[0]:
                label_x_end = cells[0][2]
        if label_x_end is None:
            label_x_end = g["x0"] + 0.55 * (g["x1"] - g["x0"])

        new_rows = reconstruct_unruled_rows(page, g["x0"] - 5, g["x1"] + 5, g["top"], bottom_limit, label_x_end)
        if len(new_rows) > len(g["rows"]):
            expanded.append({
                "rows": normalize_missing_notes_column(new_rows), "fragments": g["fragments"],
                "x0": g["x0"], "x1": g["x1"], "top": g["top"], "bottom": bottom_limit,
            })
        else:
            expanded.append(g)

    expanded.sort(key=lambda m: m["top"])
    return expanded


# ------------------------------------------------------- NEW: target search ---

def normalize(text):
    return re.sub(r"\s+", " ", text or "").strip()


LINE_X_GAP_SPLIT = 45  # points; a horizontal gap this big within a "line" means it's
                        # actually two separate columns that happen to align vertically
                        # (common on multi-column landscape note pages) -- split them
                        # apart rather than concatenating unrelated columns' text.


def build_lines(words, x_gap_split=LINE_X_GAP_SPLIT):
    """Group words on a page into text lines by vertical position, then split
    apart any line that has a big horizontal gap in it (words from different,
    unrelated columns that merely happen to share a similar vertical position)."""
    raw_lines = []
    for w in sorted(words, key=lambda w: (round(w["top"]), w["x0"])):
        placed = False
        for line in raw_lines:
            if abs(line["top"] - w["top"]) <= LINE_TOLERANCE:
                line["words"].append(w)
                placed = True
                break
        if not placed:
            raw_lines.append({"top": w["top"], "words": [w]})

    lines = []
    for raw in raw_lines:
        ws = sorted(raw["words"], key=lambda w: w["x0"])
        segment = [ws[0]]
        for w in ws[1:]:
            if w["x0"] - segment[-1]["x1"] > x_gap_split:
                lines.append({"top": raw["top"], "words": segment})
                segment = []
            segment.append(w)
        lines.append({"top": raw["top"], "words": segment})

    for line in lines:
        line["words"].sort(key=lambda w: w["x0"])
        line["text"] = " ".join(w["text"] for w in line["words"])
        line["x0"] = min(w["x0"] for w in line["words"])
        line["x1"] = max(w["x1"] for w in line["words"])
    lines.sort(key=lambda l: l["top"])
    return lines


def find_heading_hits(lines, patterns):
    """Lines on a page matching any of the target's heading patterns."""
    hits = []
    for line in lines:
        norm = normalize(line["text"]).upper()
        for pat in patterns:
            if pat.search(norm):
                hits.append(line)
                break
    return hits


def row_label(row):
    for c in row:
        if isinstance(c, str) and c.strip():
            return c.strip()
    return None


def page_scan_order(total_pages, hint_page):
    """
    Page indices (0-based) to check, in order. If a hint is given, start
    there and spiral outward so likely matches surface fast; either way
    every page in the document is still covered as a fallback.
    """
    if not hint_page:
        return list(range(total_pages))
    center = max(0, min(total_pages - 1, hint_page - 1))
    order = [center]
    lo, hi = center - 1, center + 1
    while lo >= 0 or hi < total_pages:
        if lo >= 0:
            order.append(lo)
            lo -= 1
        if hi < total_pages:
            order.append(hi)
            hi += 1
    return order


def label_similarity(a, b):
    return difflib.SequenceMatcher(None, normalize(a).upper(), normalize(b).upper()).ratio()


def table_match_score(body_rows, reference_labels, threshold=FUZZY_ROW_MATCH_THRESHOLD):
    """
    Fraction of reference_labels that fuzzy-match SOME row label in this
    table. 1.0 = every known row label was found; 0.0 = none were.
    No reference labels configured -> None (can't score content).
    """
    if not reference_labels:
        return None
    row_labels = [row_label(r) for r in body_rows if row_label(r)]
    if not row_labels:
        return 0.0
    matched = 0
    for ref in reference_labels:
        best = max((label_similarity(ref, rl) for rl in row_labels), default=0.0)
        if best >= threshold:
            matched += 1
    return matched / len(reference_labels)


def truncate_at_row(body, target):
    """
    Cut a table's body right after the row matching target's cutoff (exact-ish
    row_patterns first, then a fuzzy fallback against cutoff_reference_labels).
    Returns (possibly-shortened body, was_confirmed).
    """
    for i, row in enumerate(body):
        label = row_label(row)
        if not label:
            continue
        norm_label = normalize(label).upper()
        if any(pat.match(norm_label) for pat in target.get("row_patterns", [])):
            return body[: i + 1], True

    best_i, best_ratio = None, 0.0
    for i, row in enumerate(body):
        label = row_label(row)
        if not label:
            continue
        for ref in target.get("cutoff_reference_labels", []):
            ratio = label_similarity(label, ref)
            if ratio > best_ratio:
                best_ratio, best_i = ratio, i
    if best_i is not None and best_ratio >= 0.8:
        return body[: best_i + 1], True

    return body, False


def find_best_matching_table(pdf, page_indices, target):
    """
    Two-pass search for the table a target describes:

    Pass 1 (fast): pages whose text matches one of target['heading_patterns'].
    The table found there is only accepted if it ALSO clears
    MIN_CONTENT_SCORE against target['reference_row_labels'] -- so a heading
    match that turns out to be unrelated text doesn't win by itself.

    Pass 2 (fallback, whole document): if pass 1 didn't land on anything
    confident, every real table in the entire PDF is scored against
    reference_row_labels by fuzzy row-label overlap, and the best-scoring
    one above MIN_CONTENT_SCORE is used. This is what makes it resilient to
    a report naming the section something heading_patterns didn't anticipate.

    Returns (header, body, page_num_1_based, score, actual_heading) or
    (None, None, None, 0.0, None). actual_heading is a best-effort guess
    at what the report itself calls this section (see find_actual_heading),
    for surfacing when it differs from target['name'].
    """
    reference_labels = target.get("reference_row_labels", [])
    best = {"score": -1.0, "header": None, "body": None, "page": None, "top": None, "x0": None, "x1": None}

    checked = set()
    for idx in page_indices:
        if idx in checked or idx >= len(pdf.pages):
            continue
        checked.add(idx)
        page = pdf.pages[idx]
        lines = build_lines(page.extract_words())
        hits = find_heading_hits(lines, target["heading_patterns"])
        if not hits:
            continue
        heading_top = hits[0]["top"]

        for lookahead in range(0, 3):  # this page, then up to 2 more
            pidx = idx + lookahead
            if pidx >= len(pdf.pages):
                break
            merged = get_real_merged_tables(pdf.pages[pidx])
            if lookahead == 0:
                below = [m for m in merged if m["top"] >= heading_top - 5]
                merged = below if below else merged
            for mt in merged:
                header, body = rows_to_header_and_body(mt["rows"])
                score = table_match_score(body, reference_labels)
                score = 1.0 if score is None else score  # no reference labels -> trust the heading
                if score > best["score"]:
                    best.update(score=score, header=header, body=body, page=pidx + 1,
                                top=mt["top"], x0=mt["x0"], x1=mt["x1"])
            if merged:
                break  # only look further ahead if this page had no tables at all

        if best["score"] >= 0.6:
            break

    def finalize():
        actual_heading = None
        header, body = best["header"], best["body"]
        if best["page"] is not None:
            page = pdf.pages[best["page"] - 1]
            actual_heading = find_actual_heading(page, best["top"], x0=best["x0"], x1=best["x1"])
            if not header:
                # No column-header row was captured as part of the table
                # itself (some reports draw it entirely above the ruled
                # box) -- try to reconstruct it from the gap above.
                recon = reconstruct_column_header(page, best["top"], best["x0"], best["x1"])
                if recon:
                    # Different header lines can carry different numbers of
                    # cells (e.g. "Notes / 2023 / 2022" has 3, but a
                    # standalone "2023 / 2022" line only has 2 because it
                    # has no note-column text of its own) -- right-align
                    # each line's cells against the body's actual column
                    # count so they land under the correct value columns
                    # instead of drifting one column left. EXCEPT a
                    # standalone "Notes"/"Note" line -- that's specifically
                    # the label for the (left-most) notes column, right
                    # after the row-label column, not a value column, so it
                    # gets left-aligned to that position instead. Each line
                    # is kept as its OWN row (matching however many lines
                    # the report's own header actually spans), not merged.
                    ncols = max((len(r) for r in body), default=1)
                    header = []
                    for r in recon:
                        cells = [c for c in r if c is not None]
                        if len(cells) == 1 and re.fullmatch(r"Notes?", str(cells[0]).strip(), re.I):
                            row = [None, cells[0]] + [None] * (ncols - 2)
                        else:
                            pad = ncols - len(r)
                            row = ([None] * pad + r) if pad > 0 else r
                        header.append(row)
        return header, body, best["page"], best["score"], actual_heading

    if best["score"] >= 0.6:
        return finalize()

    if best["score"] >= MIN_CONTENT_SCORE:
        return finalize()

    # Pass 2: full-document content scan
    if reference_labels:
        print(f"    heading search inconclusive for '{target['name']}' "
              f"(best so far: {max(best['score'], 0):.0%}) -- scanning full document by content match...")
        for i, page in enumerate(pdf.pages):
            if (i + 1) % 25 == 0:
                print(f"      ...scanned {i + 1}/{len(pdf.pages)} pages")
            for mt in get_real_merged_tables(page):
                header, body = rows_to_header_and_body(mt["rows"])
                score = table_match_score(body, reference_labels)
                if score is not None and score > best["score"]:
                    best.update(score=score, header=header, body=body, page=i + 1,
                                top=mt["top"], x0=mt["x0"], x1=mt["x1"])

    if best["score"] >= MIN_CONTENT_SCORE:
        return finalize()
    return None, None, None, 0.0, None


# ---------------------------------------------------------------- writing ---

def safe_sheet_name(name: str) -> str:
    name = re.sub(r'[\[\]:*?/\\]', "_", name)
    return name[:31]


def write_title(ws, row, text, ncols):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=max(ncols, 4))
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = TITLE_FONT
    cell.fill = TITLE_FILL
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[row].height = 26
    return row + 1


def write_subtitle(ws, row, text):
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = SUBTITLE_FONT
    return row + 1


def write_section_label(ws, row, text, ncols):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=max(ncols, 1))
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = SECTION_FONT
    cell.fill = SECTION_FILL
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[row].height = 18
    return row + 1


def write_table(ws, row, header_rows, body):
    if header_rows and not isinstance(header_rows[0], (list, tuple)):
        header_rows = [header_rows]  # allow a single flat row for backward compatibility
    ncols = max([len(h) for h in header_rows] + [len(r) for r in body], default=1)
    ncols = max(ncols, 1)
    for header in header_rows:
        for c in range(ncols):
            val = header[c] if c < len(header) else None
            cell = ws.cell(row=row, column=c + 1, value=val)
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
            cell.border = BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        row += 1
    for i, data_row in enumerate(body):
        fill = ROW_FILL_ALT if i % 2 == 1 else None
        for c in range(ncols):
            val = data_row[c] if c < len(data_row) else None
            cell = ws.cell(row=row, column=c + 1, value=val)
            if val is None:
                continue
            cell.font = BODY_FONT
            cell.border = BORDER
            if fill:
                cell.fill = fill
            if isinstance(val, (int, float)):
                cell.number_format = '#,##0.00;(#,##0.00);"-"' if isinstance(val, float) and not val.is_integer() else '#,##0;(#,##0);"-"'
                cell.alignment = Alignment(horizontal="right")
            else:
                cell.alignment = Alignment(horizontal="left", wrap_text=False)
        row += 1
    return row, ncols


def autofit_columns(ws, max_col, cap=60):
    widths = {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None or cell.column > max_col:
                continue
            length = len(str(cell.value))
            widths[cell.column] = max(widths.get(cell.column, 8), min(length + 2, cap))
    for col, width in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = width


# ------------------------------------------------------------- per-file ---

def normalize_pdf_for_reading(pdf_path: Path) -> Path:
    """
    Re-save the PDF through pypdf into a temp file before handing it to
    pdfplumber. Some PDFs (seemingly ones with unusually complex page
    content streams) cause pdfplumber/pdfminer to hang indefinitely on
    perfectly valid, readable files -- re-saving through pypdf normalizes
    the structure and reliably avoids this, with no observed loss of
    content. Falls back to the original path if the re-save itself fails.
    """
    try:
        reader = PdfReader(str(pdf_path))
        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)
        tmp_path = Path(tempfile.gettempdir()) / f"normalized_{pdf_path.stem}.pdf"
        with open(tmp_path, "wb") as f:
            writer.write(f)
        return tmp_path
    except Exception as e:
        print(f"    (could not pre-normalize {pdf_path.name}, using original: {e})")
        return pdf_path


def convert_pdf(pdf_path: Path, wb: Workbook, hints: dict):
    print(f"\nProcessing: {pdf_path.name}")
    ws = wb.create_sheet(title=safe_sheet_name(pdf_path.stem))

    row = 1
    row = write_title(ws, row, pdf_path.stem.replace("-", " ").replace("_", " ").title(), 6)
    row = write_subtitle(ws, row, f"Extracted from PDF \u00b7 {pdf_path.name}")
    row += 1
    max_col_seen = 1

    read_path = normalize_pdf_for_reading(pdf_path) if HAVE_PYPDF else pdf_path
    with pdfplumber.open(read_path) as pdf:
        total_pages = len(pdf.pages)
        page_labels = load_page_labels(pdf_path)
        for target in TARGETS:
            hint = hints.get(target["hint_key"])
            order = page_scan_order(total_pages, hint)

            header, body, page_num, score, actual_heading = find_best_matching_table(pdf, order, target)

            if body is None:
                row = write_section_label(ws, row, f"{target['name']} \u2014 NOT FOUND", 6)
                cell = ws.cell(row=row, column=1, value="No heading or content match cleared the confidence threshold. See console output.")
                cell.font = NOT_FOUND_FONT
                cell.fill = NOT_FOUND_FILL
                row += 2
                print(f"  '{target['name']}' NOT FOUND in {pdf_path.name}")
                continue

            confidence_note = ""
            if target["mode"] == "table_until_row":
                body, cutoff_confirmed = truncate_at_row(body, target)
                if not cutoff_confirmed:
                    confidence_note = "  [cutoff row not confirmed \u2014 showing full matched table, please verify]"

            # Note what the report itself calls this section when it's not
            # (close to) verbatim what we searched for -- so a match found by
            # content/wording similarity is still traceable back to the source.
            called_note = ""
            if actual_heading and label_similarity(actual_heading, target["name"]) < 0.6:
                called_note = f'  \u2014 called "{actual_heading}" in this report'

            page_str = display_page(page_labels, page_num)

            label = f"{target['name']}  (page {page_str}, content match {score:.0%}){called_note}{confidence_note}"
            row = write_section_label(ws, row, label, 6)
            row, ncols = write_table(ws, row, header, body)
            max_col_seen = max(max_col_seen, ncols)
            row += 1
            called_console = f' (called "{actual_heading}" in the report)' if called_note else ""
            print(f"  found '{target['name']}' on page {page_str} (score {score:.0%}){called_console}: {len(body)} rows")

    ws.freeze_panes = "A4"
    autofit_columns(ws, max_col_seen)


def main():
    parser = argparse.ArgumentParser(description="Extract the P&L 'profit for the year' row and the G&A expenses table from annual report PDFs.")
    parser.add_argument("pdfs", nargs="+", help="One or more PDF files and/or folders. A folder is expanded to every .pdf file inside it.")
    parser.add_argument("--recursive", action="store_true", help="When a folder is given, also include PDFs in its subfolders (default: only the top level).")
    parser.add_argument("--pl-hint", type=int, default=None, help="Approximate page number of the Statement of Profit or Loss (optional, speeds up search)")
    parser.add_argument("--admin-hint", type=int, default=None, help="Approximate page number of General and administrative expenses (optional, speeds up search)")
    parser.add_argument("--out", type=str, default=None, help="Output .xlsx path (default: two_tables_extracted.xlsx next to the first PDF)")
    args = parser.parse_args()

    hints = {"pl_hint": args.pl_hint, "admin_hint": args.admin_hint}

    # Expand any folder arguments into the .pdf files they contain, keeping
    # the original order and de-duplicating in case a file is named both
    # directly and reached again via a folder.
    pdf_paths = []
    seen = set()
    for pdf_arg in args.pdfs:
        p = Path(pdf_arg)
        if p.is_dir():
            pattern = "**/*.pdf" if args.recursive else "*.pdf"
            found = sorted(p.glob(pattern))
            if not found:
                print(f"  SKIP: no .pdf files found in folder {p}")
            for f in found:
                if f.resolve() not in seen:
                    seen.add(f.resolve())
                    pdf_paths.append(f)
        elif p.exists():
            if p.resolve() not in seen:
                seen.add(p.resolve())
                pdf_paths.append(p)
        else:
            print(f"  SKIP: {p} not found (not a file or folder)")

    if not pdf_paths:
        print("No PDF files found (checked the given files/folders). Nothing to process.")
        sys.exit(1)

    print(f"Found {len(pdf_paths)} PDF file(s) to process:")
    for p in pdf_paths:
        print(f"  - {p}")

    wb = Workbook()
    wb.remove(wb.active)

    for pdf_path in pdf_paths:
        if not pdf_path.exists():
            print(f"  SKIP: {pdf_path} not found")
            continue
        try:
            convert_pdf(pdf_path, wb, hints)
        except Exception as e:
            print(f"  ERROR processing {pdf_path.name}: {e}")

    if len(wb.sheetnames) == 0:
        print("No PDFs were processed. Nothing to save.")
        sys.exit(1)

    out_path = Path(args.out) if args.out else pdf_paths[0].parent / "two_tables_extracted.xlsx"
    wb.save(out_path)
    print(f"\nDone. Workbook saved to: {out_path}")


if __name__ == "__main__":
    main()
