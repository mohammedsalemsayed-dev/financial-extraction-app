"""
telecom_extract.py -- pull TWO tables out of du and Etisalat/e& annual-report
PDFs, the same two the DFM extractor pulls, and write ONE WORKBOOK PER COMPANY:

    1. Consolidated statement of profit or loss  (du calls it "statement of
       comprehensive income")  -- from the top of the statement (Revenue ...)
       down through and INCLUDING the "Profit for the year" row. Everything
       after that (other comprehensive income, "profit attributable to ...",
       earnings per share) is dropped -- mirrors the DFM script.
    2. The operating-expenses-by-nature note the P&L's "Operating expenses" /
       "General and administrative expenses" line cross-references (Staff
       costs, Depreciation, Amortisation, Interconnect costs, Marketing, ...)
       down to and including its total.

Two companies, two profiles, two output files:
    - files named "du annual <year>.pdf"       -> profile "du"        -> du_two_tables.xlsx
    - everything else (etisalat*/eand*/en-20*  -> profile "etisalat"  -> etisalat_eand_two_tables.xlsx
      /integrated-report*)
The profile is chosen automatically from the file name; override with --company.

CORRECTNESS IS CHECKED ARITHMETICALLY (the DFM handoff's #1 recommendation):
    - note:  sum(line items) must equal the printed total, per value column.
    - P&L:   a running sum with subtotal detection -- every subtotal (Gross
             profit, Operating profit, ...) and the final "Profit for the
             year" must equal the signed sum of the primitive lines above it.
Among all candidate tables/regions that match the vocabulary, the one whose
arithmetic RECONCILES is preferred -- a wrong table (segmental note, changes
in equity, comprehensive income) almost never reconciles. Each sheet carries
a PASS / FAIL banner with the numbers, and the console prints a summary.

Deterministic: pdfplumber + pypdf + openpyxl + re. No LLM, no network.

-----------------------------------------------------------------------
SETUP:  pip install pdfplumber openpyxl pypdf
USAGE:  python telecom_extract.py .                      # every PDF in this folder
        python telecom_extract.py "du annual 2024.pdf"   # one file
        python telecom_extract.py . --outdir out         # choose where the .xlsx go
        python telecom_extract.py x.pdf --company du      # force a profile
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
# Two company PROFILES. Each is a pair of "targets" (P&L, then the expense
# note). A target carries: the heading regexes that name it, the reference
# vocabulary used to score candidate tables, "anti" labels that positively
# identify a DIFFERENT table sharing that vocabulary (segmental note, changes
# in equity, ...), and -- for the note -- "anchor" labels of which at least
# two must be present. `reconcile` says which arithmetic check applies.
#
# The profile is chosen from the file name (see profile_for_file); each
# profile writes its own workbook.

MIN_CONTENT_SCORE = 0.20   # min fraction of a target's reference vocabulary a
                            # candidate table must fuzzy-match to be considered
                            # at all (the vocab spans two companies x ~15 years,
                            # so any one report only ever matches a fraction).
FUZZY_ROW_MATCH_THRESHOLD = 0.72   # how close two row labels must be to count as "the same row"
RECON_TOL = 2              # AED'000: rounding slack allowed when checking that a
                            # reconstructed column of figures sums to its total.

# Row that ENDS the P&L (tried in priority order; the LAST match of the
# first pattern that matches anything wins -- so "Profit for the year from
# continuing operations" followed later by a bare "Profit for the year"
# total cuts at the bare total, not the interim subtotal).
_PL_ROW_PATTERNS = [
    re.compile(r"^(NET\s+)?PROFIT\s+FOR\s+THE\s+(FINANCIAL\s+)?(YEAR|PERIOD)\s*$", re.I),
    re.compile(r"^(PROFIT|LOSS)\s+FOR\s+THE\s+(YEAR|PERIOD)\s+FROM\s+CONTINUING\s+OPERATIONS\s*$", re.I),
    re.compile(r"^(NET\s+)?(PROFIT|LOSS)\s+FOR\s+THE\s+(FINANCIAL\s+)?(YEAR|PERIOD)\b", re.I),
]
_PL_CUTOFF_LABELS = [
    "Profit for the year", "Profit for the year from continuing operations",
    "Loss for the year",
]

# Labels that never appear on a real consolidated P&L but DO appear on the
# look-alikes the content score can otherwise pick (segmental information;
# statement of changes in equity; standalone other-comprehensive-income).
_PL_ANTI = [
    "Inter-segment revenue", "Segment result", "Segment results",
    "Reportable segment", "Segment assets", "Segment liabilities",
    "Balance at 1 January", "Transfer to reserves",
    "Share capital", "Share premium", "Retained earnings",
    "Total comprehensive income for the year",
    "Remeasurement of defined benefit obligations",
    "Exchange differences on translation of foreign operations",
    "Gain on net investment hedges", "Actuarial gain on defined benefit obligations",
    # Cash-flow statement -- also a running-total statement, so it can slip
    # past the arithmetic check; these lines only ever appear there.
    "Cash generated from operations", "Net cash generated from operating activities",
    "Net cash used in investing activities", "Net cash used in financing activities",
    "Purchase of property, plant and equipment", "Payment of federal royalty fee",
    "Cash and cash equivalents at 1 January",
    "Proceeds from disposal of property, plant and equipment",
    "Adjustments for non-cash items",
]

# Labels that identify a wrong sub-table inside the expense / royalty note.
_NOTE_ANTI = [
    "UAE Net Regulated Revenue", "ICT Fund Contribution",
    "Total regulated revenue", "Total regulated profit",
    "Royalty on regulated revenue", "Royalty on regulated profit",
    "Broadcasting revenue for the year", "Opening balance", "Closing balance",
    "Transfer to statutory reserve",
]

DU_PL = {
    "name": "Consolidated statement of profit or loss (through 'Profit for the year')",
    "mode": "table_until_row",
    "reconcile": "pl",
    "hint_key": "pl_hint",
    "heading_patterns": [
        re.compile(r"CONSOLIDATED\s+STATEMENT\s+OF\s+PROFIT\s+(OR|AND)\s+LOSS", re.I),
        re.compile(r"CONSOLIDATED\s+INCOME\s+STATEMENT", re.I),
        re.compile(r"\bINCOME\s+STATEMENT\b", re.I),
        # du's single combined statement -- there is no separate P&L. Won't
        # match Etisalat's "statement of OTHER comprehensive income" (the
        # word "other" breaks the phrase).
        re.compile(r"STATEMENT\s+OF\s+COMPREHENSIVE\s+INCOME", re.I),
        re.compile(r"STATEMENT\s+OF\s+PROFIT\s+OR\s+LOSS\s+AND\s+OTHER\s+COMPREHENSIVE\s+INCOME", re.I),
    ],
    "row_patterns": _PL_ROW_PATTERNS,
    "cutoff_reference_labels": _PL_CUTOFF_LABELS,
    "reference_row_labels": [
        "Revenue", "Cost of sales", "Gross profit",
        "General and administrative expenses", "Operating expenses",
        "Expected credit losses", "Expected credit losses (net of recoveries)",
        "Depreciation and amortisation", "Other income",
        "Other (expense) / income", "Finance income", "Finance expense",
        "Finance costs", "Finance income and costs",
        "Impairment of goodwill",
        "Share of loss of equity accounted investment",
        "Share of loss of associate and joint venture",
        "Share of (loss)/profit of investments accounted for using equity method",
        "Profit before Royalty", "Profit before royalty on regulated profit",
        "Profit before federal royalty on profit and corporate income tax",
        "Change in estimate for prior years' Royalty",
        "Royalty", "Federal royalty on regulated revenue",
        "Federal royalty on regulated profit", "Federal royalty on profit",
        "Corporate income tax", "Profit for the year",
        # du 2025 recut its P&L: revenue by product, per-section subtotals,
        # and the statement wraps down the left half then the right half.
        "Mobile", "Fixed", "Wholesale", "ICT and associated telecom services",
        "Total revenue", "Interconnect cost", "Commission cost",
        "Devices and other direct services cost", "Total direct costs",
        "Network and other maintenance expense", "Marketing expense",
        "Staff expense", "Administrative expense", "Other operating expense",
        "Other operating income", "Depreciation and amortization",
        "Operating profit before depreciation and amortization", "Operating profit",
        "Share of loss on equity accounted investments", "Interest income",
        "Interest expense", "Profit before financing, federal royalty and income tax",
        "Profit before federal royalty and income tax", "Federal royalty",
        "Income tax expense", "Net Profit for the year",
        "Telecommunication license and related fees",
    ],
    "anti_reference_labels": _PL_ANTI,
}

DU_NOTE = {
    "name": "Operating expenses / General and administrative expenses note",
    "mode": "full_table",
    "reconcile": "note",
    "hint_key": "admin_hint",
    "heading_patterns": [
        re.compile(r"GENERAL\s+AND\s+ADMINISTRATIVE\s+EXPENSES", re.I),
        re.compile(r"GENERAL\s*&\s*ADMINISTRATIVE\s+EXPENSES", re.I),
        re.compile(r"\bOPERATING\s+EXPENSES\b", re.I),
    ],
    # >= 2 of these must be present for a candidate to be the opex note.
    "anchor_labels": [
        "Payroll and employee related expenses", "Staff costs",
        "Depreciation and amortisation expenses",
        "Depreciation and impairment on property, plant and equipment",
        "Interconnect costs", "Network operation and maintenance",
        "Outsourcing and contracting", "Telecommunication licence and related fees",
    ],
    "reference_row_labels": [
        "Payroll and employee related expenses", "Staff costs",
        "Outsourcing and contracting", "Consulting", "Consultancy costs",
        "Telecommunications licence and related fees",
        "Telecommunication licence and related fees",
        "Sales and marketing expenses", "Marketing",
        "Depreciation and amortisation expenses",
        "Depreciation and impairment on property, plant and equipment",
        "Depreciation on right-of-use assets",
        "Amortisation and impairment on intangible assets",
        "Network operation and maintenance", "Interconnect costs",
        "Product costs", "Cost of devices and direct services",
        "Commission", "Rent and utilities", "Provision for doubtful debts",
        "Provision for receivables",
        "Impairment of property, plant and equipment", "Miscellaneous",
        "Miscellaneous expenses", "Other expenses", "Others",
    ],
    "anti_reference_labels": _NOTE_ANTI + [
        "Presentation of expenses by function", "Presentation of expenses by nature",
        "Dedicated leased lines",
    ],
    "require_reconcile": True,
}

ETI_PL = {
    "name": "Consolidated statement of profit or loss (through 'Profit for the year')",
    "mode": "table_until_row",
    "reconcile": "pl",
    "hint_key": "pl_hint",
    "heading_patterns": [
        re.compile(r"CONSOLIDATED\s+STATEMENT\s+OF\s+PROFIT\s+(OR|AND)\s+LOSS", re.I),
        re.compile(r"STATEMENT\s+OF\s+PROFIT\s+OR\s+LOSS\s+AND\s+OTHER\s+COMPREHENSIVE\s+INCOME", re.I),
        re.compile(r"CONSOLIDATED\s+INCOME\s+STATEMENT", re.I),
        # NB: Etisalat/e&'s P&L page often carries the WRONG running header
        # ("statement of financial position"), so heading match frequently
        # fails and the content + arithmetic path is what actually finds it.
    ],
    "row_patterns": _PL_ROW_PATTERNS,
    "cutoff_reference_labels": _PL_CUTOFF_LABELS,
    "reference_row_labels": [
        "Revenue", "Operating expenses",
        "Impairment loss on trade receivables and contract assets",
        "Impairment loss on other assets", "Impairment loss on other assets - net",
        "Share of results of associates and joint ventures",
        "Operating profit before federal royalty",
        "Profit before federal royalty and corporate tax",
        "Federal royalty", "Operating profit",
        "Finance and other income", "Finance and other costs",
        "Profit before tax", "Income tax expenses", "Corporate tax expenses",
        "Profit for the year from continuing operations",
        "Loss from discontinued operations", "Profit for the year",
    ],
    "anti_reference_labels": _PL_ANTI,
}

ETI_NOTE = {
    "name": "Operating expenses note (note 7 'a) Operating expenses')",
    "mode": "full_table",
    "reconcile": "note",
    "hint_key": "admin_hint",
    "heading_patterns": [
        re.compile(r"OPERATING\s+EXPENSES\s+AND\s+FEDERAL\s+ROYALTY", re.I),
        re.compile(r"\bOPERATING\s+EXPENSES\b", re.I),
    ],
    "anchor_labels": [
        "Direct cost of sales", "Staff costs", "Depreciation", "Amortisation",
        "Network and other related costs", "Regulatory expenses",
        "Marketing expenses",
    ],
    "reference_row_labels": [
        "Direct cost of sales", "Staff costs", "Depreciation", "Amortisation",
        "Network and other related costs", "Regulatory expenses",
        "Marketing expenses", "Consultancy costs", "Operating lease rentals",
        "IT costs", "Foreign exchange losses",
        "Foreign exchange (gains) / losses - net", "Other operating expenses",
        "Operating expenses (before federal royalty)",
    ],
    "anti_reference_labels": _NOTE_ANTI,
}

PROFILES = {
    "du": {
        "label": "du",
        "out": "du_two_tables.xlsx",
        "file_patterns": [re.compile(r"(^|[^a-z])du([^a-z]|$).*annual|du[ _-]?annual", re.I)],
        "targets": [DU_PL, DU_NOTE],
    },
    "etisalat": {
        "label": "Etisalat / e&",
        "out": "etisalat_eand_two_tables.xlsx",
        "file_patterns": [re.compile(r"etisalat|eand|(^|[^a-z])en-20\d\d|integrated[-_ ]?report", re.I)],
        "targets": [ETI_PL, ETI_NOTE],
    },
}


def profile_for_file(path):
    """(key, profile) for a PDF, from its file name. Falls back to du when
    the name contains a bare 'du', otherwise to the Etisalat/e& profile."""
    name = path.name
    for key, prof in PROFILES.items():
        if any(p.search(name) for p in prof["file_patterns"]):
            return key, prof
    if re.search(r"(^|[^a-z])du([^a-z]|$)", name, re.I):
        return "du", PROFILES["du"]
    return "etisalat", PROFILES["etisalat"]


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


# Characters that show up *inside* words/numbers in some of these reports
# instead of a plain space: non-breaking / figure / thin spaces, zero-width
# marks, and U+FFFD (older du note tables render every inter-word space as
# the replacement char). Collapsed to a normal space everywhere text is
# cleaned or compared, so label matching isn't defeated by them.
_WEIRD_SPACE_RE = re.compile(
    "[   -​  　﻿�]+"
)


def _despace(text):
    return _WEIRD_SPACE_RE.sub(" ", text)


def clean_cell(value):
    if value is None:
        return None
    text = _despace(str(value)).strip().replace("\n", " ")
    if text == "":
        return None
    stripped = text.replace(",", "").replace("$", "").replace("%", "").strip()
    # Negative values print as "(1,234)" -- but several of these reports
    # (Etisalat / e& especially) come out of text extraction with the
    # parentheses reversed, ")1,234(", due to bidi handling in the source
    # PDF. Treat either orientation as a negative.
    negative = (stripped.startswith("(") and stripped.endswith(")")) or \
               (stripped.startswith(")") and stripped.endswith("("))
    if negative:
        stripped = stripped[1:-1].strip()
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


MAX_LABEL_LOOKUP_DIST = 380     # pts to search left of a numbers-only table for its row labels.
                                 # Raised from the DFM script's 300: on the Etisalat/e& and older-du
                                 # landscape (1191pt-wide) spreads, a single statement's ruled number
                                 # box can sit ~350pt right of its own row-label column. The median-x0
                                 # consistency check further down still drops labels pulled in from an
                                 # unrelated neighbouring column, so widening here is safe.
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
                # A row label can WRAP over two or three physical lines inside
                # one ruled band (du's notes: "Depreciation and impairment on
                # property, plant and equipment" then "(Note 6)" a line below,
                # left-aligned). Group by physical line first -- otherwise a
                # left-aligned wrapped fragment sorts into the middle of the
                # first line by x0 and corrupts the gap analysis below.
                _lines = build_lines(label_words, x_gap_split=10_000)
                _lines.sort(key=lambda l: l["top"])
                if len(_lines) > 1:
                    # wrapped label: read top-to-bottom, left-to-right; pull a
                    # "(Note N)" continuation out into its own note cell
                    parts = []
                    for ln in _lines:
                        parts.append(" ".join(
                            w["text"] for w in sorted(ln["words"], key=lambda w: w["x0"])))
                    joined = " ".join(parts).strip()
                    m_note = re.search(r"\s*\(?notes?\s*(\d{1,2}(?:\.\d+)?)\)?\s*$",
                                       joined, re.I)
                    note_cell = None
                    if m_note:
                        note_cell = m_note.group(1)
                        joined = joined[:m_note.start()].strip()
                    # also a "(Note N)" embedded mid-label
                    joined = re.sub(r"\s*\(notes?\s*\d{1,2}(?:\.\d+)?\)\s*", " ", joined,
                                    flags=re.I).strip()
                    label_x0 = min(w["x0"] for w in label_words)
                    if note_cell is not None:
                        row = [joined, note_cell] + row
                    else:
                        row = [joined] + row
                else:
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
    # A row carrying two or more real figures (thousands/millions) is a data
    # row, never a column header -- even if a stray "Notes" word merged into
    # its label.
    big = sum(1 for c in row
              if isinstance(c, (int, float)) and not isinstance(c, bool) and abs(c) >= 1000)
    if big >= 2:
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


_TOK_NUMERIC_RE = re.compile(r"^[\(\)\[\]\-–—.,%\d\s]+$")
_TOK_UNIT_RE = re.compile(r"^(AED|USD|EGP|SAR)['’]?\s*000$", re.I)


def _tok_is_text(s):
    return bool(re.search(r"[A-Za-z]", s)) and not _TOK_NUMERIC_RE.match(s) and not _TOK_UNIT_RE.match(s.strip())


def reconstruct_unruled_rows(page, x0, x1, top, bottom, label_x_end, stop_at_text=False):
    """
    Manually rebuild rows from raw word positions in a region with NO
    ruling lines at all around its data. Words left of label_x_end on a line
    are the row label; words at/right of it are the figures, split into
    columns by a single global column model so every row's figures land in
    the SAME columns (per-line clustering drifts, which defeats the
    arithmetic check).

    stop_at_text: for landscape spreads where a SECOND statement is
    interleaved to the right -- each row's figure list stops at the first
    plainly-textual token after the figures (that's the neighbour's label);
    any textual tokens BEFORE the first figure are a wrapped continuation of
    this row's own label and are folded back into it.
    """
    words = [w for w in page.extract_words() if top <= w["top"] < bottom and x0 <= w["x0"] <= x1]
    if not words:
        return []
    lines = build_lines(words, x_gap_split=10_000)  # keep each full text line together

    # --- per line: split into a label + a list of value tokens -----------
    per_line = []
    for line in sorted(lines, key=lambda l: l["top"]):
        lw = sorted(line["words"], key=lambda w: w["x0"])
        label_words = [w for w in lw if w["x0"] < label_x_end]
        value_words = [w for w in lw if w["x0"] >= label_x_end]
        toks, cur = [], []
        for w in value_words:
            if cur and w["x0"] - cur[-1]["x1"] > 7:
                toks.append(cur)
                cur = []
            cur.append(w)
        if cur:
            toks.append(cur)
        tok_dicts = [{
            "text": " ".join(w["text"] for w in t),
            "cx": (t[0]["x0"] + t[-1]["x1"]) / 2,
        } for t in toks]
        label = " ".join(w["text"] for w in label_words) or None

        # Words before the first figure-like token are a wrapped
        # continuation of this row's own label (e.g. "Expected credit" /
        # "losses (net off recoveries)"); fold them in. With stop_at_text
        # (landscape spread) also cut at the first text token AFTER the
        # figures -- that's the neighbouring statement's label.
        if tok_dicts:
            lead, i = [], 0
            while i < len(tok_dicts) and _tok_is_text(tok_dicts[i]["text"]):
                lead.append(tok_dicts[i]["text"]); i += 1
            core = []
            while i < len(tok_dicts):
                if _tok_is_text(tok_dicts[i]["text"]):
                    if core and stop_at_text:
                        break
                    i += 1; continue
                core.append(tok_dicts[i]); i += 1
            if lead:
                label = ((label + " ") if label else "") + " ".join(lead)
            tok_dicts = core

        per_line.append({
            "label": label,
            "label_x0": label_words[0]["x0"] if label_words else None,
            "label_x1": label_words[-1]["x1"] if label_words else None,
            "top": line["top"],
            "toks": tok_dicts,
        })

    # --- one global value-column model, so every row's figures land in
    # the SAME columns (per-line gap clustering drifts row to row, which
    # then defeats the arithmetic check). Number of columns = the modal
    # token count among value-bearing lines; column centres = the K
    # widest-separated groups of all token centres.
    counts = Counter(len(pl["toks"]) for pl in per_line if pl["toks"])
    K = counts.most_common(1)[0][0] if counts else 0
    K = max(1, min(K, 4))
    all_cx = sorted(t["cx"] for pl in per_line for t in pl["toks"])
    centres = []
    if all_cx:
        order = sorted(range(1, len(all_cx)), key=lambda i: all_cx[i] - all_cx[i - 1], reverse=True)
        cuts = sorted(order[:K - 1]) + [len(all_cx)]
        start = 0
        for ci in cuts:
            grp = all_cx[start:ci]
            if grp:
                centres.append(sum(grp) / len(grp))
            start = ci
    if not centres:
        centres = [label_x_end + 40]

    def assign(toks):
        cells = [None] * len(centres)
        if len(toks) == len(centres):
            for i, t in enumerate(toks):
                cells[i] = t["text"]
            return cells
        # right-align a short row; nearest-centre for a long/odd one
        if len(toks) < len(centres):
            for off, t in enumerate(reversed(toks)):
                cells[len(centres) - 1 - off] = t["text"]
            return cells
        for t in toks:
            j = min(range(len(centres)), key=lambda k: abs(centres[k] - t["cx"]))
            cells[j] = (cells[j] + " " + t["text"]) if cells[j] else t["text"]
        return cells

    raw_rows = [{
        "label": pl["label"], "label_x0": pl["label_x0"], "label_x1": pl["label_x1"],
        "top": pl["top"], "cells": assign(pl["toks"]),
    } for pl in per_line if pl["label"] or pl["toks"]]

    # Fold a wrapped label continuation back into its row. A label like
    # "Depreciation and impairment on property, plant and / equipment
    # (Note 6)" can come back as up to three fragments -- the first bit of
    # label, the figures on their own, and the rest of the label -- so merge
    # any adjacent pair where one side has the words and the other has (or
    # neither has) the figures, as long as they sit tight together.
    _figtok = re.compile(r"^\(?[-–]?[\d,]+(\.\d+)?\)?$")

    def _hasfig(cells):
        for c in cells:
            if c is None:
                continue
            parts = str(c).strip().split()
            if parts and all(_figtok.match(p) for p in parts):
                return True
        return False

    def _all_year_figs(cells):
        """the row's figures are all bare years (a '2021 2020' column header)."""
        seen = False
        for c in cells:
            if c is None:
                continue
            for p in str(c).strip().split():
                d = re.sub(r"[^\d]", "", p)
                if not d:
                    continue
                if not (len(d) == 4 and 1990 <= int(d) <= 2099):
                    return False
                seen = True
        return seen

    def _only_small_ints(cells):
        """the row's only 'figures' are 1-2 digit ints -- a wrapped '(Note 6)'
        reference, not real values."""
        seen = False
        for c in cells:
            if c is None:
                continue
            for p in str(c).strip().split():
                d = re.sub(r"[^\d]", "", p)
                if not d:
                    continue
                if len(d) > 2:
                    return False
                seen = True
        return seen

    # a label that ends mid-phrase -- its next line is a continuation
    _HANGING = re.compile(r"\b(and|or|of|the|to|for|on|in|before|after|with|"
                          r"plant|lease|other|net|from)\s*$", re.I)

    rows = []
    for e in raw_rows:
        prev = rows[-1] if rows else None
        close = (prev is not None and 0 <= e["top"] - prev["top"] <= 15)
        left_al = (e["label_x0"] is not None and prev is not None and prev["label_x0"] is not None
                   and abs(e["label_x0"] - prev["label_x0"]) <= 16)
        right_al = (e["label_x1"] is not None and prev is not None and prev.get("label_x1") is not None
                    and abs(e["label_x1"] - prev["label_x1"]) <= 16)
        aligned = close and (left_al or right_al)
        if prev is not None and close:
            p_fig, e_fig = _hasfig(prev["cells"]), _hasfig(e["cells"])
            p_txt = bool(prev["label"] and re.search(r"[A-Za-z]", prev["label"]))
            e_txt = bool(e["label"] and re.search(r"[A-Za-z]", e["label"]))
            p_hanging = bool(prev["label"] and _HANGING.search(prev["label"].strip()))
            # a short parenthetical / lowercase-start continuation
            e_is_tail = bool(e["label"]) and (
                e["label"][:1].islower() or e["label"].lstrip().startswith("(")
                or len(e["label"].split()) <= 3)
            # e continues prev's phrase (lowercase / opens with a bracket) --
            # safe to fold even when e carries the figures
            e_cont = bool(e["label"]) and (
                e["label"][:1].islower() or e["label"].lstrip().startswith("("))
            merge = False
            if p_fig and not e_fig and e_txt and aligned and (p_hanging or e_is_tail):
                merge = True                          # figures then trailing label
            elif (p_fig and e_txt and aligned and (p_hanging or e_is_tail)
                  and _only_small_ints(e["cells"])):
                merge = True                          # trailing label + wrapped
                #                                       "(Note 6)" reference
            elif p_txt and not p_fig and not e_txt and e_fig and not _all_year_figs(e["cells"]):
                merge = True                          # a label-less figures line
                #                                       belongs to the label above
                #                                       (but not a bare "2021 2020"
                #                                       year-header row)
            elif (p_txt and not p_fig and e_txt and not e_fig
                  and (aligned or (close and e_cont))
                  and (p_hanging or e_cont)
                  and (len(e["label"].split()) <= 4
                       or len(prev["label"].split()) <= 2)):
                merge = True                          # label wrapped onto 2 lines
                #                                       ("Depreciation" / "and
                #                                       impairment on property ...")
            elif (p_txt and not p_fig and e_fig and e_txt
                  and (aligned or (close and e_cont))
                  and (p_hanging or e_cont) and len(e["label"].split()) <= 7):
                merge = True                          # label wraps; figures on the 2nd line
            if merge:
                lbl = " ".join(x for x in (prev["label"], e["label"]) if x)
                # drop a wrapped "(Note 6)" fragment that split into "(Note" + "6)"
                lbl = re.sub(r"\s*\(note\s*$", "", lbl, flags=re.I)
                lbl = re.sub(r"\s+\d{1,2}\)\s*$", "", lbl)
                prev["label"] = lbl or prev["label"]
                if e_fig and not _only_small_ints(e["cells"]):
                    prev["cells"] = e["cells"]
                prev["top"] = e["top"]
                continue
        rows.append(dict(e))

    # If the left-most value column is really a Notes-reference column
    # (every non-empty entry is a 1-2 digit integer), leave it -- callers
    # and the reconciliation both know to skip it.
    return [[r["label"]] + list(r["cells"]) for r in rows]


_PAGE_NUMTOKEN_RE = re.compile(r"^\(?[-–]?\d[\d,]*(\.\d+)?\)?$")
_NOTES_FOOTER_RE = re.compile(
    r"notes?\b.*(form|are)\s+an?\s+integral\s+part|^\s*the\s+accompanying\s+notes",
    re.I,
)


def reconstruct_page_statement(page, heading_top, heading_x0=None):
    """
    Rebuild an unruled statement/note that pdfplumber's table finder misses
    ENTIRELY (no ruling lines at all -> `find_tables()` returns nothing).
    Common in du's 2021+ portrait reports and the e& integrated reports.

    Region: just below `heading_top`, down to a "The notes on pages ... form
    an integral part" footer (or the end of the text block). On a landscape
    page carrying TWO statements side by side, the numeric tokens split into
    two x-clusters -- keep only the half that contains `heading_x0` (falling
    back to the denser half), so the neighbouring statement's columns don't
    bleed in. Far-left margin cruft (rotated page furniture) is trimmed by
    starting the label column at the 10th-percentile label x, not the min.
    """
    words = [w for w in page.extract_words() if w["top"] > (heading_top or 0) + 2]
    if not words:
        return []

    # Landscape spread that carries TWO independent statements/notes side by
    # side (du's P&L next to changes-in-equity; Etisalat's note 7 next to
    # note 8) -- keep only the half the matched heading sits in. Distinguish
    # this from ONE wide statement that merely spans a landscape page (e&'s
    # integrated-report P&L: labels far left, figures far right) by requiring
    # a heading-like line in BOTH halves at a similar height.
    mid = page.width / 2
    hx = heading_x0 if heading_x0 is not None else mid
    _HEADINGISH = re.compile(
        r"^(consolidated statement of|statement of (financial position|profit|comprehensive|"
        r"changes in equity|cash flow)|notes to the|\d{1,2}[.\)]\s+[A-Z]|[a-z]\)\s+[A-Z])", re.I)
    band = build_lines([w for w in page.extract_words()
                        if (heading_top or 0) - 4 <= w["top"] <= (heading_top or 0) + 40])
    left_head = any(_HEADINGISH.match(l["text"]) and l["x0"] < mid for l in band)
    right_head = any(_HEADINGISH.match(l["text"]) and l["x0"] >= mid for l in band)
    two_up = page.width > 720 and left_head and right_head

    def build_from(subset):
        ws = subset
        ftop = None
        for line in build_lines(ws, x_gap_split=10_000):
            if _NOTES_FOOTER_RE.search(line["text"]):
                ftop = line["top"]
                break
        if ftop is not None:
            ws = [w for w in ws if w["top"] < ftop - 1]
        ns = [w for w in ws if _PAGE_NUMTOKEN_RE.match(w["text"])]
        if len(ns) < 4:
            return []
        vxs = min(w["x0"] for w in ns)
        lxe = vxs - 6
        lbl_ws = [w for w in ws
                  if w["x1"] <= vxs - 2 and re.search(r"[A-Za-z]", w["text"])]
        lx0s = sorted(w["x0"] for w in lbl_ws)
        lft = lx0s[max(0, len(lx0s) // 10)] if lx0s else min(w["x0"] for w in ws)
        # The label column's left edge is where label LINES start, not where
        # the average label WORD starts -- a column of long labels has most of
        # its words mid-sentence, which drags a word-percentile rightwards and
        # clips the first word off every long label (du 2016/2019 landscape).
        # Take the leftmost line-start bucket that holds two or more lines
        # (a lone line further left is rotated page furniture).
        # (landscape only -- portrait single-column statements reconstruct
        # fine on the word percentile, and the line-bucket can drift left into
        # a year-header token there.)
        if page.width > 720:
            lbl_lines = build_lines(lbl_ws, x_gap_split=10_000)
            line_x0s = sorted(ln["x0"] for ln in lbl_lines)
            if line_x0s:
                from collections import Counter as _CC
                buckets = _CC(round(x / 6) for x in line_x0s)
                dense = sorted(b * 6 for b, n in buckets.items() if n >= 2)
                lft = dense[0] if dense else line_x0s[0]
        rgt = max(w["x1"] for w in ns) + 6
        bot = max(w["bottom"] for w in ws)
        rr = reconstruct_unruled_rows(page, lft - 2, rgt, (heading_top or 0) + 2,
                                      bot + 2, lxe, stop_at_text=page.width > 700)
        rr = [[clean_cell(c) for c in r] for r in rr]
        return normalize_missing_notes_column(rr)

    if two_up:
        if hx < mid:
            words = [w for w in words if w["x1"] <= mid + 0.03 * page.width]
        else:
            words = [w for w in words if w["x0"] >= hx - 24]
        rows = build_from(words)
    else:
        # A single statement can WRAP down the left half of a landscape page
        # and continue down the right half (du 2025) -- as opposed to one wide
        # statement with labels far left and figures far right (e& integrated
        # P&L). It's wrapped iff BOTH halves independently carry a column of
        # figures.
        left_ws = [w for w in words if w["x1"] <= mid + 15]
        right_ws = [w for w in words if w["x0"] >= mid - 15]

        def _bigfigs(wl):
            # real statement figures (thousands+), not stray page furniture,
            # AND clustered in a column (a genuine value column, not scatter)
            xs = sorted(w["x0"] for w in wl if _PAGE_NUMTOKEN_RE.match(w["text"])
                        and re.sub(r"[^\d]", "", w["text"]) and len(re.sub(r"[^\d]", "", w["text"])) >= 4)
            if len(xs) < 4:
                return 0
            from collections import Counter as _C
            band = _C(round(x / 15) for x in xs).most_common(1)[0][1]
            return band
        if page.width > 720 and _bigfigs(left_ws) >= 4 and _bigfigs(right_ws) >= 4:
            rows = build_from(left_ws) + build_from(right_ws)
        else:
            rows = build_from(words)

    rows = [r for r in rows if any(c is not None and str(c).strip() for c in r)]
    return _strip_page_furniture(rows)


def _drop_dead_leading_col(rows):
    """If EVERY figure-bearing row has None (or a bare note number) in cell 1,
    that column is dead padding -- remove it so these rows line up with rows
    that don't have it. Used only when stitching the two halves of a wrapped
    landscape statement together."""
    fig = [r for r in rows if any(isinstance(c, (int, float)) and not isinstance(c, bool)
                                  for c in r[1:])]
    if not fig:
        return rows
    def dead(c):
        return c is None or (isinstance(c, (int, float)) and abs(c) < 100)
    if all(len(r) > 2 and dead(r[1]) for r in fig):
        return [[r[0]] + list(r[2:]) for r in rows]
    return rows


def reconstruct_beside_ruled(page, mt):
    """A pdfplumber ruled table that boxed only the FIGURES (Etisalat's P&L,
    older du landscape statements) comes back with many rows unlabelled --
    the labels sit outside the box to the left. When that's the case, redo
    the whole region from word positions instead (labels included), keyed on
    where the label column actually starts."""
    # only for a box ruled tight around ONE or TWO value columns -- that's
    # the "figures boxed, labels outside" case. A wider box is a normal table
    # (or a 2-up spread) and re-reading it here would just pull in the
    # neighbouring column's prose.
    if mt["x1"] - mt["x0"] > 170:
        return None
    body = mt["rows"]
    fig_rows = [r for r in body
                if any(isinstance(c, (int, float)) and not isinstance(c, bool) for c in r)]
    if len(fig_rows) < 5:
        return None
    labelled = sum(1 for r in fig_rows
                   if row_label(r) and re.search(r"[A-Za-z]{3,}", row_label(r)))
    if labelled >= 0.85 * len(fig_rows):
        return None
    left_words = [w for w in page.extract_words()
                  if mt["top"] - 4 <= w["top"] <= mt["bottom"] + 6
                  and w["x1"] <= mt["x0"] - 2 and re.search(r"[A-Za-z]{2,}", w["text"])]
    if len(left_words) < 5:
        return None
    label_x0 = min(w["x0"] for w in left_words)
    # everything right of the box is the prior-year column(s); grab a little
    right_words = [w for w in page.extract_words()
                   if mt["top"] - 4 <= w["top"] <= mt["bottom"] + 6
                   and mt["x1"] < w["x0"] <= mt["x1"] + 120
                   and re.match(r"^[()\-–\d,]+$", w["text"])]
    right_lim = (max(w["x1"] for w in right_words) + 6) if right_words else mt["x1"] + 4
    rows = reconstruct_unruled_rows(page, label_x0 - 2, right_lim,
                                    mt["top"] - 2, mt["bottom"] + 4,
                                    mt["x0"] - 4, stop_at_text=False)
    rows = [[clean_cell(c) for c in r] for r in rows]
    rows = normalize_missing_notes_column(rows)
    rows = [r for r in rows if any(c is not None and str(c).strip() for c in r)]
    return _strip_page_furniture(rows) or None


_FURNITURE_RE = re.compile(
    r"^(ud|du|e&|\d{1,4}|for the year ended.*|as at .*|the year ended.*|"
    r"aed\s*['’]?000|notes?|"
    r"(31\s+december|december\s+31)(\s+\d{4})?)$", re.I)
_SOCE_FRAGMENT_RE = re.compile(
    r"^(at \d|total comprehensive|transfer (to|from) (other )?reserves?|"
    r"(final|interim)(\b|$)|dividend|balance at|other movements in equity|"
    r"fair value changes on financial asset|actuarial (loss|gain)|"
    r"items (that|which) (are|will|may))", re.I)
# rows from the share-option / ESOP schedules that sit on the same page as
# du's older "General and administrative expenses" note
_SCHEME_FRAGMENT_RE = re.compile(
    r"(grant scheme|\besop\b|share scheme|per option|stock price at|"
    r"measurement\s+expected|risk-free|retention rate|vesting|exercise price|"
    r"options? (granted|outstanding|forfeited|exercised)|black-scholes|"
    r"^(opening|closing) balance$|transfer to statutory reserve|"
    r"at 1 january|at 31 december)", re.I)


_SOCE_INLINE_RE = re.compile(
    r"\s+(final|interim)\s+(cash|scrip|special)(\s+dividend)?\b(?!\s*$)", re.I)


def _strip_page_furniture(rows):
    """Drop rotated margin text ('ud', a page number), the 'For the year
    ended 31 December' sub-title, and bare year-header rows that a word-level
    reconstruction picks up alongside the real statement -- they otherwise
    feed junk figures into the arithmetic check. Also drop everything before
    the first genuine line item (a label with real words AND a figure)."""
    def is_furniture(r):
        lbl = row_label(r)
        nums = [c for c in r if isinstance(c, (int, float)) and not isinstance(c, bool)]
        if lbl is None or not re.search(r"[A-Za-z]", lbl):
            if nums and all(2000 <= n <= 2099 for n in nums):
                return True                                   # a stray year-header row
            return len(nums) <= 1 and all(abs(n) < 3000 for n in nums)   # lone page no. / stray
        if _FURNITURE_RE.match(lbl.strip()):
            return True
        if re.fullmatch(r"(19|20)\d\d\s*[*†‡]?(\s+(19|20)\d\d\s*[*†‡]?)*", lbl.strip()):
            return True
        # a value-less fragment of the NEIGHBOURING statement-of-changes-in-
        # equity that leaked in from the right column of a landscape spread
        if not nums and _SOCE_FRAGMENT_RE.search(lbl):
            return True
        # a row of the share-option schedule that shares du's older G&A page
        if _SCHEME_FRAGMENT_RE.search(lbl):
            return True
        return False

    rows = [r for r in rows if not is_furniture(r)]

    # A changes-in-equity phrase ("Final cash", "Interim cash [dividend]") can
    # bleed from the neighbouring landscape column into the MIDDLE of a real
    # label (du 2024: "Share of loss on investment accounted Final cash for
    # using the equity method"). Splice it back out.
    for r in rows:
        for ci, c in enumerate(r):
            if isinstance(c, str) and _SOCE_INLINE_RE.search(c):
                r[ci] = re.sub(r"\s{2,}", " ", _SOCE_INLINE_RE.sub(" ", c)).strip()

    first = next((i for i, r in enumerate(rows)
                  if row_label(r) and re.search(r"[A-Za-z]{3,}", row_label(r))
                  and any(isinstance(c, (int, float)) and not isinstance(c, bool) for c in r)), 0)
    # keep a value-less section heading that sits directly above the first
    # line item (du 2025's statement opens with a bare "Revenue" heading over
    # Mobile / Fixed / Wholesale / ICT) -- it survived the furniture filter,
    # so it is real, not margin cruft
    if first > 0:
        h = rows[first - 1]
        hl = row_label(h)
        if (hl and re.fullmatch(
                r"(revenues?|income|turnover|continuing operations|"
                r"discontinued operations|operating (income|expenses|activities))",
                hl.strip(), re.I)
                and not any(isinstance(c, (int, float)) and not isinstance(c, bool) for c in h)):
            first -= 1
    return rows[first:]


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
    def is_crammed(group):
        # pdfplumber couldn't find row rules inside the box, so several
        # source rows collapsed into one cell: the cell text then carries a
        # newline, or two+ separate number groups ("(4,214,432) 54,970
        # (83,752) ..."), or a run of note-reference numbers ("21 23 23 24
        # 7"). Such a group needs rebuilding from word positions even though
        # its row count looks healthy.
        multi_num = re.compile(r"[-(]?\d[\d,]*\)?[ \n]+[-(]?\d[\d,]*")
        for r in group["rows"]:
            for c in r:
                if isinstance(c, str) and ("\n" in c or multi_num.search(c)):
                    return True
        return False

    expanded = []
    for i, g in enumerate(merged):
        if len(g["rows"]) > 2 and not is_crammed(g):
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
        # A crammed group (not a bare stub) already knows its own extent --
        # don't let the crop run down the rest of the page into unrelated
        # prose; a small margin past its own bottom edge is enough.
        if len(g["rows"]) > 2:
            bottom_limit = min(bottom_limit, g["bottom"] + 15)
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

        crop_x0 = g["x0"] - 5

        # pdfplumber frequently rules a tight box around JUST the value
        # columns of an unruled statement/note (older du notes, and the
        # Etisalat/e& and old-du landscape statements -- the numeric block
        # often comes back as one multi-line cell), leaving the row labels
        # outside the box to the LEFT. If real wordy text sits in the strip
        # immediately left of the box, at the box's own height, widen the
        # crop to pull it in and treat everything left of the box edge as
        # label text. Guarded so we don't reach into a genuinely separate
        # neighbouring column: skip if another detected table occupies that
        # left strip at the same height.
        left_blocked = any(
            gg is not g
            and gg["x1"] <= g["x0"] + 5
            and gg["x1"] > g["x0"] - MAX_LABEL_LOOKUP_DIST
            and not (gg["bottom"] < g["top"] or gg["top"] > bottom_limit)
            for gg in merged
        )
        left_strip_words = [
            w for w in page.extract_words()
            if g["x0"] - MAX_LABEL_LOOKUP_DIST <= w["x0"] < g["x0"] - 1
            and g["top"] - 2 <= (w["top"] + w["bottom"]) / 2 <= bottom_limit
            and re.search(r"[A-Za-z]{3,}", w["text"])
        ]
        if not left_blocked and len(left_strip_words) >= 3:
            crop_x0 = max(0, g["x0"] - MAX_LABEL_LOOKUP_DIST)
            label_x_end = g["x0"] - 2

        new_rows = reconstruct_unruled_rows(page, crop_x0, g["x1"] + 5, g["top"], bottom_limit, label_x_end)
        new_rows = [[clean_cell(c) for c in r] for r in new_rows]
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
    return re.sub(r"\s+", " ", _despace(text or "")).strip()


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


def table_match_score(body_rows, reference_labels, threshold=FUZZY_ROW_MATCH_THRESHOLD,
                      anti_labels=None):
    """
    Fraction of reference_labels that fuzzy-match SOME row label in this
    table. 1.0 = every known row label was found; 0.0 = none were.
    No reference labels configured -> None (can't score content).

    anti_labels: row labels that positively identify a DIFFERENT table which
    would otherwise score well on the same vocabulary (e.g. the segmental-
    information note, which shares most of the P&L's line items but also has
    "Inter-segment revenue" / "Segment result" that the real P&L never has).
    Each anti-label found subtracts 0.5 from the score, so a table carrying
    two or more of them can't win.
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
    score = matched / len(reference_labels)
    for anti in (anti_labels or []):
        if max((label_similarity(anti, rl) for rl in row_labels), default=0.0) >= threshold:
            score -= 0.5
    return score


_NOTE_HEADING_RE = re.compile(r"^\s*\d{1,2}[.\)]?\s+[A-Z]")
_DATEISH_RE = re.compile(
    r"^\s*\d{1,2}\s+(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)", re.I)


def _is_note_heading(label):
    """'7. Operating expenses' / '26 Operating expenses' -- but NOT a date
    like '31 December 2024', which the bare pattern would otherwise match."""
    s = normalize(label or "")
    return bool(_NOTE_HEADING_RE.match(s)) and not _DATEISH_RE.match(s)


def trim_note_table(body, target):
    """
    A reconstructed note frequently comes with neighbours attached: the rows
    of the PREVIOUS note above it (down to and including its "<N>. <Title>"
    heading), and the start of the NEXT note below it. Keep only the target
    note's own rows:

      * drop everything up to and including the last row whose label matches
        one of target['heading_patterns'] (the note's own heading line), if
        such a row exists;
      * then stop at the first row that begins a different note
        ("<N>. Another title"), or is plainly running prose (a long,
        number-free sentence), or -- once at least three item rows have been
        seen -- an unlabelled all-numeric row, which is the note's own total
        and the last row worth keeping.
    """
    pats = target.get("heading_patterns", [])

    def is_real_heading_row(r):
        """A row that is genuinely the NOTE'S OWN HEADING -- not a totals line
        like 'Operating expenses (before federal royalty)' that merely
        contains the same words. Real headings are '<N>. Title' or a short
        standalone phrase with no figures, parentheses or 'total'/'before'."""
        lbl = row_label(r)
        if not lbl or not any(p.search(normalize(lbl).upper()) for p in pats):
            return False
        if any(isinstance(c, (int, float)) and not isinstance(c, bool) for c in r):
            return False
        if _is_note_heading(lbl):
            return True
        return (len(lbl.split()) <= 6 and "(" not in lbl
                and not re.search(r"\b(before|after|total|net)\b", lbl, re.I))

    start = 0
    for i, r in enumerate(body):
        if is_real_heading_row(r):
            start = i + 1
    rows = body[start:]

    # drop rows belonging to a co-located share-option / statutory-reserve
    # schedule (du's older G&A note shares its page with these), and stray
    # repeated column-header rows (bare years, bare "AED 000" units)
    def _is_hdr_row(r):
        vals = [c for c in r if c is not None and str(c).strip()]
        if not vals:
            return False
        return all(
            re.fullmatch(r"\(?(19|20)\d\d\)?", str(c).strip())
            or re.fullmatch(r"(AED|USD|EGP|SAR)\s*['’]?\s*000", str(c).strip(), re.I)
            or re.fullmatch(r"(31\s+december|december\s+31)(\s+\d{4})?", str(c).strip(), re.I)
            or str(c).strip().lower() in ("notes", "note")
            for c in vals)

    rows = [r for r in rows
            if not _is_hdr_row(r)
            and not (row_label(r) and (_SCHEME_FRAGMENT_RE.search(row_label(r))
                                       or _SOCE_FRAGMENT_RE.search(row_label(r))))]

    out, seen_item, blank_run = [], False, 0
    for r in rows:
        lbl = row_label(r)
        has_words = bool(lbl and re.search(r"[A-Za-z]{3,}", lbl))
        nums = [c for c in r if isinstance(c, (int, float)) and not isinstance(c, bool)]
        # A new "<N>. Another note" heading, or a long number-free sentence
        # (running commentary between notes), ends this note.
        if out and lbl and _is_note_heading(lbl):
            break
        if out and has_words and len(lbl) > 60 and not nums:
            break
        # After we've collected the note's items and its total, a couple of
        # fully blank rows means the next note has started.
        if seen_item and not lbl and not nums:
            blank_run += 1
            if blank_run >= 2:
                break
            continue
        blank_run = 0
        # a dangling connective wrapped off the next paragraph ("and", "of the")
        # after the note's items -- stop here
        if (seen_item and has_words and not nums
                and re.fullmatch(r"(and|or|of|the|to|for|in|on)( (and|or|of|the|to|for|in|on))?",
                                 lbl.strip(), re.I)):
            break
        out.append(r)
        if has_words and nums:
            seen_item = True

    # Two note items whose labels collided in the PDF's own text layer because
    # an unrelated table is interleaved on the page: du 2012 prints
    # "Consulting" and "Telecommunications licence and related fees" as
    # "ConsultingTelecommunications licence and" / "related fees" (two rows,
    # values still correctly paired). Split the first, fold the tail into the
    # second.
    for i in range(len(out) - 1):
        a, b = out[i], out[i + 1]
        la, lb = row_label(a), row_label(b)
        if not (la and lb):
            continue
        m = re.match(r"^([A-Z][a-z]{2,})([A-Z][a-z].{3,})$", la.strip())
        a_nums = [c for c in a if isinstance(c, (int, float)) and not isinstance(c, bool)]
        b_nums = [c for c in b if isinstance(c, (int, float)) and not isinstance(c, bool)]
        if (m and a_nums and b_nums
                and lb.strip()[:1].islower() and len(lb.split()) <= 4):
            head, rest = m.group(1), m.group(2).strip()
            out[i] = [head] + list(a[1:])
            out[i + 1] = [f"{rest} {lb.strip()}"] + list(b[1:])
            break

    return out if out else body


_PL_TAIL_JUNK_RE = re.compile(
    r"\s+(other\s+comprehensive\s+(income|loss|\(loss\)|\(income\)).*"
    r"|attributable\s+to[\s:].*"
    r"|profit(\s+attributable.*)?)$", re.I)


def truncate_at_row(body, target):
    """
    Cut a table's body right after the row matching target's cutoff (exact-ish
    row_patterns first, then a fuzzy fallback against cutoff_reference_labels).
    Returns (possibly-shortened body, was_confirmed).
    """
    # Try each row_pattern in priority order; within a pattern take the LAST
    # matching row. An Etisalat-style statement shows an interim "Profit for
    # the year from continuing operations", then a discontinued-operations
    # block, then the final "Profit for the year" total -- the bare
    # "PROFIT FOR THE YEAR$" pattern is listed first precisely so that its
    # (last) match, the real bottom line, wins over the "...from continuing
    # operations" pattern which is only reached when no bare total exists.
    for pat in target.get("row_patterns", []):
        last_i = None
        for i, row in enumerate(body):
            label = row_label(row)
            if not label:
                continue
            if pat.match(normalize(label).upper()):
                last_i = i
        if last_i is not None:
            out = [list(r) for r in body[: last_i + 1]]
            # The cutoff row can come back with the NEXT section's heading
            # glued onto its label ("... Profit for the year  Other
            # comprehensive income/(loss)") -- keep only the total's own label.
            for ci, c in enumerate(out[-1]):
                if isinstance(c, str) and c.strip():
                    out[-1][ci] = _PL_TAIL_JUNK_RE.sub("", c).strip()
                    break
            return out, True

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


# ------------------------------------------------------ arithmetic checks ---
# The DFM handoff's single highest-value recommendation: don't trust an
# extraction until its numbers reconcile. Used two ways -- to PASS/FAIL each
# written table, and (more importantly) to CHOOSE between candidate tables:
# a wrong table almost never reconciles.

def _cell_num(v):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else None


# trailing note ref -- "Revenue 6 (b)" or, with Etisalat's reversed
# parentheses, "Revenue 6 )b("
_TRAIL_NOTE_RE = re.compile(r"\s+\d{1,2}(\.\d+)?\s*(\([a-z0-9]{1,3}\)|\)[a-z0-9]{1,3}\()?\s*$")
_LEAD_NOTE_RE = re.compile(r"^\s*\d{1,2}\s+(?=[\d()\-–])")
# a trailing sub-note marker with no leading digit, in any bracket/quote style:
# ")i(", "(i)", '"i)"', "“i)”"
_BARE_MARKER_RE = re.compile(
    r"\s+([)(\"“”‘’]{1,2}[a-z]{1,2}[)(\"“”‘’]{1,2})\s*$", re.I)

# label (normalised, lower-cased) -> note reference, for the note column that
# _strip_note_refs last removed. convert_pdf reads it to re-attach a display
# "Note" column to the written sheet, AFTER the arithmetic check has run on
# the clean body.
_DROPPED_NOTE_COL = {}


def _fmt_note_ref(v):
    """Render a captured note reference the way it reads in the report:
    an integer as "6", a one-decimal sub-note as "8.1", a marker as-is."""
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, int):
        return str(v)
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else f"{v:g}"
    s = str(v).strip()
    # Etisalat prints sub-note markers reversed: ")a(" -> "(a)"
    s = re.sub(r"\)([a-z0-9]{1,3})\(", r"(\1)", s, flags=re.I)
    s = re.sub(r"\s+", " ", s)
    # normalise "12(a)" / "12  (a)" -> "12 (a)"
    s = re.sub(r"^(\d{1,2}(?:\.\d+)?)\s*(\([a-z0-9]{1,3}\))$", r"\1 \2", s, flags=re.I)
    return s


def _strip_note_refs(body):
    """Note-reference numbers land inconsistently -- sometimes tucked into
    the label ("Revenue 6 (b)"), sometimes glued onto the first figure
    ("16 2,475,403") depending on how wide the label is. Strip them from
    both places so every row's figures sit in the same columns (and the
    labels read clean)."""
    _DROPPED_NOTE_COL.clear()
    _g1 = r"(?:[()]?[-–]?[\d,]+[()]?|[-–])"
    glued = re.compile(rf"^{_g1}\s+{_g1}$")

    # A bare "2013 2012" year-header row (no word label, every value a year)
    # that a word-level reconstruction picked up above the first line item --
    # it would otherwise be summed into the arithmetic check and written as a
    # stray row.
    def _is_bare_year_row(r):
        if row_label(r):
            return False
        vals = []
        for c in r:
            if isinstance(c, (int, float)) and not isinstance(c, bool):
                vals.append(c)
            elif isinstance(c, str) and re.fullmatch(r"\(?(19|20)\d\d\)?", c.strip()):
                vals.append(int(re.sub(r"[^\d]", "", c)))
            elif c is not None and str(c).strip():
                return False
        return bool(vals) and all(1990 <= v <= 2099 for v in vals)

    def _is_unit_header_row(r):
        # every non-empty cell is a column-header token: a bare year, an
        # "AED 000" unit, "Note(s)", or a "31 December" date -- a header line
        # a word-level reconstruction picked up as a body row.
        vals = [c for c in r if c is not None and str(c).strip()]
        if not vals:
            return False
        return all(
            re.fullmatch(r"\(?(19|20)\d\d\)?", str(c).strip())
            or re.fullmatch(r"(AED|USD|EGP|SAR)\s*['’]?\s*000", str(c).strip(), re.I)
            or re.fullmatch(r"(31\s+december|december\s+31)(\s+\d{4})?", str(c).strip(), re.I)
            or str(c).strip().lower() in ("notes", "note")
            for c in vals)

    def _is_squashed_header_row(r):
        # a whole column-header line collapsed into one label cell, no figures
        # ("For the year ended 31 December Note AED 000 AED 000")
        if any(isinstance(c, (int, float)) and not isinstance(c, bool) for c in r):
            return False
        lbl = row_label(r) or ""
        return bool(re.search(r"for the (year|period) ended", lbl, re.I)
                    and re.search(r"AED\s*['’]?\s*000", lbl, re.I))

    body = [r for r in body
            if not _is_bare_year_row(r) and not _is_unit_header_row(r)
            and not _is_squashed_header_row(r)]
    out = []
    for r in body:
        r = list(r)
        _note_here = None
        if r and isinstance(r[0], str):
            m_tr = _TRAIL_NOTE_RE.search(r[0])
            stripped = _TRAIL_NOTE_RE.sub("", r[0]).strip()
            # a "Notes" column-header word that merged onto the first line item
            stripped = re.sub(r"^\s*Notes?\s+(?=[A-Z])", "", stripped)
            # a bare parenthesised sub-note marker with NO leading number
            # glued to the label -- "Amortisation )i(", 'Regulatory expenses "i)"'
            m_bare = _BARE_MARKER_RE.search(stripped)
            if m_bare and re.search(r"[A-Za-z]{3,}", stripped[:m_bare.start()]):
                if not m_tr:
                    _note_here = _fmt_note_ref(
                        re.sub(r'[)("“”‘’]', "",
                               m_bare.group(1)))
                    _note_here = f"({_note_here})"
                stripped = stripped[:m_bare.start()].strip()
            if stripped and re.search(r"[A-Za-z]{3,}", stripped):
                if m_tr:
                    _note_here = _fmt_note_ref(m_tr.group().strip())
                r[0] = stripped
        # note-ref glued onto the first figure as a string ("31 (2,167,933)")
        for i in range(1, len(r)):
            if isinstance(r[i], str):
                m = _LEAD_NOTE_RE.match(r[i])
                if m:
                    if _note_here is None:
                        _note_here = r[i][:m.end()].strip()
                    r[i] = clean_cell(r[i][m.end():])
        # key the captured ref by the CLEANED label so convert_pdf can look
        # it up against the final written rows
        _lbl_clean = row_label(r)
        if _note_here and _lbl_clean:
            _DROPPED_NOTE_COL.setdefault(normalize(_lbl_clean).lower(), _note_here)
        # two figures that ended up in one cell ("(2,167,933) (2,153,590)")
        for i in range(1, len(r)):
            if isinstance(r[i], str) and glued.match(r[i].strip()):
                a, b = r[i].split()
                if i + 1 < len(r) and r[i + 1] is None:
                    r[i], r[i + 1] = clean_cell(a), clean_cell(b)
                elif i + 1 == len(r):
                    r[i] = clean_cell(a); r.append(clean_cell(b))
        out.append(r)

    # Whole-column note-reference removal. If cell index 1 (right after the
    # label) is, across the body, either empty or a small positive integer
    # -- never a real figure -- it's a Notes column: delete it from every
    # row so all figures line up in the same columns for the arithmetic
    # check. (Doing this per-row instead would shorten only the rows that
    # had a note number, re-introducing the misalignment.) The label->ref
    # map is stashed in _DROPPED_NOTE_COL (already cleared at function entry)
    # so the caller can re-attach the column to the final written sheet
    # without it ever touching the arithmetic check.
    def _isnum(x):
        return isinstance(x, (int, float)) and not isinstance(x, bool)
    wide = [r for r in out if len(r) >= 4]
    narrow_fig = [r for r in out if len(r) == 3
                  and any(_isnum(c) for c in r[1:])]
    if wide and len(out) >= 4:
        col1 = [r[1] for r in wide]
        # a note ref is a small number, possibly with one decimal ("23.1")
        def _noterefish(v):
            if v is None:
                return True
            if _isnum(v):
                return 0 < v < 100 and abs(v * 10 - round(v * 10)) < 1e-6
            # a short parenthesised note marker, either orientation: "(i)" ")i("
            return bool(re.fullmatch(r"[)(]?[a-z0-9]{1,3}[)(]?", str(v).strip(), re.I))
        noteish = sum(1 for v in col1 if _noterefish(v))
        realfig = sum(1 for v in col1 if _isnum(v) and abs(v) >= 100)
        nrefs = sum(1 for v in col1
                    if (_isnum(v) and 0 < abs(v) < 100)
                    or (v is not None and not _isnum(v) and str(v).strip()))
        # A blank column that only carries note refs is always removed so the
        # figures line up.  It becomes a re-attachable "Note" column only when
        # it holds >= 2 refs; a lone stray "(i)" marker is just dropped (its
        # own phantom column and unlabelled header are the defect, not a
        # feature).
        if realfig == 0 and noteish >= len(col1) * 0.9 and (nrefs >= 1 or narrow_fig):
            if nrefs >= 2:
                for r in out:
                    if len(r) >= 4 and r[1] is not None and str(r[1]).strip():
                        lbl = row_label(r)
                        if lbl:
                            _DROPPED_NOTE_COL[normalize(lbl).lower()] = _fmt_note_ref(r[1])
            out = [([r[0]] + list(r[2:])) if len(r) >= 4 else r for r in out]
    return out


def _value_columns(body):
    """Column indices that hold statement figures -- not the label column,
    and not a note-reference column (small integers, present on only some
    rows)."""
    if not body:
        return []
    width = max(len(r) for r in body)
    cols = []
    for c in range(1, width):
        vals = [_cell_num(r[c]) for r in body if c < len(r)]
        vals = [v for v in vals if v is not None]
        if len(vals) < 3:
            continue
        small = [v for v in vals if abs(v) < 1000]
        if len(small) == len(vals) and len(vals) < 0.7 * len(body):
            continue  # looks like a Notes-reference column
        cols.append(c)
    return cols


def _last_num_in(body, c):
    for r in reversed(body):
        v = _cell_num(r[c]) if c < len(r) else None
        if v is not None:
            return v
    return None


def _is_total_label(lbl):
    if not lbl:
        return True                      # an unlabelled figure row -> the total
    s = normalize(lbl).lower()
    return bool(re.search(r"\btotal\b", s)) or \
        bool(re.fullmatch(r"operating expenses(\s*\(before.*)?", s)) or \
        "before federal royalty" in s or "before royalty" in s


def reconcile_note(body):
    """sum(line items) == printed total, for every value column. The total
    row may be unlabelled, labelled 'Total ...', or a worded line such as
    'Operating expenses (before federal royalty)' / 'Operating expenses'
    (Etisalat/e&). Interleaved neighbour-table cruft (a value with a label
    that isn't a plausible note line) is ignored."""
    cols = _value_columns(body)
    out = []
    for c in cols:
        seq = []            # (kind, value)  kind in {"item","bare"}
        for r in body:
            v = _cell_num(r[c]) if c < len(r) else None
            if v is None:
                continue
            lbl = row_label(r)
            if lbl and not re.search(r"[A-Za-z]{3,}", lbl):
                continue                 # numeric-only junk label
            seq.append(("bare" if _is_total_label(lbl) else "item", v))

        items = [v for k, v in seq if k == "item"]
        total = next((v for k, v in reversed(seq) if k == "bare"), None)

        if total is None and len(items) >= 3:
            # no explicit total row -- treat the last line as the total if it
            # equals the sum of the ones before it (also try after dropping a
            # stray leading row that word-level reconstruction sometimes adds)
            for k in (0, 1, 2):
                if len(items) - k >= 3 and abs(items[-1] - sum(items[k:-1])) <= RECON_TOL:
                    total, items = items[-1], items[k:-1]
                    break

        if total is not None and len(items) >= 2:
            out.append({"col": c, "n": len(items), "sum": sum(items),
                        "total": total, "ok": abs(sum(items) - total) <= RECON_TOL})
    return {"kind": "note", "ok": bool(out) and any(x["ok"] for x in out), "cols": out}


_PL_SUBTOTAL_RE = re.compile(
    r"^\s*(total\b|gross (profit|margin)|net operating|operating profit|"
    r"profit before|loss before|profit/\(loss\) before|earnings before|ebitda|"
    r"(net\s+)?profit for the (year|period)|profit for the (year|period)"
    r"\s+from\s+continuing)", re.I)


def reconcile_pl(body):
    """The final 'Profit for the year' must equal the signed sum of the LEAF
    lines above it (expenses already negative). Two kinds of row are skipped
    from that sum: one whose value already equals the running cumulative
    total (a rolling subtotal like 'Gross profit'), and one whose label is a
    subtotal phrase ('Total direct costs', 'Operating profit before ...',
    'Profit before federal royalty ...') -- du's 2025 statement carries
    per-section subtotals that a purely cumulative check would double-count."""
    cols = _value_columns(body)
    out = []
    for c in cols:
        rows = [(row_label(r), _cell_num(r[c]) if c < len(r) else None) for r in body]
        rows = [(lbl, v) for lbl, v in rows if v is not None]
        if len(rows) < 3:
            continue
        target = rows[-1][1]
        running = 0.0
        leaves = 0
        for lbl, v in rows[:-1]:
            if lbl and _PL_SUBTOTAL_RE.match(lbl):
                continue
            if leaves >= 1 and abs(v - running) <= RECON_TOL:
                continue
            running += v
            leaves += 1
        out.append({"col": c, "sum": running, "final": target, "n": leaves,
                    "ok": leaves >= 2 and abs(running - target) <= RECON_TOL})
    return {"kind": "pl", "ok": bool(out) and any(x["ok"] for x in out), "cols": out}


def reconcile(target, body):
    kind = target.get("reconcile")
    if kind == "pl":
        return reconcile_pl(body)
    if kind == "note":
        return reconcile_note(body)
    return None


def recon_summary(recon):
    """One-line human summary of a reconcile() result for the sheet banner."""
    if not recon:
        return ""
    ok_cols = [x for x in recon["cols"] if x["ok"]]
    if recon["ok"]:
        x = ok_cols[0]
        if recon["kind"] == "note":
            return f"PASS  (sum of {x['n']} items = {x['sum']:,.0f} = total)"
        return f"PASS  (sum of lines = {x['sum']:,.0f} = Profit for the year)"
    if not recon["cols"]:
        return "NOT CHECKED  (couldn't identify a value column / total)"
    x = recon["cols"][0]
    if recon["kind"] == "note":
        return f"FAIL  (sum of {x['n']} items = {x['sum']:,.0f}, total says {x['total']:,.0f})"
    return f"FAIL  (sum of lines = {x['sum']:,.0f}, Profit for the year says {x['final']:,.0f})"


def _anchor_ok(body, anchors, need=2):
    if not anchors:
        return True
    labels = [row_label(r) for r in body if row_label(r)]
    hits = sum(
        1 for a in anchors
        if max((label_similarity(a, l) for l in labels), default=0.0) >= FUZZY_ROW_MATCH_THRESHOLD
    )
    return hits >= need


def _trim_for(target, body):
    if target.get("mode") == "table_until_row":
        return truncate_at_row(body, target)[0]
    if target.get("mode") == "full_table":
        return trim_note_table(body, target)
    return body


def _ensure_header(pdf, cand):
    """Give a candidate a column header if it has none, by reconstructing it
    from the gap above the table (some reports draw it outside the ruled
    box). Returns the header rows (possibly the ones it already had)."""
    header, body = cand["header"], cand["body"]
    if header or not body or cand["page"] is None:
        return header
    page = pdf.pages[cand["page"] - 1]
    recon = reconstruct_column_header(page, cand["top"], cand["x0"], cand["x1"])
    if not recon:
        return header
    ncols = max((len(r) for r in body), default=1)
    out = []
    for r in recon:
        cells = [c for c in r if c is not None]
        if len(cells) == 1 and re.fullmatch(r"Notes?", str(cells[0]).strip(), re.I):
            out.append([None, cells[0]] + [None] * (ncols - 2))
        else:
            pad = ncols - len(r)
            out.append(([None] * pad + r) if pad > 0 else r)
    return out


def find_best_matching_table(pdf, page_indices, target):
    """
    Collect every candidate table for `target` -- from ruled tables on/near a
    heading-matching page, from a whole-page word reconstruction of an
    unruled statement, and (if nothing yet reconciles) from a full-document
    scan -- then rank them:

        1. arithmetic reconciles   (reconcile_pl / reconcile_note)
        2. anchor labels present   (note only: >= 2 of anchor_labels)
        3. content score

    So a look-alike (segmental note, statement of changes in equity,
    comprehensive income) that scores well on vocabulary but doesn't add up
    loses to the real statement that does.

    Returns (header, trimmed_body, page_1based, score, actual_heading, recon)
    or (None, None, None, 0.0, None, None).
    """
    reference_labels = target.get("reference_row_labels", [])
    anti_labels = target.get("anti_reference_labels", [])
    anchors = target.get("anchor_labels", [])
    cands = []
    seen_keys = set()

    def add(header, body, page, top, x0, x1):
        if not body or len(body) < 3:
            return
        body = _strip_note_refs(body)
        note_col = dict(_DROPPED_NOTE_COL)     # label -> note ref just removed
        key = (page, round(top or 0), len(body))
        if key in seen_keys:
            return
        # Score the TRIMMED body -- otherwise trailing rows of the next
        # statement (interleaved on a landscape spread) trip the anti-labels
        # and sink a perfectly good candidate before it's even considered.
        trimmed = _trim_for(target, body)
        score = table_match_score(trimmed, reference_labels, anti_labels=anti_labels)
        score = 1.0 if score is None else score
        if score < MIN_CONTENT_SCORE:
            # A low vocabulary score (label words merged in a dense
            # reconstruction) is forgiven only if the figures add up AND the
            # table is a substantial one that still shows SOME of the
            # expected vocabulary and, for the note, its anchor lines --
            # otherwise any small unrelated table that happens to foot would
            # sneak in.
            rec = reconcile(target, trimmed)
            if not (rec and rec["ok"] and len(trimmed) >= 10 and score >= 0.08
                    and _anchor_ok(trimmed, anchors)):
                return
        seen_keys.add(key)
        cands.append({"header": header, "body": body, "page": page,
                      "top": top, "x0": x0, "x1": x1, "score": score,
                      "note_col": note_col})

    # --- Pass 1: pages whose text matches a heading pattern ---------------
    checked = set()
    for idx in page_indices:
        if idx in checked or idx >= len(pdf.pages):
            continue
        checked.add(idx)
        page = pdf.pages[idx]
        hits = find_heading_hits(build_lines(page.extract_words()), target["heading_patterns"])
        if not hits:
            continue
        heading_top = hits[0]["top"]
        heading_x0 = hits[0]["x0"]
        for lookahead in range(0, 3):
            pidx = idx + lookahead
            if pidx >= len(pdf.pages):
                break
            merged = get_real_merged_tables(pdf.pages[pidx])
            if lookahead == 0:
                below = [m for m in merged if m["top"] >= heading_top - 5]
                merged = below if below else merged
            for mt in merged:
                h, b = rows_to_header_and_body(mt["rows"])
                add(h, b, pidx + 1, mt["top"], mt["x0"], mt["x1"])
                alt = reconstruct_beside_ruled(pdf.pages[pidx], mt)
                if alt:
                    ah, ab = rows_to_header_and_body(alt)
                    add(ah, ab, pidx + 1, mt["top"] - 1, 0.0, float(pdf.pages[pidx].width))
            if merged:
                break
        pr = reconstruct_page_statement(page, heading_top, heading_x0)
        if pr:
            h, b = rows_to_header_and_body(pr)
            add(h, b, idx + 1, heading_top, 0.0, float(page.width))

    def evaluate(c):
        if target.get("mode") == "table_until_row":
            trimmed, confirmed = truncate_at_row(c["body"], target)
        else:
            trimmed, confirmed = trim_note_table(c["body"], target), True
        return trimmed, reconcile(target, trimmed), _anchor_ok(trimmed, anchors), confirmed

    def any_reconciles():
        for c in cands:
            _, rec, anc, conf = evaluate(c)
            if rec and rec["ok"] and anc and conf:
                return True
        return False

    # --- Pass 2: full-document scan, only if nothing solid yet -----------
    if reference_labels and not any_reconciles():
        print(f"    scanning full document for '{target['name']}' (content + arithmetic)...")
        for i, page in enumerate(pdf.pages):
            if (i + 1) % 40 == 0:
                print(f"      ...{i + 1}/{len(pdf.pages)} pages")
            had_ruled = False
            for mt in get_real_merged_tables(page):
                had_ruled = True
                h, b = rows_to_header_and_body(mt["rows"])
                add(h, b, i + 1, mt["top"], mt["x0"], mt["x1"])
                alt = reconstruct_beside_ruled(page, mt)
                if alt:
                    ah, ab = rows_to_header_and_body(alt)
                    add(ah, ab, i + 1, mt["top"] - 1, 0.0, float(page.width))
            hh = find_heading_hits(build_lines(page.extract_words()), target["heading_patterns"])
            if hh and not had_ruled:
                pr = reconstruct_page_statement(page, hh[0]["top"], hh[0]["x0"])
                if pr:
                    h, b = rows_to_header_and_body(pr)
                    add(h, b, i + 1, hh[0]["top"], 0.0, float(page.width))

    if not cands:
        return None, None, None, 0.0, None, None, {}

    ranked = []
    for c in cands:
        c["header"] = _ensure_header(pdf, c)
        trimmed, rec, anc, conf = evaluate(c)
        # For the expense NOTE, reject a candidate that is really the whole
        # income statement (has a Revenue row and a profit subtotal) -- some
        # du years (2014-2018) print the expense breakdown on the face of the
        # statement with no separate note, and that belongs on the P&L sheet,
        # not garbled onto the note sheet.
        if target.get("mode") == "full_table":
            labs = " || ".join((row_label(r) or "").lower() for r in trimmed)
            if re.search(r"\brevenue\b", labs) and re.search(
                    r"gross profit|profit for the year|profit before (royalty|federal)", labs):
                continue
        # For the P&L, reject a candidate laid out with more than two value
        # columns -- a real income statement has exactly current + prior
        # year. Three-plus means a segmental note (revenue by geography), a
        # statement of changes in equity, or a "before / adjustment / after"
        # restatement reconciliation.
        if target.get("mode") == "table_until_row" and len(_value_columns(trimmed)) > 2:
            continue
        ranked.append((c, trimmed, rec, anc, conf))
    if not ranked:
        return None, None, None, 0.0, None, None, {}
    def _labelled_frac(rows):
        fig = [r for r in rows if any(isinstance(c, (int, float)) and not isinstance(c, bool)
                                      for c in r[1:])]
        if not fig:
            return 0.0
        return sum(1 for r in fig if row_label(r)
                   and re.search(r"[A-Za-z]{3,}", row_label(r))) / len(fig)

    ranked.sort(key=lambda t: (
        1 if t[4] else 0,                     # P&L cutoff row actually found
        1 if (t[2] and t[2]["ok"]) else 0,    # arithmetic reconciles
        1 if t[3] else 0,                     # anchor labels present (note)
        min(len(_value_columns(t[1])), 2),    # keep BOTH year columns
        round(_labelled_frac(t[1]), 1),       # prefer rows that kept their labels
        round(t[0]["score"], 2),              # vocabulary match
        len(t[1]),
    ), reverse=True)

    c, trimmed, rec, anc, conf = ranked[0]
    # For a target that MUST reconcile (the du expense note -- a real one
    # always does now), a top candidate that doesn't add up means there is
    # no such table in this report (du 2014-2018 print the breakdown on the
    # face of the P&L). Report it as absent rather than emit a wrong table.
    if target.get("require_reconcile") and not (rec and rec["ok"]):
        return None, None, None, 0.0, None, None, {}
    page = pdf.pages[c["page"] - 1]
    actual_heading = find_actual_heading(page, c["top"], x0=c["x0"], x1=c["x1"])
    return c["header"], trimmed, c["page"], c["score"], actual_heading, rec, c.get("note_col", {})


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


def _clean_header(header_rows, body, report_year=None):
    """Rebuild the column header as (at most) two tidy rows aligned to the
    body's value columns: a year row and a units row, right-aligned so the
    years sit over the figures they head. The raw reconstructed header is
    often several fragments crammed into the wrong cells ("2018"/"2017"
    stacked, a stray "Notes", a duplicated "2023 2022" tail) -- this throws
    that away and keeps only the information (which years, which unit)."""
    if not body:
        return header_rows
    if header_rows and not isinstance(header_rows[0], (list, tuple)):
        header_rows = [header_rows]
    ncols = max((len(r) for r in body), default=1)

    flat = " ".join(str(c) for hr in (header_rows or []) for c in hr if c is not None)
    seen, yr = set(), []
    for y in re.findall(r"\b(?:19|20)\d{2}\b", flat):
        if y not in seen:
            seen.add(y); yr.append(y)
    unit_m = re.search(r"(AED|USD|EGP|SAR)\s*['’]?\s*000", flat)
    unit = "AED'000" if unit_m else "AED'000"   # every one of these reports is AED'000

    # which columns of the body actually hold figures
    vcols = _value_columns(body)
    if not vcols:
        vcols = [c for c in range(1, ncols)
                 if any(isinstance(r[c], (int, float)) for r in body if c < len(r))]
    if not vcols:
        return header_rows

    ys = sorted({int(y) for y in yr}, reverse=True)
    if report_year and (not ys or len(ys) < len(vcols)):
        # fill in missing years from the file's own year (statements always
        # run most-recent-first, one calendar year per column)
        base = ys[0] if ys else report_year
        ys = [base - k for k in range(len(vcols))]
    if not ys:
        return header_rows
    # financial statements always show the most recent year in the left-most
    # column -- order the years descending regardless of the jumble we found
    yr = ys[:len(vcols)]
    yr_row = [None] * ncols

    # a leading column that isn't a value column but carries note references
    # (small ints / "(i)" markers) -- label it "Note"
    def _is_noteref(v):
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            return 0 < abs(v) < 100
        if isinstance(v, str) and v.strip():
            # "6", "8.1", "6 (b)", "(a)", "35 (b)"
            return bool(re.fullmatch(
                r"\(?\d{1,2}(\.\d+)?\)?(\s*\([a-z0-9]{1,3}\))?|\([a-z0-9]{1,3}\)",
                v.strip(), re.I))
        return False
    for col in range(1, ncols):
        if col in vcols:
            continue
        vals = [r[col] for r in body if col < len(r) and r[col] is not None
                and str(r[col]).strip()]
        if len(vals) >= 2 and sum(_is_noteref(v) for v in vals) >= max(2, 0.6 * len(vals)):
            yr_row[col] = "Note"
            break

    for col, y in zip(vcols[:len(yr)], yr):
        # keep a restatement footnote marker if the source header carries one
        m = re.search(rf"\b{y}\s*([*†‡])", flat)
        yr_row[col] = f"{y}{m.group(1)}" if m else y
    out = [yr_row]
    if unit:
        u_row = [None] * ncols
        for col in vcols:
            u_row[col] = unit
        out.append(u_row)
    return out


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


def convert_pdf(pdf_path: Path, wb: Workbook, profile: dict, hints: dict):
    """Add one sheet for `pdf_path`. Returns [(target_name, status), ...]
    where status is 'PASS' / 'FAIL' / 'not checked' / 'NOT FOUND'."""
    print(f"\nProcessing: {pdf_path.name}   [{profile['label']}]")
    ws = wb.create_sheet(title=safe_sheet_name(pdf_path.stem))

    row = 1
    row = write_title(ws, row, pdf_path.stem.replace("-", " ").replace("_", " ").title(), 6)
    row = write_subtitle(ws, row, f"Extracted from PDF \u00b7 {pdf_path.name}  ({profile['label']} profile)")
    row += 1
    max_col_seen = 1
    statuses = []

    read_path = normalize_pdf_for_reading(pdf_path) if HAVE_PYPDF else pdf_path
    with pdfplumber.open(read_path) as pdf:
        total_pages = len(pdf.pages)
        page_labels = load_page_labels(pdf_path)
        for target in profile["targets"]:
            order = page_scan_order(total_pages, hints.get(target["hint_key"]))
            header, body, page_num, score, actual_heading, rec, note_col = \
                find_best_matching_table(pdf, order, target)

            if body is not None and len(note_col) >= 2:
                # re-attach the Notes-reference column that _strip_note_refs
                # removed for the arithmetic check -- display only, inserted
                # as a fresh column 1 so figures keep their positions
                new_body = []
                for r in body:
                    r = list(r)
                    lbl = row_label(r)
                    ref = note_col.get(normalize(lbl).lower()) if lbl else None
                    new_body.append([r[0] if r else None, ref] + (r[1:] if r else []))
                body = new_body

            if body is None:
                row = write_section_label(ws, row, f"{target['name']} \u2014 NOT FOUND", 6)
                msg = "Nothing matched the vocabulary above the confidence threshold."
                if target.get("mode") == "full_table":
                    msg = ("No separate operating-expenses note in this report. "
                           "Some du reports (2014-2018) print the expense breakdown by nature "
                           "on the face of the income statement \u2014 see the P&L table above.")
                cell = ws.cell(row=row, column=1, value=msg)
                cell.font = NOT_FOUND_FONT
                cell.fill = NOT_FOUND_FILL
                row += 2
                statuses.append((target["name"], "NOT FOUND"))
                print(f"  {target['name']}: NOT FOUND")
                continue

            called = ""
            if actual_heading and label_similarity(actual_heading, target["name"]) < 0.6:
                called = f'  \u2014 report calls it "{actual_heading}"'
            page_str = display_page(page_labels, page_num)
            summ = recon_summary(rec)
            status = "PASS" if (rec and rec["ok"]) else ("not checked" if (not rec or not rec["cols"]) else "FAIL")

            row = write_section_label(
                ws, row,
                f"{target['name']}  \u2014  page {page_str}, content match {score:.0%}{called}", 6)
            # reconciliation banner
            banner = ws.cell(row=row, column=1, value=f"Arithmetic check:  {summ}")
            banner.font = NOTE_FONT if status == "PASS" else NOT_FOUND_FONT
            if status == "FAIL":
                banner.fill = NOT_FOUND_FILL
            row += 1

            _ym = re.search(r"(19|20)\d{2}", pdf_path.stem)
            header = _clean_header(header, body, int(_ym.group()) if _ym else None)
            row, ncols = write_table(ws, row, header, body)
            max_col_seen = max(max_col_seen, ncols)
            row += 1
            statuses.append((target["name"], status))
            print(f"  {target['name']}: page {page_str}, score {score:.0%}, {len(body)} rows  ->  {summ}")

    ws.freeze_panes = "A4"
    autofit_columns(ws, max_col_seen)
    return statuses


def _gather_pdfs(args):
    paths, seen = [], set()
    for arg in args.pdfs:
        p = Path(arg)
        if p.is_dir():
            found = sorted(p.glob("**/*.pdf" if args.recursive else "*.pdf"))
            if not found:
                print(f"  SKIP: no .pdf files in {p}")
            for f in found:
                if f.resolve() not in seen:
                    seen.add(f.resolve()); paths.append(f)
        elif p.exists():
            if p.resolve() not in seen:
                seen.add(p.resolve()); paths.append(p)
        else:
            print(f"  SKIP: {p} not found")
    return paths


def main():
    parser = argparse.ArgumentParser(
        description="Extract the P&L (through 'Profit for the year') and the operating-expenses "
                    "note from du and Etisalat/e& annual reports, one workbook per company, "
                    "with an arithmetic reconciliation check on every table.")
    parser.add_argument("pdfs", nargs="+",
                        help="PDF files and/or folders (a folder = every .pdf inside it).")
    parser.add_argument("--recursive", action="store_true",
                        help="Recurse into sub-folders when a folder is given.")
    parser.add_argument("--company", choices=sorted(PROFILES), default=None,
                        help="Force a profile for every input instead of auto-detecting from the file name.")
    parser.add_argument("--outdir", type=str, default=None,
                        help="Where to write the .xlsx files (default: next to the first PDF).")
    parser.add_argument("--pl-hint", type=int, default=None, help="Approx. page of the P&L (speeds the search).")
    parser.add_argument("--admin-hint", type=int, default=None, help="Approx. page of the expenses note.")
    args = parser.parse_args()

    hints = {"pl_hint": args.pl_hint, "admin_hint": args.admin_hint}
    pdf_paths = _gather_pdfs(args)
    if not pdf_paths:
        print("No PDF files found. Nothing to do.")
        sys.exit(1)

    # Route every file to a profile.
    buckets = {}
    for p in pdf_paths:
        key = args.company or profile_for_file(p)[0]
        buckets.setdefault(key, []).append(p)

    outdir = Path(args.outdir) if args.outdir else pdf_paths[0].parent
    outdir.mkdir(parents=True, exist_ok=True)

    print(f"Found {len(pdf_paths)} PDF(s):")
    for key, files in buckets.items():
        print(f"  {PROFILES[key]['label']}: {len(files)} file(s) -> {PROFILES[key]['out']}")

    grand_summary = []
    for key, files in buckets.items():
        profile = PROFILES[key]
        wb = Workbook()
        wb.remove(wb.active)
        for pdf_path in sorted(files):
            try:
                st = convert_pdf(pdf_path, wb, profile, hints)
            except Exception as e:
                import traceback
                print(f"  ERROR on {pdf_path.name}: {e}")
                traceback.print_exc()
                st = [("(processing error)", "ERROR")]
            grand_summary.append((profile["label"], pdf_path.name, st))
        if wb.sheetnames:
            out_path = outdir / profile["out"]
            wb.save(out_path)
            print(f"\nSaved {profile['label']} -> {out_path}")

    # ---- final reconciliation scoreboard -------------------------------
    print("\n" + "=" * 78)
    print("RECONCILIATION SUMMARY  (PASS = the table's own figures add up)")
    print("=" * 78)
    for label, fname, st in grand_summary:
        cells = "   ".join(f"{n.split('(')[0].strip()[:24]}: {s}" for n, s in st)
        print(f"  [{label:>13}] {fname:<46} {cells}")
    fails = [(l, f, n, s) for l, f, st in grand_summary for n, s in st if s in ("FAIL", "NOT FOUND", "not checked", "ERROR")]
    if fails:
        print(f"\n  {len(fails)} table(s) need attention:")
        for l, f, n, s in fails:
            print(f"    {s:<11} {f}  --  {n.split('(')[0].strip()}")
    else:
        print("\n  All tables reconcile.")


if __name__ == "__main__":
    main()
