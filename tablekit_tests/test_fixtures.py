"""Extraction tests against the two committed synthetic PDFs (see
tablekit_tests/fixtures/generate_fixtures.py) -- a fabricated company, never
a real one, but internally self-consistent (every subtotal reconciles)
so the arithmetic trust-layer has something real to check.

Unlike the 24 real annual-report PDFs used elsewhere in this suite (gitignored,
present only on a dev machine -- those tests skip cleanly in CI), these two
ARE committed, so this file is CI's only non-skipped, real extraction test:
one full pass through the normal digital-text path, one through the manual
box-select "OCR this region" failsafe.

The tool has separately been tested during development against real-world
financial statements (see CHANGELOG.md) -- those files aren't committed here.
"""
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
import extract_all_tables as X  # noqa: E402

FIXTURES = Path(__file__).resolve().parent / "fixtures"
DIGITAL_PDF = FIXTURES / "sample_digital_income_statement.pdf"
SCANNED_PDF = FIXTURES / "sample_scanned_balance_sheet.pdf"


def test_digital_fixture_extracts_as_a_footing_income_statement():
    tables = list(X.scan([DIGITAL_PDF], None, 2, 2))
    assert len(tables) == 1
    t = tables[0]
    assert t["kind"] == "income statement"
    assert t["years"] == [2024, 2023]
    assert t["foots"] is True
    assert t["title"] == "CONSOLIDATED INCOME STATEMENT"
    assert t["health"]["score"] == 1.0


def test_digital_fixture_figures_are_read_as_signed_numbers():
    tables = list(X.scan([DIGITAL_PDF], None, 2, 2))
    rows = {X._row_label(r): r for r in tables[0]["rows"]}
    assert rows["Revenue"][1:] == [125000, 110000]
    assert rows["Cost of revenue"][1:] == [-72000, -64000]   # parenthesized -> negative
    assert rows["Net income"][1:] == [15200, 12540]


def test_scanned_fixture_has_no_text_layer():
    # confirms this fixture actually exercises the OCR path rather than
    # silently being read by the normal text extractor
    import pdfplumber
    with pdfplumber.open(SCANNED_PDF) as pdf:
        assert (pdf.pages[0].extract_text() or "") == ""


@pytest.mark.skipif(not X.HAVE_OCR, reason="Tesseract binary not installed on this machine")
def test_scanned_fixture_ocrs_as_a_footing_balance_sheet():
    t = X.extract_region_ocr(SCANNED_PDF, 0, (0, 0, 612, 792))
    assert t is not None
    assert t["kind"] == "statement of financial position"
    assert t["years"] == [2024, 2023]
    assert t["foots"] is True
    assert t["_ocr"] is True


@pytest.mark.skipif(not X.HAVE_OCR, reason="Tesseract binary not installed on this machine")
def test_scanned_fixture_ocr_figures_match_the_source_exactly():
    # OCR digit-misreads are a known, expected risk (see the "verify this by
    # eye" _ocr flag) -- this asserts the CURRENT read is exact so a future
    # regression (a rendering change, a Tesseract upgrade) that degrades
    # accuracy shows up here instead of only being noticed live.
    t = X.extract_region_ocr(SCANNED_PDF, 0, (0, 0, 612, 792))
    rows = {X._row_label(r): r for r in t["rows"]}
    assert rows["TOTAL ASSETS"][1:] == [126500, 113400]
    assert rows["Total liabilities"][1:] == [53200, 48600]
    assert rows["Total equity"][1:] == [73300, 64800]
    assert rows["TOTAL LIABILITIES AND EQUITY"][1:] == [126500, 113400]
