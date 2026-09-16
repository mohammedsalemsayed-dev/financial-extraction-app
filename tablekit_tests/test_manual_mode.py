"""
Tests for the Tabula-style manual box-selection feature added this session:
`tablekit/img2table_backend.py`'s region-merging/selection helpers, plus
`extract_all_tables.extract_region` and serve.py's manual-mode endpoints
(upload / search / page geometry / extraction).

The `_Region`-level tests are pure logic -- synthetic regions, no PDF, no
img2table/opencv/pandas needed -- so they run in CI same as test_units.py.
Everything that needs a real PDF (extract_region, search text, the HTTP
round trip) skips cleanly when the sample report isn't present, same
pattern as test_serve.py's test_http_end_to_end.
"""
import io
import json
import sys
import threading
import urllib.request
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
import extract_all_tables as X       # noqa: E402
import serve                          # noqa: E402
from tablekit.img2table_backend import (   # noqa: E402
    _Region, _RowRow, _split_glued_cell, _normalize_row_width, _merge_stacked,
    _drop_prose_columns, rows_in_box, HAVE_IMG2TABLE,
    _unscramble_reversed_paren_wrap, _merge_side_by_side,
    _looks_like_two_fused_statements, _SOCE_ANCHOR_RE, _split_glued_row,
    _redistribute_glued_cells, img2table_page_tables, grid_in_box,
)


def _region(x0, y0, x1, y1, rows):
    """rows: [(top, bot, [values]), ...]"""
    return _Region(x0, y0, x1, y1, rows)


# --------------------------------------------------------- glued cells -----
def test_split_glued_cell_splits_two_numbers():
    assert _split_glued_cell("3,116,600 10,148,291") == ["3,116,600", "10,148,291"]


def test_split_glued_cell_splits_number_and_nil_marker():
    assert _split_glued_cell("(17,142) –") == ["(17,142)", "–"]


def test_split_glued_cell_leaves_plain_values_alone():
    assert _split_glued_cell("1,234") == ["1,234"]
    assert _split_glued_cell(None) == [None]
    assert _split_glued_cell("Trade payables and accrual") == ["Trade payables and accrual"]


def test_normalize_row_width_pads_and_splits():
    # pads when there's nothing splittable
    assert _normalize_row_width(["label", None, None], 4) == ["label", None, None, None]
    # splits the rightmost glued cell to reach the target width
    got = _normalize_row_width(["label", "–", "–", "17,142", "(17,142) –"], 6)
    assert got == ["label", "–", "–", "17,142", "(17,142)", "–"]


# ---------------------------------------------------- whole-row splitting --
def test_split_glued_row_pulls_trailing_figures_off_a_label():
    """Regression: img2table couldn't find column boundaries at all for a
    row on du annual 2020's balance sheet and returned it as one wide
    cell -- "Property, plant and equipment 6 8,063,422 7,741,119" -- for
    a 4-column table (label, note, 2020, 2019). The note ref and both
    figures must come back as their own cells."""
    vals = ["Property, plant and equipment 6 8,063,422 7,741,119", None, None, None]
    assert _split_glued_row(vals, 4) == \
        ["Property, plant and equipment", "6", "8,063,422", "7,741,119"]


def test_split_glued_row_handles_a_subtotal_with_no_note_ref():
    """A "Total ..." row legitimately has fewer real values than the
    table's own widest row (no note-ref number of its own) -- splitting
    must not demand exactly (ncols - 1) trailing tokens."""
    vals = ["Total non-current assets 11,224,448 10,988,526", None, None, None]
    assert _split_glued_row(vals, 4) == \
        ["Total non-current assets", "11,224,448", "10,988,526"]


def test_split_glued_row_leaves_a_mid_label_wrap_alone():
    """The numbers sit in the MIDDLE of this one (a wrapped label split
    across the figures), not at the very end -- there's nothing to
    greedily pull off the tail, so the row is correctly left as-is
    rather than mis-split."""
    vals = ["Financial asset at fair value through other 11 18,368 18,368 "
            "comprehensive income", None, None, None]
    assert _split_glued_row(vals, 4) == vals


def test_split_glued_row_leaves_an_already_populated_row_alone():
    vals = ["label", "note", "1,234", "5,678"]
    assert _split_glued_row(vals, 4) == vals


def test_split_glued_row_leaves_a_short_plain_label_alone():
    # fewer than 3 tokens -- nothing to split off regardless of content
    assert _split_glued_row(["Cash and bank balances", None, None], 3) == \
        ["Cash and bank balances", None, None]


@pytest.mark.skipif(not HAVE_IMG2TABLE, reason="img2table not installed")
@pytest.mark.skipif(not (ROOT / "du annual 2020.pdf").exists(), reason="sample PDF not present")
def test_extract_region_balance_sheet_survives_column_boundary_failure():
    """End-to-end regression for THREE img2table_backend/telecom_extract
    bugs found and fixed live, all on the same table: du 2020's balance
    sheet came back with every value duplicated across all 4 columns
    ("As at 31 December" x4); once that collapse was fixed, as one glued
    cell per row ("Property, plant and equipment 6 8,063,422 7,741,119")
    instead of real cells; once THAT was split, the Notes column (6, 7,
    8, ...) still didn't get pulled out of `rows` because the table's
    two header rows ("Notes" / "AED 000") diluted the ratio just under
    the threshold. All three together made this whole statement either
    unusable or (worst case) silently wrong."""
    t = X.extract_region(ROOT / "du annual 2020.pdf", 69,
                          (0.0, 100.45420999999988, 615.275025, 841.89001))
    assert t is not None
    ppe = next(r for r in t["rows"] if r[0] == "Property, plant and equipment")
    assert ppe == ["Property, plant and equipment", 8063422, 7741119]
    assert X.row_note_ref(t, ppe) == "6"
    total = next(r for r in t["rows"] if r[0] == "Total non-current assets")
    assert total[:3] == ["Total non-current assets", 11224448, 10988526]


@pytest.mark.skipif(not HAVE_IMG2TABLE, reason="img2table not installed")
@pytest.mark.skipif(not (ROOT / "du annual 2020.pdf").exists(), reason="sample PDF not present")
def test_img2table_page_tables_rows_carry_plausible_raw_cell_geometry():
    """Phase A of the grid-preview feature: every row img2table detects on a
    real page should carry raw_cells (see _RowRow) with plausible per-cell
    x-ranges -- each cell's own x1 < x2, and every cell's x-range falls
    within the region's own overall x0/x1 bounds. This is the concrete,
    real-file proof that the capture point (after the duplicate-collapse,
    using the original cells' own bbox) produces sane geometry, not just
    that it doesn't crash on synthetic data."""
    regions = img2table_page_tables(ROOT / "du annual 2020.pdf", 69)
    real = [r for r in regions if r.x0 < 100]
    assert real, "expected to find the balance sheet's own region"
    region = real[0]
    checked = 0
    for row in region.row_rows:
        if row.raw_cells is None:
            continue
        for x1, x2, _ in row.raw_cells:
            assert x1 < x2
            assert region.x0 - 1 <= x1 and x2 <= region.x1 + 1
            checked += 1
    assert checked > 20   # this table has far more than 20 real cells total


@pytest.mark.skipif(not HAVE_IMG2TABLE, reason="img2table not installed")
@pytest.mark.skipif(not (ROOT / "du annual 2025.pdf").exists(), reason="sample PDF not present")
def test_extract_region_refuses_rather_than_show_rows_merged_by_img2table_itself():
    """du annual 2025.pdf restructured the statement of financial position
    onto a landscape page with two half-page statements side by side (left:
    non-current assets + equity; right: current assets + liabilities).
    img2table's own row-boundary detection fails outright on this layout --
    its one detected region spans the seam between the two halves, with
    cells collapsing ten different line items' figures into one, newline-
    joined ("10,288,767\\n1,515,395\\n869,600\\n..."). A real user drawing
    one box around the whole visible statement (the natural thing to do --
    the table is plainly one statement, not two) must get a clean refusal,
    not a confident-looking result missing most of the statement's rows
    and years. This is the same _looks_garbled check extract_region already
    ran against the OLDER pdfplumber-fallback path, now also checked
    against img2table's own output -- see _looks_garbled's docstring."""
    t = X.extract_region(ROOT / "du annual 2025.pdf", 130,
                          (0.0, 44.57799, 841.89001, 473.38599))
    assert t is None


@pytest.mark.skipif(not (ROOT / "en-2021-etisalat-group-annual-report.pdf").exists(),
                    reason="sample PDF not present")
def test_extract_region_recovers_labels_on_the_pdfplumber_fallback_path_too():
    """en-2021-etisalat-group-annual-report.pdf p61 is a 2-up landscape
    spread: the LEFT half is a normal P&L that img2table detects fine, but
    the RIGHT half (a separate, legitimate 'statement of profit or loss and
    other comprehensive income') is missed by img2table's table detection
    entirely, so extraction falls through to the plain pdfplumber path.
    That path used to call bare .extract_table(), which returns only cell
    text with no positional metadata -- so row_bands/matched_bbox stayed
    None and _attach_left_labels (the mechanism that recovers a label
    column sitting just outside a ruled numeric box) never even ran,
    leaving every row on this path numbers-only with no label at all.
    extract_region now uses find_table()+.extract() instead (the exact
    same result -- see pdfplumber's own Page.extract_table -- but the
    Table object also exposes each row's own bbox), so label recovery
    works here too."""
    t = X.extract_region(ROOT / "en-2021-etisalat-group-annual-report.pdf", 60,
                          (575.275025, 61.88810999999998, 1190.55005, 841.89001))
    assert t is not None
    profit = next(r for r in t["rows"] if r[0] == "Profit for the year")
    assert profit == ["Profit for the year", 11059489]
    total = next(r for r in t["rows"] if r[0] == "Total comprehensive income for the year")
    assert total == ["Total comprehensive income for the year", 10615297]


# ------------------------------------------------------------- merging -----
def test_merge_stacked_widens_a_narrower_glued_piece():
    """Regression: du 2025's SOCE closing-balance row split img2table into a
    6-col block and a 5-col block (last two columns glued in one cell) --
    merging must widen the narrower piece, not concatenate it as-is."""
    wide = _region(42, 60, 415, 300, [(60, 70, ["At 1 Jan", 1, 2, 3, 4, 5])])
    # 5 elements (col_tol=1 vs wide's 6): label + 3 loose values + 1 cell
    # gluing the last TWO columns together, same shape as the real bug
    narrow = _region(42, 315, 415, 365,
                     [(315, 325, ["At 31 Dec", 6, 7, 8, "3,116,600 10,148,291"])])
    merged = _merge_stacked([wide, narrow])
    assert len(merged) == 1
    rows = [list(v) for _, _, v in merged[0].row_rows]
    assert rows[-1] == ["At 31 Dec", 6, 7, 8, "3,116,600", "10,148,291"]


def test_merge_stacked_requires_x_alignment():
    """Regression: two Y-adjacent, same-column-count regions that sit in
    DIFFERENT horizontal columns (e.g. two unrelated notes on a dense page)
    must NOT be glued together just because they're stacked and same width."""
    left_col = _region(42, 60, 410, 160, [(60, 70, ["note A", 1, 2])])
    right_col = _region(426, 180, 793, 280, [(180, 190, ["note B", 3, 4])])
    merged = _merge_stacked([left_col, right_col])
    assert len(merged) == 2   # NOT merged -- different columns


def test_merge_stacked_still_merges_aligned_pieces():
    a = _region(42, 60, 410, 160, [(60, 70, ["row1", 1, 2])])
    b = _region(42, 180, 410, 280, [(180, 190, ["row2", 3, 4])])
    merged = _merge_stacked([a, b])
    assert len(merged) == 1
    assert merged[0].nrows == 2


def test_merge_stacked_carries_raw_cells_through():
    """A row's per-cell geometry (see _RowRow) must survive _merge_stacked --
    both the untouched first-in-run branch (plain list copy) and the
    width-normalized branch (a row from the narrower piece, textually
    widened) need to still expose the SAME raw_cells the row started with,
    since widening VALUES doesn't invent any real sub-cell x-position to
    reassign it to."""
    a = _region(42, 60, 410, 160, [
        _RowRow(60, 70, ["row1", 1, 2], raw_cells=[(42, 200, "row1"), (200, 300, 1), (300, 410, 2)]),
    ])
    b = _region(42, 180, 410, 280, [
        _RowRow(180, 190, ["row2", 3, 4], raw_cells=[(42, 200, "row2"), (200, 300, 3), (300, 410, 4)]),
    ])
    merged = _merge_stacked([a, b])
    assert len(merged) == 1
    raw = [row.raw_cells for row in merged[0].row_rows]
    assert raw == [
        [(42, 200, "row1"), (200, 300, 1), (300, 410, 2)],
        [(42, 200, "row2"), (200, 300, 3), (300, 410, 4)],
    ]


def test_drop_prose_columns_removes_a_prose_column_but_keeps_labels():
    rows = [
        ["Trade payables", "This is a long unrelated sentence about auditors", 100],
        ["Other payables", "Another long sentence describing something else here", 200],
    ]
    out = _drop_prose_columns(rows)
    assert out == [["Trade payables", 100], ["Other payables", 200]]


def test_drop_prose_columns_keeps_prose_if_it_is_the_only_text_column():
    rows = [["This is a long sentence acting as the only label column here", 100]]
    assert _drop_prose_columns(rows) == rows


def test_drop_prose_columns_safeguard_checks_for_a_real_label_not_any_string():
    """Found live on en-2022-1-eand-group-annual-report.pdf p50 (a cash flow
    statement): more than half of its genuine, wrapped-across-several-lines
    labels are long enough to trip _looks_like_prose's word-count rule
    ("Operating cash flows before changes in working capital" is 8 words),
    so the whole label column crossed the ">=50% prose" threshold for
    dropping. The "unless it's the only text column left" safeguard is
    supposed to catch exactly this, but its old implementation counted ANY
    non-empty string as a "text column" -- and a still-unparsed figure cell
    ("12,951,414") is a Python str too at this point in the pipeline, so the
    two purely-numeric columns sitting right next to the label column kept
    the safeguard from ever tripping, and the label column was dropped
    anyway, leaving every row with no description at all."""
    rows = [
        ["Operating cash flows before changes in working capital", "19,954,241", "19,975,569"],
        ["Cash flows from financing activities", None, None],
        ["Net cash generated from operating activities", "19,134,501", "18,110,856"],
    ]
    out = _drop_prose_columns(rows)
    assert out == rows   # the label column must survive -- it's the only real label column


def test_drop_prose_columns_safeguard_is_not_fooled_by_a_lone_header_cell():
    """A step further than the eand p50 case above: even checking for
    ALPHABETIC content isn't enough, because a purely-numeric column's own
    HEADER cell often carries letters too ("Notes", "AED'000 2019"). Found
    live on etisalat-group-annual-report-english-2019.pdf p46's right-half
    statement: img2table correctly captured every label in column 0, but
    over half of them are long/compound enough to trip the word-count
    prose rule, so column 0 got marked for dropping -- and a bare "any
    column has SOME alphabetic cell" safeguard saw "Notes" and "AED'000
    2019" in the OTHER columns' own header rows and judged them adequate
    replacement label columns, letting the real one go. The safeguard has
    to require alphabetic content on a healthy FRACTION of rows (the same
    threshold _has_label_column already uses), not just one cell."""
    rows = [
        [None, "Notes", "AED’000 2019", "AED’000 2018"],
        ["Profit for the year", None, "9,494,578", "10,441,920"],
        ["Other comprehensive income / (loss)", None, None, None],
        ["Items that are or may be reclassified subsequently to profit or loss:",
         None, None, None],
        ["Remeasurement of defined benefit obligations - net of tax", None,
         "37,008", "(62,667)"],
        ["Exchange differences on translation of foreign operations", None,
         "(470,726)", "(2,922,465)"],
        ["Gain on net investment hedges", "29,35", "56,416", "290,229"],
    ]
    out = _drop_prose_columns(rows)
    assert out == rows   # column 0 must survive -- "Notes"/"AED'000" alone don't count


# --------------------------------------------------------- rows_in_box -----
def test_rows_in_box_excludes_a_barely_grazed_region():
    """A region only clipped at the very edge of the drawn box (well under
    50% of its own area inside) must not qualify -- this is what stopped a
    generously-drawn box from dragging in a neighbouring note."""
    target = _region(426, 72, 793, 163,
                     [(72, 82, ["22 Trade and other payables", None, None, None]),
                      (88, 98, ["Trade payables", None, 2435421, 2135354])])
    neighbour = _region(42, 145, 412, 240,   # mostly outside the drawn box below
                        [(145, 157, ["unrelated note text", None, None])])
    page_tables = [target, neighbour]
    hit = rows_in_box(page_tables, x0=420, y0=65, x1=800, y1=170)
    assert hit is not None
    rows, bbox, row_bands = hit
    assert any("Trade payables" in (r[0] or "") for r in rows)
    assert not any("unrelated note" in (r[0] or "") for r in rows)


def test_rows_in_box_trims_to_the_drawn_range_not_the_whole_region():
    """A region that legitimately extends past the drawn box (e.g. the start
    of the NEXT section, pulled in by a stacking merge) must be trimmed --
    the returned bbox reflects only the rows that survive, not the raw
    region's full extent."""
    combined = _region(426, 72, 793, 288, [
        (72, 82, ["22 Trade and other payables", None, None]),
        (88, 98, ["Trade payables", 2435421, 2135354]),
        (227, 238, ["23 Federal royalty and corporate income tax", None, None]),
        (244, 253, ["Federal royalty on profit", 1675882, None]),
    ])
    hit = rows_in_box([combined], x0=420, y0=65, x1=800, y1=200)
    assert hit is not None
    rows, bbox, row_bands = hit
    labels = [r[0] for r in rows]
    assert any("Trade payables" in (l or "") for l in labels)
    assert not any("23 Federal" in (l or "") for l in labels), (
        "the next section's heading leaked past the drawn box's Y-range")
    assert bbox[3] < 200, f"bbox should be trimmed tight, got {bbox}"


def test_rows_in_box_returns_none_for_an_empty_page():
    assert rows_in_box([], 0, 0, 100, 100) is None


def test_rows_in_box_accepts_a_box_drawn_around_only_part_of_a_bigger_region():
    """The real bug this locks in: a user draws a box around only the
    "Assets" half of a balance-sheet region, deliberately leaving
    "Liabilities" out. That box covers well under 50% of the REGION's own
    area (checking region-coverage alone -- the original, buggy behaviour --
    would reject it) but effectively 100% of the DRAWN BOX's own area, since
    the box sits entirely inside the region. `_overlap_frac` must take
    whichever of the two fractions is more generous, not region-coverage
    alone, or a deliberately partial selection like this returns nothing."""
    region = _region(40, 0, 400, 400, [
        (10, 30, ["Assets", None, None]),
        (30, 60, ["Cash and equivalents", 500_000, 420_000]),
        (60, 90, ["Trade receivables", 300_000, 250_000]),
        (250, 280, ["Liabilities", None, None]),
        (280, 310, ["Trade payables", 200_000, 180_000]),
        (310, 340, ["Borrowings", 100_000, 90_000]),
    ])
    # drawn box: y=0..150, i.e. the Assets section only -- 150/400 = 37.5%
    # of the region's area, well under the 50% default min_overlap_frac
    hit = rows_in_box([region], x0=40, y0=0, x1=400, y1=150)
    assert hit is not None, "a box covering <50% of the region but ~100% of itself must still match"
    rows, bbox, row_bands = hit
    labels = [r[0] for r in rows]
    assert any("Cash and equivalents" in (l or "") for l in labels)
    assert not any("Liabilities" in (l or "") for l in labels), (
        "the deliberately-excluded Liabilities section leaked into a partial selection")
    assert not any("Borrowings" in (l or "") for l in labels)


# ---------------------------------------------------------- grid_in_box ----
def test_grid_in_box_returns_none_for_an_empty_page():
    assert grid_in_box([], 0, 0, 100, 100) is None


def test_grid_in_box_returns_row_bands_matching_rows_in_box():
    """row_bands is exactly what rows_in_box already computes -- grid_in_box
    just also returns it, using the same region-selection (_best_matching_
    region) so the two never disagree about which region matched."""
    region = _region(40, 0, 400, 100, [
        _RowRow(10, 30, ["Assets", None, None], raw_cells=[(40, 200, "Assets"), (200, 300, None), (300, 400, None)]),
        _RowRow(30, 60, ["Cash", 500_000, 420_000],
               raw_cells=[(40, 200, "Cash"), (200, 300, 500_000), (300, 400, 420_000)]),
    ])
    rows_hit = rows_in_box([region], x0=40, y0=0, x1=400, y1=100)
    grid_hit = grid_in_box([region], x0=40, y0=0, x1=400, y1=100)
    assert rows_hit is not None and grid_hit is not None
    _, _, row_bands_from_rows = rows_hit
    row_bands_from_grid, _, _ = grid_hit
    assert row_bands_from_grid == row_bands_from_rows


def test_grid_in_box_clusters_shared_edges_into_column_bands():
    """Adjacent real cells share (almost) the same edge value -- one cell's
    x2 sits right next to the next cell's x1 -- so clustering every kept
    row's raw_cells edges together should collapse that shared boundary
    into ONE band edge, not two, giving exactly ncols+1 boundaries for an
    n-column table."""
    region = _region(0, 0, 300, 50, [
        _RowRow(0, 10, ["Label", 1, 2],
               raw_cells=[(0, 100, "Label"), (100, 200, 1), (200, 300, 2)]),
        _RowRow(10, 20, ["Row2", 3, 4],
               raw_cells=[(0, 100, "Row2"), (100, 200, 3), (200, 300, 4)]),
    ])
    hit = grid_in_box([region], x0=0, y0=0, x1=300, y1=20)
    assert hit is not None
    _, col_bands, _ = hit
    assert col_bands == [(0, 100), (100, 200), (200, 300)]


def test_grid_in_box_returns_empty_col_bands_when_no_raw_cells_available():
    """A graceful partial result, not a failure: a region matched via a
    path with no per-cell geometry (a synthetic region built the old way,
    or img2table_page_tables never captured any) still returns real row
    bands, just an empty column list."""
    region = _region(0, 0, 300, 50, [(0, 10, ["Label", 1, 2])])   # plain tuple, no raw_cells
    hit = grid_in_box([region], x0=0, y0=0, x1=300, y1=20)
    assert hit is not None
    row_bands, col_bands, _ = hit
    assert row_bands == [(0, 10)]
    assert col_bands == []


@pytest.mark.skipif(not HAVE_IMG2TABLE, reason="img2table not installed")
@pytest.mark.skipif(not (ROOT / "du annual 2020.pdf").exists(), reason="sample PDF not present")
def test_grid_in_box_narrowing_the_box_narrows_the_detected_columns():
    """The concrete, real-file demonstration of the actual gap this feature
    closes: today, narrowing a drawn box's WIDTH has zero effect on which
    columns come back from extraction (confirmed live, in conversation,
    before this feature existed -- a box stopping well short of a table's
    second year column still returned both years). grid_in_box must behave
    differently: a box narrowed to roughly the first two columns of du
    2020's balance sheet (Assets label + Notes) must report FEWER column
    bands than the same rows under the table's full width (Assets + Notes
    + 2020 + 2019) -- the first test in this repo to exercise column-
    narrowing at all."""
    page_tables = img2table_page_tables(ROOT / "du annual 2020.pdf", 69)
    full = grid_in_box(page_tables, 0.0, 100.45420999999988, 615.275025, 400.0)
    narrow = grid_in_box(page_tables, 0.0, 100.45420999999988, 380.0, 400.0)
    assert full is not None and narrow is not None
    _, full_cols, _ = full
    _, narrow_cols, _ = narrow
    assert len(narrow_cols) < len(full_cols)


# ------------------------------------------------- extract_region (PDF) ----
def _sample_pdf():
    for name in ("du annual 2025.pdf", "du annual 2020.pdf", "du annual 2016.pdf"):
        p = ROOT / name
        if p.exists():
            return p
    return None


@pytest.mark.skipif(not HAVE_IMG2TABLE, reason="img2table not installed")
@pytest.mark.skipif(not (ROOT / "du annual 2025.pdf").exists(), reason="sample PDF not present")
def test_extract_region_soce_full_extraction():
    """End-to-end regression for the two bugs found and fixed live in this
    session: the glued closing-balance cell, and the box-vs-content bbox
    mismatch."""
    t = X.extract_region(ROOT / "du annual 2025.pdf", 133, (42, 60, 415, 365))
    assert t is not None
    assert t["kind"] == "statement of changes in equity"
    assert t["foots"] is True
    last = t["rows"][-1]
    assert last[0] == "At 31 December 2025"
    assert last[-2:] == [3116600, 10148291]   # was glued into one string pre-fix


@pytest.mark.skipif(not HAVE_IMG2TABLE, reason="img2table not installed")
@pytest.mark.skipif(not (ROOT / "du annual 2010.pdf").exists(), reason="sample PDF not present")
def test_extract_region_recovers_the_notes_reference_column():
    """Regression for a bug reported live: manually boxing du 2010's
    landscape 2-up "Consolidated statement of comprehensive income" (page
    16, left half) either dropped the Notes column entirely or glued its
    numbers onto the label ("General and administrative expenses 19").
    extract_region now runs the same _strip_note_refs pass _recon_tables
    uses, pulling the ref out into a display-only side channel
    (t["note_ref_map"] / row_note_ref) instead -- labels stay clean and
    `rows` stays note-free (so footing / value-column detection can't be
    skewed by it)."""
    t = X.extract_region(ROOT / "du annual 2010.pdf", 15,
                          (0.0, 33.836, 615.275, 308.918))
    assert t is not None
    ga = next(r for r in t["rows"] if X._row_label(r) ==
              "General and administrative expenses")
    assert ga[1:] == [-3306525, -3044490]          # not glued into the label
    assert X.row_note_ref(t, ga) == "19"
    # not the SAME field analyze() uses for its own, unrelated "which
    # column looks like note refs" column-index hint -- a collision here
    # previously meant analyze() silently clobbered the map after it was
    # set (see row_note_ref's docstring).
    assert t.get("note_ref_map")
    assert not isinstance(t.get("note_col"), dict)


@pytest.mark.skipif(not HAVE_IMG2TABLE, reason="img2table not installed")
@pytest.mark.skipif(not (ROOT / "du annual 2025.pdf").exists(), reason="sample PDF not present")
def test_extract_region_returns_none_for_a_blank_area():
    t = X.extract_region(ROOT / "du annual 2025.pdf", 0, (10, 10, 40, 40))
    assert t is None


# --------------------------------------------------------------- OCR failsafe --
# The OCR path (tablekit.img2table_backend.ocr_rows_in_box /
# extract_all_tables.extract_region_ocr) needs the actual Tesseract binary
# installed, not just pip installs -- these tests cover the plumbing that
# does NOT need it: routing (never runs unless explicitly asked, never runs
# when HAVE_OCR is off) and the text-detection check that decides whether the
# UI even offers the OCR button. Full OCR correctness needs a machine with
# Tesseract installed and is out of scope for CI.
def test_ocr_rows_in_box_pad_default_is_20_not_6():
    """Regression lock, not a live OCR check (that needs the real Tesseract
    binary + visual inspection, done by hand this session -- see
    CHANGELOG.md 0.5.0). At pad=6 a tightly-drawn box clipped/misread
    trailing digits ("$20,565,087" -> "$20,565,C" or worse); pad=20 fixed
    4 of 5 tightness variants tested live. This just makes sure no future
    edit quietly reverts the default back down."""
    import inspect
    from tablekit import img2table_backend as _i2t
    sig = inspect.signature(_i2t.ocr_rows_in_box)
    assert sig.parameters["pad"].default == 20


@pytest.mark.skipif(not (ROOT / "du annual 2025.pdf").exists(), reason="sample PDF not present")
def test_box_has_text_is_true_over_a_real_statement():
    # same box as test_extract_region_soce_full_extraction -- known to have text
    assert X.box_has_text(ROOT / "du annual 2025.pdf", 133, (42, 60, 415, 365)) is True


@pytest.mark.skipif(not (ROOT / "du annual 2025.pdf").exists(), reason="sample PDF not present")
def test_box_has_text_is_false_over_a_blank_area():
    # same blank box as test_extract_region_returns_none_for_a_blank_area
    assert X.box_has_text(ROOT / "du annual 2025.pdf", 0, (10, 10, 40, 40)) is False


def test_extract_region_ocr_returns_none_without_an_ocr_engine(monkeypatch):
    monkeypatch.setattr(X, "HAVE_OCR", False)
    assert X.extract_region_ocr(ROOT / "does-not-matter.pdf", 0, (0, 0, 10, 10)) is None


def test_ocr_rows_in_box_returns_none_without_an_ocr_engine(monkeypatch):
    from tablekit import img2table_backend as _i2t
    monkeypatch.setattr(_i2t, "HAVE_OCR", False)
    assert _i2t.ocr_rows_in_box("does-not-matter.pdf", 0, 0, 0, 10, 10) is None


def test_region_has_text_wraps_extract_all_tables(tmp_path, monkeypatch):
    """serve.py's endpoint-level wrapper resolves the file name to a path
    and delegates to X.box_has_text -- this is what /api/extract_region's
    422 response uses to decide whether to offer the OCR button."""
    pdf = tmp_path / "demo.pdf"
    pdf.write_bytes(b"%PDF-1.4 fake")
    monkeypatch.setattr(serve, "_state",
                        {"files": [pdf], "pages": None, "scans": {}, "manual": {},
                         "pngs": {}, "pagetext": {}})
    seen = {}
    def fake_has_text(path, page_index0, bbox):
        seen["args"] = (path, page_index0, bbox)
        return False
    monkeypatch.setattr(X, "box_has_text", fake_has_text)
    assert serve.region_has_text("demo.pdf", 3, [1, 2, 3, 4]) is False
    assert seen["args"] == (pdf, 2, (1, 2, 3, 4))
    with pytest.raises(KeyError):
        serve.region_has_text("not-a-file.pdf", 1, [0, 0, 1, 1])


def test_extract_region_ocr_endpoint_appends_to_manual_list_and_flags_ocr(tmp_path, monkeypatch):
    pdf = tmp_path / "demo.pdf"
    pdf.write_bytes(b"%PDF-1.4 fake")
    monkeypatch.setattr(serve, "_state",
                        {"files": [pdf], "pages": None, "scans": {}, "manual": {},
                         "pngs": {}, "pagetext": {}})
    fake_t = {"rows": [["Revenue", 100]], "title": "OCR region", "kind": "table",
             "page_label": 1, "shape": "table", "_manual": True, "_ocr": True}
    monkeypatch.setattr(X, "extract_region_ocr", lambda *a, **k: fake_t)
    d = serve.extract_region_ocr("demo.pdf", 1, [0, 0, 10, 10])
    assert d is not None
    assert d["ocr"] is True
    assert d["n"] == serve.MANUAL_OFFSET + 1
    assert serve._manual_list("demo.pdf") == [fake_t]


# --------------------------------------------------- serve.py manual mode --
def test_resolve_addresses_manual_and_auto_tables_separately(tmp_path, monkeypatch):
    pdf = tmp_path / "demo.pdf"
    pdf.write_bytes(b"%PDF-1.4 fake")
    monkeypatch.setattr(serve, "_state",
                        {"files": [pdf], "pages": None, "scans": {}, "manual": {},
                         "pngs": {}, "pagetext": {}})
    manual_t = {"rows": [["x"]], "title": "manual one", "kind": "table",
               "page_label": 1, "shape": "table", "_manual": True}
    serve._manual_list("demo.pdf").append(manual_t)
    got = serve._resolve("demo.pdf", serve.MANUAL_OFFSET + 1)
    assert got is manual_t
    with pytest.raises(KeyError):
        serve._resolve("demo.pdf", serve.MANUAL_OFFSET + 2)   # no second manual table


def test_upload_pdf_saves_file_and_registers_it(tmp_path, monkeypatch):
    monkeypatch.setattr(serve, "_state",
                        {"files": [], "pages": None, "scans": {}, "manual": {},
                         "pngs": {}, "pagetext": {}})
    monkeypatch.setattr(serve, "UPLOAD_DIR", tmp_path / "uploads")
    import base64
    b64 = base64.b64encode(b"%PDF-1.4 not a real pdf but bytes are bytes").decode()
    name = serve.upload_pdf("my report.pdf", b64)
    assert name == "my report.pdf"
    assert (tmp_path / "uploads" / "my report.pdf").exists()
    assert any(p.name == name for p in serve._state["files"])


def test_upload_pdf_sanitises_unsafe_filenames(tmp_path, monkeypatch):
    monkeypatch.setattr(serve, "_state",
                        {"files": [], "pages": None, "scans": {}, "manual": {},
                         "pngs": {}, "pagetext": {}})
    monkeypatch.setattr(serve, "UPLOAD_DIR", tmp_path / "uploads")
    import base64
    b64 = base64.b64encode(b"data").decode()
    name = serve.upload_pdf("../../evil<>:.pdf", b64)
    assert ".." not in name and "/" not in name and "\\" not in name
    assert (tmp_path / "uploads" / name).exists()


@pytest.mark.skipif(not (ROOT / "du annual 2025.pdf").exists(), reason="sample PDF not present")
def test_search_pdf_finds_and_caches_hits(monkeypatch):
    monkeypatch.setattr(serve, "_state",
                        {"files": [ROOT / "du annual 2025.pdf"], "pages": None, "scans": {},
                         "manual": {}, "pngs": {}, "pagetext": {}})
    hits = serve.search_pdf("du annual 2025.pdf", "changes in equity")
    assert hits
    assert any(h["page"] == 134 for h in hits)
    assert "du annual 2025.pdf" in serve._state["pagetext"]   # page text got cached
    cached = serve._state["pagetext"]["du annual 2025.pdf"]
    assert serve._page_texts("du annual 2025.pdf") is cached  # second call reuses the cache


def test_search_pdf_short_query_returns_nothing(monkeypatch):
    monkeypatch.setattr(serve, "_state",
                        {"files": [], "pages": None, "scans": {}, "manual": {},
                         "pngs": {}, "pagetext": {}})
    assert serve.search_pdf("anything.pdf", "a") == []


# ------------------------------------------------------- full HTTP round ---
@pytest.mark.skipif(not HAVE_IMG2TABLE, reason="img2table not installed")
@pytest.mark.skipif(not (ROOT / "du annual 2025.pdf").exists(), reason="sample PDF not present")
def test_http_manual_extraction_round_trip(monkeypatch, tmp_path):
    """Upload -> pagecount -> page_raw -> extract_region -> scan -> table ->
    export, through the real HTTP handler -- the same path the UI drives.
    Uses the SOCE table on page 134 (the one hand-verified live this
    session) rather than a guessed location, so a miss here is a real bug,
    not "nothing happened to be on that page." """
    pdf = ROOT / "du annual 2025.pdf"
    monkeypatch.setattr(serve, "_state",
                        {"files": [], "pages": None, "scans": {}, "manual": {},
                         "pngs": {}, "pagetext": {}})
    monkeypatch.setattr(serve, "UPLOAD_DIR", tmp_path / "uploads")
    from http.server import ThreadingHTTPServer
    port = serve._free_port(9200)
    srv = ThreadingHTTPServer(("127.0.0.1", port), serve.Handler)
    th = threading.Thread(target=srv.serve_forever, daemon=True)
    th.start()
    try:
        base = f"http://127.0.0.1:{port}"
        import base64
        b64 = base64.b64encode(pdf.read_bytes()).decode()
        up = json.load(urllib.request.urlopen(urllib.request.Request(
            base + "/api/upload",
            data=json.dumps({"filename": pdf.name, "data_b64": b64}).encode(),
            headers={"Content-Type": "application/json"}, method="POST")))
        name = up["file"]

        pc = json.load(urllib.request.urlopen(
            base + f"/api/pagecount?file={urllib.request.quote(name)}"))
        assert pc["pages"] > 0

        png = urllib.request.urlopen(
            base + f"/api/page_raw?file={urllib.request.quote(name)}&n=134&scale=1.5").read()
        assert png[:8] == b"\x89PNG\r\n\x1a\n"

        ext = json.load(urllib.request.urlopen(urllib.request.Request(
            base + "/api/extract_region",
            data=json.dumps({"file": name, "page": 134,
                            "bbox": [42, 60, 415, 365]}).encode(),
            headers={"Content-Type": "application/json"}, method="POST")))
        assert ext["kind"] == "statement of changes in equity"
        assert ext["foots"] is True
        n = ext["n"]
        assert n > serve.MANUAL_OFFSET

        inv = json.load(urllib.request.urlopen(
            base + f"/api/scan?file={urllib.request.quote(name)}"))
        assert any(t["n"] == n for t in inv["tables"])

        body = json.dumps({"file": name, "ns": [n]}).encode()
        req = urllib.request.Request(base + "/api/export", data=body,
                                     headers={"Content-Type": "application/json"})
        xlsx = urllib.request.urlopen(req)
        assert xlsx.status == 200
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(xlsx.read()))
        assert "Contents" in wb.sheetnames
    finally:
        srv.shutdown()


@pytest.mark.skipif(not HAVE_IMG2TABLE, reason="img2table not installed")
@pytest.mark.skipif(not (ROOT / "du annual 2020.pdf").exists(), reason="sample PDF not present")
def test_http_detect_grid_round_trip(monkeypatch, tmp_path):
    """Upload -> detect_grid, through the real HTTP handler -- the grid-
    preview feature's own round trip, separate from extract_region's.
    Confirms the endpoint is reachable, always returns 200 (never a 4xx for
    "nothing detected," which is the routine outcome of an exploratory
    drag -- see detect_grid's own docstring), and that a real, known table
    (du 2020's balance sheet) actually reports plausible row/column bands,
    not just an empty/available:false stub."""
    pdf = ROOT / "du annual 2020.pdf"
    monkeypatch.setattr(serve, "_state",
                        {"files": [], "pages": None, "scans": {}, "manual": {},
                         "pngs": {}, "pagetext": {}})
    monkeypatch.setattr(serve, "UPLOAD_DIR", tmp_path / "uploads")
    from http.server import ThreadingHTTPServer
    port = serve._free_port(9210)
    srv = ThreadingHTTPServer(("127.0.0.1", port), serve.Handler)
    th = threading.Thread(target=srv.serve_forever, daemon=True)
    th.start()
    try:
        base = f"http://127.0.0.1:{port}"
        import base64
        b64 = base64.b64encode(pdf.read_bytes()).decode()
        up = json.load(urllib.request.urlopen(urllib.request.Request(
            base + "/api/upload",
            data=json.dumps({"filename": pdf.name, "data_b64": b64}).encode(),
            headers={"Content-Type": "application/json"}, method="POST")))
        name = up["file"]

        req = urllib.request.Request(
            base + "/api/detect_grid",
            data=json.dumps({"file": name, "page": 70,
                            "bbox": [0.0, 100.45420999999988, 615.275025, 400.0]}).encode(),
            headers={"Content-Type": "application/json"}, method="POST")
        resp = urllib.request.urlopen(req)
        assert resp.status == 200   # never a 4xx for this endpoint
        d = json.load(resp)
        assert d["available"] is True
        assert len(d["rows"]) > 5
        assert len(d["cols"]) == 4

        # a request for a spot with nothing on it must still be a clean 200
        req2 = urllib.request.Request(
            base + "/api/detect_grid",
            data=json.dumps({"file": name, "page": 70, "bbox": [0.0, 0.0, 5.0, 5.0]}).encode(),
            headers={"Content-Type": "application/json"}, method="POST")
        resp2 = urllib.request.urlopen(req2)
        assert resp2.status == 200
        assert json.load(resp2)["available"] is False
    finally:
        srv.shutdown()


# ---------------------------------------------- Phase F: grid -> extraction


def test_reglue_bare_note_column_merges_an_isolated_notes_column_into_the_label():
    """A grid whose own col_bands split the Notes column out on its own
    (see _extract_via_grid's docstring) must get it merged back onto the
    label -- the same shape img2table's own column-folding produces on the
    other two paths -- rather than leaking through as an extra column."""
    raw = [
        ["Property, plant and equipment", "6", "8,063,422", "7,741,119"],
        ["Right-of-use assets", "7", "1,851,429", "1,699,651"],
        ["Total non-current assets", "", "11,224,448", "10,988,526"],
    ]
    assert X._reglue_bare_note_column(raw) == [
        ["Property, plant and equipment 6", "8,063,422", "7,741,119"],
        ["Right-of-use assets 7", "1,851,429", "1,699,651"],
        ["Total non-current assets", "11,224,448", "10,988,526"],
    ]


def test_reglue_bare_note_column_leaves_a_genuine_figure_column_alone():
    """A real (if small/thousands-formatted) figure column must never be
    mistaken for a notes column -- "18,368" is a comma-grouped THOUSANDS
    figure (a 2-digit then a 3-digit chunk), not a comma-separated list of
    independent 1-2 digit note refs like "6, 7"."""
    raw = [
        ["Interest income", "18,368", "18,368"],
        ["Contract assets", "211,216", "208,994"],
    ]
    assert X._reglue_bare_note_column(raw) == raw


def test_reglue_bare_note_column_requires_at_least_two_matching_values():
    """A single coincidental small number is too little evidence to call an
    entire column a notes column -- leave it alone rather than guess."""
    raw = [["Row A", "6", "100"], ["Row B", "", "200"]]
    assert X._reglue_bare_note_column(raw) == raw


def test_reglue_bare_note_column_tolerates_the_tables_own_header_row():
    """Found live on du annual 2020.pdf's balance sheet: this runs on
    UNFILTERED raw rows, still including the table's own column-header row
    ("Assets" | "Notes" | "2020 AED 000" | "2019 AED 000") -- the literal
    word "Notes" in that row's own note-column cell doesn't itself match
    the bare-ref pattern, and requiring unanimity let that one legitimate
    header row veto an otherwise-overwhelming notes column."""
    raw = [
        ["Assets", "Notes", "2020 AED 000", "2019 AED 000"],
        ["Property, plant and equipment", "6", "8,063,422", "7,741,119"],
        ["Right-of-use assets", "7", "1,851,429", "1,699,651"],
        ["Intangible assets and goodwill", "8", "900,215", "1,051,446"],
    ]
    out = X._reglue_bare_note_column(raw)
    assert out[1] == ["Property, plant and equipment 6", "8,063,422", "7,741,119"]
    assert out[2] == ["Right-of-use assets 7", "1,851,429", "1,699,651"]
    assert out[3] == ["Intangible assets and goodwill 8", "900,215", "1,051,446"]


@pytest.mark.skipif(not HAVE_IMG2TABLE, reason="img2table not installed")
@pytest.mark.skipif(not (ROOT / "du annual 2020.pdf").exists(), reason="sample PDF not present")
def test_extract_region_grid_none_reproduces_todays_result_unchanged():
    """Phase F's core promise: grid=None (the default -- an old client, or a
    box no detection ever ran on) must still take the exact same img2table
    path as before this feature existed. Same box, same real known-good
    figures as test_extract_region_balance_sheet_survives_column_boundary_failure."""
    t = X.extract_region(ROOT / "du annual 2020.pdf", 69,
                          (0.0, 100.45420999999988, 615.275025, 841.89001))
    assert t is not None
    ppe = next(r for r in t["rows"] if r[0] == "Property, plant and equipment")
    assert ppe == ["Property, plant and equipment", 8063422, 7741119]


@pytest.mark.skipif(not HAVE_IMG2TABLE, reason="img2table not installed")
@pytest.mark.skipif(not (ROOT / "du annual 2020.pdf").exists(), reason="sample PDF not present")
def test_extract_region_with_a_real_detected_grid_reads_the_same_correct_figures():
    """A grid built by detect_grid (exactly what the frontend sends after a
    user confirms the auto-detected preview, unedited) must read the SAME
    real, known-good figures as the ordinary img2table path -- proving
    _extract_via_grid's word-in-cell-rectangle reading is actually correct
    on a real file, not just that it runs without crashing."""
    pdf = ROOT / "du annual 2020.pdf"
    bbox = (0.0, 100.45420999999988, 615.275025, 841.89001)
    grid = X.detect_grid(pdf, 69, bbox)
    assert grid["available"] is True
    t = X.extract_region(pdf, 69, bbox, grid=grid)
    assert t is not None
    ppe = next(r for r in t["rows"] if r[0] == "Property, plant and equipment")
    assert ppe[-2:] == [8063422, 7741119]
    total = next(r for r in t["rows"] if r[0] == "Total non-current assets")
    assert total[-2:] == [11224448, 10988526]
    # found live (before _reglue_bare_note_column): this table's real
    # col_bands split the Notes column out on its own -- it reads back as
    # an already-isolated "6"/"7"/"8", never glued to anything, so
    # _strip_note_refs (which only knows how to UN-glue "label 6" in one
    # cell) never got the chance to remove it, leaking through as an
    # unwanted extra "value" column. Confirm it's gone from rows, same as
    # the ordinary img2table path already produces for this exact table.
    assert ppe == ["Property, plant and equipment", 8063422, 7741119]
    assert X.row_note_ref(t, ppe) == "6"


@pytest.mark.skipif(not HAVE_IMG2TABLE, reason="img2table not installed")
@pytest.mark.skipif(not (ROOT / "du annual 2020.pdf").exists(), reason="sample PDF not present")
def test_extract_region_with_a_user_narrowed_grid_drops_the_excluded_column():
    """The whole point of Phase F: a user who dragged/removed a column line
    to exclude the 2019 figures gets an extraction that actually reflects
    that edit -- not the full auto-detected width regardless of what they
    changed. Simulates the edit by dropping the grid's own rightmost column
    band, the same shape removeGridLine produces client-side."""
    pdf = ROOT / "du annual 2020.pdf"
    bbox = (0.0, 100.45420999999988, 615.275025, 841.89001)
    grid = X.detect_grid(pdf, 69, bbox)
    assert grid["available"] is True
    assert len(grid["cols"]) >= 2
    narrowed = dict(grid, cols=grid["cols"][:-1])   # drop the rightmost (2019) column
    t = X.extract_region(pdf, 69, bbox, grid=narrowed)
    assert t is not None
    ppe = next(r for r in t["rows"] if r[0] == "Property, plant and equipment")
    assert 8063422 in ppe
    assert 7741119 not in ppe


@pytest.mark.skipif(not HAVE_IMG2TABLE, reason="img2table not installed")
@pytest.mark.skipif(not (ROOT / "du annual 2020.pdf").exists(), reason="sample PDF not present")
def test_http_extract_region_with_grid_round_trip(monkeypatch, tmp_path):
    """Phase F's full HTTP wiring: detect_grid -> extract_region with that
    grid in the POST body, through the real HTTP handler -- confirms
    serve.py's payload.get("grid") passthrough actually reaches
    X.extract_region and produces a correct, real result end to end (edit-
    propagation itself is already covered by the pure-Python tests above)."""
    pdf = ROOT / "du annual 2020.pdf"
    monkeypatch.setattr(serve, "_state",
                        {"files": [], "pages": None, "scans": {}, "manual": {},
                         "pngs": {}, "pagetext": {}})
    monkeypatch.setattr(serve, "UPLOAD_DIR", tmp_path / "uploads")
    from http.server import ThreadingHTTPServer
    port = serve._free_port(9220)
    srv = ThreadingHTTPServer(("127.0.0.1", port), serve.Handler)
    th = threading.Thread(target=srv.serve_forever, daemon=True)
    th.start()
    try:
        base = f"http://127.0.0.1:{port}"
        import base64
        b64 = base64.b64encode(pdf.read_bytes()).decode()
        up = json.load(urllib.request.urlopen(urllib.request.Request(
            base + "/api/upload",
            data=json.dumps({"filename": pdf.name, "data_b64": b64}).encode(),
            headers={"Content-Type": "application/json"}, method="POST")))
        name = up["file"]

        bbox = [0.0, 100.45420999999988, 615.275025, 841.89001]
        gd = json.load(urllib.request.urlopen(urllib.request.Request(
            base + "/api/detect_grid",
            data=json.dumps({"file": name, "page": 70, "bbox": bbox}).encode(),
            headers={"Content-Type": "application/json"}, method="POST")))
        assert gd["available"] is True

        ext = json.load(urllib.request.urlopen(urllib.request.Request(
            base + "/api/extract_region",
            data=json.dumps({"file": name, "page": 70, "bbox": bbox,
                            "grid": gd}).encode(),
            headers={"Content-Type": "application/json"}, method="POST")))
        ppe = next(r for r in ext["rows"] if r[0] == "Property, plant and equipment")
        assert ppe[-2:] == [8063422, 7741119]
    finally:
        srv.shutdown()


# ------------------------------------------------- delete / undo (UX fix) --
def test_delete_manual_tombstones_without_shifting_other_indices(tmp_path, monkeypatch):
    pdf = tmp_path / "demo.pdf"
    pdf.write_bytes(b"%PDF-1.4 fake")
    monkeypatch.setattr(serve, "_state",
                        {"files": [pdf], "pages": None, "scans": {}, "manual": {},
                         "pngs": {}, "pagetext": {}})
    t1 = {"rows": [["a"]], "title": "one", "kind": "table", "page_label": 1,
         "shape": "table", "_manual": True}
    t2 = {"rows": [["b"]], "title": "two", "kind": "table", "page_label": 2,
         "shape": "table", "_manual": True}
    serve._manual_list("demo.pdf").append(t1)
    serve._manual_list("demo.pdf").append(t2)
    n1, n2 = serve.MANUAL_OFFSET + 1, serve.MANUAL_OFFSET + 2

    serve.delete_manual("demo.pdf", n1)

    with pytest.raises(KeyError):
        serve._resolve("demo.pdf", n1)
    assert serve._resolve("demo.pdf", n2) is t2   # n2 still resolves to t2, unshifted

    inv = serve.inventory("demo.pdf")
    ns = [row["n"] for row in inv["tables"]]
    assert n1 not in ns
    assert n2 in ns


def test_delete_manual_rejects_auto_table_numbers():
    with pytest.raises(KeyError):
        serve.delete_manual("demo.pdf", 1)   # 1 <= MANUAL_OFFSET -> not a manual table


def test_delete_manual_rejects_already_deleted(tmp_path, monkeypatch):
    pdf = tmp_path / "demo.pdf"
    pdf.write_bytes(b"%PDF-1.4 fake")
    monkeypatch.setattr(serve, "_state",
                        {"files": [pdf], "pages": None, "scans": {}, "manual": {},
                         "pngs": {}, "pagetext": {}})
    serve._manual_list("demo.pdf").append(
        {"rows": [["a"]], "title": "one", "kind": "table", "page_label": 1,
         "shape": "table", "_manual": True})
    n = serve.MANUAL_OFFSET + 1
    serve.delete_manual("demo.pdf", n)
    with pytest.raises(KeyError):
        serve.delete_manual("demo.pdf", n)


# ------------------------------------------------- quick-find (UX fix) -----
@pytest.mark.skipif(not (ROOT / "du annual 2025.pdf").exists(), reason="sample PDF not present")
def test_quick_find_statements_locates_the_core_statements(monkeypatch):
    monkeypatch.setattr(serve, "_state",
                        {"files": [ROOT / "du annual 2025.pdf"], "pages": None, "scans": {},
                         "manual": {}, "pngs": {}, "pagetext": {}})
    hits = serve.quick_find_statements("du annual 2025.pdf")
    labels = {h["label"]: h["page"] for h in hits}
    assert "Changes in equity" in labels
    # this is a plain text search, not page-structure understanding -- a
    # running header referencing "...changes in equity 133" on an EARLIER
    # notes page (p124, verified) legitimately matches before the real
    # statement (p134) does. That's an accepted characteristic of a fast
    # orientation aid, not a bug: still lands you in the right neighbourhood.
    assert 100 < labels["Changes in equity"] <= 134
    # results come back in the fixed canonical order, not page order
    assert [h["label"] for h in hits] == [l for l, _ in serve._QUICKFIND_PATTERNS
                                          if l in labels]


def test_quick_find_statements_empty_when_nothing_matches(monkeypatch):
    monkeypatch.setattr(serve, "_page_texts",
                        lambda name: ["Just some cover-page marketing copy.",
                                      "A totally unrelated page about strategy."])
    assert serve.quick_find_statements("whatever.pdf") == []


# --------------------------------------------- scan progress / cancel -----
def test_scan_stops_early_when_progress_callback_requests_cancel():
    """Pure logic on X.scan itself: a progress callback returning True must
    stop the scan after that page, not run to completion."""
    calls = []
    def progress(i, total):
        calls.append(i)
        return i >= 1   # cancel after the 2nd page (index 1)
    # an empty pdfs list can't exercise the per-page loop; verify instead
    # against a real sample if present, else just check the signature accepts it
    if (ROOT / "du annual 2025.pdf").exists():
        X.scan([ROOT / "du annual 2025.pdf"], range(0, 5), 2, 2,
               warn=lambda m: None, progress=progress)
        assert calls and max(calls) <= 1   # never progressed past the cancel point
    else:
        pytest.skip("no sample PDF present")


# ---------------------------------- ruled-numbers-only note (label loss) --
def test_looks_labelless():
    assert X._looks_labelless([]) is True
    assert X._looks_labelless([[None, 1, 2], [None, 3, 4]]) is True
    assert X._looks_labelless([["Revenue", 1], ["Costs", 2], ["Total", 3]]) is False
    # one real label out of many bare rows still counts as labelless overall
    assert X._looks_labelless([["Revenue", 1], [None, 2], [None, 3], [None, 4]]) is True


def test_looks_garbled_catches_several_rows_merged_into_one_cell():
    """Found live (du annual 2011.pdf): the pdfplumber-fallback path's
    default 'lines' strategy, with only sparse ruling to go on, merged 5
    real line items into one cell each, newline-joined -- '6\\n7.1\\n7.2\\n
    7.3\\n7.4' and '6,903,496\\n371,667\\n88,003\\n164,282\\n549,050'. The
    tell is specifically multiple newline-split PIECES that independently
    look like numbers -- a normal wrapped label isn't built from several
    number-shaped fragments."""
    garbled = [[None, "2011\nAED 000", None],
              ["6\n7.1\n7.2\n7.3\n7.4", "6,903,496\n371,667\n88,003\n164,282\n549,050", "x"]]
    assert X._looks_garbled(garbled) is True


def test_looks_garbled_leaves_a_normal_wrapped_label_alone():
    # a genuinely wrapped 2-line label/header -- one embedded newline, but
    # its pieces are text, not several number-shaped fragments
    normal = [["Net book value\nAt 31 December 2009", 334491, 115729, 450220],
              ["Capital work\nin progress", "Buildings", "Total"]]
    assert X._looks_garbled(normal) is False


def test_looks_garbled_uses_a_column_count_relative_threshold_not_a_fixed_one():
    """Found live on du annual 2020.pdf's own balance sheet: a label that
    wraps AROUND its own note-ref and figures (rather than before or after
    them) glues a few number-shaped pieces into one cell for a single real
    row -- 'Financial asset at fair value through other\\n11\\n18,368\\n
    18,368\\ncomprehensive income' (one row's own note-ref + two years'
    figures = 3 number-shaped pieces) in a 4-column table. That must NOT
    read as garbled (it's one real row, not several merged) even though 3
    is still ">= 2" -- the bare absolute threshold the original bug (below)
    was fixed with would have thrown away this whole, otherwise-correct
    statement. The same 3-piece cell in a NARROWER table (fewer columns
    than pieces) still must trip it -- a single row can never hold more
    numbers than the table has columns."""
    one_real_row_wrapped_around_its_own_figures = [
        ["As at 31 December", "2020", "2019", None],
        ["Financial asset at fair value through other\n11\n18,368\n18,368\ncomprehensive income",
         None, None, None],
    ]
    assert X._looks_garbled(one_real_row_wrapped_around_its_own_figures) is False
    too_many_numbers_for_the_table_s_own_width = [
        ["Header", None],
        ["Financial asset at fair value through other\n11\n18,368\n18,368\ncomprehensive income",
         None],
    ]
    assert X._looks_garbled(too_many_numbers_for_the_table_s_own_width) is True


def test_attach_left_labels_recovers_text_left_of_the_ruled_box():
    class FakeWord(dict):
        pass
    def W(text, x0, top):
        return {"text": text, "x0": x0, "top": top}
    class FakePage:
        def extract_words(self):
            return [W("Short", 40, 100), W("term", 65, 100), W("benefits", 90, 100),
                   W("Termination", 40, 115), W("benefits", 95, 115),
                   # a word that belongs to a DIFFERENT row (outside both bands) must not leak in
                   W("Unrelated", 40, 400)]
    rows = [[30929, 25494], [1128, 608]]
    row_bands = [(99, 111), (114, 126)]
    out = X._attach_left_labels(FakePage(), rows, row_bands, search_x0=0, region_x0=400)
    assert out == [["Short term benefits", 30929, 25494],
                   ["Termination benefits", 1128, 608]]


def test_attach_left_labels_respects_the_search_boundary():
    def W(text, x0, top):
        return {"text": text, "x0": x0, "top": top}
    class FakePage:
        def extract_words(self):
            return [W("TooFarLeft", 5, 100), W("RealLabel", 50, 100)]
    rows = [[1]]
    row_bands = [(99, 111)]
    # search_x0=40 excludes the word at x0=5 -- it's outside the user's drawn box
    out = X._attach_left_labels(FakePage(), rows, row_bands, search_x0=40, region_x0=400)
    assert out == [["RealLabel", 1]]


def test_attach_left_labels_keeps_two_wrapped_lines_in_reading_order():
    """Found live (du annual 2010.pdf p22): img2table's own cell bbox for a
    row can be taller than a single text line -- e.g. a ruled/bordered
    numbers-only row whose matching label wraps onto two lines ('Net book
    value' / 'At 31 December 2009'). The row_band search picks up words from
    BOTH lines; sorting all of them by x0 together (ignoring which line each
    came from) interleaves the two lines word-by-word instead of reading one
    line fully before the next -- 'Net At 31 book December value 2009', not
    'Net book value At 31 December 2009'. This is the regression lock for
    the fix: group by line (top) first, then sort left-to-right within it."""
    def W(text, x0, top):
        return {"text": text, "x0": x0, "top": top}
    class FakePage:
        def extract_words(self):
            return [
                # line 1, left-to-right
                W("Net", 10, 500), W("book", 25, 500), W("value", 45, 500),
                # line 2, left-to-right -- deliberately positioned so a naive
                # x0-only sort across BOTH lines would interleave with line 1
                # (matches the real-world x0s that produced the bug)
                W("At", 10, 512), W("31", 20, 512), W("December", 28, 512), W("2009", 60, 512),
            ]
    rows = [[334491, 115729, 450220]]
    row_bands = [(498, 526)]     # one wide band spanning both physical lines
    out = X._attach_left_labels(FakePage(), rows, row_bands, search_x0=0, region_x0=400)
    assert out == [["Net book value At 31 December 2009", 334491, 115729, 450220]]


@pytest.mark.skipif(not HAVE_IMG2TABLE, reason="img2table not installed")
@pytest.mark.skipif(not (ROOT / "du annual 2010.pdf").exists(), reason="sample PDF not present")
def test_extract_region_recovers_labels_for_a_ruled_numbers_only_box():
    """Regression: img2table only detects the RULED numeric grid on this
    note -- the row labels sit outside the ruled border entirely (unruled),
    so img2table's own structure detection never sees them. Before the fix,
    this returned a "table" of bare numbers with a 0% label-health score;
    some box sizes returned nothing at all ("No table found in that box")."""
    t = X.extract_region(ROOT / "du annual 2010.pdf", 22, (170, 130, 545, 260))
    assert t is not None
    assert t["title"] == "8.2 Compensation to key management personnel"
    assert t["kind"] == "note"
    labels = [r[0] for r in t["rows"] if r[0]]
    assert "Short term employee benefits" in labels
    assert "Termination benefits" in labels
    assert t["health_labels"]["score"] > 0.5


def test_split_glued_cell_handles_reversed_parens_too():
    """Found live via the comprehensive audit (en-2020-etisalat p75): img2table
    can glue two adjacent value columns into one cell on a given row just like
    the normal-parens case _split_glued_cell already handled -- except when the
    source PDF's own bidi artifact has already reversed each number's parens
    (see parse_number), the glued result is two REVERSED tokens back to back."""
    assert _split_glued_cell(")87,579( )11,915(") == [")87,579(", ")11,915("]
    assert _split_glued_cell(")9,408( )896,525( )50,749(") == \
        [")9,408(", ")896,525(", ")50,749("]
    # unaffected: normal orientation, and a single (non-glued) reversed value
    assert _split_glued_cell("(1,234) (5,678)") == ["(1,234)", "(5,678)"]
    assert _split_glued_cell(")1,234(") == [")1,234("]


def test_unscramble_reversed_paren_wrap_reorders_split_lines():
    """Found live (en-2022-1-eand-group-annual-report.pdf p49, two 'Dividends'
    rows): a reversed-parens negative number can ALSO have its digits split
    across two physical lines inside the same img2table cell, with the line
    order itself reversed by the same bidi artifact -- ')69,040\\n1,1(' for
    what the PDF printed as '(1,169,040)'. Confirmed against the real page by
    reconciling the surrounding row's own arithmetic (the 'Total' column value
    minus every other movement in the row), not just by inspection."""
    assert _unscramble_reversed_paren_wrap(") 69,040\n1,1 (") == ")1,169,040("
    assert _unscramble_reversed_paren_wrap(") 670,421\n11, (") == ")11,670,421("
    # a real two-line label must never be touched -- only bare digit/comma
    # lines inside a reversed-paren wrapper qualify
    assert _unscramble_reversed_paren_wrap(")Note\n24(") == ")Note\n24("
    assert _unscramble_reversed_paren_wrap("normal text\nwith a newline") == \
        "normal text\nwith a newline"
    # no newline at all -- nothing for this to do
    assert _unscramble_reversed_paren_wrap(")1,234(") == ")1,234("


def test_merge_side_by_side_joins_a_label_block_with_a_figures_block():
    # the documented legitimate case: one region is pure labels (no numbers),
    # the other is pure figures (no labels) -- these really are one table
    # split left/right, and should still be joined exactly as before.
    labels = _region(0, 0, 100, 50, [(0, 10, ["Revenue"]), (10, 20, ["Costs"])])
    figures = _region(105, 0, 200, 50, [(0, 10, [1000]), (10, 20, [-400])])
    out = _merge_side_by_side([labels, figures])
    assert len(out) == 1
    assert [v for _, _, v in out[0].row_rows] == [["Revenue", 1000], ["Costs", -400]]


def test_merge_side_by_side_carries_raw_cells_through():
    """raw_cells (see _RowRow) must survive _join_side_by_side, reusing the
    EXACT SAME Y-overlap match already computed for `vals` -- not a second,
    independent matching pass that could disagree with it."""
    labels = _region(0, 0, 100, 50, [
        _RowRow(0, 10, ["Revenue"], raw_cells=[(0, 100, "Revenue")]),
        _RowRow(10, 20, ["Costs"], raw_cells=[(0, 100, "Costs")]),
    ])
    figures = _region(105, 0, 200, 50, [
        _RowRow(0, 10, [1000], raw_cells=[(105, 200, 1000)]),
        _RowRow(10, 20, [-400], raw_cells=[(105, 200, -400)]),
    ])
    out = _merge_side_by_side([labels, figures])
    assert len(out) == 1
    raw = [row.raw_cells for row in out[0].row_rows]
    assert raw == [
        [(0, 100, "Revenue"), (105, 200, 1000)],
        [(0, 100, "Costs"), (105, 200, -400)],
    ]


def test_merge_side_by_side_raw_cells_is_none_when_either_side_lacks_it():
    """A merged row's raw_cells must be None (not a partial/fabricated
    guess) when either side of the match doesn't have real geometry -- e.g.
    a synthetic region built the old way (plain tuples, no raw_cells at
    all), or no Y-overlap match was found on the other side at all. A
    caller filtering by column position on a None-geometry row should pass
    it through unfiltered, never silently drop it or guess."""
    labels = _region(0, 0, 100, 50, [(0, 10, ["Revenue"])])   # plain tuple, no raw_cells
    figures = _region(105, 0, 200, 50, [
        _RowRow(0, 10, [1000], raw_cells=[(105, 200, 1000)]),
    ])
    out = _merge_side_by_side([labels, figures])
    assert len(out) == 1
    assert out[0].row_rows[0].raw_cells is None


def test_merge_side_by_side_leaves_two_independently_labelled_tables_alone():
    """Found live (en-2021-etisalat-group-annual-report.pdf p62): a balance
    sheet and a completely unrelated statement of changes in equity sit side
    by side on a landscape page and are geometrically indistinguishable from
    a genuine labels-block/figures-block split (small x-gap, large y-overlap)
    -- but unlike that legitimate case, BOTH sides already have their own row
    labels, which is the actual signature of two independently-complete
    tables. They must not be zipped together row by row."""
    balance_sheet = _region(0, 0, 100, 60, [
        (0, 10, ["Goodwill and other intangible assets", 25830041]),
        (10, 20, ["Property, plant and equipment", 43715088]),
        (20, 30, ["Right-of-use assets", 2436921]),
        (30, 40, ["Total assets", 128197066]),
    ])
    equity_statement = _region(105, 0, 200, 60, [
        (0, 10, ["Balance at 1 January 2020", 8696754]),
        (10, 20, ["Profit for the year", 9026522]),
        (20, 30, ["Other comprehensive income for the year", 376376]),
        (30, 40, ["Balance at 31 December 2020", 60550021]),
    ])
    out = _merge_side_by_side([balance_sheet, equity_statement])
    assert len(out) == 2


def test_looks_like_two_fused_statements_catches_a_real_case():
    """Found live (en-2021-etisalat-group-annual-report.pdf p62): img2table's
    own borderless-table clustering, not any merge this project performs,
    fused a balance sheet and a completely different statement of changes in
    equity into one region -- correctly detecting this is what lets
    img2table_page_tables refuse it rather than serve rows that attach one
    statement's columns to the other's."""
    fused_rows = [
        ["Non-current assets", None, None, None, None, "Share", None],
        ["Goodwill and other intangible assets", 11, 25830041, 26276442,
         None, "capital", "Reserves"],
        ["Property, plant and equipment", 13, 43715088, 45803436,
         "Balance at 1 January 2020", 8696754, 27812896],
        ["Total assets", None, 46979699, 49698687,
         "Profit for the year", None, 9026522],
    ]
    assert _looks_like_two_fused_statements(fused_rows) is True


def test_looks_like_two_fused_statements_ignores_a_passing_prose_mention():
    """A first pass at this check (a 24-file, 2428-page sweep of every real
    loaded file) found two false positives before the prose filter was
    added: accounting-policy notes (du annual 2014.pdf p26, du annual
    2017.pdf p29) that mention "transactions with the owners" in an
    ordinary 8-word sentence, nowhere near a real equity statement -- long
    enough to read as prose (_looks_like_prose's own len(s.split()) >= 8
    rule) rather than a short row-label-shaped cell. Real sentence from
    that page, verified to still trip _SOCE_ANCHOR_RE on its own (so this
    test would fail without the prose filter, not pass by coincidence)."""
    soce_phrase_in_a_sentence = "that is, as transactions with the owners in"
    assert _SOCE_ANCHOR_RE.search(soce_phrase_in_a_sentence)   # matches on its own
    prose_rows = [
        ["Total assets acquired in a business combination are", soce_phrase_in_a_sentence,
         "their capacity as owners. The difference"],
    ]
    assert _looks_like_two_fused_statements(prose_rows) is False


def test_looks_like_two_fused_statements_catches_income_statement_and_cash_flow():
    """Found live on du annual 2011.pdf p23: img2table's own clustering
    fused a side-by-side income statement and cash flow statement into one
    8-column region, zipping "Revenue"/"Cost of sales" together with "Cash
    flows from operating activities" row-by-row. A real user manually
    marking just the visible income statement got the cash flow
    statement's rows glued in beside it, because this fusion happens
    inside img2table's own region detection, before the user's box (or
    this project's own side-by-side merge logic) ever enters the
    picture."""
    fused_rows = [
        ["For the year ended 31 December", 2011, 2010,
         "For the year ended 31 December", None, 2011, 2010],
        ["Revenue", 8854683, 7074097,
         "Cash flows from operating activities", None, None, None],
        ["Cost of sales", -2953912, -2473791,
         "Net cash flows before changes in working capital", 25, 2262322, 2151051],
        ["Gross profit", 5900771, 4600306,
         "Change in inventories", None, -4962, -8369],
    ]
    assert _looks_like_two_fused_statements(fused_rows) is True


def test_looks_like_two_fused_statements_position_gate_rejects_mid_sentence_match():
    """The anchor-position check (replacing the old _looks_like_prose gate)
    must still reject an anchor phrase buried inside a longer sentence --
    not just accept anything short. "Cost of sales" appearing mid-clause in
    a policy sentence (not as its own row label) must not count as a P&L
    hit on its own."""
    prose_rows = [
        ["The classification of cost of sales between segments is disclosed in note 4"],
    ]
    assert _looks_like_two_fused_statements(prose_rows) is False


def test_looks_like_two_fused_statements_single_kind_alone_is_not_fusion():
    """A real, unfused cash flow statement (only CF anchors, no other
    statement kind's vocabulary) must not be flagged -- one kind's hit
    alone is never enough."""
    real_cash_flow_rows = [
        ["Cash flows from operating activities", None, None],
        ["Profit for the year", 11059489, 10315736],
        ["Cash and cash equivalents at end of the year", 2376371, 2785478],
    ]
    assert _looks_like_two_fused_statements(real_cash_flow_rows) is False


# ------------------------------------------------- mid-row glued figures ----
def test_redistribute_glued_cells_pulls_a_mid_row_glued_cell_apart():
    """Found live on Etisalat's balance sheet: a row already at the
    table's real width, but with several years' figures glued into ONE
    cell while the cells after it are the untouched None padding those
    figures should have landed in -- img2table found the right column
    count overall but merged three of them together on this particular
    row (a narrower gap here than on other rows)."""
    vals = ["Contract assets", 22, "432,541 221,711 205,270", None, None]
    assert _redistribute_glued_cells(vals) == \
        ["Contract assets", 22, "432,541", "221,711", "205,270"]


def test_redistribute_glued_cells_leaves_a_full_row_alone():
    vals = ["label", "note", "1,234", "5,678", "9,012"]
    assert _redistribute_glued_cells(vals) == vals


def test_redistribute_glued_cells_refuses_when_there_is_not_enough_room():
    """Splitting would need MORE slots than the row actually has left --
    correctly declined rather than overwriting or truncating."""
    vals = ["label", "432,541 221,711 205,270", "9,012"]
    assert _redistribute_glued_cells(vals) == vals
