"""
Pure-logic unit tests -- NO PDFs required, so these run in CI.

They exercise the analysis layer on hand-built table dicts:
  analyze / figure_health / label_health / cross_year_check / _deprose_labels
  / _looks_like_not_a_statement / _stitch_page_breaks / _cashflow_foots
  / reconcile_explain
"""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
import extract_all_tables as X  # noqa: E402


# --------------------------------------------------------------- fixtures ----
def _pl_rows():
    return [
        ["", 2024, 2023],
        ["Revenue", 1_000_000, 900_000],
        ["Cost of sales", -400_000, -350_000],
        ["Gross profit", 600_000, 550_000],
        ["Operating expenses", -200_000, -180_000],
        ["Finance costs", -50_000, -40_000],
        ["Profit before tax", 350_000, 330_000],
        ["Income tax", -50_000, -30_000],
        ["Profit for the year", 300_000, 300_000],
    ]


def _bs_rows():
    return [
        ["", 2024, 2023],
        ["Non-current assets", None, None],
        ["Property, plant and equipment", 800_000, 750_000],
        ["Total non-current assets", 800_000, 750_000],
        ["Current assets", None, None],
        ["Inventories", 200_000, 180_000],
        ["Total current assets", 200_000, 180_000],
        ["Total assets", 1_000_000, 930_000],
        ["Equity", None, None],
        ["Share capital", 400_000, 400_000],
        ["Retained earnings", 200_000, 150_000],
        ["Total equity", 600_000, 550_000],
        ["Total liabilities", 400_000, 380_000],
        ["Total equity and liabilities", 1_000_000, 930_000],
    ]


def _analyzed(rows, title):
    t = {"rows": rows, "title": title, "file": "x.pdf", "page_label": 1,
         "shape": X.classify(rows)}
    X.analyze(t)
    X._attach_health(t)
    return t


# ------------------------------------------------------------------ analyze --
def test_analyze_income_statement_foots():
    t = _analyzed(_pl_rows(), "Consolidated statement of profit or loss")
    assert t["kind"] == "income statement"
    assert t["years"] == [2024, 2023]
    assert t["foots"] is True
    assert [fc["ok"] for fc in t["foot_by_col"]] == [True, True]


def test_analyze_balance_sheet_foots_both_columns():
    t = _analyzed(_bs_rows(), "Consolidated statement of financial position")
    assert t["kind"] == "statement of financial position"
    assert t["foots"] is True
    assert len(t["foot_by_col"]) == 2 and all(fc["ok"] for fc in t["foot_by_col"])


def test_analyze_broken_pl_does_not_foot():
    rows = _pl_rows()
    rows[2][1] = -300_000      # cost of sales 100k too small in 2024
    t = _analyzed(rows, "statement of profit or loss")
    # current-year column is broken, prior-year still fine
    oks = [fc["ok"] for fc in t["foot_by_col"]]
    assert oks[0] is False and oks[1] is True


def test_analyze_downgrades_mdna_highlights():
    rows = [["", 2024, 2023],
            ["Revenue", 1000, 900],
            ["EBITDA", 400, 380],
            ["EBITDA margin", 40, 42],
            ["Net profit", 300, 280]]
    t = _analyzed(rows, "Profit and Loss Summary")
    assert t["kind"] == "table"          # ebitda / margin / "summary" -> not a face statement


def test_analyze_downgrades_transition_bridge():
    # previously-reported + adjustment + restated : col_a + col_b == col_c
    rows = [["", "As reported", "IFRS 16", "Restated"],
            ["Revenue", 1000, 0, 1000],
            ["Right-of-use assets", 0, 120, 120],
            ["Total assets", 1000, 120, 1120],
            ["Lease liabilities", 0, 120, 120],
            ["Profit for the year", 300, -5, 295]]
    t = _analyzed(rows, "Statement of profit or loss")
    assert t["kind"] == "table"


# ------------------------------------------------------------ figure_health --
def test_figure_health_flags_extra_digit():
    rows = _pl_rows()
    rows[5][1] = -50_000_000_000   # finance costs with an extra 6 digits
    t = {"rows": rows, "title": "p&l", "file": "x", "page_label": 1,
         "shape": X.classify(rows)}
    X.analyze(t)
    fh = X.figure_health(t)
    assert any("Finance costs" in s["label"] for s in fh["suspect"])
    assert fh["score"] < 1.0


def test_figure_health_flags_100x_year_gap():
    rows = _pl_rows()
    rows[4][2] = -1              # operating expenses: 2024=-200000, 2023=-1
    t = {"rows": rows, "title": "p&l", "file": "x", "page_label": 1,
         "shape": X.classify(rows)}
    X.analyze(t)
    fh = X.figure_health(t)
    assert any(s["why"].startswith("one year") for s in fh["suspect"])


def test_figure_health_clean_statement_scores_one():
    t = _analyzed(_pl_rows(), "statement of profit or loss")
    assert t["health"]["figures"] == 1.0


# ------------------------------------------------------------- label_health --
def test_label_health_flags_prose_and_empty():
    rows = _pl_rows()
    rows.insert(4, ["We have audited the accompanying consolidated financial "
                    "statements and in our opinion they present fairly", None, None])
    rows.insert(5, [None, 123456, 111111])          # figures, no label
    t = {"rows": rows, "title": "p&l", "file": "x", "page_label": 1,
         "shape": X.classify(rows)}
    X.analyze(t)
    lh = X.label_health(t)
    whys = " ".join(s["why"] for s in lh["suspect"])
    assert "prose" in whys and "no label" in whys


# --------------------------------------------------------- _deprose_labels ---
def test_deprose_keeps_trailing_line_item():
    rows = [["In our opinion the consolidated financial statements present "
             "fairly Total non-current assets", 800, 750]]
    out = X._deprose_labels(rows)
    assert out[0][0] == "Total non-current assets"
    assert out[0][1] == 800


def test_deprose_keeps_leading_line_item():
    rows = [["Total current assets We have audited the accompanying "
             "consolidated financial statements", 200, 180]]
    out = X._deprose_labels(rows)
    assert out[0][0].startswith("Total current assets")


def test_deprose_leaves_clean_labels_untouched():
    rows = [["Revenue", 1000, 900], ["Total assets", 5000, 4800]]
    assert X._deprose_labels([r[:] for r in rows]) == rows


# ------------------------------------------------ _looks_like_not_a_statement -
def test_downgrade_true_when_too_few_anchors():
    data = [["Some heading", None, None], ["One line", 5, 4], ["Another", 6, 5]]
    assert X._looks_like_not_a_statement(
        "income statement", data, [1, 2], [2024, 2023], []) is True


def test_downgrade_false_for_real_statement():
    rows = _pl_rows()
    data = rows[1:]
    totals = [i for i, r in enumerate(rows) if X._TOTAL_RE.match(X._row_label(r) or "")]
    assert X._looks_like_not_a_statement(
        "income statement", data, [1, 2], [2024, 2023], totals) is False


def test_downgrade_true_on_bad_year_gap():
    rows = _pl_rows()
    assert X._looks_like_not_a_statement(
        "income statement", rows[1:], [1, 2], [2024, 2010], [3, 6, 8]) is True


# ------------------------------------------------------- cross_year_check ----
def _stmt(file, year_pair, rows, kind="income statement"):
    t = {"rows": rows, "title": "s", "file": file, "page_label": 1,
         "shape": X.classify(rows)}
    X.analyze(t)
    t["years"] = year_pair
    return t


def test_cross_year_agrees_when_columns_match():
    a = _stmt("2024.pdf", [2024, 2023], _pl_rows())
    b_rows = _pl_rows()
    for r in b_rows[1:]:
        if isinstance(r[1], (int, float)):
            r[1] = r[2]                 # b's current year == a's prior year
    b = _stmt("2023.pdf", [2023, 2022], b_rows)
    X.cross_year_check([a, b])
    assert a["consistency"]["mismatch"] == 0
    assert a["consistency"]["year"] == 2023


def test_cross_year_flags_a_restatement():
    a = _stmt("2024.pdf", [2024, 2023], _pl_rows())
    b_rows = _pl_rows()
    for r in b_rows[1:]:
        if isinstance(r[1], (int, float)):
            r[1] = r[2]
    b_rows[1][1] = 850                  # b's 2023 Revenue differs from a's 2023
    b = _stmt("2023.pdf", [2023, 2022], b_rows)
    X.cross_year_check([a, b])
    assert a["consistency"]["mismatch"] >= 1
    assert any("revenue" in m[0] for m in a["consistency"]["worst"])


def test_cross_year_small_diff_is_called_a_restatement():
    a = _stmt("2024.pdf", [2024, 2023], _pl_rows())
    b_rows = _pl_rows()
    for r in b_rows[1:]:
        if isinstance(r[1], (int, float)):
            r[1] = r[2]
    b_rows[1][1] = 850_000                         # ONE line differs
    b = _stmt("2023.pdf", [2023, 2022], b_rows)
    X.cross_year_check([a, b])
    assert a["consistency"]["verdict"] == "restated"


def test_cross_year_many_diffs_is_called_misaligned():
    a = _stmt("2024.pdf", [2024, 2023], _pl_rows())
    b_rows = _pl_rows()
    for r in b_rows[1:]:                            # b's current column = a's prev + 1
        if isinstance(r[1], (int, float)):
            r[1] = r[2] + 100_000
    b = _stmt("2023.pdf", [2023, 2022], b_rows)
    X.cross_year_check([a, b])
    assert a["consistency"]["verdict"] == "columns likely misaligned"


def test_year_reversal_is_noted():
    rows = _pl_rows()
    rows[0] = ["", 2023, 2024]                      # header lists years oldest-first
    t = {"rows": rows, "title": "statement of profit or loss", "file": "x",
         "page_label": 1, "shape": X.classify(rows)}
    X.analyze(t)
    assert any("year columns may be reversed" in n for n in (t.get("notes") or []))
    # notes_i18n (webui.html's translated-note channel, see _add_note) must
    # stay in lockstep with "notes" -- same length, same order, so the
    # frontend can pair notes[i] with notes_i18n[i] by index
    notes, notes_i18n = t["notes"], t["notes_i18n"]
    assert len(notes) == len(notes_i18n)
    i = next(k for k, n in enumerate(notes) if "may be reversed" in n)
    assert notes_i18n[i]["key"] == "yearsReversed"
    assert notes_i18n[i]["vars"] == {"first": 2023, "last": 2024}


def test_add_note_keeps_notes_and_notes_i18n_in_lockstep_across_multiple_calls():
    t = {}
    X._add_note(t, "segmentalColumns", "columns look like segments...")
    X._add_note(t, "assetsOnlyIncomplete", "equity / liabilities side incomplete...")
    assert len(t["notes"]) == len(t["notes_i18n"]) == 2
    assert [m["key"] for m in t["notes_i18n"]] == ["segmentalColumns", "assetsOnlyIncomplete"]


def test_figure_health_thresholds_are_derived_not_fixed():
    # a column of similar-magnitude values: no outlier even at 10x median,
    # because the spread is tight -> a real extra-digit still trips it
    rows = [["", 2024, 2023]] + [
        [f"Item {i}", v, v] for i, v in enumerate(
            [100_000, 110_000, 95_000, 105_000, 98_000, 102_000])]
    t = {"rows": rows, "title": "note", "file": "x", "page_label": 1,
         "shape": X.classify(rows)}
    X.analyze(t)
    assert X.figure_health(t)["score"] == 1.0
    rows[3][1] = 99_000_000                          # ~1000x -> outlier
    X.analyze(t)
    assert X.figure_health(t)["score"] < 1.0


def test_cashflow_tolerance_scales_with_size():
    # a ~120-unit drift passes on a small statement but is caught relative to
    # a billions-scale one only if it exceeds 0.05% -- here 120 on a 1e6 flow
    small = [["Net cash generated from operating activities", 5000],
             ["Net cash used in investing activities", -2000],
             ["Net cash used in financing activities", -1000],
             ["Net increase in cash", 2400],                  # 400 off > 200 default
             ["Cash and cash equivalents at the end of the year", 2400]]
    assert X._cashflow_foots(small, 1) is False
    big = [["Net cash generated from operating activities", 5_000_000],
           ["Net cash used in investing activities", -2_000_000],
           ["Net cash used in financing activities", -1_000_000],
           ["Net increase in cash", 2_000_400],               # 400 off, tol scales to 1000
           ["Cash and cash equivalents at the end of the year", 2_000_400]]
    assert X._cashflow_foots(big, 1) is True


# -------------------------------------------------------- _stitch_page_breaks -
def test_stitch_merges_bs_continuation():
    top = _analyzed(_bs_rows()[:8], "Statement of financial position")  # assets only
    top["file"] = "r.pdf"; top["page_label"] = 10
    cont = {"rows": [["Share capital", 400_000, 400_000], ["Retained earnings", 200_000, 150_000],
                     ["Total equity", 600_000, 550_000], ["Total liabilities", 400_000, 380_000],
                     ["Total equity and liabilities", 1_000_000, 930_000]],
            "title": "", "file": "r.pdf", "page_label": 11,
            "shape": "table", "kind": "table", "header_idx": 0}
    X.analyze(cont)
    tables = [top, cont]
    X._stitch_page_breaks(tables)
    assert len(tables) == 1
    assert tables[0].get("_stitched_from") == 11
    assert tables[0]["foots"] is True


def test_stitch_rolls_back_a_bad_merge():
    top = _analyzed(_bs_rows()[:8], "Statement of financial position")
    top["file"] = "r.pdf"; top["page_label"] = 10
    before = [r[:] for r in top["rows"]]
    junk = {"rows": [["Lease liabilities", 1], ["Contract liabilities", 2],
                     ["random prose sentence about nothing at all here", 3]],
            "title": "", "file": "r.pdf", "page_label": 11,
            "shape": "table", "kind": "table", "header_idx": 0}
    X.analyze(junk)
    # force analyze() to blow up on the merged rows
    import unittest.mock as mock
    real = X.analyze
    calls = {"n": 0}

    def boom(t, **kw):
        calls["n"] += 1
        if calls["n"] > 0 and t is top:
            raise RuntimeError("simulated failure")
        return real(t, **kw)
    with mock.patch.object(X, "analyze", boom):
        X._stitch_page_breaks([top, junk])
    assert top["rows"] == before          # rolled back


# ------------------------------------------------------------ _cashflow_foots -
def test_cashflow_foots_opening_plus_net_plus_fx():
    data = [
        ["Net cash generated from operating activities", 500],
        ["Net cash used in investing activities", -200],
        ["Net cash used in financing activities", -100],
        ["Net increase in cash and cash equivalents", 200],
        ["Cash and cash equivalents at the beginning of the year", 1000],
        ["Effect of foreign exchange rate changes", -50],
        ["Cash and cash equivalents at the end of the year", 1150],
    ]
    assert X._cashflow_foots(data, 1) is True


def test_cashflow_foots_du_style_no_investing_subtotal():
    data = [
        ["Net cash generated from operating activities", 500],
        ["Purchase of property, plant and equipment", -150],
        ["Interest received", 30],
        ["Cash flows from financing activities", None],
        ["Dividends paid", -80],
        ["Net cash used in financing activities", -80],
        ["Net increase in cash and cash equivalents", 300],
    ]
    # operating 500 + (investing -150+30=-120) + financing -80 == net 300
    assert X._cashflow_foots(data, 1) is True


# ----------------------------------------------------- Excel export format --
def test_header_row_years_are_not_thousands_grouped_in_excel():
    """Found live: a column header showing "2,025" instead of "2025" -- the
    export's number_format was applying "#,##0" to every numeric cell,
    header row included, and a year isn't a quantity that benefits from
    thousands-grouping. Real data cells must still get it."""
    t = _analyzed(_pl_rows(), "Consolidated statement of profit or loss")
    wb = X.build_workbook([t])
    ws = wb.worksheets[1]              # worksheets[0] is the Contents index
    hi = t["header_idx"]
    r0 = 4
    # the header row's year cells (2024, 2023) -- plain, no grouping
    for cx in (2, 3):
        cell = ws.cell(row=r0 + hi, column=cx)
        assert cell.value in (2024, 2023)
        assert cell.number_format == "0"


# ---------------------------------------------------------- note-ref column --
def test_note_ref_map_survives_analyze_without_colliding_with_note_col():
    """Found live: extract_region/find_all_tables stash the PDF's own
    "Note" reference column (e.g. "19") as a display-only side channel,
    SET BEFORE analyze() runs -- but analyze() already owns an unrelated
    "note_col" field of its own (a column-INDEX hint, see analyze()'s
    `notecol` local), and unconditionally overwrites it via t.update().
    The two fields must use different keys, or the note-ref map is
    silently clobbered the instant analyze() runs -- exactly what
    happened before this was caught (see row_note_ref)."""
    rows = _pl_rows()
    t = {"rows": rows, "title": "P&L", "file": "x.pdf", "page_label": 1,
         "shape": X.classify(rows),
         "note_ref_map": {"revenue": "6", "cost of sales": "7"}}
    X.analyze(t)
    X._attach_health(t)
    assert t["note_ref_map"] == {"revenue": "6", "cost of sales": "7"}
    assert X.row_note_ref(t, rows[1]) == "6"     # Revenue
    assert X.row_note_ref(t, rows[2]) == "7"     # Cost of sales
    assert X.row_note_ref(t, rows[3]) is None    # Gross profit -- no note


def test_build_workbook_writes_note_ref_as_its_own_display_column():
    """The Notes-reference side channel (see above) has to actually reach
    the exported sheet -- as its own column right after the label, not
    merged into it -- without shifting the Δ/Δ% columns off the real
    value columns."""
    rows = _pl_rows()
    t = _analyzed(rows, "P&L")
    t["note_ref_map"] = {"revenue": "6"}
    wb = X.build_workbook([t])
    ws = wb.worksheets[1]
    r0 = 4
    revenue_i = next(i for i, r in enumerate(rows) if r[0] == "Revenue")
    assert ws.cell(row=r0 + revenue_i, column=2).value == "6"
    assert ws.cell(row=r0 + revenue_i, column=3).value == 1_000_000
    # Δ header still lands right after the shifted-over value columns
    hi = t["header_idx"]
    assert ws.cell(row=r0 + hi, column=5).value == "Δ (change)"


# ------------------------------------------------------- wrapped labels -----
def test_merge_wrapped_labels_rejoins_a_split_total_row():
    """Found live: img2table's own row detection (and the regex
    reconstruction) splits a label that wraps onto a second physical line
    into two table rows. When the SPLIT row is itself a "Total ..." line,
    the half carrying the real figures reads as some unrelated fragment
    ("amortization") instead -- which broke footing entirely, since
    nothing recognizable was left to anchor the section boundary."""
    rows = [
        ["Total net operating expenses before depreciation and", None, None],
        ["amortization", -3307608, -3347636],
    ]
    out = X._merge_wrapped_labels(rows)
    assert len(out) == 1
    assert out[0] == ["Total net operating expenses before depreciation and amortization",
                       -3307608, -3347636]


def test_merge_wrapped_labels_leaves_two_real_rows_alone():
    """The merge must be conservative: a row starting with a capital
    letter reads as a genuine new line item, not a wrap continuation --
    even when the row above it happens to carry no figures (a normal
    section header, e.g. "Direct costs")."""
    rows = [
        ["Direct costs", None, None],
        ["Interconnect cost", -100, -90],
    ]
    assert X._merge_wrapped_labels(rows) == rows


def test_merge_wrapped_labels_does_not_merge_into_a_row_that_already_has_figures():
    """A lowercase-starting label is only treated as a continuation when
    the row above is label-only -- if the row above already has its own
    figures, it's a complete, real row, and merging would silently
    overwrite its values with the row below's."""
    rows = [
        ["income before tax (restated)", 100, 90],
        ["excluding one-off items", 5, 4],
    ]
    assert X._merge_wrapped_labels(rows) == rows


# --------------------------------------------------- footing section scope --
def _sibling_sections_rows():
    """Three INDEPENDENT subtotal sections (revenue / direct costs / opex)
    with no overarching "profit for the year" row -- e.g. only the left
    half of a 2-up landscape page was boxed. "Total operating expenses"
    has nothing numerically to do with revenue or direct costs above it.
    Values are in the thousands, matching real statement magnitude -- the
    existing trailing-EPS-tail trim (values < 1000) would otherwise eat
    through a table of toy-sized numbers before this code path even runs."""
    return [
        ["", 2024, 2023],
        ["Revenue A", 100_000, 90_000],
        ["Revenue B", 50_000, 40_000],
        ["Total revenue", 150_000, 130_000],
        ["Cost A", -30_000, -25_000],
        ["Cost B", -20_000, -15_000],
        ["Total direct costs", -50_000, -40_000],
        ["Opex A", -10_000, -8_000],
        ["Opex B", -5_000, -4_000],
        ["Total operating expenses", -15_000, -12_000],
    ]


def test_income_statement_foots_a_partial_box_against_only_its_own_section():
    """Regression for the live bug: checking the WHOLE column (every
    sibling section's leaves) against the last section's own total wrongly
    reported NO FOOT. The check must narrow to just the section ending in
    the final row once the full walk demonstrably fails."""
    t = _analyzed(_sibling_sections_rows(), "Income statement")
    assert t["foots"] is True
    assert "sum of 2 line items = -15,000" in t["foot_detail"]


def _cascading_multi_total_rows():
    """A COMPLETE, normal cascading P&L using "Net income" terminology --
    never matches the "profit for the year" / OCI wording the cut-search
    looks for, so this hits the exact same "no recognised ending" code
    path as the partial-box case above. Unlike that case, "Gross profit"
    genuinely must stay in the running total feeding the final line.
    Values in the thousands for the same reason as the fixture above."""
    return [
        ["", 2024, 2023],
        ["Revenue", 100_000, 90_000],
        ["Cost of sales", -40_000, -35_000],
        ["Gross profit", 60_000, 55_000],
        ["Operating expenses", -20_000, -18_000],
        ["Net income", 40_000, 37_000],
    ]


def test_income_statement_foots_a_complete_cascade_without_narrowing():
    """Non-regression guard for the fix above: a statement with NO
    recognised ending (so the same code path runs) but where the full,
    unscoped walk already reconciles must NOT be narrowed -- narrowing it
    would throw away "Gross profit"'s contribution and break a
    previously-working statement. Found live: a synthetic "Net income"
    fixture regressed to NO FOOT the first time this was attempted."""
    t = _analyzed(_cascading_multi_total_rows(), "Income statement")
    assert t["foots"] is True
    assert "sum of 3 line items = 40,000" in t["foot_detail"]


# -------------------------------------------------------- equity footing ----
def test_equity_foots_does_not_exclude_a_movement_that_merely_contains_the_word_at():
    """Found live: the exclusion filter for "other balance rows" searched
    for the bare substring "at" anywhere in a label, matching completely
    ordinary English ("financial asset AT fair value", "obligATions") --
    not just a real "As at [date]" balance row. Once a wrapped label is
    correctly rejoined (see _merge_wrapped_labels), a real movement row
    describing fair value "at" a point in time was silently dropped from
    the sum. The exclusion must be date-anchored, same as the opening/
    closing balance-row detection itself."""
    data = [
        ["At 1 January 2024", 1000],
        ["Net profit for the year", 200],
        ["Fair value changes on financial asset at fair value through OCI", -50],
        ["At 31 December 2024", 1150],
    ]
    assert X._equity_foots(data, 1) is True


# ---------------------------------------------------------- header row ------
def test_header_row_with_label_and_years_together_is_recognized():
    """Found live: ["As at 31 December", 2010, 2009] -- a real label PLUS
    the year columns collapsed onto one row -- was being read as a DATA
    row because a year like 2010 is ">= 100", so its own "2010" got
    summed into the footing check as if it were a real figure. A
    plausible year must not disqualify a row from being the header just
    for being numerically large."""
    rows = [
        ["As at 31 December", 2024, 2023],
        ["Revenue", 1_000_000, 900_000],
        ["Cost of sales", -400_000, -350_000],
        ["Profit for the year", 600_000, 550_000],
    ]
    t = {"rows": rows, "title": "Income statement", "file": "x.pdf", "page_label": 1,
         "shape": X.classify(rows)}
    X.analyze(t)
    assert t["header_idx"] == 0
    assert t["data_start"] == 1
    assert t["foots"] is True


def test_header_row_with_a_real_large_figure_is_not_mistaken_for_one():
    """The opposite direction: a genuine data row with big figures must
    stay a data row even though its own label is short, same as before
    this fix."""
    rows = [
        ["", 2024, 2023],
        ["Revenue", 1_000_000, 900_000],
    ]
    t = {"rows": rows, "title": "x", "file": "x.pdf", "page_label": 1,
         "shape": X.classify(rows)}
    X.analyze(t)
    assert t["header_idx"] == 0
    assert t["data_start"] == 1


# ------------------------------------------------------- filename years -----
def test_doc_years_from_filename_reads_the_fiscal_year():
    assert X._doc_years_from_filename("du annual 2015.pdf") == [2015, 2014]
    assert X._doc_years_from_filename("en-2020-etisalat-group-annual-report.pdf") == [2020, 2019]
    assert X._doc_years_from_filename("no_year_here.pdf") is None


def test_finish_manual_table_falls_back_to_filename_year_when_header_is_cropped():
    """Found live: a manually-drawn box that starts right at the first
    data row (the header sits just above where the user meant to click
    "start here") left `years` empty with no fallback at all -- unlike
    the automatic scan() path, which has always had this filename-
    derived fallback for exactly this situation."""
    rows = [
        ["Revenue", 1_000_000, 900_000],
        ["Cost of sales", -400_000, -350_000],
        ["Profit for the year", 600_000, 550_000],
    ]
    t = X._finish_manual_table("du annual 2015.pdf", 40, rows, [0, 0, 100, 100], "Income statement")
    assert t["years"] == [2015, 2014]
