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
