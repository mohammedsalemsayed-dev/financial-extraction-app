"""Pure-logic unit tests for tablekit.scorer -- NO PDFs required.

Synthetic tables only: proves the alignment/scoring math itself is correct
before it's ever trusted against real hand-transcribed ground truth.
"""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
from tablekit.scorer import cell_similarity, score_row, score_table  # noqa: E402


def _pl_rows():
    return [
        ["", 2024, 2023],
        ["Revenue", 1_000_000, 900_000],
        ["Cost of sales", "(400,000)", "(350,000)"],
        ["Gross profit", 600_000, 550_000],
        ["Profit for the year", 300_000, 300_000],
    ]


# ------------------------------------------------------------- cell_similarity
def test_cell_similarity_identical_number():
    assert cell_similarity(1234, 1234) == 1.0


def test_cell_similarity_same_number_different_format():
    assert cell_similarity("(1,234)", -1234) == 1.0
    assert cell_similarity("1,234", 1234) == 1.0


def test_cell_similarity_wrong_number_is_zero_no_partial_credit():
    assert cell_similarity(1234, 1235) == 0.0
    assert cell_similarity(1234, 12340) == 0.0


def test_cell_similarity_number_vs_text_is_zero():
    assert cell_similarity(1234, "Revenue") == 0.0


def test_cell_similarity_identical_text():
    assert cell_similarity("Revenue", "Revenue") == 1.0


def test_cell_similarity_near_miss_text_partial_credit():
    s = cell_similarity("Revenue", "Revenue*")
    assert 0.0 < s < 1.0


def test_cell_similarity_both_blank_is_match():
    assert cell_similarity("", None) == 1.0
    assert cell_similarity(None, None) == 1.0


def test_cell_similarity_blank_vs_value_is_zero():
    assert cell_similarity("", "Revenue") == 0.0
    assert cell_similarity(None, 1234) == 0.0


# ----------------------------------------------------------------- score_row
def test_score_row_identical():
    row = ["Revenue", 1_000_000, 900_000]
    score, _ = score_row(row, row)
    assert score == 1.0


def test_score_row_one_wrong_cell():
    gt = ["Revenue", 1_000_000, 900_000]
    ex = ["Revenue", 1_000_000, 999_999]
    score, _ = score_row(gt, ex)
    assert 0.5 < score < 1.0


def test_score_row_extra_column_extracted():
    gt = ["Revenue", 1_000_000]
    ex = ["Revenue", 1_000_000, "Note 4"]
    score, pairs = score_row(gt, ex)
    # a trailing extra cell is a free end gap -- shouldn't tank an
    # otherwise-perfect row match
    assert score == 1.0


# --------------------------------------------------------------- score_table
def test_score_table_identical_is_perfect():
    rows = _pl_rows()
    result = score_table(rows, rows)
    assert result["score"] == 1.0
    assert len(result["rows"]) == len(rows)


def test_score_table_extra_leading_and_trailing_rows_free():
    gt = _pl_rows()
    ex = [["AED'000"]] + gt + [["Source: annual report"]]
    result = score_table(gt, ex)
    # leading/trailing extras (e.g. a units disclaimer, a footnote) are
    # free end gaps -- the extractor grabbed a slightly wider crop but got
    # every real row right, so this should still score as a perfect match
    assert result["score"] == 1.0


def test_score_table_missing_middle_row_is_penalized():
    gt = _pl_rows()
    ex = gt[:2] + gt[3:]  # drop "Cost of sales" from the middle
    result = score_table(gt, ex)
    assert result["score"] < 1.0
    notes = [r.get("note") for r in result["rows"]]
    assert "missing row (in ground truth, not extracted)" in notes


def test_score_table_wrong_value_is_penalized_but_still_aligns():
    gt = _pl_rows()
    ex = [list(r) for r in gt]
    ex[3][1] = 601_000  # Gross profit typo'd
    result = score_table(gt, ex)
    assert 0.0 < result["score"] < 1.0
    gross_profit_row = result["rows"][3]
    assert gross_profit_row["gt"][0] == "Gross profit"
    assert gross_profit_row["row_score"] < 1.0


def test_score_table_completely_different_is_low():
    gt = _pl_rows()
    ex = [["Total assets", 9_999_999, 8_888_888]]
    result = score_table(gt, ex)
    assert result["score"] < 0.3


def test_score_table_empty_vs_empty_is_perfect():
    result = score_table([], [])
    assert result["score"] == 1.0
