"""
Regression tests for extract_all_tables.py

`golden.json` is a snapshot of the detector's output on 14 real annual reports
(42 primary statements).  Any code change that alters a detected statement's
kind / years / foots verdict / shape / row values will fail here -- so the
change has to be deliberate.  Regenerate the snapshot with:

    python tablekit_tests/snapshot.py

Run:  pytest tablekit_tests/ -q      (from the project root)
"""
import json
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
import extract_all_tables as X  # noqa: E402

GOLDEN = json.loads((Path(__file__).parent / "golden.json").read_text(encoding="utf-8"))

CASE_PAGES = {
    "du annual 2010.pdf": "14-18",
    "du annual 2013.pdf": "24-27",
    "du annual 2016.pdf": "22-27",
    "du annual 2019.pdf": "70-74",
    "du annual 2020.pdf": "69-73",
    "du annual 2021.pdf": "86-91",
    "du annual 2024.pdf": "91-96",
    "du annual 2025.pdf": "130-136",
    "etisalat-group-annual-report-english-2018.pdf": "41-46",
    "etisalat-group-annual-report-english-2019.pdf": "44-49",
    "en-2020-etisalat-group-annual-report.pdf": "57-62",
    "en-2021-etisalat-group-annual-report.pdf": "59-64",
    "en-2022-1-eand-group-annual-report.pdf": "46-52",
    "integrated-reporten-2023.pdf": "105-112",
}
STMT = ("income statement", "statement of financial position",
        "statement of cash flows", "statement of changes in equity")


def _prange(s):
    a, _, b = s.partition("-")
    return range(int(a) - 1, (int(b) if b else int(a)))


def _series(t):
    vc = t.get("value_cols") or []
    data = t["rows"][t.get("data_start", 0):]
    col = vc[0] if vc else None
    return [[X._row_label(r),
             (r[col] if (col is not None and col < len(r)
                         and isinstance(r[col], (int, float))) else None)]
            for r in data if X._row_label(r)]


def _scan_case(fname):
    p = ROOT / fname
    if not p.exists():
        pytest.skip(f"{fname} not present")
    out = {}
    for t in X.scan([p], _prange(CASE_PAGES[fname]), 2, 2, warn=lambda *a: None):
        if t["kind"] in STMT:
            out[(t["page_label"], t["kind"])] = t
    return out


@pytest.mark.parametrize("fname", sorted(CASE_PAGES))
def test_statements_match_snapshot(fname):
    got = _scan_case(fname)
    expected = [g for g in GOLDEN if g["file"] == fname]
    assert expected, f"no golden rows for {fname} -- regenerate snapshot"
    for g in expected:
        k = (g["page"], g["kind"])
        assert k in got, f"{fname}: lost {g['kind']} on p{g['page']}"
        t = got[k]
        assert t.get("years") == g["years"], f"{fname} p{g['page']}: years"
        assert t.get("foots") == g["foots"], f"{fname} p{g['page']}: foots verdict"
        assert len(t["rows"]) == g["nrows"], f"{fname} p{g['page']}: row count"
        assert _series(t) == g["series"], f"{fname} p{g['page']}: row values/labels"
    # nothing spurious appeared
    assert len(got) == len(expected), (
        f"{fname}: detected {len(got)} statements, snapshot has {len(expected)}")


def test_every_snapshot_statement_foots_except_known():
    """41 of 42 snapshot statements reconcile; the 2 du old-2-up cash flows
    are the only permitted blanks."""
    blanks = [g for g in GOLDEN if g["foots"] is None]
    for g in blanks:
        assert g["kind"] == "statement of cash flows"
        assert g["file"] in ("du annual 2013.pdf", "du annual 2012.pdf")
    assert sum(1 for g in GOLDEN if g["foots"] is True) >= 40


@pytest.mark.parametrize("raw,expected", [
    ("1,234", 1234), ("1,234.56", 1234.56), ("(1,234)", -1234), ("-1,234", -1234),
    ("1,234-", -1234), ("1 234", 1234), ("1 234 567", 1234567),
    ("1.234,56", 1234.56), ("1.234.567,89", 1234567.89),
    ("12.5%", 12.5), ("$1,234", 1234), ("AED 1,234", 1234),
    ("1,234 CR", -1234), ("1,234 DR", 1234), ("1,234*", 1234),
    ("1,234¹", 1234), ("1,234 (a)", 1234), ("1,234,567", 1234567),
    ("0.43", 0.43), ("(10,448)", -10448), ("  2,860  ", 2860), ("2020", 2020),
    ("-", None), ("–", None), ("—", None), ("nil", None),
    ("n/a", None), ("", None), ("abc", None), ("10.1 -", None),
    ("1" * 400, None), ("9" * 30, None), ("1e999", None), (".", None),
    ("(1,234", None), ("1,2,3,4", 1234),
    # reversed parens -- a bidi text-ordering artifact seen in real PDFs
    # (e.g. du annual 2019.pdf, en-2022-1-eand-group-annual-report.pdf):
    # ")1,234(" instead of "(1,234)". See tablekit/parse.py::parse_number.
    (")1,234(", -1234), (")87,579(", -87579),
    # a step further than the swap above: BOTH reversed parens carried all
    # the way to the front instead of one on each side -- ") (417,358" for
    # a source "(417,358)". img2table's own cell-text assembly (not
    # pdfplumber's), found live on en-2022-1-eand-group-annual-report.pdf
    # p50's cash flow statement.
    (") (417,358", -417358), (") (297,462", -297462),
    # a number that's already comma-grouped picking up a stray extra space
    # right after one of its own commas -- img2table's own cell-text
    # assembly again, found live on the same file/page: pdfplumber's own
    # extract_words() reads the same spot as the ordinary, ungapped
    # "11,180,517".
    ("11, 180,517", 11180517), ("1, 234,567", 1234567),
    # the stray gap doesn't have to land right after a comma -- "1,1\n12,374"
    # for a source "1,112,374" splits mid-group instead (found live on
    # en-2021-etisalat-group-annual-report.pdf p61); concatenating with no
    # separator still recovers it as long as that concatenation is exactly
    # one strictly comma-grouped number.
    ("1,1\n12,374", 1112374), ("1,1 12,374", 1112374),
    # several distinct numbers that ended up in the same cell (a wrapped
    # multi-line cell flattened to spaces by normspace, or a row-joining
    # artifact upstream) must never be silently glued into one digit blob --
    # found live via the comprehensive audit as "2020202020192019" and an
    # 18-digit "193881930915196032".
    ("2020 2020 2019 2019", None), ("94,374 2,444,051 27,481", None),
    # the tightened absurd-digit-run guard (18 -> 16 digits) must still let
    # the existing +-10**15 round-trip contract through
    ("1,000,000,000,000,000", 1000000000000000),
])
def test_parse_number(raw, expected):
    assert X.parse_number(raw) == expected


def test_equity_foots_excludes_subtotal_rows_from_the_movement_sum():
    """Found live (du annual 2018.pdf p95): a hand-verified-correct equity
    statement -- every row's own columns summed to its row total, and the
    full opening-to-closing roll-forward balanced exactly -- was flagged as
    NOT footing. _equity_foots sums the movement rows between an opening and
    closing balance but wasn't excluding "Total ..." subtotal rows (e.g.
    "Total comprehensive income"), double-counting them against the rows
    they're already subtotals of."""
    header = [None, "Total"]
    data = [
        header,
        ["At 1 January 2018", 8342576],
        ["Profit for the year", 1752992],
        ["Other comprehensive income", 2687],
        ["Total comprehensive income", 1755679],   # == sum of the two rows above
        ["Cash dividends paid", -1586517],
        ["Total transactions with shareholders", -1586517],  # == the row above
        ["At 31 December 2018", 8511738],   # 8342576 + 1752992 + 2687 - 1586517
    ]
    assert X._equity_foots(data, 1) is True


def test_label_health_runs_and_scores():
    t = _scan_case("du annual 2020.pdf")[(71, "income statement")]
    h = X.label_health(t)
    assert 0.0 <= h["score"] <= 1.0
    assert "suspect" in h and isinstance(h["suspect"], list)


def test_reconcile_explain_localises_a_break():
    seq = [("Revenue", 1000), ("Cost of sales", -300), ("Other", -100),
           ("Profit for the year", 600)]
    assert X.reconcile_explain(seq)["ok"] is True
    bad = [("Revenue", 1000), ("Cost of sales", -300), ("Other", -100),
           ("Profit for the year", 800)]  # off by 200
    r = X.reconcile_explain(bad)
    assert r["ok"] is False
    assert abs(r["gap"]) == 200


def test_reconcile_explain_points_at_the_broken_row():
    """A row extracted twice is localisable; name it."""
    good = [("Revenue", 11083845), ("Operating expenses", -8253903),
            ("Expected credit losses", -227353), ("Other income", 13904),
            ("Federal royalty", -1511938), ("Finance income", 50575),
            ("Finance costs", -105859), ("Impairment of goodwill", -135830),
            ("Gain on disposal", 519374), ("Share of profit", 10099),
            ("Profit for the year", 1442914)]
    assert X.reconcile_explain(good)["ok"] is True
    doubled = good[:5] + [("Federal royalty", -1511938)] + good[5:]  # counted twice
    r = X.reconcile_explain(doubled)
    assert r["ok"] is False
    assert r["break_label"] == "Federal royalty"

    # a single mis-read digit is NOT confidently localisable -- don't pretend
    misread = [(l, (v - 100000 if l == "Federal royalty" else v)) for l, v in good]
    r2 = X.reconcile_explain(misread)
    assert r2["ok"] is False and abs(r2["gap"]) == 100000


# --------------------------------------------------------------------------- #
# Independent ground-truth anchors: figures read by hand from the raw PDF text
# --------------------------------------------------------------------------- #
_ANCHORS = json.loads(
    (Path(__file__).parent / "anchors.json").read_text(encoding="utf-8"))["anchors"]
_ANCHOR_FILES = sorted({a["file"] for a in _ANCHORS})


@pytest.mark.parametrize("fname", _ANCHOR_FILES)
def test_hand_verified_figures(fname):
    import re as _re
    got = _scan_case(fname) if fname in CASE_PAGES else None
    if got is None:
        p = ROOT / fname
        if not p.exists():
            pytest.skip(f"{fname} not present")
        got = {}
        for t in X.scan([p], None, 2, 2, warn=lambda *a: None):
            if t["kind"] in STMT:
                got[(t["page_label"], t["kind"])] = t
    for a in [x for x in _ANCHORS if x["file"] == fname]:
        t = got.get((a["page"], a["kind"]))
        assert t is not None, f"{fname}: no {a['kind']} detected on p{a['page']}"
        vc = t.get("value_cols") or []
        assert a["col"] < len(vc), f"{fname} p{a['page']}: only {len(vc)} value cols"
        col = vc[a["col"]]
        rx = _re.compile(a["label"], _re.I)
        hit = None
        for r in t["rows"][t.get("data_start", 0):]:
            lb = X._row_label(r)
            if lb and rx.search(lb) and col < len(r) and isinstance(r[col], (int, float)):
                hit = r[col]
        assert hit is not None, (
            f"{fname} p{a['page']}: no row matching /{a['label']}/ with a value in col {a['col']}")
        assert hit == a["value"], (
            f"{fname} p{a['page']} /{a['label']}/ col{a['col']}: "
            f"extractor says {hit:,}, PDF says {a['value']:,}")
