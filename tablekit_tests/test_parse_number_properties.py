"""Property-based tests for tablekit/parse.py's parse_number / coerce_cell.

The existing parametrized cases in test_golden.py::test_parse_number lock in
specific known formats by example. These complement that with properties that
should hold across a much wider space of inputs than anyone would think to
hand-write -- in particular "never crashes," which matters here specifically
because both functions run on OCR/PDF-extracted text this tool never controls
(a garbled scan, a stray Unicode character, an absurd digit run) and a crash
on one cell would take down the whole extraction, not just that cell.
"""
import sys
from pathlib import Path

from hypothesis import given, settings, strategies as st

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
from tablekit.parse import coerce_cell, parse_number  # noqa: E402


@given(st.text())
@settings(max_examples=500)
def test_parse_number_never_raises_on_arbitrary_text(s):
    parse_number(s)


@given(st.text())
@settings(max_examples=500)
def test_coerce_cell_never_raises_on_arbitrary_text(s):
    coerce_cell(s)


@given(st.integers(min_value=-10**15, max_value=10**15))
def test_round_trips_through_thousands_formatting(n):
    assert parse_number(f"{n:,}") == n


@given(st.integers(min_value=1, max_value=10**12))
def test_parens_negate_a_plain_positive_integer(n):
    plain = parse_number(f"{n:,}")
    assert parse_number(f"({n:,})") == -plain


@given(st.integers(min_value=1, max_value=10**12))
def test_percent_suffix_does_not_change_the_numeric_value(n):
    assert float(parse_number(f"{n:,}%")) == float(parse_number(f"{n:,}"))


@given(st.integers(min_value=1, max_value=10**12))
def test_cr_negates_relative_to_dr(n):
    assert parse_number(f"{n:,} CR") == -parse_number(f"{n:,} DR")


@given(st.integers(min_value=-10**12, max_value=10**12))
def test_coerce_cell_agrees_with_parse_number_on_plain_figures(n):
    # coerce_cell's NUM_RE fast path and parse_number's full parse are two
    # separately-maintained code paths meant to agree on anything simple
    # enough to hit the fast path -- this is the cross-check that a future
    # edit to one, but not the other, can't silently break without a test
    # noticing (property-based rather than a fixed example list specifically
    # because the failure mode here is "someone's edit shifts behavior for
    # some inputs but not others," which a handful of hand-picked examples
    # could easily miss).
    plain = f"{n:,}"
    assert coerce_cell(plain) == parse_number(plain)
    paren = f"({abs(n):,})"
    assert coerce_cell(paren) == parse_number(paren)
    pct = f"{n:,}%"
    assert float(coerce_cell(pct)) == float(parse_number(pct))
