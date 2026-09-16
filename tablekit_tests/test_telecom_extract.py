"""
Tests for telecom_extract.py's company-vocabulary fuzzy-matching system
(PROFILES / find_best_matching_table / the new locate-only
find_candidate_locations) -- previously entirely uncovered. These matter for
two reasons: the refactor that split find_best_matching_table into
gather_candidate_tables / _evaluate_candidate / _candidate_is_valid_shape
must be provably behavior-preserving (the regression test below locks in
its real, pre-refactor output), and find_candidate_locations is what
serve.py's /api/telecom_candidates now serves to the web UI.

Everything that needs a real PDF skips cleanly when the sample report isn't
present, same pattern as test_golden.py/test_manual_mode.py.
"""
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
import telecom_extract as te  # noqa: E402

DU_2020 = ROOT / "du annual 2020.pdf"
DU_2025 = ROOT / "du annual 2025.pdf"
EAND_2025 = ROOT / "eand-integrated-annualreporten 2025.pdf"


# ------------------------------------------------------- strict matching ---
def test_strict_profile_for_file_matches_known_companies():
    assert te.strict_profile_for_file(Path("du annual 2025.pdf"))[0] == "du"
    assert te.strict_profile_for_file(Path("eand-integrated-annualreporten 2025.pdf"))[0] == "etisalat"


def test_strict_profile_for_file_has_no_fallback_for_an_unrelated_file():
    """The concrete test for "this feature must never run on a file that
    isn't actually a du/Etisalat report" -- unlike profile_for_file (which
    always returns something, defaulting to Etisalat), the strict version
    used to gate serve.py's telecom_candidates() must return (None, None)."""
    assert te.strict_profile_for_file(Path("Microsoft 2023 Annual Report.pdf")) == (None, None)
    assert te.strict_profile_for_file(Path("Walmart 1994 Annual Report.pdf")) == (None, None)


def test_profile_for_file_still_falls_back_for_the_cli():
    """profile_for_file's own fallback behavior (used by telecom_extract.py's
    CLI, not by the new locate-only feature) must be unchanged by the
    strict_profile_for_file extraction."""
    assert te.profile_for_file(Path("Microsoft 2023 Annual Report.pdf"))[0] == "etisalat"
    assert te.profile_for_file(Path("some du report.pdf"))[0] == "du"


# --------------------------------------- behavior-preservation regression ---
@pytest.mark.skipif(not DU_2020.exists(), reason="sample PDF not present")
def test_find_best_matching_table_unchanged_on_du_2020():
    """Locks in find_best_matching_table's real, known-good output on a real
    file -- the strongest available proof that splitting it into
    gather_candidate_tables/_evaluate_candidate/_candidate_is_valid_shape
    didn't change what it returns. Page/score baseline taken from
    telecom_extract.py's own committed du_two_tables.xlsx output."""
    import pdfplumber
    with pdfplumber.open(DU_2020) as pdf:
        idx = range(len(pdf.pages))
        header, body, page, score, heading, rec, note_col = \
            te.find_best_matching_table(pdf, idx, te.DU_PL)
        assert page == 71
        assert rec and rec["ok"]

        header, body, page, score, heading, rec, note_col = \
            te.find_best_matching_table(pdf, idx, te.DU_NOTE)
        assert page == 100
        assert rec and rec["ok"]


# --------------------------------------------------- find_candidate_locations
@pytest.mark.skipif(not DU_2020.exists(), reason="sample PDF not present")
def test_find_candidate_locations_finds_both_known_tables_on_du_2020():
    import pdfplumber
    with pdfplumber.open(DU_2020) as pdf:
        idx = range(len(pdf.pages))
        pl = te.find_candidate_locations(pdf, idx, te.DU_PL)
        note = te.find_candidate_locations(pdf, idx, te.DU_NOTE)
    assert pl and pl[0]["page"] == 71 and pl[0]["reconciles"] is True
    assert note and note[0]["page"] == 100 and note[0]["reconciles"] is True
    for cands in (pl, note):
        for c in cands:
            x0, y0, x1, y1 = c["bbox"]
            assert x1 > x0 and y1 > y0
            assert set(c) == {"page", "bbox", "score", "reconciles", "anchors_ok", "confirmed"}


@pytest.mark.skipif(not EAND_2025.exists(), reason="sample PDF not present")
def test_find_candidate_locations_finds_both_known_tables_on_eand_2025():
    import pdfplumber
    with pdfplumber.open(EAND_2025) as pdf:
        idx = range(len(pdf.pages))
        pl = te.find_candidate_locations(pdf, idx, te.ETI_PL)
        note = te.find_candidate_locations(pdf, idx, te.ETI_NOTE)
    assert pl and pl[0]["page"] == 181 and pl[0]["reconciles"] is True
    assert note and note[0]["page"] == 201 and note[0]["reconciles"] is True


@pytest.mark.skipif(not DU_2025.exists(), reason="sample PDF not present")
def test_du_2025_has_a_pl_candidate_but_legitimately_no_note_candidate():
    """du annual 2025.pdf restructured its P&L to print the expense
    breakdown on the statement's own face, with no separate operating-
    expenses note that year. An empty result here is documented, correct
    behavior, not a bug to "fix" by loosening the matcher later."""
    import pdfplumber
    with pdfplumber.open(DU_2025) as pdf:
        idx = range(len(pdf.pages))
        pl = te.find_candidate_locations(pdf, idx, te.DU_PL)
        note = te.find_candidate_locations(pdf, idx, te.DU_NOTE)
    assert pl and pl[0]["page"] == 133 and pl[0]["reconciles"] is True
    assert note == []


# --------------------------------------------------------- note-ref column ---
def test_strip_note_refs_returns_the_map_instead_of_a_shared_global():
    """_strip_note_refs used to report the column it removed through a
    module-level global (_DROPPED_NOTE_COL) that every caller read back
    immediately after calling it. Now it's a plain return value -- pin the
    shape so nothing reintroduces the global.

    Body needs >= 4 rows: that's what triggers the whole-column-removal
    pass that collapses a note-bearing row's stray placeholder cell back
    down to the same width as its note-free neighbours (real statements
    are always this size; a shorter body is a separate, narrower code
    path not exercised here)."""
    body = [
        ["Revenue", 100, 90],
        ["Cost of sales", -40, -35],
        ["General and administrative expenses", None, "19 (30) (25)"],
        ["Profit for the year", 30, 30],
    ]
    rows, note_col = te._strip_note_refs(body)
    assert isinstance(note_col, dict)
    assert note_col.get("general and administrative expenses") == "19"
    ga = next(r for r in rows if r[0] == "General and administrative expenses")
    assert ga == ["General and administrative expenses", -30, -25]
    assert {len(r) for r in rows} == {3}      # every row the same width


def test_strip_note_refs_strips_a_trailing_note_number_from_the_label():
    """Same bug, the OTHER place a note ref lands: glued onto the end of
    the label itself ("Finance income 21"), not a separate cell."""
    body = [
        ["Revenue", 100, 90],
        ["Finance income 21", 5, 4],
        ["Profit for the year", 70, 65],
    ]
    rows, note_col = te._strip_note_refs(body)
    assert rows[1][0] == "Finance income"
    assert note_col.get("finance income") == "21"


def test_strip_note_refs_strips_a_comma_separated_multi_note_reference():
    """A line item can cite more than one note ("28,34" instead of a bare
    "28") -- found live gluing onto the label on both en-2021-etisalat-
    group-annual-report.pdf p61 and en-2022-1-eand-group-annual-report.pdf
    p48, both via the pdfplumber-fallback path's _attach_left_labels,
    which glues together whatever words sit in the label band without
    knowing which one is a note ref."""
    body = [
        ["Gain / (loss) on net investment hedge 28,34", 782797, -720856.0],
        ["Profit for the year", 11059489, 10315736],
    ]
    rows, note_col = te._strip_note_refs(body)
    assert rows[0][0] == "Gain / (loss) on net investment hedge"
    assert note_col.get("gain / (loss) on net investment hedge") == "28,34"


def test_strip_note_refs_reunites_one_figure_split_by_a_stray_gap():
    """A single figure whose digits got split by a stray internal
    newline/space (img2table's own cell-text assembly, not the source PDF)
    can coincidentally match the OLDER 'two glued figures in one cell'
    shape this function already handled ('(2,167,933) (2,153,590)') --
    '1,1\\n12,374' for a source '1,112,374' splits into '1,1' and '12,374',
    both individually digit/comma runs. Found live on en-2021-etisalat-
    group-annual-report.pdf p61's P&L: the naive two-figure split used to
    silently fabricate 11 and 12374 out of one real figure. Must come back
    as ONE reunited number, not two fabricated ones."""
    body = [
        ["Finance and other income", 1289120, "1,1\n12,374"],
        ["Finance and other costs", -1284136, -2361052],
    ]
    rows, note_col = te._strip_note_refs(body)
    assert rows[0] == ["Finance and other income", 1289120, 1112374]


def test_strip_note_refs_keeps_a_bare_dash_separate_from_the_figure_after_it():
    """The ORIGINAL, unrelated case the two-glued-figures check exists for:
    a nil placeholder immediately followed by a separate real figure
    ('- 120,172', meaning no figure for one year and 120,172 for the
    other) -- found live on etisalat-group-annual-report-english-2019.pdf
    p48. This must NOT be reunited into one number the way the stray-gap
    case above is: parse_number reads a leading bare '-' as a genuine
    minus sign (correctly, for every other caller), so treating the whole
    string as one number here would silently flip a legitimate 'nil,
    positive value' pair into a single fabricated negative value and lose
    the nil marker and the row's real column count."""
    body = [
        ["Net cash inflow on disposal of subsidiary and associate", "- 120,172"],
        ["Proceeds from disposal of property, plant and equipment", 87415, 87692],
    ]
    rows, note_col = te._strip_note_refs(body)
    # _strip_note_refs itself leaves a bare dash as the literal string "-"
    # (the same spacer clean_cell always returns for one) -- normalizing
    # it to None is _clean()'s job, downstream of this function.
    assert rows[0] == ["Net cash inflow on disposal of subsidiary and associate", "-", 120172]


def test_strip_note_refs_keeps_a_unique_bare_year_header_row():
    """Found live on du annual 2012.pdf's income statement: a completely
    normal 2-line header (the date-range title carrying the years on one
    line, a separate Note/units line below it) had its year line dropped
    outright, losing "2012"/"2011" with no other row repeating them
    anywhere. _is_bare_year_row's own "no real label" guard used to call
    row_label(r) -- which returns the first non-empty STRING cell
    ANYWHERE in the row, with no concept of "a real label" versus "a bare
    year string" -- so on [None, None, "2012", "2011"] (img2table's raw,
    still-string cells) it returned "2012" itself, defeating the guard for
    exactly the row shape it exists to protect. A bare-year row is only
    ever supposed to be dropped when it's a genuine DUPLICATE of another
    row's years (the original stray-extra-copy bug) -- a unique one must
    survive."""
    body = [
        [None, None, "2012", "2011"],
        ["For the year ended 31 December", "Note", "AED 000", "AED 000"],
        ["Revenue", "29", "9,841,516", "8,854,683"],
    ]
    rows, note_col = te._strip_note_refs(body)
    # (the whole-column note-ref removal is a separate mechanism, tested
    # elsewhere -- it needs more rows than this minimal fixture to trigger
    # its own ratio threshold, so column 1 survives untouched here)
    assert rows[0] == [None, None, "2012", "2011"]
    assert rows[1] == ["For the year ended 31 December", "Note", "AED 000", "AED 000"]


def test_strip_note_refs_still_drops_an_exactly_duplicated_bare_year_row():
    """The original motivating case for the bare-year-row filter must
    still work: when the SAME bare-year row is genuinely duplicated
    (identical years, both carrying nothing else), the redundant copies
    are dropped rather than both kept."""
    body = [
        [None, None, "2012", "2011"],
        [None, None, "2012", "2011"],   # stray duplicate
        ["Revenue", "29", "9,841,516", "8,854,683"],
    ]
    rows, note_col = te._strip_note_refs(body)
    year_only_rows = [r for r in rows if r[0] is None]
    assert len(year_only_rows) < 2   # the duplicate pair doesn't both survive


def test_strip_note_refs_is_safe_under_concurrent_calls():
    """Regression for a real bug found live: the server is multi-threaded
    (ThreadingHTTPServer), and telecom_extract._strip_note_refs used to
    report its result through a module-level global it cleared at the top
    of every call. serve.py's slow, multi-page quickfind/telecom-candidate
    background scan and a user's own manual extraction both called this
    function -- on different threads, for different tables -- and one
    call's global.clear() could stomp another's read-back mid-flight,
    intermittently dropping the Notes column on a perfectly good
    extraction. Hammer it from many threads with DIFFERENT note refs per
    body and assert every thread gets back exactly its own -- impossible
    to guarantee with a shared global, trivially true with a return value."""
    import threading

    def body_for(i, tag):
        # the label deliberately does NOT end in a bare number (it'd
        # collide with the trailing-note-ref-in-the-label stripper tested
        # above) -- "thread{i}" keeps it unique per thread without that
        return [
            ["Revenue", 100, 90],
            [f"Expense line thread{i}", None, f"{tag} (30) (25)"],
            ["Profit for the year", 70, 65],
        ]

    errors = []

    def worker(i):
        try:
            for _ in range(50):
                tag = 10 + i    # a distinct small "note ref" per thread
                rows, note_col = te._strip_note_refs(body_for(i, tag))
                got = note_col.get(f"expense line thread{i}")
                if got != str(tag):
                    errors.append((i, tag, got, dict(note_col)))
        except Exception as e:            # pragma: no cover -- diagnostic only
            errors.append((i, "exception", repr(e), None))

    threads = [threading.Thread(target=worker, args=(i,)) for i in range(16)]
    for t in threads:
        t.start()
    for t in threads:
        t.join()

    assert not errors, f"cross-thread contamination: {errors[:5]}"


def test_strip_note_refs_whole_column_removal_not_diluted_by_a_notes_header_row():
    """Found live on du annual 2020's balance sheet: a table with BOTH a
    "Note[s]" column-label row AND a units row ("AED 000") pushed a
    genuinely-all-note-ref column (6, 7, 8, ...) to 12/14 = 0.857, just
    under the whole-column-removal's 0.9 bar, leaving the column
    un-stripped -- every real data row's note ref then got compared to
    its OWN figures as if it were a third year, tanking figure_health."""
    body = [
        ["Assets", "Notes", 2020, 2019],
        ["Assets", "Notes", "AED 000", "AED 000"],
        ["Property, plant and equipment", 6, 8063422, 7741119],
        ["Right-of-use assets", 7, 1851429, 1699651],
        ["Intangible assets and goodwill", 8, 900215, 1051446],
        ["Contract assets", 13, 211216, 208994],
        ["Trade and other receivables", 14, 1726401, 1870556],
        ["Due from related parties", 15, 139869, 164995],
        ["Term deposits", 16, 2029327, 2948701],
        ["Cash and bank balances", 17, 213375, 268695],
    ]
    rows, note_col = te._strip_note_refs(body)
    assert note_col.get("property, plant and equipment") == "6"
    ppe = next(r for r in rows if r[0] == "Property, plant and equipment")
    assert ppe == ["Property, plant and equipment", 8063422, 7741119]
