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
    _Region, _split_glued_cell, _normalize_row_width, _merge_stacked,
    _drop_prose_columns, rows_in_box, HAVE_IMG2TABLE,
    _unscramble_reversed_paren_wrap, _merge_side_by_side,
    _looks_like_two_fused_statements, _SOCE_ANCHOR_RE,
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
