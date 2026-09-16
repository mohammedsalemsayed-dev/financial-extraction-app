"""
Tests for serve.py -- the preview / edit / export backend.

The pure-logic paths (edit application, inventory shape, cross-year wiring,
one-workbook export) run with a monkey-patched scanner and need no PDF.
One end-to-end HTTP test runs only if a sample PDF is present.
"""
import io
import json
import logging
import sys
import threading
import urllib.request
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
import extract_all_tables as X       # noqa: E402
import serve                          # noqa: E402


def _pl(years, rev=1_000_000):
    return [["", years[0], years[1]],
            ["Revenue", rev, rev - 100_000],
            ["Operating expenses", -600_000, -560_000],
            ["Finance costs", -50_000, -40_000],
            ["Profit before tax", rev - 650_000, rev - 700_000],
            ["Income tax", -50_000, -40_000],
            ["Profit for the year", rev - 700_000, rev - 740_000]]


def _make_tables():
    out = []
    for pg, kind_title, rows in [
        (5, "Statement of profit or loss", _pl([2024, 2023])),
    ]:
        t = {"rows": rows, "title": kind_title, "file": "demo.pdf",
             "page_label": pg, "page": pg, "shape": X.classify(rows),
             "bbox": [10, 10, 300, 400], "page_size": [595, 842]}
        X.analyze(t)
        X._attach_health(t)
        out.append(t)
    return out


@pytest.fixture
def wired(monkeypatch, tmp_path):
    pdf = tmp_path / "demo.pdf"
    pdf.write_bytes(b"%PDF-1.4 fake")            # never actually opened (scan is patched)
    monkeypatch.setattr(serve, "_state",
                        {"files": [pdf], "pages": None, "scans": {}, "pngs": {}})
    monkeypatch.setattr(X, "scan", lambda *a, **k: _make_tables())
    return pdf


def test_inventory_shape(wired):
    serve.scan_file("demo.pdf")   # inventory() itself no longer auto-scans
    inv = serve.inventory("demo.pdf")
    assert inv["file"] == "demo.pdf"
    assert len(inv["tables"]) == 1
    row = inv["tables"][0]
    for key in ("n", "page", "kind", "foots", "foot_by_col", "health",
                "consistency", "group", "n_suspect"):
        assert key in row
    assert row["kind"] == "income statement"
    assert row["group"] == "Financial statements"
    assert row["foots"] is True


def test_table_detail_carries_rows_and_verdicts(wired):
    d = serve.table_detail("demo.pdf", 1)
    assert d["rows"][1][0] == "Revenue"
    assert len(d["foot_by_col"]) == 2
    assert "value_cols" in d and d["value_cols"] == [1, 2]
    assert "notes_i18n" in d   # webui.html's per-note translation channel


def test_apply_edits_relabels_and_reanalyses(wired):
    tables = serve.scan_file("demo.pdf")
    rows = [r[:] for r in tables[0]["rows"]]
    rows[1][0] = "Turnover"                       # rename Revenue
    edited = serve._apply_edits(tables, {"1": {"rows": rows, "title": "My P&L"}})
    assert edited[0]["_edited"] is True
    assert edited[0]["rows"][1][0] == "Turnover"
    assert edited[0]["title"] == "My P&L"
    assert edited[0]["foots"] is True            # still reconciles after the edit


def test_apply_edits_parses_string_numbers_server_side(wired):
    tables = serve.scan_file("demo.pdf")
    rows = [r[:] for r in tables[0]["rows"]]
    rows[1][1] = "(1,234,567)"                    # user typed an accounting negative
    edited = serve._apply_edits(tables, {"1": {"rows": rows}})
    assert edited[0]["rows"][1][1] == -1234567    # parsed by extract_all_tables._cell


def test_export_is_one_workbook_with_a_sheet_per_table(wired):
    import openpyxl
    data, fn = serve.export_xlsx("demo.pdf", [1])
    wb = openpyxl.load_workbook(io.BytesIO(data))
    assert "Contents" in wb.sheetnames
    assert len([s for s in wb.sheetnames if s != "Contents"]) == 1
    assert fn.endswith(".xlsx")


def test_export_applies_edits(wired):
    import openpyxl
    tables = serve.scan_file("demo.pdf")
    rows = [r[:] for r in tables[0]["rows"]]
    rows[1][0] = "EDITED"
    data, _ = serve.export_xlsx("demo.pdf", [1], {"1": {"rows": rows, "title": "Custom"}})
    wb = openpyxl.load_workbook(io.BytesIO(data))
    assert "Custom" in wb.sheetnames
    col_a = [row[0] for row in wb["Custom"].iter_rows(values_only=True)]
    assert "EDITED" in col_a


def test_reanalyze_recomputes_verdict_after_a_bad_edit(wired):
    d0 = serve.table_detail("demo.pdf", 1)
    assert d0["foots"] is True
    rows = [r[:] for r in d0["rows"]]
    rows[1][1] = "1"                              # wreck Revenue in the current year
    d1 = serve.reanalyze("demo.pdf", 1, rows)
    assert d1["foot_by_col"][0]["ok"] is False    # current-year column now breaks
    assert d1["foot_by_col"][1]["ok"] is True     # prior year untouched


def test_reanalyze_parses_cells_server_side(wired):
    d0 = serve.table_detail("demo.pdf", 1)
    rows = [r[:] for r in d0["rows"]]
    rows[2][1] = "(600,000)"                      # typed as an accounting negative string
    d1 = serve.reanalyze("demo.pdf", 1, rows)
    assert d1["rows"][2][1] == -600000


def test_edit_one_does_not_accumulate_stale_notes_i18n_across_re_edits():
    # notes_i18n must be popped in lockstep with notes before re-analysing
    # (see the comment above the pop in _edit_one) -- otherwise a note from
    # a PREVIOUS edit lingers in notes_i18n after "notes" itself was
    # correctly rebuilt, and the two lists -- meant to be paired by index --
    # drift out of sync.
    t = _make_tables()[0]
    t["notes"] = ["stale note from a previous edit"]
    t["notes_i18n"] = [{"key": "someStaleKey", "vars": {}}]
    rows = [r[:] for r in t["rows"]]
    t2 = serve._edit_one(t, {"rows": rows})
    assert "stale note from a previous edit" not in (t2.get("notes") or [])
    assert not any(m.get("key") == "someStaleKey" for m in (t2.get("notes_i18n") or []))
    assert len(t2.get("notes") or []) == len(t2.get("notes_i18n") or [])


def test_compare_returns_a_diff(wired, monkeypatch):
    # two files, same statement, one figure restated. diff_tables() only
    # flags "restated" when A's PRIOR year lines up with B's CURRENT year
    # (comparing this year's report against last year's, the real use
    # case) -- so B must report the YEAR BEFORE A's, not the same pair, or
    # the restated branch can never fire no matter how different the figures.
    a = _make_tables()                                  # years [2024, 2023]
    b_rows = _pl([2023, 2022], rev=850_000)             # b's "2023" is a's "2023" restated
    b_table = {"rows": b_rows, "title": "Statement of profit or loss", "file": "demo.pdf",
               "page_label": 5, "page": 5, "shape": X.classify(b_rows),
               "bbox": [10, 10, 300, 400], "page_size": [595, 842]}
    X.analyze(b_table); X._attach_health(b_table)
    b = [b_table]
    files = {"a.pdf": a, "b.pdf": b}
    monkeypatch.setattr(serve, "_state",
                        {"files": [Path("a.pdf"), Path("b.pdf")],
                         "pages": None, "scans": {}, "pngs": {}})
    monkeypatch.setattr(serve, "scan_file", lambda name, force=False: files[name])
    out = serve.compare("a.pdf", 1, "b.pdf", 1)
    assert out["rows"][0][0] == "Line item"
    assert "changed" in out["verdict"] or "RESTATED" in out["verdict"]
    # `counts` is what the UI uses to build a TRANSLATED verdict sentence
    # instead of hard-coded English (webui.html's renderComparePanel) -- the
    # numbers behind it must agree with what the English verdict prose says
    for key in ("changed", "new", "removed", "restated"):
        assert key in out["counts"] and isinstance(out["counts"][key], int)
    assert out["counts"]["restated"] >= 1               # b's 2023 Revenue was restated
    assert f"{out['counts']['restated']} RESTATED" in out["verdict"]


# ---- session persistence (autosave / restart-survival / undo) -----------
def test_reanalyze_commits_the_edit_into_the_manual_list(wired):
    t = _make_tables()[0]
    serve._manual_list("demo.pdf").append(t)
    n = serve.MANUAL_OFFSET + 1
    d0 = serve.table_detail("demo.pdf", n)
    rows = [r[:] for r in d0["rows"]]
    rows[1][0] = "Turnover"
    serve.reanalyze("demo.pdf", n, rows)
    # NOT just returned to the caller -- actually committed server-side, so
    # a second, independent read (e.g. after a page refresh) sees the edit
    assert serve._manual_list("demo.pdf")[0]["rows"][1][0] == "Turnover"


def test_session_survives_a_simulated_restart(wired):
    t = _make_tables()[0]
    serve._manual_list("demo.pdf").append(t)
    serve._save_session("demo.pdf")
    # simulate a process restart: nothing left in memory for this file
    serve._state["manual"]["demo.pdf"] = []
    serve._state["deleted"]["demo.pdf"] = []
    serve._load_session("demo.pdf")
    restored = serve._manual_list("demo.pdf")
    assert len(restored) == 1
    assert restored[0]["title"] == t["title"]
    assert restored[0]["rows"] == t["rows"]


def test_delete_then_undelete_restores_the_same_table(wired):
    t = _make_tables()[0]
    serve._manual_list("demo.pdf").append(t)
    n = serve.MANUAL_OFFSET + 1
    serve.delete_manual("demo.pdf", n)
    assert serve._manual_list("demo.pdf")[0] is None
    with pytest.raises(KeyError):
        serve._resolve("demo.pdf", n)
    restored_n = serve.undelete_manual("demo.pdf")
    assert restored_n == n
    assert serve._resolve("demo.pdf", n)["title"] == t["title"]


def test_undelete_with_nothing_to_undo_returns_none(wired):
    assert serve.undelete_manual("demo.pdf") is None


def test_upload_pdf_keeps_a_unicode_filename(wired, monkeypatch, tmp_path):
    monkeypatch.setattr(serve, "UPLOAD_DIR", tmp_path / "uploads")
    import base64
    name = serve.upload_pdf("تقرير 2024.pdf", base64.b64encode(b"%PDF-1.4 fake").decode())
    assert name == "تقرير 2024.pdf"


def test_upload_pdf_still_strips_path_separators_and_reserved_chars(
        wired, monkeypatch, tmp_path):
    monkeypatch.setattr(serve, "UPLOAD_DIR", tmp_path / "uploads")
    import base64
    name = serve.upload_pdf("a/b\\c:d*e?.pdf", base64.b64encode(b"%PDF-1.4 fake").decode())
    assert "/" not in name and "\\" not in name and ":" not in name


def test_upload_pdf_renames_on_collision_with_an_already_loaded_file_not_in_upload_dir(
        wired, monkeypatch, tmp_path):
    """The old collision check was `dest.exists()` -- only catches a name
    already physically sitting in UPLOAD_DIR. Found live: a file loaded via
    a CLI arg (so it lives elsewhere, e.g. the repo root) doesn't physically
    exist AT the UPLOAD_DIR path, so uploading a NEW file with that same
    name slipped through undetected and produced the identical dead-entry
    problem the _discover_files fix above addresses -- this is the other
    half of the same bug, on the live-upload path instead of startup scan."""
    import base64
    uploads = tmp_path / "uploads"
    monkeypatch.setattr(serve, "UPLOAD_DIR", uploads)
    elsewhere = tmp_path / "elsewhere.pdf"
    elsewhere.write_bytes(b"%PDF-1.4 loaded from a CLI arg, not an upload")
    with serve._lock:
        serve._state["files"] = [elsewhere.with_name("report.pdf")]
    name = serve.upload_pdf("report.pdf", base64.b64encode(b"%PDF-1.4 new upload").decode())
    assert name != "report.pdf"
    assert name == "report_1.pdf"


def test_discover_files_rediscovers_previously_uploaded_pdfs(monkeypatch, tmp_path):
    uploads = tmp_path / "uploads"
    uploads.mkdir()
    old = uploads / "last session.pdf"
    old.write_bytes(b"%PDF-1.4 fake")
    monkeypatch.setattr(serve, "UPLOAD_DIR", uploads)
    files = serve._discover_files([])          # no CLI args -- the run_app.bat path
    assert old.resolve() in {f.resolve() for f in files}


def test_discover_files_does_not_duplicate_a_cli_arg_already_in_uploads(monkeypatch, tmp_path):
    uploads = tmp_path / "uploads"
    uploads.mkdir()
    pdf = uploads / "report.pdf"
    pdf.write_bytes(b"%PDF-1.4 fake")
    monkeypatch.setattr(serve, "UPLOAD_DIR", uploads)
    files = serve._discover_files([str(pdf)])
    assert len(files) == 1


def test_discover_files_dedupes_by_name_not_just_resolved_path(monkeypatch, tmp_path):
    """Found live: launching against a directory that has "report.pdf" while
    uploads/ ALSO has a "report.pdf" (a genuinely different file, e.g. from
    an earlier session's upload of the same-named report) listed both --
    _path()'s name-based lookup always resolves the first, so the second was
    a dead, confusing dropdown entry, not just a harmless extra. These two
    files are at different paths (unlike the identical-path case above), so
    the old resolved-path dedup didn't catch it; only a name-based one does."""
    cli_dir = tmp_path / "reports"
    cli_dir.mkdir()
    cli_pdf = cli_dir / "report.pdf"
    cli_pdf.write_bytes(b"%PDF-1.4 from the CLI-arg directory")
    uploads = tmp_path / "uploads"
    uploads.mkdir()
    (uploads / "report.pdf").write_bytes(b"%PDF-1.4 a DIFFERENT file, same name")
    monkeypatch.setattr(serve, "UPLOAD_DIR", uploads)
    files = serve._discover_files([str(cli_dir)])
    assert [f.name for f in files] == ["report.pdf"]     # not duplicated
    assert files[0] == cli_pdf                            # CLI-arg order wins


def test_free_port_returns_an_open_port():
    import socket
    p = serve._free_port(9000)
    with socket.socket() as s:
        s.bind(("127.0.0.1", p))                 # must be bindable == was free


# ---- --debug: the diagnostic capture path added after a "breaks after
# 2-3 runs" report that was never actually diagnosed (see CHANGELOG 0.7.1) --
def test_debug_state_line_does_not_crash_and_counts_manual_slots_not_just_live_ones(wired, caplog):
    # tombstoned (deleted) slots stay counted -- that's deliberate: a
    # count that only reflected LIVE tables would hide the exact kind of
    # growth (many delete/extract cycles leaving the internal list larger
    # than what the UI shows) this exists to make visible
    serve._manual_list("demo.pdf").extend([{"a": 1}, None, {"b": 2}])
    with caplog.at_level("DEBUG", logger="tablekit.serve"):
        serve._debug_state_line("/api/table")
    assert any("manual_tables=3" in r.message for r in caplog.records)


def test_cli_parses_debug_port_and_no_browser_flags(monkeypatch):
    captured = {}
    monkeypatch.setattr(serve, "run", lambda files, **kw: captured.update(kw, files=files))
    serve._cli(["report.pdf", "--debug", "--no-browser", "--port", "9999"])
    assert captured == {"files": ["report.pdf"], "port": 9999,
                        "open_browser": False, "debug": True, "allow_remote": False}


def test_cli_defaults_match_previous_behavior_with_no_flags(monkeypatch):
    # the old _cli() silently dropped every "--" flag and never passed
    # port/open_browser/debug at all -- these are the defaults that made
    # that accidentally look like it worked for the common no-flags case
    captured = {}
    monkeypatch.setattr(serve, "run", lambda files, **kw: captured.update(kw, files=files))
    serve._cli(["report.pdf"])
    assert captured == {"files": ["report.pdf"], "port": None,
                        "open_browser": True, "debug": False, "allow_remote": False}


def test_cli_parses_allow_remote_flag(monkeypatch):
    captured = {}
    monkeypatch.setattr(serve, "run", lambda files, **kw: captured.update(kw, files=files))
    serve._cli(["report.pdf", "--allow-remote"])
    assert captured["allow_remote"] is True


class _FakeServer:
    """Stands in for ThreadingHTTPServer in run() -- no real bind/serve, so
    the test doesn't hang on serve_forever() or open a real socket."""
    def __init__(self, *a, **k):
        pass

    def serve_forever(self):
        pass


def test_run_caps_pdfminers_own_debug_noise_when_debug_flag_is_on(monkeypatch, tmp_path):
    """logging.basicConfig's DEBUG level cascades to every logger without
    its own override -- including pdfminer (under pdfplumber), which logs
    every single parse token/seek/keyword. Found live: a 20-minute, 26-file
    stress session produced a 4.3 GB / 44.5-million-line debug.log, nearly
    all of it pdfminer's own byte-level parse trace, not this project's."""
    monkeypatch.setattr(serve, "ThreadingHTTPServer", _FakeServer)
    monkeypatch.setattr(serve, "ROOT", tmp_path)     # debug.log lands here, not the real repo
    pdfminer_logger = logging.getLogger("pdfminer")
    original_level = pdfminer_logger.level
    try:
        pdfminer_logger.setLevel(logging.NOTSET)
        serve.run([], open_browser=False, debug=True)
        assert pdfminer_logger.level == logging.WARNING
    finally:
        pdfminer_logger.setLevel(original_level)


def test_run_leaves_pdfminers_logger_alone_when_debug_is_off(monkeypatch, tmp_path):
    monkeypatch.setattr(serve, "ThreadingHTTPServer", _FakeServer)
    pdfminer_logger = logging.getLogger("pdfminer")
    original_level = pdfminer_logger.level
    try:
        pdfminer_logger.setLevel(logging.NOTSET)
        serve.run([], open_browser=False, debug=False)
        assert pdfminer_logger.level == logging.NOTSET      # untouched
    finally:
        pdfminer_logger.setLevel(original_level)


# ---- end-to-end HTTP (needs a real PDF) -----------------------------------
def _sample_pdf():
    for name in ("du annual 2020.pdf", "du annual 2016.pdf"):
        p = ROOT / name
        if p.exists():
            return p
    return None


@pytest.mark.skipif(_sample_pdf() is None, reason="no sample PDF present")
def test_http_end_to_end(monkeypatch):
    pdf = _sample_pdf()
    monkeypatch.setattr(serve, "_state",
                        {"files": [pdf], "pages": range(0, 80), "scans": {}, "pngs": {}})
    from http.server import ThreadingHTTPServer
    port = serve._free_port(9100)
    srv = ThreadingHTTPServer(("127.0.0.1", port), serve.Handler)
    th = threading.Thread(target=srv.serve_forever, daemon=True)
    th.start()
    try:
        base = f"http://127.0.0.1:{port}"
        files = json.load(urllib.request.urlopen(base + "/api/files"))
        assert pdf.name in files["files"]
        # the web UI has no detector to trigger any more (manual box-select
        # only) -- seed the scan directly, the same way the other serve.py
        # tests do, then read it back through the still-live /api/scan.
        serve.scan_file(pdf.name)
        inv = json.load(urllib.request.urlopen(
            base + "/api/scan?file=" + urllib.request.quote(pdf.name)))
        stmts = [t for t in inv["tables"] if t["group"] == "Financial statements"]
        assert stmts, "no statements detected in sample"
        n = stmts[0]["n"]
        png = urllib.request.urlopen(
            base + f"/api/page?file={urllib.request.quote(pdf.name)}&n={n}&scale=1.5").read()
        assert png[:8] == b"\x89PNG\r\n\x1a\n"
        body = json.dumps({"file": pdf.name, "ns": [t["n"] for t in stmts]}).encode()
        req = urllib.request.Request(base + "/api/export", data=body,
                                     headers={"Content-Type": "application/json"})
        xlsx = urllib.request.urlopen(req)
        assert xlsx.status == 200
        assert "attachment" in xlsx.headers.get("Content-Disposition", "")
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(xlsx.read()))
        assert len(wb.sheetnames) == len(stmts) + 1        # + Contents
    finally:
        srv.shutdown()


# --------------------------------------------------- telecom_candidates() ---
def test_telecom_candidates_never_opens_the_pdf_for_an_unrelated_file(monkeypatch):
    """The concrete test for "this must never run on a file that isn't
    actually a du/Etisalat report" -- the strict-match gate has to fail
    BEFORE the PDF is ever opened, not just before the result is returned."""
    import pdfplumber as pdfplumber_mod
    def _boom(*a, **k):
        raise AssertionError("pdfplumber.open should never be called for a non-matching file")
    monkeypatch.setattr(pdfplumber_mod, "open", _boom)
    monkeypatch.setattr(serve, "_state",
                        {"files": [Path("Microsoft 2023 Annual Report.pdf")],
                         "pages": None, "scans": {}, "pngs": {}, "pagetext": {}})
    result = serve.telecom_candidates("Microsoft 2023 Annual Report.pdf")
    assert result == {"available": False, "candidates": []}


def test_telecom_candidates_reuses_the_mtime_cache(monkeypatch):
    real_du_profile = X._te.PROFILES["du"]     # captured BEFORE X._te gets patched below
    calls = []
    def _fake_find(pdf, idx, target):
        calls.append(target["name"])
        return []
    fake_te = type("FakeTE", (), {
        "strict_profile_for_file": staticmethod(lambda p: ("du", real_du_profile)),
        "find_candidate_locations": staticmethod(_fake_find),
    })
    monkeypatch.setattr(X, "_te", fake_te)
    monkeypatch.setattr(X, "HAVE_RECON", True)
    monkeypatch.setattr(serve, "_path", lambda name: ROOT / "du annual 2020.pdf")
    monkeypatch.setattr(serve, "_state",
                        {"files": [], "pages": None, "scans": {}, "pngs": {}, "pagetext": {}})
    with pytest.MonkeyPatch.context() as mp:
        mp.setattr("pathlib.Path.stat", lambda self: type("S", (), {"st_mtime": 1.0})())
        r1 = serve.telecom_candidates("du annual 2020.pdf")
        r2 = serve.telecom_candidates("du annual 2020.pdf")
    assert r1 == r2 == {"available": True, "company": "du", "candidates": []}
    assert calls == ["Consolidated statement of profit or loss (through 'Profit for the year')",
                      "Operating expenses / General and administrative expenses note"]  # only ONE round


@pytest.mark.skipif(not X.HAVE_RECON, reason="telecom_extract not importable")
@pytest.mark.skipif(_sample_pdf() is None, reason="no sample PDF present")
def test_http_telecom_candidates_round_trip(monkeypatch):
    """Upload a real du report -> /api/telecom_candidates -> /api/page_raw
    with the returned bbox, through the real HTTP handler."""
    pdf = _sample_pdf()
    if "du" not in pdf.name.lower():
        pytest.skip("sample PDF isn't a du report")
    monkeypatch.setattr(serve, "_state",
                        {"files": [pdf], "pages": None, "scans": {}, "pngs": {},
                         "pagetext": {}})
    from http.server import ThreadingHTTPServer
    port = serve._free_port(9300)
    srv = ThreadingHTTPServer(("127.0.0.1", port), serve.Handler)
    th = threading.Thread(target=srv.serve_forever, daemon=True)
    th.start()
    try:
        base = f"http://127.0.0.1:{port}"
        res = json.load(urllib.request.urlopen(
            base + "/api/telecom_candidates?file=" + urllib.request.quote(pdf.name)))
        assert res["available"] is True
        assert res["candidates"], "expected at least one candidate on a real du report"
        c = res["candidates"][0]
        png = urllib.request.urlopen(
            base + f"/api/page_raw?file={urllib.request.quote(pdf.name)}"
                   f"&n={c['page']}&scale=1.5&bbox={','.join(str(v) for v in c['bbox'])}").read()
        assert png[:8] == b"\x89PNG\r\n\x1a\n"
    finally:
        srv.shutdown()


# ------------------------------------------------- remote-access passcode ---
def test_is_authed_is_inert_when_remote_passcode_is_unset(monkeypatch):
    """The whole gate must be a complete no-op for normal `python serve.py`
    usage -- this is the one assertion that most directly protects every
    other test (and every real local user) from ever seeing a login wall
    they didn't ask for."""
    monkeypatch.setattr(serve, "_REMOTE_PASSCODE", None)
    fake_handler = type("H", (), {"headers": {}})()
    assert serve._is_authed(fake_handler) is True


def test_is_authed_requires_a_known_session_when_remote_passcode_is_set(monkeypatch):
    monkeypatch.setattr(serve, "_REMOTE_PASSCODE", "test-passcode-123")
    monkeypatch.setattr(serve, "_AUTH_SESSIONS", {"good-token"})
    no_cookie = type("H", (), {"headers": {}})()
    assert serve._is_authed(no_cookie) is False
    bad_cookie = type("H", (), {"headers": {"Cookie": "tk_session=wrong-token"}})()
    assert serve._is_authed(bad_cookie) is False
    good_cookie = type("H", (), {"headers": {"Cookie": "tk_session=good-token"}})()
    assert serve._is_authed(good_cookie) is True


def test_http_cross_origin_post_refused_without_remote_mode(monkeypatch):
    """The ORIGINAL protection (predates --allow-remote entirely): in normal
    local-only usage, a POST whose Origin isn't 127.0.0.1/localhost is
    refused outright. No dedicated test existed for this before -- closing
    that gap here, since everything else in this file now touches the same
    code path."""
    monkeypatch.setattr(serve, "_state",
                        {"files": [], "pages": None, "scans": {}, "manual": {},
                         "pngs": {}, "pagetext": {}, "telecom": {}})
    monkeypatch.setattr(serve, "_REMOTE_PASSCODE", None)
    from http.server import ThreadingHTTPServer
    port = serve._free_port(9410)
    srv = ThreadingHTTPServer(("127.0.0.1", port), serve.Handler)
    th = threading.Thread(target=srv.serve_forever, daemon=True)
    th.start()
    try:
        req = urllib.request.Request(
            f"http://127.0.0.1:{port}/api/upload", data=b"{}", method="POST",
            headers={"Content-Type": "application/json",
                     "Origin": "https://example-tunnel.trycloudflare.com"})
        try:
            urllib.request.urlopen(req)
            assert False, "expected a 403"
        except urllib.error.HTTPError as e:
            assert e.code == 403
            assert "cross-origin POST refused" in e.read().decode()
    finally:
        srv.shutdown()


def test_http_remote_passcode_gate_full_flow(monkeypatch, tmp_path):
    """Full round trip through the real HTTP handler: an unauthenticated
    request gets bounced to /login, the wrong passcode doesn't grant a
    session, the right one does (and sets a cookie), and -- the actual point
    of this feature -- a cross-origin POST that would normally be refused
    succeeds once authenticated via the passcode, while one WITHOUT the
    cookie still gets refused exactly as before."""
    import http.cookiejar
    monkeypatch.setattr(serve, "_state",
                        {"files": [], "pages": None, "scans": {}, "manual": {},
                         "pngs": {}, "pagetext": {}, "telecom": {}})
    monkeypatch.setattr(serve, "_REMOTE_PASSCODE", "correct-horse-battery")
    monkeypatch.setattr(serve, "_AUTH_SESSIONS", set())
    # the POST used below to prove the passcode-authenticated bypass really
    # hits /api/upload (a real file-writing endpoint, chosen deliberately --
    # it's the simplest _NEEDS_FILE-exempt route) -- isolate it from the
    # real uploads/ dir, same pattern as test_http_manual_extraction_round_trip
    monkeypatch.setattr(serve, "UPLOAD_DIR", tmp_path / "uploads")
    from http.server import ThreadingHTTPServer
    port = serve._free_port(9400)
    srv = ThreadingHTTPServer(("127.0.0.1", port), serve.Handler)
    th = threading.Thread(target=srv.serve_forever, daemon=True)
    th.start()
    try:
        base = f"http://127.0.0.1:{port}"

        # unauthenticated GET bounces to the login page (urllib follows the
        # 302 automatically, so check the CONTENT it lands on)
        body = urllib.request.urlopen(base + "/api/files").read().decode()
        assert "Enter passcode" in body

        # wrong passcode: no cookie granted, still bounced to login
        wrong = urllib.request.urlopen(urllib.request.Request(
            base + "/login", data=b"passcode=nope", method="POST",
            headers={"Content-Type": "application/x-www-form-urlencoded"}))
        assert "Enter passcode" in wrong.read().decode()
        assert not serve._AUTH_SESSIONS

        # without a session, a POST (cross-origin or not) is bounced to
        # /login rather than let through -- it never even reaches the
        # origin check below, which exists for the OTHER case (see
        # test_http_cross_origin_post_refused_without_remote_mode)
        req = urllib.request.Request(
            base + "/api/upload", data=b"{}", method="POST",
            headers={"Content-Type": "application/json",
                     "Origin": "https://example-tunnel.trycloudflare.com"})
        assert "Enter passcode" in urllib.request.urlopen(req).read().decode()

        # correct passcode grants a session cookie
        jar = http.cookiejar.CookieJar()
        opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(jar))
        opener.open(urllib.request.Request(
            base + "/login", data=b"passcode=correct-horse-battery", method="POST",
            headers={"Content-Type": "application/x-www-form-urlencoded"}))
        assert any(c.name == "tk_session" for c in jar)

        # the SAME cross-origin POST now succeeds through the authenticated session
        req2 = urllib.request.Request(
            base + "/api/upload",
            data=json.dumps({"filename": "x.pdf", "data_b64": ""}).encode(),
            method="POST",
            headers={"Content-Type": "application/json",
                     "Origin": "https://example-tunnel.trycloudflare.com"})
        res = opener.open(req2)
        assert res.status == 200
    finally:
        srv.shutdown()
