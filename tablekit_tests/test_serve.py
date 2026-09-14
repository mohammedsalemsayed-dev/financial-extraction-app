"""
Tests for serve.py -- the preview / edit / export backend.

The pure-logic paths (edit application, inventory shape, cross-year wiring,
one-workbook export) run with a monkey-patched scanner and need no PDF.
One end-to-end HTTP test runs only if a sample PDF is present.
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


def test_free_port_returns_an_open_port():
    import socket
    p = serve._free_port(9000)
    with socket.socket() as s:
        s.bind(("127.0.0.1", p))                 # must be bindable == was free


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
