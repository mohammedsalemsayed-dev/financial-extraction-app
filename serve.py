"""
serve.py -- local preview / edit / export UI for extract_all_tables.py

    python serve.py  report.pdf  [more.pdf ...]
    python extract_all_tables.py report.pdf --serve

Opens a page on 127.0.0.1 that lists every detected table, shows each one next
to the original PDF page, lets you FIX a wrong label or value inline, tick
several, and download them as ONE Excel workbook (one sheet per table).

Local only.  No document bytes leave the machine.  Digital-text PDFs only.
"""
from __future__ import annotations

import base64
import io
import json
import logging
import re
import socket
import sys
import threading
import time
import webbrowser
from pathlib import Path
from urllib.parse import urlparse, parse_qs
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler

import pdfplumber

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))
import extract_all_tables as X  # noqa: E402

HTML = ROOT / "webui.html"
LOG = logging.getLogger("tablekit.serve")
UPLOAD_DIR = ROOT / "uploads"

STMT = ("income statement", "statement of financial position",
        "statement of cash flows", "statement of changes in equity")

# manual (user-drawn-box) tables live in a separate per-file list so they
# never collide with the auto-detect scan's own 1-based numbering -- a
# manual table's "n" sent to the browser is offset by this much, and every
# endpoint that resolves "n" to a table checks the offset to know which list
# to look in (see `_resolve`).
MANUAL_OFFSET = 100000

_state = {
    "files": [],                 # list[Path]
    "pages": None,               # optional page range (0-based) applied to every scan
    "scans": {},                 # name -> {"mtime": float, "tables": [...], "warn": [...]}
    "manual": {},                 # name -> [table dict or None (deleted), ...]
    "pngs": {},                  # (name, n, scale) -> bytes
    "pagetext": {},               # name -> [page1 text, page2 text, ...]  for search
}
_lock = threading.RLock()


# --------------------------------------------------------------------- scan ---
def _path(name):
    with _lock:
        return next((p for p in _state["files"] if p.name == name), None)


def scan_file(name: str, force=False):
    """Scan one PDF once; re-scan automatically if the file changed on disk."""
    path = _path(name)
    if path is None:
        raise KeyError(name)
    mtime = path.stat().st_mtime
    with _lock:
        cached = _state["scans"].get(name)
        if cached and not force and cached["mtime"] == mtime:
            return cached["tables"]
    LOG.info("scanning %s ...", name)
    t0 = time.time()
    warns = []
    tables = X.scan([path], _state["pages"], 2, 2,
                    warn=lambda m: warns.append(m.strip()))
    keep = []
    for t in tables:
        if t["kind"] in STMT or t["kind"] == "note":
            keep.append(t)
        elif t.get("shape") != "mostly-text" and len(t["rows"]) >= 3:
            keep.append(t)
    with _lock:
        _state["scans"][name] = {"mtime": mtime, "tables": keep, "warn": warns}
        # invalidate any rendered pages for this file
        _state["pngs"] = {k: v for k, v in _state["pngs"].items() if k[0] != name}
    LOG.info("  %s: %d tables in %.1fs", name, len(keep), time.time() - t0)
    _run_cross_year()
    return keep


def _run_cross_year():
    """Re-run the prior-year consistency check across every file scanned so far."""
    with _lock:
        allt = [t for s in _state["scans"].values() for t in s["tables"]]
    for t in allt:
        t.pop("consistency", None)
    X.cross_year_check(allt)


def _edit_one(t, e):
    """Apply one edit dict to a copy of table `t` and re-run the analysis so
    the verdicts / health reflect the fix.  ALL cell parsing happens here
    (extract_all_tables._cell), never in the browser."""
    t2 = dict(t)
    if e.get("rows"):
        t2["rows"] = [[X._cell(c) if isinstance(c, str) else c for c in r]
                      for r in e["rows"] if any(x not in (None, "") for x in r)]
    if e.get("title"):
        t2["title"] = e["title"]
        t2["_sheet_name"] = e["title"]
    t2.pop("notes", None)
    X.analyze(t2, doc_years=([max(t2["years"]), max(t2["years"]) - 1]
                             if t2.get("years") else None))
    X._attach_health(t2)
    t2["_edited"] = True
    return t2


def _apply_edits(tables, edits):
    """edits: {"<n>": {"rows": [[...]], "title": "..."}}  -> re-analysed copies.
    n here is a 1-based position in `tables` itself (not the app-wide "n"
    that can also address a manual-mode table -- see `_resolve`/`export_xlsx`
    for that)."""
    if not edits:
        return tables
    out = []
    for n, t in enumerate(tables, 1):
        e = edits.get(str(n)) or edits.get(n)
        out.append(_edit_one(t, e) if e else t)
    return out


# --------------------------------------------------------------- manual mode ---
# Tabula-style workflow: upload one file, browse its pages, draw a box around
# the table you want, extract just that region. Runs through the SAME
# classify/analyze/health pipeline as the automatic scan (X.extract_region),
# so the rest of this file's preview/edit/export code needs no branching --
# it just needs to find the right table dict given an "n" (see `_resolve`).
def _manual_list(name):
    # defensive .setdefault on _state itself too: a few tests replace
    # `_state` wholesale with a dict that predates this feature
    return _state.setdefault("manual", {}).setdefault(name, [])


def _resolve(name, n):
    """n > MANUAL_OFFSET addresses a manually-extracted table; otherwise it's
    a 1-based index into `scan_file`'s results. The web UI has no route left
    that triggers a scan (the whole-file automatic detector was pulled --
    unreliable, not worth shipping), so in practice n is always > MANUAL_OFFSET
    from the browser; the auto branch stays only as the seam tests use to
    seed table data (see test_serve.py's `wired` fixture).
    A manual entry can be `None` (tombstoned by delete_manual) -- treated
    the same as "doesn't exist"."""
    if n > MANUAL_OFFSET:
        idx = n - MANUAL_OFFSET - 1
        manual = _manual_list(name)
        if 0 <= idx < len(manual) and manual[idx] is not None:
            return manual[idx]
        raise KeyError(n)
    auto = scan_file(name)
    if 1 <= n <= len(auto):
        return auto[n - 1]
    raise KeyError(n)


def delete_manual(name, n):
    """Remove a manually-extracted table. Tombstones (sets to None) rather
    than removing from the list, so every OTHER manual table's "n" (which
    encodes its list position) stays valid -- removing outright would shift
    every later entry's effective index and silently repoint any selection
    or edit state the browser still has cached under the old number."""
    if n <= MANUAL_OFFSET:
        raise KeyError(n)
    idx = n - MANUAL_OFFSET - 1
    with _lock:
        manual = _manual_list(name)
        if not (0 <= idx < len(manual)) or manual[idx] is None:
            raise KeyError(n)
        manual[idx] = None


def page_count(name):
    path = _path(name)
    if path is None:
        raise KeyError(name)
    with pdfplumber.open(path) as pdf:
        return len(pdf.pages)


def page_raw_png(name, n, scale, highlight=None):
    highlight = (highlight or "").strip()
    key = ("raw", name, n, round(scale, 2), highlight.lower())
    with _lock:
        if key in _state["pngs"]:
            return _state["pngs"][key]
    path = _path(name)
    if path is None:
        raise KeyError(name)
    with pdfplumber.open(path) as pdf:
        if not (1 <= n <= len(pdf.pages)):
            raise KeyError(n)
        page = pdf.pages[n - 1]
        im = page.to_image(resolution=int(72 * scale))
        if highlight:
            try:
                for m in page.search(highlight, case=False):
                    im.draw_rect((m["x0"], m["top"], m["x1"], m["bottom"]),
                                stroke="#f59e0b", stroke_width=1,
                                fill=(253, 224, 71, 90))
            except Exception:
                LOG.debug("highlight search failed", exc_info=True)
        buf = io.BytesIO()
        im.save(buf, format="PNG")
        data = buf.getvalue()
    with _lock:
        if len(_state["pngs"]) > 40:
            _state["pngs"].clear()
        _state["pngs"][key] = data
    return data


def _page_texts(name):
    """Every page's extracted text, cached per file so repeat searches (and
    re-searches after typing more) don't re-open/re-parse the whole PDF."""
    with _lock:
        cached = _state.setdefault("pagetext", {}).get(name)
        if cached is not None:
            return cached
    path = _path(name)
    if path is None:
        raise KeyError(name)
    with pdfplumber.open(path) as pdf:
        texts = [pg.extract_text() or "" for pg in pdf.pages]
    with _lock:
        _state.setdefault("pagetext", {})[name] = texts
    return texts


def search_pdf(name, query, max_hits=300):
    query = (query or "").strip()
    if len(query) < 2:
        return []
    texts = _page_texts(name)
    q = query.lower()
    out = []
    for i, text in enumerate(texts):
        low = text.lower()
        count = low.count(q)
        if not count:
            continue
        pos = low.find(q)
        snip_start = max(0, pos - 40)
        snippet = text[snip_start:pos + len(query) + 40].replace("\n", " ").strip()
        out.append({"page": i + 1, "count": count, "snippet": snippet})
        if len(out) >= max_hits:
            break
    return out


_QUICKFIND_PATTERNS = [
    ("Income statement", re.compile(
        r"(consolidated\s+)?(statement of (profit or loss|comprehensive income)|income statement)", re.I)),
    ("Balance sheet", re.compile(
        r"(consolidated\s+)?statement of financial position|balance sheet", re.I)),
    ("Cash flow statement", re.compile(
        r"(consolidated\s+)?statement of cash flows?", re.I)),
    ("Changes in equity", re.compile(
        r"(consolidated\s+)?statement of changes in equity", re.I)),
]


def quick_find_statements(name):
    """A fast, text-only pass (reuses the same cached page text as search --
    no img2table, no full detection) that finds roughly where the four core
    financial statements sit, so a freshly-uploaded report doesn't just dump
    the user on its cover page with no idea which of 100+ pages to look at.
    Returns the FIRST page each kind is found on, in document order."""
    texts = _page_texts(name)
    found = {}
    for i, text in enumerate(texts):
        for label, pat in _QUICKFIND_PATTERNS:
            if label not in found and pat.search(text):
                found[label] = i + 1
        if len(found) == len(_QUICKFIND_PATTERNS):
            break
    return [{"label": label, "page": found[label]}
            for label, _ in _QUICKFIND_PATTERNS if label in found]


def upload_pdf(filename, data_b64):
    safe = re.sub(r"[^\w .()-]", "_", Path(filename or "upload.pdf").name) or "upload.pdf"
    if not safe.lower().endswith(".pdf"):
        safe += ".pdf"
    UPLOAD_DIR.mkdir(exist_ok=True)
    dest = UPLOAD_DIR / safe
    i = 1
    while dest.exists():
        dest = UPLOAD_DIR / f"{Path(safe).stem}_{i}{Path(safe).suffix}"
        i += 1
    dest.write_bytes(base64.b64decode(data_b64))
    with _lock:
        _state["files"].append(dest)
    return dest.name


def extract_region(name, page1, bbox, title=None):
    path = _path(name)
    if path is None:
        raise KeyError(name)
    t = X.extract_region(path, page1 - 1, tuple(bbox), title or None)
    if t is None:
        return None
    with _lock:
        manual = _manual_list(name)
        manual.append(t)
        idx = len(manual)
    return _detail(t, MANUAL_OFFSET + idx)


def extract_region_ocr(name, page1, bbox, title=None):
    path = _path(name)
    if path is None:
        raise KeyError(name)
    t = X.extract_region_ocr(path, page1 - 1, tuple(bbox), title or None)
    if t is None:
        return None
    with _lock:
        manual = _manual_list(name)
        manual.append(t)
        idx = len(manual)
    return _detail(t, MANUAL_OFFSET + idx)


def region_has_text(name, page1, bbox):
    path = _path(name)
    if path is None:
        raise KeyError(name)
    return X.box_has_text(path, page1 - 1, tuple(bbox))


# ------------------------------------------------------------------- shape ----
def _pill(t):
    h = t.get("health") or {}
    return {
        "n": None,
        "page": t["page_label"],
        "kind": t["kind"],
        "title": t.get("title") or "(untitled)",
        "years": t.get("years") or [],
        "foots": t.get("foots"),
        "foot_detail": t.get("foot_detail", ""),
        "foot_by_col": t.get("foot_by_col") or [],
        "notes": t.get("notes") or [],
        "consistency": t.get("consistency"),
        "health": h.get("score"),
        "health_labels": h.get("labels"),
        "health_figures": h.get("figures"),
        "n_suspect": len(h.get("suspect") or []),
        "nrows": len(t["rows"]),
        "ncols": max((len(r) for r in t["rows"]), default=0),
        "group": ("Manual selections" if t.get("_manual")
                  else ("Financial statements" if t["kind"] in STMT
                        else ("Notes" if t["kind"] == "note" else "Other tables"))),
        "stitched_from": t.get("_stitched_from"),
        "edited": bool(t.get("_edited")),
        "ocr": bool(t.get("_ocr")),
    }


def inventory(name):
    """List every table found for this file -- manual box extractions only,
    from the web UI's point of view (the whole-file automatic detector has
    no route left to trigger it). Never triggers a scan itself."""
    if _path(name) is None:
        raise KeyError(name)
    with _lock:
        auto = (_state["scans"].get(name) or {}).get("tables", [])
        manual = _manual_list(name)
        warn = (_state["scans"].get(name) or {}).get("warn", [])
    out = []
    for idx, t in enumerate(manual):
        if t is None:      # deleted
            continue
        row = _pill(t)
        row["n"] = MANUAL_OFFSET + idx + 1
        out.append(row)
    for n, t in enumerate(auto, 1):
        row = _pill(t)
        row["n"] = n
        out.append(row)
    return {"file": name, "tables": out, "warnings": warn,
            "auto_scanned": name in _state["scans"]}


def _detail(t, n):
    h = t.get("health") or {}
    return {
        "n": n, "page": t["page_label"], "kind": t["kind"],
        "title": t.get("title"), "years": t.get("years") or [],
        "foots": t.get("foots"), "foot_detail": t.get("foot_detail", ""),
        "foot_by_col": t.get("foot_by_col") or [],
        "notes": t.get("notes") or [],
        "consistency": t.get("consistency"),
        "edited": bool(t.get("_edited")),
        "ocr": bool(t.get("_ocr")),
        "health": h.get("score"), "health_labels": h.get("labels"),
        "health_figures": h.get("figures"), "scored": h.get("scored"),
        "suspect": h.get("suspect") or [],
        "header_idx": t.get("header_idx", 0),
        "data_start": t.get("data_start", 0),
        "value_cols": t.get("value_cols") or [],
        "total_rows": t.get("total_rows") or [],
        "rows": t["rows"],
    }


def table_detail(name, n):
    return _detail(_resolve(name, n), n)


def reanalyze(name, n, rows, title=None):
    """Apply an in-progress edit (raw cells) and return a fresh detail dict --
    verdicts / health recomputed by the engine, cells parsed by the engine."""
    base = _resolve(name, n)
    t2 = _edit_one(base, {"rows": rows, "title": title})
    return _detail(t2, n)


def compare(name_a, n_a, name_b, n_b):
    ta = _resolve(name_a, n_a)
    tb = _resolve(name_b, n_b)
    drows, verdict = X.diff_tables(ta, tb)
    return {"verdict": verdict, "rows": drows,
            "a": {"file": name_a, "title": ta.get("title"), "years": ta.get("years")},
            "b": {"file": name_b, "title": tb.get("title"), "years": tb.get("years")}}


def page_png(name, n, scale):
    key = (name, n, round(scale, 2))
    with _lock:
        if key in _state["pngs"]:
            return _state["pngs"][key]
    t = _resolve(name, n)
    path = _path(name)
    with pdfplumber.open(path) as pdf:
        page = pdf.pages[t["page_label"] - 1]
        im = page.to_image(resolution=int(72 * scale))
        bb = t.get("bbox")
        if bb and len(bb) == 4 and bb[2] > bb[0] and bb[3] > bb[1]:
            x0, y0, x1, y1 = bb
            try:
                im.draw_rect((x0, y0, min(x1, page.width - .5),
                              min(y1, page.height - .5)),
                             stroke="#4f46e5", stroke_width=3, fill=None)
            except Exception:
                pass
        buf = io.BytesIO()
        im.save(buf, format="PNG")
        data = buf.getvalue()
    with _lock:
        if len(_state["pngs"]) > 40:
            _state["pngs"].clear()
        _state["pngs"][key] = data
    return data


def export_xlsx(name, ns, edits=None):
    edits = edits or {}
    picked = []
    for n in ns:
        try:
            base = _resolve(name, n)
        except KeyError:
            continue
        e = edits.get(str(n)) or edits.get(n)
        picked.append(_edit_one(base, e) if e else base)
    wb = X.build_workbook(picked)
    buf = io.BytesIO()
    wb.save(buf)
    stem = Path(name).stem
    fn = (f"{stem}__{len(picked)}_tables.xlsx" if len(picked) != 1
          else f"{stem}__{(picked[0].get('title') or 'table')[:30].strip()}.xlsx")
    fn = "".join(c if (c.isalnum() or c in " _-.") else "_" for c in fn)
    return buf.getvalue(), fn


# --------------------------------------------------------------------- http ---
class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        LOG.debug("%s - %s", self.address_string(), fmt % args)

    def _send(self, code, body, ctype="application/json", extra=None):
        if isinstance(body, (dict, list)):
            body = json.dumps(body, default=str).encode("utf-8")
        elif isinstance(body, str):
            body = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        u = urlparse(self.path)
        q = parse_qs(u.query)
        try:
            if u.path in ("/", "/index.html"):
                return self._send(200, HTML.read_text(encoding="utf-8"),
                                  "text/html; charset=utf-8")
            if u.path == "/api/files":
                with _lock:
                    return self._send(200, {"files": [p.name for p in _state["files"]]})
            if u.path == "/api/status":
                return self._send(200, {"img2table": X.HAVE_IMG2TABLE, "ocr": X.HAVE_OCR})
            if u.path == "/api/scan":
                return self._send(200, inventory(q["file"][0]))
            if u.path == "/api/table":
                return self._send(200, table_detail(q["file"][0], int(q["n"][0])))
            if u.path == "/api/page":
                # cap raised from 3.0 -> 4.5 (72*4.5 = 324 DPI) so the preview
                # can actually render sharp on a high-DPI/retina display when
                # the client asks for scale*devicePixelRatio -- 3.0 (216 DPI)
                # was the ceiling even on a 1x display asking for "sharp"
                scale = max(1.0, min(float(q.get("scale", ["2"])[0]), 4.5))
                return self._send(200, page_png(q["file"][0], int(q["n"][0]), scale),
                                  "image/png")
            if u.path == "/api/pagecount":
                return self._send(200, {"pages": page_count(q["file"][0])})
            if u.path == "/api/page_raw":
                scale = max(1.0, min(float(q.get("scale", ["2"])[0]), 4.5))
                return self._send(200, page_raw_png(q["file"][0], int(q["n"][0]), scale,
                                                    q.get("hl", [""])[0]),
                                  "image/png")
            if u.path == "/api/search":
                hits = search_pdf(q["file"][0], q.get("q", [""])[0])
                return self._send(200, {"hits": hits})
            if u.path == "/api/quickfind":
                return self._send(200, {"hits": quick_find_statements(q["file"][0])})
            return self._send(404, {"error": "not found"})
        except KeyError:
            return self._send(404, {"error": "unknown file"})
        except Exception as e:
            LOG.exception("GET %s failed", self.path)
            return self._send(500, {"error": f"{type(e).__name__}: {e}"})

    def do_POST(self):
        u = urlparse(self.path)
        # reject cross-origin POSTs -- this is an open localhost endpoint that
        # reads/writes files; only our own page (or a no-Origin client) may post
        origin = self.headers.get("Origin")
        if origin and urlparse(origin).hostname not in ("127.0.0.1", "localhost"):
            return self._send(403, {"error": "cross-origin POST refused"})
        try:
            length = int(self.headers.get("Content-Length", "0"))
            payload = json.loads(self.rfile.read(length) or b"{}")
            # every mutating endpoint below needs a real, already-uploaded
            # file -- catch "no file selected yet" with one clear message
            # instead of a bare KeyError leaking to the browser (this is
            # reachable from the UI: every sidebar control used to be
            # clickable before a file was ever chosen)
            if u.path in ("/api/reanalyze", "/api/extract_region", "/api/extract_region_ocr",
                         "/api/export", "/api/delete_manual"):
                name = payload.get("file")
                if not name or _path(name) is None:
                    return self._send(400, {"error": "No file selected -- upload a PDF first."})
            if u.path == "/api/compare":
                if not payload.get("a") or not payload.get("b") \
                        or _path(payload["a"]) is None or _path(payload["b"]) is None:
                    return self._send(400, {"error": "No file selected -- upload a PDF first."})
            if u.path == "/api/export":
                data, fn = export_xlsx(payload["file"],
                                       [int(x) for x in payload["ns"]],
                                       payload.get("edits"))
                return self._send(
                    200, data,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    {"Content-Disposition": f'attachment; filename="{fn}"'})
            if u.path == "/api/reanalyze":
                return self._send(200, reanalyze(
                    payload["file"], int(payload["n"]),
                    payload["rows"], payload.get("title")))
            if u.path == "/api/compare":
                return self._send(200, compare(
                    payload["a"], int(payload["na"]),
                    payload["b"], int(payload["nb"])))
            if u.path == "/api/upload":
                name = upload_pdf(payload.get("filename"), payload["data_b64"])
                return self._send(200, {"file": name})
            if u.path == "/api/extract_region":
                bbox = [float(v) for v in payload["bbox"]]
                d = extract_region(payload["file"], int(payload["page"]), bbox,
                                   payload.get("title"))
                if d is None:
                    no_text = not region_has_text(payload["file"], int(payload["page"]), bbox)
                    return self._send(422, {
                        "error": ("No text found in that box." if no_text else
                                  "No table found in that box -- try drawing it tighter "
                                  "around just the table's rows and columns."),
                        "no_text": no_text, "ocr_available": X.HAVE_OCR})
                return self._send(200, d)
            if u.path == "/api/extract_region_ocr":
                d = extract_region_ocr(payload["file"], int(payload["page"]),
                                       [float(v) for v in payload["bbox"]],
                                       payload.get("title"))
                if d is None:
                    return self._send(422, {"error":
                        "OCR found nothing table-like in that box."})
                return self._send(200, d)
            if u.path == "/api/delete_manual":
                try:
                    delete_manual(payload["file"], int(payload["n"]))
                except KeyError:
                    return self._send(404, {"error": "that table is already gone"})
                return self._send(200, inventory(payload["file"]))
            return self._send(404, {"error": "not found"})
        except Exception as e:
            LOG.exception("POST %s failed", self.path)
            return self._send(500, {"error": f"{type(e).__name__}: {e}"})


def _free_port(start=8765, tries=20):
    for p in range(start, start + tries):
        with socket.socket() as s:
            try:
                s.bind(("127.0.0.1", p))
                return p
            except OSError:
                continue
    raise RuntimeError("no free port in range")


def run(pdf_args, host="127.0.0.1", port=None, open_browser=True):
    logging.basicConfig(level=logging.INFO, format="  %(message)s")
    files = []
    for a in pdf_args:
        p = Path(a)
        if p.is_dir():
            files += sorted(p.glob("*.pdf"))
        elif p.exists() and p.suffix.lower() == ".pdf":
            files.append(p)
        else:
            LOG.warning("skipping %s (not a PDF)", a)
    _state["files"] = files
    port = port or _free_port()
    srv = ThreadingHTTPServer((host, port), Handler)
    url = f"http://{host}:{port}/"
    print(f"\n  Table preview UI  ->  {url}")
    if X.HAVE_IMG2TABLE:
        print("  img2table engine: ON  (borderless/2-up statements + manual box-select use it)")
    else:
        print("  img2table engine: OFF -- borderless-table detection and manual box-select")
        print("    both fall back to pdfplumber's ruled-table-only extractor. Run:")
        print("    pip install img2table opencv-python-headless pandas")
    if X.HAVE_OCR:
        print("  OCR: ON  (\"OCR this region\" appears for a box with no text at all)")
    else:
        print("  OCR: OFF -- scanned/image-only pages can't be extracted. Run:")
        print("    pip install pytesseract")
        print("    then install the Tesseract binary itself (not pip-installable) --")
        print("    https://github.com/UB-Mannheim/tesseract/wiki (Windows) and make")
        print("    sure tesseract.exe is on PATH")
    if files:
        print(f"  {len(files)} PDF(s):  " + ",  ".join(f.name for f in files))
    else:
        print("  no PDFs given on the command line -- upload one from the page")
    print("  (Ctrl-C to stop)\n")
    if open_browser:
        threading.Timer(0.6, lambda: webbrowser.open(url)).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\n  stopped.")


def _cli(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)
    files = [a for a in argv if not a.startswith("--")]
    run(files)


if __name__ == "__main__":
    _cli()
