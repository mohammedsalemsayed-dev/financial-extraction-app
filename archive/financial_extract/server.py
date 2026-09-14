"""
Localhost-only HTTP server for the extraction app.

Pure Python standard library (http.server) — no web framework, no external
network. Binds to 127.0.0.1 so nothing is reachable from outside the machine.
"""
from __future__ import annotations

import json
import re
import threading
import time
import traceback
import uuid
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse, parse_qs

from . import core

WEB_DIR = Path(__file__).resolve().parent / "web"
HOST = "127.0.0.1"

# ---- in-memory session state -------------------------------------------------
class _Session:
    def __init__(self):
        self.workdir = Path(core.te.tempfile.mkdtemp(prefix="fx_app_"))
        self.docs: dict[str, dict] = {}      # doc_id -> {path, detect, extraction?}
        self.lock = threading.Lock()

    def add_pdf(self, filename: str, data: bytes) -> dict:
        safe = re.sub(r"[^A-Za-z0-9._ -]", "_", filename) or "document.pdf"
        if not safe.lower().endswith(".pdf"):
            safe += ".pdf"
        doc_id = uuid.uuid4().hex[:12]
        path = self.workdir / f"{doc_id}__{safe}"
        path.write_bytes(data)
        info = core.detect(path)
        rec = {"doc_id": doc_id, "path": path, "file": safe, "detect": info,
               "extraction": None}
        with self.lock:
            self.docs[doc_id] = rec
        return rec

    def list(self) -> list:
        with self.lock:
            return [{
                "doc_id": d["doc_id"], "file": d["file"],
                **d["detect"],
                "extracted": d["extraction"] is not None,
            } for d in self.docs.values()]


SESSION = _Session()

CONTENT_TYPES = {
    ".html": "text/html; charset=utf-8",
    ".js": "text/javascript; charset=utf-8",
    ".css": "text/css; charset=utf-8",
    ".svg": "image/svg+xml",
}


class Handler(BaseHTTPRequestHandler):
    server_version = "FinancialExtract/0.1"

    # -- helpers -----------------------------------------------------------
    def _send_json(self, obj, status=200):
        payload = json.dumps(obj).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(payload)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(payload)

    def _send_bytes(self, data, ctype, status=200, extra=None):
        self.send_response(status)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(data)

    def _body(self) -> bytes:
        n = int(self.headers.get("Content-Length", 0))
        return self.rfile.read(n) if n else b""

    def _json_body(self) -> dict:
        raw = self._body()
        return json.loads(raw.decode("utf-8")) if raw else {}

    def log_message(self, fmt, *args):
        pass  # quiet

    # -- routing ---------------------------------------------------------------
    def do_GET(self):
        u = urlparse(self.path)
        path = u.path
        try:
            if path == "/" or path == "/index.html":
                return self._static("index.html")
            if path.startswith("/web/"):
                return self._static(path[len("/web/"):])
            if path == "/api/docs":
                return self._send_json({"docs": SESSION.list()})
            if path == "/api/page":
                q = parse_qs(u.query)
                doc_id = q.get("doc_id", [""])[0]
                page = int(q.get("page", ["1"])[0])
                d = SESSION.docs.get(doc_id)
                if not d:
                    return self._send_json({"error": "unknown doc"}, 404)
                png = core.page_png(d["path"], page)
                return self._send_bytes(png, "image/png",
                                        extra={"Cache-Control": "max-age=120"})
            return self._send_json({"error": "not found"}, 404)
        except Exception as e:
            traceback.print_exc()
            return self._send_json({"error": str(e)}, 500)

    def do_POST(self):
        u = urlparse(self.path)
        path = u.path
        try:
            if path == "/api/upload":
                fname = self.headers.get("X-Filename", "document.pdf")
                data = self._body()
                if not data:
                    return self._send_json({"error": "empty upload"}, 400)
                rec = SESSION.add_pdf(fname, data)
                return self._send_json({
                    "doc_id": rec["doc_id"], "file": rec["file"], **rec["detect"],
                    "extracted": False,
                })
            if path == "/api/extract":
                body = self._json_body()
                d = SESSION.docs.get(body.get("doc_id"))
                if not d:
                    return self._send_json({"error": "unknown doc"}, 404)
                res = core.extract_document(d["path"])
                d["extraction"] = res
                return self._send_json(res)
            if path == "/api/recheck":
                body = self._json_body()
                res = core.recheck(body["reconcile_kind"], body["body"])
                return self._send_json(res)
            if path == "/api/export":
                body = self._json_body()
                docs_payload = body.get("documents", [])
                if not docs_payload:
                    return self._send_json({"error": "nothing to export"}, 400)
                xlsx = core.build_workbook(docs_payload)
                name = body.get("filename") or "financial_tables.xlsx"
                return self._send_bytes(
                    xlsx,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    extra={"Content-Disposition": f'attachment; filename="{name}"'})
            return self._send_json({"error": "not found"}, 404)
        except Exception as e:
            traceback.print_exc()
            return self._send_json({"error": str(e)}, 500)

    # -- static ----------------------------------------------------------------
    def _static(self, rel):
        rel = rel.lstrip("/")
        target = (WEB_DIR / rel).resolve()
        if not str(target).startswith(str(WEB_DIR.resolve())) or not target.is_file():
            return self._send_json({"error": "not found"}, 404)
        ctype = CONTENT_TYPES.get(target.suffix, "application/octet-stream")
        self._send_bytes(target.read_bytes(), ctype,
                         extra={"Cache-Control": "no-cache"})


def serve(port: int = 8765, open_browser: bool = True):
    httpd = ThreadingHTTPServer((HOST, port), Handler)
    url = f"http://{HOST}:{port}/"
    print(f"\n  Financial-table extraction app")
    print(f"  running at  {url}")
    print(f"  workdir     {SESSION.workdir}")
    print(f"  (localhost only — no network access for document content)")
    print(f"\n  Ctrl-C to stop.\n")
    if open_browser:
        def _open():
            time.sleep(0.6)
            try:
                import webbrowser
                webbrowser.open(url)
            except Exception:
                pass
        threading.Thread(target=_open, daemon=True).start()
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n  stopped.")
        httpd.shutdown()
