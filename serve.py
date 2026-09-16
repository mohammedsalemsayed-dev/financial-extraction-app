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
import functools
import io
import json
import logging
import pickle
import re
import secrets
import socket
import sys
import threading
import time
import webbrowser
from pathlib import Path
from typing import Any
from urllib.parse import urlparse, parse_qs
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler

import pdfplumber

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))
import extract_all_tables as X  # noqa: E402
from tablekit.parse import FormattedNumber  # noqa: E402

HTML = ROOT / "webui.html"
WEBUI_CSS = ROOT / "webui.css"
WEBUI_JS = ROOT / "webui.js"
LOG = logging.getLogger("tablekit.serve")
UPLOAD_DIR = ROOT / "uploads"
SESSION_DIR = UPLOAD_DIR / ".sessions"

STMT = ("income statement", "statement of financial position",
        "statement of cash flows", "statement of changes in equity")


def _debug_state_line(path):
    """One-line snapshot of server-side state size (DEBUG level only -- a
    no-op cost when logging isn't enabled), logged at the top of every
    request. A user-reported "breaks after 2-3 runs" issue (see CHANGELOG
    0.7.1) was never actually diagnosed -- it stopped happening, but no one
    knows which fix did it or whether something like it could recur. This
    doesn't retroactively explain that one, but it means a NEXT occurrence
    has a timeline of state growth to look at (a cache that isn't supposed
    to grow unbounded, a manual-table count that doesn't match what the UI
    shows) instead of nothing. Enable with `python serve.py --debug`."""
    with _lock:
        # .get(), not [...]: same reason as _manual_list/_deleted_list below
        # -- a few tests replace `_state` wholesale with a partial dict.
        n_scans = len(_state.get("scans", {}))
        n_pngs = len(_state.get("pngs", {}))
        n_manual = sum(len(v) for v in _state.get("manual", {}).values())
        n_files = len(_state.get("files", []))
    LOG.debug("%s  [state: files=%d scans=%d pngs=%d manual_tables=%d]",
              path, n_files, n_scans, n_pngs, n_manual)

# manual (user-drawn-box) tables live in a separate per-file list so they
# never collide with the auto-detect scan's own 1-based numbering -- a
# manual table's "n" sent to the browser is offset by this much, and every
# endpoint that resolves "n" to a table checks the offset to know which list
# to look in (see `_resolve`).
MANUAL_OFFSET = 100000

_state: dict[str, Any] = {
    "files": [],                 # list[Path]
    "pages": None,               # optional page range (0-based) applied to every scan
    "scans": {},                 # name -> {"mtime": float, "tables": [...], "warn": [...]}
    "manual": {},                 # name -> [table dict or None (deleted), ...]
    "deleted": {},                # name -> [idx, idx, ...] most-recent-last, for undelete
    "pngs": {},                  # (name, n, scale) -> bytes
    "pagetext": {},               # name -> [page1 text, page2 text, ...]  for search
    "telecom": {},                # name -> {"mtime": float, "result": {...}}
}

# Opt-in only (--allow-remote): None by default, meaning every check below
# that starts "if _REMOTE_PASSCODE" short-circuits to False and this whole
# gate is completely inert -- normal `python serve.py` usage is provably
# unaffected by any of this. When set, a random passcode (printed once at
# startup, never written to disk or logged) is required before ANY route
# responds, GET or POST, from a session that hasn't already proven it via
# /login. This exists for one specific case: deliberately exposing this
# local server through a tunnel for a short remote demo, where the origin
# check below (do_POST) can't tell "the person who owns this machine,
# reaching it through their own tunnel" apart from an actual cross-origin
# attacker -- a real passcode can. Sessions are in-memory only (a plain
# set of random tokens), so they don't survive a restart, which is fine for
# what this is for.
_REMOTE_PASSCODE: str | None = None
_AUTH_SESSIONS: set[str] = set()
_SESSION_COOKIE = "tk_session"


def _session_token(handler) -> str | None:
    raw = handler.headers.get("Cookie", "")
    for part in raw.split(";"):
        k, _, v = part.strip().partition("=")
        if k == _SESSION_COOKIE and v:
            return v
    return None


def _is_authed(handler) -> bool:
    if not _REMOTE_PASSCODE:
        return True   # gate is off entirely
    tok = _session_token(handler)
    return bool(tok and tok in _AUTH_SESSIONS)
_lock = threading.RLock()


# ------------------------------------------------------------- session save ---
# Manual extractions + hand-corrected edits are the expensive-to-reproduce
# part of a session (drawing boxes, fixing OCR misreads cell by cell) -- the
# auto-scan cache is cheap to regenerate and deliberately NOT persisted here.
# Pickle, not JSON: table dicts carry FormattedNumber cells and internal-only
# "_"-prefixed keys, and this file is only ever written and read by this same
# process in a directory it controls -- the same trust boundary as the PDFs
# themselves, so pickle's arbitrary-code-on-load risk doesn't add anything new.
def _session_path(name):
    return SESSION_DIR / (name + ".pkl")


def _save_session(name):
    with _lock:
        snapshot = {"manual": list(_manual_list(name)),
                    "deleted": list(_deleted_list(name))}
    try:
        SESSION_DIR.mkdir(parents=True, exist_ok=True)
        tmp = _session_path(name).with_suffix(".pkl.tmp")
        with open(tmp, "wb") as f:
            pickle.dump(snapshot, f)
        tmp.replace(_session_path(name))          # atomic on the same filesystem
    except Exception:
        LOG.exception("could not save session for %s", name)


def _autosaves(fn):
    """Decorates every function that mutates a file's manual-table list
    (extract_region, extract_region_ocr, delete_manual, undelete_manual,
    reanalyze) so it autosaves on the way out -- centralized here instead of
    a bare `_save_session(name)` call at the end of each function body, so a
    NEW mutating endpoint added later can't silently forget it (that used to
    be a real, easy-to-miss risk: five separate call sites, no enforcement).
    `name` is always the wrapped function's first positional argument.
    Skipped entirely if the function raises -- a failed/no-op mutation
    (e.g. delete_manual's KeyError on an already-gone table) has nothing new
    to save."""
    @functools.wraps(fn)
    def wrapper(name, *args, **kwargs):
        result = fn(name, *args, **kwargs)
        _save_session(name)
        return result
    return wrapper


def _load_session(name):
    p = _session_path(name)
    if not p.exists():
        return
    try:
        with open(p, "rb") as f:
            snapshot = pickle.load(f)
        with _lock:
            _state.setdefault("manual", {})[name] = snapshot.get("manual", [])
            _state.setdefault("deleted", {})[name] = snapshot.get("deleted", [])
        n_live = sum(1 for t in _state["manual"][name] if t is not None)
        if n_live:
            LOG.info("  restored %d saved table(s) for %s", n_live, name)
    except Exception:
        LOG.exception("could not load session for %s -- starting fresh", name)


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
    t2.pop("notes_i18n", None)   # kept in lockstep with "notes" -- see _add_note
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


def _deleted_list(name):
    # same defensive pattern as _manual_list, for the undo-delete stack
    return _state.setdefault("deleted", {}).setdefault(name, [])


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


@_autosaves
def delete_manual(name, n):
    """Remove a manually-extracted table. Tombstones (sets to None) rather
    than removing from the list, so every OTHER manual table's "n" (which
    encodes its list position) stays valid -- removing outright would shift
    every later entry's effective index and silently repoint any selection
    or edit state the browser still has cached under the old number.
    The (index, table) pair is pushed onto a small per-file undo stack
    BEFORE the slot is cleared, so `undelete_manual` can put it back."""
    if n <= MANUAL_OFFSET:
        raise KeyError(n)
    idx = n - MANUAL_OFFSET - 1
    with _lock:
        manual = _manual_list(name)
        if not (0 <= idx < len(manual)) or manual[idx] is None:
            raise KeyError(n)
        _deleted_list(name).append((idx, manual[idx]))
        manual[idx] = None


@_autosaves
def undelete_manual(name):
    """Restore the most recently deleted table for this file, if any.
    Returns the restored table's app-wide "n", or None if there was nothing
    left on the undo stack to restore."""
    with _lock:
        stack = _deleted_list(name)
        manual = _manual_list(name)
        while stack:
            idx, table = stack.pop()
            if 0 <= idx < len(manual) and manual[idx] is None:
                manual[idx] = table
                n = MANUAL_OFFSET + idx + 1
                return n
            # slot no longer empty (shouldn't normally happen -- new
            # extractions always append, never reuse a cleared index) --
            # fall through and try the next entry on the stack instead
    return None


def page_count(name):
    path = _path(name)
    if path is None:
        raise KeyError(name)
    with pdfplumber.open(path) as pdf:
        return len(pdf.pages)


def page_raw_png(name, n, scale, highlight=None, bbox=None):
    highlight = (highlight or "").strip()
    key = ("raw", name, n, round(scale, 2), highlight.lower(),
           tuple(round(v, 1) for v in bbox) if bbox else None)
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
        if bbox:
            _draw_candidate_box(im, page, bbox)
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


# du/Etisalat's telecom_extract.py targets identify themselves by hint_key;
# this just gives the browser a short, stable string instead of that key.
_TELECOM_TARGET_SHORT = {"pl_hint": "pl", "admin_hint": "note"}


def telecom_candidates(name):
    """Locate-only fuzzy-match table finder for du/Etisalat/e& files: for
    each of telecom_extract.py's 2 known targets per company (the P&L, the
    operating-expenses note), every candidate location that cleared its
    reference-vocabulary scoring gate -- NOT the full extraction. Deliberately
    gated by strict_profile_for_file (no fallback): this must never run on a
    file that isn't actually shaped like a du/Etisalat report, or it would
    read as reviving the whole-document auto-detector this project already
    removed once for being unreliable. Returns {"available": False,
    "candidates": []} immediately, without ever opening the PDF, for any
    file that doesn't strictly match."""
    path = _path(name)
    if path is None:
        raise KeyError(name)
    _te = getattr(X, "_te", None)
    if not (getattr(X, "HAVE_RECON", False) and _te):
        return {"available": False, "candidates": []}
    key, profile = _te.strict_profile_for_file(path)
    if key is None:
        return {"available": False, "candidates": []}
    mtime = path.stat().st_mtime
    with _lock:
        # .setdefault(), not _state["telecom"] -- a few tests replace _state
        # wholesale with a dict that predates this key (see _manual_list's
        # own comment for the same pattern).
        cached = _state.setdefault("telecom", {}).get(name)
        if cached and cached["mtime"] == mtime:
            return cached["result"]
    out = []
    with pdfplumber.open(path) as pdf:
        page_indices = range(len(pdf.pages))
        for target in profile["targets"]:
            for cand in _te.find_candidate_locations(pdf, page_indices, target):
                out.append({**cand,
                           "target": _TELECOM_TARGET_SHORT.get(target["hint_key"], "?"),
                           "label": target["name"]})
    result = {"available": True, "company": key, "candidates": out}
    with _lock:
        _state.setdefault("telecom", {})[name] = {"mtime": mtime, "result": result}
    return result


def _sanitize_filename(name):
    """Keep Unicode letters/digits plus a small set of safe punctuation;
    everything else becomes "_". str.isalnum() is Unicode-aware (unlike re's
    \\w, which Python's stdlib `re` -- no \\p{L} support here -- would need
    a regex to fake), so an Arabic filename survives this unchanged instead
    of collapsing into underscores. Mirrors the export-filename sanitiser in
    webui.html so the same file round-trips the same way on both ends."""
    return "".join(c if (c.isalnum() or c in " .()-") else "_" for c in name)


def upload_pdf(filename, data_b64):
    safe = _sanitize_filename(Path(filename or "upload.pdf").name) or "upload.pdf"
    if not safe.lower().endswith(".pdf"):
        safe += ".pdf"
    UPLOAD_DIR.mkdir(exist_ok=True)
    # dest.exists() alone only catches a collision with an earlier upload
    # already sitting in UPLOAD_DIR -- a same-named CLI-arg file loaded from
    # elsewhere (found live, running with several real files loaded at
    # once: see _discover_files) wouldn't physically exist AT this path, so
    # it slipped through and produced two dropdown entries under one name,
    # the second permanently unreachable (_path() resolves by name, first
    # match wins). Checking the loaded file list too closes that.
    with _lock:
        existing_names = {f.name for f in _state["files"]}
    dest = UPLOAD_DIR / safe
    i = 1
    while dest.exists() or dest.name in existing_names:
        dest = UPLOAD_DIR / f"{Path(safe).stem}_{i}{Path(safe).suffix}"
        i += 1
    dest.write_bytes(base64.b64decode(data_b64))
    with _lock:
        _state["files"].append(dest)
    _load_session(dest.name)      # picks up a prior run's saved work, if any
    return dest.name


@_autosaves
def extract_region(name, page1, bbox, title=None, grid=None):
    path = _path(name)
    if path is None:
        raise KeyError(name)
    t = X.extract_region(path, page1 - 1, tuple(bbox), title or None, grid=grid)
    if t is None:
        return None
    with _lock:
        manual = _manual_list(name)
        manual.append(t)
        idx = len(manual)
    return _detail(t, MANUAL_OFFSET + idx)


# Deliberately NOT @_autosaves -- that decorator pickles the whole session
# to disk after every successful call (see its own docstring), which is
# right for something that mutates saved state, but this fires on every
# mouse-release during exploratory dragging and mutates nothing at all; a
# disk write per drag would be pure waste.
def detect_grid(name, page1, bbox):
    path = _path(name)
    if path is None:
        raise KeyError(name)
    return X.detect_grid(path, page1 - 1, tuple(bbox))


@_autosaves
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
        "notes_i18n": t.get("notes_i18n") or [],
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


def _fmt_row(row):
    """A cell may carry a currency symbol, a '%', and/or parentheses-as-
    negative the user typed or the extractor read (see
    tablekit.parse.FormattedNumber) that the numeric VALUE itself never
    keeps -- it has to stay a plain signed number for footing, health
    scoring and every arithmetic comparison to keep working. Rather than
    change what `rows` sends (every existing consumer of a numeric cell,
    server and browser alike, expects a plain number there), send the
    original formatting as a same-shaped side channel the UI can use to
    re-append '$'/'%'/parens onto the DISPLAYED text without touching the
    value driving Δ / Δ% or the "num" right-align styling."""
    return [{"p": v.prefix, "s": v.suffix, "n": v.paren_negative}
            if isinstance(v, FormattedNumber) and (v.prefix or v.suffix or v.paren_negative)
            else None for v in row]


def _detail(t, n):
    h = t.get("health") or {}
    return {
        "n": n, "page": t["page_label"], "kind": t["kind"],
        "title": t.get("title"), "years": t.get("years") or [],
        "foots": t.get("foots"), "foot_detail": t.get("foot_detail", ""),
        "foot_by_col": t.get("foot_by_col") or [],
        "notes": t.get("notes") or [],
        "notes_i18n": t.get("notes_i18n") or [],
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
        "fmt": [_fmt_row(r) for r in t["rows"]],
        # the PDF's own "Note" reference column (e.g. "19", "21") --
        # display-only, one entry per row, never merged into "rows" itself
        # so it can't skew value-column detection or footing (see
        # extract_all_tables.row_note_ref / find_all_tables' note_ref_map).
        # Unrelated to "notes"/"notes_i18n" above (the engine's own English
        # warning text about the table) AND to t["note_col"] (an unrelated,
        # pre-existing column-INDEX hint analyze() sets for its own use).
        "note_refs": ([X.row_note_ref(t, r) for r in t["rows"]]
                      if t.get("note_ref_map") else None),
    }


def table_detail(name, n):
    return _detail(_resolve(name, n), n)


@_autosaves
def reanalyze(name, n, rows, title=None):
    """Apply an in-progress edit (raw cells) and return a fresh detail dict --
    verdicts / health recomputed by the engine, cells parsed by the engine.
    Also COMMITS the edit back into the manual table list (when `n`
    addresses one -- the normal case from the browser; the auto-scan branch
    stays transient, same as before) and flushes it to disk, so the edit
    survives a page refresh or a server restart, not just the final export.
    The browser already debounces calls here (450ms after the user stops
    typing), so each commit represents a real pause, not a keystroke.
    @_autosaves fires regardless of whether `n` actually addressed a manual
    table -- harmless: _save_session just re-snapshots current (unchanged)
    state on the test-only auto-scan branch."""
    base = _resolve(name, n)
    t2 = _edit_one(base, {"rows": rows, "title": title})
    if n > MANUAL_OFFSET:
        idx = n - MANUAL_OFFSET - 1
        with _lock:
            manual = _manual_list(name)
            if 0 <= idx < len(manual) and manual[idx] is not None:
                manual[idx] = t2
    return _detail(t2, n)


def compare(name_a, n_a, name_b, n_b):
    ta = _resolve(name_a, n_a)
    tb = _resolve(name_b, n_b)
    drows, verdict, counts = X.diff_tables(ta, tb)
    return {"verdict": verdict, "counts": counts, "rows": drows,
            "a": {"file": name_a, "title": ta.get("title"), "years": ta.get("years")},
            "b": {"file": name_b, "title": tb.get("title"), "years": tb.get("years")}}


def _draw_candidate_box(im, page, bbox):
    """Bake an indigo highlight rectangle for `bbox` ([x0,y0,x1,y1], PDF
    points) into a rendered page image -- shared by page_png (a saved
    table's own bbox) and page_raw_png (an ephemeral candidate's bbox, e.g.
    from telecom_candidates) so there's exactly one drawing convention."""
    if not (bbox and len(bbox) == 4 and bbox[2] > bbox[0] and bbox[3] > bbox[1]):
        return
    x0, y0, x1, y1 = bbox
    try:
        im.draw_rect((x0, y0, min(x1, page.width - .5),
                      min(y1, page.height - .5)),
                     stroke="#4f46e5", stroke_width=3, fill=None)
    except Exception:
        pass


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
        _draw_candidate_box(im, page, t.get("bbox"))
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

    # ---- remote-access passcode gate (--allow-remote only; see _REMOTE_PASSCODE) ----
    def _g_login(self, q):
        err = "Wrong passcode." if q.get("err") else ""
        page = (
            "<!doctype html><html><head><meta charset='utf-8'>"
            "<meta name='viewport' content='width=device-width, initial-scale=1'>"
            "<title>Tables -- sign in</title><style>"
            "body{font:16px system-ui;background:#0b0f19;color:#e6e8ef;"
            "display:flex;align-items:center;justify-content:center;height:100vh;margin:0}"
            "form{background:#141a29;padding:28px;border-radius:12px;width:260px}"
            "input{width:100%;padding:10px;margin:10px 0;border-radius:8px;border:1px solid #333;"
            "background:#0b0f19;color:#e6e8ef;font-size:16px;box-sizing:border-box}"
            "button{width:100%;padding:10px;border-radius:8px;border:0;background:#4f46e5;"
            "color:#fff;font-size:16px;cursor:pointer}"
            ".err{color:#f87171;margin:0 0 8px}"
            "</style></head><body><form method='post' action='/login'>"
            "<h2 style='margin-top:0'>Enter passcode</h2>"
            + (f"<p class='err'>{err}</p>" if err else "")
            + "<input name='passcode' type='password' autofocus autocomplete='off'>"
            "<button type='submit'>Continue</button></form></body></html>")
        return self._send(200, page, "text/html; charset=utf-8")

    def _p_login(self, raw_body):
        # form-encoded, not JSON -- the one POST route that must work BEFORE
        # any session exists, so it also has to skip do_POST's usual
        # json.loads(body) (see do_POST's early dispatch to this method)
        from urllib.parse import parse_qs as _pqs
        fields = _pqs(raw_body.decode("utf-8", "replace"))
        given = (fields.get("passcode") or [""])[0]
        if not (_REMOTE_PASSCODE and secrets.compare_digest(given, _REMOTE_PASSCODE)):
            return self._send(302, "", extra={"Location": "/login?err=1"})
        tok = secrets.token_urlsafe(32)
        _AUTH_SESSIONS.add(tok)
        return self._send(302, "", extra={
            "Location": "/",
            "Set-Cookie": f"{_SESSION_COOKIE}={tok}; Path=/; HttpOnly; SameSite=Lax",
        })

    # ---- GET routes: each takes (self, q) where q = parse_qs(query string) ----
    def _g_root(self, q):
        return self._send(200, HTML.read_text(encoding="utf-8"), "text/html; charset=utf-8")

    def _g_css(self, q):
        return self._send(200, WEBUI_CSS.read_text(encoding="utf-8"), "text/css; charset=utf-8")

    def _g_js(self, q):
        return self._send(200, WEBUI_JS.read_text(encoding="utf-8"),
                          "text/javascript; charset=utf-8")

    def _g_files(self, q):
        with _lock:
            return self._send(200, {"files": [p.name for p in _state["files"]]})

    def _g_status(self, q):
        return self._send(200, {"img2table": X.HAVE_IMG2TABLE, "ocr": X.HAVE_OCR})

    def _g_scan(self, q):
        return self._send(200, inventory(q["file"][0]))

    def _g_table(self, q):
        return self._send(200, table_detail(q["file"][0], int(q["n"][0])))

    def _g_page(self, q):
        # cap raised from 3.0 -> 4.5 (72*4.5 = 324 DPI) so the preview can
        # actually render sharp on a high-DPI/retina display when the client
        # asks for scale*devicePixelRatio -- 3.0 (216 DPI) was the ceiling
        # even on a 1x display asking for "sharp"
        scale = max(1.0, min(float(q.get("scale", ["2"])[0]), 4.5))
        return self._send(200, page_png(q["file"][0], int(q["n"][0]), scale), "image/png")

    def _g_pagecount(self, q):
        return self._send(200, {"pages": page_count(q["file"][0])})

    def _g_page_raw(self, q):
        scale = max(1.0, min(float(q.get("scale", ["2"])[0]), 4.5))
        bbox_raw = q.get("bbox", [""])[0]
        bbox = None
        if bbox_raw:
            try:
                bbox = [float(v) for v in bbox_raw.split(",")]
            except ValueError:
                bbox = None
        return self._send(200, page_raw_png(q["file"][0], int(q["n"][0]), scale,
                                            q.get("hl", [""])[0], bbox), "image/png")

    def _g_search(self, q):
        return self._send(200, {"hits": search_pdf(q["file"][0], q.get("q", [""])[0])})

    def _g_quickfind(self, q):
        return self._send(200, {"hits": quick_find_statements(q["file"][0])})

    def _g_telecom_candidates(self, q):
        return self._send(200, telecom_candidates(q["file"][0]))

    GET_ROUTES = {
        "/": _g_root, "/index.html": _g_root,
        "/webui.css": _g_css, "/webui.js": _g_js,
        "/api/files": _g_files, "/api/status": _g_status,
        "/api/scan": _g_scan, "/api/table": _g_table,
        "/api/page": _g_page, "/api/pagecount": _g_pagecount,
        "/api/page_raw": _g_page_raw, "/api/search": _g_search,
        "/api/quickfind": _g_quickfind,
        "/api/telecom_candidates": _g_telecom_candidates,
        "/login": _g_login,
    }

    def do_GET(self):
        _debug_state_line(self.path)
        u = urlparse(self.path)
        if u.path != "/login" and not _is_authed(self):
            return self._send(302, "", extra={"Location": "/login"})
        handler = self.GET_ROUTES.get(u.path)
        if handler is None:
            return self._send(404, {"error": "not found"})
        try:
            return handler(self, parse_qs(u.query))
        except KeyError:
            return self._send(404, {"error": "unknown file"})
        except Exception as e:
            LOG.exception("GET %s failed", self.path)
            return self._send(500, {"error": f"{type(e).__name__}: {e}"})

    # ---- POST routes: each takes (self, payload) -- the parsed JSON body.
    # Endpoints in _NEEDS_FILE get a shared "no file selected" 400 (below,
    # before dispatch) instead of each repeating the same check; /api/compare
    # has its own two-file version of that check since its payload shape
    # (a/b) doesn't fit the shared one-file check. ----
    def _p_export(self, payload):
        data, fn = export_xlsx(payload["file"], [int(x) for x in payload["ns"]],
                               payload.get("edits"))
        return self._send(
            200, data, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            {"Content-Disposition": f'attachment; filename="{fn}"'})

    def _p_reanalyze(self, payload):
        return self._send(200, reanalyze(
            payload["file"], int(payload["n"]), payload["rows"], payload.get("title")))

    def _p_compare(self, payload):
        if not payload.get("a") or not payload.get("b") \
                or _path(payload["a"]) is None or _path(payload["b"]) is None:
            return self._send(400, {"error": "No file selected -- upload a PDF first."})
        return self._send(200, compare(
            payload["a"], int(payload["na"]), payload["b"], int(payload["nb"])))

    def _p_upload(self, payload):
        return self._send(200, {"file": upload_pdf(payload.get("filename"), payload["data_b64"])})

    def _p_extract_region(self, payload):
        bbox = [float(v) for v in payload["bbox"]]
        d = extract_region(payload["file"], int(payload["page"]), bbox,
                           payload.get("title"), grid=payload.get("grid"))
        if d is None:
            no_text = not region_has_text(payload["file"], int(payload["page"]), bbox)
            return self._send(422, {
                "error": ("No text found in that box." if no_text else
                          "No table found in that box -- try drawing it tighter "
                          "around just the table's rows and columns."),
                "no_text": no_text, "ocr_available": X.HAVE_OCR})
        return self._send(200, d)

    def _p_detect_grid(self, payload):
        # Always 200 -- "nothing detected near this box" is the routine,
        # expected outcome of an exploratory drag (a rough box over blank
        # margin, or a page img2table can't parse), not a request error;
        # see detect_grid's own docstring. Reserve non-2xx for genuine
        # request problems (bad file/page/bbox), which the generic
        # do_POST try/except below already covers.
        bbox = [float(v) for v in payload["bbox"]]
        d = detect_grid(payload["file"], int(payload["page"]), bbox)
        return self._send(200, d)

    def _p_extract_region_ocr(self, payload):
        d = extract_region_ocr(payload["file"], int(payload["page"]),
                               [float(v) for v in payload["bbox"]], payload.get("title"))
        if d is None:
            return self._send(422, {"error": "OCR found nothing table-like in that box."})
        return self._send(200, d)

    def _p_delete_manual(self, payload):
        try:
            delete_manual(payload["file"], int(payload["n"]))
        except KeyError:
            return self._send(404, {"error": "that table is already gone"})
        return self._send(200, inventory(payload["file"]))

    def _p_undelete_manual(self, payload):
        n = undelete_manual(payload["file"])
        if n is None:
            return self._send(404, {"error": "nothing to undo"})
        inv = inventory(payload["file"])
        inv["restored_n"] = n
        return self._send(200, inv)

    POST_ROUTES = {
        "/api/export": _p_export, "/api/reanalyze": _p_reanalyze, "/api/compare": _p_compare,
        "/api/upload": _p_upload, "/api/extract_region": _p_extract_region,
        "/api/extract_region_ocr": _p_extract_region_ocr,
        "/api/delete_manual": _p_delete_manual, "/api/undelete_manual": _p_undelete_manual,
        "/api/detect_grid": _p_detect_grid,
    }
    # every one of these needs a real, already-uploaded file -- catch "no
    # file selected yet" with one clear message instead of a bare KeyError
    # leaking to the browser (reachable from the UI: every sidebar control
    # used to be clickable before a file was ever chosen)
    _NEEDS_FILE = {"/api/reanalyze", "/api/extract_region", "/api/extract_region_ocr",
                  "/api/export", "/api/delete_manual", "/api/undelete_manual",
                  "/api/detect_grid"}

    def do_POST(self):
        _debug_state_line(self.path)
        u = urlparse(self.path)
        if u.path == "/login":
            # the one POST that must work before any session exists, so it
            # runs before both the origin check and the auth gate below (its
            # own protection is the passcode comparison inside _p_login) --
            # and it's form-encoded, not JSON, so it skips the usual
            # json.loads(body) too.
            length = int(self.headers.get("Content-Length", "0"))
            return self._p_login(self.rfile.read(length) or b"")
        if not _is_authed(self):
            return self._send(302, "", extra={"Location": "/login"})
        # reject cross-origin POSTs -- this is an open localhost endpoint that
        # reads/writes files; only our own page (or a no-Origin client) may
        # post -- UNLESS the request already proved itself via the passcode
        # gate above (_is_authed already returned True, e.g. through a
        # deliberately-opened remote-access tunnel), in which case the
        # passcode already established that this is trusted.
        origin = self.headers.get("Origin")
        if origin and urlparse(origin).hostname not in ("127.0.0.1", "localhost") \
                and not _REMOTE_PASSCODE:
            return self._send(403, {"error": "cross-origin POST refused"})
        handler = self.POST_ROUTES.get(u.path)
        if handler is None:
            return self._send(404, {"error": "not found"})
        try:
            length = int(self.headers.get("Content-Length", "0"))
            payload = json.loads(self.rfile.read(length) or b"{}")
            if u.path in self._NEEDS_FILE:
                name = payload.get("file")
                if not name or _path(name) is None:
                    return self._send(400, {"error": "No file selected -- upload a PDF first."})
            return handler(self, payload)
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


def _discover_files(pdf_args):
    """CLI-arg files (or every *.pdf in a given directory) PLUS whatever a
    previous run already uploaded to UPLOAD_DIR -- without the latter,
    launching with no arguments (the normal run_app.bat path) always starts
    on a blank onboarding screen even when uploads/ still has last session's
    PDF sitting right there, which would make session persistence pointless
    (the file wouldn't even be listed). De-duplicated by NAME, not just
    resolved path (found live: launching with a directory that happens to
    share a filename with something uploaded in an earlier session -- e.g.
    the CLI-arg copy of a report AND an uploads/ copy of the same report,
    genuinely different files at genuinely different paths -- listed twice
    under the identical display name; _path()'s name-based lookup always
    resolves the FIRST one, so the second was already a dead, confusing
    entry in the dropdown, never actually reachable). CLI-arg order wins."""
    files = []
    for a in pdf_args:
        p = Path(a)
        if p.is_dir():
            files += sorted(p.glob("*.pdf"))
        elif p.exists() and p.suffix.lower() == ".pdf":
            files.append(p)
        else:
            LOG.warning("skipping %s (not a PDF)", a)
    seen_names = {f.name for f in files}
    if UPLOAD_DIR.is_dir():
        for p in sorted(UPLOAD_DIR.glob("*.pdf")):
            if p.name not in seen_names:
                files.append(p)
                seen_names.add(p.name)
    return files


def run(pdf_args, host="127.0.0.1", port=None, open_browser=True, debug=False,
        allow_remote=False):
    global _REMOTE_PASSCODE
    if allow_remote:
        # short (6 chars) is fine here -- it's not the only protection, the
        # tunnel URL itself is also an unguessable secret, this is layered
        # on top for the case where the URL alone leaks (browser history,
        # a referrer header, screen-sharing the address bar by accident)
        _REMOTE_PASSCODE = secrets.token_urlsafe(6)
    handlers = [logging.StreamHandler()]
    if debug:
        # persists past the terminal scrolling away -- if something acts up
        # partway through a session, the file (not just stdout) is what a
        # user would actually be able to send along
        fh = logging.FileHandler(ROOT / "debug.log", mode="a", encoding="utf-8")
        fh.setFormatter(logging.Formatter("%(asctime)s %(name)s %(levelname)s %(message)s"))
        handlers.append(fh)
    logging.basicConfig(level=logging.DEBUG if debug else logging.INFO,
                        format="  %(message)s", handlers=handlers, force=True)
    if debug:
        # basicConfig's DEBUG level cascades to every logger in the process
        # that doesn't have its own override -- including pdfminer (the
        # library under pdfplumber), which logs every single parse token/
        # seek/keyword at DEBUG. Measured live: a 20-minute, 26-file
        # stress session produced a 4.3 GB, 44.5-million-line debug.log,
        # nearly all of it pdfminer's own byte-level parse trace -- not
        # this project's own ~5 debug statements, and not useful for
        # diagnosing anything this flag is actually for. Capped back to
        # WARNING so --debug stays a log someone could actually read.
        logging.getLogger("pdfminer").setLevel(logging.WARNING)
        LOG.info("debug logging ON -- also writing to %s", ROOT / "debug.log")
    files = _discover_files(pdf_args)
    _state["files"] = files
    for f in files:
        _load_session(f.name)
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
    if _REMOTE_PASSCODE:
        print(f"\n  REMOTE ACCESS ON -- passcode required for every request: {_REMOTE_PASSCODE}")
        print("  (only share this passcode the same way you'd share the tunnel URL itself --")
        print("   anyone who has both can use this exactly as you can, including extracting)")
    print("  (Ctrl-C to stop)\n")
    if open_browser:
        threading.Timer(0.6, lambda: webbrowser.open(url)).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\n  stopped.")


def _cli(argv=None):
    import argparse
    ap = argparse.ArgumentParser(
        prog="serve.py",
        description="Local preview / edit / export UI for extract_all_tables.py.")
    ap.add_argument("files", nargs="*",
                    help="PDF file(s) or a directory of PDFs to preload")
    ap.add_argument("--port", type=int, default=None)
    ap.add_argument("--no-browser", action="store_true",
                    help="don't open a browser tab on start")
    ap.add_argument("--debug", action="store_true",
                    help="verbose logging (a per-request server-state line, "
                         "plus tablekit.extract's own debug output) to the "
                         "console AND debug.log next to this script")
    ap.add_argument("--allow-remote", action="store_true",
                    help="require a one-time passcode (printed at startup) before ANY "
                         "request -- for deliberately exposing this server through a "
                         "tunnel (e.g. a short remote demo). Off by default; local-only "
                         "usage is completely unaffected either way.")
    args = ap.parse_args(argv)   # argv=None -> argparse's own sys.argv[1:] default
    run(args.files, port=args.port, open_browser=not args.no_browser, debug=args.debug,
        allow_remote=args.allow_remote)


if __name__ == "__main__":
    _cli()
