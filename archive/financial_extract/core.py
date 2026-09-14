"""
Engine wrapper.  Imports telecom_extract unchanged and exposes structured
(JSON-friendly) results plus an Excel writer that honours in-app edits.

Nothing here does OCR, calls a model, or touches the network.
"""
from __future__ import annotations

import io
import re
import sys
from pathlib import Path

# telecom_extract.py lives in the project root, one level up from this package
_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

import pdfplumber  # noqa: E402
from openpyxl import Workbook  # noqa: E402

import telecom_extract as te  # noqa: E402


# --------------------------------------------------------------------------
# identification
# --------------------------------------------------------------------------
def _display_name(pdf_path: Path) -> str:
    """Strip the internal '<doc_id>__' upload prefix for display."""
    return re.sub(r"^[0-9a-f]{6,16}__", "", pdf_path.name)


def detect(pdf_path: Path) -> dict:
    """Company + year + page count for one PDF, from its file name and a
    cheap page-count read."""
    disp = _display_name(pdf_path)
    key, prof = te.profile_for_file(Path(disp))
    ym = re.search(r"\b(19|20)\d{2}\b", disp)
    year = int(ym.group()) if ym else None
    pages = None
    try:
        with pdfplumber.open(pdf_path) as pdf:
            pages = len(pdf.pages)
    except Exception:
        pass
    return {
        "company_key": key,
        "company": prof["label"],
        "year": year,
        "pages": pages,
    }


# --------------------------------------------------------------------------
# extraction
# --------------------------------------------------------------------------
def _jsonify_cell(v):
    if v is None:
        return None
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, (int, float)):
        return v
    return str(v)


def _nonempty(row) -> bool:
    return any(c is not None and str(c).strip() != "" for c in row)


def _status_of(rec) -> str:
    if not rec or not rec.get("cols"):
        return "NOT_CHECKED"
    return "PASS" if rec["ok"] else "FAIL"


def extract_document(pdf_path: Path) -> dict:
    """Run both target extractions for one PDF and return a structured
    result that mirrors what convert_pdf() would have written to a sheet."""
    disp = _display_name(pdf_path)
    disp_stem = disp[:-4] if disp.lower().endswith(".pdf") else disp
    key, prof = te.profile_for_file(Path(disp))
    ym = re.search(r"\b(19|20)\d{2}\b", disp_stem)
    report_year = int(ym.group()) if ym else None

    read_path = te.normalize_pdf_for_reading(pdf_path) if te.HAVE_PYPDF else pdf_path
    out_targets = []
    with pdfplumber.open(read_path) as pdf:
        total_pages = len(pdf.pages)
        page_labels = te.load_page_labels(pdf_path)
        for target in prof["targets"]:
            order = te.page_scan_order(total_pages, None)
            header, body, page_num, score, actual_heading, rec, note_col = \
                te.find_best_matching_table(pdf, order, target)

            if body is not None and len(note_col) >= 2:
                nb = []
                for r in body:
                    r = list(r)
                    lbl = te.row_label(r)
                    ref = note_col.get(te.normalize(lbl).lower()) if lbl else None
                    nb.append([r[0] if r else None, ref] + (r[1:] if r else []))
                body = nb

            if body is None:
                msg = "Nothing matched the target vocabulary above the confidence threshold."
                if target.get("mode") == "full_table":
                    msg = ("No separate operating-expenses note in this report. "
                           "Some du reports (2014-2018) and du 2025 print the expense "
                           "breakdown by nature on the face of the income statement "
                           "— see the statement above.")
                out_targets.append({
                    "name": target["name"],
                    "mode": target.get("mode"),
                    "reconcile_kind": target.get("reconcile"),
                    "found": False,
                    "message": msg,
                    "status": "NOT_FOUND",
                    "summary": "NOT FOUND",
                    "page": None,
                    "page_label": None,
                    "score": 0.0,
                    "heading": None,
                    "header": [],
                    "body": [],
                })
                continue

            header_rows = te._clean_header(header, body, report_year)
            if header_rows and not isinstance(header_rows[0], (list, tuple)):
                header_rows = [header_rows]
            body = [r for r in body if _nonempty(r)]
            status = _status_of(rec)
            out_targets.append({
                "name": target["name"],
                "mode": target.get("mode"),
                "reconcile_kind": target.get("reconcile"),
                "found": True,
                "status": status,
                "summary": te.recon_summary(rec),
                "reconcile": rec,
                "page": page_num,
                "page_label": te.display_page(page_labels, page_num),
                "total_pages": total_pages,
                "score": round(float(score), 4),
                "heading": actual_heading,
                "header": [[_jsonify_cell(c) for c in row] for row in header_rows],
                "body": [[_jsonify_cell(c) for c in row] for row in body],
            })

    di = detect(pdf_path)
    return {
        "file": disp,
        "stem": disp_stem,
        "company_key": key,
        "company": prof["label"],
        "year": report_year,
        "pages": di["pages"],
        "targets": out_targets,
    }


# --------------------------------------------------------------------------
# re-check after an edit
# --------------------------------------------------------------------------
def _coerce_row(row, keep_text_cols=()):
    """Parse figure cells the way the rest of the engine does.  Columns in
    `keep_text_cols` (e.g. the Note-reference column) are left verbatim."""
    out = []
    for i, c in enumerate(row):
        if i in keep_text_cols:
            out.append(c if (c is not None and str(c).strip() != "") else None)
        elif c is None or c == "":
            out.append(None)
        elif isinstance(c, (int, float)):
            out.append(c)
        else:
            out.append(te.clean_cell(str(c)))
    return out


def recheck(reconcile_kind: str, body: list) -> dict:
    """Re-run the arithmetic check on an edited body."""
    coerced = [_coerce_row(r) for r in body]
    target = {"reconcile": reconcile_kind}
    rec = te.reconcile(target, coerced)
    return {
        "status": _status_of(rec),
        "summary": te.recon_summary(rec),
        "reconcile": rec,
    }


# --------------------------------------------------------------------------
# page image
# --------------------------------------------------------------------------
# 1x1 transparent PNG, returned when a page cannot be rendered
_BLANK_PNG = bytes.fromhex(
    "89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4"
    "890000000d4944415478da6364f80f00010101005a4d6ff40000000049454e44ae426082")


def page_png(pdf_path: Path, page_1based: int, resolution: int = 120) -> bytes:
    # render from the ORIGINAL file — the pypdf-normalised copy used for text
    # extraction sometimes will not open in the image renderer
    try:
        with pdfplumber.open(pdf_path) as pdf:
            idx = max(0, min(page_1based - 1, len(pdf.pages) - 1))
            im = pdf.pages[idx].to_image(resolution=resolution)
            buf = io.BytesIO()
            im.save(buf, format="PNG")
            return buf.getvalue()
    except Exception:
        return _BLANK_PNG


# --------------------------------------------------------------------------
# Excel export (honours edits)
# --------------------------------------------------------------------------
def build_workbook(documents: list) -> bytes:
    """documents: list of dicts shaped like extract_document() output, but
    each target may carry an edited `header` / `body`.  One sheet per
    document, matching the delivered workbook layout."""
    wb = Workbook()
    wb.remove(wb.active)
    for doc in documents:
        ws = wb.create_sheet(title=te.safe_sheet_name(doc["stem"]))
        row = 1
        row = te.write_title(
            ws, row, doc["stem"].replace("-", " ").replace("_", " ").title(), 6)
        row = te.write_subtitle(
            ws, row,
            f"Extracted from PDF · {doc['file']}  ({doc['company']} profile)")
        row += 1
        max_col = 1
        for t in doc["targets"]:
            if not t.get("found"):
                row = te.write_section_label(ws, row, f"{t['name']} — NOT FOUND", 6)
                cell = ws.cell(row=row, column=1, value=t.get("message", ""))
                cell.font = te.NOT_FOUND_FONT
                cell.fill = te.NOT_FOUND_FILL
                row += 2
                continue

            called = ""
            if t.get("heading") and te.label_similarity(t["heading"], t["name"]) < 0.6:
                called = f'  — report calls it "{t["heading"]}"'
            row = te.write_section_label(
                ws, row,
                f"{t['name']}  —  page {t['page_label']}, "
                f"content match {t['score']:.0%}{called}", 6)

            # recompute the banner from the (possibly edited) body
            rc = recheck(t["reconcile_kind"], t["body"])
            banner = ws.cell(row=row, column=1, value=f"Arithmetic check:  {rc['summary']}")
            banner.font = te.NOTE_FONT if rc["status"] == "PASS" else te.NOT_FOUND_FONT
            if rc["status"] == "FAIL":
                banner.fill = te.NOT_FOUND_FILL
            row += 1

            header_rows = t.get("header") or []
            # if column 1 is the "Note" reference column, keep its cells as text
            keep = ()
            for hr in header_rows:
                if len(hr) > 1 and str(hr[1]).strip().lower() == "note":
                    keep = (1,)
                    break
            body = [_coerce_row(r, keep) for r in t.get("body", []) if _nonempty(r)]
            row, ncols = te.write_table(ws, row, header_rows, body)
            max_col = max(max_col, ncols)
            row += 1

        ws.freeze_panes = "A4"
        te.autofit_columns(ws, max_col)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
