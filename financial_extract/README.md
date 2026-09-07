# financial_extract — local extraction app (v0.1, MVP shell)

A localhost-only UI around the deterministic geometry engine in
`../telecom_extract.py`. It locates two tables in an annual-report PDF —
the consolidated income statement and the operating-expenses note — checks
that the figures reconcile arithmetically, and exports to Excel.

**No cloud AI. No local LLM. No OCR. No network access for document content.**
The backend binds to `127.0.0.1` only.

## Run

```
cd "table extraction from pdf"
python -m financial_extract           # opens http://127.0.0.1:8765/
python -m financial_extract --port 9000 --no-browser
```

Requires the same environment as `telecom_extract.py` (`pdfplumber`,
`openpyxl`, `pypdf`, `pillow`, `pypdfium2` — all already installed).

## What it does

| Step | |
|---|---|
| **Add PDFs** | drag onto the window, or *choose files*. Company + year + page count are detected from the file name. |
| **Extract** | click a document. The engine scans for both target tables, reconstructs rows/columns from page geometry, and runs the arithmetic check. |
| **Review** | the source PDF page is shown beside the extracted table. A green **PASS** / red **FAIL** banner shows whether the figures foot. |
| **Edit** | every cell is editable. 0.7 s after an edit the arithmetic check re-runs and the banner updates. |
| **Export** | *Export Excel* writes one sheet per document (Note column + year/unit header + reconciliation banner), matching the format of `du_two_tables.xlsx`. Edits are applied. |

## Layout

```
financial_extract/
  core.py      engine wrapper: detect(), extract_document(), recheck(),
               page_png(), build_workbook()  — imports telecom_extract unchanged
  server.py    stdlib http.server, JSON API + static files, 127.0.0.1 only
  __main__.py  entry point
  web/         index.html · style.css · app.js   (no build step)
```

## Endpoints (all localhost)

| Method | Path | |
|---|---|---|
| POST | `/api/upload` | raw PDF body + `X-Filename` header → `{doc_id, company, year, pages}` |
| POST | `/api/extract` | `{doc_id}` → structured extraction for both targets |
| POST | `/api/recheck` | `{reconcile_kind, body}` → re-run the arithmetic check on an edited body |
| GET | `/api/page?doc_id=&page=` | PNG of that PDF page |
| POST | `/api/export` | `{documents:[...]}` → `.xlsx` |

## Scope (v0.1)

Covers **du** and **Etisalat / e&** — the two companies the engine profiles
exist for. Adding a third company is Stage 1.3 of the build plan: a
`profiles/company_c.json` plus whatever new layout code its reports need.

Not yet built (see the build plan): cell → PDF click-through highlight,
batch queue, Windows installer, session persistence across restarts.
