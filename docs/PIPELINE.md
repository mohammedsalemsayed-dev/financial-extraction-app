# How the extractor works (maintainer's map)

Three ways a table reaches the pipeline: **automatic** (`scan()` walks the
whole document, detects and classifies every table on its own -- no UI
button for this any more, see below), **manual** (`extract_region()` — the
user draws a box in the UI around one table; no detection/classification-
from-scratch needed) and **manual + OCR** (`extract_region_ocr()` — same
box-draw UI, offered only when the box has no text layer at all, i.e. a
scanned page). All three are run through the exact same `analyze()`/health
pipeline afterward, so every mode produces identically-verified output.

```
PDF ─► scan()                                    [extract_all_tables.py]
        │  per page:
        │   find_all_tables(page)                 detection
        │     ├─ _recon_tables(page)              borderless / 2-up statements
        │     │    ├─ telecom_extract geometry     regex/word-position reconstruction (rank 0)
        │     │    └─ img2table challenger         OpenCV table detection; only competes
        │     │         (tablekit/img2table_backend) when the regex candidate does NOT
        │     │                                     already reconcile (see CONFIG/safety note)
        │     ├─ pdfplumber LINES / LINES_H / TEXT rank 1-3
        │     ├─ dedupe by bbox-overlap + exact-rows
        │     └─ _split_stacked_statements        cut a merged 2-statement block
        │   _clean(rows)                          coerce_cell + _desect + _deprose
        │   guess_title(page, bbox)               nearest heading
        │   analyze(t)                            the semantic reading  ▼
        ├─ _stitch_page_breaks(all_tables)        merge a statement running onto p+1
        ├─ cross_year_check(all_tables)           A.prev vs B.cur (needs >1 file)
        └─ _attach_health(t)                      label_health + figure_health

PDF ─► extract_region(pdf, page, bbox)            [extract_all_tables.py -- manual box-select]
        │   img2table_page_tables(pdf, page)       run img2table on the WHOLE page, uncropped
        │   rows_in_box(page_tables, bbox)         [tablekit/img2table_backend.py]
        │     ├─ candidate region(s) with >=50%    reject a region only clipped at the box's
        │     │  overlap with the drawn box          edge (stops a generous box from dragging
        │     │                                       in an unrelated neighbouring table)
        │     ├─ _merge_side_by_side / _merge_stacked   re-join img2table's own split pieces
        │     │    (X-aligned only, since 0.4.0 -- stacking must not glue two DIFFERENT
        │     │     notes together just because they're Y-adjacent with matching column count)
        │     └─ trim rows to the box's Y-range     protects against a stacking merge pulling
        │          (+ small margin), report bbox     in the START of the next section; the
        │          from the SURVIVING rows only      highlight shown back is always faithful
        │                                             to what was actually extracted
        │   (falls back to pdfplumber's ruled-table extractor if img2table isn't installed
        │    or finds nothing -- ruled tables only, no borderless support in that fallback)
        │   analyze(t) / _attach_health(t)          same verification as the automatic path

PDF ─► extract_region_ocr(pdf, page, bbox)        [extract_all_tables.py -- OCR failsafe]
        │   only offered when the drawn box has NO extractable text at all (a scan/image)
        │   ocr_rows_in_box(page, bbox)             [tablekit/img2table_backend.py]
        │     pad the crop (20pt) before running Tesseract -- a tight crop clips/misreads
        │     digits; needs pytesseract + the Tesseract binary (X.HAVE_OCR)
        │   analyze(t) / _attach_health(t)          same verification as every other path

analyze(t):
   _statement_kind(title, rows)                   income / BS / cash flow / SOCE / note / table
   header row + years                             + year-reversal / segmental notes
   value columns  (+ a note-ref column)
   total / subtotal rows
   foot verdict, PER COLUMN:
     income / note   -> _reconciles / reconcile_explain   (cut at "profit for the year")
     balance sheet   -> total identities (assets == equity+liab, or assets-side)
     cash flow       -> _cashflow_foots  (op+inv+fin==net, or open+net+fx==close,
                                          or du-style: op + Σ investing + fin == net)
     changes in equity -> _equity_foots  (close == open + Σ movements)
   downgrade to "table" if:
     _is_structural_non_statement()   transition grid / >=4 cols / EBITDA-margin
     or (not footing and _looks_like_not_a_statement())   few anchors / prose / bad years

report:
   build_workbook(tables)      one sheet per table + Contents      [extract_all_tables.py]
   build_compare_workbook()    row-level diff, --compare
   serve.py + webui.html       local preview / edit / one-workbook export
                                ONE entry point into the table list from the
                                UI: the page-picker box-draw (extract_region(),
                                the landing view; can extract more than one
                                table per page without leaving the picker).
                                scan() itself is NOT exposed as a UI action any
                                more -- the whole-file "Auto-detect all tables"
                                button was pulled (unreliable, not worth
                                shipping); scan() is still called server-side
                                by /api/scan, which only the Compare panel uses
                                (to inventory a SECOND loaded file's tables so
                                it can find the matching statement) -- see
                                serve.py's `_resolve`.
```

## img2table: required, not optional

`img2table` / `opencv-python-headless` / `pandas` (`requirements.txt`) power
the auto-detect challenger AND are the manual box-select mode's *entire*
extraction path for borderless tables -- without them, manual mode silently
falls back to `page.crop().extract_table()` (ruled tables only). `serve.py`
prints `img2table engine: ON/OFF` on startup and the UI shows an "img2table:
on/off" badge in the header; `extract_all_tables.HAVE_IMG2TABLE` is the flag
to check in code. If you see `img2table: off`, install the group above
before assuming a bug.

## Where to change things

| symptom | look at |
|---|---|
| a statement isn't detected (auto-detect) | `_recon_tables` heading regex `_RECON_HEAD_RE`; `find_all_tables` dedupe |
| a manually-drawn box returns nothing / wrong content | `tablekit/img2table_backend.py` `rows_in_box` (`min_overlap_frac`, `row_margin`); `_merge_stacked`'s `x_tol` |
| OCR misreads / drops digits from a scanned box | `tablekit/img2table_backend.py` `ocr_rows_in_box`'s `pad` (crop margin before Tesseract runs -- too tight clips characters) |
| a real statement listed as "table" | `_is_structural_non_statement`, `_looks_like_not_a_statement` |
| wrong "foots" verdict | the per-kind block in `analyze`; `_reconciles` / `_cashflow_foots` tolerances in `CONFIG` |
| a figure is wrong | `tablekit/parse.py` `parse_number` + `test_parse_number` |
| two adjacent numbers glued into one cell | `tablekit/img2table_backend.py` `_split_glued_cell` / `_normalize_row_width` |
| a label has prose in it | `_deprose_labels`, `_desect_labels`; check `label_health` flags it |
| a threshold needs tuning | `tablekit/config.py` `CONFIG` (or `extract_all_tables.toml`) |
| the UI | `serve.py` endpoints + `webui.html` (single file) |
| the highlighted box on the preview page looks wrong | it's derived from the actually-extracted rows, not the drawn input or a heuristic guess -- see `extract_region`'s `shown_bbox` and `rows_in_box`'s docstring |

## Tests

`tablekit_tests/` — `test_units.py` + `test_serve.py` + `test_manual_mode.py`
(no PDFs needed for the pure-logic parts, CI-gated), `test_golden.py`
(`golden.json` snapshot + `anchors.json` hand-verified figures + `parse_number`
cases). `test_manual_mode.py` also has PDF-gated cases (img2table region
merge/select regressions, the full upload→extract→export HTTP round trip)
that skip cleanly when the sample PDFs aren't present, same as golden/anchors.
`python tablekit_tests/snapshot.py` regenerates the golden snapshot after a
*deliberate, verified* change.
