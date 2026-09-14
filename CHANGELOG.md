# Changelog

## 0.5.0

### "Auto-detect all tables" fully removed
- The opt-in button added in 0.4.0 is gone: false positives/misses weren't
  reliable enough to ship, so manual box-select (plus its OCR failsafe) is
  now the *only* way a table reaches the UI. `scan()` itself is unchanged
  and still runs server-side -- it's what `/api/scan` uses to inventory a
  second loaded file for the Compare panel -- but there is no UI action left
  that triggers a whole-document scan. `serve.py`'s `_resolve` documents the
  seam this leaves for tests.

### Extraction bug fixes
- `rows_in_box()`'s region-overlap check compared overlap only against the
  drawn box's own area, which rejected a box deliberately drawn around just
  *part* of a bigger region (e.g. a balance sheet's assets half only) and
  silently dropped its header row. Now takes the max of overlap-vs-region and
  overlap-vs-box.
- OCR crop padding raised 6pt -> 20pt (`ocr_rows_in_box`): a tight box was
  clipping/misreading trailing digits ("$20,565,087" -> "$20,565,C").

### `%` / currency symbols preserved in the grid
- `tablekit/parse.py`'s `FormattedNumber` (a `float` subclass carrying
  `prefix`/`suffix`) lets `parse_number` / `coerce_cell` keep `%` and
  currency symbols for *display* without changing how a value behaves
  anywhere else -- footing, health-scoring and delta math all still see a
  plain float. `serve.py` sends the formatting as a parallel `fmt` field;
  `webui.html`'s `fmt()` helper re-applies it only to the shown text.

### Theming, localization, and UI polish
- Full light/dark theming via CSS custom properties, matching the CPI
  Automation Platform's cream/navy/gold palette, plus complete English/Arabic
  (RTL) localization of the UI chrome (`I18N` dict + `t()` / `applyI18n()` in
  `webui.html`). Extracted table content and server-generated verdict text
  are deliberately NOT translated (see the comment above `I18N`).
- New onboarding screen with a working drag-and-drop upload zone.
- Export tray flags at-risk tables (no foot / low health) in red.
- Compare panel: color-coded verdict banner (restated vs clean), card-styled
  diff table, inline "RESTATED" badge on the affected row.
- `.design-mockup/*.dc.html` — design-canvas mockup sources these screens
  were drawn from (main grid, upload, export review, compare), kept for
  future updates; the seeded/published canvas output itself is gitignored
  (a multi-MB copy of the design tool's own editor, not app source).

### Housekeeping
- `run_app.bat` was still launching `python -m financial_extract` -- the
  pre-rewrite MVP shell, superseded since 0.4.0 by `serve.py` + `webui.html`
  + `tablekit/` but never repointed. Fixed to launch `serve.py`.
- `financial_extract/` (the MVP shell it launched) moved to
  `archive/financial_extract/` -- same reasoning as `extract_two_tables.py`
  below: nothing imports it, and leaving it at the repo root read as a live
  alternative to the real app when it wasn't one.
- `extract_all_tables.toml.example` was missing `figure_outlier_sd` (added
  in 0.4.0's figure-health work) despite claiming to list every recognised
  key.
- `EXTRACT_ALL_TABLES.md` / `docs/PIPELINE.md` updated: both still described
  "Auto-detect all tables" as a live UI action and neither mentioned the OCR
  failsafe path at all.

## 0.4.0

### Second geometry engine
- `tablekit/img2table_backend.py` — img2table (OpenCV) as a challenger to the
  regex/word-position reconstruction for borderless and 2-up statements.
  Only competes when the regex candidate does NOT already reconcile
  ("foots passing" alone isn't sufficient correctness evidence -- unrestricted
  competition was tried and reverted after it occasionally replaced an
  already-correct table with a coincidentally-passing worse one).
- Fixed along the way: `parse_number` no longer misreads the "AED 000" /
  "USD 000" units-disclaimer as a literal zero; Y-coordinate row alignment
  (not index-zipping) when img2table splits a table into separate
  labels/figures regions; closest-heading assignment so a region on a 2-up
  page isn't attributed to the wrong statement; glued-cell splitting when a
  stacked-merge narrows a row's column count (`_split_glued_cell`); stacked
  merges now also require the pieces to be horizontally aligned, not just
  Y-adjacent with a matching column count (stopped two unrelated notes from
  being glued into one table on a dense notes page).

### Manual box-select mode ("Tabula but better")
- New default UI flow: upload a PDF, browse its pages, drag a box around a
  table, extract just that region -- runs through the SAME
  classify/analyze/health pipeline as automatic detection, so foots
  verification and the editable grid work identically either way.
  `extract_all_tables.extract_region()` / `tablekit/img2table_backend.rows_in_box`.
- Multiple tables can be extracted from the same page without leaving the
  picker (a confirmation toast replaces the old "jump to a different screen"
  flow, which made a second selection from the same page hard to find).
- The highlighted box shown back on the original-page preview is derived
  from the actually-extracted rows (trimmed to the drawn box's Y-range, with
  the matched region's own tight extent) -- not the raw drawn input, and not
  unioned with it either. Both under- and over-sized-looking highlights were
  real bugs (not just cosmetic) and are fixed: a region only clipped at the
  box's edge no longer qualifies as a match (`min_overlap_frac`), preventing
  a generously-drawn box from dragging in a neighbouring note's heading.
- "Auto-detect all tables" is now an explicit, opt-in action -- no longer
  run automatically on file load, which used to block the UI for the time it
  took to scan the whole document.
- PDF text search (`/api/search`) with results highlighted both in the
  snippet list and directly on the page image (`page.search()` bounding
  boxes), click a result to jump to that page.
- Resizable sidebar (drag the splitter, persisted) and resizable
  page/table panes (native corner-drag).

### Packaging / observability
- `img2table` / `opencv-python-headless` / `pandas` are now declared in
  `requirements.txt` and `pyproject.toml` -- previously installed ad hoc and
  undeclared, so a clean install silently lost the img2table engine (both
  the auto-detect challenger and manual mode's entire borderless-table
  path) with no error anywhere. `serve.py` now prints `img2table engine:
  ON/OFF` on startup, the CLI prints an equivalent note, and the UI shows an
  "img2table: on/off" badge in the header -- this state used to be
  invisible.

### Tests
97 → 117+: `tablekit_tests/test_manual_mode.py` -- img2table_backend region
merge/select logic (pure, no PDF needed, CI-gated) plus PDF-gated regression
cases for the exact bugs found and fixed live this session (the SOCE
glued-cell corruption, the note-22/note-23 cross-contamination), and an
HTTP round trip through upload → extract → export.

### Housekeeping
- `extract_two_tables.py` (the original DFM-only predecessor, unused by
  anything since `extract_all_tables.py` superseded it) and `HANDOFF.md`
  (its project history) moved to `archive/` -- nothing imported them, but
  leaving them at the repo root read as current documentation when it
  wasn't. `telecom_extract.py` stays where it is despite the misleading
  name: it's the generic regex/word-position reconstruction engine used for
  every company now, not telecom-specific, and is load-bearing.

## 0.3.0

### Package / tooling
- Split out `tablekit/` package: `tablekit.config` (CONFIG + TOML loader),
  `tablekit.parse` (`parse_number`, `coerce_cell`). `import extract_all_tables`
  is unchanged.
- `pyproject.toml` with `extract-tables` / `tablekit-serve` entry points.
- `.github/workflows/tests.yml` now runs the pure-logic suites (`test_units`,
  `test_serve`, parser cases) on every push — no PDFs needed. Golden + anchor
  tests run locally and skip cleanly in CI.
- `-v` / `--debug` on the CLI turns on the `tablekit.extract` logger; the
  reconstruction and foot-check paths now log instead of swallowing silently.

### Trust
- **Per-column foot verdict** — `analyze` records the worked arithmetic and an
  ok/fail for *each* value column (`foot_by_col`); a broken prior-year column is
  no longer hidden by a good current-year one. `--audit` and the UI show both.
- **Cross-year consistency** runs in `scan()` whenever >1 file is scanned (not
  just `--compare`); classifies a diff as `restated` (few lines) vs
  `columns likely misaligned` (many).
- **`figure_health`** — thresholds derived from each column's own
  log-magnitude distribution (`mean + N·sd`) rather than fixed multipliers.
- **Year-column sanity** — flags a statement whose header lists years
  oldest-first (figures may be attributed to the wrong year).
- **Segmental / multi-entity** statements get a note that the columns are not a
  comparable time series.
- Cash-flow tolerance now scales with the size of the flows.
- Balance sheet that footed on the assets side only is flagged incomplete.
- `parse_number` guards `float('inf')` / >18-digit runs (fixed an
  `OverflowError` that aborted some whole-document scans).

### Coverage
- `_deprose_labels` now also pulls a known multi-word line item out of the
  *middle* of an interleaved auditor's-report sentence (du 2010–2014).

### UI (`serve.py` + `webui.html`)
- **Row-level editing**: insert / delete / split / merge-up, plus cell edits.
  All cell parsing happens server-side (`/api/reanalyze`); the browser no
  longer re-implements the number parser.
- Edits re-reconcile live — the foot pill and health update as you type.
- **Compare in the UI** (`/api/compare`) — pick a second file, see the diff.
- Rename the sheet, set the workbook filename, reorder the export tray.
- Page-range box → re-scan a subset. Export warns when a selected table has a
  NO FOOT / low-health flag.
- Responsive: panes stack under 880px.
- POST endpoints reject cross-origin requests; scan cache invalidates on file
  mtime; the port falls back if 8765 is taken.

### Tests
92 → 97: added `test_units.py` (analyze / health / cross-year / stitch /
cash-flow), `test_serve.py` (inventory / edits / one-workbook export / HTTP
end-to-end), `anchors.json` (hand-verified figures), plus parser edge cases.

## 0.2.0
- Preview UI, label-health, `--audit` / `--json`, config file, first test suite,
  page-break stitching, fuzzy `--compare`.

## 0.1.0
- Generic detector: list every table, pick, export; `foots` check for income
  statement / balance sheet / cash flow / notes; `--compare` diff.
