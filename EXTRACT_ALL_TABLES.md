# extract_all_tables.py

Generic, company-agnostic. Four things:

1. **List** every table in a PDF, with a reading of what each one is.
2. **Preview** each table next to the original page and **pick** several to export.
3. **Export** the picked tables as **one** Excel workbook (one sheet per table).
4. **Compare** the same tables across two reports and show what changed.

Digital PDFs only (real text layer — **no OCR**). Put the file next to
`telecom_extract.py` (it borrows that file's geometry helpers to recover
borderless / landscape statements; still runs without it, just weaker on those).

### The preview UI

```
python extract_all_tables.py report.pdf --serve        # or: python serve.py report.pdf ...
```

Opens a local page (127.0.0.1 only, nothing leaves the machine). Left: every
detected table, grouped, each with a **foots** pill, a **data-health** bar, and
a red `N≠` chip when the prior-year column disagrees with another loaded
report. Click one to see the **original PDF page** (table region outlined)
above the **extracted table** — numbers right-aligned, Δ / Δ% columns shown,
total rows bold, cells that look wrong tinted amber (label) or red (figure)
with the reason on hover, the per-year worked arithmetic, and a cross-year
consistency banner.

- **Double-click a cell** to fix a wrong label or figure; **hover a row** for
  insert / delete / split / merge-up. The verdict and health **re-compute live**
  as you type — all parsing happens on the server, not in the browser.
- **Swap year columns**, rename the sheet, set the workbook filename.
- Tick several, reorder them in the **Export order** tray, and **Export N →**
  downloads **one** `.xlsx` (one sheet each + a Contents index). Export warns if
  a selected table is NO FOOT / low-health.
- **Compare** panel: pick a second loaded file → row-level diff of the same
  statement.
- `Select all statements`, `j`/`k` or `↑`/`↓` to walk the list, `space` to tick.
- Panes stack on a narrow window.

---

## 1 · List what's in the PDF

```
python extract_all_tables.py report.pdf --list
```
```
  #  page       size  what                             yrs        foots     title
  21    92   39r x 4c  statement of financial position  2024/2023  foots     Consolidated Statement of Financial Position
  22    93   25r x 4c  income statement                 2024/2023  foots     Consolidated Statement of Comprehensive Income
  23    93   11r x 5c  statement of changes in equity   2024/2023            ...
  24    94   29r x 3c  statement of cash flows          2024/2023            ...
  25   110   10r x 3c  table                            2024/2023            Term deposits / Cash and bank balances
  ...
```

For every table it works out:
| column | meaning |
|---|---|
| **what** | income statement · statement of financial position · cash flows · changes in equity · note · table |
| **yrs** | the reporting years it found in the columns |
| **foots** | `foots` = the figures add up · `NO FOOT` = they don't (extraction glitch or a genuine imbalance) · blank = no arithmetic check applies |

Arithmetic checks by statement type:
- **income statement** — walks the lines to "profit for the year", skipping
  sub-totals, and checks the leaves sum to it.
- **balance sheet** — total assets = total equity + total liabilities (or a
  net-assets / total-equity variant).
- **cash flow** — operating + investing + financing = net change · **or**
  opening + net change (+ FX retranslation) = closing · **or**, for statements
  that print no investing sub-total (du's format), operating sub-total + the
  sum of the investing lines + financing sub-total = net change. Anchors are
  matched on the distinctive *tail* of each line, so a clipped "Net cash"
  prefix still resolves. Blank only when too few of those lines survived.
- **note** — items sum to the printed total.
- **changes in equity** — no generic verdict (it's a movement matrix); shown blank.

Years for a table with no header row fall back to the file's own fiscal year
(from the file name, e.g. `en-2020-…` → 2020/2019), then to the dates the page
talks about; if prose bleed left only the comparative year, it is promoted back
to the file's fiscal pair. Restatement / IFRS-transition grids
(previously-reported + adjustment + restated) are recognised by their column
arithmetic and listed as plain `table`. MD&A "highlights" blocks (EBITDA /
margin / "…Summary" rows) and paragraph-heavy note pages are likewise kept out
of the statement types. A section header welded onto a "Total … assets" line by
the reconstruction is trimmed back.

Noise (tiny tables, prose fragments) is hidden — add `--all` to see it.

## 2 · Export the ones you want

```
python extract_all_tables.py report.pdf --only 21,22,24 --out statements.xlsx
```
`--only` takes numbers and ranges: `21,22,24` or `5-8`. Omit it to export
everything on the list.

Each table becomes a worksheet: title, source line, the data, **total rows in
bold**, and — where there are two year columns — extra **Δ (change)** and **Δ %**
columns. Plus a **Contents** sheet indexing every table with its `what` / years /
`foots`.

## 3 · Compare two reports

```
python extract_all_tables.py  new_report.pdf  old_report.pdf  --compare
```
Matches the same logical tables across the two files and writes a **diff
workbook** — one sheet per matched table:

```
Line item                         A 2024      A 2023   B 2023      B 2022   Δ A(cur−prev)   Δ% A   restated?   status
Revenue                       14,635,917  13,636,340  13,636,340  12,754,492      999,577    7.3               changed
Federal royalty on profit     (1,675,882)          –           –           –                              NEW in A
Share of loss of associate             –           –     (2,720)     (7,913)                          removed (only in B)
...
```

- **status** per line: `changed` / `NEW in A` / `removed (only in B)`
- **restated?** flags a line where report A's prior-year figure ≠ report B's
  current-year figure for the *same year* — i.e. the comparative was restated
  (or one of the two was mis-extracted). The Summary sheet says
  `prior-year columns agree` or `N RESTATED`.
- Renamed line items now mostly reconcile automatically (an alias table plus a
  fuzzy pass pairs "Federal royalty" with "Federal royalty on regulated
  profit"); a genuine rename still shows as one `removed` + one `NEW`.

(`--compare` scans whole documents; if you use `--pages`, make the range wide
enough to cover the statements in **both** files.)

---

## All options

| flag | |
|---|---|
| `--serve` | open the local preview + multi-select UI in a browser |
| `--list` | inventory only, write nothing |
| `--only 1,3,5-8` | export just those inventory numbers (one workbook) |
| `--audit` | print the worked arithmetic behind every foots / NO FOOT verdict, plus every label-health warning |
| `--json FILE` | also write the full inventory (kinds, years, verdicts, health, rows) as JSON |
| `--compare` | two PDFs → row-level diff workbook |
| `--all` | keep tiny / mostly-text fragments in the list |
| `--pages 45-60` | only scan these physical pages |
| `--out FILE.xlsx` | output path |
| `--min-rows N` / `--min-cols N` | ignore anything smaller (default 2 / 2) |
| several PDFs / a folder | scanned into one workbook |

### Trust signals

- **foots** — the arithmetic identity holds, checked **per value column** (a
  broken prior-year column is not hidden by a good current-year one). `--audit`
  and the UI print the sum computed vs the printed total for each year; on a
  doubled / stray row it names that row.
- **cross-year consistency** — when several reports are scanned together
  (`--audit a.pdf b.pdf`, or the UI with >1 file loaded), report A's prior-year
  column is checked against report B's current-year column for the same
  statement. A few differing lines is called a **restatement**; many is called
  **columns likely misaligned**. Previously only `--compare` did this.
- **label health / figure health** — two 0–100% scores. Labels: prose bled in
  from an adjacent column, a welded section header, a clipped first word, an
  empty label. Figures: a magnitude outlier for its column (thresholds derived
  from that column's own distribution — `mean + N·sd` of log-magnitudes), or a
  row whose two years differ ~100× (mis-read). Flagged rows are listed by
  `--audit` and tinted amber (label) / red (figure) in the UI.
- **year-column order** — a statement whose page header lists years oldest-first
  is flagged (figures may be attributed to the wrong year — use *Swap year
  columns*).
- **segmental / multi-entity** — a statement whose columns are entity names, not
  years, is flagged as not a comparable time series.
- **incomplete statement** — a balance sheet that footed on the assets side only
  is flagged "equity / liabilities side incomplete", not passed silently.
- **scanned PDF** — near-empty text layer → a loud warning, not an empty result.

### The number parser

`parse_number()` is a standalone, unit-tested function. It accepts `1,234` ·
`1,234.56` · `(1,234)` · `-1,234` · `1,234-` (trailing minus) · `1 234` (space
thousands) · `1.234,56` (European) · `12.5%` · `$1,234` / `AED 1,234` ·
`1,234 CR/DR` · trailing footnote marks (`1,234*`, `1,234¹`, `1,234 (a)`). A
bare dash / `nil` / `n/a` is a blank, not a zero.

### Layout & config

`tablekit/config.py` holds every tuned threshold (`CONFIG`); `tablekit/parse.py`
holds `parse_number`. `import extract_all_tables` still exposes all of them.
Override thresholds from `extract_all_tables.toml` (or `$EXTRACT_TABLES_CONFIG`)
— see `extract_all_tables.toml.example`. `pyproject.toml` installs the
`extract-tables` / `tablekit-serve` entry points. `docs/PIPELINE.md` is the
maintainer's map. `-v` on the CLI turns on the diagnostics logger.

### Tests — `pytest tablekit_tests/ -q` (97)

| suite | needs PDFs? | what |
|---|---|---|
| `test_units.py` | no | analyze / label+figure health / cross-year / stitch / cash-flow, on hand-built tables |
| `test_serve.py` | no* | inventory shape, inline edits, one-workbook export (*one HTTP end-to-end test uses a PDF if present) |
| `test_golden.py` — `test_parse_number` | no | 34 printed-figure forms |
| `test_golden.py` — `anchors.json` | yes | figures read **by hand from raw PDF text** — the independent check the snapshot can't give |
| `test_golden.py` — `golden.json` | yes | snapshot of detector output on 14 reports (42 statements); fails on any kind / years / foots / shape / value change |

CI (`.github/workflows/tests.yml`, Python 3.11 + 3.12) runs the no-PDF suites on
every push. Regenerate the snapshot **only** after a deliberate, verified
change: `python tablekit_tests/snapshot.py`.

## What's solid / what's rough

**Solid:** ruled tables and normal-layout statements — values correct. The
`foots` check is real for income statement, balance sheet, **cash flow**, and
footing notes. On the full test set — **du 2010–2025** and **Etisalat / e&
2018–2025**, plus DFM 2021 — every income statement, balance sheet and cash
flow reads `foots ✓`, with two exceptions (du 2012 & 2013 cash flow, an old
2-up layout, show no verdict; their P&L and balance sheet still foot). The
`--compare` prior-year consistency check is the strongest signal in the tool.

**Rough:**
- **Titles** are guessed from nearby headings; on dense 2-up pages they're
  sometimes the wrong half or a running page header. Rename the sheet after.
- The oldest 2-up pages (du 2010–2014) reconstruct the statement *alongside*
  the interleaved auditor's-report column, so on those pages some rows carry
  prose glued to the label — worst on du 2010's balance-sheet page (~10 lines).
  The **figures are intact and reconciled** (that's what `foots ✓` attests);
  the labels are noisy. du 2010's income statement + cash flow (a clean 2-up
  spread) come out fine. For du / Etisalat / DFM's income statement + expense
  note done *properly* (reconciled, audited), that's still `telecom_extract.py`.
- On reconstructed statements `--compare` still inflates the changed/new/removed
  counts from label-alignment noise; trust the `prior-year columns agree` line
  and eyeball the rows.
- Notes-heavy pages over-detect paragraphs as tables (mostly filtered).
- Cash-flow verdict is left **blank** (not `NO FOOT`) when the operating /
  investing / financing sub-total labels didn't survive extraction.
- Page-break stitching handles a balance sheet / P&L that runs onto the next
  page **only** when the continuation is an un-headed run of the right line
  items; a continuation buried in half a page of auditor's-report prose (du
  2025's liabilities page) is not merged — the verdict stays honest about the
  half it could check.
- **Changes-in-equity** now gets an arithmetic check (closing balance = opening
  + Σ movements, on the widest column). It reports a verdict only when the
  opening/closing balance rows are cleanly detected; otherwise blank.

**Not supported:** scanned PDFs (detected and flagged, not processed).
