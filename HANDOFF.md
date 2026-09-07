# PDF Financial Table Extraction — Project Handoff

## What this is

A Python script (`extract_two_tables.py`) that extracts two specific tables — the
**Consolidated Statement of Profit or Loss** (truncated at "Net profit for the
year") and the **General and Administrative Expenses** note — from Dubai
Financial Market (DFM) PJSC annual report PDFs, across 16 years (2010–2025),
and writes them into a single Excel workbook (`two_tables_extracted.xlsx`),
one sheet per report.

This is **not a general-purpose PDF table extractor**. It is heavily tuned to
DFM's specific report structure, wording, and line-item vocabulary. Every fix
described below came from finding a real bug in a real DFM report — there is
no theoretical/generic table-detection layer here beyond what DFM's own report
designs required.

**Everything is deterministic.** No LLM, no ML model, no API calls anywhere in
the extraction or verification path. `pdfplumber` for word/table positions,
`pypdf` for PDF normalization and page-label metadata, `openpyxl` for writing
the workbook, `re` for pattern matching, standard library otherwise.

---

## Current status

All 16 reports (2010–2025) successfully extract both tables, with content-match
confidence scores ranging 43%–100% (lower scores on older years reflect those
years' simpler statements genuinely having fewer line items than the newer
years' reference vocabulary — not extraction failure; verified via arithmetic
reconciliation, see below).

**Two files are missing from the batch and were never obtained:** 2023 and 2024
integrated reports were tested separately as individual files, not as part of
the main 16-file batch folder — check whether they need to be re-added to
`allfiles/` if continuing batch work. (Also worth double-checking: at one point
2018/2019/2020 were supplied as text-extract stubs rather than real PDF
binaries and had to be re-requested — if new files come in from the same
source, verify with `open(path,'rb').read(10)` that they start with `%PDF`
before assuming they're usable.)

---

## Architecture / pipeline (in the order the code runs)

1. **`normalize_pdf_for_reading()`** — re-saves every PDF through `pypdf`
   before `pdfplumber` touches it. **This is mandatory, not optional**: at
   least one real PDF (2023's report) caused `pdfplumber.open()` to hang
   *indefinitely* — not slow, literally never returns — even though `pypdf`
   could read it fine. Re-saving through `pypdf` first silently fixes this
   with zero data loss, for reasons never fully diagnosed (likely a
   content-stream quirk `pdfminer`/`pdfplumber` chokes on). Always do this
   first on any new file.

2. **Heading search (Pass 1)** — look for the target's heading text pattern
   (e.g. "Consolidated Statement of Profit or Loss") on each candidate page,
   then look at tables on that page and the next couple of pages.

3. **Full-document content-match scan (Pass 2, fallback)** — if Pass 1 doesn't
   find a confident match (many DFM reports print headings that don't match
   patterns cleanly — e.g. a title split across two PDF text lines, "Consolidated
   Income" / "Statement", doesn't match a single-line regex), scan every table
   on every page and score it against `reference_row_labels` (known DFM line
   items like "Trading commission fees", "iVESTOR expenses") via fuzzy string
   matching. Best-scoring table wins.

4. **Multi-tier table extraction** (`get_real_merged_tables`), tried in order:
   - `pdfplumber`'s native ruled-line table detector.
   - `attach_left_labels()` — for tables where the ruled box only wraps the
     *numbers*, not the row labels (labels sit outside the box entirely).
     Searches left of the box, within each row's own vertical band, for label
     text; also searches right for a missing second value column (some
     reports rule a box around only the current year, leaving the prior-year
     comparison column outside it).
   - `reconstruct_unruled_rows()` — full fallback for reports with almost no
     ruling at all (2022's report: the ruled box only ever wrapped a 1-row
     column-header stub, nothing else). Rebuilds rows entirely from word
     x/y positions.
   - **Gap-filling** — some reports (2012 specifically) rule a box around
     every *other* line, leaving alternating rows completely unruled and
     invisible to every above method. Detects vertical gaps between
     consecutive ruled fragments that still contain unclaimed words, and
     reconstructs just those rows. This fix alone took 2012's P&L match score
     from 44% (6 rows genuinely missing, totals didn't reconcile) to 81%
     (fully reconciling).

5. **Column-header reconstruction** (`reconstruct_column_header`) — many
   reports print "Notes / 2023 / 2022 / AED'000" *entirely above* the ruled
   table box, invisible to normal row extraction. Searches the gap above the
   matched table, using font-size and text-pattern heuristics (a real header
   is short and made of label/year/unit tokens, not prose — critically,
   distinguished from a subtitle like "for the year ended 31 December 2023"
   which also contains a year but is prose and must be rejected).
   **Important design decision, reversed once during this session:** header
   lines that print on separate PDF lines (e.g. "2023 / 2022" on one line,
   "AED'000 / AED'000" on the line below) are kept as **separate rows** in
   the output, matching the source PDF's actual layout — an earlier version
   merged them into one row for cleanliness, but the user explicitly wanted
   fidelity to the original document over compactness. `write_table()`
   accepts a **list of header rows**, not a single row, for this reason.

6. **`normalize_missing_notes_column()`** — rows with a note-reference number
   ("Investment income | 20 | 168,808 | 79,989") and rows without one
   ("Trading commission fees | 226,064 | 200,493") naturally extract with
   different column counts. This detects the pattern from rows that *do* have
   a note ref and inserts the missing blank for rows that don't, so
   everything aligns under the same header. Applied generically after
   extraction, not tied to one code path — it fixed this same misalignment
   across 2022, 2023, and 2024 simultaneously once implemented at the right
   layer.

7. **Verification (manual, not yet automated in the script)** — arithmetic
   reconciliation: `sum(income line items) == Total income`,
   `Total income + Total expenses == Net profit`, `sum(G&A items) == G&A
   total`. This is NOT currently built into the script — it was run as a
   separate one-off verification script during debugging (see "Verification
   script" below) and should be considered the single highest-value thing to
   formalize into the pipeline going forward.

---

## Bug catalog — root causes and fixes (read this before touching the code)

This is the important part. Every one of these was found by a human looking
at real extracted numbers and asking "does this look right" — not by
anticipating the bug in advance. If extending this to new files/companies,
expect to find the *same classes* of bug, even if the specifics differ.

### Cross-column contamination (the most recurring bug class)
Dense, multi-note pages (several notes packed side by side or in tight
columns) repeatedly caused text search logic to grab words from a
*different*, unrelated line item that merely happened to share a similar
vertical (or horizontal) position. Concrete instances fixed:
- A left-label search reaching too far left grabbed an unrelated note's
  numbers ("(Note 16) 8,470 11,094 Payroll and other benefits" — the note-16
  text belonged to a different row entirely).
- A "find the nearest note heading above this table" search picked up a
  *different* note's heading from an adjacent column at a similar height
  ("25. Commitments" instead of "22. General and administrative expenses").
- A header-reconstruction search on a 2-statement-per-page landscape layout
  grabbed the *other* statement's title ("POSITION" — a fragment of
  "...FINANCIAL POSITION" — instead of "PROFIT OR LOSS").
- **The "852" bug** (see dedicated section below) — the most subtle instance:
  a left-label search ran unconditionally on *every* row including header
  rows that didn't need a label search at all, and grabbed a stray number
  from a completely unrelated line ("Unearned revenue... 3,746 / 852").

**General lesson:** any "search nearby for related text" heuristic on this
class of document needs (a) a *tight* column-restricted search tried first,
widening only if nothing found, and (b) a value-plausibility check on
whatever gets found, not just a proximity/gap check. Proximity alone is not
sufficient signal on densely-packed pages.

### The "852" bug (case study — read this one in full, it's instructive)
**Symptom:** 2019's G&A table header showed `[None, 852, '2019', '2018']` —
an extra, unexplained number in the Notes-column position.
**Root cause:** the code searched left of *every* row for a label,
unconditionally, including the header row itself (which already had its own
content — no label search needed at all). Since nothing relevant existed to
find, it grabbed the nearest cluster of text within reach, which belonged to
an unrelated line ("Unearned revenue: 3,746 / 852") sitting nearby.
**Fix required three iterations to get right** (each one caught a regression
the previous one introduced — this is worth internalizing as a pattern, not
just a one-off):
1. First fix: `has_text_label(row)` — skip the search if the row already
   contains any string. **Broke:** rows whose only "string" content was a
   short note-reference number like "15, 16" got wrongly treated as
   "already labeled," losing their real label entirely.
2. Second fix: `has_real_label(row)` — exclude pure-digit strings from
   counting as a real label. **Broke again:** `clean_cell('-')` returns the
   literal string `'-'` (a dash placeholder for a blank/zero value), and the
   digit-only exclusion regex didn't account for hyphens, so a legitimate
   row whose only "text" was a dash placeholder got wrongly treated as
   already-labeled, losing ITS real label too.
3. Third, final fix: also exclude rows whose content already matches
   `looks_like_column_header()` — i.e., a row that's just bare years
   (`['2019', '2018']`, both pure integers, no string at all) still needs to
   be recognized as "doesn't need a label search" even though it has zero
   string content to check in the first place.
**Verification:** after the fix, ran a full cell-by-cell diff between the
pre-fix and post-fix workbook across all 16 sheets to confirm *only* the
intended cells changed (found exactly 3 differing rows total, across 2019 and
2020 — 2020 had the identical bug independently, "3,746 977", not
previously noticed). Also diffed cell *formatting* (font/fill/border), not
just values, since removing a phantom value should also remove its header
styling — confirmed only the 2 cells adjacent to the removed values lost
their (correctly no-longer-applicable) header styling.

### Column alignment / row-shape mismatches
- Rows with vs. without a note-reference number naturally extract with
  different column counts → `normalize_missing_notes_column()`.
- Multi-line headers reconstructed with the wrong padding (a 2-cell line like
  "2020 / 2019" landing under the wrong columns relative to a 3-cell line
  like "Notes / AED'000 / AED'000") → right-align each header line against
  the body's actual column count, with a special case for a standalone
  "Notes"/"Note" line (which should *left*-align to column B, not right-align
  to the value columns).
- A currency-unit-only row ("AED'000") landing in the wrong column entirely
  because its native ruled cell was one giant unruled cell internally →
  detect "bare unit row" and distribute one unit per column that already
  holds a bare year, rather than trusting its raw column index.
- Two "AED'000" instances only ~11pt apart on the page merging into one
  garbled cell due to the general word-clustering gap threshold → detect and
  split the specific "UNIT UNIT" duplicate-text pattern back apart.
- A year+unit combo glued onto one line internally ("2018 AED'000") when the
  source PDF actually has them on two separate lines → detect the glued
  pattern via regex and split it back into two rows, matching the source.

### Missing/incomplete rows
- 2012: alternating unruled rows (see Gap-filling above) — 6 rows were
  silently absent; arithmetic didn't reconcile until fixed.
- Trailing (second) value column sitting outside a table's own ruled bbox by
  anywhere from ~10pt to ~300pt depending on the report — required a
  progressively-widened search with a "does this look like a value"
  plausibility gate (numeric, or a dash placeholder, or a
  year+currency-unit combination) to avoid pulling in unrelated prose from
  the same widened search.

### Heading/title identification
- `find_actual_heading()` reports what the source document itself calls a
  matched table (e.g. surfacing "called 'Consolidated Income Statement' in
  the report" when the target's canonical name is "Statement of Profit or
  Loss") — useful for traceability when a match came from content-scoring
  rather than an exact heading-text match.
- Required font-size-based logic (a real title uses a distinctly larger font
  than surrounding body text) plus tie-breaking on horizontal alignment with
  the actual matched table (not just "nearest by vertical position") to avoid
  grabbing an unrelated, similarly-styled title from an adjacent column on a
  landscape 2-statement-per-page layout.

### Page numbering
- Several reports restart page numbering partway through (front matter
  numbered separately from the report body), so the physical PDF page index
  and the page number printed in the document diverge. Fixed by reading the
  PDF's actual embedded `/PageLabels` metadata via `pypdf`
  (`load_page_labels`, `display_page`) rather than guessing from visible
  text. Displayed as e.g. `"page 54 (PDF page 61)"` only when they actually
  differ, to avoid cluttering output when they match.

---

## Verification script (not yet in the main pipeline — should be)

A standalone script was used throughout debugging to catch arithmetic
inconsistencies automatically instead of relying on manual eyeballing. This
caught the 2012 missing-rows bug and confirmed every other year reconciles.
**Strongly recommend formalizing this as an automated check in the pipeline
itself**, not a side script — this is the single most valuable addition
available and was explicitly identified as such during design discussion for
a potential app version of this tool.

Core logic (reconstruct against the actual script's output structure before
reusing — this was a quick verification tool, not production code):
```python
# For each P&L sheet: parse into income-section and expense-section rows,
# sum each section, and confirm:
#   sum(income items) == "Total income" row
#   sum(expense items) == "Total expenses" row (or "Operating expenses")
#   Total income + Total expenses == "Net profit for the year" row
#     (± a "Profit before tax" / corporate tax deduction step for 2024/2025,
#      which introduced a corporate tax line the simple version doesn't
#      handle — these show as false-positive mismatches in the crude
#      version; manually verified correct via full arithmetic trace)
# For each G&A sheet: sum(line items) == total row
```

---

## Known limitations / things NOT to assume are solved

1. **Company-specific, not general-purpose.** `reference_row_labels` in the
   `TARGETS` config are DFM's actual line-item wording. A different company
   (even another UAE exchange like ADX) will need its own vocabulary rebuilt
   — the heading-search-with-content-fallback *architecture* transfers, the
   *vocabulary* does not.
2. **Tuned thresholds throughout, not derived.** Gap-cluster distances (15pt,
   80pt, 95pt, 250pt in various places), font-size multipliers, "looks like a
   note reference" digit-count patterns (1-2 digits, not 3) — all were
   arrived at by trial and error against these 16 specific files. Expect to
   need to retune when hitting a genuinely new layout.
3. **2025's P&L header remains slightly malformed** (a known, accepted
   exception) — its source PDF has a genuinely unusual layout where one
   column's year and the other column's year+unit print on different lines
   in a way that doesn't cleanly map to the general two-line-header-split
   logic. Left as-is rather than risk a fragile one-off fix; flagged here so
   it isn't mistaken for an unnoticed bug.
4. **No automated regression suite exists yet.** Every fix in this session
   was verified by manually re-running the full 16-file batch and comparing
   printed scores/row-counts, plus occasional full cell-diffs. Multiple times
   during this session, fixing one file's bug silently broke a different
   file that had been working — this pattern will keep recurring without an
   automated test harness. **Building one (a directory of PDF + verified
   expected-output pairs, replayed on every change) was identified as the
   top-priority next engineering investment**, separate from and prior to
   any new company/format support.

---

## Files in this handoff

- `extract_two_tables.py` — the current, working script (all fixes described
  above are included).
- `two_tables_extracted.xlsx` — the current output, all 16 files, verified.
- This file.

## Suggested next steps, in priority order

1. Build the automated regression test harness (PDF + verified-output pairs,
   replayed on every change) — highest leverage, prevents re-litigating
   already-fixed bugs.
2. Formalize the arithmetic reconciliation check into the pipeline itself as
   an automatic pass/fail, not a manual side-script.
3. Confirm 2023/2024 are properly included in the main batch folder if
   continuing multi-file work.
4. If extending to a new company/entity: budget for a full debugging pass
   comparable in scope to this session, not a quick tweak — expect the same
   *classes* of bug (cross-column contamination, glued header lines, missing
   unruled rows) to reappear in that company's own specific way.
