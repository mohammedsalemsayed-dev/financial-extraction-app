# Changelog

## 0.7.9

Two threads of work: another round of the same page-by-page real-file
audit from 0.7.8 (more corpus coverage, several more root causes found and
fixed the same way -- traced to a general cause, not patched per file),
and a new feature: an interactive table-grid preview for manual box-select,
built after the user asked several pointed questions about exactly what a
drawn box actually captures and whether partial selections (one row, one
column, one cell) would work correctly.

### More data-fidelity fixes (same audit discipline as 0.7.8)

- **Fixed: two unrelated statements fused into one table.** img2table's own
  borderless-table clustering can glue a completely different statement
  onto the one a user actually marked (confirmed live: `du annual
  2011.pdf`'s income statement fused row-by-row with its cash flow
  statement on the same landscape page). Detected via statement-specific
  vocabulary and refused rather than served with one statement's columns
  attached to the other's rows.
- **Fixed: a bare year-header row silently lost.** A row shaped like
  `[None, None, "2012", "2011"]` (no real label, only the year columns) was
  being treated as redundant and dropped even when it carried the ONLY
  copy of those years on the page (`du annual 2012.pdf`'s income
  statement) -- now only dropped when it's a genuine duplicate of another
  row's years elsewhere in the table.
- **Fixed: `_looks_garbled`'s fixed "&ge;2 numbers" threshold** flagged a
  legitimate wrapped row (a label that wraps around its own note-ref and
  figures) as corruption once the same check was extended to the
  img2table path. Now relative to the table's own column count instead of
  a fixed constant.
- **Fixed: a real label column dropped as prose**, twice over -- once
  because a still-unparsed numeric string ("12,951,414") passed as "real
  text", and again because any alphabetic content anywhere in a column
  (not a genuine fraction of its own cells) counted as a label. Both
  tightened; the safeguard now requires alphabetic content on a real
  portion of the column's own non-empty cells.
- **Fixed: reversed-parenthesis negatives** (`)1,234(`, a bidi artifact)
  on the manual box-select path, plus the rarer case where the digits
  also split across two reversed-order lines -- matching handling that
  already existed for automatic detection.
- **Fixed: a glued-cell "prefer the whole parse" repair** was reuniting a
  nil marker and a separate real figure (`"- 120,172"`) into one
  fabricated negative number, losing the nil marker -- excluded whenever
  either half of a glued pair is a bare dash.
- Several smaller fixes in the same vein (Notes-column stripping missed
  by a units-row diluting the ratio, a data race in a shared note-column
  global under concurrent requests, Microsoft's 10-K getting demoted from
  "income statement" for using US-GAAP wording instead of IFRS terms, and
  more) -- see the new regression tests in `tablekit_tests/` for the full,
  precise list; each one is named after the real bug it locks in.

### New: interactive grid preview for manual box-select

Built in six phases, each verified against real files before moving to
the next -- full detail in each phase's own code comments and tests:

- After releasing the mouse on a drawn box, the actual detected table
  grid (rows and columns) is now drawn over it, so what you see is
  what extraction will use -- not a guess about what's inside the box.
- Any grid line can be dragged to a new position; adjacent cells stay
  contiguous when you move a shared edge.
- "+ Row line" / "+ Column line" arm a one-click insert; dragging a line
  off the table's own edge removes it.
- The finalized grid -- auto-detected or hand-edited -- now actually
  drives extraction (`extract_all_tables._extract_via_grid`), reading
  text directly from each cell's rectangle instead of re-guessing
  structure. Omitting it reproduces the pre-existing behavior exactly
  (confirmed via a byte-identical `golden.json`).
- A grid-derived Notes-reference column (e.g. "6", "7") arrives already
  isolated in its own cell, never glued to anything -- the existing
  note-stripping logic only knows how to un-glue text, so it never got
  the chance to remove it. Fixed with a small, targeted re-glue step
  (`_reglue_bare_note_column`) so it's handled the same way either path
  produces it.

New tests throughout (`tablekit_tests/test_manual_mode.py`,
`tablekit_tests/js/test_webui_logic.js`); full pytest + JS suite green,
ruff/mypy clean, `golden.json` byte-identical.

## 0.7.8

Found live by the user again, this time on `du annual 2011.pdf`'s balance
sheet: cells bleeding into each other and duplicating. Traced to a real,
general root cause, then -- per an explicit instruction not to ship
per-file patches -- built a comprehensive audit that runs the real detector
over every page of every loaded real file (24 files, 2428 pages, ~3200
tables), simulates a user drawing a box around each detected table, and
checks the result for the same class of corruption. That audit is what
actually drove this release: every fix below was found either by it or by
following up on what it surfaced, and every fix was re-verified by re-running
it, not just by the fixture that originally caught it.

- **Fixed: several real rows silently merged into one garbled cell.** The
  pdfplumber fallback path in `extract_region` (used when img2table has no
  region to anchor a manually-drawn box) can merge multiple lines into one
  cell with an embedded newline. A newline in one cell is completely normal
  on its own (a wrapped label) -- the tell is specifically when two or more
  of the newline-split pieces independently look like a number, which a
  label's continuation lines never do. New `_looks_garbled` check
  (`extract_all_tables.py`) refuses the fallback result outright in that
  case rather than show it, falling through to the existing "no table
  found" path. New tests; full golden snapshot and pytest suite clean.
- **Fixed: numbers silently glued together with no real separator between
  them** (`tablekit/parse.py`, `parse_number`) -- e.g. `2020 2020 2019 2019`
  (four years) coming out as `2020202020192019`, or two 9-digit figures
  coming out as one 18-digit one. `normspace` turns an embedded newline
  into a space, and `parse_number` used to blindly strip every space to
  support legitimate space-grouped numbers like `1 234 567` -- with no way
  to tell "one grouped number" from "several numbers that ended up in the
  same cell" apart. Fixed by checking the space-separated tokens actually
  fit a real grouped number's shape (first token &le;3 digits, every later
  token exactly 3) before collapsing them; otherwise refuse rather than
  fabricate a number nobody printed. Also tightened the absurd-digit-run
  backstop that catches the same corruption when no separator survives at
  all (18 &rarr; 16 digits -- the widest safe value that doesn't break the
  existing &plusmn;10<sup>15</sup> round-trip property test; the largest
  real figure across every hand-verified fixture is 9 digits). Confirmed
  against the full audit: 94 absurd-number findings &rarr; 0, across every
  file that had any.
- **Fixed: correct equity statements flagged as not reconciling**
  (`extract_all_tables.py`, `_equity_foots`). Found by hand-verifying a
  `du annual 2018.pdf` equity statement was already numerically perfect --
  every row's own columns summed to its row total, and the full
  opening-to-closing roll-forward balanced exactly -- yet the tool called it
  a footing failure. `_equity_foots` sums the movement rows between an
  opening and closing balance but wasn't excluding "Total comprehensive
  income" / "Total transactions with shareholders..." subtotal rows, so it
  double-counted them against the rows they're already subtotals of.
  Deliberately narrower than reusing the existing `_HARD_TOTAL_RE`: that
  pattern also matches "Profit for the year", which is a real movement in
  an equity roll-forward, not a rollup, unlike in an income statement.
- **Fixed: negative numbers with their parentheses reversed** -- `)1,234(`
  instead of `(1,234)`, a bidi text-ordering artifact in the source PDF,
  confirmed on both an e&/Etisalat file and `du annual 2019.pdf`.
  `telecom_extract.py` (used by the auto-detect path) already special-cased
  this with a clear comment explaining it; it had just never been ported to
  `tablekit/parse.py`, which the manual box-select path actually relies on
  -- auto-detection was quietly getting these right while a user's own
  drawn box got them wrong. Ported the same handling over. Two related
  patterns found via the same audit and fixed alongside it:
  - **Multiple reversed numbers glued into one cell** -- `)87,579( )11,915(`
    -- extended `_split_glued_cell`'s existing token pattern
    (`tablekit/img2table_backend.py`) to also recognize the reversed form,
    so it un-glues both orientations the same way it already did for
    normal-order glued cells.
  - **A reversed number's digits split across two physical lines, with the
    line order ALSO reversed** -- `)69,040\n1,1(` for what was printed as
    `(1,169,040)`. Confirmed by reconciling the surrounding row's own
    arithmetic (not just by inspection) on both real occurrences found.
    Deliberately narrow -- keyed on the raw newline, which is still visible
    before `normspace` collapses it to a space and destroys the signal a
    more general "try reversing any two tokens" rule would need, and which
    would risk mis-firing on an ordinary two-number cell that has nothing
    to do with this artifact.
- **Fixed: two unrelated tables on the same landscape page merged into one
  nonsensical table.** On `en-2021-etisalat-group-annual-report.pdf` p62, a
  balance sheet and a completely different statement of changes in equity
  sit side by side; the merged result attached "Share capital" / "Reserves"
  columns from the equity statement onto balance-sheet rows like "Goodwill
  and other intangible assets". Traced to img2table's OWN internal
  borderless-table clustering fusing the two before any of this project's
  code runs, not to an explicit merge this project performs -- so hardening
  this project's own `_merge_side_by_side` (skip merging two regions that
  each already have their own label column, since that's the actual
  signature of two independently-complete tables rather than one table's
  labels half and figures half) is real and kept, but doesn't reach this
  specific bug on its own. Two structural detection approaches were tried
  and rejected before finding one that worked: a general "does this region
  have two independent label columns" check produced 384 hits dominated by
  ordinary text-heavy pages (a first sample was a table of contents), and
  is not shipped. What worked instead was content, not structure: a
  deliberately narrow check for primary-statement vocabulary from two
  DIFFERENT statement kinds in the same region (`balance at \d` /
  `transactions with owners` for a changes-in-equity statement alongside
  `total assets` / `total liabilities` for a balance sheet) -- a first pass
  at even this found 2 false positives (accounting-policy notes that
  mention "transactions with the owners" in an ordinary sentence), fixed by
  skipping any cell that reads as prose rather than a short label. With
  that filter, a full 24-file, 2428-page sweep found exactly one match: the
  real bug. `img2table_page_tables` now refuses (rather than serves) a
  region matching this signature, the same "detect and refuse" choice
  already made for a merged-rows cell -- there's no safe way to split it
  back into its two source tables here, so a box drawn over the affected
  area now correctly reports no table found (or, for a box that only
  partly overlaps, the correct figures with their labels dropped) instead
  of showing one statement's columns silently attached to the other's rows.
  New regression tests for both the structural guard and the vocabulary
  check, including the two false positives the prose filter exists for.

## 0.7.7

Found live by the user, using the app directly (not through a test): a real
note-table (du annual 2010.pdf, p22, a PP&E schedule) with a column header
split across two rows and a row label with its words scrambled. Both traced
to a real root cause; only one was safe to ship.

- **Fixed: a row label wrapped across two physical lines had its words
  interleaved by X-position instead of read in order** -- `'Net At 31 book
  December value 2009'` instead of `'Net book value At 31 December 2009'`.
  `_attach_left_labels` (`extract_all_tables.py`) recovers a row's label
  text by searching a Y-band derived from img2table's own cell bbox, which
  can be taller than one text line (a padded/bordered row, or here a
  genuinely 2-line label) -- every word inside that band was being sorted
  by x0 alone, across whichever lines happened to fall in it, instead of
  read one line at a time. Fixed by grouping words by their own line first
  (same `round(word["top"])` convention `guess_title()` already uses), THEN
  sorting left-to-right within each line. New regression test
  (`test_attach_left_labels_keeps_two_wrapped_lines_in_reading_order`).
  Verified against the full golden snapshot (60 cases, 14 real annual
  reports): no change to any of them -- this only affects the specific
  multi-line-band situation, never a normal single-line one.
- **Found, attempted, reverted: a 2-line-wrapped column header (`'Capital
  work'` / `'in progress'`) coming back as two separate rows instead of
  one.** Traced to img2table's own row-clustering treating each Y-band as
  its own row, with no concept of "this column's header wrapped." Wrote a
  geometric merge heuristic (tight gap + one row's filled columns a proper
  subset of the next's) that fixed this specific table cleanly -- and then
  ran it against the full golden snapshot before considering it done, which
  is what caught the problem: on `du annual 2013.pdf`'s cash flow statement
  and `du annual 2025.pdf`'s equity statement, the same heuristic merged a
  section header into its own first line item, and merged two DIFFERENT
  line items' figures into one garbled string cell, because those
  documents' normal section-header-to-first-line spacing (4.7pt) is
  geometrically indistinguishable from a genuine wrapped-line gap (3pt on
  the page that motivated the fix) -- there's no safe, general geometric
  rule that separates them, at least not one this pass found. Reverted
  rather than ship it: corrupting real figures elsewhere to fix a display
  quirk in one table is the wrong trade, and "found a real bug, confirmed a
  first attempt wasn't safe, reverted rather than guess again" is the
  intended outcome of testing a core-extraction change this way, not a
  failure of it. The header-split issue itself is unfixed as of this
  release.

## 0.7.6

- **`serve.py --debug` was practically unusable for any real session: 4.3 GB / 44.5 million lines from a 20-minute, 26-file stress test.** Chased down by actually running that stress test (aimed at reproducing the never-root-caused "breaks after 2-3 runs" report from 0.7.1) with `--debug` on and watching `debug.log` grow far faster than the request count could explain. `logging.basicConfig(level=DEBUG)` cascades to every logger in the process that doesn't set its own level, not just this project's own five `LOG.debug()` calls -- including `pdfminer` (underneath `pdfplumber`), which logs every single parse token, seek and keyword at DEBUG. Sampled the actual log content at several points through the file rather than assuming, confirmed >95% of lines were `pdfminer.psparser`/`pdfminer.pdfinterp`/etc., not this project's. Fixed by capping `pdfminer`'s own logger to WARNING whenever `--debug` is on -- verified with a real before/after: the same operation (three full-document scans) that previously wrote gigabytes now writes 1,085 bytes, and what's left is exactly the useful stuff (the per-request state line, the HTTP access log, startup messages).
- **The stress test that found it reproduced nothing else**: 20 cycles across 12 real files (960 API calls, ~19 minutes of continuous rapid file-switching, extraction, and page rendering) came back with zero errors and no unbounded memory growth. Doesn't retroactively explain the original 0.7.1 report -- this was a different session, a different kind of "repeated use" (server-side switching, not the browser session the original report described) -- but it's the first real stress test run against this debug-logging path since it was built, and it came back clean apart from the log-volume issue above.

## 0.7.5

Found by actually loading many real files at once (16 years of one company's
annual reports plus several of another's) instead of the small, clean sets
every automated test uses -- three real bugs, none of them hypothetical:

- **A file could appear twice in the dropdown under the identical name, with
  the second entry silently dead.** `_discover_files` (the startup file
  list: CLI-arg files plus whatever's in `uploads/`) deduped by resolved
  path, not by name -- two genuinely different files at different paths
  (a CLI-arg-loaded report, and a same-named copy sitting in `uploads/`
  from an earlier session's upload) both got listed. `_path()` resolves a
  filename to the FIRST match, so the second entry looked selectable but
  always silently showed the first file's content -- confusing, not
  crashing, which is exactly the kind of bug that survives every "does it
  crash" check. Fixed by deduping on name; CLI-arg files still win, matching
  what the function's own docstring already claimed but the code didn't
  actually do.
- **The same collision, reachable a second way.** `upload_pdf`'s
  rename-on-conflict (`report.pdf` -> `report_1.pdf`) only checked
  `UPLOAD_DIR` on disk, not the full loaded file list -- uploading a file
  whose name matched a CLI-arg-loaded file (living elsewhere, so never
  physically present in `UPLOAD_DIR`) produced the identical dead-entry
  problem live, mid-session. Fixed the same way: the collision check now
  also considers every already-loaded file's name, not just what's
  physically in the uploads folder.
- **The page-number box silently ignored Enter.** Typing a page number and
  pressing Enter -- the obvious way to do it, and how the adjacent search
  box already works -- did nothing; only clicking away (blur) navigated.
  `onchange` fires on blur, not on Enter in a bare `<input>`. Added the same
  `keydown`-checks-for-Enter handler `#searchQ` already had.

Both server-side fixes got regression tests (`test_serve.py`); the
page-number fix was verified live by dispatching a real `KeyboardEvent`
against the running page and confirming the page-image request fired with
the typed page number, since `test_webui_logic.js` deliberately doesn't
simulate DOM interaction (see its own header comment for why).

## 0.7.4

- **`webui.html` split into `webui.html` (67-line shell) + `webui.css` (347 lines) + `webui.js` (1519 lines)**, served as three separate files (`serve.py` gained two routes, `/webui.css` and `/webui.js`, alongside the existing `/`). Reconsidered rather than assumed: this project has no build step and no bundler, and keeping everything in one file to avoid needing either is itself a legitimate, common, professional choice for a no-build local tool -- splitting isn't automatically "more correct." What tipped it here is concrete, already-paid cost, not a style preference: getting ESLint and `tsc` to check the inline script at all (0.7.1, 0.7.2) needed a dedicated extraction step (`extract_inline_script.mjs`, regexing the `<script>` block out to a temp file) precisely because the JS wasn't a real file. Splitting removes that workaround rather than adding one -- `extract_inline_script.mjs` is deleted, both tools now run directly against `webui.js`. It also wasn't a "the app must still work if you just open the HTML file" concession: every feature already round-trips through `fetch()` to `serve.py`'s `/api/*` endpoints, so the app never worked without the server running regardless of where the JS lived.
  Extraction and verification were both mechanical and checked, not assumed: the CSS/JS blocks were sliced out by exact line range and diffed byte-for-byte against the original file's content before `webui.html` was touched (both matched exactly), rather than hand-copied. `tablekit_tests/js/test_webui_logic.js` (which used to regex its own inline `<script>` out of `webui.html`) now reads `webui.js` directly -- same 13 tests, all still passing, unmodified test logic. ESLint and `tsc --checkJs` were re-run against the real `webui.js`: 0 lint errors, and the same 16 already-verified-not-bugs `tsc` findings as before the split (confirming the move didn't change a single line of actual code). Re-ran the full pytest suite and did a live server smoke test after -- page load, `/webui.css`/`/webui.js` both 200 with the right `Content-Type`, extraction, export, and a fresh axe-core accessibility pass (0 violations, matching the pre-split state from earlier this version) -- not just "the diff looks right."
  Updated to match: `CONTRIBUTING.md`, `docs/PIPELINE.md` (a few now-stale "webui.html's X" / "single file" references), `eslint.config.mjs`'s header comment, `.gitignore` (dropped the now-unused `tablekit_tests/js/_webui_inline.js` generated-file entry), and `.github/workflows/tests.yml` (`frontend-logic` job's extraction step removed; `on.push.paths` and the `changelog-nudge` job's file list both gained `webui.css`/`webui.js`).

## 0.7.3

- **`serve.py`'s `do_GET`/`do_POST` were a single long if/elif chain each** (11 and 8 routes) sharing one try/except. Split into one small method per route (`_g_*`/`_p_*`) plus a class-level `GET_ROUTES`/`POST_ROUTES` dict mapping path to method; `do_GET`/`do_POST` now just look up the path and call the mapped handler, with the shared cross-cutting bits (the `_debug_state_line` call, the origin check, the `_NEEDS_FILE` "no file selected" pre-check, the try/except) staying exactly where they were. Behavior-preserving, not a rewrite: every status code, error message and response shape stayed identical -- verified against the full pytest suite (including both real-HTTP-handler tests) before and after, and against a live server smoke test exercising GET (`/`, `/api/files`, `/api/scan`, `/api/table`, `/api/page_raw`, `/api/quickfind`) and POST (`/api/export`) routes with the network tab open, not just assumed from the diff.
- **Property-based tests for `parse_number`/`coerce_cell`** (`tablekit_tests/test_parse_number_properties.py`, using `hypothesis`): the existing example-based cases in `test_golden.py` lock in specific known formats; this adds properties that hold across inputs no one would think to hand-write -- most importantly, that neither function ever raises on arbitrary text (500 fuzzed examples each), which matters because both run on OCR/PDF-extracted text this tool never controls, where a crash on one bad cell would currently take down the whole extraction rather than just that cell. Also checks: formatting an integer with thousands separators and reparsing recovers it exactly; parens negate a plain positive figure; a `%` suffix doesn't change the underlying numeric value; `CR`/`DR` are exact opposites; and `coerce_cell`'s fast path agrees with `parse_number`'s full parse on everything the fast path actually handles (the two are separately maintained, so this is the check that an edit to one without the other can't silently drift). All 7 passed on the first run against the current implementation -- no bugs found, which is itself the useful result: existing hand-picked examples don't reflect a coincidentally-narrow test, the properties actually hold.

## 0.7.2

- **`serve.py --debug`**: a diagnostic capture path the CLI (`extract_all_tables.py -v`) already had but the server never did. DEBUG-level logging to the console and a `debug.log` file, plus a one-line server-state snapshot (`files=`/`scans=`/`pngs=`/`manual_tables=` counts) logged at the top of every request. Doesn't retroactively explain the unreproduced "breaks after 2-3 runs" report from 0.7.1 -- that stopped happening on its own, root cause still unknown -- but a next occurrence now has a timeline of state growth to look at instead of nothing.
- **`serve.py`'s CLI silently dropped every flag it didn't recognize**, including `--port` and `--no-browser` -- both documented, neither wired to anything. Replaced the hand-rolled flag-stripping with real `argparse`, which fixed this as a side effect of adding `--debug`.
- **Caught by the fix above, not by design**: the new per-request state line read `_state["manual"]` directly. Several tests intentionally replace `_state` with a partial dict to stay hermetic (`serve.py` already has a documented convention for this -- `_manual_list`/`_deleted_list` both `.setdefault()` rather than index directly, specifically because "a few tests replace `_state` wholesale with a dict that predates this feature"). The new line didn't follow that convention, so it crashed `do_GET`/`do_POST` -- and therefore any test driving the real HTTP handler -- with `KeyError: 'manual'`. A first verification pass only read the first line of the combined lint+test output and reported success; the failure was in the untruncated tail. Fixed by matching the existing `.get()`-with-default pattern instead of hand-editing the test fixtures to paper over it.
- **mypy**, gradual/permissive (`ignore_missing_imports`, no `disallow_untyped_defs` -- this codebase has zero type annotations by design, and demanding full coverage would produce thousands of findings on code nobody's touching). Run as-is against the real codebase first: 15 findings, all either a genuinely missing variable annotation (3, now added) or an intentional pattern mypy can't see is deliberate (an optional-dependency None-sentinel, four fallback-stub functions with a loose `*a, **k` signature) -- given a scoped `# type: ignore[<code>]` with a comment explaining why, never a blanket ignore. Wired into the `lint` CI job; verified it also passes with zero runtime dependencies installed, so the job doesn't need the full `requirements.txt` install just to lint.
- **`renderPreview(n, d)` was called with a silently-discarded 3rd argument at 6 of its 7 call sites** (`renderPreview(n, d, true)` / `(..., false)`) -- found by running `tsc --checkJs` (permissive, no `--strict`) against the extracted inline script for the first time, which flagged "expected 0-2 arguments" at every one. `git log -S` confirms the function has taken exactly 2 parameters since the commit that introduced it (827dfb3) -- this was never a real 3rd parameter that got refactored away, callers were just always passing an extra argument JavaScript quietly drops. Removed the dead argument from all 6 sites; behavior is unchanged (it was never read), but a future reader can no longer mistake it for live control flow. `tsc --checkJs` also surfaced 16 more findings, all verified NOT bugs (an intentional `err.body = j` enrichment on a plain `Error`, plus `document.querySelector(...)` typed as the generic `Element` rather than `HTMLElement` at every `.onclick`/`.focus`/`.dataset` use) -- not wired into CI as a gate, since clearing those honestly needs 16 scattered `@ts-expect-error` casts for zero additional bug-catching value over what this pass already found. One genuinely-loose spot fixed anyway: `localStorage.setItem` was handed a `parseInt(...)` result directly; wrapped in `String(...)` (a no-op at runtime -- `setItem` already coerces -- but now type-clean too).
- **Two synthetic PDF fixtures** (`tablekit_tests/fixtures/`, generated by `generate_fixtures.py`, a fabricated "Acme Test Holdings, Inc." -- never a real company): a digital-text income statement and a scanned/no-text-layer balance sheet, both built so every subtotal reconciles exactly. Getting them to actually classify and foot correctly took three real fixes, each verified against the live detector rather than assumed: (1) the row-ruling grid needs a line under every row, not just subtotals -- one missing line merges two adjacent line items into one row; (2) `extract_all_tables.py`'s statement-heading regex (`_STMT_RE`) is tuned for IFRS wording ("income statement", "statement of profit or loss") from the du/Etisalat statements this tool was built against -- "CONSOLIDATED STATEMENT OF OPERATIONS" (US GAAP phrasing) matched nothing, "CONSOLIDATED INCOME STATEMENT" does; (3) the year-header row must carry no label text in its own row -- combined with "Year ended December 31," on the same line, its "2024"/"2023" got read back as data-row figures, throwing the reconciliation check off by exactly one year value. Both fixtures now come back `foots: true` end-to-end (CLI, `serve.py`'s live UI, and the OCR failsafe path) -- not asserted, run and read back. New `tablekit_tests/test_fixtures.py` locks this in: CI's first real (non-PDF-gated, non-skipped) extraction test, since these two -- unlike the 24 real annual reports -- are actually committed (`.gitignore` carries a narrow `!tablekit_tests/fixtures/*.pdf` exception to the blanket `*.pdf` rule). Tested against real-world financial statements during development too (Microsoft's and Walmart's public annual reports, alongside the du/Etisalat statements already covered in earlier entries) -- those aren't committed; the synthetic pair covers the same digital-text and OCR cases for CI.
- **Accessibility audit** (axe-core 4.10, run live against the served app, both themes, every reachable view -- landing/onboarding, page-picker, extracted-table preview, compare panel): 16 real violations found and fixed, not just logged. Highlights: the primary "Extract this region" button had white text on its gold background (1.79:1, should be 4.5:1) -- the single most-used action in the app was nearly unreadable; the search box and page-number input were rendering dark-mode's near-white text color on a plain (never-themed) white input background, making them functionally invisible in dark mode (1.16:1); three status-pill colors (`--none` in both themes, `--bad` in dark, `--ok` in light) fell short of 4.5:1 against their own badge backgrounds; the app had no `<h1>` anywhere and a sidebar heading skipped straight from h1 to h4; five form controls (`#pgNum`, `#fileSel`, `#cmpFile`, `#pgImg`, the per-row checkboxes) had no accessible name at all. Fixed via the existing i18n mechanism where the element was static (`data-i18n-title` already sets both `title` and `aria-label` from one key) or inline `t()` calls where it's rendered dynamically -- no new pattern introduced. New `--accent-text` token added for the 4 spots that render `--accent` as literal text (light theme's `--accent` reads fine as a border/button-fill color but fails contrast as text; dark theme's doesn't need the distinction, so `--accent-text` just equals `--accent` there). Re-verified clean (0 violations) after each fix, in both themes, across every view -- including two false leads run down and ruled out empirically rather than "fixed" on faith: a live theme-toggle occasionally left a pre-existing button's background stuck on the old theme's computed value for 1s+ in this specific automated browser (a fresh page load in the same theme via `localStorage` rendered correctly immediately, and a freshly-created `.btn` element also rendered correctly mid-session -- isolating it to a transition/automation artifact on already-mounted nodes, not a CSS or app bug), and one `region` finding that didn't reproduce on an immediate re-run in the identical state.

## 0.7.1

Four findings from re-auditing 0.7.0 itself, not new source changes --
all four were cases of asserting something worked without checking.

- **The `frontend-logic` CI job's own test-running step probably never
  worked.** `node --test tablekit_tests/js/` (directory-scan mode) --
  verified locally, twice, in two different shells (Git-Bash and native
  PowerShell, ruling out a shell-specific path-translation artifact), with
  a minimal from-scratch reproduction directory unrelated to this project:
  Node 24 doesn't discover the test file that way at all. It tries to
  `require()` the bare directory path and fails with `MODULE_NOT_FOUND`
  before a single test runs. This job pins Node 20, which the above
  couldn't directly confirm either way -- but relying on version-sensitive
  auto-discovery behavior neither check could fully verify is exactly the
  unchecked-assumption pattern the other three findings below are also
  about, so it's fixed the same way: naming the file explicitly
  (`node --test tablekit_tests/js/test_webui_logic.js`) needs no such
  trust either way. Every "all N tests pass" claim made across this whole
  review chain was run directly against that file, never through this
  directory-mode invocation -- which is exactly how this stayed
  undiscovered through every prior pass.
- **CI coverage step didn't do what its own comment said.** Commented
  "informational only, not a gate" with no `continue-on-error`, so a
  `pytest-cov` hiccup would still fail the job. Also re-ran the entire
  suite a 4th time in the same job just to attach `--cov`. Fixed: `--cov`/
  `--cov-append` ride along on two steps that already have to run, a
  separate `coverage report -m` step (zero tests, just reads the
  already-collected data) carries `continue-on-error: true` -- now it
  actually can't fail the build, instead of merely saying so.
- **eslint.config.mjs's rule list was asserted "pyflakes-equivalent",
  unverified.** It was 15 hand-picked rules. Diffed against `@eslint/js`'s
  real `recommended` config (62 rules) for the first time -- the 47 missing
  ones (`no-cond-assign`, `no-case-declarations`, `no-dupe-else-if`,
  `no-async-promise-executor`, `no-unsafe-optional-chaining`, ...) turn out
  not to fire on this codebase, but nothing had checked that before the
  claim shipped. Now uses the real `recommended` config directly. (Along
  the way: `npx -p @eslint/js -p eslint eslint ...` looked like it should
  work and doesn't -- `import "@eslint/js"` in the config resolves relative
  to the config file's own location, which npx's package cache doesn't
  satisfy. Verified failing before switching CI to a real local
  `npm install --no-save`, which resolves correctly.)
- **The save-status indicator had a happy-path-only verification, same
  gap that let the debounce race ship.** Built specifically so a failed
  autosave is visible instead of silent, confirmed working live, shipped
  with no automated test of the failure path. Added two: a forced-failure
  case (asserts the exact `saving` -> `save-error` sequence, that
  `state.edits[n]` survives so export still works, that the timer handle
  is cleaned up) and a retry-after-failure case (the next edit actually
  re-reaches the server and reports `saved`, not silently swallowed).
  Writing them surfaced one more thing, not fixed here: if `renderPreview`
  itself throws after a successful save (it would, on a malformed detail
  object -- never happens with a real server response, only hit this
  building the test with a deliberately minimal fake one), the error
  handling reports `save-error` even though the save itself succeeded.
  Real server responses are always complete, so low priority, but it's a
  misleading message in that edge case, worth knowing about.

## 0.7.0

Closing out the remaining items from the last two review passes.

### Regression tests for last session's two extraction bugs
Both fixes (rows_in_box's region/box overlap logic, OCR pad=20) were only
ever verified live, by hand -- now locked in: a synthetic-region test for a
box deliberately drawn around only part of a bigger table (the balance-sheet
"Assets only" case that motivated the fix), and a signature-default test
that fails if `ocr_rows_in_box`'s `pad` is ever quietly reverted to 6.

### Coverage measurement
`pytest-cov` wired into CI as an informational report (not a gate -- the 24
sample PDFs that exercise most of the real extraction path aren't committed,
so CI coverage is structurally lower than a local run and not a fair
threshold to fail the build on).

### ESLint
`eslint.config.mjs` (pyflakes-equivalent rules only, same reasoning as
ruff's scoped ruleset) + `tablekit_tests/js/extract_inline_script.mjs` to
pull webui.html's inline `<script>` out to a plain .js file first (ESLint
lints .js, not HTML). Resolved via `npx` in CI, no package.json/node_modules
committed. Found and removed real dead code: a duplicate unused constant
(`PREVIEW_SCALE`, a copy-paste of `PICKER_SCALE` that nothing ever read), an
unused local in `boot()`, and a vestigial unused parameter on `undoLastEdit`
left over from writing it.

### Session-save calls, centralized
Every function that mutates a file's manual-table list now autosaves via
one `@_autosaves` decorator instead of a bare `_save_session(name)` call
copy-pasted into five separate function bodies -- a sixth mutating endpoint
added later can't as easily forget it now.

### A real silent-failure gap, found by audit
`showPicker()` cached a failed page-count fetch as `pageCount = 1`,
permanently (it only fetches once per file) -- silently capping a 171-page
report's picker at page 1 for the rest of the session, with "Next" actually
DISABLED (`page >= pageCount`) and no error shown anywhere. Fixed: on
failure it now shows an error toast and leaves the count unset so the next
attempt retries, instead of locking in a wrong answer. (Audited every other
silently-swallowed error in webui.html while at it -- the rest are all
deliberate, documented, non-critical localStorage/decorative-info fallbacks.)

### Compare-verdict-style i18n for the three fixed-wording table notes
`extract_all_tables.py`'s `analyze()` generates exactly three notes with
fixed wording (a year-reversal warning, a segmental-columns warning, an
assets-only-incomplete warning) among many that are fully dynamic prose
(reconcile-explain's worked arithmetic, etc.) -- small enough to do safely,
unlike localizing server-generated text in general (still a separate,
larger project -- see the comment above I18N in webui.html). `_add_note()`
sends each of those three in two lockstep forms: the English sentence
exactly as before (`notes`, unchanged for the CLI/Excel export) and a
`{key, vars}` pair (`notes_i18n`) webui.html translates when it recognises
the key, falling back to the English text otherwise -- including for the
many notes that were never given a key at all.

## 0.6.1

### Autosave race that could silently drop an edit
- `webui.html`'s debounced commit-to-server used ONE shared timer
  (`state.reTimer`) for every table. Editing table B within table A's 450ms
  debounce window cancelled A's pending save outright -- silently, with
  `state.edits[100002]` (etc.) still showing A as "edited" in the browser
  the whole time. Reproduced live before the fix (A's edit never reached
  the server; after, both land independently) and locked in as two node
  regression tests. Fixed: `state.reTimers` keyed per table.
- Fixing that naively (a timer per table, nothing else) opens a worse hole:
  if the user switches to a DIFFERENT FILE while a save is still pending,
  `state.edits` gets wiped by `loadFile()` and the same table number can
  mean an unrelated table in the new file -- firing the stale save there
  would be silent cross-file corruption, not just a dropped edit.
  `scheduleReanalyze` now captures `file` at scheduling time and no-ops if
  the active file has changed by the time the timer fires.
- Added a small save-status indicator ("Saving…" / "● Saved" / an error
  state) next to the edit toolbar, so a save landing (or failing) is
  something the user can actually see instead of a change that's either
  silently there or silently isn't.

## 0.6.0

### Session persistence -- the big one
- `serve.py` used to hold every manually-extracted table and every hand-typed
  edit in a plain in-memory dict. A server restart (or a page refresh, for
  in-progress edits specifically -- `reanalyze()` never wrote them back
  anywhere) lost all of it, silently, with no warning. Fixed:
  `_save_session`/`_load_session` (pickle, not JSON -- table dicts carry
  `FormattedNumber` cells and internal `_`-prefixed keys; this file is only
  ever written and read by this same process, so pickle's arbitrary-code-
  on-load risk doesn't add anything new) snapshot a file's manual tables to
  `uploads/.sessions/<file>.pkl` (atomic tmp-then-rename write) after every
  extraction, edit, delete, and undelete. `reanalyze()` now COMMITS the edit
  into the manual list instead of just returning a transient preview.
  `run_app.bat`'s new no-args launch (see 0.5.0) now also rediscovers
  whatever's already sitting in `uploads/` (`_discover_files`), so the
  session it restores is actually reachable.
- Table deletion was already a tombstone (index kept, so other tables' `n`
  never shifts) but threw the deleted table's CONTENT away with no way back.
  `undelete_manual` + a per-file undo stack (`(idx, table)` pairs) makes
  "Table removed. [Undo]" a real, working toast action.
- Cell/row edits (insert/delete/split/merge-up/swap-years) get a step-back
  undo too (`snapshotForUndo`/`undoLastEdit`, a capped 20-entry stack per
  table) via an "↺ Undo edit" toolbar button -- deliberately a button, not a
  global Ctrl+Z: a table cell is `contenteditable` with its OWN native
  browser undo, and hijacking Tab/Ctrl+Z globally would fight that instead
  of complementing it.
- The export tray's selection/order and the workbook name are cheap-to-redo
  browser preferences (not authoritative data), so those are mirrored to
  `localStorage` per file and restored silently on load -- no confirm
  prompt, since worst case is re-ticking a checkbox.
- `upload_pdf`'s filename sanitiser was still the ASCII-only ripped-out-and-
  reapplied ancestor of the fix already applied to the EXPORT filename last
  session -- an Arabic-named upload was getting mangled into underscores on
  the way IN even though it round-tripped fine on the way out. Fixed
  (`_sanitize_filename`, Unicode-aware via `str.isalnum()` rather than
  fighting `re`'s lack of `\p{L}` support).

### Testing
- `tablekit_tests/js/test_webui_logic.js` -- the first automated coverage of
  `webui.html`'s frontend logic (`fmt()`, `t()`/i18n interpolation, an
  en/ar key-symmetry check, `kindName()`, `isRisky()`). Runs the REAL inline
  `<script>` inside a loose DOM stub (`dom_stub.js`), not a reimplementation
  and not a real browser -- deliberately no Playwright/Puppeteer dependency;
  see docs/PIPELINE.md. `node --test tablekit_tests/js/`, zero npm deps,
  now CI-gated (`frontend-logic` job).
- `conftest.py` (new): autouse fixture redirects `serve.SESSION_DIR` to a
  tmp dir for every test, so the autosave added above never writes into the
  real `uploads/.sessions/`.
- `ruff` wired in CI (`lint` job), scoped to pyflakes only (`select = ["F"]`
  in pyproject.toml) -- real correctness bugs, not this codebase's own
  established style (short reused names, semicolon one-liners) fighting a
  broader ruleset for no bug-catching benefit. Found and fixed: a fully
  dead function (`extract_all_tables.page_labels` -- zero callers anywhere,
  its own `pypdf` read was constructed and immediately discarded), four
  more unused-variable/import cases, two pointless f-string prefixes.
  (One near-miss: pyflakes also flagged `parse_number` as unused within
  `extract_all_tables.py` -- correctly, by its own single-file view, but
  that import is a DELIBERATE re-export `tablekit_tests/test_golden.py`
  depends on via `X.parse_number()`; restored with a documented
  `noqa: F401` rather than removed.)
- `launcher-sanity` CI job: greps `run_app.bat` for `serve.py`. The exact
  drift from 0.5.0 (the launcher silently pointing at a dead pre-rewrite
  app) now fails CI instead of waiting for someone to notice by hand.
- `serve.compare()`'s new `counts` field (see below) gets its own assertion
  in `test_compare_returns_a_diff`, which needed a real fix along the way:
  the test's two fake tables both claimed the same year pair, so the
  "restated" branch (A's PRIOR year vs B's CURRENT year) could never fire
  no matter how different the figures were -- the original loose
  `"changed" in verdict or "RESTATED" in verdict` assertion had been
  quietly tolerating that. Fixed the fixture to use an actually-offset year
  pair, matching the real use case (this year's report vs last year's).

### Compare-panel verdict, partially localized
- `diff_tables()` now returns `(rows, verdict, counts)` -- `counts` is the
  same result as plain numbers (`changed`/`new`/`removed`/`restated`)
  alongside the existing English `verdict` sentence, which stays as-is for
  the CLI/xlsx output. `webui.html` builds a translated sentence from
  `counts` when present (`cmpVerdictRestated`/`cmpVerdictClean` i18n keys),
  falling back to the raw English string against an older server response.
  This is NOT the larger project of localizing every server-generated
  string (a table's `notes`, `reconcile_explain`) -- see the comment above
  `I18N` in webui.html for why that's a separate, bigger change to
  `extract_all_tables.py`'s prose generation; this is the one contained
  piece of it that was safe to do without touching the trust layer.

### Accessibility
- Icon-only controls (theme toggle, the export tray's ↑/↓/× buttons, every
  toast's × dismiss, the row-editor's ＋/✕/⤶/⭡ buttons) now carry
  `aria-label` alongside their existing `title` -- title alone isn't
  reliably announced, and several of these have no visible text for a
  screen reader to fall back on.
- The confirm modal (`showModal`) now has `role="dialog"`, `aria-modal`,
  `aria-labelledby`, and an actual focus trap (Tab/Shift+Tab wrap between
  its two buttons instead of leaking focus into the page behind it).

### Multi-file queue
- The file `<select>` (previously hidden entirely until a second file was
  loaded, and unstyled) now shows each file's extracted-table count in its
  label and gets real styling matching the rest of the header. Not batch
  AUTOMATION -- that would need auto-detection back, which was removed for
  being unreliable (see 0.5.0) -- just visibility into which of several
  loaded files still need manual work.

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
