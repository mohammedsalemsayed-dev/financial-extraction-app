# Archive

Historical files, kept for reference but not part of the running app.
Nothing in `extract_all_tables.py`, `serve.py`, or `tablekit/` imports or
depends on anything in this folder.

- **`financial_extract/`** — the original "v0.1 MVP shell": a localhost UI
  hardcoded to two tables (income statement + opex note) for du and
  Etisalat/e& only, no OCR, its own small `web/` frontend. Fully superseded
  by `serve.py` + `webui.html` + `tablekit/` (any statement, any company,
  manual box-select, OCR failsafe, the arithmetic-reconciliation trust
  layer). `run_app.bat` used to launch this instead of `serve.py` — a
  leftover from before the rewrite, not a deliberate choice — and was fixed
  to point at the real app. Kept for reference only.
- **`extract_two_tables.py`** — the original, DFM-only predecessor: hardcoded
  to pull exactly two named tables (profit-or-loss, G&A expenses) out of one
  specific company's report wording. `extract_all_tables.py` fully supersedes
  it (generic detection across every table, any company, with the
  arithmetic-reconciliation trust layer this never had). Superseded, not
  removed, in case its DFM-specific heading/vocabulary patterns are ever
  useful reference for tuning a new company's support.
- **`HANDOFF.md`** — the project-history writeup for `extract_two_tables.py`
  above (bug catalog, DFM-specific tuning notes). Describes that earlier
  script, not the current app — kept for the debugging lessons documented
  in it (cross-column contamination, glued header lines, etc. — the *same
  classes* of bug this session re-found and re-fixed in the current engine),
  not as current documentation. See `docs/PIPELINE.md` and `CHANGELOG.md`
  for how the app actually works today.
