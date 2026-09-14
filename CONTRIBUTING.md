# Contributing

Read `docs/PIPELINE.md` first if you're touching detection/analysis code --
it's the maintainer's map of how a PDF turns into a verified table. This
file is about the mechanics: setup, tests, and what a PR needs before it's
mergeable.

## Setup

```
pip install -r requirements.txt
pip install -e ".[dev]"
```

`requirements.txt` covers the runtime dependencies (including the img2table/
opencv second geometry engine -- not optional in practice, see its own
comments there); `.[dev]` adds pytest, ruff, mypy, hypothesis. OCR
(`pytesseract`) additionally needs the Tesseract binary installed separately
-- not pip-installable, see `requirements.txt`.

## Running things locally

```
pytest tablekit_tests/ -q                              # everything
ruff check extract_all_tables.py serve.py telecom_extract.py tablekit/ tablekit_tests/
mypy extract_all_tables.py serve.py telecom_extract.py tablekit/
```

`.github/workflows/tests.yml` is the actual source of truth for what gates a
PR -- if something passes locally but you're unsure it matches CI, read that
file rather than guess.

A few suites behave differently depending on what's on your machine, not as
a bug but by design:

- **The 24 real annual-report PDFs** this project was developed against
  (`golden.json`'s snapshot cases, most of `test_manual_mode.py`) aren't
  committed -- `.gitignore` blanket-excludes `*.pdf` at the repo root. They
  run and gate locally if you have them; they skip cleanly in CI. Don't try
  to make CI exercise them by committing real reports, even ones that seem
  publicly available -- see the copyright note below.
- **`tablekit_tests/fixtures/*.pdf`** (a fabricated "Acme Test Holdings,
  Inc.", never a real company) are the exception to that `*.pdf` rule and
  ARE committed, specifically so CI has at least one real, non-skipped,
  end-to-end extraction test. Regenerate them with
  `python tablekit_tests/fixtures/generate_fixtures.py` (needs
  `pip install reportlab` first -- not a project dependency, only used here).
- **OCR-specific tests** (`test_scanned_fixture_ocrs_as_a_footing_balance_sheet`
  and similar) skip without the Tesseract binary. Windows: see the link in
  `requirements.txt`.
- **The frontend suites** (`node --test tablekit_tests/js/test_webui_logic.js`,
  `npx eslint --config eslint.config.mjs webui.js`) need Node 20+ and, for
  lint, a local `npm install --no-save eslint@9 @eslint/js@9` -- no
  `package.json` is committed, matching how ESLint is pulled fresh in CI.
  `webui.html`'s script lives in `webui.js` as a real file (split out in
  0.7.4, see `CHANGELOG.md`) specifically so tools like these run directly
  against it, no extraction step needed. `tsc --allowJs --checkJs --noEmit
  webui.js` is a useful local check but isn't wired into CI -- see 0.7.2 in
  `CHANGELOG.md` for why (the remaining findings are `Element` vs
  `HTMLElement` DOM-typing noise, not real bugs, and clearing them honestly
  needs more scattered casts than the check is worth as a gate).

## Before opening a PR

- **Run the tests that apply to what you touched**, at minimum. If you
  changed detection/analysis and have the real PDFs locally, run the golden
  suite -- CI can't catch a regression there for you.
- **Verify claims, don't assert them.** This codebase's history (see
  `CHANGELOG.md`, especially the 0.7.x entries) has repeatedly found real
  bugs specifically by re-checking something that "should" have worked --
  a CI step whose core command silently never ran, a lint config asserted
  equivalent to a stricter one without ever being diffed against it, a debug
  line that crashed the one code path nobody happened to test. If you're
  not sure something works, run it and look, rather than describe what it
  should do.
- **New tooling (a linter, a type checker, a new test category) should be
  gradual, not a wall.** Run it against the real codebase first and look at
  what it actually finds before deciding what to enable -- see
  `[tool.mypy]` and `[tool.ruff.lint]` in `pyproject.toml` for the reasoning
  and the precedent. A tool that reports thousands of findings on code
  nobody's touching trains people to ignore it.
- **Update `CHANGELOG.md`** for anything a user or future contributor would
  want to know about -- not just what changed, but why, especially if the
  change was a fix for something that was asserted to work and wasn't. CI's
  `changelog-nudge` job warns (doesn't block) if source changed without it.
- **Never commit a real third-party PDF**, even a publicly-available annual
  report. Documenting that the tool was tested against one is fine and
  encouraged (see recent `CHANGELOG.md` entries for the pattern); committing
  the actual file is a redistribution question this project treats as
  out of scope -- use or extend the synthetic fixtures instead.
