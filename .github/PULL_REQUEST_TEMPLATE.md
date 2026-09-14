<!-- See CONTRIBUTING.md for the reasoning behind each of these. -->

## What changed, and why

<!-- The "why" matters more than the "what" here -- the diff already shows
what changed. -->

## Verification

<!-- What did you actually run, and what did it show? "Should work" isn't
verification -- run it and say what happened. -->

- [ ] Ran the relevant test suite(s) locally (`pytest tablekit_tests/ -q`, or a
      narrower subset if that's what applies) -- results:
- [ ] `ruff check` / `mypy` clean on anything Python this touches
- [ ] If this touches `webui.html`: exercised it in a browser, not just read
      the diff
- [ ] `CHANGELOG.md` updated (or this doesn't need it -- a typo, a comment, etc.)

## Anything committed that shouldn't be

- [ ] No real third-party PDF (an annual report, a sample someone sent you) --
      only the synthetic fixtures under `tablekit_tests/fixtures/`, or nothing
