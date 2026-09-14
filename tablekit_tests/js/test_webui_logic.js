"use strict";
// Regression tests for the PURE logic in webui.js (webui.html's script,
// split into its own file in 0.7.3 -- see CHANGELOG.md) -- number/currency
// formatting, i18n lookup + interpolation, the kind-name fallback chain.
// These run the real, unmodified file (not a reimplementation of it) inside
// a loose DOM stub (see dom_stub.js) so a change to webui.js itself is what
// gets tested, not a copy that can drift from it.
// Deliberately NOT a real browser: no rendering, no layout, no click
// simulation -- see docs/PIPELINE.md for why this project doesn't pull in a
// browser-automation dependency for that instead. Run with:
//   node tablekit_tests/js/test_webui_logic.js
const assert = require("node:assert");
const test = require("node:test");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const { makeDocumentStub, makeLocalStorageStub } = require("./dom_stub.js");

const JS_PATH = path.join(__dirname, "..", "..", "webui.js");

function loadSandbox() {
  const source = fs.readFileSync(JS_PATH, "utf8");

  const sandbox = {
    console,
    document: makeDocumentStub(),
    localStorage: makeLocalStorageStub(),
    window: {
      devicePixelRatio: 1,
      innerWidth: 1200,
      matchMedia: () => ({ matches: false }),
      addEventListener: () => {},
    },
    navigator: { language: "en" },
    setTimeout, clearTimeout, setInterval, clearInterval,
    // the script's own boot() fires two real fetches (unawaited-with-catch
    // on the second) the instant it loads -- resolve both harmlessly rather
    // than reject, so loading the sandbox never produces an unhandled
    // rejection. A shape that satisfies both /api/status and /api/files.
    fetch: () => Promise.resolve({
      ok: true, status: 200,
      json: () => Promise.resolve({ files: [], img2table: false, ocr: false }),
    }),
  };
  sandbox.window.localStorage = sandbox.localStorage;
  vm.createContext(sandbox);
  // top-level `const`/`let` bindings (fmt, I18N, ...) do NOT become
  // properties of the sandbox global the way `function`/`var` do -- append
  // one line, in the SAME scope, that reaches them as plain identifiers and
  // republishes what this harness needs onto the context object itself.
  const withExports = source + "\n;globalThis.__TEST_EXPORTS__ = { fmt, I18N, state };\n";
  vm.runInContext(withExports, sandbox, { filename: "webui.js" });
  Object.assign(sandbox, sandbox.__TEST_EXPORTS__);
  return sandbox;
}

test("fmt() leaves plain numbers alone and re-applies a prefix/suffix only for display", () => {
  const { fmt } = loadSandbox();
  assert.strictEqual(fmt(1234), "1,234");
  assert.strictEqual(fmt(1234.5), "1,234.5");
  assert.strictEqual(fmt(26.6, { p: "", s: "%" }), "26.6%");
  assert.strictEqual(fmt(67344574000, { p: "$", s: "" }), "$67,344,574,000");
  assert.strictEqual(fmt("already a string"), "already a string");
  assert.strictEqual(fmt(null), "");
});

test("t() falls back EN -> key, and interpolates {placeholders}", () => {
  const sandbox = loadSandbox();
  const { t } = sandbox;
  assert.strictEqual(t("brand"), "Tables");
  assert.strictEqual(t("exportOrderN", { n: 3 }), "Export order — 3");
  // a genuinely unknown key falls back to returning the key itself, not a crash
  assert.strictEqual(t("__does_not_exist__"), "__does_not_exist__");
});

test("t() switches dictionaries with document.documentElement.lang", () => {
  const sandbox = loadSandbox();
  sandbox.document.documentElement.lang = "ar";
  assert.strictEqual(sandbox.t("cancelBtn"), "إلغاء");
  sandbox.document.documentElement.lang = "en";
  assert.strictEqual(sandbox.t("cancelBtn"), "Cancel");
});

test("every I18N key present in English also exists in Arabic (and vice versa)", () => {
  const { I18N } = loadSandbox();
  const enKeys = new Set(Object.keys(I18N.en));
  const arKeys = new Set(Object.keys(I18N.ar));
  const missingFromAr = [...enKeys].filter(k => !arKeys.has(k));
  const missingFromEn = [...arKeys].filter(k => !enKeys.has(k));
  assert.deepStrictEqual(missingFromAr, [], "keys in en but missing from ar");
  assert.deepStrictEqual(missingFromEn, [], "keys in ar but missing from en");
});

test("kindName() prefers the active translation, falls back to the English kind name", () => {
  const sandbox = loadSandbox();
  const { kindName } = sandbox;
  assert.strictEqual(kindName("income statement"), "Income statement");
  sandbox.document.documentElement.lang = "ar";
  assert.notStrictEqual(kindName("income statement"), "Income statement");
  sandbox.document.documentElement.lang = "en";
  // an unrecognised kind falls all the way back to the raw string, not a crash
  assert.strictEqual(kindName("__unknown_kind__"), "__unknown_kind__");
});

test("isRisky() flags NO FOOT or low health, matching serve.py's health_bad_below (0.75)", () => {
  const { isRisky } = loadSandbox();
  assert.strictEqual(isRisky({ foots: false, health: 0.99 }), true);
  assert.strictEqual(isRisky({ foots: true, health: 0.5 }), true);
  assert.strictEqual(isRisky({ foots: true, health: 0.75 }), false);   // boundary: not < 0.75
  assert.strictEqual(isRisky({ foots: true, health: 0.9 }), false);
  assert.strictEqual(isRisky({ foots: null, health: null }), false);  // unscored, not flagged
  assert.strictEqual(isRisky(null), false);
  assert.strictEqual(isRisky(undefined), false);
});

test("noteLines() translates a recognised note key, falls back to English for an unrecognised one", () => {
  const sandbox = loadSandbox();
  const { noteLines } = sandbox;
  const d = {
    notes: [
      "year columns may be reversed — the page header lists years oldest-first (2019…2021); figures could be attributed to the wrong year. Use 'Swap year columns' if so.",
      "some fully-dynamic reconcile-explain sentence with no i18n key at all",
    ],
    notes_i18n: [
      { key: "yearsReversed", vars: { first: 2019, last: 2021 } },
      null,   // extract_all_tables.py only ever sends {key,vars} for the
              // fixed-wording subset -- most real notes leave this null
    ],
  };
  const en = noteLines(d);
  assert.ok(en[0].includes("2019") && en[0].includes("2021"));
  assert.strictEqual(en[1], d.notes[1], "no key recognised -> falls back to the raw English text");

  sandbox.document.documentElement.lang = "ar";
  const ar = noteLines(d);
  assert.notStrictEqual(ar[0], en[0], "a recognised key must actually translate under ar");
  assert.strictEqual(ar[1], d.notes[1], "still falls back to English -- there's no ar text to fall back TO here");
});

test("noteLines() falls back cleanly when notes_i18n is entirely absent (older server response)", () => {
  const { noteLines } = loadSandbox();
  const d = { notes: ["plain english note, no notes_i18n field at all"] };
  assert.deepStrictEqual(noteLines(d), d.notes);
});

test("upload-filename sanitiser keeps Unicode letters, still strips path separators", () => {
  const { document } = loadSandbox();
  void document; // sandbox already loaded; re-derive the same regex used in webui.js
  const sanitize = (s) => s.replace(/[^\p{L}\p{N} .()-]/gu, "_");
  assert.strictEqual(sanitize("تقرير 2024.pdf"), "تقرير 2024.pdf");
  assert.strictEqual(sanitize("a/b\\c:d*e?.pdf"), "a_b_c_d_e_.pdf");
});

// A recording fetch mock robust to both call shapes in this file: bare GETs
// (jget calls fetch(url) with no second argument -- boot()'s own /api/status
// and /api/files calls still fire in the background when a test reassigns
// sandbox.fetch, since loadSandbox() doesn't block on boot()'s promise) and
// POSTs (jpost always passes {body: "..."}).
function makeRecordingFetch(calls) {
  return (url, opts) => {
    const body = opts && opts.body ? JSON.parse(opts.body) : null;
    calls.push({ url, body });
    return Promise.resolve({
      ok: true, status: 200,
      json: () => Promise.resolve(body ? { n: body.n, rows: [] } : { files: [], img2table: false, ocr: false }),
    });
  };
}

// Regression test for a real bug found and fixed in this session: the
// debounced autosave used to key off ONE shared timer (state.reTimer) for
// every table, so editing table B within table A's 450ms debounce window
// cancelled A's pending save outright -- silently, with the browser still
// showing A as "edited". Reproduced live in a real browser before the fix
// landed. scheduleReanalyze now keys the timer per table (state.reTimers).
test("scheduleReanalyze: editing a second table does not cancel the first table's pending save", async () => {
  const sandbox = loadSandbox();
  const calls = [];
  sandbox.fetch = makeRecordingFetch(calls);
  sandbox.state.file = "a.pdf";
  sandbox.state.edits[100001] = { rows: [["A_EDIT"]] };
  sandbox.state.edits[100002] = { rows: [["B_EDIT"]] };

  sandbox.scheduleReanalyze(100001);   // table A: schedules a save in 450ms
  sandbox.scheduleReanalyze(100002);   // table B: must NOT cancel A's timer

  await new Promise(r => setTimeout(r, 650));   // both debounces have long since fired

  const reanalyzeCalls = calls.filter(c => c.url === "/api/reanalyze");
  const savedNs = reanalyzeCalls.map(c => c.body.n).sort((a, b) => a - b);
  assert.deepStrictEqual(savedNs, [100001, 100002],
    "both tables' edits should have reached the server, not just the most recently edited one");
});

// A second, related risk a naive per-table-timer fix could introduce: if
// the user switches to a DIFFERENT FILE while a save is still pending, that
// pending call must not later fire against the new file (state.edits keys
// are just numbers -- the same n could mean an unrelated table in the new
// file, and sending stale content there would be silent cross-file
// corruption, not just a dropped edit).
test("scheduleReanalyze: a pending save does not fire against a file switched to afterward", async () => {
  const sandbox = loadSandbox();
  const calls = [];
  sandbox.fetch = makeRecordingFetch(calls);
  sandbox.state.file = "a.pdf";
  sandbox.state.edits[100001] = { rows: [["A_EDIT"]] };
  sandbox.scheduleReanalyze(100001);
  sandbox.state.file = "b.pdf";   // switched files before the 450ms debounce fired

  await new Promise(r => setTimeout(r, 650));

  const reanalyzeCalls = calls.filter(c => c.url === "/api/reanalyze");
  assert.strictEqual(reanalyzeCalls.length, 0,
    "a save scheduled against file a.pdf must not fire once the active file is b.pdf");
});

// The save-status indicator (setSaveStatus/"Saving…"/"Saved"/error) exists
// specifically because the debounce race above was invisible by design --
// building it and confirming the happy path live, without ever locking in
// the FAILURE path it exists for, would repeat exactly the gap that let the
// race ship unnoticed. Spies on setSaveStatus rather than inspecting DOM
// state: both are top-level `function` declarations in the same sandboxed
// script, so reassigning sandbox.setSaveStatus after load is visible to
// scheduleReanalyze's own call to it, the same mechanism the fetch mock
// above already relies on -- and it sidesteps dom_stub.js's stub elements
// not persisting attribute state across separate querySelector() calls,
// which would make asserting on "the live DOM" here more theatre than test.
test("scheduleReanalyze: a failed save reports save-error, not a silent drop", async () => {
  const sandbox = loadSandbox();
  // only /api/reanalyze fails -- boot()'s own /api/status + /api/files
  // calls (fired in the background the instant the sandbox loads) still
  // need to resolve harmlessly, or boot()'s un-awaited second jget() call
  // becomes an unhandled rejection unrelated to what this test checks
  sandbox.fetch = (url) => String(url).includes("/api/reanalyze")
    ? Promise.reject(new Error("simulated network failure"))
    : Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve({ files: [], img2table: false, ocr: false }) });
  const statusCalls = [];
  sandbox.setSaveStatus = (n, kind) => statusCalls.push({ n, kind });
  sandbox.state.file = "a.pdf";
  sandbox.state.active = 100001;
  sandbox.state.edits[100001] = { rows: [["A_EDIT"]] };

  sandbox.scheduleReanalyze(100001);
  await new Promise(r => setTimeout(r, 650));

  assert.deepStrictEqual(statusCalls, [
    { n: 100001, kind: "saving" },
    { n: 100001, kind: "save-error" },
  ]);
  // the timer handle is cleaned up on failure same as on success, or a
  // stale reference could interfere with the next scheduleReanalyze(n) --
  // see the unconditional `delete state.reTimers[n]` at the top of the
  // real timer callback, before the try/catch that can fail
  assert.strictEqual(sandbox.state.reTimers[100001], undefined);
  // state.edits[n] must survive the failure -- it's what a retry (the next
  // edit) or export sends, regardless of whether any autosave ever landed
  assert.deepStrictEqual(sandbox.state.edits[100001], { rows: [["A_EDIT"]] });
});

test("scheduleReanalyze: a retry after a failed save succeeds and reports saved", async () => {
  const sandbox = loadSandbox();
  let attempt = 0;
  sandbox.fetch = (url, opts) => {
    if (String(url).includes("/api/reanalyze")) {
      attempt += 1;
      if (attempt === 1) return Promise.reject(new Error("simulated network failure"));
      const body = JSON.parse(opts.body);
      return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve({ n: body.n, rows: [] }) });
    }
    return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve({ files: [], img2table: false, ocr: false }) });
  };
  const statusCalls = [];
  sandbox.setSaveStatus = (n, kind) => statusCalls.push({ n, kind });
  sandbox.state.file = "a.pdf";
  // deliberately NOT 100001: a successful reanalyze() also calls
  // renderPreview() when its table is the active one, and this test's fake
  // {n, rows:[]} response is nowhere near the full shape a real server
  // reply has (kind/title/years/notes/notes_i18n/value_cols/...) --
  // renderPreview would throw on it. This test is about the retry actually
  // reaching the server and the status sequence, not about rendering, so
  // leaving no table active skips that call entirely rather than needing a
  // hand-built full detail object just to avoid a crash unrelated to what's
  // under test here.
  sandbox.state.active = null;
  sandbox.state.edits[100001] = { rows: [["FIRST_EDIT"]] };

  sandbox.scheduleReanalyze(100001);           // fails
  await new Promise(r => setTimeout(r, 650));
  sandbox.state.edits[100001] = { rows: [["RETRY_EDIT"]] };
  sandbox.scheduleReanalyze(100001);           // retry -- must actually re-fire, not be skipped
  await new Promise(r => setTimeout(r, 650));

  assert.strictEqual(attempt, 2, "the retry must reach the server a second time, not be swallowed");
  assert.deepStrictEqual(statusCalls.map(c => c.kind), ["saving", "save-error", "saving", "saved"]);
});
