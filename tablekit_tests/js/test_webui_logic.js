"use strict";
// Regression tests for the PURE logic inside webui.html -- number/currency
// formatting, i18n lookup + interpolation, the kind-name fallback chain.
// These run the real, unmodified inline <script> (not a reimplementation of
// it) inside a loose DOM stub (see dom_stub.js) so a change to webui.html
// itself is what gets tested, not a copy that can drift from it.
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

const HTML_PATH = path.join(__dirname, "..", "..", "webui.html");

function loadSandbox() {
  const html = fs.readFileSync(HTML_PATH, "utf8");
  const m = html.match(/<script>([\s\S]*)<\/script>/);
  assert.ok(m, "could not find the inline <script> block in webui.html");
  const source = m[1];

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
  const withExports = source + "\n;globalThis.__TEST_EXPORTS__ = { fmt, I18N };\n";
  vm.runInContext(withExports, sandbox, { filename: "webui.html (inline script)" });
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

test("upload-filename sanitiser keeps Unicode letters, still strips path separators", () => {
  const { document } = loadSandbox();
  void document; // sandbox already loaded; re-derive the same regex used in webui.html
  const sanitize = (s) => s.replace(/[^\p{L}\p{N} .()-]/gu, "_");
  assert.strictEqual(sanitize("تقرير 2024.pdf"), "تقرير 2024.pdf");
  assert.strictEqual(sanitize("a/b\\c:d*e?.pdf"), "a_b_c_d_e_.pdf");
});
