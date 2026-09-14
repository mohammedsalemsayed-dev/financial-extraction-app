// ESLint config for webui.js (webui.html's script, split out to its own
// file in 0.7.3 -- see CHANGELOG.md; no extraction step needed any more,
// ESLint just lints the real file directly). No package.json / node_modules
// committed to the repo --
// CI installs both this and @eslint/js fresh into the checkout
// (`npm install --no-save`, gone with the runner afterward), matching the
// project's zero-committed-JS-dependency approach (see
// tablekit_tests/js/dom_stub.js). Needs a real local install, not
// `npx -p @eslint/js -p eslint eslint ...`: the `import "@eslint/js"`
// below resolves relative to THIS file's own location, which npx's package
// cache doesn't satisfy -- verified failing (ERR_MODULE_NOT_FOUND) before
// switching to a local install, which works.
//
// Uses @eslint/js's own `recommended` config rather than a hand-picked rule
// list: an earlier version of this file picked ~15 rules that "sounded
// right" and asserted (in the commit message, unchecked) that this was
// pyflakes-equivalent rigor. It wasn't verified against anything -- diffing
// it against the real `recommended` set (62 rules) found the gap was
// entirely lucky: everything recommended-but-missing (no-cond-assign,
// no-case-declarations, no-dupe-else-if, no-async-promise-executor,
// no-unsafe-optional-chaining, ...) turned out not to fire on this
// codebase, but nothing had actually checked that before shipping the
// claim. This is the real thing now, not a guess at it.
import js from "@eslint/js";

export default [
  {
    ...js.configs.recommended,
    languageOptions: {
      ecmaVersion: 2022,
      sourceType: "script",
      globals: {
        window: "readonly", document: "readonly", localStorage: "readonly",
        navigator: "readonly", fetch: "readonly", console: "readonly",
        setTimeout: "readonly", clearTimeout: "readonly",
        setInterval: "readonly", clearInterval: "readonly",
        FileReader: "readonly", URL: "readonly", Blob: "readonly",
        KeyboardEvent: "readonly", Event: "readonly", Set: "readonly",
        Promise: "readonly", requestAnimationFrame: "readonly",
        getComputedStyle: "readonly",
      },
    },
    rules: {
      ...js.configs.recommended.rules,
      // Both overrides verified against this file's actual patterns before
      // being added, not assumed:
      // - catch(e){ /* deliberately ignored */ } is common and intentional
      //   throughout (a failed localStorage read, a non-critical status
      //   fetch) -- flagging every unused catch binding is noise.
      // - allowEmptyCatch does the same for no-empty specifically (a
      //   DIFFERENT rule than no-unused-vars -- recommended's no-empty
      //   flags empty BLOCKS regardless of the catch binding, so this one
      //   needs its own opt-out or every catch(e){} above trips it too).
      "no-unused-vars": ["error", { caughtErrors: "none" }],
      "no-empty": ["error", { allowEmptyCatch: true }],
    },
  },
];
