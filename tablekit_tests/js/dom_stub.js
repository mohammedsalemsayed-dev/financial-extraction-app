"use strict";
// A deliberately loose DOM stub -- NOT a real DOM. Its only job is to let
// webui.html's inline <script> execute top-to-bottom without throwing (it
// wires a handful of element.onclick/.oninput/.addEventListener at parse
// time), so the PURE logic defined alongside that wiring -- fmt(), t(),
// kindName(), the i18n dictionary, the risky-table check, etc. -- becomes
// callable and testable from Node. Anything actually rendering to a page
// still needs a real browser; see the Playwright-alternative note in
// CHANGELOG.md/docs/PIPELINE.md for why this project doesn't pull in a
// browser-automation dependency for that instead.

function makeStubElement(tag) {
  const store = { value: "", textContent: "", innerHTML: "", checked: false,
                   dataset: {}, style: {}, hidden: false, disabled: false,
                   className: "", title: "" };
  const classList = { add(){}, remove(){}, toggle(){}, contains(){ return false; } };
  const handler = {
    get(target, prop) {
      if (prop === "classList") return classList;
      if (prop === "dataset") return store.dataset;
      if (prop === "style") return store.style;
      if (prop === "children" || prop === "childNodes") return [];
      if (prop === "querySelector") return () => makeStubElement();
      if (prop === "querySelectorAll") return () => [];
      if (prop === "addEventListener" || prop === "removeEventListener") return () => {};
      if (prop === "appendChild" || prop === "prepend" || prop === "insertBefore"
          || prop === "remove" || prop === "focus" || prop === "click"
          || prop === "getContext" || prop === "scrollIntoView") {
        return () => (prop === "getContext" ? makeStubElement() : undefined);
      }
      if (prop in store) return store[prop];
      if (prop === "firstChild") return makeStubElement();
      if (typeof prop === "symbol") return undefined;
      // catch-all: anything else accessed as a function is a harmless no-op
      return (..._args) => makeStubElement();
    },
    set(target, prop, value) { store[prop] = value; return true; },
  };
  return new Proxy({}, handler);
}

function makeDocumentStub() {
  const root = makeStubElement("html");
  root.dataset = {};
  root.classList = { add(){}, remove(){}, toggle(){}, contains(){ return false; } };
  const body = makeStubElement("body");
  body.classList = { add(){}, remove(){}, toggle(){}, contains(){ return false; } };
  return {
    documentElement: root,
    body,
    title: "",
    querySelector: () => makeStubElement(),
    querySelectorAll: () => [],
    getElementById: () => makeStubElement(),
    createElement: () => makeStubElement(),
    addEventListener: () => {},
    removeEventListener: () => {},
  };
}

function makeLocalStorageStub() {
  const data = {};
  return {
    getItem: (k) => (k in data ? data[k] : null),
    setItem: (k, v) => { data[k] = String(v); },
    removeItem: (k) => { delete data[k]; },
  };
}

module.exports = { makeDocumentStub, makeLocalStorageStub, makeStubElement };
