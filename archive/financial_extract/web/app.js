"use strict";

const state = {
  docs: [],            // {doc_id, file, company, company_key, year, pages, status}
  extractions: {},     // doc_id -> extract_document result (body arrays are live/edited)
  selected: null,
  pageByCard: {},      // `${doc_id}:${i}` -> current page shown
};

const $ = (s, r = document) => r.querySelector(s);
const el = (tag, cls, txt) => {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (txt != null) e.textContent = txt;
  return e;
};

function toast(msg, isErr) {
  const t = $("#toast");
  t.textContent = msg;
  t.className = "toast" + (isErr ? " err" : "");
  t.hidden = false;
  clearTimeout(toast._t);
  toast._t = setTimeout(() => (t.hidden = true), isErr ? 4200 : 2200);
}

async function api(path, opts) {
  const r = await fetch(path, opts);
  const ct = r.headers.get("content-type") || "";
  if (ct.includes("application/json")) {
    const j = await r.json();
    if (!r.ok || j.error) throw new Error(j.error || ("HTTP " + r.status));
    return j;
  }
  if (!r.ok) throw new Error("HTTP " + r.status);
  return r;
}

/* ---------------- upload ---------------- */
async function uploadFiles(files) {
  for (const f of files) {
    if (!f.name.toLowerCase().endsWith(".pdf")) { toast(`skipped ${f.name} (not a PDF)`, true); continue; }
    try {
      const buf = await f.arrayBuffer();
      const rec = await api("/api/upload", {
        method: "POST",
        headers: { "X-Filename": f.name, "Content-Type": "application/pdf" },
        body: buf,
      });
      state.docs.push({ ...rec, status: "idle" });
      renderDocList();
      // auto-extract the first one for immediate feedback
      if (state.docs.length === 1) selectDoc(rec.doc_id);
    } catch (e) {
      toast(`upload failed: ${e.message}`, true);
    }
  }
  refreshExportBtn();
}

/* ---------------- doc list ---------------- */
function statusDot(s) {
  const d = el("span", "dot");
  if (s === "PASS") d.classList.add("pass");
  else if (s === "FAIL") d.classList.add("fail");
  else if (s === "PARTIAL") d.classList.add("partial");
  else if (s === "busy") d.classList.add("busy");
  return d;
}

function docOverallStatus(doc_id) {
  const ex = state.extractions[doc_id];
  if (!ex) {
    const d = state.docs.find((x) => x.doc_id === doc_id);
    return d && d.status === "busy" ? "busy" : "idle";
  }
  const st = ex.targets.map((t) => t.status);
  if (st.includes("FAIL")) return "FAIL";
  if (st.every((s) => s === "PASS")) return "PASS";
  return "PARTIAL";
}

function renderDocList() {
  const list = $("#docList");
  list.innerHTML = "";
  for (const d of state.docs) {
    const li = el("li", "doc" + (d.doc_id === state.selected ? " active" : ""));
    li.onclick = () => selectDoc(d.doc_id);
    li.appendChild(el("div", "doc-name", d.file));
    const sub = el("div", "doc-sub");
    sub.appendChild(statusDot(docOverallStatus(d.doc_id)));
    sub.appendChild(el("span", null,
      `${d.company} · ${d.year || "?"} · ${d.pages || "?"}p`));
    li.appendChild(sub);
    list.appendChild(li);
  }
  $("#docCount").textContent = state.docs.length;
}

/* ---------------- extract + render ---------------- */
async function selectDoc(doc_id) {
  state.selected = doc_id;
  renderDocList();
  if (!state.extractions[doc_id]) {
    const d = state.docs.find((x) => x.doc_id === doc_id);
    d.status = "busy";
    renderDocList();
    renderStageLoading(d);
    try {
      const res = await api("/api/extract", {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ doc_id }),
      });
      state.extractions[doc_id] = res;
    } catch (e) {
      toast(`extraction failed: ${e.message}`, true);
      d.status = "idle";
      renderDocList();
      return;
    }
    d.status = "done";
  }
  renderDocList();
  renderStage(doc_id);
  refreshExportBtn();
}

function renderStageLoading(d) {
  const stage = $("#stage");
  stage.innerHTML = "";
  const h = el("div", "dochead");
  h.appendChild(el("h2", null, d.file));
  h.appendChild(el("div", "meta", "locating tables and checking arithmetic…"));
  stage.appendChild(h);
}

function isNumeric(v) {
  return typeof v === "number";
}

function renderStage(doc_id) {
  const ex = state.extractions[doc_id];
  const stage = $("#stage");
  stage.innerHTML = "";

  const h = el("div", "dochead");
  h.appendChild(el("h2", null, ex.file));
  h.appendChild(el("div", "meta",
    `${ex.company} · report year ${ex.year || "?"} · ${ex.pages} pages`));
  stage.appendChild(h);

  ex.targets.forEach((t, i) => stage.appendChild(renderCard(doc_id, t, i)));
}

function renderCard(doc_id, t, idx) {
  const card = el("div", "card");

  const head = el("div", "card-head");
  head.appendChild(el("h3", null, t.name));
  if (t.found) {
    head.appendChild(el("div", "sub",
      `page ${t.page_label} · content match ${(t.score * 100).toFixed(0)}%` +
      (t.heading ? ` · report calls it “${t.heading}”` : "")));
    head.appendChild(renderBanner(t));
  }
  card.appendChild(head);

  if (!t.found) {
    card.appendChild(el("div", "notfound", t.message));
    return card;
  }

  const body = el("div", "card-body");
  body.appendChild(renderPdfPane(doc_id, t, idx));
  body.appendChild(renderTablePane(doc_id, t, idx));
  card.appendChild(body);

  const foot = el("div", "card-foot");
  foot.appendChild(el("span", null,
    `${t.body.length} rows · ${t.mode === "table_until_row" ? "statement" : "note"}`));
  foot.appendChild(el("span", "spacer"));
  const recheck = el("button", "btn small", "Re-check arithmetic");
  recheck.onclick = () => doRecheck(doc_id, idx);
  foot.appendChild(recheck);
  card.appendChild(foot);

  return card;
}

function renderBanner(t) {
  const cls = t.status === "PASS" ? "pass" : t.status === "FAIL" ? "fail" : "partial";
  const b = el("div", "banner " + cls);
  b.dataset.role = "banner";
  const tag = el("span", "b-tag",
    t.status === "PASS" ? "PASS" : t.status === "FAIL" ? "FAIL" : "CHECK");
  b.appendChild(tag);
  b.appendChild(el("span", null, t.summary.replace(/^(PASS|FAIL|NOT CHECKED)\s*/, "")));
  return b;
}

function renderPdfPane(doc_id, t, idx) {
  const key = `${doc_id}:${idx}`;
  if (state.pageByCard[key] == null) state.pageByCard[key] = t.page;
  const pane = el("div", "pdf-pane");

  const nav = el("div", "pdf-nav");
  const prev = el("button", "btn small", "‹");
  const next = el("button", "btn small", "›");
  const label = el("span", null, "");
  const spacer = el("span", "spacer");
  const scroll = el("div", "pdf-scroll");
  const img = el("img");
  scroll.appendChild(img);

  const load = () => {
    const p = state.pageByCard[key];
    label.textContent = `PDF page ${p} / ${t.total_pages}`;
    img.src = `/api/page?doc_id=${doc_id}&page=${p}&_=${Date.now()}`;
  };
  prev.onclick = () => { state.pageByCard[key] = Math.max(1, state.pageByCard[key] - 1); load(); };
  next.onclick = () => { state.pageByCard[key] = Math.min(t.total_pages, state.pageByCard[key] + 1); load(); };

  nav.append(prev, next, label, spacer);
  pane.append(nav, scroll);
  load();
  return pane;
}

function renderTablePane(doc_id, t, idx) {
  const pane = el("div", "tbl-pane");
  const table = el("table", "extract");
  table.dataset.card = idx;

  const ncols = Math.max(
    ...t.header.map((r) => r.length),
    ...t.body.map((r) => r.length), 1);

  // header
  const thead = el("thead");
  t.header.forEach((hr) => {
    const tr = el("tr");
    for (let c = 0; c < ncols; c++) {
      tr.appendChild(el("th", null, hr[c] != null ? String(hr[c]) : ""));
    }
    thead.appendChild(tr);
  });
  table.appendChild(thead);

  // body
  const tbody = el("tbody");
  t.body.forEach((row, r) => {
    const tr = el("tr");
    const label = (row.find((v) => typeof v === "string" && v.trim()) || "").toLowerCase();
    if (/\btotal\b|profit for the year|gross profit|operating profit/.test(label)) tr.classList.add("subtotal");
    else if (row.slice(1).every((v) => v == null || v === "")) tr.classList.add("section");

    for (let c = 0; c < ncols; c++) {
      const v = row[c];
      const td = el("td");
      td.contentEditable = "true";
      td.dataset.r = r; td.dataset.c = c;
      if (c === 0) {
        td.textContent = v != null ? String(v) : "";
      } else if (isNumeric(v)) {
        td.className = "num";
        td.textContent = v.toLocaleString("en-US");
      } else {
        // note-ref column or blank
        td.className = c === 1 && !isNumeric(v) ? "note" : "num";
        td.textContent = v != null ? String(v) : "";
      }
      td.addEventListener("input", () => onCellEdit(doc_id, idx, td));
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
  pane.appendChild(table);
  return pane;
}

/* ---------------- editing ---------------- */
let editTimer = null;
function onCellEdit(doc_id, idx, td) {
  td.classList.add("edited");
  const t = state.extractions[doc_id].targets[idx];
  const r = +td.dataset.r, c = +td.dataset.c;
  const raw = td.textContent.trim();
  // store the raw string; the server's clean_cell() handles "(1,234)",
  // ")1,234(", "-", "Nil" etc. consistently with the rest of the engine
  t.body[r][c] = raw === "" ? null : raw;
  clearTimeout(editTimer);
  editTimer = setTimeout(() => doRecheck(doc_id, idx), 700);
}

async function doRecheck(doc_id, idx) {
  const t = state.extractions[doc_id].targets[idx];
  try {
    const res = await api("/api/recheck", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ reconcile_kind: t.reconcile_kind, body: t.body }),
    });
    t.status = res.status;
    t.summary = res.summary;
    t.reconcile = res.reconcile;
    // update just the banner + dot
    const card = $$(".card")[idx];
    const oldB = card && card.querySelector('[data-role="banner"]');
    if (oldB) oldB.replaceWith(renderBanner(t));
    renderDocList();
    refreshExportBtn();
  } catch (e) {
    toast(`re-check failed: ${e.message}`, true);
  }
}
const $$ = (s, r = document) => [...r.querySelectorAll(s)];

/* ---------------- export ---------------- */
function refreshExportBtn() {
  const any = Object.keys(state.extractions).length > 0;
  $("#exportBtn").disabled = !any;
}

async function doExport() {
  const documents = state.docs
    .filter((d) => state.extractions[d.doc_id])
    .map((d) => {
      const ex = state.extractions[d.doc_id];
      return {
        file: ex.file, stem: ex.stem, company: ex.company,
        targets: ex.targets.map((t) => ({
          name: t.name, found: t.found, message: t.message,
          mode: t.mode, reconcile_kind: t.reconcile_kind,
          page_label: t.page_label, score: t.score, heading: t.heading,
          header: t.header, body: t.body,
        })),
      };
    });
  if (!documents.length) return;
  const companies = [...new Set(documents.map((d) => d.company))];
  const fname = (companies.length === 1
    ? companies[0].replace(/[^A-Za-z0-9]+/g, "_")
    : "financial") + "_tables.xlsx";
  try {
    const r = await api("/api/export", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ documents, filename: fname }),
    });
    const blob = await r.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = fname; a.click();
    URL.revokeObjectURL(url);
    toast(`exported ${documents.length} sheet${documents.length > 1 ? "s" : ""}`);
  } catch (e) {
    toast(`export failed: ${e.message}`, true);
  }
}

/* ---------------- wiring ---------------- */
function init() {
  const drop = $("#drop"), input = $("#fileInput");
  $("#pickBtn").onclick = () => input.click();
  input.onchange = () => { uploadFiles([...input.files]); input.value = ""; };

  ["dragenter", "dragover"].forEach((ev) =>
    drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add("over"); }));
  ["dragleave", "drop"].forEach((ev) =>
    drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.remove("over"); }));
  drop.addEventListener("drop", (e) => uploadFiles([...e.dataTransfer.files]));
  // allow dropping anywhere on the stage too
  document.addEventListener("dragover", (e) => e.preventDefault());
  document.addEventListener("drop", (e) => {
    if (e.target.closest(".rail")) return;
    e.preventDefault();
    if (e.dataTransfer.files.length) uploadFiles([...e.dataTransfer.files]);
  });

  $("#exportBtn").onclick = doExport;
  loadExistingDocs();
}

async function loadExistingDocs() {
  try {
    const j = await api("/api/docs");
    if (!j.docs || !j.docs.length) return;
    state.docs = j.docs.map((d) => ({ ...d, status: "idle" }));
    renderDocList();
  } catch (e) { /* first run, nothing there */ }
}

document.addEventListener("DOMContentLoaded", init);
