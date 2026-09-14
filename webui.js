const $ = s => document.querySelector(s);
// side-by-side is the DEFAULT for every table (not opt-in per table) --
// only an explicit prior toggle (stored either way) overrides that default
let _savedSideBySide = true;
try{
  const saved = localStorage.getItem("tk_side_by_side");
  if (saved !== null) _savedSideBySide = saved === "1";
}catch(e){}
const state = {
  files:[], file:null, tables:[], order:[], sel:new Set(),
  active:null, zoom:false, edits:{}, reTimers:{}, cmp:null,
  page:1, pageCount:null, box:null, searchQ:"", searchHits:null, quickfind:null,
  sideBySide:_savedSideBySide, pickerZoom:false, undoStack:{},
};

// -------------------------------------------------- per-file UI state -----
// Selection / export order / the workbook name are cheap-to-redo browser
// preferences, not authoritative data (unlike the extracted tables
// themselves, which serve.py now autosaves server-side -- see
// _save_session in serve.py) -- so these live in localStorage, mirrored on
// every change and silently restored on load. Per-viewer convenience only:
// wrapped in try/catch since a private window or blocked site data can make
// localStorage throw, and the app must still work fine without it.
function uiStateKey(file){ return "tk_ui_" + file; }
function saveUiState(){
  if (!state.file) return;
  try{
    localStorage.setItem(uiStateKey(state.file), JSON.stringify({
      sel: [...state.sel], order: state.order,
      fileName: $("#fileName") ? $("#fileName").value : "",
    }));
  }catch(e){}
}
function loadUiState(file){
  try{
    const raw = localStorage.getItem(uiStateKey(file));
    return raw ? JSON.parse(raw) : null;
  }catch(e){ return null; }
}
// kind identifiers stay fixed regardless of language (they're matched
// against the server's own `kind` field); only the DISPLAYED name -- via
// kindName() below, which looks the current translation up through t() --
// changes with the language toggle. KIND_NAME itself is only the
// English fallback t() uses when a key is missing from the active dict.
const KIND_I18N_KEY = {
  "income statement":"kind_income",
  "statement of financial position":"kind_balance",
  "statement of cash flows":"kind_cashflow",
  "statement of changes in equity":"kind_equity",
  "note":"kind_note","table":"kind_table",
};
const KIND_NAME = {
  "income statement":"Income statement",
  "statement of financial position":"Balance sheet",
  "statement of cash flows":"Cash flow statement",
  "statement of changes in equity":"Statement of changes in equity",
  "note":"Note","table":"Table",
};
function kindName(kind){ return t(KIND_I18N_KEY[kind]) || KIND_NAME[kind] || kind; }
const REAL_STMT = new Set(Object.keys(KIND_I18N_KEY).filter(k=>k!=="note"&&k!=="table"));
// Render resolution: base DPI-equivalent * the screen's actual pixel
// density, so a retina/high-DPI display doesn't get a 1x-sharp image
// stretched to look soft -- capped to match the server's own ceiling
// (serve.py: 4.5). Box-draw math divides by this same constant, so it
// stays correct regardless of what it resolves to.
const DPR = Math.min(2, window.devicePixelRatio || 1);
const PICKER_SCALE = Math.min(4.5, 2.2 * DPR);
// `f` is the matching entry from the response's "fmt" side-channel (same
// shape as "rows") -- {p: prefix, s: suffix} the source text carried (a
// currency symbol, a '%') that the VALUE itself can't keep (it has to stay
// a plain number for Δ / Δ% and every arithmetic check), so the display
// text re-adds it without touching the underlying value at all.
const fmt = (v, f) => (typeof v === "number")
  ? (f ? (f.p||"") : "") + v.toLocaleString(undefined,{maximumFractionDigits:2}) + (f ? (f.s||"") : "")
  : (v ?? "");
const nameOf = tbl => REAL_STMT.has(tbl.kind)
  ? kindName(tbl.kind) + (tbl.years && tbl.years[0] ? ` ${tbl.years[0]}` : "")
  : (tbl.title||"").slice(0,46);

function footPill(f, byCol){
  // a per-column verdict is stricter than the summary flag: if ANY checkable
  // column fails, say so, even when another column happens to reconcile
  if (byCol && byCol.length){
    const oks = byCol.map(c=>c.ok);
    if (oks.some(o=>o===false) && oks.some(o=>o===true))
      return `<span class="pill bad">✗ ${t('footsOneYearBad')}</span>`;
    if (oks.every(o=>o===false)) return `<span class="pill bad">✗ ${t('footsNo')}</span>`;
    if (oks.some(o=>o===true))   return `<span class="pill ok">✓ ${t('footsYes')}</span>`;
  }
  if (f === true)  return `<span class="pill ok">✓ ${t('footsYes')}</span>`;
  if (f === false) return `<span class="pill bad">✗ ${t('footsNo')}</span>`;
  return `<span class="pill none">– ${t('footsNoCheck')}</span>`;
}
function healthBar(h){
  if (h == null) return "";
  const pc = Math.round(h*100);
  const col = h>=.9 ? "var(--ok)" : h>=.75 ? "#65a30d" : "var(--warn)";
  return `<span class="hbar" title="${t('dataHealthTitle',{pc})}"><i style="width:${pc}%;background:${col}"></i></span>`;
}
async function jget(u){ const r = await fetch(u); const j = await r.json();
  if (j && j.error) throw new Error(j.error); return j; }
async function jpost(u,b){ const r = await fetch(u,{method:"POST",
  headers:{"Content-Type":"application/json"},body:JSON.stringify(b)});
  const j = await r.json();
  if(!r.ok||j.error){ const err = new Error(j.error||r.statusText); err.body = j; throw err; }
  return j; }

// ------------------------------------------------------ modal / toasts ----
// In-app replacements for native alert()/confirm() -- those are jarring,
// look nothing like the rest of the UI, and (confirm) block the whole page.
function showModal({title, body, okText="Remove", danger=true}){
  return new Promise(resolve=>{
    const host = document.createElement("div");
    host.className = "modal-backdrop";
    host.innerHTML = `
      <div class="modal-box" role="dialog" aria-modal="true" aria-labelledby="modalTitle">
        <div class="modal-title" id="modalTitle">${title}</div>
        <div class="modal-body">${body}</div>
        <div class="modal-actions">
          <button class="btn" data-a="cancel">${t('cancelBtn')}</button>
          <button class="btn-primary ${danger?"btn-danger":""}" data-a="ok">${okText}</button>
        </div>
      </div>`;
    document.body.appendChild(host);
    const done = v => { host.remove(); document.removeEventListener("keydown", onKey); resolve(v); };
    const cancelBtn = host.querySelector('[data-a="cancel"]');
    const okBtn = host.querySelector('[data-a="ok"]');
    cancelBtn.onclick = () => done(false);
    okBtn.onclick = () => done(true);
    host.addEventListener("mousedown", e => { if (e.target === host) done(false); });
    // only two focusable elements ever exist in here -- a plain wrap-around
    // is enough to keep Tab from leaving the dialog into the page behind it
    const onKey = e => {
      if (e.key === "Escape") return done(false);
      if (e.key !== "Tab") return;
      e.preventDefault();
      (document.activeElement === okBtn ? cancelBtn : okBtn).focus();
    };
    document.addEventListener("keydown", onKey);
    okBtn.focus();
  });
}
function confirmDialog(msg, okText){
  return showModal({title:t('areYouSure'), body:msg, okText:okText||t('removeBtn')});
}
function showErrorToast(msg){
  const host = $("#globalToasts");
  const el = document.createElement("div");
  el.className = "extract-toast toast-error";
  el.innerHTML = `⚠<span class="spacer"></span><button class="close" data-act="dismiss" title="${t('dismissTitle')}" aria-label="${t('dismissTitle')}">×</button>`;
  el.insertBefore(document.createTextNode(msg), el.querySelector(".spacer"));
  host.prepend(el);
  el.querySelector('[data-act="dismiss"]').onclick = () => el.remove();
  setTimeout(()=>{ if (el.isConnected) el.remove(); }, 9000);
}
// a plain informational toast (no action) -- e.g. "restored N tables"
function showInfoToast(msg, ms=6000){
  const host = $("#globalToasts");
  const el = document.createElement("div");
  el.className = "extract-toast";
  el.innerHTML = `${msg}<span class="spacer"></span><button class="close" data-act="dismiss" title="${t('dismissTitle')}" aria-label="${t('dismissTitle')}">×</button>`;
  host.prepend(el);
  el.querySelector('[data-act="dismiss"]').onclick = () => el.remove();
  setTimeout(()=>{ if (el.isConnected) el.remove(); }, ms);
}
// a toast with one action button (e.g. "Table removed. [Undo]") -- the
// action always dismisses the toast whether it throws or not, so a failed
// undo doesn't leave a dead button sitting on screen
function showActionToast(msg, actionLabel, onAction, ms=8000){
  const host = $("#globalToasts");
  const el = document.createElement("div");
  el.className = "extract-toast";
  el.innerHTML = `${msg}<span class="spacer"></span>
    <button class="btn" data-act="go">${actionLabel}</button>
    <button class="close" data-act="dismiss" title="${t('dismissTitle')}" aria-label="${t('dismissTitle')}">×</button>`;
  host.prepend(el);
  el.querySelector('[data-act="dismiss"]').onclick = () => el.remove();
  el.querySelector('[data-act="go"]').onclick = async () => {
    el.remove();
    try{ await onAction(); }catch(e){ showErrorToast(e.message); }
  };
  setTimeout(()=>{ if (el.isConnected) el.remove(); }, ms);
}
// the OCR failsafe: only ever offered (never run automatically) when a box
// came back with literally no text at all -- see doExtractRegion's catch.
function showOcrOfferToast(bbox){
  const host = $("#globalToasts");
  const el = document.createElement("div");
  el.className = "extract-toast toast-error";
  el.innerHTML = `${t('noTextInBox')}
    <span class="spacer"></span>
    <button class="btn btn-accent" data-act="ocr">${t('ocrThisRegion')}</button>
    <button class="close" data-act="dismiss" title="${t('dismissTitle')}" aria-label="${t('dismissTitle')}">×</button>`;
  host.prepend(el);
  el.querySelector('[data-act="dismiss"]').onclick = () => el.remove();
  el.querySelector('[data-act="ocr"]').onclick = () => { el.remove(); doExtractRegionOcr(bbox); };
  setTimeout(()=>{ if (el.isConnected) el.remove(); }, 15000);
}

// -------------------------------------------------------------- theme -----
function _applyTheme(mode){
  document.documentElement.dataset.theme = mode;
  const icon = $("#themeIcon");
  if (icon) icon.innerHTML = mode === "dark"
    ? '<path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"/>'
    : '<circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.93 4.93l1.41 1.41M17.66 17.66l1.41 1.41M2 12h2M20 12h2M6.34 17.66l-1.41 1.41M19.07 4.93l-1.41 1.41"/>';
}
(function initTheme(){
  let saved = null;
  try{ saved = localStorage.getItem("tk_theme"); }catch(e){}
  const mode = saved || (window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light");
  _applyTheme(mode);
})();
document.addEventListener("DOMContentLoaded", ()=>{
  const btn = $("#themeToggle");
  if (btn) btn.onclick = ()=>{
    const next = document.documentElement.dataset.theme === "dark" ? "light" : "dark";
    _applyTheme(next);
    try{ localStorage.setItem("tk_theme", next); }catch(e){}
  };
});

// --------------------------------------------------------------- i18n -----
// Static UI chrome -- every button, label, tooltip and message this file
// itself renders. NOT translated: the extracted TABLE CONTENT (it reflects
// whatever language the source PDF was written in, not app UI), and
// messages the SERVER generates dynamically (a table's `notes`, the
// cross-year consistency verdict, the compare panel's verdict text) --
// those come back as already-formatted English sentences from
// extract_all_tables.py, and localizing them would mean teaching the
// Python analysis engine to emit i18n keys instead of prose, a separate,
// much larger change than this pass.
const I18N = {
  en: {
    brand: "Tables", pageTitle: "Table preview", export: "Export", upload: "Upload PDF", uploading: "Uploading…",
    fileName_ph: "workbook name",
    sideBySide: "⬌ Side by side", stackPanes: "⬍ Stack panes",
    layoutBtnTitle: "switch between side-by-side and stacked panes (applies whenever you're viewing a table)",
    themeToggleTitle: "Switch light/dark mode",
    outputFileTitle: "output file name",
    howItWorks: "How it works",
    step1_h: "Pick a page", step1_b: "Browse the PDF and find the table you want.",
    step2_h: "Draw a box", step2_b: "Drag a rectangle around just the rows and columns you need.",
    step3_h: "Check the verdict", step3_b: "Every extraction is checked against its own printed totals — a foots pill tells you if it adds up.",
    dzTitle: "Drop a PDF here, or click Upload PDF above",
    dzBody: "Everything runs locally — nothing leaves this machine.",
    emptyStart: "Upload a PDF (top left) to browse its pages and draw a box around the table you want.",

    selectTable: "+ Select a table", selectTableTitle: "go back to the page picker to draw another box",
    selectAllStatements: "Select all statements", clearBtn: "Clear",
    uploadPrompt: "Upload a PDF to get started.",
    noTablesYet: "no tables yet", nTables: "{n} tables",
    noTablesYetHint: "No tables yet — draw a box on the right to extract one.",
    detailEmpty: "Extract a table (drag a box, then click Extract) or pick one from the list on the left to preview it here.",
    borderlessFull: "Borderless tables: full support",
    borderlessFullTitle: "Full support: works on borderless / hand-drawn-box tables, not just ruled ones",
    borderlessLimited: "Borderless tables: limited",
    borderlessLimitedTitle: "pip install img2table opencv-python-headless pandas -- until then, only tables with visible ruling lines extract reliably",

    groupManual: "Manual selections", groupStatements: "Financial statements",
    groupNotes: "Notes", groupOther: "Other tables",
    editedTag: "edited", selectRowCb: "select {name}",
    ocrPill: "OCR", ocrPillTitle: "read by OCR, not the PDF's own text -- double-check the figures",
    suspectPillTitle: "{n} row(s) look wrong",
    consistencyPillTitle: "prior-year column disagrees with {vs}",
    nothingToShow: "Nothing to show.",
    removeTableTitle: "remove this extracted table",
    confirmRemoveTable: "Remove this extracted table? You can undo this right after.",
    removeBtn: "Remove", cancelBtn: "Cancel", areYouSure: "Are you sure?",
    removeFailed: "Couldn't remove that table: {msg}",
    tableRemoved: "Table removed.", undoBtn: "Undo", undoFailed: "Couldn't undo: {msg}",
    nothingToUndo: "Nothing left to undo.",
    restoredTablesToast: "Restored {n} table(s) from your last session.",
    undoEditTitle: "undo the last edit to this table", undoEditBtn: "↺ Undo edit",

    exportOrder: "Export order", exportOrderN: "Export order — {n}",
    tickToAdd: "tick tables to add them",
    exportBtnN: "Export {n} →",
    moveUpTitle: "move up in the export order", moveDownTitle: "move down in the export order",
    removeFromOrderTitle: "remove from the export order",
    queueTableCount: "{n} table(s)",

    selectATableH2: "Select a table", pageWord: "page", ofTotal: "of {total}",
    prevBtn: "‹ Prev", nextBtn: "Next ›",
    searchPh: "Search this PDF's text… (press Enter)", searchBtn: "Search",
    dragBoxHint: "Drag a box around the table you want (rows and columns only), then click Extract. You can extract more than one table from the same page.",
    pageCap: "Page {page}", fitWidth: "Fit width", actualSize: "Actual size",
    extractBtn: "Extract this region →", extracting: "Extracting…", runningOcr: "Running OCR…",
    locatingStatements: "locating the financial statements…", jumpTo: "Jump to:",
    typeAtLeast2: "type at least 2 characters", searching: "searching…",
    searchFailed: "search failed: {msg}", noMatches: "no matches",
    pageCountFailed: "Couldn't get this PDF's page count: {msg} — try again",
    // per-note translations for the fixed-wording subset of extract_all_tables.py's
    // analyze() notes (see _add_note there) -- NOT every note has a key here (most
    // are fully dynamic prose, e.g. reconcile_explain's worked arithmetic), so
    // renderPreview falls back to the raw English text when a note's key is
    // missing or unrecognised, same as an older server response with no notes_i18n at all
    noteYearsReversed: "year columns may be reversed — the page header lists years oldest-first ({first}…{last}); figures could be attributed to the wrong year. Use 'Swap year columns' if so.",
    noteSegmentalColumns: "columns look like segments / entities, not reporting years — the figures across columns may not be a comparable time series.",
    noteAssetsOnlyIncomplete: "equity / liabilities side incomplete — only the assets side was captured (check the following PDF page for the rest)",
    dragToResizeTitle: "drag to resize",
    selectFileTitle: "selected file", pageNumTitle: "page number", pgImgAlt: "PDF page preview",

    extractionFailed: "Extraction failed: {msg}", ocrFailed: "OCR failed: {msg}",
    noTextInBox: "⚠ No text in that box (looks like a scan).",
    ocrThisRegion: "OCR this region →", dismissTitle: "dismiss",
    extractedPrefix: "✓ Extracted", viewBtn: "View →",

    insertRowBelow: "insert row below", deleteRow: "delete row",
    splitLabel: "split label at the middle", mergeUp: "merge into row above",
    pageLabel: "page {page}", labelsFigures: "labels {l}% · figures {f}%",
    ocrVerifyByEye: "OCR — verify by eye", editedRecomputed: "edited — verdict recomputed",
    sheetNamePh: "sheet name (optional)", swapYears: "Swap year columns",
    undoEdits: "Undo my edits",
    editHint: "click a cell to fix it · hover a row for ＋ ✕ ⤶ ⭡ · exports use your version",
    savingStatus: "Saving…", savedStatus: "Saved", saveErrorStatus: "Couldn't save — will retry when you edit again, or export still uses your version",
    extractedTableRows: "Extracted table — {n} rows", amberRedReview: "amber/red = review",
    priorYearVs: "prior-year ({year}) column vs <b>{vs}</b>:", agree: "agree ✓",
    differBase: "{mismatch} of {checked} differ —", differsWord: "differs",

    footsYes: "foots", footsNo: "no foot", footsOneYearBad: "one year doesn't foot",
    footsNoCheck: "no check", dataHealthTitle: "data health {pc}%",
    kind_income: "Income statement", kind_balance: "Balance sheet",
    kind_cashflow: "Cash flow statement", kind_equity: "Statement of changes in equity",
    kind_note: "Note", kind_table: "Table",

    compareTitle: "Compare", compareVs: "this table vs the same statement in",
    compareBtn: "Compare →", scanningMatching: "scanning {file} & matching…",
    noKindFoundIn: "no {kind} found in {file}", compareFailed: "compare failed: {msg}",
    cmpVerdictRestated: "{changed} changed, {new} new, {removed} removed, {restated} RESTATED",
    cmpVerdictClean: "{changed} changed, {new} new, {removed} removed, prior-year columns agree",

    uploadFailed: "Upload failed: {msg}",
    riskyExportWarn: "{risky} of the {total} selected table(s) have a warning (NO FOOT or low data-health):",
    exportAnywayQ: "Export anyway?", exportAnyway: "Export anyway",
    building: "Building…", exportFailed: "Export failed: {msg}",
  },
  ar: {
    brand: "الجداول", pageTitle: "معاينة الجداول", export: "تصدير", upload: "رفع PDF", uploading: "جارٍ الرفع…",
    fileName_ph: "اسم ملف العمل",
    sideBySide: "⬌ جنبًا إلى جنب", stackPanes: "⬍ تكديس اللوحات",
    layoutBtnTitle: "التبديل بين اللوحات جنبًا إلى جنب والمكدّسة (يسري متى ما كنت تعرض جدولًا)",
    themeToggleTitle: "تبديل الوضع الفاتح/الداكن",
    outputFileTitle: "اسم ملف الإخراج",
    howItWorks: "كيف يعمل",
    step1_h: "اختر صفحة", step1_b: "تصفح PDF واعثر على الجدول الذي تريده.",
    step2_h: "ارسم مربعًا", step2_b: "اسحب مستطيلًا حول الصفوف والأعمدة التي تحتاجها فقط.",
    step3_h: "تحقق من النتيجة", step3_b: "يتم فحص كل استخراج مقابل إجمالياته المطبوعة — شارة المطابقة تخبرك إن كانت الأرقام صحيحة.",
    dzTitle: "أفلت ملف PDF هنا، أو اضغط \"تحميل PDF\" أعلاه",
    dzBody: "كل شيء يعمل محليًا — لا تغادر أي بيانات هذا الجهاز.",
    emptyStart: "ارفع ملف PDF (أعلى اليسار) لتصفح صفحاته ورسم مربع حول الجدول الذي تريده.",

    selectTable: "+ اختر جدولًا", selectTableTitle: "العودة إلى متصفح الصفحات لرسم مربع آخر",
    selectAllStatements: "اختر كل القوائم", clearBtn: "مسح",
    uploadPrompt: "ارفع ملف PDF للبدء.",
    noTablesYet: "لا توجد جداول بعد", nTables: "{n} جدول",
    noTablesYetHint: "لا توجد جداول بعد — ارسم مربعًا على اليمين لاستخراج واحد.",
    detailEmpty: "استخرج جدولًا (ارسم مربعًا ثم اضغط استخراج) أو اختر واحدًا من القائمة على اليسار لمعاينته هنا.",
    borderlessFull: "الجداول غير المحددة بحدود: دعم كامل",
    borderlessFullTitle: "دعم كامل: يعمل مع الجداول غير المحددة بحدود / المربعات المرسومة يدويًا، وليس فقط ذات الخطوط",
    borderlessLimited: "الجداول غير المحددة بحدود: دعم محدود",
    borderlessLimitedTitle: "ثبّت pip install img2table opencv-python-headless pandas -- إلى أن يتم ذلك، يمكن استخراج الجداول ذات الخطوط المرئية فقط بشكل موثوق",

    groupManual: "التحديدات اليدوية", groupStatements: "القوائم المالية",
    groupNotes: "الإيضاحات", groupOther: "جداول أخرى",
    editedTag: "معدَّل", selectRowCb: "تحديد {name}",
    ocrPill: "OCR", ocrPillTitle: "تمت قراءته بالتعرف الضوئي وليس من نص PDF الأصلي -- تحقق من الأرقام",
    suspectPillTitle: "{n} صفًا يبدو غير صحيح",
    consistencyPillTitle: "عمود السنة السابقة لا يطابق {vs}",
    nothingToShow: "لا يوجد ما يُعرض.",
    removeTableTitle: "إزالة هذا الجدول المستخرج",
    confirmRemoveTable: "إزالة هذا الجدول المستخرج؟ يمكنك التراجع عن هذا مباشرة بعدها.",
    removeBtn: "إزالة", cancelBtn: "إلغاء", areYouSure: "هل أنت متأكد؟",
    removeFailed: "تعذّرت إزالة الجدول: {msg}",
    tableRemoved: "تمت إزالة الجدول.", undoBtn: "تراجع", undoFailed: "تعذّر التراجع: {msg}",
    nothingToUndo: "لا يوجد ما يمكن التراجع عنه.",
    restoredTablesToast: "تم استرجاع {n} جدول(جداول) من جلستك الأخيرة.",
    undoEditTitle: "التراجع عن آخر تعديل لهذا الجدول", undoEditBtn: "↺ تراجع عن التعديل",

    exportOrder: "ترتيب التصدير", exportOrderN: "ترتيب التصدير — {n}",
    tickToAdd: "حدد الجداول لإضافتها",
    exportBtnN: "تصدير {n} →",
    moveUpTitle: "نقل لأعلى في ترتيب التصدير", moveDownTitle: "نقل لأسفل في ترتيب التصدير",
    removeFromOrderTitle: "إزالة من ترتيب التصدير",
    queueTableCount: "{n} جدول",

    selectATableH2: "اختر جدولًا", pageWord: "صفحة", ofTotal: "من {total}",
    prevBtn: "‹ السابق", nextBtn: "التالي ›",
    searchPh: "ابحث في نص PDF… (اضغط Enter)", searchBtn: "بحث",
    dragBoxHint: "ارسم مربعًا حول الجدول الذي تريده (الصفوف والأعمدة فقط)، ثم اضغط استخراج. يمكنك استخراج أكثر من جدول من نفس الصفحة.",
    pageCap: "صفحة {page}", fitWidth: "احتواء العرض", actualSize: "الحجم الفعلي",
    extractBtn: "استخراج هذه المنطقة →", extracting: "جارٍ الاستخراج…", runningOcr: "جارٍ التعرف الضوئي…",
    locatingStatements: "جارٍ تحديد موقع القوائم المالية…", jumpTo: "الانتقال إلى:",
    typeAtLeast2: "اكتب حرفين على الأقل", searching: "جارٍ البحث…",
    searchFailed: "فشل البحث: {msg}", noMatches: "لا توجد نتائج",
    pageCountFailed: "تعذّر الحصول على عدد صفحات هذا الملف: {msg} — حاول مجددًا",
    noteYearsReversed: "قد تكون أعمدة السنوات معكوسة — يسرد رأس الصفحة السنوات من الأقدم إلى الأحدث ({first}…{last})؛ قد تُنسب الأرقام إلى السنة الخطأ. استخدم \"تبديل أعمدة السنوات\" إذا كان الأمر كذلك.",
    noteSegmentalColumns: "تبدو الأعمدة كقطاعات / كيانات، وليست سنوات تقرير — قد لا تكون الأرقام عبر الأعمدة سلسلة زمنية قابلة للمقارنة.",
    noteAssetsOnlyIncomplete: "جانب حقوق الملكية / الالتزامات غير مكتمل — تم التقاط جانب الأصول فقط (تحقق من صفحة PDF التالية لمعرفة الباقي)",
    dragToResizeTitle: "اسحب لتغيير الحجم",
    selectFileTitle: "الملف المحدد", pageNumTitle: "رقم الصفحة", pgImgAlt: "معاينة صفحة PDF",

    extractionFailed: "فشل الاستخراج: {msg}", ocrFailed: "فشل التعرف الضوئي: {msg}",
    noTextInBox: "⚠ لا يوجد نص في هذا المربع (يبدو أنه مسح ضوئي).",
    ocrThisRegion: "تعرّف ضوئيًا على هذه المنطقة →", dismissTitle: "إغلاق",
    extractedPrefix: "✓ تم استخراج", viewBtn: "عرض →",

    insertRowBelow: "إدراج صف أسفل", deleteRow: "حذف الصف",
    splitLabel: "تقسيم التسمية من المنتصف", mergeUp: "دمج مع الصف الأعلى",
    pageLabel: "صفحة {page}", labelsFigures: "التسميات {l}% · الأرقام {f}%",
    ocrVerifyByEye: "تعرّف ضوئي — تحقق بعينك", editedRecomputed: "معدَّل — أُعيد حساب النتيجة",
    sheetNamePh: "اسم الورقة (اختياري)", swapYears: "تبديل أعمدة السنوات",
    undoEdits: "التراجع عن تعديلاتي",
    editHint: "اضغط على خلية لتصحيحها · مرّر فوق صف لـ ＋ ✕ ⤶ ⭡ · التصدير يستخدم نسختك",
    savingStatus: "جارٍ الحفظ…", savedStatus: "تم الحفظ", saveErrorStatus: "تعذّر الحفظ — ستتم إعادة المحاولة عند التعديل مجددًا، أو يستخدم التصدير نسختك دائمًا",
    extractedTableRows: "الجدول المستخرج — {n} صفًا", amberRedReview: "كهرماني/أحمر = يحتاج مراجعة",
    priorYearVs: "عمود السنة السابقة ({year}) مقابل <b>{vs}</b>:", agree: "متطابق ✓",
    differBase: "{mismatch} من أصل {checked} يختلف —", differsWord: "يختلف",

    footsYes: "متطابق", footsNo: "غير متطابق", footsOneYearBad: "سنة واحدة غير متطابقة",
    footsNoCheck: "بدون فحص", dataHealthTitle: "جودة البيانات {pc}%",
    kind_income: "قائمة الدخل", kind_balance: "قائمة المركز المالي",
    kind_cashflow: "قائمة التدفقات النقدية", kind_equity: "قائمة التغيرات في حقوق الملكية",
    kind_note: "إيضاح", kind_table: "جدول",

    compareTitle: "مقارنة", compareVs: "قارن هذا الجدول بنفس القائمة في",
    compareBtn: "قارن →", scanningMatching: "جارٍ فحص {file} والمطابقة…",
    noKindFoundIn: "لم يُعثر على {kind} في {file}", compareFailed: "فشلت المقارنة: {msg}",
    cmpVerdictRestated: "{changed} معدَّل، {new} جديد، {removed} محذوف، {restated} أُعيد إدراجه",
    cmpVerdictClean: "{changed} معدَّل، {new} جديد، {removed} محذوف، أعمدة السنة السابقة متطابقة",

    uploadFailed: "فشل الرفع: {msg}",
    riskyExportWarn: "{risky} من أصل {total} جدولًا محددًا يحمل تحذيرًا (لا مطابقة أو جودة بيانات منخفضة):",
    exportAnywayQ: "تصدير على أي حال؟", exportAnyway: "تصدير على أي حال",
    building: "جارٍ الإنشاء…", exportFailed: "فشل التصدير: {msg}",
  },
};
function t(key, vars){
  const dict = I18N[document.documentElement.lang] || I18N.en;
  let s = dict[key] != null ? dict[key] : (I18N.en[key] != null ? I18N.en[key] : key);
  if (vars) for (const k in vars) s = s.replace(new RegExp("\\{"+k+"\\}","g"), vars[k]);
  return s;
}
const ONB_ICONS = [
  '<path d="M4 3h10l4 4v14H4z"/><path d="M14 3v4h4"/>',
  '<rect x="4" y="6" width="16" height="12" rx="1" stroke-dasharray="3 2"/><path d="M4 6l16 12M20 6L4 18" opacity="0"/>',
  '<circle cx="12" cy="12" r="9"/><path d="M8 12l3 3 5-6"/>',
];
function renderOnboarding(){
  const host = $("#onboarding");
  if (!host) return;
  const dict = I18N[document.documentElement.lang] || I18N.en;
  const steps = [
    ["step1_h","step1_b"], ["step2_h","step2_b"], ["step3_h","step3_b"],
  ];
  host.innerHTML = `
    <h2 class="onb-title" data-i18n="howItWorks">${dict.howItWorks}</h2>
    <div class="onb-steps">
      ${steps.map(([h,b],i)=>`
        <div class="onb-step">
          <div class="onb-icon"><svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">${ONB_ICONS[i]}</svg></div>
          <div class="onb-h" data-i18n="${h}">${dict[h]}</div>
          <div class="onb-b" data-i18n="${b}">${dict[b]}</div>
        </div>`).join("")}
    </div>
    <div class="dropzone">
      <div class="dz-box" id="dzBox">
        <div class="dz-icon">
          <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 16V4M6 10l6-6 6 6"/><path d="M4 20h16"/></svg>
        </div>
        <div class="dz-h" data-i18n="dzTitle">${dict.dzTitle}</div>
        <div class="dz-b" data-i18n="dzBody">${dict.dzBody}</div>
      </div>
    </div>`;
  wireDropzone();
}
function wireDropzone(){
  const dz = $("#dzBox");
  if (!dz) return;
  const over = e => { e.preventDefault(); dz.classList.add("dragover"); };
  const leave = () => dz.classList.remove("dragover");
  dz.addEventListener("dragover", over);
  dz.addEventListener("dragleave", leave);
  dz.addEventListener("drop", async e => {
    e.preventDefault(); leave();
    const f = e.dataTransfer.files && e.dataTransfer.files[0];
    if (!f) return;
    if (!/\.pdf$/i.test(f.name)){ showErrorToast(t("uploadFailed",{msg:"not a PDF"})); return; }
    await uploadPdfFile(f);
  });
}
function applyI18n(lang){
  const dict = I18N[lang] || I18N.en;
  document.querySelectorAll("[data-i18n]").forEach(el=>{
    const k = el.dataset.i18n;
    if (dict[k] != null) el.textContent = dict[k];
  });
  document.querySelectorAll("[data-i18n-ph]").forEach(el=>{
    const k = el.dataset.i18nPh;
    if (dict[k] != null) el.setAttribute("placeholder", dict[k]);
  });
  document.querySelectorAll("[data-i18n-title]").forEach(el=>{
    const k = el.dataset.i18nTitle;
    if (dict[k] != null){
      el.setAttribute("title", dict[k]);
      // same text for screen readers -- title alone isn't reliably
      // announced, and several of these (themeToggle, the tray's ↑/↓/×
      // buttons) carry no visible text at all for aria-label to redundantly
      // repeat over
      el.setAttribute("aria-label", dict[k]);
    }
  });
  document.documentElement.lang = lang;
  document.documentElement.dir = lang === "ar" ? "rtl" : "ltr";
  document.body.classList.toggle("lang-ar", lang === "ar");
  const lt = $("#langToggle");
  if (lt) lt.textContent = lang === "ar" ? "ع / EN" : "EN / ع";
  document.title = t("pageTitle");
  renderOnboarding();
  // everything below is dynamically-rendered content that doesn't carry a
  // data-i18n attribute (it's built from `state` at render time) -- re-run
  // whichever of these is currently on screen so a language switch mid-
  // session updates it too, not just the static chrome and the landing page
  if (typeof setLayoutBtnLabel === "function" && $("#layoutBtn")) setLayoutBtnLabel();
  if (state.files && state.files.length && $("#list")){
    renderList([]);
    renderTray();
    syncExport();
  }
  renderEngineBadge();
  if ($("#pickerPane")) renderPicker();
  if (state.active != null && state._lastDetail) renderPreview(state.active, state._lastDetail);
}
(function initLang(){
  let saved = null;
  try{ saved = localStorage.getItem("tk_lang"); }catch(e){}
  applyI18n(saved || "en");
})();
document.addEventListener("DOMContentLoaded", ()=>{
  const btn = $("#langToggle");
  if (btn) btn.onclick = ()=>{
    const next = document.documentElement.lang === "ar" ? "en" : "ar";
    applyI18n(next);
    try{ localStorage.setItem("tk_lang", next); }catch(e){}
  };
});

// ---------------------------------------------------------------- boot ----
// controls that do nothing useful (or error) before a file is chosen --
// left fully clickable pre-fix, which used to crash with a raw exception
// when clicked on the empty landing screen
const FILE_GATED_IDS = ["pickerBtn","selAllStmt","selClear","layoutBtn","fileName"];
function setControlsEnabled(on){
  FILE_GATED_IDS.forEach(id=>{ const el=$("#"+id); if (el) el.disabled = !on; });
}

// the file-picker <select> doubles as a lightweight multi-file queue: each
// option's label carries that file's table count so you can see at a
// glance which reports still need work without switching to each one.
// /api/scan never triggers a scan itself (serve.py's inventory() docstring)
// -- it only reads back what's already there, so fetching it for every
// loaded file here is cheap, not a re-extraction.
// cheaper than refreshFileQueueLabels() -- no fetch, just reflects a count
// already in memory onto the currently-selected file's own <option>
function updateCurrentFileQueueLabel(){
  const sel = $("#fileSel");
  if (!sel || sel.hidden) return;
  const opt = [...sel.options].find(o=>o.value===state.file);
  if (opt) opt.textContent = state.tables.length
    ? `${state.file} — ${t('queueTableCount',{n:state.tables.length})}` : state.file;
}
async function refreshFileQueueLabels(){
  const sel = $("#fileSel");
  if (!sel || !state.files.length) return;
  const counts = await Promise.all(state.files.map(async name=>{
    try{ const inv = await jget("/api/scan?file="+encodeURIComponent(name)); return inv.tables.length; }
    catch(e){ return 0; }
  }));
  sel.innerHTML = state.files.map((name,i)=>
    `<option value="${name}"${name===state.file?" selected":""}>${name}${counts[i]?` — ${t('queueTableCount',{n:counts[i]})}`:""}</option>`
  ).join("");
}
async function boot(){
  setControlsEnabled(false);
  try{
    const st = await jget("/api/status");
    state._engineOk = st.img2table;
    renderEngineBadge();
  }catch(e){ /* non-critical */ }
  const f = await jget("/api/files");
  state.files = f.files;
  const sel = $("#fileSel");
  sel.innerHTML = f.files.map(n=>`<option>${n}</option>`).join("");
  sel.hidden = f.files.length <= 1;
  sel.onchange = () => loadFile(sel.value);
  if (f.files.length > 1) refreshFileQueueLabels();
  if (f.files.length) loadFile(f.files[0], {isBoot:true});
  else {
    $("#list").innerHTML = `<div class="msg">${t('uploadPrompt')}</div>`;
    $("#preview").innerHTML = `<div class="onboarding" id="onboarding"></div>`;
    renderOnboarding();
  }
}
function renderEngineBadge(){
  const b = $("#engineBadge");
  if (!b || state._engineOk == null) return;
  b.innerHTML = state._engineOk
    ? `<span class="pill ok" title="${t('borderlessFullTitle')}">${t('borderlessFull')}</span>`
    : `<span class="pill warn" title="${t('borderlessLimitedTitle')}">${t('borderlessLimited')}</span>`;
}

function ensureSplitLayout(){
  if ($("#splitRegion")) return;
  $("#preview").innerHTML = `
    <div class="split-region ${state.sideBySide?"side-by-side":""}" id="splitRegion">
      <div class="split-pane" id="pickerPane"></div>
      <div class="split-pane" id="detailPane">
        <div class="empty">${t('detailEmpty')}</div>
      </div>
    </div>`;
}

async function loadFile(name, opts={}){
  state.file = name; state.tables = []; state.order = []; state.sel.clear();
  state.active = null; state.edits = {}; state.cmp = null; state.undoStack = {};
  // any timer still in here belongs to the file we're LEAVING -- scheduleReanalyze
  // guards against it firing against the wrong file, so this is just tidying the
  // tracking object, not relying on it to cancel anything
  state.reTimers = {};
  state.page = 1; state.pageCount = null; state.box = null;
  state.searchQ = ""; state.searchHits = null; state.quickfind = null;
  setControlsEnabled(true);
  syncExport(); renderTray();
  $("#fileName").value = name.replace(/\.pdf$/i,"") + " — " + t('brand').toLowerCase();
  $("#list").innerHTML = `<div class="msg">${t('noTablesYetHint')}</div>`;
  ensureSplitLayout();
  $("#detailPane").innerHTML = `<div class="empty">${t('detailEmpty')}</div>`;
  try{
    const inv = await jget("/api/scan?file="+encodeURIComponent(name));
    state.tables = inv.tables; state.order = inv.tables.map(t=>t.n);
    // re-apply whatever this browser last had ticked/ordered/named for this
    // exact file -- the tables themselves are already back (serve.py's own
    // session autosave put them in `inv`); this is just the lighter,
    // per-viewer preferences layered back on top. Silent and automatic
    // rather than a confirm prompt: worst case the user re-ticks a box,
    // there's nothing here that isn't trivially redone by hand.
    const saved = loadUiState(name);
    const liveNs = new Set(state.order);
    if (saved){
      (saved.sel||[]).forEach(n=>{ if (liveNs.has(n)) state.sel.add(n); });
      if (saved.order && saved.order.length){
        const kept = saved.order.filter(n=>liveNs.has(n));
        const extra = state.order.filter(n=>!kept.includes(n));
        state.order = [...kept, ...extra];
      }
      if (saved.fileName) $("#fileName").value = saved.fileName;
      syncExport(); renderTray();
    }
    if (opts.isBoot && state.tables.length){
      showInfoToast(t("restoredTablesToast", {n: state.tables.length}));
    }
    renderList(inv.warnings||[]);
  }catch(e){ /* nothing extracted for this file yet -- fine */ }
  showPicker();
  // fire-and-forget: a fast text-only pass to point at roughly where the
  // core statements are, so a 100+ page report doesn't leave the user
  // guessing after landing on page 1 (usually a cover with nothing on it)
  jget("/api/quickfind?file="+encodeURIComponent(name)).then(res=>{
    state.quickfind = res.hits || [];
    if (state.file === name) renderQuickfind();
  }).catch(()=>{});
}

const GROUP_KEY = {
  "Manual selections":"groupManual", "Financial statements":"groupStatements",
  "Notes":"groupNotes", "Other tables":"groupOther",
};
function renderList(warnings){
  const tabs = state.tables;
  $("#count").textContent = tabs.length ? t("nTables",{n:tabs.length}) : t("noTablesYet");
  let html = "";
  if (warnings && warnings.length)
    html += `<div class="warnbox">${warnings.map(w=>w.replace(/^!!\s*/,"")).join("<br>")}</div>`;
  for (const g of ["Manual selections","Financial statements","Notes","Other tables"]){
    const items = tabs.filter(x=>x.group===g);
    if (!items.length) continue;
    html += `<div class="grouphd">${t(GROUP_KEY[g])}</div>`;
    for (const x of items){
      const nm = nameOf(x);
      const yrs = x.years.length ? x.years.join("/") : "–";
      const st = x.stitched_from ? ` +p${x.stitched_from}` : "";
      const consBad = x.consistency && x.consistency.mismatch
        ? `<span class="pill bad" title="${t('consistencyPillTitle',{vs:x.consistency.vs})}">${x.consistency.mismatch}≠</span>` : "";
      const canDelete = g === "Manual selections";
      html += `
        <div class="row" data-n="${x.n}">
          <input type="checkbox" data-n="${x.n}" ${state.sel.has(x.n)?"checked":""} aria-label="${t('selectRowCb',{name:nm}).replace(/"/g,"&quot;")}">
          <div class="body" data-n="${x.n}">
            <div class="name">${nm}${state.edits[x.n]?`<span class="tag">${t('editedTag')}</span>`:''}</div>
            <div class="sub">${t('pageLabel',{page:x.page})}${st} · ${yrs} · ${x.nrows}×${x.ncols}</div>
            <div class="meta">${footPill(x.foots)} ${healthBar(x.health)}
              ${x.ocr?`<span class="pill warn" title="${t('ocrPillTitle')}">${t('ocrPill')}</span>`:``}
              ${x.n_suspect?`<span class="pill warn" title="${t('suspectPillTitle',{n:x.n_suspect})}">${x.n_suspect}⚠</span>`:``}
              ${consBad}
              ${x.notes.length?`<span class="pill warn" title="${noteLines(x).join(' / ')}">!</span>`:``}
            </div>
          </div>
          ${canDelete?`<button class="row-del" data-del="${x.n}" title="${t('removeTableTitle')}">×</button>`:``}
        </div>`;
    }
  }
  $("#list").innerHTML = html || `<div class="msg">${t('nothingToShow')}</div>`;
  $("#list").querySelectorAll(".body").forEach(el=>el.onclick=()=>openTable(+el.dataset.n));
  $("#list").querySelectorAll('input[type=checkbox]').forEach(cb=>
    cb.onchange=()=>toggleSel(+cb.dataset.n, cb.checked));
  $("#list").querySelectorAll('[data-del]').forEach(b=>
    b.onclick=(e)=>{ e.stopPropagation(); deleteManualTable(+b.dataset.del); });
}

async function deleteManualTable(n){
  const ok = await confirmDialog(t("confirmRemoveTable"));
  if (!ok) return;
  const wasSelected = state.sel.has(n);
  const wasOrderIdx = state.order.indexOf(n);
  try{
    const inv = await jpost("/api/delete_manual", {file:state.file, n});
    state.tables = inv.tables; state.order = inv.tables.map(x=>x.n);
    state.sel.delete(n); delete state.edits[n];
    syncExport(); renderTray(); renderList(inv.warnings||[]);
    if (state.active === n){
      state.active = null;
      $("#detailPane").innerHTML = `<div class="empty">${t('detailEmpty')}</div>`;
    }
    updateCurrentFileQueueLabel();
    saveUiState();
    showActionToast(t("tableRemoved"), t("undoBtn"), () => undeleteTable(n, wasSelected, wasOrderIdx));
  }catch(e){ showErrorToast(t("removeFailed",{msg:e.message})); }
}

async function undeleteTable(n, wasSelected, wasOrderIdx){
  try{
    const inv = await jpost("/api/undelete_manual", {file:state.file});
    state.tables = inv.tables; state.order = inv.tables.map(x=>x.n);
    if (wasOrderIdx >= 0 && state.order.includes(n)){
      state.order = state.order.filter(x=>x!==n);
      state.order.splice(Math.min(wasOrderIdx, state.order.length), 0, n);
    }
    if (wasSelected) state.sel.add(n);
    syncExport(); renderTray(); renderList(inv.warnings||[]);
    updateCurrentFileQueueLabel();
    saveUiState();
  }catch(e){ showErrorToast(t("undoFailed",{msg:e.message})); }
}

function toggleSel(n,on){
  if (on){ state.sel.add(n); if(!state.order.includes(n)) state.order.push(n); }
  else state.sel.delete(n);
  syncExport(); renderTray(); saveUiState();
}
// extract_all_tables.py's analyze() -- see _add_note there -- sends each
// note in TWO forms: plain English (d.notes[i], unchanged, what the CLI and
// Excel export always used) and, for the fixed-wording subset, a {key,vars}
// pair in d.notes_i18n[i] at the SAME index. Only the keys listed here are
// ones this file actually has a translation for; most notes (the fully
// dynamic reconcile-explain prose) have no key at all and fall back to the
// English text as-is -- same as an older server response with no
// notes_i18n field.
const NOTE_I18N_KEY = {
  yearsReversed: "noteYearsReversed",
  segmentalColumns: "noteSegmentalColumns",
  assetsOnlyIncomplete: "noteAssetsOnlyIncomplete",
};
function noteLines(d){
  const en = d.notes || [];
  const meta = d.notes_i18n || [];
  return en.map((text, i)=>{
    const m = meta[i];
    const key = m && NOTE_I18N_KEY[m.key];
    return key ? t(key, m.vars || {}) : text;
  });
}
// same "worth a second look before exporting" threshold the confirm-dialog
// uses at export time (health_bad_below in tablekit/config.py) -- one place
// so the tray's warning icon and the export-time confirmation never disagree
function isRisky(tbl){
  return !!tbl && (tbl.foots===false || (tbl.health!=null && tbl.health<0.75));
}
function renderTray(){
  const chosen = state.order.filter(n=>state.sel.has(n));
  const tray = $("#tray");
  if (!chosen.length){ tray.innerHTML = `<h2>${t('exportOrder')}</h2><div class="empty">${t('tickToAdd')}</div>`; return; }
  tray.innerHTML = `<h2>${t('exportOrderN',{n:chosen.length})}</h2>` + chosen.map((n,i)=>{
    const tbl = state.tables.find(x=>x.n===n);
    const risky = isRisky(tbl);
    return `<div class="item" data-n="${n}"><span class="nm${risky?" risky":""}">${i+1}. ${nameOf(tbl)}${risky?" ⚠":""}</span>
      <button data-up="${n}" ${i===0?"disabled":""} aria-label="${t('moveUpTitle')}" title="${t('moveUpTitle')}">↑</button>
      <button data-down="${n}" ${i===chosen.length-1?"disabled":""} aria-label="${t('moveDownTitle')}" title="${t('moveDownTitle')}">↓</button>
      <button data-rm="${n}" aria-label="${t('removeFromOrderTitle')}" title="${t('removeFromOrderTitle')}">×</button></div>`;
  }).join("");
  tray.querySelectorAll("[data-up]").forEach(b=>b.onclick=()=>moveSel(+b.dataset.up,-1));
  tray.querySelectorAll("[data-down]").forEach(b=>b.onclick=()=>moveSel(+b.dataset.down,1));
  tray.querySelectorAll("[data-rm]").forEach(b=>b.onclick=()=>{
    const cb=$(`#list input[data-n="${b.dataset.rm}"]`); if(cb) cb.checked=false;
    toggleSel(+b.dataset.rm,false); renderList([]);
  });
}
function moveSel(n,dir){
  const chosen = state.order.filter(x=>state.sel.has(x));
  const i=chosen.indexOf(n), j=i+dir;
  if (j<0||j>=chosen.length) return;
  [chosen[i],chosen[j]]=[chosen[j],chosen[i]];
  state.order = [...chosen, ...state.order.filter(x=>!state.sel.has(x))];
  renderTray(); saveUiState();
}
function syncExport(){
  const b=$("#exportBtn"), n=state.sel.size;
  b.disabled = n===0; b.textContent = n ? t('exportBtnN',{n}) : t('export');
}

// --------------------------------------------------------------- picker ----
// Tabula-style manual mode: browse the raw pages of the uploaded file, drag
// a box around the table you want, extract just that region. This is the
// ONLY way to grab a table -- the whole-file automatic detector was removed
// (its false positives/misses weren't reliable enough to ship).
async function showPicker(){
  // NOTE: this only (re)draws the picker pane -- it no longer deselects
  // whatever table is showing in the detail pane, since both are always
  // visible together now (see ensureSplitLayout). "+ Select a table" jumps
  // focus back to the picker without losing your place in the result.
  if (!state.file) return;
  ensureSplitLayout();
  if (state.pageCount == null){
    try{
      const pc = await jget("/api/pagecount?file="+encodeURIComponent(state.file));
      state.pageCount = pc.pages;
    }catch(e){
      // leave state.pageCount at null (NOT 1) -- this only fetches once
      // per file, so caching a wrong "1" here would silently cap the
      // picker at a single page for the rest of the session even on a
      // 171-page report; null instead makes the next "+ Select a table"
      // retry the fetch, and the toast makes the failure visible now
      showErrorToast(t("pageCountFailed",{msg:e.message}));
    }
  }
  renderPicker();
  await loadPickerPage(state.page);
}

function renderPicker(){
  $("#pickerPane").innerHTML = `
    <div class="pv-head">
      <h2>${t('selectATableH2')}</h2>
      <span class="where">${t('pageWord')}
        <input class="pgnum" id="pgNum" value="${state.page}" aria-label="${t('pageNumTitle')}"> ${t('ofTotal',{total:state.pageCount||"?"})}</span>
      <button class="btn" id="pgPrev" ${state.page<=1?"disabled":""}>${t('prevBtn')}</button>
      <button class="btn" id="pgNext" ${state.pageCount&&state.page>=state.pageCount?"disabled":""}>${t('nextBtn')}</button>
    </div>
    <div class="search-box">
      <input id="searchQ" placeholder="${t('searchPh')}" value="${state.searchQ||""}">
      <button class="btn" id="searchBtn">${t('searchBtn')}</button>
      ${state.searchQ?`<button class="btn" id="searchClear">×</button>`:""}
    </div>
    <div id="searchResults"></div>
    <div id="quickfindRow"></div>
    <div id="extractToasts"></div>
    <div class="pv-lines">${t('dragBoxHint')}</div>
    <div class="card">
      <div class="cap">${t('pageCap',{page:state.page})}<span class="spacer"></span>
        <button class="mini" id="pickerZoomBtn">${state.pickerZoom?t('fitWidth'):t('actualSize')}</button>
        <button class="btn-primary" id="extractBtn" disabled>${t('extractBtn')}</button></div>
      <div class="pagebox ${state.pickerZoom?"actual":""}" id="pickerBox"><div class="picker-stage" id="pickerStage">
        <img id="pgImg" alt="${t('pgImgAlt')}"><canvas id="boxCanvas"></canvas>
      </div></div>
    </div>`;
  $("#pgPrev").onclick = ()=>gotoPage(state.page-1);
  $("#pgNext").onclick = ()=>gotoPage(state.page+1);
  $("#pgNum").onchange = e=>gotoPage(parseInt(e.target.value,10)||1);
  $("#extractBtn").onclick = doExtractRegion;
  $("#pickerZoomBtn").onclick = ()=>{
    state.pickerZoom = !state.pickerZoom;
    $("#pickerBox").classList.toggle("actual", state.pickerZoom);
    $("#pickerZoomBtn").textContent = state.pickerZoom ? t('fitWidth') : t('actualSize');
  };
  $("#searchBtn").onclick = runSearch;
  $("#searchQ").onkeydown = e=>{ if (e.key==="Enter"){ e.preventDefault(); runSearch(); } };
  const sc = $("#searchClear");
  if (sc) sc.onclick = ()=>{ state.searchQ=""; state.searchHits=null; renderPicker(); loadPickerPage(state.page); };
  if (state.searchHits) renderSearchResults();
  renderQuickfind();
}

function renderQuickfind(){
  const el = $("#quickfindRow");
  if (!el) return;
  const hits = state.quickfind;
  if (hits === null || hits === undefined){
    el.innerHTML = `<div class="quickfind"><span class="msg" style="padding:0">${t('locatingStatements')}</span></div>`;
    return;
  }
  if (!hits.length){ el.innerHTML = ""; return; }
  el.innerHTML = `<div class="quickfind">
    <span class="msg" style="padding:0">${t('jumpTo')}</span>
    ${hits.map(h=>`<button class="chip" data-pg="${h.page}">${h.label} · p${h.page}</button>`).join("")}
  </div>`;
  el.querySelectorAll(".chip").forEach(b=> b.onclick = ()=>gotoPage(+b.dataset.pg));
}

async function runSearch(){
  const q = $("#searchQ").value.trim();
  state.searchQ = q;
  if (q.length < 2){ $("#searchResults").innerHTML =
    `<div class="msg" style="padding:6px 2px">${t('typeAtLeast2')}</div>`; return; }
  $("#searchResults").innerHTML = `<div class="msg" style="padding:6px 2px"><span class="spin"></span>${t('searching')}</div>`;
  try{
    const res = await jget(`/api/search?file=${encodeURIComponent(state.file)}&q=${encodeURIComponent(q)}`);
    state.searchHits = res.hits;
    renderSearchResults();
    loadPickerPage(state.page);   // re-render the current page with matches highlighted
  }catch(e){ $("#searchResults").innerHTML = `<div class="msg">${t('searchFailed',{msg:e.message})}</div>`; }
}

function escapeRe(s){ return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); }
function markMatches(text, q){
  const esc = text.replace(/</g,"&lt;");
  if (!q) return esc;
  const re = new RegExp("("+escapeRe(q).replace(/</g,"&lt;")+")", "ig");
  return esc.replace(re, "<mark>$1</mark>");
}

function renderSearchResults(){
  const hits = state.searchHits || [];
  const el = $("#searchResults");
  if (!el) return;
  if (!hits.length){ el.innerHTML = `<div class="msg" style="padding:6px 2px">${t('noMatches')}</div>`; return; }
  el.innerHTML = `<div class="search-hits">${hits.map(h=>`
    <div class="hit" data-pg="${h.page}">
      <span class="pg">p${h.page}</span>
      <span class="snip">…${markMatches(h.snippet, state.searchQ)}…</span>
      <span class="cnt">${h.count}×</span>
    </div>`).join("")}</div>`;
  el.querySelectorAll(".hit").forEach(row=>
    row.onclick = ()=>gotoPage(+row.dataset.pg));
}

function gotoPage(n){
  n = Math.max(1, Math.min(state.pageCount||1, n));
  if (n === state.page && $("#pgImg")) return;
  state.page = n; state.box = null;
  renderPicker();
  loadPickerPage(n);
}

async function loadPickerPage(n){
  const img = $("#pgImg");
  let src = `/api/page_raw?file=${encodeURIComponent(state.file)}&n=${n}&scale=${PICKER_SCALE}`;
  if (state.searchQ) src += `&hl=${encodeURIComponent(state.searchQ)}`;
  img.src = src;
  img.alt = t('pageCap', {page: n});
  await new Promise(res=>{ img.onload = res; img.onerror = res; });
  const cv = $("#boxCanvas");
  if (!cv) return;   // user navigated away while the image was loading
  // the canvas's DRAWING BUFFER stays at the image's full native resolution
  // (crisp lines) -- it's the DISPLAYED size that's left to CSS (max-width:
  // 100%; height:auto, same as the img), so the page fits the pane instead
  // of forcing horizontal scroll for every single page. wireCanvas()
  // converts pointer coordinates from that displayed size back to this
  // buffer's coordinate space before using them for anything.
  cv.width = img.naturalWidth; cv.height = img.naturalHeight;
  wireCanvas(cv);
}

function wireCanvas(cv){
  const ctx = cv.getContext("2d");
  let drag = null;
  // pointer events report offsetX/offsetY in the canvas's DISPLAYED (CSS)
  // size, which can now be smaller than its drawing-buffer resolution
  // (scaled down to fit the pane) -- convert to buffer space before using.
  const toBuffer = (x, y) => {
    const rx = cv.width / cv.clientWidth, ry = cv.height / cv.clientHeight;
    return [x * rx, y * ry];
  };
  const redraw = (x0,y0,x1,y1)=>{
    ctx.clearRect(0,0,cv.width,cv.height);
    const x=Math.min(x0,x1), y=Math.min(y0,y1), w=Math.abs(x1-x0), h=Math.abs(y1-y0);
    const hex = (getComputedStyle(document.documentElement).getPropertyValue("--accent").trim()) || "#d4af37";
    const r = parseInt(hex.slice(1,3),16), g = parseInt(hex.slice(3,5),16), b = parseInt(hex.slice(5,7),16);
    ctx.fillStyle = `rgba(${r},${g},${b},.18)`; ctx.fillRect(x,y,w,h);
    ctx.strokeStyle = hex; ctx.lineWidth = 2; ctx.strokeRect(x,y,w,h);
  };
  cv.onpointerdown = e=>{
    cv.setPointerCapture(e.pointerId);
    const [bx,by] = toBuffer(e.offsetX, e.offsetY);
    drag = {x0:bx, y0:by, dispX0:e.offsetX, dispY0:e.offsetY};
    const eb = $("#extractBtn"); if (eb) eb.disabled = true;
  };
  cv.onpointermove = e=>{
    if (!drag) return;
    const [bx,by] = toBuffer(e.offsetX, e.offsetY);
    redraw(drag.x0, drag.y0, bx, by);
  };
  cv.onpointerup = e=>{
    if (!drag) return;
    // the "too small, ignore it" check is in DISPLAYED pixels (what the user
    // actually dragged), not buffer pixels -- otherwise, once the page is
    // scaled down to fit the pane, a tiny on-screen wiggle could translate
    // to well over the buffer-space threshold and register as a real box
    const dispW = Math.abs(e.offsetX - drag.dispX0), dispH = Math.abs(e.offsetY - drag.dispY0);
    const [bx,by] = toBuffer(e.offsetX, e.offsetY);
    const x0=drag.x0, y0=drag.y0, x1=bx, y1=by;
    drag = null;
    if (dispW < 6 || dispH < 6){ ctx.clearRect(0,0,cv.width,cv.height); return; }
    // state.box stays in the canvas's BUFFER (native render-pixel) space,
    // same as before -- doExtractRegion's ÷PICKER_SCALE conversion is unchanged
    state.box = [Math.min(x0,x1), Math.min(y0,y1), Math.max(x0,x1), Math.max(y0,y1)];
    const eb = $("#extractBtn"); if (eb) eb.disabled = false;
  };
}

async function doExtractRegion(){
  if (!state.box) return;
  const b = $("#extractBtn");
  b.disabled = true; b.textContent = t("extracting");
  try{
    const pdfBox = state.box.map(v=>v/PICKER_SCALE);
    const d = await jpost("/api/extract_region",
      {file:state.file, page:state.page, bbox:pdfBox});
    await _afterExtract(d);
  }catch(e){
    // OCR is a failsafe the user clicks, never something that fires on its
    // own -- offer it only when the box is genuinely textless (a scan) and
    // the server actually has an OCR engine to run.
    if (e.body && e.body.no_text && e.body.ocr_available){
      showOcrOfferToast(state.box.map(v=>v/PICKER_SCALE));
    } else {
      showErrorToast(t("extractionFailed",{msg:e.message}));
    }
  } finally {
    const eb = $("#extractBtn");
    if (eb){ eb.disabled = true; eb.textContent = t("extractBtn"); }
  }
}

async function doExtractRegionOcr(pdfBox){
  const b = $("#extractBtn");
  if (b){ b.disabled = true; b.textContent = t("runningOcr"); }
  try{
    const d = await jpost("/api/extract_region_ocr",
      {file:state.file, page:state.page, bbox:pdfBox});
    await _afterExtract(d);
  }catch(e){
    showErrorToast(t("ocrFailed",{msg:e.message}));
  } finally {
    if (b){ b.disabled = true; b.textContent = t("extractBtn"); }
  }
}

async function _afterExtract(d){
  const inv = await jget("/api/scan?file="+encodeURIComponent(state.file));
  state.tables = inv.tables; state.order = inv.tables.map(x=>x.n);
  renderList(inv.warnings||[]);
  updateCurrentFileQueueLabel();
  state._lastDetail = d;
  // stay on the picker (same page) so another table can be grabbed right
  // away -- clear the box and show a small confirmation. The result also
  // shows up immediately in the detail pane right beside the picker
  // (they're both always on screen now), no click needed to "go look at it".
  state.box = null;
  const cv = $("#boxCanvas");
  if (cv) cv.getContext("2d").clearRect(0,0,cv.width,cv.height);
  showExtractToast(d);
  state.active = d.n;
  $("#list").querySelectorAll(".row").forEach(r=>r.classList.toggle("active", +r.dataset.n===d.n));
  renderPreview(d.n, d);
}

function showExtractToast(d){
  const host = $("#extractToasts");
  if (!host) return;
  const el = document.createElement("div");
  el.className = "extract-toast";
  el.innerHTML = `${t('extractedPrefix')} <b>${(nameOf(d)||t('kind_table')).replace(/</g,"&lt;")}</b>
    <span class="spacer"></span>
    <button class="btn" data-act="view">${t('viewBtn')}</button>
    <button class="close" data-act="dismiss" title="${t('dismissTitle')}" aria-label="${t('dismissTitle')}">×</button>`;
  host.prepend(el);
  el.querySelector('[data-act="view"]').onclick = ()=>{
    state.active = d.n;
    $("#list").querySelectorAll(".row").forEach(r=>r.classList.toggle("active", +r.dataset.n===d.n));
    renderPreview(d.n, d);
  };
  el.querySelector('[data-act="dismiss"]').onclick = ()=> el.remove();
  setTimeout(()=>{ if (el.isConnected) el.remove(); }, 8000);
}

// -------------------------------------------------------------- preview ----
function baseRows(n, d){
  const e = state.edits[n];
  return (e && e.rows ? e.rows : d.rows).map(r=>r.slice());
}
// Undo, one row-mutation at a time: every place that's about to CHANGE
// rows (cell edit, insert/delete/split/merge-up, swap year columns) snap-
// shots the rows as they stood right before, onto a small per-table stack.
// undoLastEdit() below is the only thing that ever pops it -- it restores
// via pushEdit() directly rather than routing back through one of those
// three call sites, so undoing never pushes a new snapshot of its own.
function snapshotForUndo(n, d){
  const arr = state.undoStack[n] = state.undoStack[n] || [];
  arr.push(baseRows(n, d));
  if (arr.length > 20) arr.shift();
}
function undoLastEdit(n){
  const arr = state.undoStack[n];
  if (!arr || !arr.length){ showInfoToast(t("nothingToUndo"), 3000); return; }
  const rows = arr.pop();
  pushEdit(n, rows);
  renderPreview(n, state._lastDetail);
}
function pushEdit(n, rows, extra){
  const e = state.edits[n] = Object.assign({dirty:new Set()}, state.edits[n]||{}, extra||{});
  if (rows) e.rows = rows;
  renderList([]);
  scheduleReanalyze(n);
}
// visible only while the ACTIVE table has a save in flight/just landed --
// harmless no-op if the toolbar for a different table is on screen right
// now (setSaveStatus itself no-ops when its target isn't the shown table)
function setSaveStatus(n, kind){
  const el = $("#saveStatus");
  if (!el || state.active !== n) return;
  if (!kind){ el.hidden = true; return; }
  el.hidden = false;
  el.className = "save-status " + kind;
  el.textContent = kind==="saving" ? t("savingStatus")
                  : kind==="saved" ? t("savedStatus") : t("saveErrorStatus");
  if (kind === "saved") setTimeout(()=>{ if (el.className.includes("saved")) el.hidden = true; }, 2500);
}
// Debounced commit-to-server for an edit, one independent timer PER TABLE
// (state.reTimers[n]) -- NOT a single shared timer. A single shared timer
// was a real, reproduced bug: switching to table B within another table
// A's 450ms window cancelled A's pending save outright, silently, with
// state.edits[n] still showing A as "edited" in the browser the whole
// time. `file` is captured now, at scheduling time, not re-read from
// state.file when the timer fires -- if the user has switched to a
// DIFFERENT FILE by then, state.edits was already wiped by loadFile() and
// this table number could now mean something else entirely in the new
// file; firing against a stale `file`+fresh-`state.edits[n]` combination
// is exactly how one file's typed edit could land in another file's
// table, so this just no-ops rather than risk that.
function scheduleReanalyze(n){
  const file = state.file;
  clearTimeout(state.reTimers[n]);
  setSaveStatus(n, "saving");
  state.reTimers[n] = setTimeout(async ()=>{
    delete state.reTimers[n];
    if (state.file !== file) return;
    try{
      const e = state.edits[n] || {};
      const d = await jpost("/api/reanalyze",
        {file, n, rows:e.rows||null, title:e.title||null});
      // renderPreview rebuilds the toolbar (a fresh #saveStatus, hidden by
      // default) -- set the "saved" indicator AFTER, or the re-render wipes
      // it before it's ever seen
      if (state.active === n){
        state._lastDetail = d;
        renderPreview(n, d);
      }
      setSaveStatus(n, "saved");
    }catch(err){
      // state.edits[n] still holds the edit either way -- a later edit to
      // this table retries the save, and export always sends it fresh
      // regardless of whether any autosave ever landed
      setSaveStatus(n, "save-error");
    }
  }, 450);
}

async function openTable(n){
  state.active = n;
  $("#list").querySelectorAll(".row").forEach(r=>r.classList.toggle("active",+r.dataset.n===n));
  const d = await jget(`/api/table?file=${encodeURIComponent(state.file)}&n=${n}`);
  state._lastDetail = d;
  renderPreview(n, d);
}

function renderPreview(n, d){
  const nm = kindName(d.kind)||t('kind_table');
  const yrs = d.years.length ? d.years.join(" / ") : "";
  const ed = state.edits[n];
  const rows = baseRows(n, d);
  const dirty = (ed && ed.dirty) || new Set();
  const hi=d.header_idx, ds=d.data_start, vc=d.value_cols||[], tot=new Set(d.total_rows||[]);
  const susp={}; (d.suspect||[]).forEach(s=>susp[s.row]=s);
  const showDelta = vc.length>=2;
  const ncols = Math.max(...rows.map(r=>r.length), 1);

  let bodyRows = "";
  rows.forEach((r,i)=>{
    const isHdr=i<=hi, di=i-ds;
    const cls=[isHdr?"hdr":"", tot.has(i)?"total":""].filter(Boolean).join(" ");
    let tds = `<td class="rowctl">`+
      `<button data-act="ins" data-i="${i}" title="${t('insertRowBelow')}" aria-label="${t('insertRowBelow')}">＋</button>`+
      `<button data-act="del" data-i="${i}" title="${t('deleteRow')}" aria-label="${t('deleteRow')}">✕</button>`+
      `<button data-act="split" data-i="${i}" title="${t('splitLabel')}" aria-label="${t('splitLabel')}">⤶</button>`+
      `<button data-act="mergeup" data-i="${i}" title="${t('mergeUp')}" aria-label="${t('mergeUp')}" ${i===0?"disabled":""}>⭡</button></td>`;
    // the "fmt" side-channel is keyed to the SERVER's own rows -- once the
    // user has typed an edit for this table, row/column positions can no
    // longer be trusted to still line up with it, so stick to plain numbers
    // (never wrong, just loses the $/% cosmetics) rather than risk
    // mismatched formatting on an edited cell.
    const fRow = (!ed && d.fmt && d.fmt[i]) || null;
    for (let ci=0; ci<ncols; ci++){
      const c=r[ci], num=typeof c==="number";
      let kl = num ? "num" : "";
      if (ci===0 && !isHdr && susp[di]) kl += susp[di].kind==="figure" ? " figbad" : " suspect";
      if (dirty.has(i+","+ci)) kl += " dirty";
      const title = (ci===0 && susp[di]) ? ` title="${susp[di].why}"` : "";
      // data-orig: the EXACT text this cell was rendered with (e.g. "26.6%",
      // not the bare 26.6 behind it) -- commitCell diffs against this, not
      // against rows[r][c], so merely clicking into a $/%-formatted cell and
      // clicking back out isn't mistaken for an edit.
      const shown = fmt(c, fRow && fRow[ci]);
      tds += `<td class="${kl}"${title} data-r="${i}" data-c="${ci}" data-orig="${shown.replace(/"/g,"&quot;")}" contenteditable="plaintext-only">${shown}</td>`;
    }
    if (showDelta && !isHdr){
      const a=r[vc[0]], b=r[vc[1]];
      if (typeof a==="number" && typeof b==="number"){
        const dd=a-b, pc=b?(100*dd/Math.abs(b)):null;
        tds += `<td class="delta">${dd.toLocaleString()}</td><td class="delta">${pc==null?"":pc.toFixed(1)+"%"}</td>`;
      } else tds += `<td class="delta"></td><td class="delta"></td>`;
    } else if (showDelta){
      tds += `<td class="delta">Δ</td><td class="delta">Δ%</td>`;
    }
    bodyRows += `<tr class="${cls}">${tds}</tr>`;
  });

  const consist = d.consistency ? (()=>{
    const c=d.consistency, ok=c.mismatch===0;
    const worst=(c.worst||[]).map(w=>`${w[0]}: ${Number(w[1]).toLocaleString()} vs ${Number(w[2]).toLocaleString()}`).join(" · ");
    const tail = ok ? t('agree')
      : t('differBase',{mismatch:c.mismatch, checked:c.checked}) + ` <b>${c.verdict||t('differsWord')}</b>`;
    return `<div class="pv-consist ${ok?"ok":"bad"}">${t('priorYearVs',{year:c.year, vs:c.vs})}
       ${tail}${worst?` — ${worst}`:""}</div>`;
  })() : "";
  const notes = noteLines(d).map(x=>`<div class="pv-note">${x}</div>`).join("");
  const lines = (d.foot_by_col||[]).map(fc=>{
    const m=fc.ok===true?"✓":fc.ok===false?"✗":"?"; return `${fc.year} ${m}  ${fc.worked}`;
  }).join("\n") || d.foot_detail || "";

  ensureSplitLayout();
  $("#detailPane").innerHTML = `
    <div class="pv-head">
      <h2>${nm}${yrs?` <span class="where">· ${yrs}</span>`:""}</h2>
      <span class="where">${t('pageLabel',{page:d.page})}</span>
      ${footPill(d.foots, d.foot_by_col)}
      <span class="where">${t('labelsFigures',{l:Math.round((d.health_labels??1)*100), f:Math.round((d.health_figures??1)*100)})}</span>
      ${d.ocr?`<span class="pill warn" title="${t('ocrPillTitle')}">${t('ocrVerifyByEye')}</span>`:``}
      ${d.edited?`<span class="pill warn">${t('editedRecomputed')}</span>`:``}
    </div>
    ${notes}${consist}
    <div class="pv-lines">${lines}</div>

    <div class="toolbar">
      <input class="sheet" id="sheetName" placeholder="${t('sheetNamePh')}" value="${(ed&&ed.title)||""}">
      ${vc.length>=2?`<button class="btn" id="swapYears">${t('swapYears')}</button>`:""}
      ${(state.undoStack[n]||[]).length?`<button class="btn" id="undoEditBtn" title="${t('undoEditTitle')}">${t('undoEditBtn')}</button>`:""}
      ${ed?`<button class="btn" id="resetEdits">${t('undoEdits')}</button>`:""}
      <span class="hint">${t('editHint')}</span>
      ${state.reTimers[n]
        ? `<span class="save-status saving" id="saveStatus">${t('savingStatus')}</span>`
        : `<span class="save-status" id="saveStatus" hidden></span>`}
    </div>

    <div class="card">
      <div class="cap">${t('extractedTableRows',{n:rows.length})}
        ${(d.suspect||[]).length?`<span class="spacer"></span><span style="color:var(--warn)">${t('amberRedReview')}</span>`:``}
      </div>
      <div class="tablebox"><table class="tbl"><tbody>${bodyRows}</tbody></table></div>
    </div>

    <div id="compare"></div>`;

  // the picker (right there in the OTHER pane) is now the only PDF preview --
  // no more redundant second copy of the same page. Just make sure it's
  // actually showing the page this table came from (a no-op if it already is).
  if (state.page !== d.page) gotoPage(d.page);

  wirePreview(n, d);
  renderComparePanel(n, d);
}

function wirePreview(n, d){
  $("#sheetName").oninput = e => pushEdit(n, null, {title:e.target.value});
  const sy=$("#swapYears"); if(sy) sy.onclick=()=>{
    snapshotForUndo(n, d);
    const rows=baseRows(n,d); const [a,b]=d.value_cols;
    rows.forEach(r=>{const t=r[a];r[a]=r[b];r[b]=t;});
    pushEdit(n, rows); openTable(n);
  };
  const re=$("#resetEdits"); if(re) re.onclick=()=>{ delete state.edits[n]; state.undoStack[n]=[]; renderList([]); openTable(n); };
  const ue=$("#undoEditBtn"); if(ue) ue.onclick=()=> undoLastEdit(n);

  $("#detailPane").querySelectorAll("td[contenteditable]").forEach(td=>{
    td.addEventListener("blur", ()=>commitCell(n, d, +td.dataset.r, +td.dataset.c, td));
    td.addEventListener("keydown", e=>{ if(e.key==="Enter"){ e.preventDefault(); td.blur(); }});
  });
  $("#detailPane").querySelectorAll("td.rowctl button").forEach(b=>{
    b.onclick = ()=> rowAction(n, d, b.dataset.act, +b.dataset.i);
  });
}

function commitCell(n, d, r, c, td){
  const rows = baseRows(n, d);
  const raw = td.textContent.trim();
  if (raw === (td.dataset.orig ?? "")) return;
  snapshotForUndo(n, d);
  rows[r][c] = raw;                       // keep the RAW string; server parses it
  const e = state.edits[n] = Object.assign({dirty:new Set()}, state.edits[n]||{});
  e.rows = rows; e.dirty.add(r+","+c);
  renderList([]);
  scheduleReanalyze(n);
}

function rowAction(n, d, act, i){
  snapshotForUndo(n, d);
  let rows = baseRows(n, d);
  if (act === "ins"){
    rows.splice(i+1, 0, new Array(Math.max(...rows.map(x=>x.length),1)).fill(""));
  } else if (act === "del"){
    rows.splice(i, 1);
  } else if (act === "mergeup" && i>0){
    const lab = [rows[i-1][0], rows[i][0]].filter(Boolean).join(" ").trim();
    // keep the figures from whichever row has them
    const figs = rows[i].slice(1).some(v=>v!=="" && v!=null) ? rows[i] : rows[i-1];
    rows[i-1] = [lab, ...figs.slice(1)];
    rows.splice(i, 1);
  } else if (act === "split"){
    const lab = String(rows[i][0]||"");
    const words = lab.split(/\s+/);
    if (words.length >= 2){
      const mid = Math.ceil(words.length/2);
      const top = [words.slice(0,mid).join(" "), ...rows[i].slice(1)];
      const bot = [words.slice(mid).join(" "), ...rows[i].slice(1).map(()=>"")];
      rows.splice(i, 1, top, bot);
    }
  }
  pushEdit(n, rows);
  const d2 = state._lastDetail;
  renderPreview(n, d2);
}

// -------------------------------------------------------------- compare ----
function renderComparePanel(n, d){
  const el = $("#compare");
  if (state.files.length < 2){ el.innerHTML = ""; return; }
  const others = state.files.filter(f=>f!==state.file);
  el.innerHTML = `
    <div class="cmp-head"><b>${t('compareTitle')}</b>
      <label class="where" for="cmpFile">${t('compareVs')}</label>
      <select id="cmpFile">${others.map(f=>`<option>${f}</option>`).join("")}</select>
      <button class="btn" id="cmpRun">${t('compareBtn')}</button></div>
    <div id="cmpOut"></div>`;
  $("#cmpRun").onclick = async ()=>{
    const bfile = $("#cmpFile").value;
    $("#cmpOut").innerHTML = `<div class="msg"><span class="spin"></span>${t('scanningMatching',{file:bfile})}</div>`;
    try{
      const binv = await jget("/api/scan?file="+encodeURIComponent(bfile));
      const match = binv.tables.find(x=>x.kind===d.kind);
      if (!match){ $("#cmpOut").innerHTML = `<div class="msg">${t('noKindFoundIn',{kind:kindName(d.kind), file:bfile})}</div>`; return; }
      const cmp = await jpost("/api/compare",{a:state.file,na:n,b:bfile,nb:match.n});
      const hdr = cmp.rows[0];
      const isWarn = /RESTATED/.test(cmp.verdict);
      // `counts` (plain numbers) lets the verdict SENTENCE be translated;
      // the free-form per-row `notes`/reconcile-explain text elsewhere is
      // still English-only server prose -- see the note above I18N for why
      // (a much larger change to extract_all_tables.py itself). Fall back
      // to the raw English string if an older server response lacks it.
      const c = cmp.counts;
      const verdictText = c
        ? t(isWarn?"cmpVerdictRestated":"cmpVerdictClean",
            {changed:c.changed, new:c.new, removed:c.removed, restated:c.restated})
        : cmp.verdict;
      let rowsHtml = "";
      for (let r=1;r<cmp.rows.length;r++){
        const row = cmp.rows[r];
        const restated = row[7] ? " restated" : "";
        rowsHtml += `<tr class="${restated}">` + row.map((v,ci)=>
          `<td class="${(ci>=1&&ci<=6&&typeof v==='number')?'num':''}">${v==null?"":(typeof v==='number'?v.toLocaleString():v)}${ci===0&&restated?`<span class="restated-tag">RESTATED</span>`:""}</td>`
        ).join("") + `</tr>`;
      }
      $("#cmpOut").innerHTML = `<div class="cmp-verdict ${isWarn?"warn":"ok"}">${isWarn?"⚠":"✓"} ${verdictText}</div>
        <div style="max-height:44vh;overflow:auto"><table>
          <thead><tr>${hdr.map(h=>`<th>${h}</th>`).join("")}</tr></thead>
          <tbody>${rowsHtml}</tbody></table></div>`;
    }catch(e){ $("#cmpOut").innerHTML = `<div class="msg">${t('compareFailed',{msg:e.message})}</div>`; }
  };
}

// ---------------------------------------------------------------- header ----
function fileToBase64(file){
  return new Promise((resolve,reject)=>{
    const r = new FileReader();
    r.onload = ()=> resolve(String(r.result).split(",")[1]);
    r.onerror = reject;
    r.readAsDataURL(file);
  });
}
async function uploadPdfFile(f){
  const btn = $("#uploadBtn"); const label = btn.firstChild;
  label.textContent = t("uploading");
  try{
    const data_b64 = await fileToBase64(f);
    const res = await jpost("/api/upload", {filename:f.name, data_b64});
    if (!state.files.includes(res.file)) state.files.push(res.file);
    const sel = $("#fileSel");
    sel.hidden = state.files.length <= 1;
    await loadFile(res.file);
    if (state.files.length > 1) refreshFileQueueLabels();
  }catch(err){ showErrorToast(t("uploadFailed",{msg:err.message})); }
  finally{ label.textContent = t("upload"); }
}
$("#uploadInput").onchange = async (e)=>{
  const f = e.target.files[0]; if (!f) return;
  await uploadPdfFile(f);
  e.target.value = "";
};

$("#pickerBtn").onclick = ()=>showPicker();
$("#fileName").addEventListener("input", saveUiState);

// always-visible header control (not tied to any one table's view) --
// arranges the PERSISTENT picker+detail split (#splitRegion) -- both panes
// are always present once a file is loaded, so this always has something
// visible to change, regardless of what's currently in the detail pane
function setLayoutBtnLabel(){
  $("#layoutBtn").textContent = state.sideBySide ? t("stackPanes") : t("sideBySide");
}
setLayoutBtnLabel();
$("#layoutBtn").onclick = ()=>{
  state.sideBySide = !state.sideBySide;
  try{ localStorage.setItem("tk_side_by_side", state.sideBySide ? "1" : "0"); }catch(e){}
  setLayoutBtnLabel();
  const split = $("#splitRegion");
  if (split) split.classList.toggle("side-by-side", state.sideBySide);
};
$("#selAllStmt").onclick = ()=>{
  state.tables.filter(t=>REAL_STMT.has(t.kind)).forEach(t=>toggleSel(t.n,true));
  renderList([]);
};
$("#selClear").onclick = ()=>{ state.sel.clear(); syncExport(); renderTray(); renderList([]); };

$("#exportBtn").onclick = async ()=>{
  const ns = state.order.filter(n=>state.sel.has(n));
  const risky = ns.map(n=>state.tables.find(x=>x.n===n)).filter(isRisky);
  if (risky.length){
    const ok = await confirmDialog(
      t('riskyExportWarn',{risky:risky.length, total:ns.length}) + `\n\n`
      + risky.map(x=>"  • "+nameOf(x)).join("\n")
      + `\n\n` + t('exportAnywayQ'), t('exportAnyway'));
    if (!ok) return;
  }
  const edits = {};
  for (const n of ns){ const e=state.edits[n];
    if (e && (e.rows||e.title)) edits[n] = {rows:e.rows||null, title:e.title||null}; }
  const b=$("#exportBtn"); b.disabled=true; b.textContent=t("building");
  try{
    const res = await fetch("/api/export",{method:"POST",headers:{"Content-Type":"application/json"},
      body:JSON.stringify({file:state.file, ns, edits})});
    if(!res.ok) throw new Error((await res.json()).error||res.statusText);
    const blob = await res.blob();
    // Unicode-aware (\p{L}\p{N}, not the ASCII-only \w) so an Arabic
    // filename -- the auto-filled default, or anything typed by hand --
    // survives sanitisation instead of collapsing into underscores.
    let fn = ($("#fileName").value.trim()||t('brand').toLowerCase()).replace(/[^\p{L}\p{N} .()-]/gu,"_");
    if(!/\.xlsx$/i.test(fn)) fn += ".xlsx";
    const a=document.createElement("a"); a.href=URL.createObjectURL(blob); a.download=fn; a.click();
    URL.revokeObjectURL(a.href);
  }catch(e){ showErrorToast(t("exportFailed",{msg:e.message})); }
  syncExport();
};

// ------------------------------------------------------------- splitter ----
// Adjustable panel widths: drag the splitter to resize the sidebar; the
// page-picker and table panes below use native CSS `resize` (drag their
// bottom-right corner) since each already scrolls independently.
(function(){
  const aside = document.querySelector("aside"), split = $("#splitter");
  try{
    const saved = localStorage.getItem("tk_sidebar_w");
    if (saved) aside.style.width = saved+"px";
  }catch(e){}
  let dragging = false;
  split.addEventListener("mousedown", e=>{
    dragging = true; split.classList.add("active");
    e.preventDefault();
  });
  window.addEventListener("mousemove", e=>{
    if (!dragging) return;
    const w = Math.max(200, Math.min(window.innerWidth*0.7, e.clientX));
    aside.style.width = w+"px";
  });
  window.addEventListener("mouseup", ()=>{
    if (!dragging) return;
    dragging = false; split.classList.remove("active");
    try{ localStorage.setItem("tk_sidebar_w", String(parseInt(aside.style.width,10))); }catch(e){}
  });
})();

document.addEventListener("keydown", e=>{
  if (["INPUT","TD","TEXTAREA","SELECT"].includes(document.activeElement.tagName)) return;
  if (!state.tables.length) return;
  const ord = state.tables.map(t=>t.n);
  const i = ord.indexOf(state.active);
  if (e.key==="ArrowDown"||e.key==="j"){ e.preventDefault(); openTable(ord[Math.min(ord.length-1,i+1)] ?? ord[0]); }
  else if (e.key==="ArrowUp"||e.key==="k"){ e.preventDefault(); openTable(ord[Math.max(0,i-1)] ?? ord[0]); }
  else if (e.key===" " && state.active!=null){ e.preventDefault();
    const cb=$(`#list input[data-n="${state.active}"]`); if(cb){ cb.checked=!cb.checked; toggleSel(state.active,cb.checked); } }
});

boot();
