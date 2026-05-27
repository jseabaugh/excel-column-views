/**
 * Column Views Add-in
 * ─────────────────────────────────────────────────────────────
 * Saves named column-visibility configurations per worksheet,
 * stored in the workbook's Custom XML Parts so they persist
 * with the file across saves and re-opens.
 *
 * Compatibility: Excel Desktop (Windows/Mac) + Excel Online
 */

"use strict";

/* ── Constants ────────────────────────────────────────────────── */
const XML_NAMESPACE = "https://column-views-addin/v1";
const XML_PART_KEY  = "ColumnViewsData";

/* ── App State ────────────────────────────────────────────────── */
const state = {
  views: {},          // { [sheetName]: View[] }
  activeViews: {},    // { [sheetName]: string (view name) | null }
  currentSheet: null,
  editingViewName: null,   // null = creating new
  columnDefs: [],          // { letter, name } for current sheet
};

/* ── DOM refs ─────────────────────────────────────────────────── */
const $ = id => document.getElementById(id);

const dom = {
  sheetBadge:       $("sheet-badge"),
  panelViews:       $("panel-views"),
  panelEditor:      $("panel-editor"),
  viewsList:        $("views-list"),
  emptyState:       $("empty-state"),
  editorModeLabel:  $("editor-mode-label"),
  viewNameInput:    $("view-name-input"),
  columnsList:      $("columns-list"),
  colCountLabel:    $("col-count-label"),
  toast:            $("toast"),
  btnNewView:       $("btn-new-view"),
  btnShowAll:       $("btn-show-all"),
  btnBack:          $("btn-back"),
  btnSaveView:      $("btn-save-view"),
  btnCancelEditor:  $("btn-cancel-editor"),
  btnCheckAll:      $("btn-check-all"),
  btnUncheckAll:    $("btn-uncheck-all"),
  headerRowInput:   $("header-row-input"),
  btnReloadCols:    $("btn-reload-cols"),
};

/* ── Toast ────────────────────────────────────────────────────── */
let toastTimer = null;
function showToast(msg, type = "info", duration = 2200) {
  dom.toast.textContent = msg;
  dom.toast.className = `toast show ${type}`;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { dom.toast.className = "toast"; }, duration);
}

/* ── Panel navigation ─────────────────────────────────────────── */
function showPanel(name) {
  dom.panelViews .classList.toggle("active", name === "views");
  dom.panelEditor.classList.toggle("active", name === "editor");
}

/* ── Persistence: Custom XML Parts ───────────────────────────── */
async function saveToWorkbook() {
  try {
    await Excel.run(async ctx => {
      const parts = ctx.workbook.customXmlParts;
      const payload = JSON.stringify({ views: state.views, activeViews: state.activeViews });
      const xmlContent = `<${XML_PART_KEY} xmlns="${XML_NAMESPACE}">${escapeXml(payload)}</${XML_PART_KEY}>`;

      // Find & delete existing part
      const existing = parts.getByNamespace(XML_NAMESPACE);
      existing.load("items");
      await ctx.sync();
      for (const item of existing.items) item.delete();
      await ctx.sync();

      // Write new part
      parts.add(xmlContent);
      await ctx.sync();
    });
  } catch (err) {
    console.error("saveToWorkbook error:", err);
  }
}

async function loadFromWorkbook() {
  try {
    await Excel.run(async ctx => {
      const parts = ctx.workbook.customXmlParts;
      const existing = parts.getByNamespace(XML_NAMESPACE);
      existing.load("items");
      await ctx.sync();

      if (existing.items.length === 0) return;

      const part = existing.items[0];
      const xmlProxy = part.getXml();
      await ctx.sync();

      const xml = xmlProxy.value;
      const match = xml.match(/>([^<]+)</);
      if (!match) return;

      const data = JSON.parse(unescapeXml(match[1]));
      state.views       = data.views       || {};
      state.activeViews = data.activeViews || {};
    });
  } catch (err) {
    console.warn("loadFromWorkbook (no saved data yet):", err);
  }
}

function escapeXml(str) {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function unescapeXml(str) {
  return str
    .replace(/&quot;/g, '"')
    .replace(/&gt;/g, ">")
    .replace(/&lt;/g, "<")
    .replace(/&amp;/g, "&");
}

/* ── Column reading ───────────────────────────────────────────── */
async function readColumns(sheetName, headerRow = 1) {
  return Excel.run(async ctx => {
    const sheet = ctx.workbook.worksheets.getItem(sheetName);

    // Step 1: get used range to find start column and column count
    const usedRange = sheet.getUsedRange(true);
    usedRange.load(["columnIndex", "columnCount"]);
    await ctx.sync();
    const startCol = usedRange.columnIndex || 0;   // absolute 0-based column offset
    const colCount = Math.min(usedRange.columnCount || 0, 200);
    if (colCount === 0) return { defs: [], hiddenMap: {} };

    // Step 2: read the header row using absolute sheet coordinates
    // headerRow is 1-based (user input), getRangeByIndexes is 0-based
    const headerRowIndex = Math.max(0, headerRow - 1);
    const headerRange = sheet.getRangeByIndexes(headerRowIndex, startCol, 1, colCount);
    headerRange.load("values");
    await ctx.sync();
    const headerValues = headerRange.values[0] || [];

    const defs = [];
    for (let i = 0; i < colCount; i++) {
      const absColIndex = startCol + i;
      const letter = columnIndexToLetter(absColIndex);
      const headerName = headerValues[i];
      defs.push({
        letter,
        index: absColIndex,   // absolute column index for hiding/showing
        name: headerName !== "" && headerName !== null && headerName !== undefined
          ? String(headerName)
          : letter,
      });
    }

    // Step 3: detect currently hidden columns using absolute indexes
    const hiddenMap = {};
    try {
      for (let i = 0; i < colCount; i++) {
        const absColIndex = startCol + i;
        const cell = sheet.getRangeByIndexes(0, absColIndex, 1, 1);
        cell.load("columnHidden");
        await ctx.sync();
        hiddenMap[absColIndex] = cell.columnHidden;
      }
    } catch (_) {}

    return { defs, hiddenMap };
  });
}

function columnLetterToIndex(letter) {
  let index = 0;
  for (let i = 0; i < letter.length; i++) {
    index = index * 26 + letter.charCodeAt(i) - 64;
  }
  return index - 1;
}

function columnIndexToLetter(index) {
  let letter = "";
  let n = index;
  while (n >= 0) {
    letter = String.fromCharCode(65 + (n % 26)) + letter;
    n = Math.floor(n / 26) - 1;
  }
  return letter;
}

/* ── Apply a View ─────────────────────────────────────────────── */
async function applyView(sheetName, viewName) {
  const sheetViews = state.views[sheetName] || [];
  const view = sheetViews.find(v => v.name === viewName);
  if (!view) return;

  try {
    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getItem(sheetName);

      for (const col of view.columnStates) {
        // col.index is the absolute 0-based column index on the sheet
        const colIndex = typeof col.index === "number" ? col.index : columnLetterToIndex(col.letter);
        const range = sheet.getRangeByIndexes(0, colIndex, 1, 1);
        range.columnHidden = !col.visible;
      }

      await ctx.sync();
    });

    state.activeViews[sheetName] = viewName;
    await saveToWorkbook();
    renderViewsList();
    showToast(`View "${viewName}" applied`, "success");
  } catch (err) {
    console.error("applyView error:", err);
    showToast("Failed to apply view", "error");
  }
}

/* ── Show All Columns ─────────────────────────────────────────── */
async function showAllColumns(sheetName) {
  try {
    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getItem(sheetName);

      // Collect every unique column index ever used across all views for this sheet
      const sheetViews = state.views[sheetName] || [];
      const colIndexes = new Set();
      for (const view of sheetViews) {
        for (const col of view.columnStates) {
          const idx = typeof col.index === "number"
            ? col.index
            : columnLetterToIndex(col.letter);
          colIndexes.add(idx);
        }
      }

      // Unhide each column individually
      for (const idx of colIndexes) {
        const range = sheet.getRangeByIndexes(0, idx, 1, 1);
        range.columnHidden = false;
      }

      await ctx.sync();
    });

    state.activeViews[sheetName] = null;
    await saveToWorkbook();
    renderViewsList();
    showToast("All columns shown", "success");
  } catch (err) {
    console.error("showAllColumns error:", err);
    showToast(`Error: ${err.message || "could not show all columns"}`, "error");
  }
}

/* ── Delete a View ────────────────────────────────────────────── */
async function deleteView(sheetName, viewName) {
  state.views[sheetName] = (state.views[sheetName] || []).filter(v => v.name !== viewName);
  if (state.activeViews[sheetName] === viewName) {
    state.activeViews[sheetName] = null;
  }
  await saveToWorkbook();
  renderViewsList();
  showToast(`View "${viewName}" deleted`);
}

/* ── Render Views List ────────────────────────────────────────── */
function renderViewsList() {
  const sheet = state.currentSheet;
  const sheetViews = (state.views[sheet] || []);
  const activeView = state.activeViews[sheet] || null;

  dom.viewsList.innerHTML = "";

  if (sheetViews.length === 0) {
    dom.viewsList.appendChild(dom.emptyState);
    return;
  }

  for (const view of sheetViews) {
    const isActive = view.name === activeView;
    const hiddenCount  = view.columnStates.filter(c => !c.visible).length;
    const visibleCount = view.columnStates.filter(c =>  c.visible).length;

    const card = document.createElement("div");
    card.className = `view-card${isActive ? " active-view" : ""}`;
    card.innerHTML = `
      <div class="view-card-top" title="Click to apply">
        <div class="view-indicator"></div>
        <div class="view-name">${escapeHtml(view.name)}</div>
        <div class="view-meta">${visibleCount} shown / ${hiddenCount} hidden</div>
      </div>
      <div class="view-card-actions">
        <button class="view-action-btn apply" data-action="apply">
          <svg width="11" height="11" viewBox="0 0 11 11" fill="none">
            <path d="M2 5.5h7M6.5 3l2.5 2.5L6.5 8" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/>
          </svg>
          Apply
        </button>
        <button class="view-action-btn edit" data-action="edit">
          <svg width="11" height="11" viewBox="0 0 11 11" fill="none">
            <path d="M7 2l2 2-5 5H2V7l5-5z" stroke="currentColor" stroke-width="1.3" stroke-linejoin="round"/>
          </svg>
          Edit
        </button>
        <button class="view-action-btn delete" data-action="delete">
          <svg width="11" height="11" viewBox="0 0 11 11" fill="none">
            <path d="M2 3h7M4.5 3V2h2v1M5.5 5v3M4 5l.2 3M7 5l-.2 3" stroke="currentColor" stroke-width="1.3" stroke-linecap="round"/>
          </svg>
          Delete
        </button>
      </div>`;

    // Capture name once for all handlers on this card
    const capturedName = view.name;
    const capturedSheet = sheet;

    card.querySelector(".view-card-top").addEventListener("click", () => applyView(capturedSheet, capturedName));

    // Wire each button individually rather than via forEach to avoid closure issues
    const applyBtn  = card.querySelector(".view-action-btn.apply");
    const editBtn   = card.querySelector(".view-action-btn.edit");
    const deleteBtn = card.querySelector(".view-action-btn.delete");

    applyBtn.addEventListener("click", async e => {
      e.stopPropagation();
      await applyView(capturedSheet, capturedName);
    });

    editBtn.addEventListener("click", e => {
      e.stopPropagation();
      openEditor(capturedName);
    });

    deleteBtn.addEventListener("click", async e => {
      e.stopPropagation();
      if (deleteBtn.dataset.armed === "true") {
        deleteBtn.dataset.armed = "false";
        await deleteView(capturedSheet, capturedName);
      } else {
        deleteBtn.dataset.armed = "true";
        deleteBtn.textContent = "Confirm?";
        deleteBtn.style.color = "var(--red)";
        deleteBtn.style.fontWeight = "700";
        setTimeout(() => {
          if (deleteBtn.dataset.armed === "true") {
            deleteBtn.dataset.armed = "false";
            deleteBtn.textContent = "Delete";
            deleteBtn.style.color = "";
            deleteBtn.style.fontWeight = "";
          }
        }, 3000);
      }
    });

    dom.viewsList.appendChild(card);
  }
}

function escapeHtml(str) {
  const d = document.createElement("div");
  d.textContent = str;
  return d.innerHTML;
}

/* ── Open Editor ──────────────────────────────────────────────── */
async function openEditor(viewNameToEdit = null) {
  state.editingViewName = viewNameToEdit;
  dom.editorModeLabel.textContent = viewNameToEdit ? "EDIT VIEW" : "NEW VIEW";

  // Pre-fill name if editing
  if (viewNameToEdit) {
    dom.viewNameInput.value = viewNameToEdit;
    // Restore saved header row if editing
    const existing = (state.views[state.currentSheet] || []).find(v => v.name === viewNameToEdit);
    dom.headerRowInput.value = existing?.headerRow ?? 1;
  } else {
    dom.viewNameInput.value = "";
    dom.headerRowInput.value = 1;
  }

  // Show loading
  dom.columnsList.innerHTML = `
    <div class="loading-spinner">
      <div class="spinner"></div>
      <span>Reading columns…</span>
    </div>`;
  dom.colCountLabel.textContent = "loading…";

  showPanel("editor");

  await reloadColumns(viewNameToEdit);
}

/* ── Reload columns with current header row value ────────────── */
async function reloadColumns(viewNameToEdit = state.editingViewName) {
  const headerRow = Math.max(1, parseInt(dom.headerRowInput.value, 10) || 1);

  dom.columnsList.innerHTML = `
    <div class="loading-spinner">
      <div class="spinner"></div>
      <span>Reading columns…</span>
    </div>`;
  dom.colCountLabel.textContent = "loading…";

  // Load columns
  try {
    const { defs, hiddenMap } = await readColumns(state.currentSheet, headerRow);
    state.columnDefs = defs;

    // Determine initial checked state
    let checkedMap = {};
    if (viewNameToEdit) {
      // Load from saved view
      const existing = (state.views[state.currentSheet] || []).find(v => v.name === viewNameToEdit);
      if (existing) {
        existing.columnStates.forEach(cs => { checkedMap[cs.letter] = cs.visible; });
      }
    } else {
      // Default: reflect current sheet state
      defs.forEach(d => { checkedMap[d.letter] = !hiddenMap[d.index]; });
    }

    renderColumnsList(defs, checkedMap);
    dom.colCountLabel.textContent = `${defs.length} col${defs.length !== 1 ? "s" : ""}`;
  } catch (err) {
    dom.columnsList.innerHTML = `<div class="loading-spinner" style="color:var(--red)">Failed to read columns</div>`;
    console.error(err);
  }
}

/* ── Render Columns List ──────────────────────────────────────── */
function renderColumnsList(defs, checkedMap) {
  dom.columnsList.innerHTML = "";
  for (const def of defs) {
    const checked = checkedMap[def.letter] !== false; // default visible
    const item = document.createElement("div");
    item.className = `col-item${checked ? " checked" : ""}`;
    item.dataset.letter = def.letter;
    item.innerHTML = `
      <div class="col-checkbox"></div>
      <div class="col-letter">${def.letter}</div>
      <div class="col-name">${escapeHtml(def.name !== def.letter ? def.name : "")}</div>`;

    item.addEventListener("click", () => {
      item.classList.toggle("checked");
    });

    dom.columnsList.appendChild(item);
  }
}

/* ── Get current checked states from DOM ─────────────────────── */
function getCheckedStates() {
  const items = dom.columnsList.querySelectorAll(".col-item");
  const states = [];
  items.forEach(item => {
    states.push({
      letter: item.dataset.letter,
      visible: item.classList.contains("checked"),
    });
  });
  return states;
}

/* ── Save View ────────────────────────────────────────────────── */
async function saveView() {
  const name = dom.viewNameInput.value.trim();
  if (!name) {
    showToast("Please enter a view name", "error");
    dom.viewNameInput.focus();
    return;
  }

  const sheet = state.currentSheet;
  const columnStates = getCheckedStates();

  if (!state.views[sheet]) state.views[sheet] = [];

  const existingIdx = state.views[sheet].findIndex(v => v.name === name);

  if (state.editingViewName && state.editingViewName !== name) {
    // Renamed — remove old, add new
    state.views[sheet] = state.views[sheet].filter(v => v.name !== state.editingViewName);
    if (state.activeViews[sheet] === state.editingViewName) {
      state.activeViews[sheet] = null;
    }
  }

  const headerRow = Math.max(1, parseInt(dom.headerRowInput.value, 10) || 1);
  const newView = { name, columnStates, headerRow, updatedAt: Date.now() };

  if (existingIdx !== -1) {
    state.views[sheet][existingIdx] = newView;
  } else {
    state.views[sheet].push(newView);
  }

  await saveToWorkbook();
  renderViewsList();
  showPanel("views");
  showToast(`View "${name}" saved`, "success");
}

/* ── Sheet change handling ────────────────────────────────────── */
async function onSheetChanged() {
  try {
    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");
      await ctx.sync();
      state.currentSheet = sheet.name;
    });
    dom.sheetBadge.textContent = state.currentSheet || "—";
    renderViewsList();
  } catch (err) {
    console.error("onSheetChanged:", err);
  }
}

/* ── Init ─────────────────────────────────────────────────────── */
async function init() {
  // Wire up buttons
  dom.btnNewView      .addEventListener("click", () => openEditor(null));
  dom.btnShowAll      .addEventListener("click", () => showAllColumns(state.currentSheet));
  dom.btnBack         .addEventListener("click", () => showPanel("views"));
  dom.btnCancelEditor .addEventListener("click", () => showPanel("views"));
  dom.btnSaveView     .addEventListener("click", saveView);

  dom.btnReloadCols.addEventListener("click", () => reloadColumns());
  dom.headerRowInput.addEventListener("keydown", e => { if (e.key === "Enter") reloadColumns(); });

  dom.btnCheckAll.addEventListener("click", () => {
    dom.columnsList.querySelectorAll(".col-item").forEach(i => i.classList.add("checked"));
  });
  dom.btnUncheckAll.addEventListener("click", () => {
    dom.columnsList.querySelectorAll(".col-item").forEach(i => i.classList.remove("checked"));
  });

  await Office.onReady();

  // Load saved data
  await loadFromWorkbook();

  // Get active sheet
  await Excel.run(async ctx => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    sheet.load("name");
    await ctx.sync();
    state.currentSheet = sheet.name;

    // Watch for sheet change
    ctx.workbook.worksheets.onActivated.add(async () => {
      await onSheetChanged();
    });
    await ctx.sync();
  });

  dom.sheetBadge.textContent = state.currentSheet || "—";
  renderViewsList();
  showPanel("views");
}

Office.onReady(() => init());
