// ─── Halda Training Checklist — Apps Script ───────────────────────────────────
// Paste this entire file into your Apps Script project (replace existing code),
// then deploy a new version of the Web app if prompted.

const SHEET_ID   = '1jeuqHrKuEgcZsDhYvB5ezD7tE4v9U0ygbjwV2k3MoLc';
const SHEET_NAME = 'Halda Training Checklist Data';

const COL = {
  CATEGORY:    0,
  DOT:         1,
  TASK:        2,
  SUBTASK:     3,
  URL:         4,
  GROUP_LABEL: 5,
  CHECKED:     6,
  ROW_ID:      7,
};

/** Cached GET payload (JSON string of { ok: true, data }). Max ~100 KB in CacheService. */
const CACHE_KEY = 'halda_training_get_v1';
const CACHE_TTL_SEC = 120;

function clearTrainingCache() {
  CacheService.getScriptCache().remove(CACHE_KEY);
}

// ─── GET ──────────────────────────────────────────────────────────────────────
// Supports JSONP via ?callback=fnName — works from file://, localhost, GitHub Pages
function doGet(e) {
  const cache = CacheService.getScriptCache();
  const callback = e && e.parameter && e.parameter.callback;

  const cached = cache.get(CACHE_KEY);
  if (cached) {
    if (callback) {
      return ContentService
        .createTextOutput(callback + '(' + cached + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(cached).setMimeType(ContentService.MimeType.JSON);
  }

  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const rows  = sheet.getDataRange().getValues();
  const data  = rows.slice(1).map(r => ({
    category:   r[COL.CATEGORY],
    dot:        r[COL.DOT],
    task:       r[COL.TASK],
    subtask:    r[COL.SUBTASK],
    url:        r[COL.URL],
    groupLabel: r[COL.GROUP_LABEL],
    checked:    r[COL.CHECKED] === true || String(r[COL.CHECKED]).toUpperCase() === 'TRUE',
    rowId:      String(r[COL.ROW_ID]),
  }));

  const payload = JSON.stringify({ ok: true, data });
  try {
    cache.put(CACHE_KEY, payload, CACHE_TTL_SEC);
  } catch (ignore) {
    // CacheService entry limit is ~100 KB; GET still succeeds without caching.
  }

  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + payload + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(payload)
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── POST ─────────────────────────────────────────────────────────────────────
function doPost(e) {
  const payload = JSON.parse(e.postData.contents);
  const { action } = payload;
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  try {
    if (action === 'set_checked') return setChecked(sheet, payload);
    if (action === 'add_row')     return addRow(sheet, payload);
    if (action === 'update_row')  return updateRow(sheet, payload);
    if (action === 'delete_row')  return deleteRow(sheet, payload);
    if (action === 'rename_category') return renameCategory(sheet, payload);
    return json({ ok: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return json({ ok: false, error: err.message });
  }
}

// ─── Helpers ──────────────────────────────────────────────────────────────────
function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Ensures pending spreadsheet writes finish before the web app responds (helps other clients reload). */
function flushSpreadsheet_() {
  try {
    SpreadsheetApp.flush();
  } catch (ignore) {}
}

/** Only reads the ROW_ID column — faster than getDataRange().getValues() for lookups. */
function findRow(sheet, rowId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;

  const ids = sheet.getRange(2, COL.ROW_ID + 1, lastRow, COL.ROW_ID + 1).getValues();
  const want = String(rowId);
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === want) return i + 2;
  }
  return -1;
}

function nextRowId(sheet) {
  const lastRow = sheet.getLastRow();
  let max = 0;
  if (lastRow >= 2) {
    const ids = sheet.getRange(2, COL.ROW_ID + 1, lastRow, COL.ROW_ID + 1).getValues();
    ids.forEach(function (r) {
      const n = parseInt(String(r[0]).replace(/\D/g, ''), 10);
      if (!isNaN(n) && n > max) max = n;
    });
  }
  return 'r' + (max + 1);
}

// ─── Actions ──────────────────────────────────────────────────────────────────
function setChecked(sheet, payload) {
  const rowId = payload.rowId;
  const checked = payload.checked;
  const row = findRow(sheet, rowId);
  if (row < 0) return json({ ok: false, error: 'Row not found: ' + rowId });
  sheet.getRange(row, COL.CHECKED + 1).setValue(checked);
  flushSpreadsheet_();
  clearTrainingCache();
  return json({ ok: true });
}

function addRow(sheet, payload) {
  const category = payload.category;
  const dot = payload.dot;
  const task = payload.task;
  const subtask = payload.subtask;
  const url = payload.url;
  const groupLabel = payload.groupLabel;

  const id = nextRowId(sheet);
  const newRow = [category, dot, task, subtask || '', url || '', groupLabel || '', false, id];
  sheet.appendRow(newRow);
  flushSpreadsheet_();
  clearTrainingCache();
  return json({ ok: true, rowId: id });
}

function updateRow(sheet, payload) {
  const rowId = payload.rowId;
  const fields = payload.fields;
  const row = findRow(sheet, rowId);
  if (row < 0) return json({ ok: false, error: 'Row not found: ' + rowId });
  const colMap = {
    category: COL.CATEGORY,
    dot: COL.DOT,
    task: COL.TASK,
    subtask: COL.SUBTASK,
    url: COL.URL,
    groupLabel: COL.GROUP_LABEL,
    checked: COL.CHECKED,
  };
  Object.keys(fields).forEach(function (key) {
    if (key in colMap) sheet.getRange(row, colMap[key] + 1).setValue(fields[key]);
  });
  flushSpreadsheet_();
  clearTrainingCache();
  return json({ ok: true });
}

function deleteRow(sheet, payload) {
  const rowId = payload.rowId;
  const row = findRow(sheet, rowId);
  if (row < 0) return json({ ok: false, error: 'Row not found: ' + rowId });
  sheet.deleteRow(row);
  flushSpreadsheet_();
  clearTrainingCache();
  return json({ ok: true });
}

function renameCategory(sheet, payload) {
  const oldCategory = String(payload.oldCategory || '');
  const newCategory = String(payload.newCategory || '');
  if (!oldCategory || !newCategory) {
    return json({ ok: false, error: 'rename_category requires oldCategory and newCategory' });
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    flushSpreadsheet_();
    clearTrainingCache();
    return json({ ok: true });
  }
  const colVals = sheet.getRange(2, COL.CATEGORY + 1, lastRow, COL.CATEGORY + 1).getValues();
  for (let i = 0; i < colVals.length; i++) {
    if (String(colVals[i][0]) === oldCategory) {
      sheet.getRange(i + 2, COL.CATEGORY + 1).setValue(newCategory);
    }
  }
  flushSpreadsheet_();
  clearTrainingCache();
  return json({ ok: true });
}
