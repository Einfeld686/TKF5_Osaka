// 共通ユーティリティCode

/** 共通ユーティリティ群 **/

function getActive_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function notifyUser_(message, level) {
  const type = String(level || 'INFO').toUpperCase();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss && typeof ss.toast === 'function') ss.toast(message, type, 5);
  } catch (e) {}
  if (type === 'ERROR') {
    try {
      SpreadsheetApp.getUi().alert(message);
    } catch (e) {}
  }
  try { Logger.log(`[${type}] ${message}`); } catch (e) {}
  if (typeof console !== 'undefined') {
    const logFn = (type === 'ERROR' && console.error) ? console.error : console.log;
    if (logFn) logFn(`[${type}] ${message}`);
  }
}

function applyValidation_(sheet, headerMap, rules) {
  if (!sheet || !headerMap || !rules) return;
  const startRow = 2;
  const numRows = Math.max(0, sheet.getMaxRows() - 1);
  if (numRows <= 0) return;

  for (const header in rules) {
    if (!Object.prototype.hasOwnProperty.call(rules, header)) continue;
    const colIdx = headerMap.get(header);
    if (colIdx == null) continue;

    const range = sheet.getRange(startRow, colIdx + 1, numRows, 1);
    const rule = rules[header];

    if (rule === 'CHECKBOX') {
      const dv = SpreadsheetApp.newDataValidation().requireCheckbox().build();
      range.setDataValidation(dv);
      continue;
    }
    if (Array.isArray(rule)) {
      const dv = SpreadsheetApp.newDataValidation()
        .requireValueInList(rule, true)
        .setAllowInvalid(false)
        .build();
      range.setDataValidation(dv);
    }
  }
}

function readTable_(sheet) {
  const range = sheet.getDataRange();
  const values = range.getDisplayValues();
  if (values.length === 0) return { headerIndex: 0, headerMap: new Map(), data: [] };
  const header = values[0].map(h => String(h).trim());
  const headerMap = new Map(header.map((h, i) => [h, i]));
  const data = values.slice(1);
  return { headerIndex: 0, headerMap, data };
}

function ensureHeaders_(sheet, headers) {
  const lr = sheet.getLastRow(), lc = sheet.getLastColumn();
  if (lr === 0 || lc === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return new Map(headers.map((h, i) => [h, i]));
  }
  const cur = sheet.getRange(1, 1, 1, lc).getDisplayValues()[0];
  const map = new Map(cur.map((h, i) => [h, i]));
  const missing = headers.filter(h => !map.has(h));
  if (missing.length) {
    sheet.getRange(1, lc + 1, 1, missing.length).setValues([missing]);
  }
  const newHeader = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  return new Map(newHeader.map((h, i) => [h, i]));
}

function rowToObj_(row, headerMap) {
  const obj = {};
  for (const [k, idx] of headerMap.entries()) obj[k] = row[idx];
  return obj;
}

function normalizeBool_(v) {
  const s = String(v || '').trim().toLowerCase();
  return s === 'true' || s === '1' || s === 'yes' || s === 'y';
}

function normalizeDate_(d) {
  try {
    if (Object.prototype.toString.call(d) === '[object Date]' && !isNaN(d)) {
      const x = new Date(d); x.setHours(0,0,0,0); return x;
    }
    if (typeof d === 'number') {
      const epoch = new Date(1899, 11, 30);
      const x = new Date(epoch.getTime() + d * 86400000); x.setHours(0,0,0,0); return x;
    }
    const x = new Date(d); if (!isNaN(x)) { x.setHours(0,0,0,0); return x; }
  } catch (_) {}
  return null;
}

function formatYmd_(d, tz) {
  if (!d) return '';
  const date = (Object.prototype.toString.call(d) === '[object Date]') ? d : normalizeDate_(d);
  if (!date) return String(d);
  return Utilities.formatDate(date, tz || APP_CONFIG.TIMEZONE, 'yyyy-MM-dd');
}

function buildCellLink_(ss, sheet, a1) {
  return `${ss.getUrl()}#gid=${sheet.getSheetId()}&range=${encodeURIComponent(a1)}`;
}
