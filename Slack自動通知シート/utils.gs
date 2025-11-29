// 共通ユーティリティCode

/** 共通ユーティリティ群 **/

function getActive_() { return SpreadsheetApp.getActiveSpreadsheet(); }

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

/**
 * 指定列に入力規則を適用する
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Map<string, number>} headerMap
 * @param {Object} rules - { HEADER_NAME: ['Option1', 'Option2'] } or { HEADER_NAME: 'CHECKBOX' }
 */
function applyValidation_(sheet, headerMap, rules) {
  const maxRows = sheet.getMaxRows();
  if (maxRows <= 1) return; // データ行がない、またはヘッダーのみの場合はスキップ
  
  // データ領域は2行目から最後まで
  const numRows = maxRows - 1;
  
  for (const [header, rule] of Object.entries(rules)) {
    if (!headerMap.has(header)) continue;
    const colIdx = headerMap.get(header) + 1; // 1-based
    const range = sheet.getRange(2, colIdx, numRows, 1);
    
    let validation = null;
    if (rule === 'CHECKBOX') {
      validation = SpreadsheetApp.newDataValidation()
        .requireCheckbox()
        .setAllowInvalid(false)
        .build();
    } else if (Array.isArray(rule)) {
      validation = SpreadsheetApp.newDataValidation()
        .requireValueInList(rule, true) // true = dropdown
        .setAllowInvalid(false)
        .build();
    }
    
    if (validation) {
      range.setDataValidation(validation);
    }
  }
}

/**
 * ユーザーへの通知（Toast / Alert）
 * @param {string} message
 * @param {string} level - 'INFO' | 'WARN' | 'ERROR'
 */
function notifyUser_(message, level = 'INFO') {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    console.log(`[${level}] ${message}`);
    return;
  }
  
  // Toastは常に表示（控えめに）
  const title = level === 'ERROR' ? 'エラー' : (level === 'WARN' ? '注意' : '通知');
  // エラー時は長めに表示
  const seconds = level === 'ERROR' ? 10 : 5;
  ss.toast(message, title, seconds);
  
  // ERRORの場合はダイアログも出して気づかせる
  if (level === 'ERROR') {
    try {
      SpreadsheetApp.getUi().alert(`【エラー】\n${message}`);
    } catch (e) {
      // UIが取得できない場合（トリガー実行時など）は無視
      console.warn('Cannot show alert dialog:', e);
    }
  }
  
  console.log(`[${level}] ${message}`);
}

