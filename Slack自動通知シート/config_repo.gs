// 台帳(Config)の読み取りCode

/** 台帳(Config)の読み取り・検索 **/

/** 台帳を行オブジェクト配列で取得 */
function cfgLoadAll() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.CONFIG);
  if (!sh) return [];
  const { headerMap, data } = readTable_(sh);
  return data.map(row => {
    const obj = rowToObj_(row, headerMap);
    obj[CONFIG_HEADERS.ENABLED] = normalizeBool_(obj[CONFIG_HEADERS.ENABLED]);
    obj[CONFIG_HEADERS.TASK_SHEET_NAME] = obj[CONFIG_HEADERS.TASK_SHEET_NAME] || SHEET_NAMES.TASKS_DEFAULT;
    obj[CONFIG_HEADERS.MEMBERS_SHEET_NAME] = obj[CONFIG_HEADERS.MEMBERS_SHEET_NAME] || SHEET_NAMES.MEMBERS_DEFAULT;
    obj[CONFIG_HEADERS.MODE] = obj[CONFIG_HEADERS.MODE] || 'central';
    obj[CONFIG_HEADERS.NOTIFY_TIMING] = obj[CONFIG_HEADERS.NOTIFY_TIMING] || 'immediate';
    return obj;
  });
}

/** スプレッドシートIDで台帳行を検索 */
function cfgFindBySpreadsheetId(spreadsheetId) {
  const all = cfgLoadAll().filter(r => r[CONFIG_HEADERS.ENABLED]);
  for (const r of all) {
    const id = extractSpreadsheetId_(r[CONFIG_HEADERS.SPREADSHEET_URL] || '');
    if (id === spreadsheetId) return r;
  }
  return null;
}

/** URL or ID から ID を抽出 */
function extractSpreadsheetId_(urlOrId) {
  if (!urlOrId) return '';
  if (/^[a-zA-Z0-9-_]+$/.test(urlOrId)) return urlOrId;
  const m = String(urlOrId).match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  return m ? m[1] : '';
}

/**
 * Configシートの初期化・修復
 * - ヘッダーの不足分を追加
 * - 入力規則の適用
 * - 列幅の調整
 */
function initConfigSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAMES.CONFIG);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAMES.CONFIG);
  }

  // ヘッダー定義から配列を作成
  const headers = Object.values(CONFIG_HEADERS);
  const lcBefore = sh.getLastColumn();
  const existingHeaders = (sh.getLastRow() >= 1 && lcBefore > 0)
    ? sh.getRange(1, 1, 1, lcBefore).getDisplayValues()[0].map(h => String(h).trim())
    : [];
  const missingHeaders = headers.filter(h => existingHeaders.indexOf(h) === -1);
  
  // ヘッダー確保
  const headerMap = ensureHeaders_(sh, headers);
  
  // 入力規則の適用ルール作成
  const rules = {};
  rules[CONFIG_HEADERS.ENABLED] = 'CHECKBOX';
  rules[CONFIG_HEADERS.MODE] = CONFIG_VALIDATIONS.MODE;
  rules[CONFIG_HEADERS.NOTIFY_TIMING] = CONFIG_VALIDATIONS.NOTIFY_TIMING;
  
  // 入力規則適用
  applyValidation_(sh, headerMap, rules);
  
  // 列幅調整（列を新規追加した場合のみデフォルト幅を設定）
  if (missingHeaders.length) {
    const DEFAULT_WIDTH = 120;
    const WIDE_WIDTH = 300;
    for (const header of missingHeaders) {
      const col = headerMap.get(header);
      if (col == null) continue;
      const width = (header === CONFIG_HEADERS.SPREADSHEET_URL || header === CONFIG_HEADERS.SLACK_WEBHOOK)
        ? WIDE_WIDTH
        : DEFAULT_WIDTH;
      sh.setColumnWidth(col + 1, width);
    }
  }

  return sh;
}

/**
 * Configシートのステータス列を更新する
 * @param {string} spreadsheetIdOrUrl - 対象のSpreadsheet ID もしくは URL
 * @param {string} status - 'OK' | 'ERROR' 等
 * @param {string} msg - 詳細メッセージ
 */
function updateConfigStatus_(spreadsheetIdOrUrl, status, msg) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.CONFIG);
  if (!sh) return;
  
  const { headerMap, data } = readTable_(sh);
  if (!headerMap.has(CONFIG_HEADERS.STATUS) || !headerMap.has(CONFIG_HEADERS.LAST_UPDATED) || !headerMap.has(CONFIG_HEADERS.SPREADSHEET_URL)) return;
  
  const statusCol = headerMap.get(CONFIG_HEADERS.STATUS) + 1;
  const updatedCol = headerMap.get(CONFIG_HEADERS.LAST_UPDATED) + 1;
  const urlColIdx = headerMap.get(CONFIG_HEADERS.SPREADSHEET_URL);
  const targetId = extractSpreadsheetId_(spreadsheetIdOrUrl);

  // 行を探す
  for (let i = 0; i < data.length; i++) {
    const rowUrl = data[i][urlColIdx];
    const rowId = extractSpreadsheetId_(rowUrl);
    
    const matchById = targetId && rowId === targetId;
    const matchByUrl = !targetId && rowUrl === spreadsheetIdOrUrl;
    
    if (matchById || matchByUrl) {
      const rowNum = i + 2; // header + 1-based
      const ts = Utilities.formatDate(new Date(), APP_CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
      
      sh.getRange(rowNum, statusCol).setValue(`${status}: ${msg}`);
      sh.getRange(rowNum, updatedCol).setValue(ts);
      break; // 1行見つけたら終了
    }
  }
}
