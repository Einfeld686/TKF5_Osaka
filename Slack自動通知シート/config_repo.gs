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
  
  // ヘッダー確保
  const headerMap = ensureHeaders_(sh, headers);
  
  // 入力規則の適用ルール作成
  const rules = {};
  rules[CONFIG_HEADERS.ENABLED] = 'CHECKBOX';
  rules[CONFIG_HEADERS.MODE] = CONFIG_VALIDATIONS.MODE;
  rules[CONFIG_HEADERS.NOTIFY_TIMING] = CONFIG_VALIDATIONS.NOTIFY_TIMING;
  
  // 入力規則適用
  applyValidation_(sh, headerMap, rules);
  
  // 列幅調整（任意）
  sh.setColumnWidths(1, headers.length, 120);
  // URL列などは広めに
  if (headerMap.has(CONFIG_HEADERS.SPREADSHEET_URL)) {
    sh.setColumnWidth(headerMap.get(CONFIG_HEADERS.SPREADSHEET_URL) + 1, 300);
  }
  if (headerMap.has(CONFIG_HEADERS.SLACK_WEBHOOK)) {
    sh.setColumnWidth(headerMap.get(CONFIG_HEADERS.SLACK_WEBHOOK) + 1, 300);
  }

  return sh;
}

/**
 * Configシートのステータス列を更新する
 * @param {string} spreadsheetId - 対象のSpreadsheet ID (URLから抽出したもの)
 * @param {string} status - 'OK' | 'ERROR' 等
 * @param {string} msg - 詳細メッセージ
 */
function updateConfigStatus_(spreadsheetId, status, msg) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.CONFIG);
  if (!sh) return;
  
  const { headerMap, data } = readTable_(sh);
  if (!headerMap.has(CONFIG_HEADERS.STATUS) || !headerMap.has(CONFIG_HEADERS.LAST_UPDATED)) return;
  
  const statusCol = headerMap.get(CONFIG_HEADERS.STATUS) + 1;
  const updatedCol = headerMap.get(CONFIG_HEADERS.LAST_UPDATED) + 1;
  const urlColIdx = headerMap.get(CONFIG_HEADERS.SPREADSHEET_URL);

  // 行を探す
  for (let i = 0; i < data.length; i++) {
    const rowUrl = data[i][urlColIdx];
    const rowId = extractSpreadsheetId_(rowUrl);
    
    if (rowId === spreadsheetId) {
      const rowNum = i + 2; // header + 1-based
      const ts = Utilities.formatDate(new Date(), APP_CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
      
      sh.getRange(rowNum, statusCol).setValue(`${status}: ${msg}`);
      sh.getRange(rowNum, updatedCol).setValue(ts);
      break; // 1行見つけたら終了
    }
  }
}

