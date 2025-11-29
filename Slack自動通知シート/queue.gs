// 夜間スキャン→Queue化、朝の送信Code

/** 夜間スキャン→Queue化、朝の送信（軽処理） **/

function nightlyScanAndQueue() {
  const configs = cfgLoadAll().filter(r => r[CONFIG_HEADERS.ENABLED]);
  const today0 = new Date(); today0.setHours(0,0,0,0);
  const queueSheet = getActive_().getSheetByName(SHEET_NAMES.QUEUE) || getActive_().insertSheet(SHEET_NAMES.QUEUE);
  const qHeaderMap = ensureHeaders_(queueSheet, Object.values(QUEUE_HEADERS));
  const existingKeys = new Set(readQueueKeys_()); // 重複防止

  let processedCells = 0;

  for (const cfg of configs) {
    if (cfg[CONFIG_HEADERS.NOTIFY_TIMING] === 'immediate') continue; // 朝配信不要の拠点はスキップ可

    const spreadsheetId = extractSpreadsheetId_(cfg[CONFIG_HEADERS.SPREADSHEET_URL]);
    if (!spreadsheetId) continue;
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const tasks = ss.getSheetByName(cfg[CONFIG_HEADERS.TASK_SHEET_NAME]);
    if (!tasks) continue;

    // メタ行
    const lastCol = tasks.getLastColumn();
    const lastRow = tasks.getLastRow();
    if (lastCol < TASK_SHEET_LAYOUT.TASK_START_COLUMN || lastRow < TASK_SHEET_LAYOUT.STAFF_START_ROW) continue;

    const taskNames = tasks.getRange(TASK_SHEET_LAYOUT.TASK_NAME_ROW, TASK_SHEET_LAYOUT.TASK_START_COLUMN, 1, lastCol - TASK_SHEET_LAYOUT.TASK_START_COLUMN + 1).getDisplayValues()[0];
    const deadlines = tasks.getRange(TASK_SHEET_LAYOUT.DEADLINE_ROW, TASK_SHEET_LAYOUT.TASK_START_COLUMN, 1, lastCol - TASK_SHEET_LAYOUT.TASK_START_COLUMN + 1).getValues()[0];
    const fileUrls  = tasks.getRange(TASK_SHEET_LAYOUT.FILE_URL_ROW, TASK_SHEET_LAYOUT.TASK_START_COLUMN, 1, lastCol - TASK_SHEET_LAYOUT.TASK_START_COLUMN + 1).getDisplayValues()[0];

    const staffNames = tasks.getRange(TASK_SHEET_LAYOUT.STAFF_START_ROW, TASK_SHEET_LAYOUT.STAFF_NAME_COLUMN, lastRow - TASK_SHEET_LAYOUT.STAFF_START_ROW + 1, 1).getDisplayValues().map(r => String(r[0] || '').trim());

    const checkboxValues = tasks.getRange(
      TASK_SHEET_LAYOUT.STAFF_START_ROW, TASK_SHEET_LAYOUT.CHECKBOX_START_COLUMN,
      lastRow - TASK_SHEET_LAYOUT.STAFF_START_ROW + 1,
      lastCol - TASK_SHEET_LAYOUT.CHECKBOX_START_COLUMN + 1
    ).getValues();

    const slackMap = buildSlackIdMap_(ss, cfg[CONFIG_HEADERS.MEMBERS_SHEET_NAME]);

    for (let r = 0; r < staffNames.length; r++) {
      const staff = staffNames[r]; if (!staff) continue;
      for (let c = 0; c < taskNames.length; c++) {
        processedCells++;
        if (processedCells > APP_CONFIG.NIGHTLY_SCAN_CELL_LIMIT) return; // チャンク上限で早期終了（次回へ）

        const taskName = (taskNames[c] || '').toString().trim();
        if (!taskName) continue;

        const done = (checkboxValues[r][c] === true || checkboxValues[r][c] === 'TRUE');
        if (done) continue;

        const dl = normalizeDate_(deadlines[c]);
        if (!dl) continue;

        if (dl.getTime() - today0.getTime() > 0) continue; // 当日以降のみ対象

        const row = TASK_SHEET_LAYOUT.STAFF_START_ROW + r;
        const col = TASK_SHEET_LAYOUT.CHECKBOX_START_COLUMN + c;
        const a1 = tasks.getRange(row, col).getA1Notation();

        const key = [
          cfg[CONFIG_HEADERS.SITE_CODE],
          tasks.getSheetId(), row, col,
          taskName, formatYmd_(dl, APP_CONFIG.TIMEZONE)
        ].join('|');

        if (existingKeys.has(key)) continue;

        const rec = {};
        rec[QUEUE_HEADERS.KEY] = key;
        rec[QUEUE_HEADERS.STATUS] = 'PENDING';
        rec[QUEUE_HEADERS.SITE_CODE] = cfg[CONFIG_HEADERS.SITE_CODE];
        rec[QUEUE_HEADERS.SPREADSHEET_ID] = spreadsheetId;
        rec[QUEUE_HEADERS.SHEET_NAME] = tasks.getName();
        rec[QUEUE_HEADERS.STAFF_NAME] = staff;
        rec[QUEUE_HEADERS.STAFF_SLACK_ID] = slackMap.get(staff) || '';
        rec[QUEUE_HEADERS.TASK_NAME] = taskName;
        rec[QUEUE_HEADERS.DEADLINE_ISO] = dl.toISOString();
        rec[QUEUE_HEADERS.FILE_URL] = (fileUrls[c] || '').toString().trim();
        rec[QUEUE_HEADERS.CELL_A1] = a1;
        rec[QUEUE_HEADERS.CELL_LINK] = buildCellLink_(ss, tasks, a1);
        rec[QUEUE_HEADERS.REASON] = 'overdue_or_today';
        rec[QUEUE_HEADERS.WEBHOOK] = cfg[CONFIG_HEADERS.SLACK_WEBHOOK];
        rec[QUEUE_HEADERS.TEMPLATE_ID] = 'overdue';
        rec[QUEUE_HEADERS.CREATED_AT] = new Date().toISOString();

        // 末尾に追加
        const values = Object.values(QUEUE_HEADERS).map(h => rec[h] || '');
        const lr = queueSheet.getLastRow();
        if (lr === 0) ensureHeaders_(queueSheet, Object.values(QUEUE_HEADERS));
        queueSheet.getRange(lr + 1 === 1 ? 2 : lr + 1, 1, 1, values.length).setValues([values]);

        existingKeys.add(key);
      }
    }
  }
}

/** 朝の送信（QueueからPENDINGを送るだけ） */
function morningDispatch() {
  const qsh = getActive_().getSheetByName(SHEET_NAMES.QUEUE);
  if (!qsh) return;
  const { headerMap, data } = readTable_(qsh);

  const idxStatus = headerMap.get(QUEUE_HEADERS.STATUS);
  const idxWebhook = headerMap.get(QUEUE_HEADERS.WEBHOOK);

  let sent = 0;
  for (let i = 0; i < data.length; i++) {
    if (sent >= APP_CONFIG.DISPATCH_LIMIT_PER_RUN) break;

    const row = data[i];
    if (row[idxStatus] !== 'PENDING') continue;

    const rec = {};
    for (const [k, idx] of headerMap.entries()) rec[k] = row[idx];

    const webhook = rec[QUEUE_HEADERS.WEBHOOK];
    if (!webhook) continue;

    const payload = buildOverduePayload_(rec);
    const ok = postToSlack_(webhook, payload);
    if (ok) {
      // シート上で更新（status/sent_at）
      qsh.getRange(i + 2, idxStatus + 1).setValue('SENT');
      const idxSentAt = headerMap.get(QUEUE_HEADERS.SENT_AT);
      if (idxSentAt != null) qsh.getRange(i + 2, idxSentAt + 1).setValue(new Date());

      // Outboxにも記録（任意：必要に応じて）
      logOutbox_(rec);

      sent++;
      Utilities.sleep(APP_CONFIG.RATE_SLEEP_MS);
    }
  }
}

/** Outboxへの記録（Queueヘッダを共用して保存） */
function logOutbox_(rec) {
  const osh = getActive_().getSheetByName(SHEET_NAMES.OUTBOX) || getActive_().insertSheet(SHEET_NAMES.OUTBOX);
  const map = ensureHeaders_(osh, Object.values(QUEUE_HEADERS));
  const row = Object.values(QUEUE_HEADERS).map(h => (h === QUEUE_HEADERS.SENT_AT ? new Date().toISOString() : (rec[h] || '')));
  const lr = osh.getLastRow();
  osh.getRange(lr + 1 === 1 ? 2 : lr + 1, 1, 1, row.length).setValues([row]);
}

/** 既存キーの取得（重複予約防止） */
function readQueueKeys_() {
  const sh = getActive_().getSheetByName(SHEET_NAMES.QUEUE);
  if (!sh || sh.getLastRow() < 2) return [];
  const { headerMap, data } = readTable_(sh);
  const idx = headerMap.get(QUEUE_HEADERS.KEY);
  return data.map(r => r[idx]).filter(Boolean);
}
