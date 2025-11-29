// 即時通知ハンドラ（onEdit）Code

/** 即時通知ハンドラ（中央集約：onEdit実行） **/

function handleEdit(e) {
  try {
    if (!e || !e.range || !e.source) return;

    const spreadsheetId = e.source.getId();
    const cfg = cfgFindBySpreadsheetId(spreadsheetId);
    if (!cfg) return;
    if (cfg[CONFIG_HEADERS.MODE] === 'local') return; // 即時は拠点側に任せる設定
    if (cfg[CONFIG_HEADERS.NOTIFY_TIMING] === 'morning') return; // 朝のみ

    const sheet = e.range.getSheet();
    const targetSheetName = cfg[CONFIG_HEADERS.TASK_SHEET_NAME];
    if (sheet.getName() !== targetSheetName) return;

    const row = e.range.getRow();
    const col = e.range.getColumn();
    if (row < TASK_SHEET_LAYOUT.STAFF_START_ROW) return;
    if (col < TASK_SHEET_LAYOUT.CHECKBOX_START_COLUMN) return;

    // チェックボックス変更のみ対象
    const newVal = e.value;
    const isCheckboxEdit = (newVal === 'TRUE' || newVal === 'FALSE' || newVal === undefined);
    if (!isCheckboxEdit) return;

    // 文脈抽出
    const staffName = sheet.getRange(row, TASK_SHEET_LAYOUT.STAFF_NAME_COLUMN).getDisplayValue().trim();
    const taskName  = sheet.getRange(TASK_SHEET_LAYOUT.TASK_NAME_ROW, col).getDisplayValue().trim();
    const deadline  = sheet.getRange(TASK_SHEET_LAYOUT.DEADLINE_ROW, col).getValue();
    const fileUrl   = sheet.getRange(TASK_SHEET_LAYOUT.FILE_URL_ROW, col).getDisplayValue().trim();
    const a1        = sheet.getRange(row, col).getA1Notation();
    const link      = buildCellLink_(e.source, sheet, a1);

    // Slack ID 引き
    const slackId   = lookupSlackId_(e.source, cfg[CONFIG_HEADERS.MEMBERS_SHEET_NAME], staffName);
    const checked   = (newVal === 'TRUE');
    const deadlineText = formatYmd_(deadline, APP_CONFIG.TIMEZONE);

    const payload = buildImmediatePayload_(cfg[CONFIG_HEADERS.SITE_CODE], {
      staffName, slackId, taskName, deadlineText, fileUrl, cellLink: link, checked
    });

    const webhook = cfg[CONFIG_HEADERS.SLACK_WEBHOOK];
    if (webhook) postToSlack_(webhook, payload);

  } catch (err) {
    console.error('[handleEdit] ', err);
  }
}
