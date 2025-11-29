// onEditトリガー同期＆メニューCode

/** onEdit（インストール型）トリガーの台帳同期＆メニュー **/

function syncTriggersFromConfig() {
  const rows = cfgLoadAll().filter(r => r[CONFIG_HEADERS.ENABLED]);
  const want = new Set(rows.map(r => extractSpreadsheetId_(r[CONFIG_HEADERS.SPREADSHEET_URL])));

  const trigs = ScriptApp.getProjectTriggers();
  const map = new Map();
  trigs.forEach(t => {
    if (t.getEventType() === ScriptApp.EventType.ON_EDIT && t.getTriggerSourceId()) {
      map.set(t.getTriggerSourceId(), t);
    }
  });

  // 追加
  for (const id of want) {
    if (!map.has(id)) {
      ScriptApp.newTrigger('handleEdit').forSpreadsheet(id).onEdit().create();
    }
  }
  // 削除
  for (const [srcId, trig] of map) {
    if (!want.has(srcId)) ScriptApp.deleteTrigger(trig);
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Central')
    .addItem('台帳反映（トリガー同期）', 'syncTriggersFromConfig')
    .addItem('夜間スキャン→Queue', 'nightlyScanAndQueue')
    .addItem('朝の送信', 'morningDispatch')
    .addToUi();
}
