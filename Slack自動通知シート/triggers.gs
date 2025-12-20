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

function setupTimeTriggers_() {
  const targets = [
    {
      handler: 'nightlyScanAndQueue',
      hour: APP_CONFIG.NIGHTLY_TRIGGER_HOUR,
      minute: APP_CONFIG.NIGHTLY_TRIGGER_NEAR_MINUTE
    },
    {
      handler: 'morningDispatch',
      hour: APP_CONFIG.MORNING_TRIGGER_HOUR,
      minute: APP_CONFIG.MORNING_TRIGGER_NEAR_MINUTE
    }
  ];

  const targetHandlers = new Set(targets.map(t => t.handler));
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getEventType() === ScriptApp.EventType.CLOCK) {
      const handler = t.getHandlerFunction();
      if (targetHandlers.has(handler)) ScriptApp.deleteTrigger(t);
    }
  });

  const allowedMinutes = new Set([0, 15, 30, 45]);

  targets.forEach(t => {
    const hour = (typeof t.hour === 'number' && t.hour >= 0 && t.hour <= 23) ? t.hour : 0;
    let builder = ScriptApp.newTrigger(t.handler).timeBased().everyDays(1).atHour(hour);
    if (allowedMinutes.has(t.minute)) builder = builder.nearMinute(t.minute);
    builder.create();
  });
}

function setupCentral_() {
  notifyUser_('初期セットアップを開始します...', 'INFO');
  try {
    initConfigSheet();
    syncTriggersFromConfig();
    setupTimeTriggers_();
    notifyUser_('初期セットアップが完了しました。', 'INFO');
  } catch (e) {
    notifyUser_(`初期セットアップに失敗しました: ${e.message}`, 'ERROR');
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Central')
    .addItem('初期セットアップ', 'setupCentral_')
    .addItem('台帳反映（トリガー同期）', 'syncTriggersFromConfig')
    .addItem('夜間スキャン→Queue', 'nightlyScanAndQueue')
    .addItem('朝の送信', 'morningDispatch')
    .addToUi();
}
