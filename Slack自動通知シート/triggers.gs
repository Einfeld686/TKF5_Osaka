// onEditトリガー同期＆メニューCode

/** onEdit（インストール型）トリガーの台帳同期＆メニュー **/

function syncTriggersFromConfig() {
  notifyUser_('トリガー同期を開始します...', 'INFO');
  try {
    const rows = cfgLoadAll().filter(r => r[CONFIG_HEADERS.ENABLED]);
    const want = new Set(rows.map(r => extractSpreadsheetId_(r[CONFIG_HEADERS.SPREADSHEET_URL])));

    const trigs = ScriptApp.getProjectTriggers();
    const map = new Map();
    trigs.forEach(t => {
      if (t.getEventType() === ScriptApp.EventType.ON_EDIT && t.getTriggerSourceId()) {
        map.set(t.getTriggerSourceId(), t);
      }
    });

    let added = 0;
    let removed = 0;

    // 追加
    for (const id of want) {
      if (!map.has(id)) {
        ScriptApp.newTrigger('handleEdit').forSpreadsheet(id).onEdit().create();
        added++;
        updateConfigStatus_(id, 'OK', 'Trigger created');
      } else {
        // 既存OK
        updateConfigStatus_(id, 'OK', 'Trigger active');
      }
    }
    // 削除
    for (const [srcId, trig] of map) {
      if (!want.has(srcId)) {
        ScriptApp.deleteTrigger(trig);
        removed++;
        // 削除されたIDはConfigにないかもしれないが、あれば更新
        updateConfigStatus_(srcId, 'DISABLED', 'Trigger removed');
      }
    }
    
    notifyUser_(`同期完了: 追加=${added}, 削除=${removed}`, 'INFO');
  } catch (e) {
    notifyUser_(`同期失敗: ${e.message}`, 'ERROR');
  }
}

function onOpen() {
  // 自動でConfigシートの検証・修復を行う（サイレント実行）
  try {
    initConfigSheet();
  } catch (e) {
    console.warn('Auto-init failed:', e);
  }

  SpreadsheetApp.getUi()
    .createMenu('Central')
    .addItem('台帳反映（トリガー同期）', 'syncTriggersFromConfig')
    .addItem('夜間スキャン→Queue', 'nightlyScanAndQueue')
    .addItem('朝の送信', 'morningDispatch')
    .addSeparator()
    .addItem('設定シートの検証・修復', 'manualInitConfig')
    .addToUi();
}

/** 手動実行用のラッパー */
function manualInitConfig() {
  try {
    initConfigSheet();
    notifyUser_('設定シートの検証・修復が完了しました', 'INFO');
  } catch (e) {
    notifyUser_(`検証失敗: ${e.message}`, 'ERROR');
  }
}
