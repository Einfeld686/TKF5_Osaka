// 定数・列定義・運用パラメータCode

/*** ユーザー指定の定義（1-based index） ***/
const MEMBER_SHEET_COLUMNS = {
  NAME: 3,
  SLACK_ID: 2,
  ROLE: 5
};
const TASK_SHEET_LAYOUT = {
  STAFF_NAME_COLUMN: 4,
  TASK_START_COLUMN: 5,
  CHECKBOX_START_COLUMN: 5,
  TASK_NAME_ROW: 4,
  DEADLINE_ROW: 5,
  FILE_URL_ROW: 6,
  STAFF_START_ROW: 7
};

/*** シート名・ヘッダ名（必要に応じて調整） ***/
const SHEET_NAMES = {
  CONFIG: 'Config',
  QUEUE: 'Queue',
  OUTBOX: 'Outbox',
  TASKS_DEFAULT: 'Tasks',
  MEMBERS_DEFAULT: 'Members'
};

// Configヘッダ（台帳の1行目にこの見出しを置く）
const CONFIG_HEADERS = {
  ENABLED: 'enabled',                  // TRUE/FALSE
  SITE_CODE: 'site_code',              // 例: TOKYO
  SPREADSHEET_URL: 'spreadsheet_url',  // 各拠点のURL（セル入力）
  TASK_SHEET_NAME: 'task_sheet',       // 未指定なら Tasks
  MEMBERS_SHEET_NAME: 'members_sheet', // 未指定なら Members
  SLACK_WEBHOOK: 'slack_webhook',      // 平文でOK（方針どおり）
  MODE: 'mode',                        // 'central' | 'local' | 'none'
  NOTIFY_TIMING: 'notify_timing',      // 'immediate' | 'morning' | 'both'
  STATUS: 'status',                    // [NEW] 実行結果/エラー
  LAST_UPDATED: 'last_updated'         // [NEW] 最終実行日時
};

// [NEW] Configシートの入力規則定義
const CONFIG_VALIDATIONS = {
  MODE: ['central', 'local', 'none'],
  NOTIFY_TIMING: ['immediate', 'morning', 'both']
};

// Queueヘッダ（Queueシートの1行目）
const QUEUE_HEADERS = {
  KEY: 'key',
  STATUS: 'status',            // PENDING | SENT | SKIP
  SITE_CODE: 'site_code',
  SPREADSHEET_ID: 'spreadsheet_id',
  SHEET_NAME: 'sheet_name',
  STAFF_NAME: 'staff_name',
  STAFF_SLACK_ID: 'staff_slack_id',
  TASK_NAME: 'task_name',
  DEADLINE_ISO: 'deadline_iso',
  FILE_URL: 'file_url',
  CELL_A1: 'cell_a1',
  CELL_LINK: 'cell_link',
  REASON: 'reason',            // overdue_or_today など
  WEBHOOK: 'slack_webhook',
  TEMPLATE_ID: 'template_id',  // basic | overdue 等
  CREATED_AT: 'created_at',
  SENT_AT: 'sent_at'
};

/*** 運用パラメータ（時間超過やレート制限対策） ***/
const APP_CONFIG = {
  TIMEZONE: 'Asia/Tokyo',
  DISPATCH_LIMIT_PER_RUN: 80,  // 朝に送る最大件数/実行
  RATE_SLEEP_MS: 150,          // Slack送信間隔(軽い間引き)
  NIGHTLY_SCAN_CELL_LIMIT: 5000 // 夜間スキャンの最大セル数/実行（ざっくり上限）
};
