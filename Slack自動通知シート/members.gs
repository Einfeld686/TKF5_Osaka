// Members参照（氏名→Slack ID）Code

/** Members 参照（氏名→Slack ID） **/

/** 拠点ファイルの Members から氏名→Slack ID マップを構築 */
function buildSlackIdMap_(ss, membersSheetName) {
  const map = new Map();
  const sh = ss.getSheetByName(membersSheetName || SHEET_NAMES.MEMBERS_DEFAULT);
  if (!sh) return map;
  const last = sh.getLastRow();
  if (last < 2) return map;

  const needCols = Math.max(MEMBER_SHEET_COLUMNS.NAME, MEMBER_SHEET_COLUMNS.SLACK_ID);
  const values = sh.getRange(2, 1, last - 1, needCols).getDisplayValues();
  for (const row of values) {
    const name = String(row[MEMBER_SHEET_COLUMNS.NAME - 1] || '').trim();
    const sidRaw = String(row[MEMBER_SHEET_COLUMNS.SLACK_ID - 1] || '').trim();
    if (!name || !sidRaw) continue;
    const sid = sidRaw.replace(/[<@>]/g, ''); // "<@U12345>" → "U12345" に寄せる
    map.set(name, sid);
  }
  return map;
}

/** 単発ルックアップ（氏名→Slack ID） */
function lookupSlackId_(ss, membersSheetName, staffName) {
  const map = buildSlackIdMap_(ss, membersSheetName);
  return (map.get(staffName) || '');
}
