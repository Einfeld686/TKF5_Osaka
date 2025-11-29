// Slack投稿＆メッセージ生成Code

/** Slack 投稿とメッセージ生成 **/

function postToSlack_(webhookUrl, payload) {
  try {
    const res = UrlFetchApp.fetch(webhookUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    if (code >= 300) {
      console.error('Slack error', code, res.getContentText());
      return false;
    }
    return true;
  } catch (e) {
    console.error('Slack fetch error', e);
    return false;
  }
}

function buildImmediatePayload_(siteCode, { staffName, slackId, taskName, deadlineText, fileUrl, cellLink, checked }) {
  const who = slackId ? `<@${slackId}>` : staffName;
  const statusEmoji = checked ? ':white_check_mark:' : ':white_large_square:';
  const title = `${statusEmoji} [${siteCode}] ${taskName}`;

  const fields = [
    { type: 'mrkdwn', text: `*担当*\n${who}` },
    { type: 'mrkdwn', text: `*期限*\n${deadlineText || '-'}` }
  ];

  const blocks = [
    { type: 'section', text: { type: 'mrkdwn', text: `*${title}*` } },
    { type: 'section', fields }
  ];

  const actions = [];
  if (cellLink) actions.push({ type: 'button', text: { type: 'plain_text', text: 'セルを開く' }, url: cellLink });
  if (fileUrl)  actions.push({ type: 'button', text: { type: 'plain_text', text: '関連ファイル' }, url: fileUrl });
  if (actions.length) blocks.push({ type: 'actions', elements: actions });

  return { text: title, blocks };
}

function buildOverduePayload_(rec) {
  const who = rec[QUEUE_HEADERS.STAFF_SLACK_ID] ? `<@${rec[QUEUE_HEADERS.STAFF_SLACK_ID]}>` : rec[QUEUE_HEADERS.STAFF_NAME];
  const deadlineYmd = (rec[QUEUE_HEADERS.DEADLINE_ISO] || '').slice(0, 10);
  const title = `:alarm_clock: [${rec[QUEUE_HEADERS.SITE_CODE]}] 期限到来/超過: ${rec[QUEUE_HEADERS.TASK_NAME]}`;

  const fields = [
    { type: 'mrkdwn', text: `*担当*\n${who}` },
    { type: 'mrkdwn', text: `*期限*\n${deadlineYmd || '-'}` }
  ];

  const blocks = [
    { type: 'section', text: { type: 'mrkdwn', text: `*${title}*` } },
    { type: 'section', fields }
  ];
  const actions = [];
  if (rec[QUEUE_HEADERS.CELL_LINK]) actions.push({ type: 'button', text: { type: 'plain_text', text: 'セルを開く' }, url: rec[QUEUE_HEADERS.CELL_LINK] });
  if (rec[QUEUE_HEADERS.FILE_URL])  actions.push({ type: 'button', text: { type: 'plain_text', text: '関連ファイル' }, url: rec[QUEUE_HEADERS.FILE_URL] });
  if (actions.length) blocks.push({ type: 'actions', elements: actions });

  return { text: title, blocks };
}
