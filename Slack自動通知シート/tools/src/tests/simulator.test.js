const test = require('node:test');
const assert = require('node:assert/strict');
const { loadGasFiles } = require('../test-helpers');
const {
  resetMockState,
  getSlackRequests,
  getActiveSpreadsheetMock
} = require('../mock-gas');

loadGasFiles();

function withFixedDate(isoString, fn) {
  const RealDate = Date;
  class MockDate extends RealDate {
    constructor(...args) {
      if (args.length === 0) {
        super(isoString);
      } else {
        super(...args);
      }
    }
    static now() { return new RealDate(isoString).getTime(); }
    static parse(str) { return RealDate.parse(str); }
    static UTC(...args) { return RealDate.UTC(...args); }
  }
  global.Date = MockDate;
  try {
    return fn();
  } finally {
    global.Date = RealDate;
  }
}

test('handleEdit sends Slack payload', () => {
  resetMockState();
  const e = {
    source: SpreadsheetApp.openById('mock-doc-osaka'),
    range: {
      getRow: () => 7,
      getColumn: () => 5,
      getSheet: () => e.source.getSheetByName('Tasks'),
      getA1Notation: () => 'E7',
      getDisplayValue: () => 'TRUE',
      getValue: () => 'TRUE'
    },
    value: 'TRUE'
  };

  handleEdit(e);

  const requests = getSlackRequests();
  assert.equal(requests.length, 1);
  assert.match(requests[0].payload.text, /\[OSAKA\]/);
  assert.match(requests[0].payload.text, /日次報告/);
});

test('nightlyScanAndQueue adds queue entries', () => {
  resetMockState();
  const ss = getActiveSpreadsheetMock();
  const queueSheet = ss.getSheetByName('Queue');
  const initialRows = queueSheet.getLastRow();

  withFixedDate('2025-12-17T00:00:00+09:00', () => {
    nightlyScanAndQueue();
  });

  const afterRows = queueSheet.getLastRow();
  assert.equal(afterRows, initialRows + 4);
});

test('morningDispatch updates queue and writes outbox', () => {
  resetMockState();
  const ss = getActiveSpreadsheetMock();
  const queueSheet = ss.getSheetByName('Queue');
  const header = queueSheet.getRange(1, 1, 1, queueSheet.getLastColumn()).getValues()[0];
  const idxKey = header.indexOf('key');
  const idxStatus = header.indexOf('status');
  const idxSentAt = header.indexOf('sent_at');
  const targetKey = 'TOKYO|123|10|5|週次MTG資料|2025-12-17';

  morningDispatch();

  const requests = getSlackRequests();
  assert.equal(requests.length, 1);

  const data = queueSheet.getRange(2, 1, queueSheet.getLastRow() - 1, queueSheet.getLastColumn()).getValues();
  const targetRow = data.find(row => row[idxKey] === targetKey);
  assert.ok(targetRow);
  assert.equal(targetRow[idxStatus], 'SENT');
  assert.ok(targetRow[idxSentAt]);

  const outbox = ss.getSheetByName('Outbox');
  assert.ok(outbox);
  assert.equal(outbox.getLastRow(), 2);
});
