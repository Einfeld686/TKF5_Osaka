const fs = require('fs');
const path = require('path');
const { printLog, printSlack } = require('./display');

// モックデータのロード
const MOCK_DIR = path.join(__dirname, '../mock');
const configData = JSON.parse(fs.readFileSync(path.join(MOCK_DIR, 'config.json'), 'utf8'));
const tasksData = JSON.parse(fs.readFileSync(path.join(MOCK_DIR, 'tasks.json'), 'utf8'));
const membersData = JSON.parse(fs.readFileSync(path.join(MOCK_DIR, 'members.json'), 'utf8'));
const queueData = JSON.parse(fs.readFileSync(path.join(MOCK_DIR, 'queue.json'), 'utf8'));
const slackRequests = [];
const spreadsheetStore = new Map();
const triggerStore = [];

// Utilities Mock
global.Utilities = {
    formatDate: (date, tz, fmt) => {
        // 簡易的な実装: yyyy-MM-dd のみ対応
        const d = new Date(date);
        return d.toISOString().split('T')[0];
    },
    sleep: (ms) => { /* no-op */ }
};

// UrlFetchApp Mock
global.UrlFetchApp = {
    fetch: (url, params) => {
        printLog('API', 'UrlFetchApp.fetch called');
        if (url.includes('slack.com')) {
            const payload = JSON.parse(params.payload);
            printSlack(payload);
            slackRequests.push({ url, payload, params });
        }
        return {
            getResponseCode: () => 200,
            getContentText: () => 'ok'
        };
    }
};

// Sheet Mock
class SheetMock {
    constructor(name, data) {
        this.name = name;
        this.data = data; // 2D array representation
        this.validations = [];
        this.maxRows = Math.max(this.data.length, 1000);
    }

    getName() { return this.name; }
    getSheetId() { return 12345; } // dummy
    getLastRow() { return this.data.length; }
    getLastColumn() { return this.data.reduce((max, row) => Math.max(max, row.length), 0); }
    getMaxRows() { return this.maxRows; }

    _ensureSize(row, col) {
        while (this.data.length < row) this.data.push([]);
        const targetRow = this.data[row - 1];
        while (targetRow.length < col) targetRow.push('');
        if (row > this.maxRows) this.maxRows = row;
    }

    getRange(row, col, numRows = 1, numCols = 1) {
        return new RangeMock(this, row, col, numRows, numCols);
    }

    getDataRange() {
        return new RangeMock(this, 1, 1, this.getLastRow(), this.getLastColumn());
    }
}

// Range Mock
class RangeMock {
    constructor(sheet, row, col, numRows, numCols) {
        this.sheet = sheet;
        this.row = row;
        this.col = col;
        this.numRows = numRows;
        this.numCols = numCols;
    }

    getRow() { return this.row; }
    getColumn() { return this.col; }
    getSheet() { return this.sheet; }

    getValue() {
        const r = this.sheet.data[this.row - 1];
        return r ? r[this.col - 1] : '';
    }

    getValues() {
        const res = [];
        for (let i = 0; i < this.numRows; i++) {
            const r = this.sheet.data[this.row - 1 + i] || [];
            const rowData = [];
            for (let j = 0; j < this.numCols; j++) {
                rowData.push(r[this.col - 1 + j]);
            }
            res.push(rowData);
        }
        return res;
    }

    getDisplayValue() { return String(this.getValue()); }
    getDisplayValues() { return this.getValues().map(r => r.map(c => String(c))); }

    setValue(val) {
        printLog('Sheet', `SetValue [${this.sheet.name}] ${val}`);
        this.sheet._ensureSize(this.row, this.col);
        this.sheet.data[this.row - 1][this.col - 1] = val;
    }

    setValues(values) {
        printLog('Sheet', `SetValues [${this.sheet.name}] (${values.length} rows)`);
        for (let i = 0; i < values.length; i++) {
            const rowValues = values[i] || [];
            for (let j = 0; j < rowValues.length; j++) {
                const r = this.row + i;
                const c = this.col + j;
                this.sheet._ensureSize(r, c);
                this.sheet.data[r - 1][c - 1] = rowValues[j];
            }
        }
    }

    setDataValidation(rule) {
        this.sheet.validations.push({
            row: this.row,
            col: this.col,
            numRows: this.numRows,
            numCols: this.numCols,
            rule
        });
    }

    getA1Notation() {
        const colStr = String.fromCharCode(64 + this.col); // 簡易: A-Zのみ
        return `${colStr}${this.row}`;
    }
}

// Spreadsheet Mock
class SpreadsheetMock {
    constructor(id, url) {
        this.id = id;
        this.url = url;
        this.sheets = new Map();
    }

    getId() { return this.id; }
    getUrl() { return this.url; }
    toast(msg, title, seconds) {
        const label = title ? ` [${title}]` : '';
        printLog('Toast', `${msg}${label}`);
    }

    getSheetByName(name) {
        if (this.sheets.has(name)) return this.sheets.get(name);

        // Config Sheet
        if (name === 'Config') {
            // JSONから2次元配列へ変換
            const headers = ['enabled', 'site_code', 'spreadsheet_url', 'task_sheet', 'members_sheet', 'slack_webhook', 'mode', 'notify_timing', 'status', 'last_updated'];
            const data = [headers];
            configData.forEach(c => {
                data.push(headers.map(k => {
                    if (typeof c[k] === 'boolean') return c[k] ? 'TRUE' : 'FALSE';
                    return c[k] || '';
                }));
            });
            const sh = new SheetMock(name, data);
            this.sheets.set(name, sh);
            return sh;
        }

        // Tasks Sheet
        if (name === 'Tasks') {
            // tasks.json からマトリクス構築
            // Layout based on constants.gs: TASK_SHEET_LAYOUT
            // Row 4: Task Name
            // Row 5: Deadline
            // Row 6: File URL
            // Row 7~: Staff, Checks

            const maxCols = 4 + tasksData.taskNames.length; // Staff Col(4) + Tasks
            const data = [];

            // 1-3: 空行
            data.push(new Array(maxCols).fill(''));
            data.push(new Array(maxCols).fill(''));
            data.push(new Array(maxCols).fill(''));

            // 4: Task Names
            const row4 = new Array(maxCols).fill('');
            tasksData.taskNames.forEach((n, i) => row4[4 + i] = n);
            data.push(row4);

            // 5: Deadlines
            const row5 = new Array(maxCols).fill('');
            tasksData.deadlines.forEach((d, i) => row5[4 + i] = new Date(d));
            data.push(row5);

            // 6: URLs
            const row6 = new Array(maxCols).fill('');
            tasksData.fileUrls.forEach((u, i) => row6[4 + i] = u);
            data.push(row6);

            // 7~: Staff rows
            tasksData.staffs.forEach(staff => {
                const row = new Array(maxCols).fill('');
                row[3] = staff.name; // Col 4 (index 3)
                staff.checks.forEach((chk, i) => row[4 + i] = chk);
                data.push(row);
            });

            const sh = new SheetMock(name, data);
            this.sheets.set(name, sh);
            return sh;
        }

        // Members Sheet
        if (name === 'Members') {
            // 1: Header
            const data = [['Name', 'SlackID', 'Role']];
            membersData.forEach(m => {
                data.push([m.name, m.slack_id, '']);
            });
            const sh = new SheetMock(name, data);
            this.sheets.set(name, sh);
            return sh;
        }

        // Queue Sheet
        if (name === 'Queue') {
            const headers = ['key', 'status', 'site_code', 'spreadsheet_id', 'sheet_name', 'staff_name', 'staff_slack_id', 'task_name', 'deadline_iso', 'file_url', 'cell_a1', 'cell_link', 'reason', 'slack_webhook', 'template_id', 'created_at', 'sent_at'];
            const data = [headers];
            queueData.forEach(q => {
                data.push(headers.map(k => q[k] || ''));
            });
            const sh = new SheetMock(name, data);
            this.sheets.set(name, sh);
            return sh;
        }

        return null;
    }

    insertSheet(name) {
        printLog('Sheet', `InsertSheet: ${name}`);
        const sh = new SheetMock(name, []);
        this.sheets.set(name, sh);
        return sh;
    }
}

// Global SpreadsheetApp
const ACTIVE_SPREADSHEET_ID = 'main-doc-id';
const ACTIVE_SPREADSHEET_URL = 'http://main-doc';

function getOrCreateSpreadsheet(id, url) {
    if (!spreadsheetStore.has(id)) {
        spreadsheetStore.set(id, new SpreadsheetMock(id, url));
    }
    return spreadsheetStore.get(id);
}

global.SpreadsheetApp = {
    getActiveSpreadsheet: () => {
        // デフォルトでConfig用のシートを返すなどの挙動
        return getOrCreateSpreadsheet(ACTIVE_SPREADSHEET_ID, ACTIVE_SPREADSHEET_URL);
    },
    openById: (id) => {
        // どのIDでも同じモック構造を返す（簡易化のため）
        return getOrCreateSpreadsheet(id, `http://doc/${id}`);
    },
    getUi: () => ({
        createMenu: () => {
            const menu = {
                addItem: () => menu,
                addToUi: () => { }
            };
            return menu;
        },
        alert: (message) => printLog('UI', `Alert: ${message}`)
    }),
    newDataValidation: () => {
        const state = { type: '', values: [], allowInvalid: true };
        const builder = {
            requireCheckbox: () => { state.type = 'checkbox'; return builder; },
            requireValueInList: (values, showDropdown) => {
                state.type = 'list';
                state.values = values || [];
                state.showDropdown = !!showDropdown;
                return builder;
            },
            setAllowInvalid: (allow) => { state.allowInvalid = !!allow; return builder; },
            build: () => ({ ...state })
        };
        return builder;
    }
};

const EVENT_TYPES = { ON_EDIT: 'ON_EDIT', CLOCK: 'CLOCK' };

global.ScriptApp = {
    getProjectTriggers: () => triggerStore.slice(),
    newTrigger: (handler) => ({
        forSpreadsheet: (id) => ({
            onEdit: () => ({
                create: () => {
                    triggerStore.push({
                        handler,
                        eventType: EVENT_TYPES.ON_EDIT,
                        sourceId: id,
                        getEventType() { return this.eventType; },
                        getTriggerSourceId() { return this.sourceId; },
                        getHandlerFunction() { return this.handler; }
                    });
                }
            })
        }),
        timeBased: () => {
            const timeTrigger = {
                handler,
                eventType: EVENT_TYPES.CLOCK,
                hour: null,
                minute: null,
                getEventType() { return this.eventType; },
                getHandlerFunction() { return this.handler; }
            };
            const builder = {
                everyDays: () => builder,
                atHour: (hour) => { timeTrigger.hour = hour; return builder; },
                nearMinute: (minute) => { timeTrigger.minute = minute; return builder; },
                create: () => { triggerStore.push(timeTrigger); }
            };
            return builder;
        }
    }),
    deleteTrigger: (trigger) => {
        const idx = triggerStore.indexOf(trigger);
        if (idx >= 0) triggerStore.splice(idx, 1);
    },
    EventType: EVENT_TYPES
};

function resetMockState() {
    slackRequests.length = 0;
    triggerStore.length = 0;
    spreadsheetStore.clear();
}

module.exports = {
    getSpreadsheetMock: (id) => getOrCreateSpreadsheet(id, 'http://mock'),
    getActiveSpreadsheetMock: () => getOrCreateSpreadsheet(ACTIVE_SPREADSHEET_ID, ACTIVE_SPREADSHEET_URL),
    getSlackRequests: () => slackRequests,
    resetMockState
};
