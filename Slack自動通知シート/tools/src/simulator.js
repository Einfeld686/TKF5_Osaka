const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { printHeader, printLog } = require('./display');
require('./mock-gas'); // GAS環境モックのロード

// プロジェクトルート
const ROOT_DIR = path.join(__dirname, '../../');

// GASファイルのロード順序
const GAS_FILES = [
    'constants.gs',
    'utils.gs',
    'members.gs',
    'config_repo.gs',
    'slack.gs',
    'queue.gs',
    'handle_edit.gs',
    'triggers.gs'
];

// GASコードをグローバルコンテキストにロード
function loadGasFiles() {
    printLog('System', 'Loading GAS files...');
    GAS_FILES.forEach(file => {
        const filePath = path.join(ROOT_DIR, file);
        if (fs.existsSync(filePath)) {
            const code = fs.readFileSync(filePath, 'utf8');
            vm.runInThisContext(code);
            // printLog('System', `Loaded ${file}`);
        } else {
            console.error(`File not found: ${filePath}`);
        }
    });
    printLog('System', 'All files loaded successfully.');
}

// シミュレーション実行
async function run() {
    const mode = process.argv[2];

    loadGasFiles();

    // Patch: Missing notifyUser_ function
    global.notifyUser_ = (msg, type) => {
        const color = type === 'ERROR' ? 'red' : (type === 'WARN' ? 'yellow' : 'cyan');
        printLog('Notify', `[${type}] ${msg}`);
    };

    switch (mode) {
        case 'edit':
            printHeader('🔔 即時通知シミュレーション (onEdit)');
            await runEditSim();
            break;
        case 'nightly':
            printHeader('🌙 夜間スキャンシミュレーション');
            await runNightlySim();
            break;
        case 'morning':
            printHeader('☀️ 朝の送信シミュレーション');
            await runMorningSim();
            break;
        default:
            console.log('Usage: node src/simulator.js [edit|nightly|morning]');
    }
}

// 1. 即時通知シミュレーション
async function runEditSim() {
    printLog('Sim', 'Scenario: TasksシートのチェックボックスをTRUEに変更');

    // モックイベントオブジェクト作成
    // tasks.json の "佐藤 健" (Row 7, Col 5 [index 0]) のタスク "日次報告" を編集したことにする
    const e = {
        source: SpreadsheetApp.openById('mock-doc-osaka'), // OSAKA (notify_immediate)
        range: {
            getRow: () => 7, // 佐藤 健
            getColumn: () => 5, // 日次報告
            getSheet: () => e.source.getSheetByName('Tasks'),
            getA1Notation: () => 'E7',
            getDisplayValue: () => 'TRUE',
            getValue: () => 'TRUE'
        },
        value: 'TRUE'
    };

    printLog('Input', 'Event: mock-doc-osaka / Tasks / E7 / TRUE');

    try {
        handleEdit(e);
    } catch (err) {
        console.error(err);
    }
}

// 2. 夜間スキャンシミュレーション
async function runNightlySim() {
    printLog('Sim', 'Running nightlyScanAndQueue()...');
    try {
        nightlyScanAndQueue();
    } catch (err) {
        console.error(err);
    }
}

// 3. 朝送信シミュレーション
async function runMorningSim() {
    printLog('Sim', 'Running morningDispatch()...');
    try {
        // QueueにあるPENDINGを送信
        morningDispatch();
    } catch (err) {
        console.error(err);
    }
}

// 実行
run();
