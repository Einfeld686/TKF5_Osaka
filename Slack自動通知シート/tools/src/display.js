const chalk = require('chalk');

function printHeader(title) {
    console.log(chalk.cyan('╔════════════════════════════════════════════════════════╗'));
    console.log(chalk.cyan(`║  ${padCenter(title, 52)}  ║`));
    console.log(chalk.cyan('╠════════════════════════════════════════════════════════╣'));
}

function printFooter() {
    console.log(chalk.cyan('╚════════════════════════════════════════════════════════╝'));
}

function printLog(step, message, detail = '') {
    const stepStr = `[${step}]`.padEnd(4);
    console.log(chalk.green(`║  ${stepStr} ${message.padEnd(46)} ║`));
    if (detail) {
        console.log(chalk.gray(`║       → ${detail.padEnd(44)} ║`));
    }
}

function printSlack(payload) {
    console.log(chalk.cyan('╠════════════════════════════════════════════════════════╣'));
    console.log(chalk.yellow('║  📤 Slack送信内容 (Preview):                            ║'));
    console.log(chalk.cyan('║  ────────────────────────────────────────              ║'));

    const text = payload.text || '';
    console.log(chalk.white(`║  ${text.padEnd(52)}  ║`));

    if (payload.blocks) {
        for (const block of payload.blocks) {
            if (block.fields) {
                for (const f of block.fields) {
                    const lines = f.text.split('\n');
                    for (const line of lines) {
                        console.log(chalk.white(`║  ${line.replace(/\*/g, '').padEnd(52)}  ║`));
                    }
                }
            }
        }
    }
    printFooter();
}

function padCenter(str, length) {
    const padding = length - getStringWidth(str);
    const left = Math.floor(padding / 2);
    const right = padding - left;
    return ' '.repeat(left) + str + ' '.repeat(right);
}

function getStringWidth(str) {
    let width = 0;
    for (let i = 0; i < str.length; i++) {
        const c = str.charCodeAt(i);
        // 簡易的な全角判定
        if ((c >= 0x3000 && c <= 0xffff) || (c >= 0xff01 && c <= 0xff60)) {
            width += 2;
        } else {
            width += 1;
        }
    }
    return width;
}

module.exports = {
    printHeader,
    printFooter,
    printLog,
    printSlack
};
