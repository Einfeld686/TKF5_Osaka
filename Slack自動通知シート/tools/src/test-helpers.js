const fs = require('fs');
const path = require('path');
const vm = require('vm');
require('./mock-gas');

const ROOT_DIR = path.join(__dirname, '../../');
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

let loaded = false;

function loadGasFiles() {
  if (loaded) return;
  GAS_FILES.forEach(file => {
    const filePath = path.join(ROOT_DIR, file);
    if (fs.existsSync(filePath)) {
      const code = fs.readFileSync(filePath, 'utf8');
      vm.runInThisContext(code);
    }
  });
  loaded = true;
}

module.exports = { loadGasFiles };
