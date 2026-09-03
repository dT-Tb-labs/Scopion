// Loads the extension's classic scripts into one vm sandbox, the way the
// content script shares one global scope in Chrome (and run_tests.js does for
// Apps Script). Run `node extension/build.js --lib-only` first.
const fs = require('fs');
const vm = require('vm');
const path = require('path');

const ext = path.join(__dirname, '..');
const FILES = ['lib/formula.js', 'lib/trace.js', 'adapter.js', 'audit.js', 'state.js', 'dom.js'];

function loadSandbox(extra) {
  const sandbox = Object.assign({
    console, Date, Math, JSON, String, Number, Array, Object, RegExp, Promise, Error,
    isNaN, isFinite, parseInt, parseFloat, setTimeout, clearTimeout,
    Logger: { log() {} }
  }, extra || {});
  vm.createContext(sandbox);
  for (const f of FILES) {
    const p = path.join(ext, f);
    if (!fs.existsSync(p)) continue; // later tasks add files; earlier tests must still run
    vm.runInContext(fs.readFileSync(p, 'utf8'), sandbox, { filename: f });
  }
  return sandbox;
}

module.exports = { loadSandbox };
