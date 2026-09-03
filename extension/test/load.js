// Loads the extension's classic scripts into one shared scope, the way the
// content script shares one global scope in Chrome (and run_tests.js does for
// Apps Script). Run `node extension/build.js --lib-only` first.
//
// Built with `new Function` in this process's own realm, not `vm` — vm.createContext
// mints a separate V8 realm, so every {}/[] literal produced by code run through it
// gets a different Array/Object than the one in this file, and assert.deepStrictEqual
// (what assert/strict's deepEqual is) treats those as unequal even with identical
// contents. Top-level `var`/`function` declarations are recovered by regex — every
// .gs file here only ever declares single names at top level (no `var a, b;`) — and
// returned as an object, mimicking Apps Script's shared globals.
const fs = require('fs');
const path = require('path');

const ext = path.join(__dirname, '..');
const FILES = ['lib/formula.js', 'lib/trace.js', 'adapter.js', 'audit.js', 'state.js', 'dom.js'];
const DECL = /^(?:var|function)\s+([A-Za-z_$][\w$]*)/gm;

function loadSandbox(extra) {
  const globals = Object.assign({ Logger: { log() {} } }, extra || {});
  const names = new Set();
  let source = '';
  for (const f of FILES) {
    const p = path.join(ext, f);
    if (!fs.existsSync(p)) continue; // later tasks add files; earlier tests must still run
    const code = fs.readFileSync(p, 'utf8');
    for (const m of code.matchAll(DECL)) names.add(m[1]);
    source += code + '\n';
  }
  const params = Object.keys(globals);
  const body = source + '\nreturn {' + [...names, ...params].join(',') + '};';
  const fn = new Function(...params, body);
  return fn(...params.map((k) => globals[k]));
}

module.exports = { loadSandbox };
