/**
 * Builds tests/harness.html: the real Sidebar.html with a google.script.run
 * mock injected ahead of it. The mock executes the REAL server code
 * (Formula.gs + Trace.gs + Code.gs) against the same in-memory spreadsheet the
 * e2e harness uses, with a configurable artificial latency, so Playwright
 * drives the sidebar exactly as Apps Script would serve it.
 *
 *   node tests/build-harness.js  ->  tests/harness.html
 */
const fs = require('fs');
const path = require('path');
const root = path.join(__dirname, '..');

const sidebar = fs.readFileSync(path.join(root, 'Sidebar.html'), 'utf8');
const server = ['Formula.gs', 'Trace.gs', 'Code.gs']
  .map((f) => fs.readFileSync(path.join(root, f), 'utf8'))
  .join('\n');

// The e2e file holds the SpreadsheetApp mock; reuse its class definitions.
const e2e = fs.readFileSync(path.join(root, 'run_e2e.js'), 'utf8');
const mockStart = e2e.indexOf('function colToA(');
const mockEnd = e2e.indexOf('// --- load the add-on');
const mockClasses = e2e.slice(mockStart, mockEnd);

const bootstrap = `
<script>
// ---- in-browser Apps Script stand-in (real server code + mock spreadsheet) --
(function () {
${mockClasses}

const inputs = sheetOf('Inputs', {
  1: { 1: 'Growth', 2: 0.05 },
  2: { 1: 'Base', 2: 100 },
  3: { 1: 'Shift', 2: 1 },
  4: { 1: 'Empty', 2: '' }
});
const hidden = sheetOf('Hidden Calc', { 1: { 1: 7 }, 2: { 1: 8 }, 3: { 1: 9 } }, true);
const model = sheetOf('Model', {
  1: { 1: "=Inputs!B2*(1+Inputs!B1)" },
  2: { 1: "=OFFSET('Hidden Calc'!A1,Inputs!B3,0)" },
  3: { 1: '=INDEX(Inputs!B1:B4,4)' },
  4: { 1: '=A1+A2' },
  5: { 1: '=INDIRECT("Inputs!B2")' },
  6: { 1: '=SUM(Inputs!B1:B3)' },
  7: { 1: '=OFFSET(Inputs!B1,MATCH(1,Inputs!A:A,0),0)' }
});
const ss = new Spreadsheet([model, inputs, hidden], [
  { getName: () => 'Growth', getRange: () => inputs.getRange('B1') },
  { getName: () => 'Shift', getRange: () => inputs.getRange('B3') }
]);
ss.setActiveRange(model.getRange('A2'));
window.__ss = ss; // test hook: Playwright inspects grid state through this

const props = {};
window.PropertiesService = {
  getDocumentProperties: () => ({
    getProperty: (k) => (k in props ? props[k] : null),
    setProperty: (k, v) => { props[k] = v; }
  })
};
window.SpreadsheetApp = { getActive: () => ss };
window.Logger = { log: () => {} };

// ---- real server code -------------------------------------------------------
${server}

// the server functions live inside this IIFE; expose them for the shim
['scopionObserve', 'scopionNavigate', 'scopionAudit', 'scopionGetSettings',
 'scopionSetSetting', 'scopionRehide', 'scopionClearHighlight',
 'scopionAuditCore'].forEach((n) => { window[n] = eval(n); });

// ---- google.script.run shim with latency ------------------------------------
window.__rpcLog = [];
window.__latency = 250;
window.__closed = false;
window.google = { script: { run: null, host: { close: () => { window.__closed = true; } } } };
function makeRunner() {
  let ok = function () {}, fail = function () {};
  const runner = {
    withSuccessHandler(f) { ok = f; return runner; },
    withFailureHandler(f) { fail = f; return runner; }
  };
  ['scopionObserve', 'scopionNavigate', 'scopionAudit', 'scopionGetSettings',
   'scopionSetSetting', 'scopionRehide', 'scopionClearHighlight'].forEach((name) => {
    runner[name] = function () {
      const args = Array.from(arguments);
      window.__rpcLog.push({ name, args: JSON.parse(JSON.stringify(args)) });
      setTimeout(() => {
        try {
          const result = window[name].apply(null, args);
          ok(JSON.parse(JSON.stringify(result === undefined ? null : result)));
        } catch (e) { fail(e); }
      }, window.__latency);
    };
  });
  return runner;
}
Object.defineProperty(window.google.script, 'run', { get: makeRunner });
})();
</script>
`;

// function replacement: $' / $& inside the payload must stay literal
const html = sidebar.replace('<script>\n\'use strict\';', () => bootstrap + '<script>\n\'use strict\';');
if (html === sidebar) throw new Error('injection anchor not found');
fs.writeFileSync(path.join(__dirname, 'harness.html'), html);
console.log('tests/harness.html written,', html.length, 'bytes');
