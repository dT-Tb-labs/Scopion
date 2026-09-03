# Scopion Chrome Extension Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** A Chrome extension that reproduces the Excel ACE floating window in Google Sheets: shortcut → strip-shaped panel next to the selected cell listing its precedents; ↑/↓ walk them (grid selection follows), Enter drills in, Back returns, OK/Esc closes and restores.

**Architecture:** MV3 extension. A content script on `docs.google.com/spreadsheets/*` reads the selection from the Sheets DOM (name box, formula bar, selection border), asks the background service worker for cell data via Sheets REST API v4 (OAuth through `chrome.identity`), runs the existing `Formula.gs`/`Trace.gs` parser+resolver unmodified against an in-memory snapshot shaped like `SpreadsheetApp`, and renders a Shadow-DOM panel. Jumps are done by typing into the name box.

**Tech Stack:** Plain JS (classic scripts sharing one global scope — same model as Apps Script), Chrome MV3 (`identity`, `storage`, `commands`), Sheets API v4 `spreadsheets.get`. Tests: Node 26 built-in `node:test` + `vm` (no npm dependencies, same loading trick as `run_tests.js`).

**Spec:** `docs/superpowers/specs/2026-09-03-scopion-chrome-extension-design.md`

## Global Constraints

- `Formula.gs` and `Trace.gs` are **not edited**. They are copied into `extension/lib/` by `extension/build.js`.
- Dependents are out of scope; `findDependents` is never called.
- Nothing is ever written to the spreadsheet. OAuth scope is exactly `https://www.googleapis.com/auth/spreadsheets.readonly`.
- Credentials never enter git: `extension/oauth.local.json`, `extension/manifest.json`, `extension/lib/`, `extension/*.pem` are git-ignored.
- Default shortcut `Ctrl+Shift+A` on every platform (not `Cmd+Shift+A`).
- "Include hidden sheets" defaults to **off**.
- Code style follows the repo: `function` declarations, `var` for module globals, JSDoc comments that say *why*. No bundler, no framework.
- Every task ends with `node run_tests.js && node run_e2e.js && node --test extension/test/` green, then a commit. Commit messages: `<type>: <terse description>`, no AI attribution.
- Deviation from spec, recorded here: `panel.css` is a JS string inside `panel.js` (Shadow DOM needs the CSS inside the shadow root; a manifest `css` entry cannot reach it). DOM tests cover the pure helpers of `dom.js`; selector glue is verified by the manual smoke test (Task 10) — Node has no DOM and adding jsdom is a dependency the spec does not want.

## File structure

```
extension/
  build.js                 copy Formula.gs/Trace.gs → lib/; render manifest.json from template + oauth.local.json
  manifest.template.json   MV3 manifest with empty key/client_id
  oauth.example.json       shape of oauth.local.json
  background.js            shortcut → tab message; Sheets API proxy with OAuth token handling
  adapter.js               Snapshot / SnapSheet / SnapRange / SnapNamedRange (SpreadsheetApp look-alike over API JSON)
  audit.js                 one audit: rectsToFetch → fetch → findPrecedents → second fetch → buildRows
  state.js                 pure walk state (highlight index, history)
  dom.js                   Sheets page adapters (selectors live here only) + pure helpers
  panel.js                 Shadow-DOM panel (ACE geometry), keyboard, drag, highlight overlay
  content.js               wiring: toggle → audit → walk/drill/back/close
  test/
    load.js                vm sandbox loader (mirrors run_tests.js)
    fake-api.js            in-memory Sheets API for audit tests
    fixtures/meta.json     spreadsheets.get metadata for the Model/Inputs/Hidden Calc scenario
    build.test.js adapter.test.js audit.test.js state.test.js dom.test.js background.test.js
```

Load order in the content script (and in `test/load.js`): `lib/formula.js, lib/trace.js, adapter.js, audit.js, state.js, dom.js, panel.js, content.js`. Later files may call functions from earlier ones; never the reverse.

---

### Task 1: Scaffold — build script, manifest template, test loader

**Files:**
- Create: `extension/build.js`, `extension/manifest.template.json`, `extension/oauth.example.json`, `extension/test/load.js`, `extension/test/build.test.js`
- Modify: `.gitignore`

**Interfaces:**
- Produces: `node extension/build.js [--lib-only] [path/to/oauth.json]`; `loadSandbox(extraGlobals) -> sandbox` from `test/load.js` (globals: every function of the loaded files).

- [ ] **Step 1: Write the failing build test**

`extension/test/build.test.js`:
```js
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');

const ext = path.join(__dirname, '..');
const root = path.join(ext, '..');

test('build copies the Apps Script sources into lib/ unchanged', () => {
  execFileSync('node', [path.join(ext, 'build.js'), '--lib-only']);
  assert.equal(fs.readFileSync(path.join(ext, 'lib/formula.js'), 'utf8'), fs.readFileSync(path.join(root, 'Formula.gs'), 'utf8'));
  assert.equal(fs.readFileSync(path.join(ext, 'lib/trace.js'), 'utf8'), fs.readFileSync(path.join(root, 'Trace.gs'), 'utf8'));
});

test('build renders manifest.json from the template and the local oauth file', () => {
  const tmp = path.join(ext, 'test', 'oauth.tmp.json');
  fs.writeFileSync(tmp, JSON.stringify({ client_id: 'cid.apps.googleusercontent.com', key: 'MIIB' }));
  try {
    execFileSync('node', [path.join(ext, 'build.js'), tmp]);
    const m = JSON.parse(fs.readFileSync(path.join(ext, 'manifest.json'), 'utf8'));
    assert.equal(m.manifest_version, 3);
    assert.equal(m.oauth2.client_id, 'cid.apps.googleusercontent.com');
    assert.equal(m.key, 'MIIB');
    assert.deepEqual(m.oauth2.scopes, ['https://www.googleapis.com/auth/spreadsheets.readonly']);
    assert.equal(m.commands['toggle-scopion'].suggested_key.default, 'Ctrl+Shift+A');
    assert.equal(m.commands['toggle-scopion'].suggested_key.mac, 'Ctrl+Shift+A');
  } finally {
    fs.unlinkSync(tmp);
    fs.rmSync(path.join(ext, 'manifest.json'), { force: true });
  }
});

test('build refuses to render a manifest without credentials', () => {
  assert.throws(
    () => execFileSync('node', [path.join(ext, 'build.js'), path.join(ext, 'test', 'does-not-exist.json')], { stdio: 'pipe' }),
    (e) => e.status === 1 && /oauth\.example\.json/.test(String(e.stderr))
  );
});
```

- [ ] **Step 2: Run it to verify it fails**

Run: `node --test extension/test/build.test.js`
Expected: FAIL — `Cannot find module .../extension/build.js`

- [ ] **Step 3: Write `.gitignore` additions, template, example, build script**

Append to `.gitignore`:
```
extension/lib/
extension/manifest.json
extension/oauth.local.json
extension/*.pem
extension/test/oauth.tmp.json
```

`extension/manifest.template.json`:
```json
{
  "manifest_version": 3,
  "name": "Scopion",
  "version": "0.1.0",
  "description": "Active Cell Explorer for Google Sheets: list a formula's precedents and walk them with the arrow keys.",
  "key": "",
  "permissions": ["identity", "storage"],
  "host_permissions": ["https://docs.google.com/*", "https://sheets.googleapis.com/*"],
  "background": { "service_worker": "background.js" },
  "content_scripts": [
    {
      "matches": ["https://docs.google.com/spreadsheets/*"],
      "js": ["lib/formula.js", "lib/trace.js", "adapter.js", "audit.js", "state.js", "dom.js", "panel.js", "content.js"],
      "run_at": "document_idle"
    }
  ],
  "commands": {
    "toggle-scopion": {
      "suggested_key": { "default": "Ctrl+Shift+A", "mac": "Ctrl+Shift+A" },
      "description": "Open Scopion on the selected cell (re-audit when already open)"
    }
  },
  "oauth2": {
    "client_id": "",
    "scopes": ["https://www.googleapis.com/auth/spreadsheets.readonly"]
  }
}
```

`extension/oauth.example.json`:
```json
{
  "client_id": "1234567890-abc.apps.googleusercontent.com",
  "key": "<base64 DER public key of extension/scopion.pem — see README 'Chrome extension' section>"
}
```

`extension/build.js`:
```js
#!/usr/bin/env node
/**
 * build.js — the extension is assembled, not hand-maintained:
 *   lib/formula.js, lib/trace.js  are byte copies of ../Formula.gs, ../Trace.gs
 *   manifest.json                  is manifest.template.json + oauth.local.json
 * The Apps Script files stay the single source of truth for the parser and
 * resolver; credentials stay out of git.
 *
 *   node extension/build.js              full build (needs extension/oauth.local.json)
 *   node extension/build.js --lib-only   lib/ only, for tests
 *   node extension/build.js path.json    use another oauth file
 */
const fs = require('fs');
const path = require('path');

const ext = __dirname;
const root = path.join(ext, '..');
const args = process.argv.slice(2);
const libOnly = args.includes('--lib-only');
const oauthPath = args.find((a) => a !== '--lib-only') || path.join(ext, 'oauth.local.json');

fs.mkdirSync(path.join(ext, 'lib'), { recursive: true });
for (const [src, dst] of [['Formula.gs', 'formula.js'], ['Trace.gs', 'trace.js']]) {
  fs.copyFileSync(path.join(root, src), path.join(ext, 'lib', dst));
}
if (libOnly) {
  console.log('built extension/lib');
  process.exit(0);
}

if (!fs.existsSync(oauthPath)) {
  console.error('missing ' + oauthPath + ' — copy extension/oauth.example.json to extension/oauth.local.json and fill it in');
  process.exit(1);
}
const local = JSON.parse(fs.readFileSync(oauthPath, 'utf8'));
if (!local.client_id || !local.key) {
  console.error(oauthPath + ' needs both "client_id" and "key" (oauth credentials)');
  process.exit(1);
}
const manifest = JSON.parse(fs.readFileSync(path.join(ext, 'manifest.template.json'), 'utf8'));
manifest.oauth2.client_id = local.client_id;
manifest.key = local.key;
fs.writeFileSync(path.join(ext, 'manifest.json'), JSON.stringify(manifest, null, 2) + '\n');
console.log('built extension/lib and extension/manifest.json');
```

`extension/test/load.js`:
```js
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
```

- [ ] **Step 4: Run the build test**

Run: `node --test extension/test/build.test.js`
Expected: 3 passing.

- [ ] **Step 5: Commit**

```bash
git add .gitignore extension/build.js extension/manifest.template.json extension/oauth.example.json extension/test/load.js extension/test/build.test.js
git commit -m "chore: chrome extension scaffold, build script and test loader"
```

---

### Task 2: `adapter.js` — snapshot that looks like SpreadsheetApp

**Files:**
- Create: `extension/adapter.js`, `extension/test/fixtures/meta.json`, `extension/test/adapter.test.js`

**Interfaces:**
- Consumes (from `lib/trace.js`): `rect(sheetName,r1,c1,r2,c2)`, `rectToA1(rect)`, `a1ToRect(a1, sheetName)`, `rectWithin(inner, outer)`.
- Produces:
  - `new Snapshot(metaJson)` — `metaJson` is the `spreadsheets.get` response with `fields=namedRanges,sheets.properties(...)`.
  - `Snapshot#getSheets()`, `#getSheetByName(name) -> SnapSheet|null`, `#getNamedRanges() -> SnapNamedRange[]`
  - `Snapshot#addGridData(gridJson)` — merges a `spreadsheets.get?includeGridData=true` response.
  - `Snapshot#markFetched(rect)`, `Snapshot#isFetched(rect) -> boolean`
  - `SnapSheet`: `getName getSheetId isSheetHidden getLastRow getLastColumn getMaxRows getMaxColumns getDataRange getRange(a1 | r,c,nr,nc) cell(r,c) -> {v,d,f}`; fields `rowCount`, `columnCount`.
  - `SnapRange`: `getSheet getRow getColumn getLastRow getLastColumn getNumRows getNumColumns getA1Notation getValues getDisplayValues getFormulas getValue getDisplayValue getFormula`
  - `SnapNamedRange`: `getName() getRange()`
  - `apiRange(snap, rect) -> "'Sheet'!A1:B2"` clipped to the sheet grid; `quoteSheetName(name)`.

- [ ] **Step 1: Write the fixture and the failing tests**

`extension/test/fixtures/meta.json`:
```json
{
  "sheets": [
    { "properties": { "sheetId": 0, "title": "Model", "gridProperties": { "rowCount": 1000, "columnCount": 26 } } },
    { "properties": { "sheetId": 1, "title": "Inputs", "gridProperties": { "rowCount": 1000, "columnCount": 26 } } },
    { "properties": { "sheetId": 2, "title": "Hidden Calc", "hidden": true, "gridProperties": { "rowCount": 1000, "columnCount": 26 } } },
    { "properties": { "sheetId": 3, "title": "O'Connor", "gridProperties": { "rowCount": 50, "columnCount": 10 } } }
  ],
  "namedRanges": [
    { "namedRangeId": "n1", "name": "Growth", "range": { "sheetId": 1, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 1, "endColumnIndex": 2 } },
    { "namedRangeId": "n2", "name": "Shift", "range": { "sheetId": 1, "startRowIndex": 2, "endRowIndex": 3, "startColumnIndex": 1, "endColumnIndex": 2 } },
    { "namedRangeId": "n3", "name": "WholeCol", "range": { "sheetId": 1, "startColumnIndex": 0, "endColumnIndex": 1 } },
    { "namedRangeId": "n4", "name": "Orphan", "range": { "sheetId": 99, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 1 } }
  ]
}
```

`extension/test/adapter.test.js`:
```js
const test = require('node:test');
const assert = require('node:assert/strict');
const { loadSandbox } = require('./load');
const meta = require('./fixtures/meta.json');

const S = loadSandbox();

function grid(sheetId, startRow, startColumn, rows) {
  // rows: array of arrays of cell specs {n:number}|{s:string}|{f:formula,n:number}|{err:'N_A'}|null
  return { sheets: [{ properties: { sheetId }, data: [{ startRow, startColumn, rowData: rows.map((r) => ({
    values: r.map((c) => {
      if (c === null) return {};
      const v = {};
      if (c.err) { v.effectiveValue = { errorValue: { type: c.err } }; v.formattedValue = '#N/A'; }
      else if (typeof c.n === 'number') { v.effectiveValue = { numberValue: c.n }; v.formattedValue = String(c.n); }
      else if (typeof c.s === 'string') { v.effectiveValue = { stringValue: c.s }; v.formattedValue = c.s; }
      if (c.f) v.userEnteredValue = { formulaValue: c.f }; else if (v.effectiveValue) v.userEnteredValue = v.effectiveValue;
      return v;
    }) })) }] }] };
}

test('sheets and hidden state come from properties', () => {
  const snap = new S.Snapshot(meta);
  assert.equal(snap.getSheets().length, 4);
  assert.equal(snap.getSheetByName('Hidden Calc').isSheetHidden(), true);
  assert.equal(snap.getSheetByName('Model').isSheetHidden(), false);
  assert.equal(snap.getSheetByName('Nope'), null);
  assert.equal(snap.getSheetByName("O'Connor").getMaxRows(), 50);
});

test('named ranges become getName/getRange, unbounded axes fill the grid, orphans are dropped', () => {
  const snap = new S.Snapshot(meta);
  const names = snap.getNamedRanges().map((n) => n.getName());
  assert.deepEqual(names, ['Growth', 'Shift', 'WholeCol']);
  const growth = snap.getNamedRanges()[0].getRange();
  assert.equal(growth.getSheet().getName(), 'Inputs');
  assert.equal(growth.getA1Notation(), 'B1');
  const whole = snap.getNamedRanges()[2].getRange();
  assert.equal(whole.getRow(), 1); assert.equal(whole.getLastRow(), 1000);
  assert.equal(whole.getColumn(), 1); assert.equal(whole.getLastColumn(), 1);
});

test('buildNamedRangeMap from Trace.gs works on the snapshot', () => {
  const snap = new S.Snapshot(meta);
  const map = S.buildNamedRangeMap(snap);
  assert.deepEqual(map.GROWTH, { name: 'Growth', rect: { sheetName: 'Inputs', r1: 1, c1: 2, r2: 1, c2: 2 } });
});

test('grid data lands at startRow/startColumn and reads back as values, display values, formulas', () => {
  const snap = new S.Snapshot(meta);
  snap.addGridData(grid(1, 1, 1, [[{ n: 100 }], [{ n: 0.05 }], [{ f: '=1+1', n: 2 }], [null], [{ err: 'N_A' }]]));
  const inputs = snap.getSheetByName('Inputs');
  assert.equal(inputs.getRange('B2').getValue(), 100);
  assert.equal(inputs.getRange('B2').getDisplayValue(), '100');
  assert.equal(inputs.getRange('B2').getFormula(), '');
  assert.equal(inputs.getRange('B4').getFormula(), '=1+1');
  assert.equal(inputs.getRange('B4').getValue(), 2);
  assert.deepEqual(inputs.getRange('B5').getValues(), [['']]);
  assert.equal(inputs.getRange('B6').getValue(), '#N/A');
  assert.deepEqual(inputs.getRange(2, 2, 2, 1).getValues(), [[100], [0.05]]);
  assert.deepEqual(inputs.getRange('B2:B3').getDisplayValues(), [['100'], ['0.05']]);
  assert.deepEqual(inputs.getRange('B2:B4').getFormulas(), [[''], [''], ['=1+1']]);
});

test('cells never fetched read as blank, and getDataRange spans the whole grid', () => {
  const snap = new S.Snapshot(meta);
  const oc = snap.getSheetByName("O'Connor");
  assert.equal(oc.getRange('C7').getValue(), '');
  const all = oc.getDataRange();
  assert.equal(all.getNumRows(), 50); assert.equal(all.getNumColumns(), 10);
  assert.equal(all.getValues().length, 50);
});

test('whole-column A1 is clipped to the grid by getRange', () => {
  const snap = new S.Snapshot(meta);
  const r = snap.getSheetByName('Inputs').getRange('A:A');
  assert.equal(r.getLastRow(), 1000);
  assert.equal(r.getA1Notation(), 'A1:A1000');
});

test('fetched bookkeeping', () => {
  const snap = new S.Snapshot(meta);
  const r = S.rect('Inputs', 1, 2, 4, 2);
  assert.equal(snap.isFetched(S.rect('Inputs', 2, 2, 2, 2)), false);
  snap.markFetched(r);
  assert.equal(snap.isFetched(S.rect('Inputs', 2, 2, 2, 2)), true);
  assert.equal(snap.isFetched(S.rect('Inputs', 5, 2, 5, 2)), false);
  assert.equal(snap.isFetched(S.rect('Nope', 1, 1, 1, 1)), true, 'an unknown sheet has nothing to fetch');
});

test('apiRange quotes and clips', () => {
  const snap = new S.Snapshot(meta);
  assert.equal(S.apiRange(snap, S.rect('Inputs', 1, 1, S.MAX_ROW, 1)), "'Inputs'!A1:A1000");
  assert.equal(S.apiRange(snap, S.rect("O'Connor", 2, 3, 2, 3)), "'O''Connor'!C2");
  assert.equal(S.apiRange(snap, S.rect('Model', 1, 1, 3, 2)), "'Model'!A1:B3");
});
```

- [ ] **Step 2: Run to verify failure**

Run: `node extension/build.js --lib-only && node --test extension/test/adapter.test.js`
Expected: FAIL — `S.Snapshot is not a constructor`

- [ ] **Step 3: Write `extension/adapter.js`**

```js
/**
 * adapter.js — Sheets REST API v4 JSON, reshaped into the handful of
 * SpreadsheetApp calls Trace.gs makes (getSheetByName, getRange, getValues,
 * getFormula, getNamedRanges...). Same shape as the fake in run_e2e.js, so the
 * parser/resolver run unmodified in the browser.
 *
 * A cell that was never fetched reads as blank. That is deliberate: the
 * snapshot cannot call the network from inside a synchronous Trace.gs walk, so
 * audit.js fetches every rectangle first and asks isFetched() before trusting
 * a value. rowCount/columnCount stand in for getLastRow/getLastColumn — the
 * API does not expose a "used range", and over-reporting only costs a clip.
 */

var ERROR_TEXT = {
  DIV_0: '#DIV/0!', N_A: '#N/A', NAME: '#NAME?', NUM: '#NUM!',
  REF: '#REF!', VALUE: '#VALUE!', ERROR: '#ERROR!', NULL_VALUE: '#NULL!'
};
var BLANK_CELL = { v: '', d: '', f: '' };

function cellFromApi(v) {
  if (!v) return BLANK_CELL;
  var ev = v.effectiveValue || {}, raw;
  if (ev.errorValue) raw = ERROR_TEXT[ev.errorValue.type] || '#ERROR!';
  else if (ev.numberValue !== undefined && ev.numberValue !== null) raw = ev.numberValue;
  else if (ev.boolValue !== undefined && ev.boolValue !== null) raw = ev.boolValue;
  else if (ev.stringValue !== undefined && ev.stringValue !== null) raw = ev.stringValue;
  else raw = '';
  var ue = v.userEnteredValue || {};
  return {
    v: raw,
    d: v.formattedValue === undefined || v.formattedValue === null ? '' : String(v.formattedValue),
    f: ue.formulaValue || ''
  };
}

function Snapshot(meta) {
  this.sheets = [];
  this.byName = {};
  this.byId = {};
  var props = (meta && meta.sheets) || [];
  for (var i = 0; i < props.length; i++) {
    var s = new SnapSheet(this, props[i].properties || {});
    this.sheets.push(s);
    this.byName[s.title] = s;
    this.byId[s.sheetId] = s;
  }
  this.named = [];
  var nrs = (meta && meta.namedRanges) || [];
  for (var n = 0; n < nrs.length; n++) {
    var nr = nrs[n], g = nr.range || {}, sheet = this.byId[g.sheetId];
    if (!sheet) continue; // a name whose sheet is gone
    var r1 = (g.startRowIndex || 0) + 1, c1 = (g.startColumnIndex || 0) + 1;
    var r2 = g.endRowIndex === undefined || g.endRowIndex === null ? sheet.rowCount : g.endRowIndex;
    var c2 = g.endColumnIndex === undefined || g.endColumnIndex === null ? sheet.columnCount : g.endColumnIndex;
    this.named.push(new SnapNamedRange(nr.name, new SnapRange(sheet, r1, c1, r2 - r1 + 1, c2 - c1 + 1)));
  }
}
Snapshot.prototype.getSheets = function () { return this.sheets; };
Snapshot.prototype.getSheetByName = function (name) { return this.byName[name] || null; };
Snapshot.prototype.getNamedRanges = function () { return this.named; };

/** Merge a spreadsheets.get(includeGridData=true) response. */
Snapshot.prototype.addGridData = function (resp) {
  var sheets = (resp && resp.sheets) || [];
  for (var i = 0; i < sheets.length; i++) {
    var sheet = this.byId[(sheets[i].properties || {}).sheetId];
    if (!sheet) continue;
    var blocks = sheets[i].data || [];
    for (var b = 0; b < blocks.length; b++) {
      var blk = blocks[b], r0 = (blk.startRow || 0) + 1, c0 = (blk.startColumn || 0) + 1;
      var rows = blk.rowData || [];
      for (var r = 0; r < rows.length; r++) {
        var vals = rows[r].values || [];
        for (var c = 0; c < vals.length; c++) sheet.put(r0 + r, c0 + c, cellFromApi(vals[c]));
      }
    }
  }
};
/** The API trims trailing empty cells from a block, so "what did we ask for"
 *  is recorded by the caller, not inferred from the response. */
Snapshot.prototype.markFetched = function (r) {
  var s = this.byName[r.sheetName];
  if (s) s.fetched.push(r);
};
Snapshot.prototype.isFetched = function (r) {
  var s = this.byName[r.sheetName];
  if (!s) return true;
  for (var i = 0; i < s.fetched.length; i++) if (rectWithin(r, s.fetched[i])) return true;
  return false;
};

function SnapSheet(parent, props) {
  this.parent = parent;
  this.title = props.title || '';
  this.sheetId = props.sheetId;
  this.hidden = !!props.hidden;
  var g = props.gridProperties || {};
  this.rowCount = g.rowCount || 1000;
  this.columnCount = g.columnCount || 26;
  this.cells = {};   // row -> col -> {v, d, f}
  this.fetched = []; // rects we asked the API for
}
SnapSheet.prototype.put = function (r, c, cell) {
  if (!this.cells[r]) this.cells[r] = {};
  this.cells[r][c] = cell;
};
SnapSheet.prototype.cell = function (r, c) { return (this.cells[r] && this.cells[r][c]) || BLANK_CELL; };
SnapSheet.prototype.getName = function () { return this.title; };
SnapSheet.prototype.getSheetId = function () { return this.sheetId; };
SnapSheet.prototype.isSheetHidden = function () { return this.hidden; };
SnapSheet.prototype.getLastRow = function () { return this.rowCount; };
SnapSheet.prototype.getLastColumn = function () { return this.columnCount; };
SnapSheet.prototype.getMaxRows = function () { return this.rowCount; };
SnapSheet.prototype.getMaxColumns = function () { return this.columnCount; };
SnapSheet.prototype.getDataRange = function () { return new SnapRange(this, 1, 1, this.rowCount, this.columnCount); };
SnapSheet.prototype.getRange = function (a, b, c, d) {
  if (typeof a === 'string') {
    var r = a1ToRect(a, this.title);
    if (!r) throw new Error('Not a valid reference: ' + a);
    var r2 = Math.min(r.r2, this.rowCount), c2 = Math.min(r.c2, this.columnCount);
    return new SnapRange(this, r.r1, r.c1, r2 - r.r1 + 1, c2 - r.c1 + 1);
  }
  return new SnapRange(this, a, b, c || 1, d || 1);
};

function SnapRange(sheet, r1, c1, nr, nc) {
  this.sheet = sheet; this.r1 = r1; this.c1 = c1;
  this.nr = Math.max(nr, 0); this.nc = Math.max(nc, 0);
}
SnapRange.prototype.getSheet = function () { return this.sheet; };
SnapRange.prototype.getRow = function () { return this.r1; };
SnapRange.prototype.getColumn = function () { return this.c1; };
SnapRange.prototype.getLastRow = function () { return this.r1 + this.nr - 1; };
SnapRange.prototype.getLastColumn = function () { return this.c1 + this.nc - 1; };
SnapRange.prototype.getNumRows = function () { return this.nr; };
SnapRange.prototype.getNumColumns = function () { return this.nc; };
SnapRange.prototype.getA1Notation = function () {
  return rectToA1(rect(this.sheet.title, this.r1, this.c1, this.getLastRow(), this.getLastColumn()));
};
SnapRange.prototype.grid = function (pick) {
  var out = [];
  for (var r = 0; r < this.nr; r++) {
    var line = [];
    for (var c = 0; c < this.nc; c++) line.push(pick(this.sheet.cell(this.r1 + r, this.c1 + c)));
    out.push(line);
  }
  return out;
};
SnapRange.prototype.getValues = function () { return this.grid(function (x) { return x.v; }); };
SnapRange.prototype.getDisplayValues = function () { return this.grid(function (x) { return x.d; }); };
SnapRange.prototype.getFormulas = function () { return this.grid(function (x) { return x.f; }); };
SnapRange.prototype.getValue = function () { return this.sheet.cell(this.r1, this.c1).v; };
SnapRange.prototype.getDisplayValue = function () { return this.sheet.cell(this.r1, this.c1).d; };
SnapRange.prototype.getFormula = function () { return this.sheet.cell(this.r1, this.c1).f; };

function SnapNamedRange(name, range) { this.name = name; this.range = range; }
SnapNamedRange.prototype.getName = function () { return this.name; };
SnapNamedRange.prototype.getRange = function () { return this.range; };

function quoteSheetName(name) { return "'" + String(name).replace(/'/g, "''") + "'"; }

/** rect -> API range string, clipped to the grid so "A:A" never asks for ten million rows. */
function apiRange(snap, r) {
  var s = snap.getSheetByName(r.sheetName);
  var clipped = rect(r.sheetName, r.r1, r.c1,
    Math.min(r.r2, s ? s.rowCount : r.r2), Math.min(r.c2, s ? s.columnCount : r.c2));
  return quoteSheetName(r.sheetName) + '!' + rectToA1(clipped);
}
```

- [ ] **Step 4: Run the adapter tests**

Run: `node --test extension/test/adapter.test.js`
Expected: 8 passing.

- [ ] **Step 5: Commit**

```bash
git add extension/adapter.js extension/test/fixtures/meta.json extension/test/adapter.test.js
git commit -m "feat(ext): snapshot adapter shaped like SpreadsheetApp over Sheets API JSON"
```

---

### Task 3: `audit.js` — one audit against the snapshot

**Files:**
- Create: `extension/audit.js`, `extension/test/fake-api.js`, `extension/test/audit.test.js`

**Interfaces:**
- Consumes: `Snapshot`, `apiRange`, `buildNamedRangeMap`, `findPrecedents`, `ValueCache`, `extractRefs`, `refToRect`, `a1ToRect`, `rectToA1`, `numToCol`, `rectsOverlap`, `rectWithin`, `MAX_ROW`.
- Produces:
  - `auditCell(api, snap, sheetName, a1, formula, settings) -> Promise<AuditResult>` where `api = { getGrid(rangeStrings) -> Promise<gridJson> }`, `settings = { includeHidden, showNames, showExternal }`.
  - `AuditResult = { origin: { sheetName, a1, formula, names: string[] }, rows: Row[], unresolved: [{raw, reason}], hasBlank: boolean, unknownNames: string[] }`
  - `Row = { external, sheetName, flag: ''|'H'|'EX'|'?', address, value, subFormula, viaName, dynamic, isBlank, jumpable, url? }`
  - `rectsToFetch(formula, sheetName, a1, snap, namedRanges) -> rect[]`
  - `formatNumber(n)`, `safeUrl(text)`, `BLANK_LABEL`, `DEFAULT_SETTINGS`, `MAX_TOTAL_CELLS = 5000`, `MAX_FETCH_CELLS = 50000`.

- [ ] **Step 1: Write the fake API and the failing tests**

`extension/test/fake-api.js`:
```js
// An in-memory Sheets API: getGrid(ranges) answers from a cell table and records
// every range it was asked for, so tests can assert what an audit fetched.
// cells: { 'Sheet': { row: { col: {v, d, f} } } }  (d defaults to String(v), f to '')
function makeFakeApi(S, meta, cells) {
  const byId = {};
  for (const s of meta.sheets) byId[s.properties.title] = s.properties.sheetId;
  const calls = [];
  return {
    calls,
    getGrid(ranges) {
      calls.push(ranges.slice());
      const sheets = [];
      for (const rs of ranges) {
        const bang = rs.lastIndexOf('!');
        const sheetName = rs.substring(1, bang - 1).split("''").join("'");
        const r = S.a1ToRect(rs.substring(bang + 1), sheetName);
        const rowData = [];
        for (let row = r.r1; row <= r.r2; row++) {
          const values = [];
          for (let col = r.c1; col <= r.c2; col++) {
            const c = cells[sheetName] && cells[sheetName][row] && cells[sheetName][row][col];
            if (!c) { values.push({}); continue; }
            const v = {};
            if (typeof c.v === 'number') v.effectiveValue = { numberValue: c.v };
            else if (typeof c.v === 'boolean') v.effectiveValue = { boolValue: c.v };
            else if (c.v !== '' && c.v !== undefined) v.effectiveValue = { stringValue: String(c.v) };
            v.formattedValue = c.d !== undefined ? c.d : (c.v === undefined ? '' : String(c.v));
            if (c.f) v.userEnteredValue = { formulaValue: c.f };
            else if (v.effectiveValue) v.userEnteredValue = v.effectiveValue;
            values.push(v);
          }
          rowData.push({ values });
        }
        sheets.push({ properties: { sheetId: byId[sheetName] }, data: [{ startRow: r.r1 - 1, startColumn: r.c1 - 1, rowData }] });
      }
      return Promise.resolve({ sheets });
    }
  };
}
module.exports = { makeFakeApi };
```

`extension/test/audit.test.js`:
```js
const test = require('node:test');
const assert = require('node:assert/strict');
const { loadSandbox } = require('./load');
const { makeFakeApi } = require('./fake-api');
const meta = require('./fixtures/meta.json');

const S = loadSandbox();
const F = (f) => ({ v: 0, d: '0', f }); // a formula cell whose value does not matter

// Same scenario as run_e2e.js: Inputs holds the constants, Model the formulas.
const CELLS = {
  Inputs: {
    1: { 1: { v: 'Growth' }, 2: { v: 0.05 } },
    2: { 1: { v: 'Base' }, 2: { v: 100 } },
    3: { 1: { v: 'Shift' }, 2: { v: 1 } },
    4: { 1: { v: 'Empty' } }
  },
  'Hidden Calc': { 1: { 1: { v: 7 } }, 2: { 1: { v: 8 } }, 3: { 1: { v: 9 } } },
  Model: {
    1: { 1: F('=Inputs!B2*(1+Inputs!B1)') },
    2: { 1: F("=OFFSET('Hidden Calc'!A1,Inputs!B3,0)") },
    3: { 1: F('=INDEX(Inputs!B1:B4,4)') },
    5: { 1: F('=INDIRECT("Inputs!B2")') },
    7: { 1: F('=OFFSET(Inputs!B1,MATCH(1,Inputs!A:A,0),0)') },
    8: { 1: F("=OFFSET('Hidden Calc'!A1,Shift,0)") },
    9: { 1: F('=IMPORTRANGE("https://docs.google.com/spreadsheets/d/abcdefghijklmnopqrstuvwxyz","Sheet1!A1")+Inputs!B2') },
    10: { 1: F('=SUM(Inputs!B1:B3)') }
  }
};
function run(a1, settings) {
  const snap = new S.Snapshot(meta);
  const api = makeFakeApi(S, meta, CELLS);
  const formula = CELLS.Model[S.a1ToRect(a1, 'Model').r1][1].f;
  return S.auditCell(api, snap, 'Model', a1, formula, settings).then((res) => ({ res, api, snap }));
}
const addrs = (res) => res.rows.map((r) => r.sheetName + '!' + r.address);

test('plain cross-sheet refs: one fetch covering origin and both references', async () => {
  const { res, api } = await run('A1');
  assert.deepEqual(addrs(res), ['Inputs!B2', 'Inputs!B1']);
  assert.deepEqual(res.rows.map((r) => r.value), ['100', '0.05']);
  assert.equal(api.calls.length, 1);
  assert.deepEqual(api.calls[0].slice().sort(), ["'Inputs'!B1", "'Inputs'!B2", "'Model'!A1"]);
  assert.equal(res.hasBlank, false);
  assert.equal(res.origin.formula, '=Inputs!B2*(1+Inputs!B1)');
});

test('OFFSET into a hidden sheet: resolved target fetched second, hidden rows follow the setting', async () => {
  const shown = await run('A2', { includeHidden: true });
  assert.deepEqual(addrs(shown.res), ['Hidden Calc!A2', 'Hidden Calc!A1', 'Inputs!B3']);
  assert.deepEqual(shown.res.rows.map((r) => r.flag), ['H', 'H', '']);
  assert.equal(shown.res.rows[0].value, '8');
  assert.equal(shown.api.calls.length, 2);
  assert.deepEqual(shown.api.calls[1], ["'Hidden Calc'!A2"]);
  const hidden = await run('A2');
  assert.deepEqual(addrs(hidden.res), ['Inputs!B3'], 'hidden sheets are off by default');
});

test('INDEX onto an empty cell raises the blank alert; the range row shows its total', async () => {
  const { res } = await run('A3');
  assert.deepEqual(addrs(res), ['Inputs!B4', 'Inputs!B1:B4']);
  assert.equal(res.rows[0].isBlank, true);
  assert.equal(res.rows[0].value, S.BLANK_LABEL);
  assert.equal(res.rows[1].value, '101.05 (3)');
  assert.equal(res.hasBlank, true);
});

test('INDIRECT with a literal cannot be prefetched, so it costs a second call', async () => {
  const { res, api } = await run('A5');
  assert.deepEqual(addrs(res), ['Inputs!B2']);
  assert.equal(api.calls.length, 2);
});

test('an unresolvable OFFSET is reported, and the MATCH column was fetched clipped to the grid', async () => {
  const { res, api } = await run('A7');
  assert.equal(res.unresolved.length, 1);
  assert.equal(res.unresolved[0].raw, 'OFFSET(Inputs!B1,MATCH(1,Inputs!A:A,0),0)');
  assert.ok(api.calls[0].includes("'Inputs'!A1:A1000"));
});

test('a named range as an OFFSET argument resolves through the snapshot names', async () => {
  const { res } = await run('A8', { includeHidden: true });
  assert.deepEqual(addrs(res), ['Hidden Calc!A2', 'Hidden Calc!A1', 'Inputs!B3']);
  assert.equal(res.unresolved.length, 0);
  assert.deepEqual(res.unknownNames, []);
});

test('IMPORTRANGE is an external row, dropped when the setting is off', async () => {
  const on = await run('A9');
  assert.equal(on.res.rows[0].flag, 'EX');
  assert.equal(on.res.rows[0].jumpable, false);
  assert.equal(on.res.rows[0].url, 'https://docs.google.com/spreadsheets/d/abcdefghijklmnopqrstuvwxyz');
  assert.deepEqual(addrs(on.res).slice(1), ['Inputs!B2']);
  const off = await run('A9', { showExternal: false });
  assert.deepEqual(addrs(off.res), ['Inputs!B2']);
});

test('the origin reports the names covering it', async () => {
  const snap = new S.Snapshot(meta);
  const api = makeFakeApi(S, meta, CELLS);
  const res = await S.auditCell(api, snap, 'Inputs', 'B1', '', {});
  assert.deepEqual(res.origin.names, ['Growth']);
  assert.deepEqual(res.rows, [], 'not a formula cell');
  const off = await S.auditCell(api, snap, 'Inputs', 'B1', '', { showNames: false });
  assert.deepEqual(off.origin.names, []);
});

test('a name the snapshot does not know is reported so the caller can refresh metadata', async () => {
  const snap = new S.Snapshot(meta);
  const api = makeFakeApi(S, meta, CELLS);
  const res = await S.auditCell(api, snap, 'Model', 'A1', '=NewName+Inputs!B2', {});
  assert.deepEqual(res.unknownNames, ['NewName']);
});

test('every jumpable row was actually fetched before it was built', async () => {
  for (const a1 of ['A1', 'A2', 'A3', 'A5', 'A8', 'A10']) {
    const { res, snap } = await run(a1, { includeHidden: true });
    for (const row of res.rows.filter((r) => r.jumpable)) {
      const r = S.a1ToRect(row.address, row.sheetName);
      assert.equal(snap.isFetched(S.rect(row.sheetName, r.r1, r.c1, r.r1, r.c1)), true, a1 + ' ' + row.address);
    }
  }
});

test('a selection audits its top-left cell', async () => {
  const { res } = await run('A1:B3');
  assert.equal(res.origin.a1, 'A1');
});
```

- [ ] **Step 2: Run to verify failure**

Run: `node --test extension/test/audit.test.js`
Expected: FAIL — `S.auditCell is not a function`

- [ ] **Step 3: Write `extension/audit.js`**

```js
/**
 * audit.js — one audit of one cell. Port of scopionAuditCore + materializeRows
 * (Code.gs) without dependents, backgrounds, or the neighbour flag.
 *
 * The parser and resolver are synchronous and read through the snapshot, so
 * the order is: decide what they will need, fetch it, run them, fetch what the
 * resolver pointed at, build rows. Two API calls worst case.
 *
 * `api.getGrid(rangeStrings)` is injected so tests run without the network.
 */

var BLANK_LABEL = '---BLANK CELL---';
var MAX_TOTAL_CELLS = 5000;   // a range row past this shows "(N cells)" instead of a total
var MAX_FETCH_CELLS = 50000;  // a reference past this is fetched as its top-left cell only
var DEFAULT_SETTINGS = { includeHidden: false, showNames: true, showExternal: true };

function rectCells(r) { return (r.r2 - r.r1 + 1) * (r.c2 - r.c1 + 1); }
function topLeft(r) { return rect(r.sheetName, r.r1, r.c1, r.r1, r.c1); }

/** Clip to the sheet grid; shrink to the top-left cell when the block is too big to pull. */
function fetchRectFor(snap, r) {
  var s = snap.getSheetByName(r.sheetName);
  if (!s) return null;
  var clipped = rect(r.sheetName, r.r1, r.c1, Math.min(r.r2, s.rowCount), Math.min(r.c2, s.columnCount));
  if (clipped.r1 > clipped.r2 || clipped.c1 > clipped.c2) return null; // entirely off the grid
  return rectCells(clipped) <= MAX_FETCH_CELLS ? clipped : topLeft(clipped);
}

function dedupeRects(rs) {
  var seen = {}, out = [];
  for (var i = 0; i < rs.length; i++) {
    var r = rs[i];
    if (!r) continue;
    var k = [r.sheetName, r.r1, r.c1, r.r2, r.c2].join('|');
    if (seen[k]) continue;
    seen[k] = true;
    out.push(r);
  }
  return out;
}

/** The origin cell plus every static reference and known name in the formula. */
function rectsToFetch(formula, sheetName, a1, snap, namedRanges) {
  var out = [fetchRectFor(snap, a1ToRect(a1, sheetName))];
  var refs = formula ? extractRefs(formula) : [];
  for (var i = 0; i < refs.length; i++) {
    var r = refToRect(refs[i], sheetName, namedRanges);
    if (r) out.push(fetchRectFor(snap, r));
  }
  return dedupeRects(out);
}

/** Names used by the formula that the snapshot has never heard of. */
function unknownNamesIn(formula, namedRanges) {
  var out = [];
  var refs = formula ? extractRefs(formula) : [];
  for (var i = 0; i < refs.length; i++) {
    if (refs[i].kind === 'name' && !namedRanges[refs[i].a1.toUpperCase()]) out.push(refs[i].a1);
  }
  return out;
}

function fetchRects(api, snap, rects) {
  var todo = rects.filter(function (r) { return !snap.isFetched(r); });
  if (!todo.length) return Promise.resolve();
  return api.getGrid(todo.map(function (r) { return apiRange(snap, r); })).then(function (resp) {
    snap.addGridData(resp);
    todo.forEach(function (r) { snap.markFetched(r); });
  });
}

function auditCell(api, snap, sheetName, a1, formula, settings) {
  settings = Object.assign({}, DEFAULT_SETTINGS, settings || {});
  var namedRanges = buildNamedRangeMap(snap);
  var box = a1ToRect(a1, sheetName);
  if (!box) return Promise.reject(new Error('Not a valid reference: ' + a1));
  a1 = numToCol(box.c1) + box.r1; // a selection audits its top-left cell
  formula = formula && String(formula).charAt(0) === '=' ? String(formula) : '';

  return fetchRects(api, snap, rectsToFetch(formula, sheetName, a1, snap, namedRanges))
    .then(function () {
      var cache = new ValueCache(snap);
      var found = findPrecedents(snap, sheetName, a1, namedRanges, cache);
      // OFFSET/INDEX/INDIRECT results land wherever they point; pull the ones we do not hold.
      var extra = found.targets
        .filter(function (t) { return t.rect; })
        .map(function (t) { return fetchRectFor(snap, t.rect); });
      return fetchRects(api, snap, dedupeRects(extra)).then(function () {
        var rows = buildRows(snap, found.targets, cache, settings);
        return {
          origin: {
            sheetName: sheetName,
            a1: a1,
            formula: found.formula || formula,
            names: settings.showNames ? namesCovering(namedRanges, sheetName, a1) : []
          },
          rows: rows,
          unresolved: found.unresolved,
          hasBlank: rows.some(function (r) { return r.isBlank; }),
          // From the formula the page showed us, not the snapshot's: the caller
          // refreshes metadata on a miss, and the page is what the user is looking at.
          unknownNames: unknownNamesIn(formula, namedRanges)
        };
      });
    });
}

/** Port of materializeRows without backgrounds; the snapshot already holds the values. */
function buildRows(snap, targets, cache, settings) {
  var rows = [];
  for (var i = 0; i < targets.length; i++) {
    var t = targets[i];
    if (t.external) {
      if (!settings.showExternal) continue;
      rows.push({ external: true, sheetName: '(external)', flag: 'EX', address: t.a1, value: '',
        subFormula: t.raw || '', viaName: '', dynamic: false, url: safeUrl(t.url), isBlank: false, jumpable: false });
      continue;
    }
    var rct = t.rect;
    var sheet = snap.getSheetByName(rct.sheetName);
    if (!sheet) {
      rows.push({ external: false, sheetName: rct.sheetName, flag: '?', address: rectToA1(rct), value: '(sheet not found)',
        subFormula: t.raw || '', viaName: '', dynamic: !!t.dynamic, isBlank: false, jumpable: false });
      continue;
    }
    if (sheet.isSheetHidden() && !settings.includeHidden) continue;
    var flag = sheet.isSheetHidden() ? 'H' : '';
    var isRange = rct.r1 !== rct.r2 || rct.c1 !== rct.c2;
    if (rct.r1 > sheet.getMaxRows() || rct.c1 > sheet.getMaxColumns()) {
      // Past the grid there is nothing to read — and an empty referenced range is
      // exactly the modelling error the blank alert exists for.
      rows.push({ external: false, sheetName: sheet.getName(), flag: flag, address: rectToA1(rct), value: BLANK_LABEL,
        subFormula: t.raw || '', viaName: t.viaName || '', dynamic: !!t.dynamic, isBlank: true, jumpable: false });
      continue;
    }
    var cell = sheet.cell(rct.r1, rct.c1);
    // "Blank" means structurally empty. A formula returning "" is not a missing input.
    var isBlank = !isRange && cell.d === '' && cell.f === '';
    rows.push({
      external: false, sheetName: sheet.getName(), flag: flag, address: rectToA1(rct),
      value: isBlank ? BLANK_LABEL : (isRange ? rangeTotal(cache, rct) : cell.d),
      subFormula: t.raw || '', viaName: t.viaName || '', dynamic: !!t.dynamic,
      isBlank: isBlank, jumpable: true
    });
  }
  return rows;
}

/** Defined names whose range covers the audited cell (ACE's txtNames box). */
function namesCovering(namedRanges, sheetName, a1) {
  var cell = a1ToRect(a1, sheetName);
  var out = [];
  for (var key in namedRanges) {
    if (rectsOverlap(namedRanges[key].rect, cell)) out.push(namedRanges[key].name);
  }
  return out;
}

/** A range row shows what the block adds up to, or an em dash when nothing in it is numeric. */
function rangeTotal(cache, rct) {
  var cells = rectCells(rct);
  if (cells > MAX_TOTAL_CELLS) return '(' + formatNumber(cells) + ' cells)';
  var t = cache.sumRect(rct);
  if (!t.count) return '—';
  return formatNumber(t.sum) + ' (' + t.count + ')';
}

function formatNumber(n) {
  if (typeof n !== 'number' || !isFinite(n)) return String(n);
  if (n !== 0 && (Math.abs(n) >= 1e15 || Math.abs(n) < 1e-4)) return n.toExponential(2);
  var rounded = Math.round(n * 100) / 100;
  var parts = String(rounded).split('.');
  parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return parts.join('.');
}

/** IMPORTRANGE's first argument is spreadsheet content: only a plain https link or a bare key is handed back. */
function safeUrl(url) {
  var text = String(url || '').trim();
  if (/^https:\/\/[^\s"'<>]+$/.test(text)) return text;
  if (/^[A-Za-z0-9_-]{20,}$/.test(text)) return 'https://docs.google.com/spreadsheets/d/' + text;
  return '';
}
```

- [ ] **Step 4: Run the audit tests**

Run: `node --test extension/test/audit.test.js`
Expected: 11 passing. If `INDIRECT` or the `MATCH` case fails, the cause is almost always a rectangle that was not fetched — print `api.calls` and compare against `rectsToFetch`.

- [ ] **Step 5: Run everything and commit**

Run: `node run_tests.js && node run_e2e.js && node --test extension/test/`
Expected: all green.

```bash
git add extension/audit.js extension/test/fake-api.js extension/test/audit.test.js
git commit -m "feat(ext): audit a cell through the snapshot with two fetches at most"
```

---

### Task 4: `state.js` — pure walk state

**Files:**
- Create: `extension/state.js`, `extension/test/state.test.js`

**Interfaces:**
- Produces:
  - `createWalk(origin, rows) -> Walk` where `origin = { sheetName, a1 }`, `Walk = { origin, rows, active: number, history: origin[] }`
  - `walkMove(walk, delta) -> Walk` (clamped; from -1 the first move lands on row 0)
  - `walkSelect(walk, index) -> Walk`
  - `walkDrill(walk, newOrigin, newRows) -> Walk` (pushes the old origin)
  - `walkBack(walk) -> { origin, history } | null`
  - `walkRoot(walk) -> origin` (where OK returns to: the first origin of the session)

- [ ] **Step 1: Write the failing tests**

`extension/test/state.test.js`:
```js
const test = require('node:test');
const assert = require('node:assert/strict');
const { loadSandbox } = require('./load');
const S = loadSandbox();

const o = (a1) => ({ sheetName: 'Model', a1 });
const rows = [{ address: 'B1' }, { address: 'B2' }, { address: 'B3' }];

test('a new walk highlights the first row when there is one', () => {
  assert.equal(S.createWalk(o('A1'), rows).active, 0);
  assert.equal(S.createWalk(o('A1'), []).active, -1);
});

test('move clamps at both ends and never goes negative', () => {
  let w = S.createWalk(o('A1'), rows);
  w = S.walkMove(w, -1); assert.equal(w.active, 0);
  w = S.walkMove(w, 5); assert.equal(w.active, 2);
  w = S.walkMove(w, -1); assert.equal(w.active, 1);
  assert.equal(S.walkMove(S.createWalk(o('A1'), []), 1).active, -1);
  assert.equal(S.walkSelect(w, 2).active, 2);
  assert.equal(S.walkSelect(w, 9).active, 1, 'out of range keeps the current row');
});

test('drill pushes the old origin; back pops it; root is the first origin', () => {
  let w = S.createWalk(o('A1'), rows);
  w = S.walkDrill(w, o('B2'), [{ address: 'C1' }]);
  assert.deepEqual(w.origin, o('B2'));
  assert.deepEqual(w.history, [o('A1')]);
  assert.equal(w.active, 0);
  w = S.walkDrill(w, o('C1'), []);
  assert.deepEqual(S.walkRoot(w), o('A1'));
  const back = S.walkBack(w);
  assert.deepEqual(back, { origin: o('B2'), history: [o('A1')] });
  assert.equal(S.walkBack(S.createWalk(o('A1'), rows)), null);
});

test('state functions do not mutate their input', () => {
  const w = S.createWalk(o('A1'), rows);
  S.walkMove(w, 1); S.walkDrill(w, o('B2'), []);
  assert.equal(w.active, 0); assert.deepEqual(w.history, []);
});
```

- [ ] **Step 2: Run to verify failure**

Run: `node --test extension/test/state.test.js`
Expected: FAIL — `S.createWalk is not a function`

- [ ] **Step 3: Write `extension/state.js`**

```js
/**
 * state.js — the walk, without the DOM: which row is highlighted and how we
 * got here. ACE kept this in module globals (PastHistory, myTest); keeping it
 * pure here is what makes it testable without a browser.
 */
function createWalk(origin, rows) {
  return { origin: origin, rows: rows, active: rows.length ? 0 : -1, history: [] };
}
function walkMove(w, delta) {
  if (!w.rows.length) return w;
  var from = w.active < 0 ? 0 : w.active;
  var next = Math.max(0, Math.min(w.rows.length - 1, from + delta));
  return Object.assign({}, w, { active: next });
}
function walkSelect(w, index) {
  if (index < 0 || index >= w.rows.length) return w;
  return Object.assign({}, w, { active: index });
}
/** New Origin: the highlighted target becomes the origin; where we stood is remembered. */
function walkDrill(w, newOrigin, newRows) {
  return { origin: newOrigin, rows: newRows, active: newRows.length ? 0 : -1, history: w.history.concat([w.origin]) };
}
/** Back: the origin to re-audit and the history that remains, or null at the root. */
function walkBack(w) {
  if (!w.history.length) return null;
  return { origin: w.history[w.history.length - 1], history: w.history.slice(0, -1) };
}
function walkRoot(w) { return w.history.length ? w.history[0] : w.origin; }
```

- [ ] **Step 4: Run and commit**

Run: `node --test extension/test/state.test.js` → 4 passing.

```bash
git add extension/state.js extension/test/state.test.js
git commit -m "feat(ext): pure walk state with history"
```

---

### Task 5: `dom.js` — Sheets page adapters

**Files:**
- Create: `extension/dom.js`, `extension/test/dom.test.js`

**Interfaces:**
- Produces (pure): `spreadsheetIdFromPath(pathname) -> string|null`, `parseNameBox(value) -> {a1}|{name}|null`, `unionRects(rects) -> {x,y,w,h}|null`, `topLeftA1(address) -> "B5"`.
- Produces (page): `SHEETS_SEL`, `sheetsSelfCheck(doc) -> string[]` (missing selectors), `SheetsDom = { spreadsheetId(), activeSheetName(), selection(), formula(), cellRect(), sheetTabs(), jump(sheetName, a1) -> Promise<void> }`.
- Nothing in this file touches `document`/`location` at load time (Node loads it).

- [ ] **Step 1: Write the failing tests**

`extension/test/dom.test.js`:
```js
const test = require('node:test');
const assert = require('node:assert/strict');
const { loadSandbox } = require('./load');
const S = loadSandbox();

test('spreadsheet id from the path', () => {
  assert.equal(S.spreadsheetIdFromPath('/spreadsheets/d/11GMudQJ1VmzTgscq4sY82ceb1lHywWKisegzQPQC3P0/edit'), '11GMudQJ1VmzTgscq4sY82ceb1lHywWKisegzQPQC3P0');
  assert.equal(S.spreadsheetIdFromPath('/spreadsheets/u/0/'), null);
});

test('the name box shows an address, a range, or a name', () => {
  assert.deepEqual(S.parseNameBox('B5'), { a1: 'B5' });
  assert.deepEqual(S.parseNameBox('$b$5'), { a1: 'B5' });
  assert.deepEqual(S.parseNameBox('A1:C3'), { a1: 'A1' });
  assert.deepEqual(S.parseNameBox('Growth'), { name: 'Growth' });
  assert.equal(S.parseNameBox(''), null);
});

test('union of border pieces', () => {
  const r = S.unionRects([
    { left: 45, top: 165, right: 147, bottom: 167 }, { left: 45, top: 185, right: 147, bottom: 187 },
    { left: 45, top: 165, right: 47, bottom: 187 }, { left: 145, top: 165, right: 147, bottom: 187 }
  ]);
  assert.deepEqual(r, { x: 45, y: 165, w: 102, h: 22 });
  assert.equal(S.unionRects([]), null);
});

test('top-left of a range address', () => {
  assert.equal(S.topLeftA1('B2:D9'), 'B2');
  assert.equal(S.topLeftA1('C7'), 'C7');
  assert.equal(S.topLeftA1('A:A'), 'A1');
});

test('self check names what is missing', () => {
  const fake = { querySelector: (sel) => (sel === '#t-name-box' ? {} : null) };
  assert.deepEqual(S.sheetsSelfCheck(fake), ['#t-formula-bar-input', '.docs-sheet-tab']);
});
```

- [ ] **Step 2: Run to verify failure**

Run: `node --test extension/test/dom.test.js` → FAIL `S.spreadsheetIdFromPath is not a function`

- [ ] **Step 3: Write `extension/dom.js`**

```js
/**
 * dom.js — everything that touches the Google Sheets page lives here, and
 * nowhere else. The grid is a canvas, so the only readable state is the name
 * box, the formula bar, the selection border overlay and the sheet tabs; the
 * only way to move the selection is to type into the name box.
 * Verified against the live page on 2026-09-03; if Sheets renames a selector,
 * sheetsSelfCheck() says which one and the panel shows it.
 */
var SHEETS_SEL = {
  nameBox: '#t-name-box',
  formulaBar: '#t-formula-bar-input',
  activeBorder: '.active-cell-border',
  tab: '.docs-sheet-tab',
  tabName: '.docs-sheet-tab-name',
  activeTab: '.docs-sheet-active-tab'
};
var JUMP_SETTLE_MS = 300;

function spreadsheetIdFromPath(pathname) {
  var m = /\/spreadsheets\/d\/([A-Za-z0-9_-]+)/.exec(pathname || '');
  return m ? m[1] : null;
}

/** "B5" / "$B$5" / "A1:C3" -> {a1: top-left}; anything else is a named range's name. */
function parseNameBox(value) {
  var v = String(value || '').trim();
  if (!v) return null;
  var m = /^\$?([A-Za-z]{1,3})\$?([0-9]{1,8})(?::\$?[A-Za-z]{1,3}\$?[0-9]{1,8})?$/.exec(v);
  if (m) return { a1: m[1].toUpperCase() + m[2] };
  return { name: v };
}

function unionRects(rects) {
  if (!rects.length) return null;
  var x1 = Infinity, y1 = Infinity, x2 = -Infinity, y2 = -Infinity;
  for (var i = 0; i < rects.length; i++) {
    var r = rects[i];
    x1 = Math.min(x1, r.left); y1 = Math.min(y1, r.top);
    x2 = Math.max(x2, r.right); y2 = Math.max(y2, r.bottom);
  }
  return { x: x1, y: y1, w: x2 - x1, h: y2 - y1 };
}

function topLeftA1(address) {
  var r = a1ToRect(address, '');
  return r ? numToCol(r.c1) + r.r1 : address;
}

function sheetsSelfCheck(doc) {
  var missing = [];
  [SHEETS_SEL.nameBox, SHEETS_SEL.formulaBar, SHEETS_SEL.tab].forEach(function (sel) {
    if (!doc.querySelector(sel)) missing.push(sel);
  });
  return missing;
}

var SheetsDom = {
  spreadsheetId: function () { return spreadsheetIdFromPath(location.pathname); },
  activeSheetName: function () {
    var el = document.querySelector(SHEETS_SEL.activeTab + ' ' + SHEETS_SEL.tabName);
    return el ? el.textContent.trim() : null;
  },
  selection: function () {
    var nb = document.querySelector(SHEETS_SEL.nameBox);
    return nb ? parseNameBox(nb.value) : null;
  },
  /** The formula text, or '' for a constant cell (the bar shows the value then). */
  formula: function () {
    var fb = document.querySelector(SHEETS_SEL.formulaBar);
    var t = fb ? fb.textContent : '';
    return t.charAt(0) === '=' ? t : '';
  },
  cellRect: function () {
    var els = document.querySelectorAll(SHEETS_SEL.activeBorder);
    var rects = [];
    for (var i = 0; i < els.length; i++) rects.push(els[i].getBoundingClientRect());
    return unionRects(rects);
  },
  sheetTabs: function () {
    var tabs = document.querySelectorAll(SHEETS_SEL.tab), out = [];
    for (var i = 0; i < tabs.length; i++) {
      var n = tabs[i].querySelector(SHEETS_SEL.tabName);
      out.push({ name: n ? n.textContent.trim() : '', hidden: tabs[i].offsetParent === null,
        active: tabs[i].classList.contains('docs-sheet-active-tab') });
    }
    return out;
  },
  /**
   * Move the grid selection by typing into the name box. Sheets rewrites the
   * box with the landed address, so "value no longer what we typed" is the
   * settle signal; the timeout covers a rejected reference.
   */
  jump: function (sheetName, a1) {
    var nb = document.querySelector(SHEETS_SEL.nameBox);
    if (!nb) return Promise.reject(new Error('Sheets name box not found'));
    var target = quoteSheetName(sheetName) + '!' + a1;
    nb.focus();
    nb.value = target;
    nb.dispatchEvent(new Event('input', { bubbles: true }));
    ['keydown', 'keypress', 'keyup'].forEach(function (type) {
      nb.dispatchEvent(new KeyboardEvent(type, { key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true, cancelable: true }));
    });
    return new Promise(function (resolve) {
      var t0 = Date.now();
      (function poll() {
        if (nb.value !== target || Date.now() - t0 > JUMP_SETTLE_MS) return resolve();
        setTimeout(poll, 30);
      })();
    });
  }
};
```

- [ ] **Step 4: Run and commit**

Run: `node --test extension/test/dom.test.js` → 5 passing.

```bash
git add extension/dom.js extension/test/dom.test.js
git commit -m "feat(ext): sheets page adapters — name box, formula bar, selection rect, jump"
```

---

### Task 6: `background.js` — shortcut and Sheets API proxy

**Files:**
- Create: `extension/background.js`, `extension/test/background.test.js`

**Interfaces:**
- Produces: `metaUrl(spreadsheetId)`, `gridUrl(spreadsheetId, ranges)`, `apiGet(url, deps) -> Promise<json>` with `deps = { fetch, getToken(interactive) -> Promise<string>, removeToken(token) -> Promise }`.
- Messages handled: `{type:'scopion:meta', spreadsheetId}` and `{type:'scopion:grid', spreadsheetId, ranges}` → `sendResponse({ok:true,data}|{ok:false,error})`. Command `toggle-scopion` → `chrome.tabs.sendMessage(tabId, {type:'scopion:toggle'})`.

- [ ] **Step 1: Write the failing tests**

`extension/test/background.test.js`:
```js
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const vm = require('vm');
const path = require('path');

function load() {
  const sandbox = { console, Promise, Error, JSON, encodeURIComponent, String, Object, Array, Math, Date };
  vm.createContext(sandbox);
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'background.js'), 'utf8'), sandbox, { filename: 'background.js' });
  return sandbox;
}

test('urls ask for exactly the fields the adapter reads', () => {
  const B = load();
  const u = new URL(B.metaUrl('abc'));
  assert.equal(u.pathname, '/v4/spreadsheets/abc');
  assert.equal(u.searchParams.get('fields'), 'namedRanges,sheets.properties(sheetId,title,hidden,gridProperties(rowCount,columnCount))');
  const g = new URL(B.gridUrl('abc', ["'Inputs'!B1", "'O''Connor'!A1:B2"]));
  assert.equal(g.searchParams.get('includeGridData'), 'true');
  assert.deepEqual(g.searchParams.getAll('ranges'), ["'Inputs'!B1", "'O''Connor'!A1:B2"]);
  assert.equal(g.searchParams.get('fields'), 'sheets(properties.sheetId,data(startRow,startColumn,rowData.values(effectiveValue,formattedValue,userEnteredValue)))');
});

function fakeDeps(statuses, tokens) {
  const log = [];
  return {
    log,
    fetch: async (url, opts) => {
      const status = statuses.shift();
      log.push(['fetch', opts.headers.Authorization, status]);
      return { status, ok: status >= 200 && status < 300, json: async () => ({ status }), text: async () => 'body ' + status };
    },
    getToken: async (interactive) => { log.push(['getToken', interactive]); const t = tokens.shift(); if (!t) throw new Error('no token'); return t; },
    removeToken: async (t) => { log.push(['removeToken', t]); }
  };
}

test('a good token is used once, silently', async () => {
  const B = load();
  const deps = fakeDeps([200], ['t1']);
  assert.deepEqual(await B.apiGet('u', deps), { status: 200 });
  assert.deepEqual(deps.log, [['getToken', false], ['fetch', 'Bearer t1', 200]]);
});

test('a rejected token is dropped and the call retried once with an interactive token', async () => {
  const B = load();
  const deps = fakeDeps([401, 200], ['stale', 'fresh']);
  assert.deepEqual(await B.apiGet('u', deps), { status: 200 });
  assert.deepEqual(deps.log, [['getToken', false], ['fetch', 'Bearer stale', 401], ['removeToken', 'stale'], ['getToken', true], ['fetch', 'Bearer fresh', 200]]);
});

test('no cached token yet: fall through to the interactive prompt', async () => {
  const B = load();
  const deps = fakeDeps([200], [null, 'first']);
  await B.apiGet('u', deps);
  assert.deepEqual(deps.log.slice(0, 2), [['getToken', false], ['getToken', true]]);
});

test('other errors surface with the status and a slice of the body', async () => {
  const B = load();
  await assert.rejects(B.apiGet('u', fakeDeps([403], ['t'])), /Sheets API 403: body 403/);
});
```

- [ ] **Step 2: Run to verify failure**

Run: `node --test extension/test/background.test.js` → FAIL (`background.js` missing)

- [ ] **Step 3: Write `extension/background.js`**

```js
/**
 * background.js — the service worker does two things the content script
 * cannot: receive the keyboard command, and call the Sheets API with an OAuth
 * token (chrome.identity is not available to content scripts). It holds no
 * state; every request carries its spreadsheet id.
 */
var SHEETS_API = 'https://sheets.googleapis.com/v4/spreadsheets/';
var META_FIELDS = 'namedRanges,sheets.properties(sheetId,title,hidden,gridProperties(rowCount,columnCount))';
var GRID_FIELDS = 'sheets(properties.sheetId,data(startRow,startColumn,rowData.values(effectiveValue,formattedValue,userEnteredValue)))';

function metaUrl(spreadsheetId) {
  return SHEETS_API + encodeURIComponent(spreadsheetId) + '?fields=' + encodeURIComponent(META_FIELDS);
}
function gridUrl(spreadsheetId, ranges) {
  return SHEETS_API + encodeURIComponent(spreadsheetId) + '?includeGridData=true&fields=' + encodeURIComponent(GRID_FIELDS) +
    ranges.map(function (r) { return '&ranges=' + encodeURIComponent(r); }).join('');
}

/** GET with a cached token; on 401 drop the token and retry once interactively. */
function apiGet(url, deps) {
  function call(token) { return deps.fetch(url, { headers: { Authorization: 'Bearer ' + token } }); }
  return deps.getToken(false).catch(function () { return deps.getToken(true); })
    .then(function (token) {
      return call(token).then(function (res) {
        if (res.status !== 401) return res;
        return deps.removeToken(token).then(function () { return deps.getToken(true); }).then(call);
      });
    })
    .then(function (res) {
      if (res.ok) return res.json();
      return res.text().then(function (body) { throw new Error('Sheets API ' + res.status + ': ' + String(body).slice(0, 200)); });
    });
}

if (typeof chrome !== 'undefined' && chrome.runtime && chrome.runtime.onMessage) {
  var chromeDeps = {
    fetch: function (u, o) { return fetch(u, o); },
    getToken: function (interactive) {
      return new Promise(function (resolve, reject) {
        chrome.identity.getAuthToken({ interactive: interactive }, function (result) {
          var token = result && typeof result === 'object' ? result.token : result;
          if (chrome.runtime.lastError || !token) reject(new Error(chrome.runtime.lastError ? chrome.runtime.lastError.message : 'no token'));
          else resolve(token);
        });
      });
    },
    removeToken: function (token) {
      return new Promise(function (resolve) { chrome.identity.removeCachedAuthToken({ token: token }, resolve); });
    }
  };

  chrome.runtime.onMessage.addListener(function (msg, sender, sendResponse) {
    if (!msg || !msg.type) return false;
    var p;
    if (msg.type === 'scopion:meta') p = apiGet(metaUrl(msg.spreadsheetId), chromeDeps);
    else if (msg.type === 'scopion:grid') p = apiGet(gridUrl(msg.spreadsheetId, msg.ranges || []), chromeDeps);
    else return false;
    p.then(function (data) { sendResponse({ ok: true, data: data }); },
           function (e) { sendResponse({ ok: false, error: String(e && e.message || e) }); });
    return true; // keep the channel open for the async response
  });

  chrome.commands.onCommand.addListener(function (command, tab) {
    if (command === 'toggle-scopion' && tab && tab.id !== undefined) {
      chrome.tabs.sendMessage(tab.id, { type: 'scopion:toggle' }, function () { void chrome.runtime.lastError; });
    }
  });
}
```

- [ ] **Step 4: Run and commit**

Run: `node --test extension/test/background.test.js` → 5 passing.

```bash
git add extension/background.js extension/test/background.test.js
git commit -m "feat(ext): background worker — shortcut relay and Sheets API proxy with token retry"
```

---

### Task 7: `panel.js` — the ACE strip

**Files:**
- Create: `extension/panel.js`

**Interfaces:**
- Consumes: nothing from earlier tasks except `formatNumber` is *not* needed (values arrive as strings).
- Produces: `createPanel(handlers) -> Panel`.
  - `handlers = { onWalk(index), onDrill(), onBack(), onClose(), onNewOrigin(), onAdvanced(), onSetting(key, value), onExternal(index) }`
  - `Panel = { open(nearRect), close(), render(view), focus(), setPosition({x,y}), getPosition(), highlight(rect|null, kind:'walk'|'origin'), notice(text|null), isOpen() }`
  - `view = { originFormula, names: string[], rows: Row[], active: number, hasBlank, unresolved: [{raw,reason}], canBack, advanced: boolean, settings: {includeHidden, showNames, showExternal}, busy: boolean }`
- No automated test (DOM). Verified in Task 10.

- [ ] **Step 1: Write `extension/panel.js`**

Geometry is ACE's, in px (1 pt ≈ 1.33 px): strip 910×110 (reduced) / 910×151 (advanced). List 467 px wide; origin formula at x=473; RANGE NAME at y=80; OK at x=832,y=72; advanced row at y=120. Column widths: basic `[200, 0, 120, 147, 0]`, with flags or dynamic calls `[147, 27, 80, 107, 107]` (ACE's `ChangeColumns`).

```js
/**
 * panel.js — the ACE window (frmACE), rebuilt: a modeless strip you can drag,
 * a multi-column list of precedents on the left, the origin formula and the
 * range names on the right, OK/Back/New Origin/Advanced along the bottom.
 * Shadow DOM keeps Sheets' CSS out and ours in. No framework: the list is
 * small and re-rendered whole.
 */
var PANEL_CSS = [
  ':host{all:initial}',
  '.win{position:fixed;z-index:2147483000;width:910px;background:#f0f0f0;border:1px solid #6d6d6d;box-shadow:0 6px 24px rgba(0,0,0,.35);',
  '  font:11px Tahoma,"Segoe UI",system-ui,sans-serif;color:#000;user-select:none}',
  '.title{height:22px;line-height:22px;padding:0 8px;background:linear-gradient(#fdfdfd,#dcdcdc);border-bottom:1px solid #a0a0a0;cursor:move;display:flex;justify-content:space-between}',
  '.title .x{cursor:pointer;padding:0 6px}.title .x:hover{background:#c42b1c;color:#fff}',
  '.body{position:relative;height:88px}.body.adv{height:129px}',
  '.list{position:absolute;left:0;top:0;width:467px;height:82px;overflow-y:auto;background:#fff;border:1px inset #999;outline:none}',
  '.list.blank{background:#ffff99}.list.unfocused .row.on{background:#d9d9d9;color:#000}',
  '.row{display:grid;height:16px;line-height:16px;white-space:nowrap;cursor:default}',
  '.row>span{overflow:hidden;text-overflow:ellipsis;padding:0 3px}.row .v{text-align:right;font-variant-numeric:tabular-nums}',
  '.row.on{background:#0078d7;color:#fff}.row.ext{color:#5a3d8a}.row .flag{text-align:center}',
  '.origin{position:absolute;left:473px;top:0;width:430px;height:60px;background:#fff;border:1px inset #999;padding:2px 4px;overflow:auto;white-space:pre-wrap;word-break:break-all;font-family:Consolas,Menlo,monospace}',
  '.names{position:absolute;left:473px;top:64px;width:350px;height:18px;background:#fff;border:1px inset #999;padding:0 4px;line-height:18px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}',
  'button{position:absolute;height:22px;padding:0 8px;font:inherit;background:linear-gradient(#fff,#e1e1e1);border:1px solid #707070;border-radius:2px;cursor:pointer}',
  'button:disabled{color:#888;cursor:default}button:focus{outline:1px dotted #000}',
  '.ok{left:832px;top:62px;width:70px}.adv-row{position:absolute;top:100px;left:0;right:0;height:26px;display:none}.body.adv .adv-row{display:block}',
  '.adv-row label{position:absolute;top:5px;white-space:nowrap}.adv-row input{vertical-align:-2px}',
  '.back{left:632px;top:100px;width:70px}.advbtn{left:744px;top:100px;width:80px}.neworigin{left:832px;top:100px;width:70px}',
  '.body:not(.adv) .back,.body:not(.adv) .advbtn,.body:not(.adv) .neworigin{top:64px;display:none}',
  '.notice{position:absolute;left:473px;top:84px;width:350px;color:#8a4b00;font-size:10px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}',
  '.busy .list{opacity:.6}',
  '.hl{position:fixed;z-index:2147482999;pointer-events:none;box-sizing:border-box}',
  '.hl.walk{background:rgba(0,176,80,.28);border:2px solid #00b050}.hl.origin{border:2px dashed #00b050}'
].join('\n');

var COLS_BASIC = [200, 0, 120, 147, 0];
var COLS_FULL = [147, 27, 80, 107, 107];

function createPanel(handlers) {
  var host = document.createElement('div');
  host.id = 'scopion-host';
  var root = host.attachShadow({ mode: 'open' });
  var style = document.createElement('style');
  style.textContent = PANEL_CSS;
  root.appendChild(style);

  var win = document.createElement('div');
  win.className = 'win';
  win.hidden = true;
  win.innerHTML =
    '<div class="title"><span>Scopion — Active Cell Explorer</span><span class="x" title="OK (Esc)">✕</span></div>' +
    '<div class="body">' +
      '<div class="list" tabindex="0" role="listbox" aria-label="Precedents"></div>' +
      '<div class="origin"></div>' +
      '<div class="names"></div>' +
      '<div class="notice"></div>' +
      '<button class="ok">OK</button>' +
      '<div class="adv-row">' +
        '<label style="left:8px"><input type="checkbox" data-key="showExternal"> Display Workbook information</label>' +
        '<label style="left:224px"><input type="checkbox" data-key="showNames"> Display Range Name</label>' +
        '<label style="left:408px"><input type="checkbox" data-key="includeHidden"> Include Hidden Sheets</label>' +
      '</div>' +
      '<button class="back">Back</button>' +
      '<button class="advbtn">Advanced</button>' +
      '<button class="neworigin">New Origin</button>' +
    '</div>';
  root.appendChild(win);

  var walkHl = document.createElement('div'); walkHl.className = 'hl walk'; walkHl.hidden = true;
  var originHl = document.createElement('div'); originHl.className = 'hl origin'; originHl.hidden = true;
  root.appendChild(walkHl); root.appendChild(originHl);

  var $ = function (sel) { return win.querySelector(sel); };
  var list = $('.list'), body = $('.body');
  var pos = { x: 80, y: 120 };
  var view = null;

  function esc(s) { return String(s === undefined || s === null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }

  function renderList() {
    var rows = view.rows || [];
    var anyFlag = rows.some(function (r) { return r.flag; });
    var anyDyn = rows.some(function (r) { return r.dynamic || r.external; });
    var cols = anyFlag || anyDyn ? COLS_FULL : COLS_BASIC;
    var tpl = cols.map(function (w) { return w + 'px'; }).join(' ');
    list.className = 'list' + (view.hasBlank ? ' blank' : '');
    list.innerHTML = rows.length ? '' : '<div class="row" style="grid-template-columns:1fr"><span>' +
      esc(view.originFormula ? 'No references found.' : 'Not a formula cell.') + '</span></div>';
    rows.forEach(function (r, i) {
      var el = document.createElement('div');
      el.className = 'row' + (i === view.active ? ' on' : '') + (r.external ? ' ext' : '');
      el.style.gridTemplateColumns = tpl;
      el.setAttribute('role', 'option');
      el.dataset.index = i;
      el.innerHTML =
        '<span title="' + esc(r.sheetName) + '">' + esc(r.sheetName) + '</span>' +
        '<span class="flag">' + esc(r.flag) + '</span>' +
        '<span>' + esc(r.address) + '</span>' +
        '<span class="v">' + esc(r.value) + '</span>' +
        '<span title="' + esc(r.subFormula) + '">' + esc(r.dynamic || r.external ? r.subFormula : '') + '</span>';
      list.appendChild(el);
    });
    var on = list.querySelector('.row.on');
    if (on && on.scrollIntoView) on.scrollIntoView({ block: 'nearest' });
  }

  function render(v) {
    view = v;
    body.className = 'body' + (v.advanced ? ' adv' : '') + (v.busy ? ' busy' : '');
    $('.origin').textContent = v.originFormula ? 'ORIGIN FORMULA: ' + v.originFormula : 'ORIGIN FORMULA: (not a formula)';
    $('.names').textContent = v.names && v.names.length ? 'RANGE NAME: ' + v.names.join(' | ') : '';
    $('.back').disabled = !v.canBack;
    $('.advbtn').textContent = v.advanced ? 'Simple' : 'Advanced';
    ['showExternal', 'showNames', 'includeHidden'].forEach(function (k) {
      $('input[data-key="' + k + '"]').checked = !!(v.settings && v.settings[k]);
    });
    if (v.unresolved && v.unresolved.length) {
      notice(v.unresolved.length + ' dynamic reference' + (v.unresolved.length > 1 ? 's' : '') + ' unresolved: ' +
        v.unresolved.map(function (u) { return u.raw + ' (' + u.reason + ')'; }).join('; '));
    }
    renderList();
  }

  function notice(text) { $('.notice').textContent = text || ''; $('.notice').title = text || ''; }

  function place(p) {
    var w = 910, h = view && view.advanced ? 151 : 110;
    pos = { x: Math.max(0, Math.min(p.x, innerWidth - w)), y: Math.max(0, Math.min(p.y, innerHeight - h)) };
    win.style.left = pos.x + 'px'; win.style.top = pos.y + 'px';
  }

  // Keyboard — the ACE ListBox: arrows walk (and jump), Enter is New Origin, Back is Backspace/←, Esc is OK.
  list.addEventListener('keydown', function (ev) {
    var k = ev.key;
    if (!view) return;
    if (k === 'ArrowDown') handlers.onWalk(Math.min(view.rows.length - 1, Math.max(0, view.active) + (view.active < 0 ? 0 : 1)));
    else if (k === 'ArrowUp') handlers.onWalk(Math.max(0, view.active - 1));
    else if (k === 'Home') handlers.onWalk(0);
    else if (k === 'End') handlers.onWalk(view.rows.length - 1);
    else if (k === 'PageDown') handlers.onWalk(Math.min(view.rows.length - 1, view.active + 5));
    else if (k === 'PageUp') handlers.onWalk(Math.max(0, view.active - 5));
    else if (k === 'Enter' || k === 'ArrowRight') handlers.onDrill();
    else if (k === 'Backspace' || k === 'ArrowLeft') handlers.onBack();
    else if (k === 'Escape') handlers.onClose();
    else return;
    ev.preventDefault(); ev.stopPropagation();
  });
  list.addEventListener('click', function (ev) {
    var row = ev.target.closest('.row');
    if (!row || row.dataset.index === undefined) return;
    var i = Number(row.dataset.index);
    if (view.rows[i] && view.rows[i].external) handlers.onExternal(i); else handlers.onWalk(i);
    list.focus();
  });
  list.addEventListener('focus', function () { list.classList.remove('unfocused'); });
  list.addEventListener('blur', function () { list.classList.add('unfocused'); });
  win.addEventListener('keydown', function (ev) { if (ev.key === 'Escape') { handlers.onClose(); ev.preventDefault(); } });

  $('.ok').onclick = handlers.onClose;
  $('.x').onclick = handlers.onClose;
  $('.back').onclick = handlers.onBack;
  $('.neworigin').onclick = handlers.onNewOrigin;
  $('.advbtn').onclick = handlers.onAdvanced;
  win.querySelectorAll('input[data-key]').forEach(function (cb) {
    cb.onchange = function () { handlers.onSetting(cb.dataset.key, cb.checked); };
  });

  // Drag by the title bar; the position is the caller's to persist (getPosition()).
  var drag = null;
  $('.title').addEventListener('mousedown', function (ev) {
    if (ev.target.classList.contains('x')) return;
    drag = { dx: ev.clientX - pos.x, dy: ev.clientY - pos.y };
    ev.preventDefault();
  });
  window.addEventListener('mousemove', function (ev) { if (drag) place({ x: ev.clientX - drag.dx, y: ev.clientY - drag.dy }); });
  window.addEventListener('mouseup', function () { if (drag) { drag = null; if (handlers.onMoved) handlers.onMoved(pos); } });

  function setHl(el, rect) {
    if (!rect) { el.hidden = true; return; }
    el.hidden = false;
    el.style.left = rect.x + 'px'; el.style.top = rect.y + 'px';
    el.style.width = rect.w + 'px'; el.style.height = rect.h + 'px';
  }

  document.documentElement.appendChild(host);

  return {
    isOpen: function () { return !win.hidden; },
    open: function (nearRect, savedPos) {
      win.hidden = false;
      if (savedPos) place(savedPos);
      else if (nearRect) place({ x: nearRect.x, y: nearRect.y + nearRect.h + 8 });
      else place(pos);
    },
    close: function () { win.hidden = true; setHl(walkHl, null); setHl(originHl, null); notice(''); },
    render: render,
    focus: function () { list.focus(); },
    setPosition: place,
    getPosition: function () { return pos; },
    highlight: function (rect, kind) { setHl(kind === 'origin' ? originHl : walkHl, rect); },
    notice: notice
  };
}
```

- [ ] **Step 2: Syntax check and commit**

Run: `node --check extension/panel.js && node run_tests.js && node --test extension/test/`
Expected: no syntax error; suites green (panel.js is not loaded by the sandbox).

```bash
git add extension/panel.js
git commit -m "feat(ext): the ACE strip as a shadow-dom panel"
```

---

### Task 8: `content.js` — wiring

**Files:**
- Create: `extension/content.js`

**Interfaces:**
- Consumes: `Snapshot`, `auditCell`, `DEFAULT_SETTINGS`, `createWalk/walkMove/walkSelect/walkDrill/walkBack/walkRoot`, `SheetsDom`, `sheetsSelfCheck`, `topLeftA1`, `createPanel`, `a1ToRect`, `numToCol`, `buildNamedRangeMap`.
- Produces: the running feature. Listens for `{type:'scopion:toggle'}`.

- [ ] **Step 1: Write `extension/content.js`**

```js
/**
 * content.js — glue. The shortcut arrives as a message from background.js;
 * everything else is: read the selection, audit it, show the strip, and turn
 * key presses into jumps. Sheets moves keyboard focus to its grid after every
 * jump, so the panel takes it back each time.
 */
(function () {
  if (window.__scopionLoaded) return; // Chrome re-injects on extension reload
  window.__scopionLoaded = true;

  var S = {
    spreadsheetId: null, snap: null, settings: null, panel: null,
    walk: null, advanced: false, busy: false, lastJump: null, unhidden: [], savedPos: null, tick: null
  };

  function rpc(msg) {
    return new Promise(function (resolve, reject) {
      chrome.runtime.sendMessage(msg, function (r) {
        if (chrome.runtime.lastError) return reject(new Error(chrome.runtime.lastError.message));
        if (!r || !r.ok) return reject(new Error(r ? r.error : 'no response from background'));
        resolve(r.data);
      });
    });
  }
  var api = { getGrid: function (ranges) { return rpc({ type: 'scopion:grid', spreadsheetId: S.spreadsheetId, ranges: ranges }); } };

  function ensureSnapshot(force) {
    if (S.snap && !force) return Promise.resolve(S.snap);
    return rpc({ type: 'scopion:meta', spreadsheetId: S.spreadsheetId }).then(function (meta) {
      S.snap = new Snapshot(meta);
      return S.snap;
    });
  }

  function storageGet(area, keys) {
    return new Promise(function (resolve) { chrome.storage[area].get(keys, function (v) { resolve(v || {}); }); });
  }
  function loadPrefs() {
    return Promise.all([storageGet('sync', ['settings']), storageGet('local', ['panelPos'])]).then(function (r) {
      S.settings = Object.assign({}, DEFAULT_SETTINGS, r[0].settings || {});
      S.savedPos = r[1].panelPos || null;
    });
  }

  function view(extra) {
    var w = S.walk;
    return Object.assign({
      originFormula: w ? w.originFormula : '', names: w ? w.names : [], rows: w ? w.rows : [], active: w ? w.active : -1,
      hasBlank: w ? w.hasBlank : false, unresolved: w ? w.unresolved : [], canBack: !!(w && w.history.length),
      advanced: S.advanced, settings: S.settings, busy: S.busy
    }, extra || {});
  }
  function render() { S.panel.render(view()); }
  function setBusy(b) { S.busy = b; if (S.panel.isOpen()) render(); }
  function fail(e) { setBusy(false); S.panel.notice('Scopion: ' + (e && e.message ? e.message : e)); }

  /** Where the grid selection is right now, as {sheetName, a1}. A named cell is resolved through the snapshot. */
  function currentCell() {
    var sheetName = SheetsDom.activeSheetName();
    var sel = SheetsDom.selection();
    if (!sheetName || !sel) throw new Error('Select a cell first.');
    if (sel.a1) return { sheetName: sheetName, a1: sel.a1 };
    var map = buildNamedRangeMap(S.snap);
    var hit = map[sel.name.toUpperCase()];
    if (!hit) throw new Error('Cannot resolve "' + sel.name + '" to a cell.');
    return { sheetName: hit.rect.sheetName, a1: numToCol(hit.rect.c1) + hit.rect.r1 };
  }

  /** Audit a cell. Refresh metadata once when the formula names something the snapshot does not know. */
  function audit(cell, formula, retried) {
    return ensureSnapshot().then(function (snap) {
      return auditCell(api, snap, cell.sheetName, cell.a1, formula, S.settings);
    }).then(function (res) {
      if (res.unknownNames.length && !retried) return ensureSnapshot(true).then(function () { return audit(cell, formula, true); });
      return res;
    });
  }

  function walkFrom(res, history) {
    var w = createWalk({ sheetName: res.origin.sheetName, a1: res.origin.a1 }, res.rows);
    w.history = history || [];
    w.originFormula = res.origin.formula; w.names = res.origin.names; w.hasBlank = res.hasBlank; w.unresolved = res.unresolved;
    return w;
  }

  /** ACE's CheckXLNone: if the user clicked elsewhere in the grid, go back to where the walk was before acting. */
  function resync() {
    if (!S.lastJump) return Promise.resolve();
    var now;
    try { now = currentCell(); } catch (e) { return Promise.resolve(); }
    if (now.sheetName === S.lastJump.sheetName && now.a1 === S.lastJump.a1) return Promise.resolve();
    return jump(S.lastJump);
  }

  function jump(cell) {
    var sheet = S.snap && S.snap.getSheetByName(cell.sheetName);
    var wasHidden = sheet && sheet.isSheetHidden() && S.unhidden.indexOf(cell.sheetName) < 0;
    return SheetsDom.jump(cell.sheetName, cell.a1).then(function () {
      S.lastJump = cell;
      if (wasHidden) {
        S.unhidden.push(cell.sheetName);
        S.panel.notice('Sheets unhid "' + cell.sheetName + '" to show it. Re-hide it by hand when you are done.');
      }
      paintHighlights();
      S.panel.focus();
    });
  }

  function paintHighlights() {
    var rect = SheetsDom.cellRect();
    var w = S.walk;
    var onOrigin = w && S.lastJump && S.lastJump.sheetName === w.origin.sheetName && S.lastJump.a1 === w.origin.a1;
    S.panel.highlight(rect && !onOrigin ? rect : null, 'walk');
    S.panel.highlight(rect && onOrigin ? rect : null, 'origin');
  }

  function open() {
    setBusy(true);
    return loadPrefs().then(function () { return ensureSnapshot(); }).then(function () {
      var cell = currentCell();
      return audit(cell, SheetsDom.formula()).then(function (res) {
        S.walk = walkFrom(res, []);
        S.lastJump = { sheetName: res.origin.sheetName, a1: res.origin.a1 };
        S.unhidden = [];
        setBusy(false);
        S.panel.open(SheetsDom.cellRect(), S.savedPos);
        render();
        paintHighlights();
        S.panel.focus();
        if (S.tick) clearInterval(S.tick);
        S.tick = setInterval(paintHighlights, 250); // the grid scrolls under us; follow the cell
      });
    }).catch(function (e) {
      // A failure before the strip is up must still be visible somewhere.
      if (!S.panel.isOpen()) { S.panel.open(SheetsDom.cellRect(), S.savedPos); S.panel.render(view()); }
      fail(e);
    });
  }

  function walkTo(index) {
    if (S.busy || !S.walk) return;
    var row = S.walk.rows[index];
    if (!row) return;
    S.walk = walkSelect(S.walk, index);
    render();
    if (!row.jumpable) return;
    setBusy(true);
    resync().then(function () { return jump({ sheetName: row.sheetName, a1: topLeftA1(row.address) }); })
      .then(function () { setBusy(false); }).catch(fail);
  }

  /** New Origin / Enter: audit the cell the walk is standing on. */
  function drill() {
    if (S.busy || !S.walk) return;
    var row = S.walk.rows[S.walk.active];
    if (!row || !row.jumpable) return;
    var cell = { sheetName: row.sheetName, a1: topLeftA1(row.address) };
    setBusy(true);
    resync().then(function () { return jump(cell); })
      .then(function () { return audit(cell, SheetsDom.formula()); })
      .then(function (res) {
        var next = walkDrill(S.walk, cell, res.rows);
        S.walk = walkFrom(res, next.history);
        setBusy(false); render(); paintHighlights(); S.panel.focus();
      }).catch(fail);
  }

  function newOrigin() {
    if (S.busy) return;
    setBusy(true);
    var cell;
    try { cell = currentCell(); } catch (e) { return fail(e); }
    jump(cell).then(function () { return audit(cell, SheetsDom.formula()); })
      .then(function (res) {
        var history = S.walk ? S.walk.history.concat([S.walk.origin]) : [];
        S.walk = walkFrom(res, history);
        setBusy(false); render(); paintHighlights(); S.panel.focus();
      }).catch(fail);
  }

  function back() {
    if (S.busy || !S.walk) return;
    var b = walkBack(S.walk);
    if (!b) return;
    setBusy(true);
    resync().then(function () { return jump(b.origin); })
      .then(function () { return audit(b.origin, SheetsDom.formula()); })
      .then(function (res) {
        S.walk = walkFrom(res, b.history);
        setBusy(false); render(); paintHighlights(); S.panel.focus();
      }).catch(fail);
  }

  /** OK / Esc: return to where the session started, clear the overlays, close. */
  function close() {
    if (S.tick) { clearInterval(S.tick); S.tick = null; }
    var root = S.walk ? walkRoot(S.walk) : null;
    chrome.storage.local.set({ panelPos: S.panel.getPosition() });
    var p = root ? SheetsDom.jump(root.sheetName, root.a1) : Promise.resolve();
    p.then(function () {
      S.panel.close();
      S.walk = null; S.lastJump = null; S.busy = false;
      if (S.unhidden.length) console.info('Scopion left these sheets visible: ' + S.unhidden.join(', '));
    });
  }

  function setSetting(key, value) {
    S.settings[key] = value;
    chrome.storage.sync.set({ settings: S.settings });
    // Settings change what the list shows, so re-audit the current origin (ACE: ChangeColumns + rebuild).
    if (!S.walk) return render();
    setBusy(true);
    audit(S.walk.origin, S.walk.originFormula).then(function (res) {
      S.walk = walkFrom(res, S.walk.history);
      setBusy(false); render();
    }).catch(fail);
  }

  function init() {
    S.spreadsheetId = SheetsDom.spreadsheetId();
    if (!S.spreadsheetId) return;
    S.panel = createPanel({
      onWalk: walkTo,
      onDrill: drill,
      onBack: back,
      onClose: close,
      onNewOrigin: newOrigin,
      onAdvanced: function () { S.advanced = !S.advanced; render(); },
      onSetting: setSetting,
      onExternal: function (i) { var r = S.walk && S.walk.rows[i]; if (r && r.url) window.open(r.url, '_blank', 'noopener'); },
      onMoved: function (pos) { chrome.storage.local.set({ panelPos: pos }); }
    });
    chrome.runtime.onMessage.addListener(function (msg) {
      if (!msg || msg.type !== 'scopion:toggle') return;
      var missing = sheetsSelfCheck(document);
      if (missing.length) { S.panel.open(null, S.savedPos); S.panel.render(view()); S.panel.notice('Sheets layout changed; missing ' + missing.join(', ')); return; }
      if (!S.panel.isOpen()) open(); else newOrigin();
    });
  }

  init();
})();
```

- [ ] **Step 2: Syntax check and commit**

Run: `node --check extension/content.js && node --test extension/test/`

```bash
git add extension/content.js
git commit -m "feat(ext): wire shortcut, audit, walk, drill, back and close"
```

---

### Task 9: README section — building, credentials, loading unpacked

**Files:**
- Modify: `README.md` (append a section after "Getting `Cmd+Shift+A` back")

- [ ] **Step 1: Append to README.md**

````markdown
## Chrome extension (Scopion)

The extension reproduces the ACE window next to the selected cell and ships its
own shortcut. It reads the sheet through the Sheets API, so it needs an OAuth
client that you create once. Nothing is ever written to the spreadsheet.

1. **Key** (makes the extension ID stable, which the OAuth client is tied to):
   ```bash
   openssl genrsa -out extension/scopion.pem 2048
   openssl rsa -in extension/scopion.pem -pubout -outform DER | openssl base64 -A; echo
   ```
   The printed base64 is the `key` value. Extension ID:
   ```bash
   openssl rsa -in extension/scopion.pem -pubout -outform DER | node -e "const c=require('crypto');let b=[];process.stdin.on('data',d=>b.push(d)).on('end',()=>{const h=c.createHash('sha256').update(Buffer.concat(b)).digest('hex').slice(0,32);console.log([...h].map(x=>String.fromCharCode(97+parseInt(x,16))).join(''))})"
   ```
2. **Google Cloud**: create a project → *APIs & Services ▸ Library* → enable
   **Google Sheets API** → *OAuth consent screen*: External, publishing status
   *Testing*, add your Google account as a test user, add the scope
   `https://www.googleapis.com/auth/spreadsheets.readonly` → *Credentials ▸ Create
   credentials ▸ OAuth client ID*, application type **Chrome Extension**, Item ID =
   the extension ID from step 1. Copy the client ID.
3. `cp extension/oauth.example.json extension/oauth.local.json` and fill in
   `client_id` and `key`. This file is git-ignored.
4. `node extension/build.js` → writes `extension/lib/` and `extension/manifest.json`.
5. `chrome://extensions` → Developer mode → **Load unpacked** → the `extension/`
   folder. Check the ID matches step 1.
6. Open a spreadsheet, select a formula cell, press **Ctrl+Shift+A**. The first
   run asks for Google sign-in and read-only access. Rebind the key at
   `chrome://extensions/shortcuts` if it clashes.

Keys: ↑/↓ walk (the selection follows), Enter = New Origin, Backspace/← = Back,
Esc = OK. "Include Hidden Sheets" is off by default because jumping into a hidden
sheet makes Sheets unhide it; the panel says so when it happens.

Tests: `node extension/build.js --lib-only && node --test extension/test/`.
Publishing on the Chrome Web Store additionally needs Google's verification of
the sensitive `spreadsheets.readonly` scope.
````

- [ ] **Step 2: Commit**

```bash
git add README.md
git commit -m "docs: chrome extension setup"
```

---

### Task 10: Manual smoke test on the live sheet, then record the result

**Files:**
- Modify: `tests/README.md` (append the extension smoke checklist), `memory.md` (state, verification commands)

Prerequisites: the user has completed README steps 1–5 (credentials are theirs to create; the agent must not enter them).

- [ ] **Step 1: Smoke checklist** — run on `https://docs.google.com/spreadsheets/d/11GMudQJ1VmzTgscq4sY82ceb1lHywWKisegzQPQC3P0/edit`, sheet **Model**, and tick each line:

```
[ ] Select B6 (=B1*(1+Growth)); Ctrl+Shift+A → strip appears below-right of B6 within ~1 s
[ ] List shows Model!B1 = 100 and Inputs-side Growth row; ORIGIN FORMULA shows the formula
[ ] ↓ moves the highlight to the second row AND the grid selection moves there; green overlay on the cell
[ ] ↑ returns; origin cell shows the dashed outline when standing on it
[ ] Enter on B1 → list re-audits from B1 ("Not a formula cell." with empty list), Back enabled
[ ] Backspace → back to B6 with the original list
[ ] Select B7 (=OFFSET('Hidden Calc'!A1,Shift,0)); Ctrl+Shift+A while open → New Origin re-audits
[ ] Hidden rows absent by default; Advanced → tick Include Hidden Sheets → Hidden Calc!A2 (H) appears
[ ] ↓ onto Hidden Calc!A2 → sheet switches, notice says Sheets unhid it
[ ] Esc → selection returns to B7, overlays gone, strip closed
[ ] Right-click the Hidden Calc tab → シートを非表示 (restore the test sheet)
[ ] Drag the strip by its title; Esc; reopen → same position
[ ] Click a cell elsewhere, then press ↓ in the strip → the walk first returns to its last cell (resync), then moves
```

- [ ] **Step 2: Append the checklist to `tests/README.md`** under a heading `## Extension smoke test`, verbatim.

- [ ] **Step 3: Update `memory.md`** (English, ≤60 lines total): current state = extension built and smoke-tested (or which lines failed), verification commands (`node extension/build.js --lib-only && node --test extension/test/`), the Sheets DOM facts table from the spec, open items (re-hide needs write scope; Web Store verification; neighbour flag not yet ported).

- [ ] **Step 4: Commit**

```bash
git add tests/README.md memory.md
git commit -m "docs: extension smoke test and project memory"
```

---

## Self-review notes

- Spec coverage: architecture (T1–T8), one-audit flow incl. two fetches (T3), navigation rules incl. resync/hidden notice/root return (T8), keyboard map (T7), panel geometry + column switching (T7), auth + credentials outside git (T1, T6, T9), failure modes: self-check (T5/T8), settle wait (T5), focus retake (T8), quota (T3 ≤2 calls), large ranges (T3 `MAX_FETCH_CELLS`/`MAX_TOTAL_CELLS`), tests (T1–T6), manual smoke (T10).
- Not covered on purpose (spec "later"): re-hide, neighbour flag port, Firefox, Web Store.
- Names used across tasks: `Snapshot`, `apiRange`, `auditCell`, `DEFAULT_SETTINGS`, `createWalk/walkMove/walkSelect/walkDrill/walkBack/walkRoot`, `SheetsDom`, `sheetsSelfCheck`, `topLeftA1`, `createPanel`, `metaUrl/gridUrl/apiGet` — consistent between definition and use.
