/**
 * End-to-end harness: runs aceAudit / aceJump against a fake spreadsheet, so the
 * half of the add-on that talks to SpreadsheetApp is exercised without a live
 * document. Complements run_tests.js, which only covers formula parsing.
 *
 *   node run_e2e.js
 */
const fs = require('fs');
const vm = require('vm');

// --- minimal SpreadsheetApp stand-in ---------------------------------------
function colToA(n) { let s = ''; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; }
function parseA1(a1) {
  const m = /^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/i.exec(a1.replace(/\$/g, ''));
  if (!m) throw new Error('mock getRange: unsupported a1 "' + a1 + '"');
  const col = (s) => s.toUpperCase().split('').reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
  return { r1: +m[2], c1: col(m[1]), r2: m[4] ? +m[4] : +m[2], c2: m[3] ? col(m[3]) : col(m[1]) };
}

class Range {
  constructor(sheet, r1, c1, nr, nc) { this.sheet = sheet; this.r1 = r1; this.c1 = c1; this.nr = nr; this.nc = nc; }
  _cell(dr, dc) { return this.sheet._cell(this.r1 + dr, this.c1 + dc); }
  getSheet() { return this.sheet; }
  getRow() { return this.r1; } getColumn() { return this.c1; }
  getLastRow() { return this.r1 + this.nr - 1; } getLastColumn() { return this.c1 + this.nc - 1; }
  getCell(r, c) { return new Range(this.sheet, this.r1 + r - 1, this.c1 + c - 1, 1, 1); }
  getA1Notation() {
    const a = colToA(this.c1) + this.r1, b = colToA(this.c1 + this.nc - 1) + (this.r1 + this.nr - 1);
    return a === b ? a : a + ':' + b;
  }
  getFormula() { return this._cell(0, 0).formula; }
  getValue() { return this._cell(0, 0).value; }
  getDisplayValue() { const v = this._cell(0, 0).value; return v === '' || v === null ? '' : String(v); }
  getBackground() { return this._cell(0, 0).background; }
  _grid(pick) {
    const out = [];
    for (let r = 0; r < this.nr; r++) { const row = [];
      for (let c = 0; c < this.nc; c++) row.push(pick(this._cell(r, c)));
      out.push(row); }
    return out;
  }
  getFormulas() { return this._grid((c) => c.formula); }
  getValues() { return this._grid((c) => c.value); }
  getDisplayValues() { return this._grid((c) => (c.value === '' || c.value === null ? '' : String(c.value))); }
  getBackgrounds() { return this._grid((c) => c.background); }
}

class Sheet {
  constructor(name, cells, hidden) { this.name = name; this.cells = cells; this.hidden = !!hidden; }
  _cell(r, c) { return (this.cells[r] && this.cells[r][c]) || { formula: '', value: '', background: '#ffffff' }; }
  getName() { return this.name; }
  isSheetHidden() { return this.hidden; }
  showSheet() { this.hidden = false; } hideSheet() { this.hidden = true; }
  activate() { this.parent.active = this; }
  getLastRow() { return Math.max(0, ...Object.keys(this.cells).map(Number)); }
  getMaxRows() { return this.maxRows || 1000; }      // a new Google sheet is 1000 x 26
  getMaxColumns() { return this.maxCols || 26; }
  getLastColumn() {
    let max = 0;
    for (const r of Object.keys(this.cells)) max = Math.max(max, ...Object.keys(this.cells[r]).map(Number));
    return max;
  }
  getDataRange() { return new Range(this, 1, 1, Math.max(this.getLastRow(), 1), Math.max(this.getLastColumn(), 1)); }
  getRange(a, b, c, d) {
    if (typeof a === 'string') { const p = parseA1(a); return new Range(this, p.r1, p.c1, p.r2 - p.r1 + 1, p.c2 - p.c1 + 1); }
    return new Range(this, a, b, c || 1, d || 1);
  }
}

class Spreadsheet {
  constructor(sheets, names) {
    this.sheets = sheets; sheets.forEach((s) => { s.parent = this; });
    this.names = names || []; this.active = sheets[0]; this.activeRange = sheets[0].getRange('A1');
  }
  getId() { return 'mock'; }
  getSheets() { return this.sheets; }
  getSheetByName(n) { return this.sheets.find((s) => s.name === n) || null; }
  getActiveSheet() { return this.active; }
  getActiveRange() { return this.activeRange; }
  setActiveRange(r) { this.activeRange = r; this.active = r.sheet; return r; }
  getNamedRanges() { return this.names; }
}

/** cells: {"1": {"1": "=A2", ...}} — a string is a formula if it starts with "=". */
function sheetOf(name, spec, hidden) {
  const cells = {};
  for (const r of Object.keys(spec)) {
    cells[r] = {};
    for (const c of Object.keys(spec[r])) {
      const raw = spec[r][c];
      const isFormula = typeof raw === 'string' && raw.charAt(0) === '=';
      cells[r][c] = {
        formula: isFormula ? raw : '',
        value: isFormula ? (spec[r][c + '.value'] !== undefined ? spec[r][c + '.value'] : 0) : raw,
        background: '#ffffff'
      };
    }
  }
  return new Sheet(name, cells, hidden);
}

// --- load the add-on --------------------------------------------------------
const props = {};
const sandbox = {
  console, Date, Math, JSON, String, Number, Array, Object, RegExp, isNaN, parseInt,
  Logger: { log: () => {} },
  PropertiesService: {
    getDocumentProperties: () => ({
      getProperty: (k) => (k in props ? props[k] : null),
      setProperty: (k, v) => { props[k] = v; }
    })
  },
  SpreadsheetApp: { getActive: () => sandbox.__ss }
};
vm.createContext(sandbox);
for (const f of ['Formula.gs', 'Trace.gs', 'Code.gs']) {
  vm.runInContext(fs.readFileSync(f, 'utf8'), sandbox, { filename: f });
}

// --- scenario ---------------------------------------------------------------
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
  { getName: () => 'Growth', getRange: () => inputs.getRange('B1') }
]);
sandbox.__ss = ss;

// --- assertions -------------------------------------------------------------
let failed = 0;
function check(label, actual, expected) {
  const a = JSON.stringify(actual), e = JSON.stringify(expected);
  if (a === e) { console.log('  ok   ' + label); return; }
  failed++;
  console.log('  FAIL ' + label + '\n    expected: ' + e + '\n    actual:   ' + a);
}
function addrs(res) { return res.rows.map((r) => r.sheetName + '!' + r.address); }

console.log('precedents');
let res = sandbox.aceAudit({ sheetName: 'Model', a1: 'A1', mode: 'precedents' });
check('plain cross-sheet refs', addrs(res), ['Inputs!B2', 'Inputs!B1']);
check('no blank alert', res.hasBlank, false);

res = sandbox.aceAudit({ sheetName: 'Model', a1: 'A2', mode: 'precedents' });
check('OFFSET into a hidden sheet resolves', addrs(res),
  ["Hidden Calc!A1", 'Inputs!B3', "Hidden Calc!A2"]);
check('the hidden sheet is flagged H', res.rows.filter((r) => r.flag === 'H').map((r) => r.address), ['A1', 'A2']);

res = sandbox.aceAudit({ sheetName: 'Model', a1: 'A3', mode: 'precedents' });
check('INDEX target', addrs(res), ['Inputs!B1:B4', 'Inputs!B4']);
check('an empty INDEX target raises the blank alert', res.hasBlank, true);
check('and is labelled', res.rows[1].value, '---BLANK CELL---');

res = sandbox.aceAudit({ sheetName: 'Model', a1: 'A5', mode: 'precedents' });
check('INDIRECT with a literal', addrs(res), ['Inputs!B2']);

res = sandbox.aceAudit({ sheetName: 'Model', a1: 'A6', mode: 'precedents' });
check('SUM is not double-counted', addrs(res), ['Inputs!B1:B3']);

res = sandbox.aceAudit({ sheetName: 'Model', a1: 'A7', mode: 'precedents' });
check('an unresolvable OFFSET is reported, not dropped', res.unresolved.length, 1);
check('with the offending text', res.unresolved[0].raw, 'OFFSET(Inputs!B1,MATCH(1,Inputs!A:A,0),0)');

console.log('named ranges');
res = sandbox.aceAudit({ sheetName: 'Inputs', a1: 'B1', mode: 'precedents' });
check('the covering name is reported', res.namedRanges, ['Growth']);

console.log('dependents');
res = sandbox.aceAudit({ sheetName: 'Inputs', a1: 'B2', mode: 'dependents' });
check('every reader of Inputs!B2, including the ranges covering it',
  addrs(res).sort(), ['Model!A1', 'Model!A3', 'Model!A5', 'Model!A6'].sort());
res = sandbox.aceAudit({ sheetName: 'Model', a1: 'A1', mode: 'dependents' });
check('a same-sheet reader', addrs(res), ['Model!A4']);
res = sandbox.aceAudit({ sheetName: 'Inputs', a1: 'B3', mode: 'dependents' });
check('a reader through an OFFSET argument', addrs(res).sort(),
  ['Model!A2', 'Model!A3', 'Model!A6'].sort());

console.log('settings');
sandbox.aceSetSetting('traverseHidden', false);
res = sandbox.aceAudit({ sheetName: 'Model', a1: 'A2', mode: 'precedents' });
check('hidden-sheet rows drop out when the toggle is off', addrs(res), ['Inputs!B3']);
sandbox.aceSetSetting('traverseHidden', true);

console.log('navigation');
const jump = sandbox.aceJump('Hidden Calc', 'A2');
check('jumping unhides', jump.unhidden, true);
check('and reports the cell', jump.a1, 'A2');
check('the selection moved', ss.getActiveRange().getA1Notation(), 'A2');
check('re-hiding works', sandbox.aceRehide(['Hidden Calc']), []); // active sheet is skipped
sandbox.aceJump('Model', 'A1');
check('re-hiding once off the sheet', sandbox.aceRehide(['Hidden Calc']), ['Hidden Calc']);

console.log('grid bounds');
let outOfGrid = '';
try { sandbox.aceJump('Model', 'A5000'); } catch (e) { outOfGrid = e.message; }
check('a cell past the grid is refused instead of throwing inside getRange',
  outOfGrid.indexOf('outside the grid') > 0, true);

console.log('input validation');
let threw = '';
try { sandbox.aceJump('Model', 'DROP TABLE'); } catch (e) { threw = e.message; }
check('a bogus address is rejected', threw.indexOf('Not a valid reference') === 0, true);
threw = '';
try { sandbox.aceAudit({ sheetName: 'Nope', a1: 'A1' }); } catch (e) { threw = e.message; }
check('a bogus sheet is rejected', threw.indexOf('No such sheet') === 0, true);

console.log(failed ? '\n' + failed + ' e2e check(s) FAILED' : '\nall e2e checks passed');
process.exit(failed ? 1 : 0);
