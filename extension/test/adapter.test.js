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
