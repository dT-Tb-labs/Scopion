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
