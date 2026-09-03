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
