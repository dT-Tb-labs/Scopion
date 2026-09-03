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
