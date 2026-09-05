/**
 * state.js — the walk, without the DOM: which row is highlighted and how we
 * got here. ACE kept this in module globals (PastHistory, myTest); keeping it
 * pure here is what makes it testable without a browser.
 */
function createWalk(origin, rows) {
  return { origin: origin, rows: rows, active: rows.length ? 0 : -1, history: [] };
}
/**
 * ACE's list begins with the origin cell itself (PopulateList calls FindName
 * on oCheck before walking the arrows), so ↑ from the first precedent lands
 * back on the origin. Row 0 is that cell; the precedents follow.
 */
function withOriginRow(origin, rows) {
  return [{
    external: false, sheetName: origin.sheetName, flag: '', address: origin.a1,
    value: origin.value === undefined ? '' : origin.value, subFormula: '', viaName: '',
    dynamic: false, isBlank: false, jumpable: true, isOrigin: true
  }].concat(rows);
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
