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
  DIVIDE_BY_ZERO: '#DIV/0!', N_A: '#N/A', NAME: '#NAME?', NUM: '#NUM!',
  REF: '#REF!', VALUE: '#VALUE!', ERROR: '#ERROR!', NULL_VALUE: '#NULL!',
  LOADING: 'Loading...'
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
/**
 * A snapshot with no cell data at all, built from the sheet tabs in the page.
 * Used when the Sheets API refuses the document (an .xlsx opened in Sheets):
 * the walk still works from the formula bar and the name box; values do not.
 * tabs: [{name, hidden}] in tab order.
 */
Snapshot.fromTabs = function (tabs) {
  var snap = new Snapshot({
    sheets: (tabs || []).map(function (t, i) {
      return { properties: { sheetId: i, title: t.name, hidden: !!t.hidden, gridProperties: { rowCount: 1000, columnCount: 26 } } };
    })
  });
  snap.dataless = true;
  return snap;
};
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
