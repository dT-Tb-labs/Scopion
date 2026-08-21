/**
 * Trace.gs — turns a cell into a list of what it reads (precedents) or what
 * reads it (dependents), plus the values/colours needed to render those rows.
 *
 * Everything here is read-only against the user's document. The only write the
 * add-on ever performs is unhiding a sheet the user explicitly asked to jump to.
 */

var MAX_MATERIALIZE_BBOX = 200000; // cells; above this, read targets one by one
var SCAN_DEADLINE_MS = 240000;     // stop scanning well inside the 6-minute cap

var MAX_GRID_CACHE_CELLS = 200000;  // above this, read single cells instead of the whole grid
var MAX_SCAN_CELLS = 2000000; // refuse to pull a grid large enough to exhaust the invocation

function sheetTooLarge(sheet) {
  return sheet.getLastRow() * sheet.getLastColumn() > MAX_SCAN_CELLS;
}

/** {sheetName, r1, c1, r2, c2} — r2/c2 are inclusive. */
function rect(sheetName, r1, c1, r2, c2) {
  return { sheetName: sheetName, r1: r1, c1: c1, r2: r2, c2: c2 };
}

function rectContains(r, sheetName, row, col) {
  return r.sheetName === sheetName && row >= r.r1 && row <= r.r2 && col >= r.c1 && col <= r.c2;
}

function rectToA1(r) {
  var full = r.r1 === 1 && r.r2 >= MAX_ROW;
  var wide = r.c1 === 1 && r.c2 >= MAX_COL;
  if (full && !wide) {
    return numToCol(r.c1) + ':' + numToCol(r.c2);
  }
  if (wide && !full) {
    return r.r1 + ':' + r.r2;
  }
  var a = numToCol(r.c1) + r.r1;
  var b = numToCol(r.c2) + r.r2;
  return a === b ? a : a + ':' + b;
}

/** Parse "A1", "A1:B2", "A:A", "1:1" into a rectangle on `sheetName`. */
function a1ToRect(a1, sheetName) {
  if (!a1) return null;
  var parts = String(a1).replace(/\$/g, '').split(':');
  var left = parseRefPart(parts[0]);
  if (!left) return null;
  var right = parts.length > 1 ? parseRefPart(parts[1]) : left;
  if (!right) return null;

  var r1 = left.row === null ? 1 : left.row;
  var r2 = right.row === null ? MAX_ROW : right.row;
  var c1 = left.col === null ? 1 : left.col;
  var c2 = right.col === null ? MAX_COL : right.col;

  return rect(sheetName, Math.min(r1, r2), Math.min(c1, c2), Math.max(r1, r2), Math.max(c1, c2));
}

/** name (upper-case) -> rectangle, built once per invocation. */
function buildNamedRangeMap(ss) {
  var map = {};
  var names = ss.getNamedRanges();
  for (var i = 0; i < names.length; i++) {
    var nr = names[i];
    var r;
    try {
      r = nr.getRange();
    } catch (e) {
      continue; // a name pointing at a deleted range
    }
    map[nr.getName().toUpperCase()] = {
      name: nr.getName(),
      rect: rect(r.getSheet().getName(), r.getRow(), r.getColumn(),
                 r.getLastRow(), r.getLastColumn())
    };
  }
  return map;
}

/**
 * Turn one extracted reference into a rectangle. Returns null for anything that
 * is not a local range (external refs and unknown names are handled by callers).
 */
function refToRect(ref, contextSheetName, namedRanges) {
  if (ref.kind === 'name') {
    var hit = namedRanges[ref.a1.toUpperCase()];
    return hit ? hit.rect : null;
  }
  if (ref.kind !== 'ref') return null;
  return a1ToRect(ref.a1, ref.sheet || contextSheetName);
}

// ---------------------------------------------------------------------------
// Dynamic targets: OFFSET / INDEX / INDIRECT
//
// SUM is deliberately NOT in this list. The Excel original scanned for it, but
// SUM does not produce a reference — its arguments are already ordinary
// precedents, so scanning for it double-counted every SUM range.
// ---------------------------------------------------------------------------

var DYNAMIC_FUNCS = ['OFFSET', 'INDEX', 'INDIRECT'];

/**
 * Evaluate an argument to a number, using only literals, simple arithmetic and
 * the *values* of single-cell references. Returns null when the argument needs
 * more than that (MATCH, QUERY, custom functions...), which the caller reports
 * as an explicitly unresolved row rather than guessing.
 */
function evalNumericArg(argText, contextSheetName, values, namedRanges) {
  var text = String(argText).trim();
  if (text === '') return null;

  if (/^-?[0-9]+(\.[0-9]+)?$/.test(text)) return Number(text);

  var call = evalSupportedCall(text, contextSheetName, values, namedRanges);
  if (call !== null) return call;

  // MATCH(...)+1 and COUNT(...)-1 are everyday model idioms, so substitute any
  // supported call with its value and let the arithmetic below finish the job.
  // Substituted back-to-front, or an earlier replacement would shift the
  // positions of the ones after it.
  var inner = findCalls(text, SUPPORTED_CALLS);
  for (var c = inner.length - 1; c >= 0; c--) {
    if (inner[c].text === text) continue; // the bare-call case, handled above
    var v = evalSupportedCall(inner[c].text, contextSheetName, values, namedRanges);
    if (v === null) return null;
    text = text.substring(0, inner[c].index) + v +
           text.substring(inner[c].index + inner[c].text.length);
  }

  var refs = extractRefs(text);
  var substituted = text;
  for (var i = 0; i < refs.length; i++) {
    var r = refs[i];
    var v = singleCellValue(r, contextSheetName, values, namedRanges);
    if (typeof v !== 'number') return null;
    // Bounded replace: a plain split on "A1" would also corrupt "A10".
    substituted = replaceRefToken(substituted, r.raw, String(v));
  }

  // Only a bare arithmetic expression survives; anything else is unresolvable.
  if (!/^[-+*/() 0-9.]+$/.test(substituted)) return null;
  var value = arithmetic(substituted);
  return (value === null || isNaN(value)) ? null : value;
}

/**
 * Resolve an argument that must denote a range.
 *
 * Only an argument that IS a reference, or IS a nested dynamic call, is
 * accepted. Picking the first reference out of an arbitrary expression makes
 * OFFSET(OFFSET(A1,1,0),1,0) silently report A2 instead of A3, and a
 * confidently wrong precedent is worse than a reported gap.
 */
function resolveRefArg(argText, contextSheetName, values, namedRanges, depth) {
  var text = String(argText === undefined ? '' : argText).trim();
  if (text === '') return null;
  if ((depth || 0) > 8) return null;

  var nested = findCalls(text, DYNAMIC_FUNCS);
  if (nested.length && nested[0].text === text) {
    var inner = resolveDynamicTargets(text, contextSheetName, values, namedRanges, (depth || 0) + 1);
    return inner.length ? inner[0].rect : null;
  }

  var refs = extractRefs(text);
  if (refs.length !== 1) return null;
  var bare = text.replace(/\$/g, '');
  if (refs[0].raw !== text && refs[0].raw !== bare) return null;
  return refToRect(refs[0], contextSheetName, namedRanges);
}

/**
 * The value of a reference that denotes exactly one cell, whether it is written
 * as an address or as a named range. Named ranges are the normal way a real
 * model writes an OFFSET argument (OFFSET(Base, Shift, 0)), so refusing them
 * here reported every such formula as unresolvable.
 */
function singleCellValue(ref, contextSheetName, values, namedRanges) {
  if (ref.kind === 'name') {
    var nr = (namedRanges || {})[ref.a1.toUpperCase()];
    if (!nr) return null;
    var r = nr.rect;
    if (r.r1 !== r.r2 || r.c1 !== r.c2) return null; // a multi-cell name is not a scalar
    return values.get(r.sheetName, numToCol(r.c1) + r.r1);
  }
  if (ref.kind === 'ref' && isFullCellRef(ref.a1)) {
    return values.get(ref.sheet || contextSheetName, ref.a1);
  }
  return null;
}

/**
 * Evaluate the small set of functions that real models put INSIDE a dynamic
 * reference. Without these, INDEX(range, MATCH(...)) and
 * OFFSET(base, 0, 0, COUNT(range), 1) — the two idioms most financial models
 * are built from — could only ever be reported as unresolved.
 *
 * Deliberately narrow: an unsupported shape returns null and the caller says
 * so, which is safe. Guessing would put a wrong cell under a right-looking
 * address.
 */
var SUPPORTED_CALLS = ['MATCH', 'COUNT', 'COUNTA', 'ROWS', 'COLUMNS'];

function evalSupportedCall(argText, contextSheetName, values, namedRanges) {
  var text = String(argText === undefined ? '' : argText).trim();
  var calls = findCalls(text, SUPPORTED_CALLS);
  if (!calls.length || calls[0].text !== text) return null; // must BE the call
  var call = calls[0];
  var rangeOf = function (i) {
    return resolveRefArg(call.args[i], contextSheetName, values, namedRanges);
  };

  if (call.name === 'ROWS' || call.name === 'COLUMNS') {
    var r = rangeOf(0);
    if (!r) return null;
    return call.name === 'ROWS' ? (r.r2 - r.r1 + 1) : (r.c2 - r.c1 + 1);
  }

  if (call.name === 'COUNT' || call.name === 'COUNTA') {
    var cr = rangeOf(0);
    if (!cr) return null;
    var cells = values.getRect(cr);
    if (!cells.length) return null;
    var n = 0;
    for (var i = 0; i < cells.length; i++) {
      var v = cells[i];
      if (call.name === 'COUNT') { if (typeof v === 'number') n++; }
      else if (v !== '' && v !== null && v !== undefined) n++;
    }
    return n;
  }

  // MATCH(key, range, 0). Only exact match: types 1 and -1 assume the range is
  // sorted, and a wrong assumption there silently returns the wrong row.
  var type = call.args.length > 2 ? String(call.args[2]).trim() : '1';
  if (type !== '0') return null;
  var range = rangeOf(1);
  if (!range) return null;
  var key = evalScalar(call.args[0], contextSheetName, values, namedRanges);
  if (key === null) return null;
  var list = values.getRect(range);
  if (!list.length) return null;
  for (var j = 0; j < list.length; j++) {
    var cell = list[j];
    if (cell === key) return j + 1;
    if (typeof cell === 'string' && typeof key === 'string' &&
        cell.toUpperCase() === key.toUpperCase()) return j + 1;
  }
  return null; // #N/A — not a number we can offset by
}

/** A literal, or the value of one cell. Used for a MATCH key. */
function evalScalar(argText, contextSheetName, values, namedRanges) {
  var text = String(argText === undefined ? '' : argText).trim();
  if (text === '') return null;
  var str = /^"((?:[^"]|"")*)"$/.exec(text);
  if (str) return str[1].split('""').join('"');
  if (/^-?[0-9]+(\.[0-9]+)?$/.test(text)) return Number(text);
  var refs = extractRefs(text);
  if (refs.length !== 1) return null;
  var bare = text.replace(/\$/g, '');
  if (refs[0].raw !== text && refs[0].raw !== bare) return null;
  return singleCellValue(refs[0], contextSheetName, values, namedRanges);
}

/** OFFSET(A1,,2) is legal: an omitted argument is zero, not an unresolvable one. */
function blankArgIsZero(argText, contextSheetName, values, namedRanges) {
  if (argText === undefined || String(argText).trim() === '') return 0;
  return evalNumericArg(argText, contextSheetName, values, namedRanges);
}

/**
 * Evaluate "1+2*(3-1)" without eval(). Recursive descent, because a
 * shunting-yard pass has to special-case unary signs and gets "2*-3" wrong.
 *   expr   := term (("+"|"-") term)*
 *   term   := factor (("*"|"/") factor)*
 *   factor := ("+"|"-") factor | number | "(" expr ")"
 */
function arithmetic(expr) {
  var tokens = String(expr).match(/[0-9]*\.?[0-9]+|[-+*\/()]/g);
  if (!tokens) return null;
  var pos = 0;
  var failed = false;

  function peek() { return pos < tokens.length ? tokens[pos] : null; }

  function factor() {
    var t = peek();
    if (t === '+' || t === '-') {
      pos++;
      var operand = factor();
      if (failed) return 0;
      return t === '-' ? -operand : operand;
    }
    if (t === '(') {
      pos++;
      var inner = expression();
      if (peek() !== ')') { failed = true; return 0; }
      pos++;
      return inner;
    }
    if (t !== null && /^[0-9.]/.test(t)) { pos++; return Number(t); }
    failed = true;
    return 0;
  }

  function term() {
    var value = factor();
    while (!failed && (peek() === '*' || peek() === '/')) {
      var op = tokens[pos++];
      var rhs = factor();
      if (failed) return 0;
      if (op === '/') {
        if (rhs === 0) { failed = true; return 0; }
        value = value / rhs;
      } else {
        value = value * rhs;
      }
    }
    return value;
  }

  function expression() {
    var value = term();
    while (!failed && (peek() === '+' || peek() === '-')) {
      var op = tokens[pos++];
      var rhs = term();
      if (failed) return 0;
      value = op === '+' ? value + rhs : value - rhs;
    }
    return value;
  }

  var result = expression();
  if (failed || pos !== tokens.length || isNaN(result)) return null;
  return result;
}

/** Replace a reference token only where it stands alone, so A1 does not match A10. */
function replaceRefToken(text, token, value) {
  var escaped = token.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  var re = new RegExp('(^|[^A-Za-z0-9_$!.])' + escaped + '(?![A-Za-z0-9_$])', 'g');
  return text.replace(re, function (m, lead) { return lead + value; });
}

/**
 * One getValues() per sheet, reused for every dynamic-argument lookup.
 * Without this, resolving OFFSET arguments during a whole-spreadsheet dependents
 * scan would issue one API round trip per formula, which alone exhausts the
 * execution limit on any real model. Sheets too large to hold fall back to
 * memoised single-cell reads.
 */
function ValueCache(ss) {
  this.ss = ss;
  this.sheets = {};
  this.singles = {};
}

/** Values of a rectangle, from the same cached grid. [] when unavailable. */
ValueCache.prototype.getRect = function (r) {
  if (!r) return [];
  var grid = this.grid(r.sheetName);
  if (!grid || !grid.length) return [];
  var out = [];
  var lastRow = Math.min(r.r2, grid.length);
  var lastCol = Math.min(r.c2, grid[0].length);
  for (var row = r.r1; row <= lastRow; row++) {
    for (var col = r.c1; col <= lastCol; col++) out.push(grid[row - 1][col - 1]);
  }
  return out;
};

/**
 * Total of the numeric cells in a rectangle. The Excel original showed this
 * for SUM ranges only; a modeller wants it for every multi-cell reference —
 * "is this block the size I think it is" is the question a range row is
 * there to answer.
 */
ValueCache.prototype.sumRect = function (r) {
  var cells = this.getRect(r);
  var sum = 0, count = 0;
  for (var i = 0; i < cells.length; i++) {
    if (typeof cells[i] === 'number') { sum += cells[i]; count++; }
  }
  return { sum: sum, count: count };
};

/** The cached used-range grid for a sheet, or null when it cannot be held. */
ValueCache.prototype.grid = function (sheetName) {
  var grid = this.sheets[sheetName];
  if (grid === undefined) {
    var sheet = this.ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() === 0) grid = null;
    else if (sheet.getLastRow() * sheet.getLastColumn() > MAX_GRID_CACHE_CELLS) grid = false;
    else grid = sheet.getDataRange().getValues();
    this.sheets[sheetName] = grid;
  }
  return grid === false ? null : grid;
};

ValueCache.prototype.get = function (sheetName, a1) {
  var r = a1ToRect(a1, sheetName);
  if (!r) return null;

  var grid = this.sheets[sheetName];
  if (grid === undefined) {
    var sheet = this.ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() === 0) {
      grid = null;
    } else if (sheet.getLastRow() * sheet.getLastColumn() > MAX_GRID_CACHE_CELLS) {
      grid = false;
    } else {
      grid = sheet.getDataRange().getValues();
    }
    this.sheets[sheetName] = grid;
  }

  if (grid === false) {
    var key = sheetName + '!' + a1;
    if (!(key in this.singles)) {
      var sh = this.ss.getSheetByName(sheetName);
      this.singles[key] = sh ? sh.getRange(r.r1, r.c1).getValue() : null;
    }
    return this.singles[key];
  }
  if (!grid) return null;
  if (grid.length === 0 || r.r1 > grid.length || r.c1 > grid[0].length) return null;
  return grid[r.r1 - 1][r.c1 - 1];
};

/** Evaluate an argument to a reference string, for INDIRECT. */
function evalTextArg(argText, contextSheetName, values, namedRanges) {
  var text = String(argText).trim();
  var pieces = splitTopLevel(text, '&');
  var result = '';

  for (var i = 0; i < pieces.length; i++) {
    var p = pieces[i].trim();
    var strMatch = /^"((?:[^"]|"")*)"$/.exec(p);
    if (strMatch) { result += strMatch[1].split('""').join('"'); continue; }
    if (isFullCellRef(p.replace(/\$/g, ''))) {
      var v = values.get(contextSheetName, p.replace(/\$/g, ''));
      if (v === null) return null;
      result += String(v);
      continue;
    }
    var qualified = /^(?:'((?:[^']|'')*)'|([A-Za-z0-9_. ]+))!(\$?[A-Za-z]{1,3}\$?[0-9]+)$/.exec(p);
    if (qualified) {
      var sName = qualified[1] !== undefined ? qualified[1].split("''").join("'") : qualified[2];
      var qv = values.get(sName, qualified[3].replace(/\$/g, ''));
      if (qv === null) return null;
      result += String(qv);
      continue;
    }
    var named = (namedRanges || {})[p.toUpperCase()];
    if (named) {
      var nv = singleCellValue({ kind: 'name', a1: p }, contextSheetName, values, namedRanges);
      if (nv !== null) { result += String(nv); continue; }
    }
    return null; // needs a function we do not implement
  }
  return result;
}

/** Split on a single-character operator, ignoring anything inside quotes or brackets. */
function splitTopLevel(text, sep) {
  var out = [], cur = '', depth = 0, inStr = false, inQuote = false;
  for (var i = 0; i < text.length; i++) {
    var c = text.charAt(i);
    if (inStr) { cur += c; if (c === '"') { if (text.charAt(i + 1) === '"') { cur += '"'; i++; } else inStr = false; } continue; }
    if (inQuote) { cur += c; if (c === "'") { if (text.charAt(i + 1) === "'") { cur += "'"; i++; } else inQuote = false; } continue; }
    if (c === '"') { inStr = true; cur += c; continue; }
    if (c === "'") { inQuote = true; cur += c; continue; }
    if (c === '(') depth++;
    if (c === ')') depth--;
    if (c === sep && depth === 0) { out.push(cur); cur = ''; continue; }
    cur += c;
  }
  out.push(cur);
  return out;
}

/**
 * Resolve every OFFSET/INDEX/INDIRECT call in `formula` to a rectangle.
 * Unresolvable calls come back with `rect: null` and a reason, so the UI can
 * show them as known-but-unresolved instead of silently omitting them.
 */
function resolveDynamicTargets(formula, contextSheetName, values, namedRanges, depth) {
  var calls = findCalls(formula, DYNAMIC_FUNCS);
  var out = [];

  for (var i = 0; i < calls.length; i++) {
    var call = calls[i];
    var resolved = null;
    var reason = '';
    var callBase = null;

    try {
      if (call.name === 'OFFSET') {
        var base = resolveRefArg(call.args[0], contextSheetName, values, namedRanges, depth);
        var dr = blankArgIsZero(call.args[1], contextSheetName, values, namedRanges);
        var dc = blankArgIsZero(call.args[2], contextSheetName, values, namedRanges);
        if (!base) reason = 'base reference not static';
        else if (dr === null || dc === null) reason = 'row/column offset needs an unsupported function';
        else {
          var h = call.args.length > 3 && call.args[3] !== ''
            ? evalNumericArg(call.args[3], contextSheetName, values, namedRanges) : (base.r2 - base.r1 + 1);
          var w = call.args.length > 4 && call.args[4] !== ''
            ? evalNumericArg(call.args[4], contextSheetName, values, namedRanges) : (base.c2 - base.c1 + 1);
          if (h === null || w === null) reason = 'height/width needs an unsupported function';
          else {
            var r1 = base.r1 + dr, c1 = base.c1 + dc;
            if (r1 < 1 || c1 < 1) reason = 'offset lands outside the sheet';
            else resolved = rect(base.sheetName, r1, c1, r1 + h - 1, c1 + w - 1);
          }
        }
      } else if (call.name === 'INDEX') {
        var range = resolveRefArg(call.args[0], contextSheetName, values, namedRanges, depth);
        callBase = range;
        var rn = call.args.length > 1
          ? blankArgIsZero(call.args[1], contextSheetName, values, namedRanges) : 1;
        var cn = call.args.length > 2 && String(call.args[2]).trim() !== ''
          ? evalNumericArg(call.args[2], contextSheetName, values, namedRanges) : null;
        if (!range) reason = 'first argument is not a static range';
        else if (rn === null) reason = 'row number needs an unsupported function';
        else if (cn === null && range.r1 === range.r2 && range.c1 !== range.c2) {
          // INDEX(A1:E1, 3) indexes along the only row, so it selects a column.
          var oneRowCol = range.c1 + (rn > 0 ? rn - 1 : 0);
          resolved = rn === 0
            ? rect(range.sheetName, range.r1, range.c1, range.r1, range.c2)
            : rect(range.sheetName, range.r1, oneRowCol, range.r1, oneRowCol);
        } else {
          // row 0 (or column 0) means "the whole column (or row) of the range".
          var tr1 = rn === 0 ? range.r1 : range.r1 + rn - 1;
          var tr2 = rn === 0 ? range.r2 : tr1;
          // An omitted or zero column means the whole row of the range; on a
          // single-column range that collapses back to the one cell.
          var tc1, tc2;
          if (cn === null || cn === 0) {
            tc1 = range.c1;
            tc2 = range.c2;
          } else {
            tc1 = range.c1 + cn - 1;
            tc2 = tc1;
          }
          resolved = rect(range.sheetName, tr1, tc1, tr2, tc2);
        }
      } else if (call.name === 'INDIRECT') {
        if (call.args.length > 1 && /FALSE/i.test(call.args[1])) {
          reason = 'R1C1 form is not supported';
        } else {
          var text = evalTextArg(call.args[0] || '', contextSheetName, values, namedRanges);
          if (text === null) reason = 'reference text needs an unsupported function';
          else {
            var bang = text.lastIndexOf('!');
            var sheetName = contextSheetName, addr = text;
            if (bang >= 0) {
              sheetName = text.substring(0, bang).replace(/^'|'$/g, '').split("''").join("'");
              addr = text.substring(bang + 1);
            }
            var named = namedRanges[text.toUpperCase()];
            resolved = named ? named.rect : a1ToRect(addr, sheetName);
            if (!resolved) reason = 'result "' + text + '" is not a valid reference';
          }
        }
      }
    } catch (e) {
      reason = 'resolution failed: ' + e.message;
    }

    // Sheets returns #REF! when an index runs past its range; reporting the
    // cell just outside would be a fabricated precedent.
    if (resolved && !reason && callBase && !rectWithin(resolved, callBase)) {
      resolved = null;
      reason = 'index falls outside the range (Sheets returns #REF!)';
    }
    out.push({ rect: resolved, raw: call.text, reason: reason, dynamic: true, pos: call.index });
  }
  return out;
}

// ---------------------------------------------------------------------------
// Precedents / dependents
// ---------------------------------------------------------------------------

function findPrecedents(ss, sheetName, a1, namedRanges, values) {
  values = values || new ValueCache(ss);
  var sheet = ss.getSheetByName(sheetName);
  var formula = sheet.getRange(a1).getFormula();
  var targets = [];
  var unresolved = [];

  if (formula) {
    var refs = extractRefs(formula);
    for (var i = 0; i < refs.length; i++) {
      var ref = refs[i];
      if (ref.kind === 'external') {
        targets.push({ external: true, url: ref.url, a1: ref.a1, raw: ref.raw, pos: ref.pos || 0 });
        continue;
      }
      if (ref.kind === 'name' && !namedRanges[ref.a1.toUpperCase()]) continue; // a function we do not know, not a name
      var r = refToRect(ref, sheetName, namedRanges);
      if (r) targets.push({ rect: r, raw: ref.raw, pos: ref.pos || 0,
                            viaName: ref.kind === 'name' ? ref.a1 : '' });
    }

    var dyn = resolveDynamicTargets(formula, sheetName, values, namedRanges);
    for (var d = 0; d < dyn.length; d++) {
      if (dyn[d].rect) {
        targets.push({ rect: dyn[d].rect, raw: dyn[d].raw, dynamic: true, pos: dyn[d].pos || 0 });
      } else {
        unresolved.push({ raw: dyn[d].raw, reason: dyn[d].reason });
      }
    }

    // Reading order. A modeller checks a formula left to right, so the walk
    // must follow the formula, not the order the resolver happened to produce
    // (static refs first, dynamic targets appended). A dynamic call sorts at
    // the position of the call itself, so OFFSET(...)'s resolved target sits
    // where OFFSET is written, ahead of its own arguments.
    targets = targets.map(function (t, i) { return { t: t, i: i }; })
      .sort(function (a, b) {
        var d = (a.t.pos || 0) - (b.t.pos || 0);
        return d !== 0 ? d : a.i - b.i;
      })
      .map(function (w) { return w.t; });
  }

  return { formula: formula, targets: targets, unresolved: unresolved };
}

/**
 * Scan every sheet's formulas and return the cells whose references cover the
 * target. No pre-built index: one getFormulas() call per sheet, parsed in
 * memory. The elapsed time is reported to the UI so the cost is visible rather
 * than assumed — add an index only if these numbers justify one.
 */
function findDependents(ss, targetSheetName, a1, namedRanges, values) {
  values = values || new ValueCache(ss);
  var target = a1ToRect(a1, targetSheetName);
  var sheets = ss.getSheets();
  var targets = [];
  var skipped = [];
  var started = Date.now();
  var scannedFormulas = 0;

  for (var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];
    var sheetName = sheet.getName();

    if (Date.now() - started > SCAN_DEADLINE_MS) {
      skipped.push(sheetName + ' (time limit)');
      continue;
    }
    if (sheet.getLastRow() === 0) continue;
    if (sheetTooLarge(sheet)) {
      skipped.push(sheetName + ' (' + (sheet.getLastRow() * sheet.getLastColumn()) + ' cells)');
      continue;
    }

    var formulas = sheet.getDataRange().getFormulas();
    for (var r = 0; r < formulas.length; r++) {
      if ((r & 255) === 0 && Date.now() - started > SCAN_DEADLINE_MS) {
        skipped.push(sheetName + ' (time limit, stopped at row ' + (r + 1) + ')');
        break;
      }
      var row = formulas[r];
      for (var c = 0; c < row.length; c++) {
        var f = row[c];
        if (!f) continue;
        scannedFormulas++;

        var refs = extractRefs(f);
        var hit = null;
        for (var i = 0; i < refs.length && !hit; i++) {
          if (refs[i].kind === 'external') continue;
          if (refs[i].kind === 'name' && !namedRanges[refs[i].a1.toUpperCase()]) continue;
          var rr = refToRect(refs[i], sheetName, namedRanges);
          if (rr && rectsOverlap(rr, target)) hit = refs[i].raw;
        }

        // A dependency can also be created dynamically; resolve those only for
        // formulas that mention one of the dynamic functions, to keep the scan cheap.
        if (!hit && /OFFSET\s*\(|INDEX\s*\(|INDIRECT\s*\(/i.test(f)) {
          var dyn = resolveDynamicTargets(f, sheetName, values, namedRanges);
          for (var d = 0; d < dyn.length && !hit; d++) {
            if (dyn[d].rect && rectsOverlap(dyn[d].rect, target)) hit = dyn[d].raw;
          }
        }

        if (hit) {
          targets.push({
            rect: rect(sheetName, r + 1, c + 1, r + 1, c + 1),
            raw: hit,
            sourceFormula: f
          });
        }
      }
    }
  }

  return {
    targets: targets,
    skipped: skipped,
    elapsedMs: Date.now() - started,
    scannedFormulas: scannedFormulas
  };
}

/** True when `inner` sits entirely inside `outer`. */
function rectWithin(inner, outer) {
  return inner.sheetName === outer.sheetName &&
    inner.r1 >= outer.r1 && inner.r2 <= outer.r2 &&
    inner.c1 >= outer.c1 && inner.c2 <= outer.c2;
}

function rectsOverlap(a, b) {
  return a.sheetName === b.sheetName &&
    a.r1 <= b.r2 && b.r1 <= a.r2 &&
    a.c1 <= b.c2 && b.c1 <= a.c2;
}
