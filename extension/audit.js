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
    // JSON.stringify, not join('|'): a sheet name containing '|' would collide
    // with a neighbouring rect's key and drop a fetch (a false blank).
    var k = JSON.stringify([r.sheetName, r.r1, r.c1, r.r2, r.c2]);
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
  // Guard first: a stale walk (or a jump that landed on a sheet the snapshot
  // never heard of) must fail loudly, not fall through and read another sheet.
  if (!snap.getSheetByName(sheetName)) {
    return Promise.reject(new Error('Sheet "' + sheetName + '" is not in the snapshot — reopen Scopion.'));
  }
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
        // A fresh cache, not `cache`: the resolver's cache may have memoised
        // pre-fetch blanks, or a whole-grid snapshot taken before the dynamic
        // targets arrived, and either would print a stale value here.
        var rows = buildRows(snap, found.targets, new ValueCache(snap), settings);
        return {
          origin: {
            sheetName: sheetName,
            a1: a1,
            // ACE's list starts with the origin itself; the panel needs its value.
            value: snap.getSheetByName(sheetName).cell(box.r1, box.c1).d,
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

// namesCovering/rangeTotal/formatNumber/safeUrl below have a twin copy in the frozen Code.gs — change one, change both.
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
