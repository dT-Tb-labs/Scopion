/**
 * Code.gs — entry points and the RPC surface the sidebar calls.
 *
 * ACE for Google Sheets: a port of the Excel VBA add-in of the same name
 * (PwC, 2002-2012). Original trigger was Ctrl+Shift+A; Sheets only offers
 * Ctrl+Alt+Shift+<1-9> macro slots, so this binds slot 1 and README.md explains
 * the Karabiner mapping that restores the original chord.
 */

var SETTING_KEYS = ['showExternal', 'traverseHidden', 'showNamedRanges'];
var BLANK_LABEL = '---BLANK CELL---';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('ACE')
    .addItem('Audit active cell', 'aceOpen')
    .addToUi();
}

/** Macro entry point. Bound to Ctrl+Alt+Shift+1 by appsscript.json. */
function aceOpen() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('ACE — cell audit');
  SpreadsheetApp.getUi().showSidebar(html);
}

// ---------------------------------------------------------------------------
// Settings (the Excel original stored these as defined names inside the workbook;
// DocumentProperties is the Apps Script equivalent and does not touch the grid)
// ---------------------------------------------------------------------------

function aceGetSettings() {
  var props = PropertiesService.getDocumentProperties();
  var out = {};
  for (var i = 0; i < SETTING_KEYS.length; i++) {
    var raw = props.getProperty('ace.' + SETTING_KEYS[i]);
    out[SETTING_KEYS[i]] = raw === null ? true : raw === 'true';
  }
  return out;
}

function aceSetSetting(key, value) {
  if (SETTING_KEYS.indexOf(key) < 0) throw new Error('Unknown setting: ' + key);
  PropertiesService.getDocumentProperties()
    .setProperty('ace.' + key, value ? 'true' : 'false');
  return aceGetSettings();
}

// ---------------------------------------------------------------------------
// Audit
// ---------------------------------------------------------------------------

/**
 * @param {?Object} request {sheetName, a1, mode} — omit to audit the current
 *   selection. `mode` is 'precedents' or 'dependents'.
 */
function aceAudit(request) {
  var started = Date.now();
  var ss = SpreadsheetApp.getActive();
  var settings = aceGetSettings();
  var mode = (request && request.mode) === 'dependents' ? 'dependents' : 'precedents';

  var sheetName, a1;
  if (request && request.sheetName) {
    var validated = validateTarget(ss, request.sheetName, request.a1);
    sheetName = validated.sheetName;
    // Audit the top-left cell: getRange() rejects the unbounded "C:C" form, and
    // a range has no single formula to report anyway.
    var box = a1ToRect(validated.a1, sheetName);
    a1 = numToCol(box.c1) + box.r1;
  } else {
    var active = ss.getActiveRange();
    if (!active) throw new Error('Select a cell first.');
    sheetName = active.getSheet().getName();
    a1 = active.getCell(1, 1).getA1Notation();
  }

  var namedRanges = buildNamedRangeMap(ss);
  var originSheet = ss.getSheetByName(sheetName);
  var originCell = originSheet.getRange(a1);

  var found, unresolved = [], skipped = [], scanStats = null;
  if (mode === 'precedents') {
    found = findPrecedents(ss, sheetName, a1, namedRanges);
    unresolved = found.unresolved;
  } else {
    var dep = findDependents(ss, sheetName, a1, namedRanges);
    found = { formula: originCell.getFormula(), targets: dep.targets };
    skipped = dep.skipped;
    scanStats = { elapsedMs: dep.elapsedMs, scannedFormulas: dep.scannedFormulas };
  }

  var visible = {};
  var sheets = ss.getSheets();
  for (var s = 0; s < sheets.length; s++) visible[sheets[s].getName()] = !sheets[s].isSheetHidden();

  var targets = found.targets;
  if (!settings.traverseHidden) {
    targets = targets.filter(function (t) {
      return t.external || visible[t.rect.sheetName] !== false;
    });
  }
  if (!settings.showExternal) {
    targets = targets.filter(function (t) { return !t.external; });
  }

  var rows = materializeRows(ss, targets, visible);

  return {
    origin: {
      sheetName: sheetName,
      a1: a1,
      formula: found.formula || '',
      value: displayValue(originCell)
    },
    mode: mode,
    rows: rows,
    unresolved: unresolved,
    skippedSheets: skipped,
    scanStats: scanStats,
    namedRanges: settings.showNamedRanges ? namesCovering(namedRanges, sheetName, a1) : [],
    hasBlank: rows.some(function (r) { return r.isBlank; }),
    settings: settings,
    elapsedMs: Date.now() - started
  };
}

/** Reject anything from the browser that is not a real sheet and a legal A1 range. */
function validateTarget(ss, sheetName, a1) {
  var sheet = ss.getSheetByName(String(sheetName));
  if (!sheet) throw new Error('No such sheet: ' + sheetName);
  var r = a1ToRect(String(a1), sheet.getName());
  if (!r) throw new Error('Not a valid reference: ' + a1);
  // A formula may reference a cell outside the current grid (a deleted row, a
  // sheet that shrank). getRange() throws there, so it is caught here instead.
  if (r.r1 > sheet.getMaxRows() || r.c1 > sheet.getMaxColumns()) {
    throw new Error(a1 + ' is outside the grid of "' + sheet.getName() + '".');
  }
  return { sheetName: sheet.getName(), a1: String(a1).replace(/\$/g, '') };
}

/** Defined names whose range covers the audited cell (the original's txtNames box). */
function namesCovering(namedRanges, sheetName, a1) {
  var cell = a1ToRect(a1, sheetName);
  var out = [];
  for (var key in namedRanges) {
    var nr = namedRanges[key];
    if (rectsOverlap(nr.rect, cell)) out.push(nr.name);
  }
  return out;
}

/**
 * Read value, background and formula for every target, batching by sheet.
 * One getValues/getBackgrounds/getFormulas per sheet when the targets' bounding
 * box is small enough, otherwise per target — a formula referencing both A1 and
 * ZZ100000 would otherwise pull the whole sheet.
 */
function materializeRows(ss, targets, visible) {
  var bySheet = {};
  var rows = [];

  for (var i = 0; i < targets.length; i++) {
    var t = targets[i];
    if (t.external) {
      rows.push({
        index: i, external: true, sheetName: '(external)', flag: 'EX',
        address: t.a1, value: '', background: '', subFormula: t.raw,
        url: safeUrl(t.url), isBlank: false, jumpable: false
      });
      continue;
    }
    if (!bySheet[t.rect.sheetName]) bySheet[t.rect.sheetName] = [];
    bySheet[t.rect.sheetName].push({ index: i, target: t });
  }

  for (var sheetName in bySheet) {
    var sheet = ss.getSheetByName(sheetName);
    var items = bySheet[sheetName];
    if (!sheet) {
      for (var m = 0; m < items.length; m++) {
        rows.push(missingSheetRow(items[m], sheetName));
      }
      continue;
    }

    var maxRow = Math.max(sheet.getLastRow(), 1);
    var maxCol = Math.max(sheet.getLastColumn(), 1);
    var box = boundingBox(items, maxRow, maxCol);
    var useBatch = box && (box.r2 - box.r1 + 1) * (box.c2 - box.c1 + 1) <= MAX_MATERIALIZE_BBOX;

    var values = null, backgrounds = null, formulas = null;
    if (useBatch) {
      var range = sheet.getRange(box.r1, box.c1, box.r2 - box.r1 + 1, box.c2 - box.c1 + 1);
      values = range.getDisplayValues();
      backgrounds = range.getBackgrounds();
      formulas = range.getFormulas();
    }

    for (var k = 0; k < items.length; k++) {
      var item = items[k];
      var rct = item.target.rect;
      var isRange = rct.r1 !== rct.r2 || rct.c1 !== rct.c2;

      // Beyond the used range there is nothing to read. Clamping to the last
      // used cell would show a different cell's value under the right address.
      var withinGrid = rct.r1 <= sheet.getMaxRows() && rct.c1 <= sheet.getMaxColumns();
      if (rct.r1 > maxRow || rct.c1 > maxCol) {
        rows.push({
          index: item.index, external: false, sheetName: sheetName,
          flag: visible[sheetName] === false ? 'H' : '',
          address: rectToA1(rct), value: BLANK_LABEL,
          background: '', subFormula: item.target.raw || '',
          viaName: item.target.viaName || '', dynamic: !!item.target.dynamic,
          sourceFormula: item.target.sourceFormula || '',
          // A range entirely outside the used area is empty too, and an empty
          // referenced range is exactly the modelling error the alert is for.
          isBlank: true, jumpable: withinGrid
        });
        continue;
      }
      var r = rct.r1, c = rct.c1;

      var display, background, formula;
      if (useBatch) {
        display = values[r - box.r1][c - box.c1];
        background = backgrounds[r - box.r1][c - box.c1];
        formula = formulas[r - box.r1][c - box.c1];
      } else {
        var cellRange = sheet.getRange(r, c);
        display = cellRange.getDisplayValue();
        background = cellRange.getBackground();
        formula = cellRange.getFormula();
      }

      // "Blank" means structurally empty. A formula returning "" displays as
      // empty but is not a missing input, so it must not raise the alert.
      var isBlank = !isRange && display === '' && formula === '';

      rows.push({
        index: item.index,
        external: false,
        sheetName: sheetName,
        flag: visible[sheetName] === false ? 'H' : '',
        address: rectToA1(rct),
        value: isBlank ? BLANK_LABEL : (isRange ? '(' + rectToA1(rct) + ')' : display),
        background: background,
        subFormula: item.target.raw || '',
        viaName: item.target.viaName || '',
        dynamic: !!item.target.dynamic,
        sourceFormula: item.target.sourceFormula || '',
        isBlank: isBlank,
        jumpable: true
      });
    }
  }

  rows.sort(function (a, b) { return a.index - b.index; });
  return rows;
}

function missingSheetRow(item, sheetName) {
  return {
    index: item.index, external: false, sheetName: sheetName, flag: '?',
    address: rectToA1(item.target.rect), value: '(sheet not found)',
    background: '', subFormula: item.target.raw || '', isBlank: false, jumpable: false
  };
}

function boundingBox(items, maxRow, maxCol) {
  var r1 = null, c1 = null, r2 = null, c2 = null;
  for (var i = 0; i < items.length; i++) {
    var rct = items[i].target.rect;
    if (rct.r1 > maxRow || rct.c1 > maxCol) continue; // handled as blank, never read
    var r = rct.r1, c = rct.c1;
    r1 = r1 === null ? r : Math.min(r1, r);
    c1 = c1 === null ? c : Math.min(c1, c);
    r2 = r2 === null ? r : Math.max(r2, r);
    c2 = c2 === null ? c : Math.max(c2, c);
  }
  return r1 === null ? null : { r1: r1, c1: c1, r2: r2, c2: c2 };
}

/**
 * IMPORTRANGE's first argument is spreadsheet content, so it is untrusted input
 * to the sidebar. Only a plain https link is ever handed back for opening.
 */
function safeUrl(url) {
  var text = String(url || '').trim();
  if (/^https:\/\/[^\s"'<>]+$/.test(text)) return text;
  // IMPORTRANGE accepts a bare spreadsheet key as well as a full URL.
  if (/^[A-Za-z0-9_-]{20,}$/.test(text)) {
    return 'https://docs.google.com/spreadsheets/d/' + text;
  }
  return '';
}

function displayValue(range) {
  try { return range.getDisplayValue(); } catch (e) { return ''; }
}

// ---------------------------------------------------------------------------
// Navigation
// ---------------------------------------------------------------------------

/**
 * Move the selection to a target. The Excel original coloured the cell green and
 * restored the old colour on the way out; here the native selection outline does
 * that job, so the document is never written to for highlighting.
 *
 * A hidden sheet is the one case that needs a write: Sheets cannot activate a
 * hidden sheet. The sheet is unhidden and the fact is reported, so the sidebar
 * can tell the user their document changed.
 */
function aceJump(sheetName, a1) {
  var ss = SpreadsheetApp.getActive();
  var validated = validateTarget(ss, sheetName, a1);
  var sheet = ss.getSheetByName(validated.sheetName);
  var unhidden = false;

  if (sheet.isSheetHidden()) {
    sheet.showSheet();
    unhidden = true;
  }
  sheet.activate();
  // The address may be a whole column ("C:C") or row ("5:5"), which getRange()
  // rejects in that form. Select the range's top-left cell instead.
  var target = a1ToRect(validated.a1, validated.sheetName);
  var range = sheet.getRange(target.r1, target.c1);
  ss.setActiveRange(range);

  return { sheetName: validated.sheetName, a1: range.getA1Notation(), unhidden: unhidden };
}

/** Re-hide sheets this session unhid, when the user asks to undo that. */
function aceRehide(sheetNames) {
  var ss = SpreadsheetApp.getActive();
  var activeName = ss.getActiveSheet().getName();
  var rehidden = [];
  for (var i = 0; i < sheetNames.length; i++) {
    var sheet = ss.getSheetByName(sheetNames[i]);
    if (!sheet || sheet.getName() === activeName) continue;
    sheet.hideSheet();
    rehidden.push(sheet.getName());
  }
  return rehidden;
}
