/**
 * @OnlyCurrentDoc
 *
 * Code.gs — entry points and the RPC surface the Scopion sidebar calls.
 *
 * Scopion: a formula auditor for Google Sheets, descended from the Excel VBA
 * add-in "ACE" (PwC, 2002-2012). The Excel original hung off a hotkey;
 * Marketplace add-ons cannot register shortcuts, so Scopion is a persistent
 * sidebar that FOLLOWS the grid selection (scopionObserve polling) and gives
 * arrow-key walking of the reference list once the sidebar has focus.
 */

var SETTING_KEYS = ['showExternal', 'traverseHidden', 'showNamedRanges', 'followMode'];
var BLANK_LABEL = '---BLANK CELL---';
var WALK_HIGHLIGHT = '#ccff90'; // the Excel original painted the walked cell ColorIndex 4

function onOpen() {
  // AuthMode.NONE-safe: menu construction only, no spreadsheet reads.
  SpreadsheetApp.getUi()
    .createMenu('Scopion')
    .addItem('Open Scopion', 'showScopionSidebar')
    .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function showScopionSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Scopion');
  SpreadsheetApp.getUi().showSidebar(html);
}

/** Macro alias for the bound-script build (Ctrl+Alt+Shift+1). */
function scopionOpen() {
  showScopionSidebar();
}

// ---------------------------------------------------------------------------
// Settings (DocumentProperties; the Excel original used workbook defined names)
// ---------------------------------------------------------------------------

function scopionGetSettings() {
  var out = {};
  var props = null;
  try {
    props = PropertiesService.getDocumentProperties();
  } catch (e) {
    // view-only access: no document properties, fall through to defaults
  }
  for (var i = 0; i < SETTING_KEYS.length; i++) {
    var raw = props ? props.getProperty('scopion.' + SETTING_KEYS[i]) : null;
    out[SETTING_KEYS[i]] = raw === null ? true : raw === 'true';
  }
  return out;
}

function scopionSetSetting(key, value) {
  if (SETTING_KEYS.indexOf(key) < 0) throw new Error('Unknown setting: ' + key);
  // Best-effort: a viewer without edit rights cannot persist settings; the
  // sidebar keeps its in-memory state, so failing loudly here would only
  // break the toggle for read-only users.
  try {
    PropertiesService.getDocumentProperties()
      .setProperty('scopion.' + key, value ? 'true' : 'false');
  } catch (e) {}
  return scopionGetSettings();
}

// ---------------------------------------------------------------------------
// Selection identity: sheetId:row:col. Immutable against sheet renames, which
// a name-keyed protocol would silently mis-resolve mid-session.
// ---------------------------------------------------------------------------

function sheetById(ss, sheetId) {
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === sheetId) return sheets[i];
  }
  return null;
}

function selectionInfo(ss) {
  var range = ss.getActiveRange();
  if (!range) return null;
  var sheet = range.getSheet();
  var cell = range.getCell(1, 1); // multi-cell selections audit the top-left
  return {
    key: sheet.getSheetId() + ':' + cell.getRow() + ':' + cell.getColumn(),
    sheetId: sheet.getSheetId(),
    sheetName: sheet.getName(),
    row: cell.getRow(),
    column: cell.getColumn(),
    a1: cell.getA1Notation(),
    isSingleCell: range.getNumRows() === 1 && range.getNumColumns() === 1
  };
}

/**
 * The follow-mode poll. Cheap when nothing changed; when the selection HAS
 * changed it captures the coordinates and returns the full audit in the same
 * round trip — a separate peek-then-audit pair would race against the user
 * moving again in between.
 */
function scopionObserve(request) {
  var ss = SpreadsheetApp.getActive();
  // A sidebar that was closed mid-walk leaves a painted cell behind. The first
  // poll of a new session clears it.
  if (request && request.firstRun) scopionClearHighlight();
  var sel = selectionInfo(ss);
  if (!sel) return { changed: false, selection: null };
  if (request && request.knownKey === sel.key) {
    return { changed: false, selection: { key: sel.key } };
  }
  var audit = scopionAuditCore({
    sheetName: sel.sheetName,
    a1: sel.a1,
    mode: request && request.mode
  });
  return { changed: true, selection: sel, audit: audit };
}

/**
 * The walk highlight is owned by the DOCUMENT, not the sidebar.
 *
 * A sidebar reload (or simply closing and reopening it) used to orphan the
 * painted cell forever, because only the client remembered what to restore.
 * Keeping the record in DocumentProperties means any later run — including the
 * next boot — can clean it up.
 */
function readHighlightRecord() {
  try {
    var raw = PropertiesService.getDocumentProperties().getProperty('scopion.hl');
    return raw ? JSON.parse(raw) : null;
  } catch (e) { return null; }
}

function writeHighlightRecord(record) {
  try {
    var props = PropertiesService.getDocumentProperties();
    if (record) props.setProperty('scopion.hl', JSON.stringify(record));
    else props.deleteProperty('scopion.hl');
  } catch (e) {}
}

/**
 * Restore one previously highlighted cell to its recorded background.
 * Failures (protected range, deleted sheet) are swallowed: losing a restore
 * must never break navigation.
 */
function restoreHighlight(ss, restore) {
  if (!restore || !restore.target || restore.target.sheetId == null) return;
  try {
    var sheet = sheetById(ss, restore.target.sheetId);
    if (!sheet) return;
    var cell = sheet.getRange(restore.target.row, restore.target.column);
    if (restore.background && restore.background !== '#ffffff') {
      cell.setBackground(restore.background);
    } else {
      cell.setBackground(null);
    }
  } catch (e) {}
}

/** Standalone restore, for ending a walk without another jump. */
function scopionClearHighlight() {
  var record = readHighlightRecord();
  if (record) {
    restoreHighlight(SpreadsheetApp.getActive(), record);
    writeHighlightRecord(null);
  }
  return true;
}

/**
 * Navigation. action "jump" moves the grid selection only (arrow-key walking);
 * "navigateAndAudit" moves it AND audits the target as the new origin in one
 * round trip (Enter drill-in / Back).
 */
function scopionNavigate(request) {
  if (!request || (request.action !== 'jump' && request.action !== 'navigateAndAudit')) {
    throw new Error('Unknown navigation action.');
  }
  var ss = SpreadsheetApp.getActive();
  var sheet = sheetById(ss, request.target.sheetId);
  if (!sheet) throw new Error('Sheet no longer exists.');
  var row = Math.floor(request.target.row);
  var column = Math.floor(request.target.column);
  if (!(row >= 1) || !(column >= 1) ||
      row > sheet.getMaxRows() || column > sheet.getMaxColumns()) {
    throw new Error('Target is outside the grid.');
  }

  var unhidden = false;
  if (sheet.isSheetHidden()) {
    sheet.showSheet();
    unhidden = true;
  }
  sheet.activate();
  var range = sheet.getRange(row, column);
  ss.setActiveRange(range);

  // Walk highlight: paint the landed-on cell so the eye finds it instantly
  // (the selection outline alone is easy to lose on a dense model). The
  // previous walked cell is restored in the same round trip; the client keeps
  // the single {target, background} record and sends it back on the next hop.
  var highlight = null;
  if (request.highlight && request.highlight.apply) {
    restoreHighlight(ss, readHighlightRecord()); // whatever is painted now
    try {
      var prevBackground = range.getBackground();
      range.setBackground(WALK_HIGHLIGHT);
      writeHighlightRecord({
        target: { sheetId: sheet.getSheetId(), row: row, column: column },
        background: prevBackground
      });
      highlight = { applied: true };
    } catch (e) {
      writeHighlightRecord(null);
      highlight = { applied: false }; // protected range
    }
  } else if (request.highlight) {
    restoreHighlight(ss, readHighlightRecord());
    writeHighlightRecord(null);
    highlight = { applied: false };
  }

  var selection = {
    key: sheet.getSheetId() + ':' + row + ':' + column,
    sheetId: sheet.getSheetId(),
    sheetName: sheet.getName(),
    row: row,
    column: column,
    a1: range.getA1Notation(),
    sheetWasUnhidden: unhidden,
    highlight: highlight
  };

  if (request.action === 'navigateAndAudit') {
    return {
      selection: selection,
      audit: scopionAuditCore({
        sheetName: sheet.getName(),
        a1: range.getA1Notation(),
        mode: request.mode
      })
    };
  }
  return { selection: selection };
}

/** Mode switches / refresh on the current origin (no navigation). */
function scopionAudit(request) {
  // Prefer the immutable target: a sheetName captured before a rename would
  // resolve to "No such sheet" even though the sheet still exists.
  if (request && request.target && request.target.sheetId != null) {
    var sheet = sheetById(SpreadsheetApp.getActive(), request.target.sheetId);
    if (!sheet) throw new Error('Sheet no longer exists.');
    return scopionAuditCore({
      sheetName: sheet.getName(),
      a1: sheet.getRange(request.target.row, request.target.column).getA1Notation(),
      mode: request.mode
    });
  }
  return scopionAuditCore(request);
}

// ---------------------------------------------------------------------------
// Audit
// ---------------------------------------------------------------------------

/**
 * @param {?Object} request {sheetName, a1, mode} — omit to audit the current
 *   selection. `mode` is 'precedents' or 'dependents'.
 */
function scopionAuditCore(request) {
  var started = Date.now();
  var ss = SpreadsheetApp.getActive();
  var settings = scopionGetSettings();
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

  // Navigation identity for each row (sheetId:row:col of the row's top-left).
  var idBySheet = {};
  var allSheets = ss.getSheets();
  for (var sIdx = 0; sIdx < allSheets.length; sIdx++) {
    idBySheet[allSheets[sIdx].getName()] = allSheets[sIdx].getSheetId();
  }
  for (var rIdx = 0; rIdx < rows.length; rIdx++) {
    var rw = rows[rIdx];
    if (rw.external || !rw.jumpable) continue;
    var box = a1ToRect(rw.address, rw.sheetName);
    if (box && idBySheet[rw.sheetName] !== undefined) {
      rw.target = { sheetId: idBySheet[rw.sheetName], row: box.r1, column: box.c1 };
    } else {
      rw.jumpable = false;
    }
  }

  var originSheetId = null;
  var origSheets = ss.getSheets();
  for (var oIdx = 0; oIdx < origSheets.length; oIdx++) {
    if (origSheets[oIdx].getName() === sheetName) { originSheetId = origSheets[oIdx].getSheetId(); break; }
  }
  var originBox = a1ToRect(a1, sheetName);
  return {
    origin: {
      sheetName: sheetName,
      sheetId: originSheetId,
      a1: a1,
      key: originSheetId + ':' + originBox.r1 + ':' + originBox.c1,
      target: { sheetId: originSheetId, row: originBox.r1, column: originBox.c1 },
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

/** Re-hide sheets this session unhid, when the user asks to undo that. */
function scopionRehide(sheetNames, originTarget) {
  var ss = SpreadsheetApp.getActive();
  // After walking into a hidden sheet, that sheet IS the active sheet and
  // could never be re-hidden. Return to the audit origin first.
  if (originTarget && originTarget.sheetId != null) {
    var origin = sheetById(ss, originTarget.sheetId);
    if (origin && !origin.isSheetHidden() &&
        sheetNames.indexOf(origin.getName()) < 0) {
      origin.activate();
      ss.setActiveRange(origin.getRange(originTarget.row, originTarget.column));
    }
  }
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
