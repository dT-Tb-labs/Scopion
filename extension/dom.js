/**
 * dom.js — everything that touches the Google Sheets page lives here, and
 * nowhere else. The grid is a canvas, so the only readable state is the name
 * box, the formula bar, the selection border overlay and the sheet tabs; the
 * only way to move the selection is to type into the name box.
 * Verified against the live page on 2026-09-03; if Sheets renames a selector,
 * sheetsSelfCheck() says which one and the panel shows it.
 */
var SHEETS_SEL = {
  nameBox: '#t-name-box',
  formulaBar: '#t-formula-bar-input',
  activeBorder: '.active-cell-border',
  selectionBorder: '.selection-border',
  tab: '.docs-sheet-tab',
  tabName: '.docs-sheet-tab-name',
  activeTab: '.docs-sheet-active-tab'
};
var JUMP_SETTLE_MS = 300;

function spreadsheetIdFromPath(pathname) {
  // Multi-account Chrome profiles put /u/0/ between /spreadsheets/ and /d/.
  var m = /\/spreadsheets(?:\/u\/\d+)?\/d\/([A-Za-z0-9_-]+)/.exec(pathname || '');
  return m ? m[1] : null;
}

/** "B5" / "$B$5" / "A1:C3" -> {a1: top-left}; anything else is a named range's name. */
function parseNameBox(value) {
  var v = String(value || '').trim();
  if (!v) return null;
  var m = /^\$?([A-Za-z]{1,3})\$?([0-9]{1,8})(?::\$?[A-Za-z]{1,3}\$?[0-9]{1,8})?$/.exec(v);
  if (m) return { a1: m[1].toUpperCase() + m[2] };
  return { name: v };
}

function unionRects(rects) {
  if (!rects.length) return null;
  var x1 = Infinity, y1 = Infinity, x2 = -Infinity, y2 = -Infinity;
  for (var i = 0; i < rects.length; i++) {
    var r = rects[i];
    x1 = Math.min(x1, r.left); y1 = Math.min(y1, r.top);
    x2 = Math.max(x2, r.right); y2 = Math.max(y2, r.bottom);
  }
  return { x: x1, y: y1, w: x2 - x1, h: y2 - y1 };
}

function topLeftA1(address) {
  var r = a1ToRect(address, '');
  return r ? numToCol(r.c1) + r.r1 : address;
}

function sheetsSelfCheck(doc) {
  var missing = [];
  [SHEETS_SEL.nameBox, SHEETS_SEL.formulaBar, SHEETS_SEL.tab].forEach(function (sel) {
    if (!doc.querySelector(sel)) missing.push(sel);
  });
  return missing;
}

/** Border pieces that are actually drawn. Sheets keeps hidden ones in the DOM at 0×0 at the page corner. */
function borderRects(selector) {
  var els = document.querySelectorAll(selector), rects = [];
  for (var i = 0; i < els.length; i++) {
    var r = els[i].getBoundingClientRect();
    if (r.width > 0 || r.height > 0) rects.push(r);
  }
  return rects;
}

var SheetsDom = {
  spreadsheetId: function () { return spreadsheetIdFromPath(location.pathname); },
  activeSheetName: function () {
    var el = document.querySelector(SHEETS_SEL.activeTab + ' ' + SHEETS_SEL.tabName);
    return el ? el.textContent.trim() : null;
  },
  selection: function () {
    var nb = document.querySelector(SHEETS_SEL.nameBox);
    return nb ? parseNameBox(nb.value) : null;
  },
  /** The formula text, or '' for a constant cell (the bar shows the value then). */
  formula: function () {
    var fb = document.querySelector(SHEETS_SEL.formulaBar);
    var t = fb ? fb.textContent : '';
    return t.charAt(0) === '=' ? t : '';
  },
  cellRect: function () {
    return unionRects(borderRects(SHEETS_SEL.activeBorder));
  },
  /**
   * The whole selection: a multi-cell selection draws four `.selection-border`
   * pieces around the block (verified live), a single cell draws none — then
   * the active-cell border is the selection.
   */
  selectionRect: function () {
    return unionRects(borderRects(SHEETS_SEL.selectionBorder)) || unionRects(borderRects(SHEETS_SEL.activeBorder));
  },
  sheetTabs: function () {
    var tabs = document.querySelectorAll(SHEETS_SEL.tab), out = [];
    for (var i = 0; i < tabs.length; i++) {
      var n = tabs[i].querySelector(SHEETS_SEL.tabName);
      out.push({ name: n ? n.textContent.trim() : '', hidden: tabs[i].offsetParent === null,
        active: tabs[i].classList.contains('docs-sheet-active-tab') });
    }
    return out;
  },
  /**
   * Move the grid selection by typing into the name box. Sheets rewrites the
   * box with the landed address, but a cross-sheet jump can leave the formula
   * bar (and even the name box, for a frame) still showing the previous cell —
   * reading it that early mis-resolves dynamic targets. So the settle signal is
   * not just "value changed": it is "value changed to something that parses as
   * the target address (or the target name)". The timeout covers a rejected
   * reference, where the box never settles on the target at all.
   */
  jump: function (sheetName, a1) {
    var nb = document.querySelector(SHEETS_SEL.nameBox);
    if (!nb) return Promise.reject(new Error('Sheets name box not found'));
    var target = quoteSheetName(sheetName) + '!' + a1;
    // A range address ("B1:B3") lands with its top-left cell in the name box's parse.
    var targetA1 = topLeftA1(a1).toUpperCase();
    nb.focus();
    nb.value = target;
    nb.dispatchEvent(new Event('input', { bubbles: true }));
    ['keydown', 'keypress', 'keyup'].forEach(function (type) {
      nb.dispatchEvent(new KeyboardEvent(type, { key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true, cancelable: true }));
    });
    return new Promise(function (resolve) {
      var t0 = Date.now();
      (function poll() {
        if (Date.now() - t0 > JUMP_SETTLE_MS) return resolve();
        if (nb.value !== target) {
          var parsed = parseNameBox(nb.value);
          if (parsed && (parsed.a1 === targetA1 || parsed.name)) return resolve();
        }
        setTimeout(poll, 30);
      })();
    });
  }
};
