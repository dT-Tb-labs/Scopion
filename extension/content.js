/**
 * content.js — glue. The shortcut arrives as a message from background.js;
 * everything else is: read the selection, audit it, show the strip, and turn
 * key presses into jumps. Sheets moves keyboard focus to its grid after every
 * jump, so the panel takes it back each time.
 */
(function () {
  if (window.__scopionLoaded) return; // Chrome re-injects on extension reload
  window.__scopionLoaded = true;

  var S = {
    spreadsheetId: null, snap: null, settings: null, panel: null,
    walk: null, advanced: false, busy: false, lastJump: null, unhidden: [], savedPos: null, tick: null
  };

  function rpc(msg) {
    return new Promise(function (resolve, reject) {
      chrome.runtime.sendMessage(msg, function (r) {
        if (chrome.runtime.lastError) return reject(new Error(chrome.runtime.lastError.message));
        if (!r || !r.ok) return reject(new Error(r ? r.error : 'no response from background'));
        resolve(r.data);
      });
    });
  }
  /**
   * Debug trace on the host element's dataset. Content scripts live in an
   * isolated world, so page-side tooling (and a curious developer's console)
   * can only see what we put in the DOM. Last 20 entries, newest last.
   */
  var TRACE = [];
  function trace(label, data) {
    TRACE.push({ t: Date.now(), label: label, data: data });
    if (TRACE.length > 20) TRACE.shift();
    var host = document.getElementById('scopion-host');
    if (host) host.dataset.trace = JSON.stringify(TRACE);
  }

  var api = {
    getGrid: function (ranges) {
      trace('grid:request', ranges);
      return rpc({ type: 'scopion:grid', spreadsheetId: S.spreadsheetId, ranges: ranges }).then(function (resp) {
        trace('grid:response', (resp.sheets || []).map(function (s) {
          return { sheetId: s.properties && s.properties.sheetId, blocks: (s.data || []).map(function (b) {
            return { startRow: b.startRow || 0, startColumn: b.startColumn || 0, rows: (b.rowData || []).length,
              firstCell: b.rowData && b.rowData[0] && b.rowData[0].values && b.rowData[0].values[0] };
          }) };
        }));
        return resp;
      });
    }
  };

  function ensureSnapshot(force) {
    if (S.snap && !force) return Promise.resolve(S.snap);
    return rpc({ type: 'scopion:meta', spreadsheetId: S.spreadsheetId }).then(function (meta) {
      S.snap = new Snapshot(meta);
      S.apiError = '';
      return S.snap;
    }, function (e) {
      // The Sheets API refuses non-native documents (an .xlsx opened in Sheets)
      // and can fail for quota or auth reasons; the walk still works from the
      // page alone, so degrade instead of dying. The notice says why.
      trace('meta:error', String(e && e.message ? e.message : e));
      S.apiError = String(e && e.message ? e.message : e);
      S.snap = Snapshot.fromTabs(SheetsDom.sheetTabs());
      return S.snap;
    });
  }

  function storageGet(area, keys) {
    return new Promise(function (resolve) { chrome.storage[area].get(keys, function (v) { resolve(v || {}); }); });
  }
  function loadPrefs() {
    return Promise.all([storageGet('sync', ['settings']), storageGet('local', ['panelPos'])]).then(function (r) {
      S.settings = Object.assign({}, DEFAULT_SETTINGS, r[0].settings || {});
      S.savedPos = r[1].panelPos || null;
    });
  }

  function view(extra) {
    var w = S.walk;
    return Object.assign({
      originFormula: w ? w.originFormula : '', originSheet: w ? w.origin.sheetName : '',
      origin: w ? w.origin : null, history: w ? w.history : [],
      names: w ? w.names : [], rows: w ? w.rows : [], active: w ? w.active : -1,
      hasBlank: w ? w.hasBlank : false, unresolved: w ? w.unresolved : [], canBack: !!(w && w.history.length),
      advanced: S.advanced, settings: S.settings, busy: S.busy, unhidden: S.unhidden,
      dataless: !!(S.snap && S.snap.dataless), apiError: S.apiError || ''
    }, extra || {});
  }
  function render() { S.panel.render(view()); }
  function setBusy(b) { S.busy = b; if (S.panel.isOpen()) render(); }
  function fail(e) {
    trace('error', String(e && e.message ? e.message : e));
    setBusy(false);
    S.panel.notice('Scopion: ' + (e && e.message ? e.message : e));
  }

  /** Where the grid selection is right now, as {sheetName, a1}. A named cell is resolved through the snapshot. */
  function currentCell() {
    var sheetName = SheetsDom.activeSheetName();
    var sel = SheetsDom.selection();
    if (!sheetName || !sel) throw new Error('Select a cell first.');
    if (sel.a1) return { sheetName: sheetName, a1: sel.a1 };
    var map = buildNamedRangeMap(S.snap);
    var hit = map[sel.name.toUpperCase()];
    if (!hit) throw new Error('Cannot resolve "' + sel.name + '" to a cell.');
    return { sheetName: hit.rect.sheetName, a1: numToCol(hit.rect.c1) + hit.rect.r1 };
  }

  /** Audit a cell. Refresh metadata once when the formula names something the snapshot does not know. */
  function audit(cell, formula, retried) {
    trace('audit:start', { cell: cell, formula: formula, settings: S.settings });
    return ensureSnapshot().then(function (snap) {
      trace('snapshot', { sheets: snap.getSheets().map(function (s) { return s.getName() + '#' + s.getSheetId() + (s.isSheetHidden() ? ' (hidden)' : ''); }),
        names: snap.getNamedRanges().map(function (n) { return n.getName(); }) });
      return auditCell(api, snap, cell.sheetName, cell.a1, formula, S.settings);
    }).then(function (res) {
      trace('audit:result', { origin: res.origin, rows: res.rows.length, unresolved: res.unresolved, unknownNames: res.unknownNames });
      if (res.unknownNames.length && !retried) return ensureSnapshot(true).then(function () { return audit(cell, formula, true); });
      return res;
    });
  }

  function walkFrom(res, history) {
    var w = createWalk({ sheetName: res.origin.sheetName, a1: res.origin.a1 }, withOriginRow(res.origin, res.rows));
    w.history = history || [];
    w.originFormula = res.origin.formula; w.names = res.origin.names; w.hasBlank = res.hasBlank; w.unresolved = res.unresolved;
    return w;
  }

  /** ACE's CheckXLNone: if the user clicked elsewhere in the grid, go back to where the walk was before acting. */
  function resync() {
    if (!S.lastJump) return Promise.resolve();
    var now;
    try { now = currentCell(); } catch (e) { return Promise.resolve(); }
    if (now.sheetName === S.lastJump.sheetName && now.a1 === S.lastJump.a1) return Promise.resolve();
    return jump(S.lastJump);
  }

  /** cell = {sheetName, a1 (top-left), address? (full range — ACE selects the whole block)}. */
  function jump(cell) {
    var sheet = S.snap && S.snap.getSheetByName(cell.sheetName);
    var wasHidden = sheet && sheet.isSheetHidden() && S.unhidden.indexOf(cell.sheetName) < 0;
    return SheetsDom.jump(cell.sheetName, cell.address || cell.a1).then(function () {
      S.lastJump = cell;
      // The caller's render() (part of the view now, not a side-channel notice) shows this.
      if (wasHidden) S.unhidden.push(cell.sheetName);
      paintHighlights();
      S.panel.focus();
    });
  }

  function paintHighlights() {
    var rect = SheetsDom.selectionRect();
    var w = S.walk;
    var onOrigin = w && S.lastJump && S.lastJump.sheetName === w.origin.sheetName && S.lastJump.a1 === w.origin.a1;
    S.panel.highlight(rect && !onOrigin ? rect : null, 'walk');
    S.panel.highlight(rect && onOrigin ? rect : null, 'origin');
  }

  function open() {
    setBusy(true);
    // Force a fresh snapshot on every open: values and formulas change between
    // opens, and one extra metadata call per open is cheap.
    return loadPrefs().then(function () { return ensureSnapshot(true); }).then(function () {
      var cell = currentCell();
      return audit(cell, SheetsDom.formula()).then(function (res) {
        S.walk = walkFrom(res, []);
        S.lastJump = { sheetName: res.origin.sheetName, a1: res.origin.a1 };
        S.unhidden = [];
        setBusy(false);
        S.panel.open(SheetsDom.selectionRect(), S.savedPos);
        render();
        paintHighlights();
        S.panel.focus();
        if (S.tick) clearInterval(S.tick);
        S.tick = setInterval(paintHighlights, 250); // the grid scrolls under us; follow the cell
      });
    }).catch(function (e) {
      // A failure before the strip is up must still be visible somewhere.
      if (!S.panel.isOpen()) { S.panel.open(SheetsDom.selectionRect(), S.savedPos); S.panel.render(view()); }
      fail(e);
    });
  }

  function walkTo(index) {
    if (S.busy || !S.walk) return;
    var row = S.walk.rows[index];
    if (!row) return;
    S.walk = walkSelect(S.walk, index);
    render();
    if (!row.jumpable) return;
    setBusy(true);
    resync().then(function () { return jump({ sheetName: row.sheetName, a1: topLeftA1(row.address), address: row.address }); })
      .then(function () {
        // Without API data, a visited constant cell still shows its value in the formula bar.
        if (S.snap && S.snap.dataless && S.walk && S.walk.rows[index] === row) {
          var bar = document.querySelector(SHEETS_SEL.formulaBar);
          var text = bar ? bar.textContent : '';
          if (text && text.charAt(0) !== '=' && !row.isOrigin) row.value = text;
        }
        setBusy(false);
      }).catch(fail);
  }

  /** New Origin / Enter: audit the cell the walk is standing on. */
  function drill() {
    if (S.busy || !S.walk) return;
    var row = S.walk.rows[S.walk.active];
    if (!row || !row.jumpable || row.isOrigin) return; // the origin is already the origin
    var cell = { sheetName: row.sheetName, a1: topLeftA1(row.address) };
    setBusy(true);
    resync().then(function () { return jump(cell); })
      .then(function () { return audit(cell, SheetsDom.formula()); })
      .then(function (res) {
        var next = walkDrill(S.walk, cell, res.rows);
        S.walk = walkFrom(res, next.history);
        setBusy(false); render(); paintHighlights(); S.panel.focus();
      }).catch(fail);
  }

  function newOrigin() {
    if (S.busy) return;
    setBusy(true);
    var cell;
    try { cell = currentCell(); } catch (e) { return fail(e); }
    jump(cell).then(function () { return audit(cell, SheetsDom.formula()); })
      .then(function (res) {
        var history = S.walk ? S.walk.history.concat([S.walk.origin]) : [];
        S.walk = walkFrom(res, history);
        setBusy(false); render(); paintHighlights(); S.panel.focus();
      }).catch(fail);
  }

  function back() {
    if (S.busy || !S.walk) return;
    var b = walkBack(S.walk);
    if (!b) return;
    setBusy(true);
    resync().then(function () { return jump(b.origin); })
      .then(function () { return audit(b.origin, SheetsDom.formula()); })
      .then(function (res) {
        S.walk = walkFrom(res, b.history);
        setBusy(false); render(); paintHighlights(); S.panel.focus();
      }).catch(fail);
  }

  /** Breadcrumb click: return to the origin at history[depth], dropping everything after it. */
  function backTo(depth) {
    if (S.busy || !S.walk || depth < 0 || depth >= S.walk.history.length) return;
    var target = S.walk.history[depth], history = S.walk.history.slice(0, depth);
    setBusy(true);
    resync().then(function () { return jump(target); })
      .then(function () { return audit(target, SheetsDom.formula()); })
      .then(function (res) {
        S.walk = walkFrom(res, history);
        setBusy(false); render(); paintHighlights(); S.panel.focus();
      }).catch(fail);
  }

  /**
   * Sheets unhid these when the walk entered them; hide them again on the way
   * out (ACE restored visibility on exit). The sheet the user is left on
   * cannot be hidden, so it is skipped. Needs cell data (an API-backed
   * snapshot) for the sheet ids.
   */
  function rehideUnhidden(stay) {
    if (!S.unhidden.length || !S.snap || S.snap.dataless) return Promise.resolve();
    var keep = stay ? SheetsDom.activeSheetName() : (S.walk ? walkRoot(S.walk).sheetName : null);
    var ids = S.unhidden.filter(function (n) { return n !== keep; })
      .map(function (n) { var s = S.snap.getSheetByName(n); return s ? s.getSheetId() : null; })
      .filter(function (id) { return id !== null && id !== undefined; });
    S.unhidden = [];
    if (!ids.length) return Promise.resolve();
    return rpc({ type: 'scopion:rehide', spreadsheetId: S.spreadsheetId, sheetIds: ids });
  }

  /** OK / Esc return to where the session started; Enter (stay=true) keeps the selection where the walk left it. */
  function close(stay) {
    if (S.tick) { clearInterval(S.tick); S.tick = null; }
    var root = S.walk && !stay ? walkRoot(S.walk) : null;
    chrome.storage.local.set({ panelPos: S.panel.getPosition() });
    var warn = function (what) { return function (e) { console.warn('Scopion: ' + what + ': ' + (e && e.message ? e.message : e)); }; };
    var p = root ? SheetsDom.jump(root.sheetName, root.a1) : Promise.resolve();
    // The panel must close even if the return jump or the re-hide fails — nobody sees a notice once it is gone.
    p.catch(warn('could not return to the origin'))
      .then(function () { return rehideUnhidden(!!stay); }).catch(warn('could not re-hide sheets'))
      .then(function () {
        S.panel.close();
        SheetsDom.focusGrid(); // the list had the keyboard; give it back to the cells
        S.walk = null; S.lastJump = null; S.busy = false;
      });
  }

  function setSetting(key, value) {
    if (S.busy) { render(); return; } // puts the checkbox back to the stored value
    S.settings[key] = value;
    chrome.storage.sync.set({ settings: S.settings });
    // Settings change what the list shows, so re-audit the current origin (ACE: ChangeColumns + rebuild).
    if (!S.walk) return render();
    setBusy(true);
    audit(S.walk.origin, S.walk.originFormula).then(function (res) {
      S.walk = walkFrom(res, S.walk.history);
      setBusy(false); render();
    }).catch(fail);
  }

  function init() {
    S.spreadsheetId = SheetsDom.spreadsheetId();
    if (!S.spreadsheetId) return;
    S.panel = createPanel({
      onWalk: walkTo,
      onDrill: drill,
      onBack: back,
      onCrumb: backTo,
      onClose: function () { close(false); },
      onCommit: function () { close(true); },
      onNewOrigin: newOrigin,
      onAdvanced: function () { S.advanced = !S.advanced; render(); },
      onSetting: setSetting,
      onExternal: function (i) { var r = S.walk && S.walk.rows[i]; if (r && r.url) window.open(r.url, '_blank', 'noopener'); },
      onMoved: function (pos) { chrome.storage.local.set({ panelPos: pos }); }
    });
    function toggle() {
      if (S.busy) return; // an open()/newOrigin() is already in flight
      var missing = sheetsSelfCheck(document);
      if (missing.length) { S.panel.open(null, S.savedPos); S.panel.render(view()); S.panel.notice('Sheets layout changed; missing ' + missing.join(', ')); return; }
      if (!S.panel.isOpen()) open(); else newOrigin();
    }
    chrome.runtime.onMessage.addListener(function (msg, sender, sendResponse) {
      if (msg && msg.type === 'scopion:toggle') {
        toggle();
        sendResponse({ ok: true }); // background.js reads "no response" as "no content script here"
      }
    });
    // Same door for page-side tooling (the smoke harness cannot press the
    // real shortcut): document.dispatchEvent(new CustomEvent('scopion:toggle')).
    document.addEventListener('scopion:toggle', toggle);
  }

  init();
})();
