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
  var api = { getGrid: function (ranges) { return rpc({ type: 'scopion:grid', spreadsheetId: S.spreadsheetId, ranges: ranges }); } };

  function ensureSnapshot(force) {
    if (S.snap && !force) return Promise.resolve(S.snap);
    return rpc({ type: 'scopion:meta', spreadsheetId: S.spreadsheetId }).then(function (meta) {
      S.snap = new Snapshot(meta);
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
      originFormula: w ? w.originFormula : '', names: w ? w.names : [], rows: w ? w.rows : [], active: w ? w.active : -1,
      hasBlank: w ? w.hasBlank : false, unresolved: w ? w.unresolved : [], canBack: !!(w && w.history.length),
      advanced: S.advanced, settings: S.settings, busy: S.busy
    }, extra || {});
  }
  function render() { S.panel.render(view()); }
  function setBusy(b) { S.busy = b; if (S.panel.isOpen()) render(); }
  function fail(e) { setBusy(false); S.panel.notice('Scopion: ' + (e && e.message ? e.message : e)); }

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
    return ensureSnapshot().then(function (snap) {
      return auditCell(api, snap, cell.sheetName, cell.a1, formula, S.settings);
    }).then(function (res) {
      if (res.unknownNames.length && !retried) return ensureSnapshot(true).then(function () { return audit(cell, formula, true); });
      return res;
    });
  }

  function walkFrom(res, history) {
    var w = createWalk({ sheetName: res.origin.sheetName, a1: res.origin.a1 }, res.rows);
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

  function jump(cell) {
    var sheet = S.snap && S.snap.getSheetByName(cell.sheetName);
    var wasHidden = sheet && sheet.isSheetHidden() && S.unhidden.indexOf(cell.sheetName) < 0;
    return SheetsDom.jump(cell.sheetName, cell.a1).then(function () {
      S.lastJump = cell;
      if (wasHidden) {
        S.unhidden.push(cell.sheetName);
        S.panel.notice('Sheets unhid "' + cell.sheetName + '" to show it. Re-hide it by hand when you are done.');
      }
      paintHighlights();
      S.panel.focus();
    });
  }

  function paintHighlights() {
    var rect = SheetsDom.cellRect();
    var w = S.walk;
    var onOrigin = w && S.lastJump && S.lastJump.sheetName === w.origin.sheetName && S.lastJump.a1 === w.origin.a1;
    S.panel.highlight(rect && !onOrigin ? rect : null, 'walk');
    S.panel.highlight(rect && onOrigin ? rect : null, 'origin');
  }

  function open() {
    setBusy(true);
    return loadPrefs().then(function () { return ensureSnapshot(); }).then(function () {
      var cell = currentCell();
      return audit(cell, SheetsDom.formula()).then(function (res) {
        S.walk = walkFrom(res, []);
        S.lastJump = { sheetName: res.origin.sheetName, a1: res.origin.a1 };
        S.unhidden = [];
        setBusy(false);
        S.panel.open(SheetsDom.cellRect(), S.savedPos);
        render();
        paintHighlights();
        S.panel.focus();
        if (S.tick) clearInterval(S.tick);
        S.tick = setInterval(paintHighlights, 250); // the grid scrolls under us; follow the cell
      });
    }).catch(function (e) {
      // A failure before the strip is up must still be visible somewhere.
      if (!S.panel.isOpen()) { S.panel.open(SheetsDom.cellRect(), S.savedPos); S.panel.render(view()); }
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
    resync().then(function () { return jump({ sheetName: row.sheetName, a1: topLeftA1(row.address) }); })
      .then(function () { setBusy(false); }).catch(fail);
  }

  /** New Origin / Enter: audit the cell the walk is standing on. */
  function drill() {
    if (S.busy || !S.walk) return;
    var row = S.walk.rows[S.walk.active];
    if (!row || !row.jumpable) return;
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

  /** OK / Esc: return to where the session started, clear the overlays, close. */
  function close() {
    if (S.tick) { clearInterval(S.tick); S.tick = null; }
    var root = S.walk ? walkRoot(S.walk) : null;
    chrome.storage.local.set({ panelPos: S.panel.getPosition() });
    var p = root ? SheetsDom.jump(root.sheetName, root.a1) : Promise.resolve();
    // The panel must close even if the return jump fails — nobody sees a notice once it is gone.
    p.catch(function (e) {
      console.warn('Scopion: could not return to the origin: ' + (e && e.message ? e.message : e));
    }).then(function () {
      S.panel.close();
      S.walk = null; S.lastJump = null; S.busy = false;
      if (S.unhidden.length) console.info('Scopion left these sheets visible: ' + S.unhidden.join(', '));
    });
  }

  function setSetting(key, value) {
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
      onClose: close,
      onNewOrigin: newOrigin,
      onAdvanced: function () { S.advanced = !S.advanced; render(); },
      onSetting: setSetting,
      onExternal: function (i) { var r = S.walk && S.walk.rows[i]; if (r && r.url) window.open(r.url, '_blank', 'noopener'); },
      onMoved: function (pos) { chrome.storage.local.set({ panelPos: pos }); }
    });
    chrome.runtime.onMessage.addListener(function (msg) {
      if (!msg || msg.type !== 'scopion:toggle') return;
      var missing = sheetsSelfCheck(document);
      if (missing.length) { S.panel.open(null, S.savedPos); S.panel.render(view()); S.panel.notice('Sheets layout changed; missing ' + missing.join(', ')); return; }
      if (!S.panel.isOpen()) open(); else newOrigin();
    });
  }

  init();
})();
