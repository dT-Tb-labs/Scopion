/**
 * panel.js — the ACE window (frmACE), rebuilt: a modeless strip you can drag,
 * a multi-column list of precedents on the left, the origin formula and the
 * range names on the right, OK/Back/New Origin/Advanced along the bottom.
 * Shadow DOM keeps Sheets' CSS out and ours in. No framework: the list is
 * small and re-rendered whole.
 */
var PANEL_CSS = [
  ':host{all:initial}',
  '.win{position:fixed;z-index:2147483000;width:910px;background:#f0f0f0;border:1px solid #6d6d6d;box-shadow:0 6px 24px rgba(0,0,0,.35);',
  '  font:11px Tahoma,"Segoe UI",system-ui,sans-serif;color:#000;user-select:none}',
  '.title{height:22px;line-height:22px;padding:0 8px;background:linear-gradient(#fdfdfd,#dcdcdc);border-bottom:1px solid #a0a0a0;cursor:move;display:flex;justify-content:space-between}',
  '.title .x{cursor:pointer;padding:0 6px}.title .x:hover{background:#c42b1c;color:#fff}',
  '.body{position:relative;height:88px}.body.adv{height:129px}',
  '.list{position:absolute;left:0;top:0;width:467px;height:82px;overflow-y:auto;background:#fff;border:1px inset #999;outline:none}',
  '.list.blank{background:#ffff99}.list.unfocused .row.on{background:#d9d9d9;color:#000}',
  '.row{display:grid;height:16px;line-height:16px;white-space:nowrap;cursor:default}',
  '.row>span{overflow:hidden;text-overflow:ellipsis;padding:0 3px}.row .v{text-align:right;font-variant-numeric:tabular-nums}',
  '.row.on{background:#0078d7;color:#fff}.row.ext{color:#5a3d8a}.row .flag{text-align:center}',
  '.origin{position:absolute;left:473px;top:0;width:430px;height:60px;background:#fff;border:1px inset #999;padding:2px 4px;overflow:auto;white-space:pre-wrap;word-break:break-all;font-family:Consolas,Menlo,monospace}',
  '.names{position:absolute;left:473px;top:64px;width:350px;height:18px;background:#fff;border:1px inset #999;padding:0 4px;line-height:18px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}',
  'button{position:absolute;height:22px;padding:0 8px;font:inherit;background:linear-gradient(#fff,#e1e1e1);border:1px solid #707070;border-radius:2px;cursor:pointer}',
  'button:disabled{color:#888;cursor:default}button:focus{outline:1px dotted #000}',
  '.ok{left:832px;top:62px;width:70px}.adv-row{position:absolute;top:100px;left:0;right:0;height:26px;display:none}.body.adv .adv-row{display:block}',
  '.adv-row label{position:absolute;top:5px;white-space:nowrap}.adv-row input{vertical-align:-2px}',
  '.back{left:632px;top:100px;width:70px}.advbtn{left:744px;top:100px;width:80px}.neworigin{left:832px;top:100px;width:70px}',
  '.body:not(.adv) .back,.body:not(.adv) .advbtn,.body:not(.adv) .neworigin{top:64px;display:none}',
  '.notice{position:absolute;left:473px;top:84px;width:350px;color:#8a4b00;font-size:10px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}',
  '.busy .list{opacity:.6}',
  '.hl{position:fixed;z-index:2147482999;pointer-events:none;box-sizing:border-box}',
  '.hl.walk{background:rgba(0,176,80,.28);border:2px solid #00b050}.hl.origin{border:2px dashed #00b050}'
].join('\n');

var COLS_BASIC = [200, 0, 120, 147, 0];
var COLS_FULL = [147, 27, 80, 107, 107];

function createPanel(handlers) {
  var host = document.createElement('div');
  host.id = 'scopion-host';
  var root = host.attachShadow({ mode: 'open' });
  var style = document.createElement('style');
  style.textContent = PANEL_CSS;
  root.appendChild(style);

  var win = document.createElement('div');
  win.className = 'win';
  win.hidden = true;
  win.innerHTML =
    '<div class="title"><span>Scopion — Active Cell Explorer</span><span class="x" title="OK (Esc)">✕</span></div>' +
    '<div class="body">' +
      '<div class="list" tabindex="0" role="listbox" aria-label="Precedents"></div>' +
      '<div class="origin"></div>' +
      '<div class="names"></div>' +
      '<div class="notice"></div>' +
      '<button class="ok">OK</button>' +
      '<div class="adv-row">' +
        '<label style="left:8px"><input type="checkbox" data-key="showExternal"> Display Workbook information</label>' +
        '<label style="left:224px"><input type="checkbox" data-key="showNames"> Display Range Name</label>' +
        '<label style="left:408px"><input type="checkbox" data-key="includeHidden"> Include Hidden Sheets</label>' +
      '</div>' +
      '<button class="back">Back</button>' +
      '<button class="advbtn">Advanced</button>' +
      '<button class="neworigin">New Origin</button>' +
    '</div>';
  root.appendChild(win);

  var walkHl = document.createElement('div'); walkHl.className = 'hl walk'; walkHl.hidden = true;
  var originHl = document.createElement('div'); originHl.className = 'hl origin'; originHl.hidden = true;
  root.appendChild(walkHl); root.appendChild(originHl);

  var $ = function (sel) { return win.querySelector(sel); };
  var list = $('.list'), body = $('.body');
  var pos = { x: 80, y: 120 };
  var view = null;

  function esc(s) { return String(s === undefined || s === null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }

  function renderList() {
    var rows = view.rows || [];
    var anyFlag = rows.some(function (r) { return r.flag; });
    var anyDyn = rows.some(function (r) { return r.dynamic || r.external; });
    var cols = anyFlag || anyDyn ? COLS_FULL : COLS_BASIC;
    var tpl = cols.map(function (w) { return w + 'px'; }).join(' ');
    list.classList.toggle('blank', !!view.hasBlank);
    list.classList.toggle('unfocused', list.getRootNode().activeElement !== list);
    list.innerHTML = rows.length ? '' : '<div class="row" style="grid-template-columns:1fr"><span>' +
      esc(view.originFormula ? 'No references found.' : 'Not a formula cell.') + '</span></div>';
    rows.forEach(function (r, i) {
      var el = document.createElement('div');
      el.className = 'row' + (i === view.active ? ' on' : '') + (r.external ? ' ext' : '');
      el.style.gridTemplateColumns = tpl;
      el.setAttribute('role', 'option');
      el.setAttribute('aria-selected', i === view.active ? 'true' : 'false');
      el.dataset.index = i;
      el.innerHTML =
        '<span title="' + esc(r.sheetName) + '">' + esc(r.sheetName) + '</span>' +
        '<span class="flag">' + esc(r.flag) + '</span>' +
        '<span>' + esc(r.address) + '</span>' +
        '<span class="v">' + esc(r.value) + '</span>' +
        '<span title="' + esc(r.subFormula) + '">' + esc(r.dynamic || r.external ? r.subFormula : '') + '</span>';
      list.appendChild(el);
    });
    var on = list.querySelector('.row.on');
    if (on && on.scrollIntoView) on.scrollIntoView({ block: 'nearest' });
  }

  /**
   * v: {originFormula, names, rows, active, hasBlank, unresolved, canBack,
   *     advanced, settings, busy, unhidden}. unhidden is the list of sheet
   * names Sheets unhid during this session's jumps — shown as part of the
   * notice, since a side-channel notice() call gets overwritten by the next
   * render() (every caller renders right after a jump).
   */
  function render(v) {
    view = v;
    body.className = 'body' + (v.advanced ? ' adv' : '') + (v.busy ? ' busy' : '');
    $('.origin').textContent = v.originFormula ? 'ORIGIN FORMULA: ' + v.originFormula : 'ORIGIN FORMULA: (not a formula)';
    $('.names').textContent = v.names && v.names.length ? 'RANGE NAME: ' + v.names.join(' | ') : '';
    $('.back').disabled = !v.canBack;
    $('.advbtn').textContent = v.advanced ? 'Simple' : 'Advanced';
    ['showExternal', 'showNames', 'includeHidden'].forEach(function (k) {
      $('input[data-key="' + k + '"]').checked = !!(v.settings && v.settings[k]);
    });
    var unresolvedText = v.unresolved && v.unresolved.length ?
      v.unresolved.length + ' dynamic reference' + (v.unresolved.length > 1 ? 's' : '') + ' unresolved: ' +
        v.unresolved.map(function (u) { return u.raw + ' (' + u.reason + ')'; }).join('; ') : '';
    var unhiddenText = v.unhidden && v.unhidden.length ?
      'Sheets unhid: ' + v.unhidden.join(', ') + ' — re-hide by hand when done.' : '';
    notice(unresolvedText && unhiddenText ? unresolvedText + ' · ' + unhiddenText : (unresolvedText || unhiddenText));
    renderList();
  }

  function notice(text) { $('.notice').textContent = text || ''; $('.notice').title = text || ''; }

  function place(p) {
    var w = 910, h = view && view.advanced ? 151 : 110;
    pos = { x: Math.max(0, Math.min(p.x, innerWidth - w)), y: Math.max(0, Math.min(p.y, innerHeight - h)) };
    win.style.left = pos.x + 'px'; win.style.top = pos.y + 'px';
  }

  // Keyboard — the ACE ListBox: arrows walk (and jump), Enter is New Origin, Back is Backspace/←, Esc is OK.
  list.addEventListener('keydown', function (ev) {
    var k = ev.key;
    if (!view) return;
    if (k === 'ArrowDown') handlers.onWalk(Math.min(view.rows.length - 1, Math.max(0, view.active) + (view.active < 0 ? 0 : 1)));
    else if (k === 'ArrowUp') handlers.onWalk(Math.max(0, view.active - 1));
    else if (k === 'Home') handlers.onWalk(0);
    else if (k === 'End') handlers.onWalk(view.rows.length - 1);
    else if (k === 'PageDown') handlers.onWalk(Math.min(view.rows.length - 1, view.active + 5));
    else if (k === 'PageUp') handlers.onWalk(Math.max(0, view.active - 5));
    else if (k === 'Enter' || k === 'ArrowRight') handlers.onDrill();
    else if (k === 'Backspace' || k === 'ArrowLeft') handlers.onBack();
    else if (k === 'Escape') handlers.onClose();
    else return;
    ev.preventDefault(); ev.stopPropagation();
  });
  list.addEventListener('click', function (ev) {
    var row = ev.target.closest('.row');
    if (!row || row.dataset.index === undefined) return;
    var i = Number(row.dataset.index);
    if (view.rows[i] && view.rows[i].external) handlers.onExternal(i); else handlers.onWalk(i);
    list.focus();
  });
  list.addEventListener('focus', function () { list.classList.remove('unfocused'); });
  list.addEventListener('blur', function () { list.classList.add('unfocused'); });
  win.addEventListener('keydown', function (ev) { if (ev.key === 'Escape') { handlers.onClose(); ev.preventDefault(); } });

  $('.ok').onclick = handlers.onClose;
  $('.x').onclick = handlers.onClose;
  $('.back').onclick = handlers.onBack;
  $('.neworigin').onclick = handlers.onNewOrigin;
  $('.advbtn').onclick = handlers.onAdvanced;
  win.querySelectorAll('input[data-key]').forEach(function (cb) {
    cb.onchange = function () { handlers.onSetting(cb.dataset.key, cb.checked); };
  });

  // Drag by the title bar; the position is the caller's to persist (getPosition()).
  var drag = null;
  $('.title').addEventListener('mousedown', function (ev) {
    if (ev.target.classList.contains('x')) return;
    drag = { dx: ev.clientX - pos.x, dy: ev.clientY - pos.y };
    ev.preventDefault();
  });
  window.addEventListener('mousemove', function (ev) { if (drag) place({ x: ev.clientX - drag.dx, y: ev.clientY - drag.dy }); });
  window.addEventListener('mouseup', function () { if (drag) { drag = null; if (handlers.onMoved) handlers.onMoved(pos); } });

  function setHl(el, rect) {
    if (!rect) { el.hidden = true; return; }
    el.hidden = false;
    el.style.left = rect.x + 'px'; el.style.top = rect.y + 'px';
    el.style.width = rect.w + 'px'; el.style.height = rect.h + 'px';
  }

  document.documentElement.appendChild(host);

  return {
    isOpen: function () { return !win.hidden; },
    open: function (nearRect, savedPos) {
      win.hidden = false;
      if (savedPos) place(savedPos);
      else if (nearRect) place({ x: nearRect.x, y: nearRect.y + nearRect.h + 8 });
      else place(pos);
    },
    close: function () { win.hidden = true; setHl(walkHl, null); setHl(originHl, null); notice(''); },
    render: render,
    focus: function () { list.focus(); },
    setPosition: place,
    getPosition: function () { return pos; },
    highlight: function (rect, kind) { setHl(kind === 'origin' ? originHl : walkHl, rect); },
    notice: notice
  };
}
