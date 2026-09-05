/**
 * panel.js — the ACE window (frmACE), rebuilt as a liquid-glass strip: a
 * modeless, draggable pane that blurs the grid behind it, a multi-column list
 * of precedents on the left, the origin formula and the range names on the
 * right, OK/Back/New Origin/Advanced along the bottom. Shadow DOM keeps
 * Sheets' CSS out and ours in. No framework: the list is small and
 * re-rendered whole.
 *
 * Reading aids: every reference row gets a colour, and the same colour marks
 * that reference inside the origin formula; rows on another sheet than the
 * origin are bold, rows on the same sheet are dimmed.
 */
var PANEL_CSS = [
  ':host{all:initial}',
  '.win{position:fixed;z-index:2147483000;width:910px;border-radius:14px;box-sizing:border-box;',
  '  background:rgba(250,251,255,.66);-webkit-backdrop-filter:blur(22px) saturate(170%);backdrop-filter:blur(22px) saturate(170%);',
  '  box-shadow:0 18px 50px rgba(20,30,60,.28),0 2px 8px rgba(20,30,60,.10),inset 0 1px 0 rgba(255,255,255,.95),inset 0 0 0 1px rgba(255,255,255,.45),0 0 0 1px rgba(30,40,80,.12);',
  '  font:12px -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,"Helvetica Neue",sans-serif;color:#1c2333;user-select:none;',
  '  animation:scopion-pop .16s ease-out}',
  '@keyframes scopion-pop{from{opacity:0;transform:translateY(6px) scale(.985)}to{opacity:1;transform:none}}',
  '.title{height:30px;line-height:30px;padding:0 10px 0 12px;cursor:move;display:flex;align-items:center;justify-content:space-between;',
  '  border-bottom:1px solid rgba(30,40,80,.08);font-weight:600;letter-spacing:.01em}',
  '.title .name{display:flex;align-items:center;gap:7px}.title .mark{width:14px;height:14px;border-radius:4px;background:linear-gradient(135deg,#1c2a6b,#070b24);position:relative}',
  '.title .mark::after{content:"";position:absolute;left:4px;top:4px;width:4px;height:4px;border-radius:50%;background:#ffd28a;box-shadow:5px 3px 0 -1px #f7d774,-2px 5px 0 -1.5px #f7d774}',
  '.title .x{width:22px;height:22px;line-height:22px;text-align:center;border-radius:50%;cursor:pointer;color:#3a4460;',
  '  background:rgba(255,255,255,.55);box-shadow:inset 0 1px 0 rgba(255,255,255,.9),0 0 0 1px rgba(30,40,80,.12);transition:background .12s,transform .08s}',
  '.title .x:hover{background:#ff5f57;color:#fff}.title .x:active{transform:scale(.94)}',
  '.body{position:relative;height:100px}.body.adv{height:140px}',
  '.names .trail{flex:1;min-width:0;overflow:hidden;white-space:nowrap}.names .rn{flex:none;max-width:45%;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:#5b6580}',
  '.crumb{display:inline-block;max-width:150px;height:18px;line-height:18px;padding:0 8px;border-radius:999px;vertical-align:middle;overflow:hidden;text-overflow:ellipsis;',
  '  color:#1c2333;cursor:pointer;background:rgba(255,255,255,.55);box-shadow:inset 0 1px 0 rgba(255,255,255,.9),0 0 0 1px rgba(30,40,80,.12);transition:background .12s}',
  '.crumb:hover{background:rgba(255,255,255,.9)}',
  '.crumb.cur{cursor:default;font-weight:600;color:#0b2a6b;background:linear-gradient(180deg,rgba(84,132,255,.30),rgba(84,132,255,.18));box-shadow:inset 0 1px 0 rgba(255,255,255,.7),0 0 0 1px rgba(84,132,255,.45)}',
  '.names .sep{margin:0 4px;color:#8a93aa}',
  '.pane{background:rgba(255,255,255,.42);border-radius:10px;box-shadow:inset 0 1px 0 rgba(255,255,255,.85),0 0 0 1px rgba(30,40,80,.10)}',
  '.list{position:absolute;left:8px;top:6px;width:459px;height:88px;overflow-y:auto;outline:none;padding:3px}',
  '.list.blank{background:rgba(255,214,0,.28)}',
  '.row{display:grid;height:20px;line-height:20px;white-space:nowrap;cursor:default;border-radius:7px;transition:background .12s,box-shadow .12s}',
  '.row>span{overflow:hidden;text-overflow:ellipsis;padding:0 4px}.row .v{text-align:right;font-variant-numeric:tabular-nums}',
  '.row .flag{text-align:center;font-size:10px;color:#7a2e00}',
  '.row .dot{display:inline-block;width:8px;height:8px;border-radius:50%;margin:0 6px 0 0;vertical-align:-1px;box-shadow:inset 0 0 0 1px rgba(0,0,0,.08)}',
  '.row.same .sheet{opacity:.5}.row.x .sheet{font-weight:700}.row.ext{color:#5a3d8a}.row.is-origin{font-weight:700}',
  '.row.on{background:linear-gradient(180deg,rgba(84,132,255,.30),rgba(84,132,255,.18));box-shadow:inset 0 1px 0 rgba(255,255,255,.7),0 0 0 1px rgba(84,132,255,.45);color:#0b2a6b}',
  '.list.unfocused .row.on{background:rgba(30,40,80,.10);color:#1c2333;box-shadow:0 0 0 1px rgba(30,40,80,.16)}',
  '.row:hover:not(.on){background:rgba(255,255,255,.55)}',
  '.origin{position:absolute;left:475px;top:6px;width:427px;height:60px;padding:5px 8px;overflow:auto;white-space:pre-wrap;word-break:break-all;box-sizing:border-box;',
  '  font:12px ui-monospace,SFMono-Regular,Menlo,Consolas,monospace;line-height:1.45}',
  '.origin .lbl{color:#5b6580;font-weight:600;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;font-size:10px;letter-spacing:.06em;margin-right:6px}',
  '.tok{border-radius:4px;padding:0 3px;font-weight:600}',
  '.names{position:absolute;left:475px;top:70px;width:340px;height:24px;line-height:22px;padding:0 6px;box-sizing:border-box;display:flex;align-items:center;gap:6px;font-size:11px;color:#3a4460}',
  '.names .lbl{color:#5b6580;font-weight:600;font-size:10px;letter-spacing:.06em;margin-right:6px}',
  'button{position:absolute;height:24px;padding:0 12px;font:inherit;font-weight:600;color:#1c2333;cursor:pointer;border:0;border-radius:999px;',
  '  background:rgba(255,255,255,.6);box-shadow:inset 0 1px 0 rgba(255,255,255,.9),0 0 0 1px rgba(30,40,80,.14),0 1px 2px rgba(20,30,60,.08);transition:background .12s,transform .08s}',
  'button:hover{background:rgba(255,255,255,.9)}button:active{transform:scale(.97)}button:disabled{opacity:.45;cursor:default}',
  'button:focus-visible{outline:2px solid rgba(84,132,255,.7);outline-offset:1px}',
  '.ok{left:822px;top:70px;width:80px;color:#fff;background:linear-gradient(180deg,#5c8cff,#3b6cf0);box-shadow:inset 0 1px 0 rgba(255,255,255,.45),0 2px 6px rgba(59,108,240,.35)}',
  '.ok:hover{background:linear-gradient(180deg,#6b97ff,#4576f5)}',
  '.adv-row{position:absolute;top:104px;left:8px;right:8px;height:30px;display:none}.body.adv .adv-row{display:block}',
  '.adv-row label{position:absolute;top:6px;white-space:nowrap;color:#3a4460}.adv-row input{vertical-align:-2px;accent-color:#3b6cf0}',
  '.back{left:610px;top:106px;width:72px}.advbtn{left:704px;top:106px;width:96px}.neworigin{left:812px;top:106px;width:90px}',
  '.body:not(.adv) .back,.body:not(.adv) .advbtn,.body:not(.adv) .neworigin{display:none}',
  '.notice{position:absolute;left:475px;top:96px;width:340px;color:#8a4b00;font-size:10px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}',
  '.body.adv .notice{top:136px}',
  '.busy .list{opacity:.6}',
  '.hl{position:fixed;z-index:2147482999;pointer-events:none;box-sizing:border-box;border-radius:3px;transition:left .12s,top .12s,width .12s,height .12s}',
  '.hl.walk{background:rgba(255,178,96,.22);border:2px solid #e39a3b;box-shadow:0 0 0 3px rgba(227,154,59,.18)}',
  '.hl.origin{border:2px dashed #e39a3b}'
].join('\n');

var COLS_BASIC = [200, 0, 120, 133, 0];
var COLS_FULL = [147, 27, 80, 107, 92];
var TOKEN_COLORS = ['#e4572e', '#2b6fdb', '#1f9d55', '#8e44ad', '#e67e22', '#0e9aa7', '#d63384', '#7a5c2e'];

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
    '<div class="title"><span class="name"><span class="mark"></span>Scopion</span><span class="x" title="OK (Esc)">✕</span></div>' +
    '<div class="body">' +
      '<div class="list pane" tabindex="0" role="listbox" aria-label="Precedents"></div>' +
      '<div class="origin pane"></div>' +
      '<div class="names pane" aria-label="History and range names"></div>' +
      '<div class="notice"></div>' +
      '<button class="ok">OK</button>' +
      '<div class="adv-row">' +
        '<label style="left:4px"><input type="checkbox" data-key="showExternal"> Display Workbook information</label>' +
        '<label style="left:220px"><input type="checkbox" data-key="showNames"> Display Range Name</label>' +
        '<label style="left:404px"><input type="checkbox" data-key="includeHidden"> Include Hidden Sheets</label>' +
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

  /** Reference rows get a colour each; the origin row has none. */
  function rowColors(rows) {
    var k = 0;
    return rows.map(function (r) { return r.isOrigin ? null : TOKEN_COLORS[k++ % TOKEN_COLORS.length]; });
  }

  /**
   * The origin formula with each reference wrapped in its row's colour.
   * Trace.gs positions index the formula without its leading "=", and a
   * sheet-qualified ref's position points at the cell part, so the raw text
   * is searched for just before that position. A dynamic call (OFFSET(...))
   * colours only its function name — its argument refs are rows of their own.
   */
  function formulaHtml(formula, rows, colors) {
    var body = formula.charAt(0) === '=' ? formula.slice(1) : formula;
    var marks = [];
    rows.forEach(function (r, i) {
      if (r.isOrigin || !r.subFormula || !colors[i]) return;
      var raw = r.subFormula;
      var len = r.dynamic && raw.indexOf('(') > 0 ? raw.indexOf('(') : raw.length;
      var idx = body.indexOf(raw, Math.max(0, (r.pos || 0) - raw.length));
      if (idx < 0) idx = body.indexOf(raw);
      if (idx < 0) return;
      marks.push({ start: idx, end: idx + len, color: colors[i] });
    });
    marks.sort(function (a, b) { return a.start - b.start; });
    var out = '', cursor = 0;
    marks.forEach(function (m) {
      if (m.start < cursor) return; // overlapping the previous mark: keep the earlier one
      out += esc(body.slice(cursor, m.start)) +
        '<span class="tok" style="background:' + m.color + '22;color:' + m.color + '">' + esc(body.slice(m.start, m.end)) + '</span>';
      cursor = m.end;
    });
    return (formula.charAt(0) === '=' ? '=' : '') + out + esc(body.slice(cursor));
  }

  function renderList(colors) {
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
      var sameSheet = !r.external && r.sheetName === view.originSheet;
      el.className = 'row' + (i === view.active ? ' on' : '') + (r.external ? ' ext' : '') +
        (r.isOrigin ? ' is-origin' : (sameSheet ? ' same' : ' x'));
      el.style.gridTemplateColumns = tpl;
      el.setAttribute('role', 'option');
      el.setAttribute('aria-selected', i === view.active ? 'true' : 'false');
      el.dataset.index = i;
      var dot = colors[i] ? '<span class="dot" style="background:' + colors[i] + '"></span>' : '<span class="dot" style="background:#e39a3b"></span>';
      el.innerHTML =
        '<span class="sheet" title="' + esc(r.sheetName) + '">' + dot + esc(r.sheetName) + '</span>' +
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
   * v: {originFormula, originSheet, names, rows, active, hasBlank, unresolved,
   *     canBack, advanced, settings, busy, unhidden}. unhidden is the list of
   * sheet names Sheets unhid during this session's jumps — shown as part of
   * the notice, since a side-channel notice() call gets overwritten by the
   * next render() (every caller renders right after a jump).
   */
  function render(v) {
    view = v;
    body.className = 'body' + (v.advanced ? ' adv' : '') + (v.busy ? ' busy' : '');
    var colors = rowColors(v.rows || []);
    $('.origin').innerHTML = '<span class="lbl">ORIGIN FORMULA</span>' +
      (v.originFormula ? formulaHtml(v.originFormula, v.rows || [], colors) : '<span style="color:#5b6580">(not a formula)</span>');
    // Breadcrumbs: every origin drilled through, oldest first, then the current one.
    var crumbs = (v.history || []).map(function (h, i) {
      var label = h.sheetName + '!' + h.a1;
      return '<span class="crumb" data-depth="' + i + '" title="Back to ' + esc(label) + '">' + esc(label) + '</span>';
    });
    if (v.origin) crumbs.push('<span class="crumb cur">' + esc(v.origin.sheetName + '!' + v.origin.a1) + '</span>');
    // The trail shares the box beside OK with the range names (ACE's txtNames), which are usually empty.
    $('.names').innerHTML = '<span class="trail">' + crumbs.join('<span class="sep">›</span>') + '</span>' +
      (v.names && v.names.length ? '<span class="rn" title="' + esc(v.names.join(' | ')) + '"><span class="lbl">RANGE NAME</span>' + esc(v.names.join(' | ')) + '</span>' : '');
    $('.back').disabled = !v.canBack;
    $('.advbtn').textContent = v.advanced ? 'Simple' : 'Advanced';
    ['showExternal', 'showNames', 'includeHidden'].forEach(function (k) {
      $('input[data-key="' + k + '"]').checked = !!(v.settings && v.settings[k]);
    });
    var unresolvedText = v.unresolved && v.unresolved.length ?
      v.unresolved.length + ' dynamic reference' + (v.unresolved.length > 1 ? 's' : '') + ' unresolved: ' +
        v.unresolved.map(function (u) { return u.raw + ' (' + u.reason + ')'; }).join('; ') : '';
    var unhiddenText = v.unhidden && v.unhidden.length ?
      'Sheets unhid: ' + v.unhidden.join(', ') + (v.dataless ? ' — re-hide by hand when done.' : ' — hidden again when Scopion closes.') : '';
    var datalessText = v.dataless ? 'No cell data for this file (Excel format?) — references and jumps only.' : '';
    notice([datalessText, unresolvedText, unhiddenText].filter(Boolean).join(' · '));
    renderList(colors);
  }

  function notice(text) { $('.notice').textContent = text || ''; $('.notice').title = text || ''; }

  function height() { return 30 + (view && view.advanced ? 140 : 100); }

  function place(p) {
    var w = 910, h = height();
    pos = { x: Math.max(0, Math.min(p.x, innerWidth - w)), y: Math.max(0, Math.min(p.y, innerHeight - h)) };
    win.style.left = pos.x + 'px'; win.style.top = pos.y + 'px';
  }

  /** Never cover the selection: to its right when that fits, else below, else above. */
  function placeNear(r) {
    var w = 910, h = height(), gap = 10;
    if (r.x + r.w + gap + w <= innerWidth) place({ x: r.x + r.w + gap, y: r.y });
    else if (r.y + r.h + gap + h <= innerHeight) place({ x: r.x, y: r.y + r.h + gap });
    else if (r.y - gap - h >= 0) place({ x: r.x, y: r.y - gap - h });
    else place({ x: r.x, y: innerHeight - h });
  }

  // Keyboard — the ACE ListBox: arrows walk (and jump), → drills in, ←/Backspace goes back, Enter stays, Esc returns.
  list.addEventListener('keydown', function (ev) {
    var k = ev.key;
    if (!view) return;
    if (k === 'ArrowDown') handlers.onWalk(Math.min(view.rows.length - 1, Math.max(0, view.active) + (view.active < 0 ? 0 : 1)));
    else if (k === 'ArrowUp') handlers.onWalk(Math.max(0, view.active - 1));
    else if (k === 'Home') handlers.onWalk(0);
    else if (k === 'End') handlers.onWalk(view.rows.length - 1);
    else if (k === 'PageDown') handlers.onWalk(Math.min(view.rows.length - 1, view.active + 5));
    else if (k === 'PageUp') handlers.onWalk(Math.max(0, view.active - 5));
    else if (k === 'Enter') handlers.onCommit();
    else if (k === 'ArrowRight') handlers.onDrill();
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
  $('.names').addEventListener('click', function (ev) {
    var c = ev.target.closest('.crumb');
    if (c && c.dataset.depth !== undefined && handlers.onCrumb) handlers.onCrumb(Number(c.dataset.depth));
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
    /** Opens next to the selection rectangle (see placeNear); a saved position is used only when there is no rectangle. */
    open: function (nearRect, savedPos) {
      win.hidden = false;
      if (nearRect) placeNear(nearRect);
      else if (savedPos) place(savedPos);
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
