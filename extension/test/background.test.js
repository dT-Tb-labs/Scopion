const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const vm = require('vm');
const path = require('path');

function load(chrome) {
  const sandbox = { console, Promise, Error, JSON, encodeURIComponent, String, Object, Array, Math, Date, chrome };
  vm.createContext(sandbox);
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'background.js'), 'utf8'), sandbox, { filename: 'background.js' });
  return sandbox;
}

/** Just enough of the chrome.* surface for background.js to register its listeners; records what it opens and sends. */
function fakeChrome() {
  const listeners = {}, created = [], sent = [];
  const on = (name) => ({ addListener: (fn) => { listeners[name] = fn; } });
  const chrome = {
    noReceiver: false, // true = no content script: Chrome sets lastError and calls back with no reply
    runtime: { onMessage: on('message'), onInstalled: on('installed'), lastError: null, getURL: (p) => 'chrome-extension://id/' + p },
    tabs: {
      create: (o) => created.push(o.url),
      sendMessage: (id, msg, cb) => {
        sent.push([id, msg.type]);
        chrome.runtime.lastError = chrome.noReceiver ? { message: 'Receiving end does not exist' } : null;
        cb(chrome.noReceiver ? undefined : { ok: true }); // content.js replies {ok:true} synchronously
        chrome.runtime.lastError = null;
      }
    },
    commands: { onCommand: on('command') },
    action: { onClicked: on('clicked') },
    identity: { getAuthToken() {}, removeCachedAuthToken() {} }
  };
  return { chrome, listeners, created, sent };
}

test('urls ask for exactly the fields the adapter reads', () => {
  const B = load();
  const u = new URL(B.metaUrl('abc'));
  assert.equal(u.pathname, '/v4/spreadsheets/abc');
  assert.equal(u.searchParams.get('fields'), 'namedRanges,sheets.properties(sheetId,title,hidden,gridProperties(rowCount,columnCount))');
  const g = new URL(B.gridUrl('abc', ["'Inputs'!B1", "'O''Connor'!A1:B2"]));
  assert.equal(g.searchParams.get('includeGridData'), 'true');
  assert.deepEqual(g.searchParams.getAll('ranges'), ["'Inputs'!B1", "'O''Connor'!A1:B2"]);
  assert.equal(g.searchParams.get('fields'), 'sheets(properties.sheetId,data(startRow,startColumn,rowData.values(effectiveValue,formattedValue,userEnteredValue)))');
});

function fakeDeps(statuses, tokens) {
  const log = [];
  return {
    log,
    fetch: async (url, opts) => {
      const status = statuses.shift();
      log.push(['fetch', opts.headers.Authorization, status]);
      return { status, ok: status >= 200 && status < 300, json: async () => ({ status }), text: async () => 'body ' + status };
    },
    getToken: async (interactive) => { log.push(['getToken', interactive]); const t = tokens.shift(); if (!t) throw new Error('no token'); return t; },
    removeToken: async (t) => { log.push(['removeToken', t]); }
  };
}

test('a good token is used once, silently', async () => {
  const B = load();
  const deps = fakeDeps([200], ['t1']);
  assert.deepEqual(await B.apiGet('u', deps), { status: 200 });
  assert.deepEqual(deps.log, [['getToken', false], ['fetch', 'Bearer t1', 200]]);
});

test('a rejected token is dropped and the call retried once with an interactive token', async () => {
  const B = load();
  const deps = fakeDeps([401, 200], ['stale', 'fresh']);
  assert.deepEqual(await B.apiGet('u', deps), { status: 200 });
  assert.deepEqual(deps.log, [['getToken', false], ['fetch', 'Bearer stale', 401], ['removeToken', 'stale'], ['getToken', true], ['fetch', 'Bearer fresh', 200]]);
});

test('no cached token yet: fall through to the interactive prompt', async () => {
  const B = load();
  const deps = fakeDeps([200], [null, 'first']);
  await B.apiGet('u', deps);
  assert.deepEqual(deps.log.slice(0, 2), [['getToken', false], ['getToken', true]]);
});

test('after a declined consent the API is tried silently: no interactive prompt, a "not signed in" error instead', async () => {
  const B = load();
  const deps = fakeDeps([200], [null, 'would-be-prompted']);
  await assert.rejects(B.apiGet('u', deps, false), /not signed in/);
  assert.deepEqual(deps.log, [['getToken', false]]); // never getToken(true)
  const deps401 = fakeDeps([401, 200], ['stale', 'fresh']);
  await assert.rejects(B.apiGet('u', deps401, false), /not signed in: token rejected/);
  assert.deepEqual(deps401.log, [['getToken', false], ['fetch', 'Bearer stale', 401], ['removeToken', 'stale']]);
});

test('re-hide posts one updateSheetProperties per sheet with the token', async () => {
  const B = load();
  const seen = [];
  const deps = {
    fetch: async (url, opts) => { seen.push([url, opts.method, opts.headers['Content-Type'], opts.headers.Authorization, JSON.parse(opts.body)]); return { status: 200, ok: true, json: async () => ({}), text: async () => '' }; },
    getToken: async () => 'tok', removeToken: async () => {}
  };
  await B.apiPost(B.rehideUrl('abc'), B.rehideBody([2, 5]), deps);
  assert.deepEqual(seen[0].slice(0, 4), ['https://sheets.googleapis.com/v4/spreadsheets/abc:batchUpdate', 'POST', 'application/json', 'Bearer tok']);
  assert.deepEqual(seen[0][4], { requests: [
    { updateSheetProperties: { properties: { sheetId: 2, hidden: true }, fields: 'hidden' } },
    { updateSheetProperties: { properties: { sheetId: 5, hidden: true }, fields: 'hidden' } }
  ] });
});

test('a toolbar click with no content script opens onboarding; a spreadsheet tab gets the reload hint', () => {
  const B = load();
  assert.equal(B.onboardingUrlFor(undefined), 'onboarding.html');
  assert.equal(B.onboardingUrlFor({ id: 3 }), 'onboarding.html'); // tab.url absent: a host we have no permission for
  assert.equal(B.onboardingUrlFor({ id: 3, url: 'https://docs.google.com/document/d/x/edit' }), 'onboarding.html');
  assert.equal(B.onboardingUrlFor({ id: 3, url: 'https://docs.google.com/spreadsheets/' }), 'onboarding.html'); // Sheets home: nothing to reload
  assert.equal(B.onboardingUrlFor({ id: 3, url: 'https://docs.google.com/spreadsheets/d/x/edit#gid=0' }), 'onboarding.html#reload');
  assert.equal(B.onboardingUrlFor({ id: 3, url: 'https://docs.google.com/spreadsheets/u/1/d/x/edit' }), 'onboarding.html#reload');
});

test('onboarding opens once on install, not on update', () => {
  const f = fakeChrome(); load(f.chrome);
  f.listeners.installed({ reason: 'install' });
  f.listeners.installed({ reason: 'update' });
  f.listeners.installed({ reason: 'chrome_update' });
  assert.deepEqual(f.created, ['chrome-extension://id/onboarding.html']);
});

test('toolbar click: toggles when a content script answers, opens onboarding when none does; the shortcut stays silent', () => {
  const sheets = { id: 7, url: 'https://docs.google.com/spreadsheets/d/x/edit' };
  const f = fakeChrome(); load(f.chrome);
  f.listeners.clicked(sheets);
  assert.deepEqual(f.sent, [[7, 'scopion:toggle']]);
  assert.deepEqual(f.created, []);
  f.chrome.noReceiver = true;
  f.listeners.clicked(sheets);                 // spreadsheet open since before the install
  f.listeners.clicked({ id: 8 });              // any other site: tab.url absent
  f.listeners.command('toggle-scopion', { id: 8 });
  f.listeners.command('something-else', sheets);
  assert.deepEqual(f.created, ['chrome-extension://id/onboarding.html#reload', 'chrome-extension://id/onboarding.html']);
  assert.deepEqual(f.sent.slice(1), [[7, 'scopion:toggle'], [8, 'scopion:toggle'], [8, 'scopion:toggle']]);
});

test('other errors surface with the status and a slice of the body', async () => {
  const B = load();
  await assert.rejects(B.apiGet('u', fakeDeps([403], ['t'])), /Sheets API 403: body 403/);
});
