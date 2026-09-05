const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const vm = require('vm');
const path = require('path');

function load() {
  const sandbox = { console, Promise, Error, JSON, encodeURIComponent, String, Object, Array, Math, Date };
  vm.createContext(sandbox);
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'background.js'), 'utf8'), sandbox, { filename: 'background.js' });
  return sandbox;
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

test('other errors surface with the status and a slice of the body', async () => {
  const B = load();
  await assert.rejects(B.apiGet('u', fakeDeps([403], ['t'])), /Sheets API 403: body 403/);
});
