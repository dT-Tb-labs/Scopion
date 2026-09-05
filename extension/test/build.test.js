const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');

const ext = path.join(__dirname, '..');
const root = path.join(ext, '..');

test('build copies the Apps Script sources into lib/ unchanged', () => {
  execFileSync('node', [path.join(ext, 'build.js'), '--lib-only']);
  assert.equal(fs.readFileSync(path.join(ext, 'lib/formula.js'), 'utf8'), fs.readFileSync(path.join(root, 'Formula.gs'), 'utf8'));
  assert.equal(fs.readFileSync(path.join(ext, 'lib/trace.js'), 'utf8'), fs.readFileSync(path.join(root, 'Trace.gs'), 'utf8'));
});

test('build renders manifest.json from the template and the local oauth file', () => {
  const tmp = path.join(ext, 'test', 'oauth.tmp.json');
  // A developer's real manifest.json (loaded unpacked in Chrome) must survive
  // the test run: park it and put it back, never delete it.
  const manifestPath = path.join(ext, 'manifest.json');
  const parked = fs.existsSync(manifestPath) ? fs.readFileSync(manifestPath) : null;
  fs.writeFileSync(tmp, JSON.stringify({ client_id: 'cid.apps.googleusercontent.com', key: 'MIIB' }));
  try {
    execFileSync('node', [path.join(ext, 'build.js'), tmp]);
    const m = JSON.parse(fs.readFileSync(path.join(ext, 'manifest.json'), 'utf8'));
    assert.equal(m.manifest_version, 3);
    // Localised fields reach the rendered manifest as placeholders Chrome resolves from _locales.
    assert.equal(m.default_locale, 'en');
    assert.equal(m.name, '__MSG_appName__');
    assert.equal(m.description, '__MSG_appDesc__');
    assert.equal(m.action.default_title, '__MSG_actionTitle__');
    assert.equal(m.commands['toggle-scopion'].description, '__MSG_cmdToggle__');
    assert.equal(m.oauth2.client_id, 'cid.apps.googleusercontent.com');
    assert.equal(m.key, 'MIIB');
    // Read/write: the one write is re-hiding sheets the walk had to unhide.
    assert.deepEqual(m.oauth2.scopes, ['https://www.googleapis.com/auth/spreadsheets']);
    assert.equal(m.commands['toggle-scopion'].suggested_key.default, 'Ctrl+Shift+A');
    // Chrome turns "Ctrl" into Command on macOS; MacCtrl is the Control key the listing and onboarding promise.
    assert.equal(m.commands['toggle-scopion'].suggested_key.mac, 'MacCtrl+Shift+A');
    assert.deepEqual(m.host_permissions, ['https://docs.google.com/spreadsheets/*', 'https://sheets.googleapis.com/*']);
  } finally {
    fs.unlinkSync(tmp);
    if (parked) fs.writeFileSync(manifestPath, parked);
    else fs.rmSync(manifestPath, { force: true });
  }
});

test('--pack zips only the shipping files, without the manifest key, plus key.pem when given', () => {
  const tmp = path.join(ext, 'test', 'oauth.tmp.json');
  const out = fs.mkdtempSync(path.join(require('os').tmpdir(), 'scopion-dist-'));
  const pem = path.join(out, 'fake.pem');
  const manifestPath = path.join(ext, 'manifest.json');
  const parked = fs.existsSync(manifestPath) ? fs.readFileSync(manifestPath) : null;
  fs.writeFileSync(tmp, JSON.stringify({ client_id: 'cid.apps.googleusercontent.com', key: 'MIIB' }));
  fs.writeFileSync(pem, 'not a real key');
  try {
    execFileSync('node', [path.join(ext, 'build.js'), tmp, '--pack', '--pem', pem, '--out', out]);
    const version = JSON.parse(fs.readFileSync(path.join(ext, 'manifest.template.json'), 'utf8')).version;
    const zip = path.join(out, 'scopion-' + version + '.zip');
    const listed = execFileSync('unzip', ['-Z1', zip]).toString().trim().split('\n').sort();
    assert.deepEqual(listed, ['_locales/en/messages.json', '_locales/ja/messages.json', 'adapter.js', 'audit.js', 'background.js', 'content.js', 'dom.js',
      'icons/icon128.png', 'icons/icon16.png', 'icons/icon32.png', 'icons/icon48.png', 'key.pem', 'lib/formula.js', 'lib/trace.js', 'manifest.json',
      'onboarding.html', 'onboarding.js', 'panel.js', 'state.js']);
    const m = JSON.parse(execFileSync('unzip', ['-p', zip, 'manifest.json']).toString());
    assert.equal(m.key, undefined);
    assert.equal(m.oauth2.client_id, 'cid.apps.googleusercontent.com');
    // The developer's unpacked manifest keeps its key: only the zip drops it.
    assert.equal(JSON.parse(fs.readFileSync(manifestPath, 'utf8')).key, 'MIIB');
  } finally {
    fs.unlinkSync(tmp);
    fs.rmSync(out, { recursive: true, force: true });
    if (parked) fs.writeFileSync(manifestPath, parked);
    else fs.rmSync(manifestPath, { force: true });
  }
});

test('every __MSG__ the manifest and onboarding page use exists in both locales, and the locales agree', () => {
  const read = (p) => JSON.parse(fs.readFileSync(path.join(ext, p), 'utf8'));
  const en = read('_locales/en/messages.json'), ja = read('_locales/ja/messages.json');
  assert.deepEqual(Object.keys(ja).sort(), Object.keys(en).sort());
  const template = fs.readFileSync(path.join(ext, 'manifest.template.json'), 'utf8');
  const used = new Set([...template.matchAll(/__MSG_(\w+)__/g)].map((m) => m[1]));
  const page = fs.readFileSync(path.join(ext, 'onboarding.html'), 'utf8');
  for (const m of page.matchAll(/data-i18n="(\w+)"/g)) used.add(m[1]);
  assert.ok(used.size >= 8, 'template and page should reference locale keys');
  for (const key of used) {
    assert.ok(en[key] && en[key].message, 'en missing ' + key);
    assert.ok(ja[key] && ja[key].message, 'ja missing ' + key);
  }
  assert.equal(JSON.parse(template).default_locale, 'en');
  // The store's summary field and the manifest description share the 132-character limit.
  assert.ok(en.appDesc.message.length <= 132, 'en description over 132: ' + en.appDesc.message.length);
  assert.ok(ja.appDesc.message.length <= 132, 'ja description over 132: ' + ja.appDesc.message.length);
});

test('build refuses to render a manifest without credentials', () => {
  assert.throws(
    () => execFileSync('node', [path.join(ext, 'build.js'), path.join(ext, 'test', 'does-not-exist.json')], { stdio: 'pipe' }),
    (e) => e.status === 1 && /oauth\.example\.json/.test(String(e.stderr))
  );
});
