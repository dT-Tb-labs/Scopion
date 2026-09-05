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
    assert.equal(m.oauth2.client_id, 'cid.apps.googleusercontent.com');
    assert.equal(m.key, 'MIIB');
    // Read/write: the one write is re-hiding sheets the walk had to unhide.
    assert.deepEqual(m.oauth2.scopes, ['https://www.googleapis.com/auth/spreadsheets']);
    assert.equal(m.commands['toggle-scopion'].suggested_key.default, 'Ctrl+Shift+A');
    assert.equal(m.commands['toggle-scopion'].suggested_key.mac, 'Ctrl+Shift+A');
    assert.deepEqual(m.host_permissions, ['https://docs.google.com/spreadsheets/*', 'https://sheets.googleapis.com/*']);
  } finally {
    fs.unlinkSync(tmp);
    if (parked) fs.writeFileSync(manifestPath, parked);
    else fs.rmSync(manifestPath, { force: true });
  }
});

test('build refuses to render a manifest without credentials', () => {
  assert.throws(
    () => execFileSync('node', [path.join(ext, 'build.js'), path.join(ext, 'test', 'does-not-exist.json')], { stdio: 'pipe' }),
    (e) => e.status === 1 && /oauth\.example\.json/.test(String(e.stderr))
  );
});
