#!/usr/bin/env node
/**
 * build.js — the extension is assembled, not hand-maintained:
 *   lib/formula.js, lib/trace.js  are byte copies of ../Formula.gs, ../Trace.gs
 *   manifest.json                  is manifest.template.json + oauth.local.json
 * The Apps Script files stay the single source of truth for the parser and
 * resolver; credentials stay out of git.
 *
 *   node extension/build.js              full build (needs extension/oauth.local.json)
 *   node extension/build.js --lib-only   lib/ only, for tests
 *   node extension/build.js path.json    use another oauth file
 *   node extension/build.js --watch     rebuild on change and bump dev-reload.json,
 *       which the unpacked extension polls (dev-reload.js) to reload itself
 *   node extension/build.js --pack [--pem scopion.pem] [--out dir]
 *       full build, then dist/scopion-<version>.zip with only the files the
 *       extension runs — no tests, credentials, sources of the icons, or the
 *       manifest "key" (the store assigns the key; the .pem, zipped as key.pem
 *       on the first upload, is what keeps the extension ID).
 */
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');

const ext = __dirname;
const root = path.join(ext, '..');
const args = process.argv.slice(2);
const libOnly = args.includes('--lib-only');
const pack = args.includes('--pack');
const watch = args.includes('--watch');
function flag(name) { const i = args.indexOf(name); return i >= 0 ? args[i + 1] : null; }
const pemPath = flag('--pem');
const outDir = flag('--out') || path.join(root, 'dist');
const flagValues = [pemPath, outDir].filter(Boolean);
const oauthPath = args.find((a) => !a.startsWith('--') && !flagValues.includes(a)) || path.join(ext, 'oauth.local.json');

/** The files that ship. Anything not listed here stays out of the zip on purpose. */
const SHIP = ['manifest.json', 'background.js', 'content.js', 'adapter.js', 'audit.js', 'state.js', 'dom.js', 'panel.js',
  'onboarding.html', 'onboarding.js', '_locales/en/messages.json', '_locales/ja/messages.json',
  'lib/formula.js', 'lib/trace.js', 'icons/icon16.png', 'icons/icon32.png', 'icons/icon48.png', 'icons/icon128.png'];

fs.mkdirSync(path.join(ext, 'lib'), { recursive: true });
for (const [src, dst] of [['Formula.gs', 'formula.js'], ['Trace.gs', 'trace.js']]) {
  fs.copyFileSync(path.join(root, src), path.join(ext, 'lib', dst));
}
if (libOnly) {
  console.log('built extension/lib');
  process.exit(0);
}

if (!fs.existsSync(oauthPath)) {
  console.error('missing ' + oauthPath + ' — copy extension/oauth.example.json to extension/oauth.local.json and fill it in');
  process.exit(1);
}
const local = JSON.parse(fs.readFileSync(oauthPath, 'utf8'));
if (!local.client_id || !local.key) {
  console.error(oauthPath + ' needs both "client_id" and "key" (oauth credentials)');
  process.exit(1);
}
const manifest = JSON.parse(fs.readFileSync(path.join(ext, 'manifest.template.json'), 'utf8'));
manifest.oauth2.client_id = local.client_id;
manifest.key = local.key;
fs.writeFileSync(path.join(ext, 'manifest.json'), JSON.stringify(manifest, null, 2) + '\n');
console.log('built extension/lib and extension/manifest.json');
if (watch) {
  // Re-run this script on every source change; the stamp tells the extension to reload.
  const stamp = path.join(ext, 'dev-reload.json');
  const skip = /(^|\/)(lib|manifest\.json|dev-reload\.json|test|node_modules)(\/|$)/;
  let timer = null;
  const rebuild = () => {
    timer = null;
    try { execFileSync(process.execPath, [__filename, oauthPath], { stdio: 'inherit' }); } catch (e) { return; }
    fs.writeFileSync(stamp, JSON.stringify({ t: Date.now() }) + '\n');
    console.log('dev-reload stamped');
  };
  const onChange = (_, file) => { if (file && skip.test(file)) return; clearTimeout(timer); timer = setTimeout(rebuild, 300); };
  fs.watch(ext, { recursive: true }, onChange);
  for (const f of ['Formula.gs', 'Trace.gs']) fs.watch(path.join(root, f), onChange);
  rebuild();
  console.log('watching extension/ and Formula.gs, Trace.gs — Ctrl+C to stop');
  return;
}
if (!pack) process.exit(0);

// Stage the shipping files in a clean directory so the zip cannot pick up
// oauth.local.json, *.pem, test/ or the build script by accident.
const stage = fs.mkdtempSync(path.join(require('os').tmpdir(), 'scopion-pack-'));
for (const rel of SHIP) {
  fs.mkdirSync(path.dirname(path.join(stage, rel)), { recursive: true });
  fs.copyFileSync(path.join(ext, rel), path.join(stage, rel));
}
delete manifest.key;
fs.writeFileSync(path.join(stage, 'manifest.json'), JSON.stringify(manifest, null, 2) + '\n');
if (pemPath) fs.copyFileSync(pemPath, path.join(stage, 'key.pem'));
fs.mkdirSync(outDir, { recursive: true });
const zipPath = path.join(outDir, 'scopion-' + manifest.version + '.zip');
fs.rmSync(zipPath, { force: true });
execFileSync('zip', ['-qrDX', zipPath, '.'], { cwd: stage });
fs.rmSync(stage, { recursive: true, force: true });
console.log('packed ' + path.relative(process.cwd(), zipPath) + (pemPath ? ' (with key.pem)' : ''));
