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
 */
const fs = require('fs');
const path = require('path');

const ext = __dirname;
const root = path.join(ext, '..');
const args = process.argv.slice(2);
const libOnly = args.includes('--lib-only');
const oauthPath = args.find((a) => a !== '--lib-only') || path.join(ext, 'oauth.local.json');

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
