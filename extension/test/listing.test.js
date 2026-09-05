const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');

const store = path.join(__dirname, '..', '..', 'docs', 'store');

function section(md, name) {
  const part = md.split(/^## /m).find((s) => new RegExp('^' + name + '[ \\t]*\\n').test(s));
  assert.ok(part, 'section ' + name);
  return part.slice(name.length).trim();
}

// Chrome Web Store field limits: summary 132 characters, description 16,000.
for (const lang of ['en', 'ja']) {
  test('listing copy (' + lang + ') fits the store fields and matches the manifest description', () => {
    const md = fs.readFileSync(path.join(store, 'listing.' + lang + '.md'), 'utf8');
    const summary = section(md, 'Summary'), description = section(md, 'Description');
    assert.ok(summary.length <= 132, 'summary ' + summary.length + ' > 132');
    assert.ok(description.length > 300 && description.length <= 16000, 'description length ' + description.length);
    assert.ok(!/[*_#`]/.test(description), 'description must be plain text (no markdown)');
    const locale = JSON.parse(fs.readFileSync(path.join(__dirname, '..', '_locales', lang, 'messages.json'), 'utf8'));
    assert.equal(summary, locale.appDesc.message, 'summary and manifest description are the same string');
  });
}

/** PNG width/height straight from the IHDR chunk; no image library needed. */
function pngSize(file) {
  const b = fs.readFileSync(file);
  assert.equal(b.toString('ascii', 1, 4), 'PNG', file + ' is not a PNG');
  return { w: b.readUInt32BE(16), h: b.readUInt32BE(20) };
}

// Store image specs: screenshots 1280×800, small promo tile 440×280 (developer.chrome.com/docs/webstore/images).
test('store images exist at the exact pixel sizes the dashboard accepts', () => {
  const assets = path.join(store, 'assets');
  const shots = fs.existsSync(assets) ? fs.readdirSync(assets).filter((f) => /^screenshot-\d\.png$/.test(f)) : [];
  assert.ok(shots.length >= 1 && shots.length <= 5, 'need 1-5 screenshot-N.png, found ' + shots.length);
  for (const f of shots) assert.deepEqual(pngSize(path.join(assets, f)), { w: 1280, h: 800 }, f);
  assert.deepEqual(pngSize(path.join(assets, 'promo-440x280.png')), { w: 440, h: 280 });
});
