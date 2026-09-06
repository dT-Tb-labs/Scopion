# Chrome Web Store listing — Scopion

Everything the developer dashboard asks for, so a submission is copy-and-paste.
Facts here must match `extension/manifest.template.json`; change both together.

## Package

- Build: `node extension/build.js --pack`. The zip has no `key` field, no
  credentials, no tests.
- The OAuth client (GCP project 839296398878, type "Chrome Extension") is
  bound to the unpacked ID `ihjnpijodijcbigekdihoamchmkgdgie`. The store
  assigns its own signing key, so the Item ID it shows after the first upload
  may differ: **compare, and if it differs, set the OAuth client's Item ID to
  the store's** (one field; the unpacked development copy then needs its own
  client or a second Item ID entry). Do not publish until the two match.
  `--pem extension/scopion.pem` would copy the private key into the zip as
  `key.pem` to try to keep the ID; that hands Google the signing key on an
  undocumented promise, so it is not the default (review finding, round 1).

## Store listing

- **Name:** Scopion (manifest `__MSG_appName__`; the manifest has
  `default_locale: en` and `_locales/{en,ja}`, so the dashboard offers one
  listing per language — fill **both** English and Japanese).
- **Summary (≤132 chars) and detailed description:** paste from
  `listing.en.md` / `listing.ja.md` (plain text; `extension/test/listing.test.js`
  checks the limits and that the summary equals the manifest description).
- **Category:** Productivity → Tools. **Default language:** English.
- **Version:** 1.0.0 for the first public upload (decided 2026-09-05: the walk
  is user-verified live, and a 0.x on a public listing reads as beta). The
  listing text names no other product by name; "Excel" appears only as the
  platform the original tool ran on.
- **Images** (PNG, in `docs/store/assets/`; the test checks the sizes):
  - Store icon: from the manifest 128 icon (96×96 artwork, 16 px padding) —
    rendered by `extension/icons/render.swift`.
  - Screenshots `screenshot-1..N.png`: 1280×800, full bleed, real captures of
    the panel on the test spreadsheet (`sh docs/store/capture.sh N X Y` grabs
    the 1280×800 viewport region of the screen and resamples the Retina
    capture). **Status 2026-09-06:** `screenshot-1..4.png` exist, all from
    the test spreadsheet with the Japanese Sheets UI (`?hl=en` does not
    override the account language; the EN listing reuses them until the
    account language is switched): 1 origin row, INDEX/MATCH; 2 Calc!B35
    walked one step into Inputs!B6:B10 (range row, token colours, cross-sheet
    bold); 3 the same walk on the hidden-sheet row (H flag, unhide notice);
    4 Calc!B6 drilled into Model!B6 with the Growth name row (breadcrumbs).
    Capture recipe: AppleScript `set bounds of window` to {0,39,1280,1016}
    gives a 1280×800 viewport at X=0 Y=216; `screencapture` includes the
    real cursor, so park it at the viewport's bottom-right corner first.
  - Small promo tile `promo-440x280.png`: icon + tagline, no screenshot;
    source `promo-tile.html`, rendered at a 440×280 viewport.
  - Marquee 1400×560: not made; add later if wanted.
- **Homepage:** `https://dt-tb-labs.github.io/Scopion/` (GitHub Pages of this
  repo, source folder `docs/`, Jekyll minimal theme; `superpowers/` and `store/`
  excluded). **Support URL:** `https://github.com/dT-Tb-labs/Scopion/issues`.
  **Privacy policy URL:** `https://dt-tb-labs.github.io/Scopion/privacy/`
  (`docs/privacy.md`). The same two URLs go on the OAuth consent screen; the
  domain to verify in Search Console is `dt-tb-labs.github.io` (URL-prefix
  property, HTML-file method — the file goes in `docs/`).

## Privacy practices tab

- **Single purpose:** Show which cells a selected Google Sheets formula reads
  and let the user jump between them.
- **Permission justifications**
  - `identity` — obtains an OAuth token so the Sheets API will return the
    sheet list, named ranges and referenced cell values of the open spreadsheet.
  - `storage` — three display checkboxes (sync) and the panel's last position
    (local). No spreadsheet content is stored.
  - Host `https://docs.google.com/spreadsheets/*` — the content script reads
    the selected cell and formula bar, moves the selection via the name box,
    and draws the panel.
  - Host `https://sheets.googleapis.com/*` — the API calls above, from the
    service worker.
  - OAuth scopes: `.../auth/spreadsheets.readonly` at install (manifest);
    `.../auth/spreadsheets` requested **incrementally** (chrome.identity
    `scopes` override) the first time a re-hide is needed — jumping into a
    hidden sheet makes Google Sheets unhide it; Scopion hides it again when
    the panel closes. That `batchUpdate` is the only write. Both scopes must
    be listed on the GCP consent screen and both are sensitive → verification.
- **Remote code:** No.
- **Data usage:** collects *Website content* (spreadsheet cells and formulas,
  in memory while the panel is open); no other category. Not sold, not used
  for unrelated purposes, not used for creditworthiness or lending. Certify
  the three statements.

## Google OAuth verification (blocks publishing)

The `spreadsheets` scope is sensitive. Before submitting the store item:

1. OAuth consent screen: publishing status *In production*; app name
   "Scopion"; logo; homepage and privacy policy URLs on the same verified
   domain; authorised domain added.
2. Submit for verification with a short screencast: open a sheet, press
   Ctrl+Shift+A, walk a formula into a hidden sheet, close, show the sheet
   re-hidden. Explain why read-only does not suffice (the re-hide).
3. Expect a few days to weeks; the store review can run in parallel but the
   extension will show an "unverified app" screen to users until it passes.

## Review notes (Test instructions tab)

Test account not needed — any Google account with a spreadsheet works. Steps:
open a spreadsheet, select a cell containing a formula, press Ctrl+Shift+A
(or click the toolbar icon), accept the Google sign-in, use ↑↓ to walk, → to
drill into a cell, ← to go back, Esc to return to the start and close.
On install the extension opens its bundled `onboarding.html` once
(`chrome.runtime.onInstalled`, reason `install` only); clicking the toolbar
icon on a tab that is not a spreadsheet opens the same page. No new
permissions are involved.

## First-upload order

1. Push the repo to GitHub (public), enable Pages from `main` / `docs`, and
   confirm the two URLs above resolve. The Jekyll site has never been built
   locally: check on the first deploy that the minimal theme's `default`
   layout applies, `/Scopion/privacy/` renders, and nothing from
   `docs/superpowers` or `docs/store` is published.
2. Verify `dt-tb-labs.github.io` in Search Console; set the OAuth consent
   screen to In production with the homepage, privacy URL and logo; submit
   for verification (see above).
3. `node extension/build.js --pack`, upload the zip, compare the Item ID with
   the OAuth client (fix the client if they differ), fill both language
   listings, upload the images, answer the privacy practices tab, submit for
   review.
