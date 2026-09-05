# Chrome Web Store listing — Scopion

Everything the developer dashboard asks for, so a submission is copy-and-paste.
Facts here must match `extension/manifest.template.json`; change both together.

## Package

- Build: `node extension/build.js --pack --pem extension/scopion.pem` (first
  upload; later uploads without `--pem`). The zip has no `key` field, no
  credentials, no tests.
- Extension ID must stay `ihjnpijodijcbigekdihoamchmkgdgie` — the OAuth client
  (GCP project 839296398878, type "Chrome Extension") is bound to it. After
  the first upload, compare the Item ID in the dashboard with that value
  before doing anything else. If they differ, update the OAuth client's Item
  ID instead of publishing.

## Store listing

- **Name:** Scopion
- **Summary (≤132 chars):** the manifest `description`.
- **Category:** Productivity → Tools. **Language:** English.
- **Detailed description:** what it does (precedent trace for the selected
  cell, arrow-key walk with the selection following, cross-sheet and hidden
  sheets, breadcrumbs, IMPORTRANGE rows), the shortcut, and that the only
  write is re-hiding sheets it unhid. Do not mention competitors or Google
  trademarks in a way that implies endorsement ("for Google Sheets" is fine).
- **Images** (all PNG, taken from the live panel on a real model):
  - Store icon: from the manifest 128 icon (96×96 artwork, 16 px padding) —
    already rendered by `extension/icons/render.swift`.
  - Screenshots: 1–5 at 1280×800, full bleed, no padding.
  - Small promo tile: 440×280 (required for featured placement).
  - Marquee 1400×560: optional.
- **Homepage / support URL:** the repository. **Privacy policy URL:** a public
  page serving `docs/store/PRIVACY.md` (the same URL goes on the OAuth
  consent screen).

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
  - OAuth scope `.../auth/spreadsheets` (not `.readonly`) — jumping into a
    hidden sheet makes Google Sheets unhide it; Scopion hides it again when
    the panel closes. That `batchUpdate` is the only write.
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
