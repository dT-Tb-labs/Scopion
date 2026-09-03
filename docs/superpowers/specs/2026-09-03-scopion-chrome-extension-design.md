# Scopion as a Chrome extension — design

Date: 2026-09-03. Status: approved in conversation, pending user review of this text.

## Goal

Reproduce the Excel VBA add-in **ACE** (PwC "Active Cell Explorer", `frmACE`) in
Google Sheets with the *same* UI/UX: a small modeless floating window next to the
selected cell, listing that cell's precedents; arrow keys walk the list and the
grid selection follows; Enter drills in; Back returns; OK closes and restores.
Distributed as a Chrome extension so the keyboard shortcut ships with it.

## Non-goals

- Dependents (the user dropped them; also removes the whole-workbook scan).
- Painting cells in the document. Highlight is an overlay drawn by the extension.
- What-If / Goal Seek / format cycling (competitor features, not ACE).
- Browsers other than Chromium. Firefox is a possible later port.
- The Apps Script build is frozen, not deleted. Nothing is removed there.

## Why an extension (decided)

Apps Script UI is limited to a sidebar or a centred dialog; neither can be placed
next to a cell, and published add-ons cannot register keyboard shortcuts. A
content script can read the active cell, position a panel at its rectangle, move
the selection, and `chrome.commands` gives a distributable shortcut.

## Platform facts, verified live on the test sheet (2026-09-03)

| Need | Mechanism | Verified |
|---|---|---|
| Active cell address | `#t-name-box` (`<input>`) `.value`. Shows the covering named-range name when the cell is one. | yes |
| Active cell formula | `#t-formula-bar-input` `.textContent` (formula text for formula cells, value text otherwise) | yes |
| Active cell screen rect | union of the four `.active-cell-border` elements' bounding rects | yes (102×22 px) |
| Jump to a cell | set `#t-name-box` value, dispatch `input` then `keydown/keypress/keyup` Enter. Sheet-qualified `'Sheet'!A1` works. | yes |
| Jump to a hidden sheet | works, **and unhides the sheet** (document mutation) | yes |
| Sheet tabs / hidden state | `.docs-sheet-tab` → `.docs-sheet-tab-name`; hidden tab has `offsetParent === null` | yes |
| Values of referenced cells | not in the DOM (canvas). Needs Sheets REST API v4 | — |

The grid is a canvas; the extension never reads cell values from the DOM.

## Architecture

```
extension/
  manifest.json          MV3; commands, identity, host permission docs.google.com
  background.js          shortcut → message to tab; OAuth token; Sheets API fetch
  content.js             panel (Shadow DOM), DOM adapters (name box, formula bar,
                         selection rect, sheet tabs), keyboard, highlight overlay
  audit.js               orchestrates one audit: refs → fetch → resolve → rows
  adapter.js             snapshot-backed Spreadsheet/Sheet/Range facade consumed
                         by Trace.gs (same shape as the fake in run_e2e.js)
  lib/formula.js         copied from ../Formula.gs by build (source of truth stays .gs)
  lib/trace.js           copied from ../Trace.gs
  panel.css
build.js                 copies the two .gs files into extension/lib/*.js
```

Formula.gs and Trace.gs are reused unmodified. Their only SpreadsheetApp touch
points are `ValueCache` (`getSheetByName`, `getDataRange().getValues()`,
`getRange(...).getValues()`, `getLastRow/Column`, `getMaxRows/Columns`),
`buildNamedRangeMap` (`getNamedRanges` → `getName`, `getRange`), and
`findPrecedents` (`getRange(a1).getFormula()`). `findDependents` is not called.

### One audit (precedents of the selected cell)

1. Content script reads sheet name (active tab), address (name box), formula
   (formula bar). Non-formula cell → panel shows "Not a formula cell", empty list.
2. `extractRefs(formula)` (sync) → referenced ranges and names.
3. Background fetches, in **one** call:
   `GET spreadsheets/{id}?includeGridData=true&ranges=<each ref range>&fields=sheets(properties(sheetId,title,hidden,gridProperties),data(startRow,startColumn,rowData.values(effectiveValue,formattedValue,userEnteredValue)))`
   Named ranges + the sheet list are fetched once per document open
   (`fields=namedRanges,sheets.properties`) and refreshed on New Origin when the
   formula names something unknown.
4. `adapter.js` builds a snapshot spreadsheet from the response; `audit.js` runs
   `findPrecedents(ss, sheet, a1, namedRanges, new ValueCache(ss))`.
   The adapter holds a sparse map of fetched cells per sheet plus the sheet's
   `gridProperties` (rowCount/columnCount serve `getLastRow/Column` and
   `getMaxRows/Columns`). `getRange(...).getValues()` materialises any rectangle
   from the map; `getDataRange().getValues()` (reached by `ValueCache` after
   `RECT_READS_BEFORE_GRID` reads, bounded by `MAX_GRID_CACHE_CELLS`) materialises
   the whole grid the same way. A cell that was never fetched reads as blank, so
   `audit.js` must prefetch every rectangle `extractRefs` names (step 3) and
   every resolved dynamic target (step 5) before rows are built; the fixture
   test asserts that no row is built from an unfetched cell.
5. Targets resolved from OFFSET/INDEX/INDIRECT that fall outside the fetched
   rectangles are fetched in a second call (same shape), then rows are built.
6. Rows carry the current `materializeRows` fields: `sheetName, flag (H|EX|''),
   address, value (formattedValue; range rows show the sum and count like today),
   subFormula, isBlank, external, url, jumpable`. Blank target → value
   `---BLANK CELL---` and the list turns yellow (ACE behaviour).

Cost: two API round trips worst case, typically one; ~0.3–0.6 s.

### Navigation

- Walk (↑/↓ or click): jump via the name box to the row's `'Sheet'!A1`; draw the
  green highlight overlay at the new active-cell rect; the origin keeps a
  distinct outline overlay. Panel keeps keyboard focus (re-focus after the jump,
  since Sheets moves focus to its grid editor).
- New Origin (Enter / button): the current row becomes the origin; push the old
  origin onto history; re-audit.
- Back (Backspace, ←, button): pop history; jump; re-audit. Disabled when empty.
- OK (Esc, button, ×): jump back to the original origin, remove overlays,
  remember panel position, close.
- Stale selection: if the user clicked elsewhere in the grid between two panel
  actions, the panel first re-selects the last known cell (ACE's `CheckXLNone`
  branch does exactly this) so the walk never acts on a surprise cell.
- Hidden sheets: setting "Include hidden sheets" is **off by default**. When on,
  a jump into a hidden sheet unhides it (Sheets behaviour); the panel says so
  and lists the sheets it unhid. Re-hiding needs write scope and is not in v1.
- IMPORTRANGE rows: flag `EX`, not jumpable, click opens the source URL.
- Protected sheets are irrelevant (nothing is written).

### Keyboard

| Key | Action |
|---|---|
| ↑ / ↓ | move highlight, jump immediately (ACE ListBox click semantics) |
| Home / End / PageUp / PageDown | as today |
| Enter | New Origin |
| Backspace / ← | Back |
| Esc | OK |
| shortcut (default `Ctrl+Shift+A`, all platforms) | open on the selected cell; when open, re-audit the selected cell (New Origin) |

`Cmd+Shift+A` is Chrome's tab search on macOS, so it is not the default. Users
rebind at `chrome://extensions/shortcuts`.

### Panel (faithful to `frmACE`)

Geometry decoded from the form binary (points; form client 682×113, reduced
height 82 client): list `lstResults` at (0,0) 350 wide, columns
`workbook 0 | sheet 150 | flag 0 | address 90 | value 110 | colour 0 | offset 0`
in basic mode and `0|130|20|90|110|0|0` when hidden-sheet flags are on,
`0|130|0|60|80|0|80` when an OFFSET/INDEX/INDIRECT column is present (widths
switch exactly as `ChangeColumns` does). `lstOriginal` (origin formula) at
(355,0); `txtNames` (RANGE NAME: …) at (355,60); `OK` at (624,54). The
Advanced row at y=90 holds `Display Workbook information`, `Display Range Name`,
`Include Hidden Sheets` checkboxes at x=6/168/306, `Back` 474, `Advanced` 558,
`New Origin` 624. ACE's `Precedents`/`Dependents` radios (456/522) are omitted:
there is no dependents mode to switch to.

Rendering: Shadow DOM, Tahoma/system 8pt-equivalent (11px), Windows-form grey.
Draggable by its title bar "PwC Active Cell Explorer (ACE)" → "Scopion";
position persisted in `chrome.storage.local` per screen size; first open
positions the panel 8 px below-right of the active cell, clamped to the viewport.
Settings (three checkboxes) persist in `chrome.storage.sync`.

### Auth

- Scope: `https://www.googleapis.com/auth/spreadsheets.readonly` only.
- `chrome.identity.getAuthToken({interactive:true})` on first use; token cached
  by Chrome. 401 → clear token, retry once interactively.
- The user creates the Google Cloud project and an OAuth client of type
  *Chrome Extension* (needs the extension ID; the manifest pins a `key` so the
  ID is stable across machines). Consent screen in *Testing* mode is enough for
  personal use (≤100 test users). A public Web Store listing requires Google's
  sensitive-scope verification. Credentials never enter this repo: `manifest.json`
  reads the client id from `extension/oauth.local.json`, git-ignored, and the
  build fails loudly when it is missing.

### Failure modes and their answers

1. Sheets changes a DOM id/class → adapters live in one file with a self-check
   on load; missing selector shows "Sheets layout changed" in the panel instead
   of failing silently.
2. Name-box jump races the formula-bar update → after a jump, wait for the
   name box value to equal the target (poll ≤300 ms) before reading anything.
3. Focus stolen by Sheets after a jump → panel re-takes focus; a hint row shows
   when the panel does not own the keyboard (same rule as the current sidebar).
4. Quota: 60 read requests/min/user. One audit ≤2 calls; walking does not call
   the API. Fine.
5. Large ranges: `includeGridData` on `A1:Z500` is 13,000 cells; the current
   `MAX_TOTAL_CELLS` rule applies — over the cap the row shows "(N cells)" and
   the range is fetched as its top-left cell only.

## Testing

- `node run_tests.js` unchanged (parser).
- `node run_e2e.js` unchanged (Apps Script build).
- New `extension/test/audit.test.js`: `audit.js` + `adapter.js` against a recorded
  Sheets API response fixture — same scenarios as `run_e2e.js` precedents section.
- New `extension/test/dom.test.js`: adapters against a saved copy of the Sheets
  DOM fragments (name box, formula bar, four border divs, sheet tabs).
- Manual smoke on the test spreadsheet
  `11GMudQJ1VmzTgscq4sY82ceb1lHywWKisegzQPQC3P0`: open, walk B6→B1, drill B7 →
  hidden-sheet notice, Back, Esc returns to origin, Hidden Calc re-hidden by
  hand afterwards.

## Out of scope / later

- Re-hide sheets (needs `spreadsheets` write scope).
- Row/column neighbour consistency flag (implemented in the Apps Script build
  today; port to the extension in a follow-up — it needs the 3×3 block fetch).
- Firefox port, Web Store publication, i18n.
