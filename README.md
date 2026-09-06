# ACE for Google Sheets

A port of the Excel VBA add-in **ACE** (PwC; Nick Gage 2002, Tim Mosedale 2006,
Chris McNeill 2007, Matthew Howard-Cairns 2009, Ryuji Moritani 2012) to Google
Apps Script.

Select a cell, press the shortcut, and a sidebar lists every cell that cell
reads — or every cell that reads it. Click a row to jump there; the sidebar
re-audits from the new position, so you walk a model one hop at a time.

## Install

1. Open the spreadsheet → **Extensions ▸ Apps Script**.
2. Create the files below and paste in the matching contents:
   `Code.gs`, `Formula.gs`, `Trace.gs`, `Test.gs`, `Sidebar.html`.
3. Project settings → tick **Show "appsscript.json" manifest file**, then replace
   the manifest with the `appsscript.json` in this repo. That is what registers
   the macro and its shortcut.
4. Reload the spreadsheet. An **ACE** menu appears, and `Ctrl+Alt+Shift+1` runs it.
5. Optional: run `runTests()` once from the editor to confirm the parser works
   in your environment (it needs no document access).

## Getting `Cmd+Shift+A` back

Google Sheets has no API for custom shortcuts — macros can only occupy
`Ctrl+Alt+Shift+1..9`, on macOS as well as Windows. The original chord is
restored on the macOS side instead:

1. Install [Karabiner-Elements](https://karabiner-elements.pqrs.org/).
2. Copy `karabiner-ace.json` into `~/.config/karabiner/assets/complex_modifications/`.
3. Karabiner-Elements → **Complex Modifications ▸ Add rule** → enable
   *Cmd+Shift+A -> Ctrl+Alt+Shift+1 (ACE macro)*.

The rule is scoped to browsers, so `Cmd+Shift+A` keeps its normal meaning
everywhere else.

## Chrome extension (Scopion)

The extension reproduces the ACE window next to the selected cell and ships its
own shortcut. It reads the sheet through the Sheets API, so it needs an OAuth
client that you create once. Its only write is hiding again the sheets it had
to unhide for a walk (see "Keys" below). Install from the Chrome Web Store when
it is listed; the steps below are for running it from source.

1. **Key** (makes the extension ID stable, which the OAuth client is tied to):
   ```bash
   openssl genrsa -out extension/scopion.pem 2048
   openssl rsa -in extension/scopion.pem -pubout -outform DER | openssl base64 -A; echo
   ```
   The printed base64 is the `key` value. Extension ID:
   ```bash
   openssl rsa -in extension/scopion.pem -pubout -outform DER | node -e "const c=require('crypto');let b=[];process.stdin.on('data',d=>b.push(d)).on('end',()=>{const h=c.createHash('sha256').update(Buffer.concat(b)).digest('hex').slice(0,32);console.log([...h].map(x=>String.fromCharCode(97+parseInt(x,16))).join(''))})"
   ```
2. **Google Cloud**: create a project → *APIs & Services ▸ Library* → enable
   **Google Sheets API** → *OAuth consent screen*: External, publishing status
   *Testing*, add your Google account as a test user, add both scopes
   `https://www.googleapis.com/auth/spreadsheets.readonly` and
   `https://www.googleapis.com/auth/spreadsheets` → *Credentials ▸ Create
   credentials ▸ OAuth client ID*, application type **Chrome Extension**, Item ID =
   the extension ID from step 1. Copy the client ID.
3. `cp extension/oauth.example.json extension/oauth.local.json` and fill in
   `client_id` and `key`. This file is git-ignored.
4. `node extension/build.js` → writes `extension/lib/` and `extension/manifest.json`.
5. `chrome://extensions` → Developer mode → **Load unpacked** → the `extension/`
   folder. Check the ID matches step 1.
6. Open a spreadsheet, select a formula cell, press **Ctrl+Shift+A**. The first
   run asks for Google sign-in and access to your spreadsheets. Rebind the key at
   `chrome://extensions/shortcuts` if it clashes.

Keys: ↑/↓ walk (the selection follows; row 0 is the origin), → = New Origin
(drill into the highlighted cell), Backspace/← = Back, Enter = close and stay
where the walk left you, Esc/OK = return to the origin and close. The toolbar
icon opens Scopion too. Hidden sheets are walked like any other: Sheets unhides
a sheet when you jump into it, and Scopion hides it again when you close — that
re-hide is the only write it makes. Reads run under `spreadsheets.readonly`;
the `spreadsheets` edit scope is requested incrementally the first time a
re-hide is needed, and declining it leaves the sheets visible with a notice. The breadcrumb row under the list shows
every origin you drilled through; click one to go back there. An .xlsx opened
in Sheets is refused by the Sheets API: Scopion then
runs from the page alone — references and jumps work, values show as "—".

Tests: `node extension/build.js --lib-only && node --test extension/test/*.test.js`.

Web Store upload: `node extension/build.js --pack` writes
`dist/scopion-<version>.zip` holding only the files the extension runs
(no tests, build script, credentials or `key`). The store assigns its own key,
so the store item's ID may differ from the unpacked one: after the first upload
compare the Item ID with the OAuth client's and update the client if they
differ. `--pem extension/scopion.pem` puts the private key into the zip as
`key.pem`, which is the undocumented way to keep the ID — use it only if you
would rather hand Google the key than re-point the OAuth client. The manifest
strings come from `extension/_locales/{en,ja}` and the
store listing is entered per language from `docs/store/listing.en.md` /
`listing.ja.md`; the dashboard checklist is `docs/store/listing.md`. On install
(and on a toolbar click outside a spreadsheet) the extension opens
`onboarding.html`, localised the same way. The privacy policy and homepage the
store and Google's OAuth verification require are served from `docs/` by GitHub
Pages. Publishing additionally needs Google's verification of the sensitive
`spreadsheets` scope.

## What changed in the port, and why

| Excel original | Here | Reason |
|---|---|---|
| `Range.NavigateArrow` walks the dependency arrows | The formula is parsed and its references resolved | Apps Script exposes no dependency graph at all |
| Dependents come from the same arrow walk | Every sheet's formulas are scanned and matched against the target | Same reason; the elapsed time and formula count are reported in the sidebar |
| Active cell painted `ColorIndex 4`, old colour restored on exit | Nothing is painted; `setActiveRange` selects the cell | The selection outline already does this natively. Painting would write to the document, pollute undo, fight conditional formatting, fail on protected ranges, and disturb other editors — and it would stay behind if the sidebar died |
| Column of the target's colour | Kept, as a swatch read from `getBackgrounds()` | Input-vs-formula colour conventions are a real audit signal |
| Every sheet unhidden before the walk, restored afterwards | Hidden sheets are read in place | The Apps Script range API ignores sheet visibility. Only *jumping* to a hidden sheet needs it unhidden, and the sidebar then offers **Re-hide sheets** |
| `H` / `VH` / `EX` flags | `H` and `EX` | Google Sheets has no "very hidden" state, so `VH` could never occur |
| `FUNCTIONFINDER` scans `OFFSET(`, `INDEX(`, `SUM(`, `INDIRECT(` | `OFFSET`, `INDEX`, `INDIRECT` only | `SUM` does not produce a reference. Its arguments are ordinary precedents that the parser already returns, so scanning for it double-counted every `SUM` range |
| Settings stored as workbook defined names (`ACERangeNameSetting` …) | `DocumentProperties` | Same persistence, without adding names to the user's file |
| External workbook references | `IMPORTRANGE` rows, flagged `EX`, clicking opens the source file | Sheets has no other cross-file reference, and no script can move another file's selection |
| Blank target → `---BLANK CELL---`, list turns yellow | Same | A blank precedent in a financial model is usually a mistake, which is the whole point of the alert |

## Limits worth knowing

- **Dynamic references are resolved, not evaluated.** `OFFSET`/`INDEX`/`INDIRECT`
  arguments are worked out from literals, arithmetic, and the values of
  single-cell references. `OFFSET(A1,MATCH(...),0)` cannot be resolved without a
  formula engine, so it is listed explicitly as *unresolved* with the reason.
  It is never silently dropped — a missing precedent is worse than a visible gap.
  For the same reason an argument that merely *contains* a reference
  (`OFFSET(SOMEFN(A1),1,0)`) is reported unresolved rather than guessed at.
- **`INDIRECT(ref, FALSE)`** — the R1C1 form is not parsed; those rows come back
  unresolved.
- **`LET`/`LAMBDA` bindings** are treated as local to the whole formula rather
  than to their real scope. A workbook named range shadowed by a `LET` variable
  of the same name is therefore not reported. Erring this way drops a reference
  rather than inventing one.
- **A referenced range is one row, not one row per cell** — `=SUM(A1:A3)` lists
  `A1:A3`, matching the Excel original (which showed `--` for a multi-cell
  precedent). Conversely, every cell inside a referenced range does count as a
  dependent of that formula, which is also what Excel's tracing does.
- **Precedents are one level deep**, matching the original. Click through for depth.
- **Dependents are scanned live**, with no persistent index. One `getFormulas()`
  call per sheet, parsed in memory; sheets over 2,000,000 cells, and anything not
  reached within four minutes, are named in the sidebar rather than quietly
  omitted. If the reported scan times turn out to be slow on your models, that
  measurement is the argument for adding an index — building one first would have
  added edit triggers, generation counters, and a hidden index sheet for a cost
  nobody had checked.
- **Array formulas** are parsed like any other formula; a reference inside
  `ARRAYFORMULA`/`QUERY` is found, but the cells such a formula *spills* into are
  not treated as dependents.
- **References outside the sheet's grid** (a formula left pointing past the last
  row after a delete) are shown as blank and are not clickable, because
  `getRange()` throws on them.

## Tests

Two harnesses, both runnable without a Google account:

```bash
node run_tests.js   # 30 groups: formula parsing, reference maths, the arithmetic evaluator
node run_e2e.js     # aceAudit / aceJump against a fake spreadsheet
```

`runTests()` also runs from the Apps Script editor. `run_e2e.js` stubs
`SpreadsheetApp` — it proves the audit logic end to end, but it cannot prove that
Apps Script itself behaves as the stub does. The first run against a real
spreadsheet is still the real test.
