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
