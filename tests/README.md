# Playwright harness

`node tests/build-harness.js` bakes `tests/harness.html`: the real `Sidebar.html`
with a `google.script.run` shim that executes the REAL server code
(`Formula.gs` + `Trace.gs` + `Code.gs`) against the same in-memory spreadsheet
mock `run_e2e.js` uses. Serve `tests/` over http (file:// is blocked by
Playwright) and drive it like the live sidebar.

Test hooks exposed on `window`:
- `__ss` — the mock spreadsheet (assert grid selection and cell backgrounds)
- `__rpcLog` — every RPC with its arguments, in order
- `__latency` — artificial round-trip delay in ms (default 250)

What it proves: state machine, keyboard walk, highlight paint/restore,
latest-wins jump queue, follow-mode polling. What it cannot prove: real Apps
Script focus behaviour — the one-hop-then-dead bug was only visible live.

## Walk smoke test

The keyboard walk is the product. Nothing in `run_e2e.js` covers the client
half of it, so it has to be driven by hand — run this after any change to
`Sidebar.html`, `materializeRows`, or `ValueCache`:

```bash
node tests/build-harness.js && (cd tests && python3 -m http.server 8731)
```

Open `http://localhost:8731/harness.html`, then in the console:

```js
const list = document.getElementById('list');
const ss = window.__ss;
const painted = () => { const p = []; for (const sh of ss.sheets)
  for (const r of Object.keys(sh.cells)) for (const c of Object.keys(sh.cells[r])) {
    const b = sh.cells[r][c].background;
    if (b && b.toLowerCase() === '#ccff90') p.push(sh.name + '!' + r + ',' + c); }
  return p; };
const where = () => ss.activeRange.sheet.name + '!' + ss.activeRange.getA1Notation();
const fire = (k) => list.dispatchEvent(
  new KeyboardEvent('keydown', { key: k, bubbles: true, cancelable: true }));

list.focus();
// then, one at a time, waiting ~2.5s after each:
//   fire('ArrowDown')  -> where() Hidden Calc!A1, painted() one cell
//   fire('ArrowDown')  -> where() Inputs!B3,      painted() ONE cell (previous restored)
//   fire('ArrowUp')    -> where() Hidden Calc!A1
//   fire('Escape')     -> where() Model!A2, painted() [], window.__closed === true
```

`painted()` returning more than one cell means a restore was missed — that is
the orphaned-highlight bug, and it is the thing this check exists to catch.

## Extension smoke test

Run on the test spreadsheet (Model sheet) after `node extension/build.js` and loading `extension/` unpacked; needs the OAuth client from README 'Chrome extension'.

```
[ ] Select B6 (=B1*(1+Growth)); Ctrl+Shift+A → strip appears below-right of B6 within ~1 s
[ ] List shows Model!B1 = 100 and Inputs-side Growth row; ORIGIN FORMULA shows the formula
[ ] ↓ moves the highlight to the second row AND the grid selection moves there; green overlay on the cell
[ ] ↑ returns; origin cell shows the dashed outline when standing on it
[ ] Enter on B1 → list re-audits from B1 ("Not a formula cell." with empty list), Back enabled
[ ] Backspace → back to B6 with the original list
[ ] Select B7 (=OFFSET('Hidden Calc'!A1,Shift,0)); Ctrl+Shift+A while open → New Origin re-audits
[ ] Hidden rows absent by default; Advanced → tick Include Hidden Sheets → Hidden Calc!A2 (H) appears
[ ] ↓ onto Hidden Calc!A2 → sheet switches, notice says Sheets unhid it
[ ] Esc → selection returns to B7, overlays gone, strip closed
[ ] Right-click the Hidden Calc tab → シートを非表示 (restore the test sheet)
[ ] Drag the strip by its title; Esc; reopen → same position
[ ] Click a cell elsewhere, then press ↓ in the strip → the walk first returns to its last cell (resync), then moves
```
