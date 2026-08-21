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
