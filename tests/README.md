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
