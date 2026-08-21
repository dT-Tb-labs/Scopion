# Project Memory
最終更新: 2026-08-21

## 現在の状態
Scopion = formula auditor for Google Sheets, ported from the Excel VBA add-in ACE.
Bound script on test sheet 11GMudQJ1VmzTgscq4sY82ceb1lHywWKisegzQPQC3P0,
Apps Script project 1UfOTEheizdN02zCw_4wuAGCFMh63foz_cbdPGwmizjVOHNRL_yYLELM6.
Cloud == local, verified by checksum (Formula+Trace+Code concatenated = 326576712 / 67612;
Sidebar.html = 845800809 / 32064). Deploy method: edit monaco in the editor tab, then
verify the checksum before saving — string transport silently corrupts characters.
Tests: `node run_tests.js` (31/31), `node run_e2e.js`. Walk smoke test: tests/README.md.

Launch: **Cmd+Option+Shift+1** (verified live). Menu Scopion > Open Scopion also works.
Esc = return to origin + clear highlight + close (verified in harness).

## 直近の作業
Measured on a generated 51,600-formula model (built and deleted twice via a temp
menu item; the generator is not in the repo):
- Latency is round trips, NOT bytes. getLastRow/getMaxRows are server calls.
  One audit: 52 calls to fetch 1,242 cells. Now 22 after caching sheet extent
  per audit (ValueCache.dims) and re-enabling the materializeRows batch.
- Spreadsheet first touch costs ~395ms; getNamedRanges only ~62ms. An earlier
  claim that buildNamedRangeMap was the 497ms cost was wrong.
- Sidebar wall-clock is 2.7-4.0s on that model, variance too wide to resolve
  further round-trip cuts. Perf work stopped there deliberately.

Fixed: dependents capped at 200 with the remainder reported; oversized range rows
print "(N cells)" instead of an unreadable exponent; formatNumber falls back to
exponent notation outside 1e-4..1e15.

## 次のタスク / 未解決
- User is testing the current build (2026-08-21 evening). Awaiting feedback.
- Proposed next features, not started: (1) flag hardcoded numbers inside formulas,
  (2) flag a formula that differs from its row neighbours, (3) breadcrumb trail,
  (4) number keys 1-9 to jump. (1) and (2) are the high-value ones.
- **Marketplace add-ons cannot have the shortcut.** Google docs, verbatim: "You
  cannot distribute macro definitions using a Sheets Google Workspace add-on."
  Only route to keep Cmd+Option+Shift+1 for distributed users is a companion
  Chrome extension that traps the key and clicks the add-on menu item. Needs the
  add-on to exist first.
- A simple onOpen CANNOT show a sidebar (runs unauthorized, showSidebar throws,
  completes in 0.7s silently). An installable open trigger can, but needs the
  script.scriptapp scope and re-authorization. Auto-open was built and reverted
  (commit 1e94041) — user wants the shortcut, not auto-open.
- karabiner-ace.json = right-Command double-tap -> Cmd+Option+Shift+1. Karabiner
  is NOT installed on this machine; untested. Modifier was Ctrl and was wrong —
  Sheets registers the macro as Cmd on macOS.
- Known limits: MATCH exact-only; =SUM(INDEX(..):INDEX(..)) reports both endpoints
  not the span; sidebar iframe input is unreachable from browser automation, so
  the walk must be checked via tests/harness.html.
