# Project Memory
最終更新: 2026-09-04

## 現在の状態
Scopion = formula auditor for Google Sheets. Apps Script build frozen: bound script on test sheet 11GMudQJ1VmzTgscq4sY82ceb1lHywWKisegzQPQC3P0, launched via Cmd+Option+Shift+1 macro (Karabiner rule `Cmd+Shift+A`). Deploy: edit monaco in the Apps Script editor tab, then verify the checksum before saving — string transport silently corrupts characters. Row/column neighbour flag shipped (commit dac2d5a), not yet ported to extension. Chrome extension (`extension/` on branch feat/chrome-extension) spec: `docs/superpowers/specs/2026-09-03-scopion-chrome-extension-design.md`, plan: `docs/superpowers/plans/2026-09-03-scopion-chrome-extension.md`.

## 直近の作業
Tasks 1–9 complete and reviewed. Task 10 (smoke test + docs) in progress: checklist appended to tests/README.md, live test pending user's GCP OAuth client (README steps 1–3). Verification: `node run_tests.js` (35/35), `node run_e2e.js`, `node extension/build.js --lib-only && node --test extension/test/*.test.js` (38 tests). Sheets DOM facts verified live 2026-09-03: name-box (#t-name-box) value = address or named-range name, formula-bar (#t-formula-bar-input) textContent = formula, `.active-cell-border` ×4 union = cell rect, jump via name box + Enter (`'Sheet'!A1` works), hidden-sheet jump unhides. API enum: `ErrorValue.type` is `DIVIDE_BY_ZERO`, not `DIV_0`.

## 次のタスク / 未解決
- Web Store publication needs Google sensitive-scope verification; re-hide needs `spreadsheets` write scope.
- Port row/column neighbour flag to extension (3×3 block fetch).
- Default shortcut is `Ctrl+Shift+A` on every platform, including Mac; `Cmd+Shift+A` is avoided because it is Chrome's own tab-search shortcut on macOS.
- Known limits: MATCH is exact-match only; `=SUM(INDEX(..):INDEX(..))` reports both endpoints, not the span between them; the sidebar iframe is unreachable from browser automation, so the walk is checked via tests/harness.html.
- Competitors: SheetTrace (shortcuts), SheetWhiz (What-If/Goal Seek), Formula Tracer Sidebar (Marketplace). Neighbour diff is the differentiator.
