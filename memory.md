# Project Memory
最終更新: 2026-09-05

## 現在の状態
Scopion = formula auditor for Google Sheets, an ACE (Excel VBA) port. Two builds:
- Apps Script (frozen): bound script on test sheet 11GMudQJ1VmzTgscq4sY82ceb1lHywWKisegzQPQC3P0, macro Cmd+Option+Shift+1, Karabiner rule Cmd+Shift+A. Row/column neighbour flag shipped there (dac2d5a), not yet ported to the extension. Deploy = edit monaco in the editor tab, verify the checksum before saving (string transport corrupts characters).
- Chrome extension (`extension/`, MV3): the product now. User-verified live 2026-09-05: shortcut Ctrl+Shift+A + toolbar icon, origin row first, ↑↓ walk with the selection following, range rows select the whole block, → drill, ←/Backspace back, Enter close-and-stay, Esc return-and-close, liquid-glass panel, formula-token ↔ row colour link, cross-sheet rows bold, smart placement (right → below → above), DOM-only mode when the Sheets API refuses the file (.xlsx opened in Sheets: values "—", jumps work). Scorpius icon: extension/icons (`swift extension/icons/render.swift scorpio.svg extension/icons`).
Spec: docs/superpowers/specs/2026-09-03-scopion-chrome-extension-design.md. Plan: docs/superpowers/plans/2026-09-03-scopion-chrome-extension.md.

## 直近の作業
Branch feat/chrome-extension (worktree .worktrees/chrome-extension): all 10 plan tasks, final review, fix wave, then live smoke fixes. The user's Chrome loads the extension unpacked from `.worktrees/chrome-extension/extension` (id ihjnpijodijcbigekdihoamchmkgdgie, GCP project 839296398878; oauth.local.json + scopion.pem are git-ignored and live only there). Every code change needs chrome://extensions ↻ AND a reload of the sheet tab (content scripts do not re-inject).
Verification: `node run_tests.js` (35/35), `node run_e2e.js`, `node extension/build.js --lib-only && node --test extension/test/*.test.js` (42; the `--test <dir>` form is broken on Node 26.5 here). Smoke checklist: tests/README.md. Automation trigger: `document.dispatchEvent(new CustomEvent('scopion:toggle'))`; debug trace in `#scopion-host` dataset.trace.
Sheets DOM facts (verified live): `#t-name-box` value = address / range / covering named-range name; `#t-formula-bar-input` = formula; `.active-cell-border` ×4 = cell rect, `.selection-border` ×4 = range rect (hidden pieces are 0×0 at the page corner — skip); jump = name box + Enter (`'Sheet'!A1`, `B1:B3`, bare names work; a hidden sheet gets unhidden); `#waffle-rich-text-editor`.focus() returns the keyboard to the grid. API enum: ErrorValue.type is DIVIDE_BY_ZERO.
Test sheet sheets: Model, Inputs (constants, lookup table), Calc (50 formulas: cross-sheet, hidden, OFFSET/INDEX/INDIRECT, names, blanks, whole col/row, LET, ARRAYFORMULA, errors), Hidden Calc.
Known limits: MATCH exact-only; `=SUM(INDEX(..):INDEX(..))` reports endpoints; the Apps Script sidebar iframe is unreachable from automation (walk checked via tests/harness.html).

## 次のタスク / 未解決
- Merge feat/chrome-extension into main; keep the worktree while Chrome points at it, or copy oauth.local.json + scopion.pem into `Scopion/extension`, build there, and load unpacked from that path (same id thanks to the pinned key).
- Web Store publication needs Google's sensitive-scope verification; re-hiding sheets needs the `spreadsheets` write scope; port the neighbour flag to the extension (3×3 block fetch).
- Parked: dom.js jump settle accepts any name-shaped name-box value (narrow race on cross-sheet jumps into named cells) — compare against namesCovering() if it ever bites.
- Competitors: SheetTrace (shortcuts), SheetWhiz (What-If/Goal Seek/Auto-Color), Formula Tracer Sidebar. Neighbour diff is the differentiator none has.
