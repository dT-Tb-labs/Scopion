# Project Memory
最終更新: 2026-09-05

## 現在の状態
Scopion = formula auditor for Google Sheets, an ACE (Excel VBA) port. Two builds:
- Apps Script (frozen): bound script on test sheet 11GMudQJ1VmzTgscq4sY82ceb1lHywWKisegzQPQC3P0, macro Cmd+Option+Shift+1, Karabiner rule Cmd+Shift+A. Row/column neighbour flag shipped there (dac2d5a), not yet ported to the extension. Deploy = edit monaco in the editor tab, verify the checksum before saving (string transport corrupts characters).
- Chrome extension (`extension/`, MV3) = the product. main == feat/chrome-extension @ 532efc0. User-verified live 2026-09-05: Ctrl+Shift+A + toolbar icon, origin row first, ↑↓ walk with the selection following, range rows select the whole block, → drill, ←/Backspace back, Enter close-and-stay, Esc return-and-close, liquid-glass panel, formula-token ↔ row colour link, cross-sheet rows bold, smart placement (right → below → above), breadcrumbs in the box beside OK (click = back to that origin), hidden sheets walked by default and re-hidden on close via batchUpdate (OAuth scope is now `spreadsheets`, read/write; re-hide is the only write), DOM-only mode when the Sheets API refuses the file (.xlsx opened in Sheets: values "—", jumps work). Scorpius icon: extension/icons (`swift extension/icons/render.swift scorpio.svg extension/icons`). Store description ≤132 chars, aimed at financial modellers.
Spec: docs/superpowers/specs/2026-09-03-scopion-chrome-extension-design.md. Plan: docs/superpowers/plans/2026-09-03-scopion-chrome-extension.md.

## 直近の作業
Worktree .worktrees/chrome-extension (branch feat/chrome-extension) is KEPT on purpose: the user's Chrome loads the extension unpacked from `.worktrees/chrome-extension/extension` (id ihjnpijodijcbigekdihoamchmkgdgie, GCP project 839296398878; oauth.local.json is in both checkouts, scopion.pem only in the worktree — needed only to regenerate the key). Every code change needs chrome://extensions ↻ AND a reload of the sheet tab; a scope change re-prompts consent.
Verification: `node run_tests.js` (35/35), `node run_e2e.js`, `node extension/build.js --lib-only && node --test extension/test/*.test.js` (43; the `--test <dir>` form is broken on Node 26.5 here). Smoke checklist: tests/README.md. Automation trigger: `document.dispatchEvent(new CustomEvent('scopion:toggle'))`; debug trace in `#scopion-host` dataset.trace. build.test.js parks and restores a real manifest.json.
Sheets DOM facts (verified live): `#t-name-box` = address / range / covering named-range name; `#t-formula-bar-input` = formula; `.active-cell-border` ×4 = cell rect, `.selection-border` ×4 = range rect (hidden pieces are 0×0 at the page corner — skip); jump = name box + Enter (`'Sheet'!A1`, `B1:B3`, bare names; a hidden sheet gets unhidden); `#waffle-rich-text-editor`.focus() returns the keyboard to the grid. API enum: ErrorValue.type is DIVIDE_BY_ZERO. Sheets API refuses .xlsx-in-Sheets documents.
Test sheet sheets: Model, Inputs (constants, lookup table), Calc (50 formulas: cross-sheet, hidden, OFFSET/INDEX/INDIRECT, names, blanks, whole col/row, LET, ARRAYFORMULA, errors), Hidden Calc.
Known limits: MATCH exact-only; `=SUM(INDEX(..):INDEX(..))` reports endpoints; the Apps Script sidebar iframe is unreachable from automation (walk checked via tests/harness.html).

## 次のタスク / 未解決
- Chrome Web Store competitors (2026-09-05): SheetTrace (DOM-based precedents, free), SheetWhiz + Formula Explorer (arrow-key navigation, Goal Seek/What-If, YC), XLKeys, ShortieCuts. Scopion lacks Dependents (dropped on purpose) — a weakness for financial modellers; differentiators: origin row/range selection, token colour link, hidden re-hide, xlsx mode, breadcrumbs.
- Web Store publication needs Google's verification of the sensitive `spreadsheets` scope; port the neighbour flag to the extension (3×3 block fetch); consider Dependents via API scan.
- Parked: dom.js jump settle accepts any name-shaped name-box value (narrow race on cross-sheet jumps into named cells) — compare against namesCovering() if it ever bites.
- Scratch dirs to delete when the user agrees: `.superpowers/` (main) and `.worktrees/chrome-extension/.superpowers/` (git-ignored SDD ledgers).
