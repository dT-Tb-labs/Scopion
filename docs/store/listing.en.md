# Store listing — English

Plain text: the store's description field renders no markdown. `extension/test/listing.test.js` checks the lengths.

## Summary

Trace a formula's inputs in Google Sheets: Ctrl+Shift+A, walk the referenced cells with ↑↓, across sheets. For financial models.

## Description

Scopion shows you every cell a formula reads, then walks you through them one keystroke at a time.

Select a formula cell and press Ctrl+Shift+A (or click the toolbar icon). A compact panel opens next to the cell and lists its references: the sheet, the address, the current value, and — for a range — its total. Press ↓ and the spreadsheet selection jumps to the first reference; keep pressing and you follow the formula across sheets, into hidden sheets, through named ranges. Each reference in the formula is coloured to match its row, so you can read the formula and the list as one.

Built for financial models
• Origin row: the list starts with the cell you audited, so ↑ from the first input brings you home.
• Drill: → makes the highlighted cell the new origin; ← or Backspace goes back. Breadcrumbs show the trail you took, and clicking one returns there.
• Blank inputs are flagged loudly: a referenced cell that is empty is usually the error you were looking for.
• Range rows select the whole block and show its sum and count.
• IMPORTRANGE rows open the source spreadsheet.
• OFFSET, INDEX and INDIRECT are resolved from literals and single-cell values. What cannot be resolved is listed as unresolved with the reason — never silently dropped.
• Hidden sheets are walked like any other. Sheets unhides a sheet when you jump into it; Scopion hides it again when you close the panel (except on an Excel file, where the panel asks you to re-hide by hand).

Keys
↑ ↓ walk the references (the selection follows) · → new origin · ← / Backspace back · Enter close and stay · Esc / OK return to the start and close. Rebind the shortcut at chrome://extensions/shortcuts if another extension has it.

What it accesses, and the one thing it writes
Scopion reads the open spreadsheet through the Google Sheets API — sheet names, named ranges, and the cells the formula points at — only when you invoke it. The only change it ever makes is hiding again the sheets it had to unhide for the walk. It has no server: nothing is stored, logged or sent anywhere except to Google's API. Three display settings and the panel's last position live in your Chrome profile. Details in the privacy policy.

Good to know
• An Excel file opened in Sheets (.xlsx) has no API data: references and jumps work, values show as —, and hidden sheets you walked into stay visible until you re-hide them.
• Precedents are one level deep, like the Excel original; drill in for depth. Dependents are not traced.
• The panel is in English; the store page and the welcome page are also available in Japanese.

Scopion is a port of the formula auditor that generations of modellers used inside Excel, rebuilt for Google Sheets.
