---
layout: default
title: Privacy Policy
permalink: /privacy/
---

# Scopion Privacy Policy

_Last updated: 2026-09-05_

Scopion is a Chrome extension that shows, for the selected cell in a Google
Sheets spreadsheet, the cells its formula reads, and lets you jump between
them. This policy describes what data the extension touches and what happens
to it.

## What Scopion accesses

- **The spreadsheet open in your tab.** When you press the shortcut or the
  toolbar icon, Scopion reads the selected cell's address and formula from the
  Google Sheets page, and requests from the Google Sheets API the sheet names,
  named ranges and the values of the cells the formula references. Nothing is
  read until you invoke it, and only from the spreadsheet you invoked it on.
- **Your Google account, for authorisation only.** Scopion uses Chrome's
  identity API to obtain an OAuth token for the `spreadsheets.readonly` scope
  so that the Sheets API will answer. The first time it needs to hide a sheet
  again (below) it asks, separately, for the `spreadsheets` edit scope; you can
  decline and re-hide by hand. Scopion never sees your password and does not read
  your profile, email address or Drive file list.
- **Preferences.** Three display settings (show hidden sheets, show range
  names, show external references) are stored with `chrome.storage.sync`; the
  panel's last position is stored with `chrome.storage.local`. Both live in
  your Chrome profile.

## What Scopion writes

The only change Scopion makes to a spreadsheet is to **hide again the sheets it
had to unhide** so that you could jump into them, when you close the panel —
and only after you have granted the edit scope when first asked; if you
decline, the panel tells you which sheets to re-hide by hand.
It never edits cell contents, formatting or sharing. For an Excel file opened
in Sheets the Sheets API is unavailable, so Scopion cannot re-hide: the panel
says so and asks you to re-hide by hand.

## Where the data goes

Spreadsheet data travels only between your browser and Google's Sheets API
(`sheets.googleapis.com`), over HTTPS, and is held in the page's memory while
the panel is open. Scopion has **no server of its own**: it does not transmit,
store, log or analyse your spreadsheet data, your identity or your usage
anywhere else. There is no analytics, advertising or crash-reporting code.

## Limited Use disclosure

Scopion's use of information received from Google APIs adheres to the
[Chrome Web Store User Data Policy](https://developer.chrome.com/docs/webstore/program-policies/user-data-faq),
including the Limited Use requirements. In particular, Google user data is
used only to provide the feature described above, is never transferred to
third parties, is never used for advertising, and is never read by a human.

## Data retention and deletion

Nothing persists beyond the settings and panel position noted above. Removing
the extension deletes them. You can revoke Scopion's access to your Google
account at any time at <https://myaccount.google.com/permissions>.

## Changes and contact

Changes to this policy will be published at this address with a new date.
Questions: open an issue at <https://github.com/dT-Tb-labs/Scopion/issues>.
