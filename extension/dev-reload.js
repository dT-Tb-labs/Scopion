/**
 * dev-reload.js — development only, never shipped (not in build.js SHIP).
 * `node extension/build.js --watch` rewrites dev-reload.json after every
 * rebuild; the worker polls that stamp and reloads the extension when it
 * changes, then reloads the open Sheets tabs so the content script is fresh.
 * Only unpacked extensions serve their files live from disk, so the poll is
 * a no-op for the packed build even if this file were present.
 */
(function () {
  var STAMP = chrome.runtime.getURL('dev-reload.json');
  var PENDING = 'devReloadPending';
  var seen = null;

  chrome.storage.local.get(PENDING, function (v) {
    if (!v[PENDING]) return;
    chrome.storage.local.remove(PENDING);
    chrome.tabs.query({ url: 'https://docs.google.com/spreadsheets/*' }, function (tabs) {
      tabs.forEach(function (t) { chrome.tabs.reload(t.id); });
    });
  });

  function poll() {
    // An extension API call resets the worker's idle timer, so polling keeps it alive.
    chrome.runtime.getPlatformInfo(function () {});
    fetch(STAMP, { cache: 'no-store' }).then(function (r) { return r.text(); }).then(function (t) {
      if (seen === null) seen = t;
      else if (t !== seen) {
        chrome.storage.local.set({ devReloadPending: true }, function () { chrome.runtime.reload(); });
        return;
      }
      setTimeout(poll, 1000);
    }).catch(function () { /* no stamp: not a --watch build, stop polling */ });
  }
  poll();
})();
