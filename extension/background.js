/**
 * background.js — the service worker does two things the content script
 * cannot: receive the keyboard command, and call the Sheets API with an OAuth
 * token (chrome.identity is not available to content scripts). It holds no
 * state; every request carries its spreadsheet id.
 */
// Dev-only self-reload (see dev-reload.js); the file is absent from the packed build.
try { importScripts('dev-reload.js'); } catch (e) { /* not a dev build */ }

var SHEETS_API = 'https://sheets.googleapis.com/v4/spreadsheets/';
/**
 * Reads use the manifest's default scope (spreadsheets.readonly), so the first
 * consent says "view" only. The one write — re-hiding sheets — asks for the
 * edit scope incrementally, the first time it is actually needed.
 */
var WRITE_SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];
var META_FIELDS = 'namedRanges,sheets.properties(sheetId,title,hidden,gridProperties(rowCount,columnCount))';
var GRID_FIELDS = 'sheets(properties.sheetId,data(startRow,startColumn,rowData.values(effectiveValue,formattedValue,userEnteredValue)))';

function metaUrl(spreadsheetId) {
  return SHEETS_API + encodeURIComponent(spreadsheetId) + '?fields=' + encodeURIComponent(META_FIELDS);
}
function gridUrl(spreadsheetId, ranges) {
  return SHEETS_API + encodeURIComponent(spreadsheetId) + '?includeGridData=true&fields=' + encodeURIComponent(GRID_FIELDS) +
    ranges.map(function (r) { return '&ranges=' + encodeURIComponent(r); }).join('');
}

function rehideUrl(spreadsheetId) { return SHEETS_API + encodeURIComponent(spreadsheetId) + ':batchUpdate'; }
/** The only write Scopion ever makes: hide again the sheets it had to unhide to walk into. */
function rehideBody(sheetIds) {
  return { requests: sheetIds.map(function (id) {
    return { updateSheetProperties: { properties: { sheetId: id, hidden: true }, fields: 'hidden' } };
  }) };
}

/**
 * Where a toolbar click goes when no content script answers: the onboarding
 * page. tab.url is populated only for hosts the extension has permission for,
 * so a spreadsheet URL here means "Sheets tab opened before the install, not
 * yet reloaded" and the page shows its reload hint; undefined means any other site.
 */
function onboardingUrlFor(tab) {
  var url = tab && tab.url ? String(tab.url) : '';
  // A document URL only (/d/<id>): the Sheets home page also matches the content-script
  // pattern but has nothing to reload into.
  return 'onboarding.html' + (/^https:\/\/docs\.google\.com\/spreadsheets\/(?:u\/\d+\/)?d\//.test(url) ? '#reload' : '');
}

/**
 * Authorised fetch with a cached token; on 401 drop the token and retry once.
 * `interactive` (default true) allows the Google consent window when no token
 * is cached; the content script passes false after the user has declined once,
 * so a declined consent is not re-opened on every panel open. `scopes`
 * (optional) overrides the manifest scopes for this call — WRITE_SCOPES for
 * the re-hide.
 */
function apiCall(url, init, deps, interactive, scopes) {
  var mayPrompt = interactive !== false;
  function call(token) {
    var headers = Object.assign({ Authorization: 'Bearer ' + token }, (init && init.headers) || {});
    return deps.fetch(url, Object.assign({}, init || {}, { headers: headers }));
  }
  function prompt(err) {
    if (!mayPrompt) throw new Error('not signed in: ' + (err && err.message ? err.message : err));
    return deps.getToken(true, scopes);
  }
  return deps.getToken(false, scopes).catch(prompt)
    .then(function (token) {
      return call(token).then(function (res) {
        if (res.status !== 401) return res;
        return deps.removeToken(token).then(function () { return prompt(new Error('token rejected')); }).then(call);
      });
    })
    .then(function (res) {
      if (res.ok) return res.json();
      return res.text().then(function (body) { throw new Error('Sheets API ' + res.status + ': ' + String(body).slice(0, 200)); });
    });
}
function apiGet(url, deps, interactive) { return apiCall(url, null, deps, interactive); }
function apiPost(url, body, deps, interactive, scopes) {
  return apiCall(url, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) }, deps, interactive, scopes);
}

if (typeof chrome !== 'undefined' && chrome.runtime && chrome.runtime.onMessage) {
  var chromeDeps = {
    fetch: function (u, o) { return fetch(u, o); },
    getToken: function (interactive, scopes) {
      var details = { interactive: interactive };
      if (scopes) details.scopes = scopes; // incremental consent for the write
      return new Promise(function (resolve, reject) {
        chrome.identity.getAuthToken(details, function (result) {
          var token = result && typeof result === 'object' ? result.token : result;
          if (chrome.runtime.lastError || !token) reject(new Error(chrome.runtime.lastError ? chrome.runtime.lastError.message : 'no token'));
          else resolve(token);
        });
      });
    },
    removeToken: function (token) {
      return new Promise(function (resolve) { chrome.identity.removeCachedAuthToken({ token: token }, resolve); });
    }
  };

  chrome.runtime.onMessage.addListener(function (msg, sender, sendResponse) {
    if (!msg || !msg.type) return false;
    var p, interactive = msg.interactive !== false;
    if (msg.type === 'scopion:meta') p = apiGet(metaUrl(msg.spreadsheetId), chromeDeps, interactive);
    else if (msg.type === 'scopion:rehide') {
      if (!msg.sheetIds || !msg.sheetIds.length) { sendResponse({ ok: true, data: null }); return true; }
      p = apiPost(rehideUrl(msg.spreadsheetId), rehideBody(msg.sheetIds), chromeDeps, interactive, WRITE_SCOPES);
    }
    else if (msg.type === 'scopion:grid') {
      // An empty ranges list means "the whole spreadsheet" to the Sheets API,
      // not "nothing" — never send that request.
      if (!msg.ranges || !msg.ranges.length) { sendResponse({ ok: false, error: 'no ranges' }); return true; }
      p = apiGet(gridUrl(msg.spreadsheetId, msg.ranges), chromeDeps, interactive);
    }
    else return false;
    p.then(function (data) { sendResponse({ ok: true, data: data }); },
           function (e) { sendResponse({ ok: false, error: String(e && e.message || e) }); });
    return true; // keep the channel open for the async response
  });

  // The shortcut and the toolbar icon are two doors to the same room: a user
  // whose key is taken by another extension still has the icon. A click that
  // reaches no content script opens the onboarding page instead of doing
  // nothing; the shortcut stays silent there (an accidental chord on another
  // site should not open a tab).
  function openOnboarding(tab) {
    chrome.tabs.create({ url: chrome.runtime.getURL(onboardingUrlFor(tab)) });
  }
  // "No content script" is decided by the reply, not by lastError alone: the
  // content script answers {ok:true} synchronously, so an absent reply means
  // nobody was listening (lastError is read so Chrome does not log it unchecked).
  function sendToggle(tab, onNoReceiver) {
    if (!tab || tab.id === undefined) return;
    chrome.tabs.sendMessage(tab.id, { type: 'scopion:toggle' }, function (reply) {
      void chrome.runtime.lastError;
      if (!reply && onNoReceiver) onNoReceiver(tab);
    });
  }
  chrome.commands.onCommand.addListener(function (command, tab) {
    if (command === 'toggle-scopion') sendToggle(tab, null);
  });
  chrome.action.onClicked.addListener(function (tab) { sendToggle(tab, openOnboarding); });
  chrome.runtime.onInstalled.addListener(function (details) {
    if (details && details.reason === 'install') openOnboarding(null);
  });
}
