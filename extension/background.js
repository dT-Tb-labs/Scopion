/**
 * background.js — the service worker does two things the content script
 * cannot: receive the keyboard command, and call the Sheets API with an OAuth
 * token (chrome.identity is not available to content scripts). It holds no
 * state; every request carries its spreadsheet id.
 */
var SHEETS_API = 'https://sheets.googleapis.com/v4/spreadsheets/';
var META_FIELDS = 'namedRanges,sheets.properties(sheetId,title,hidden,gridProperties(rowCount,columnCount))';
var GRID_FIELDS = 'sheets(properties.sheetId,data(startRow,startColumn,rowData.values(effectiveValue,formattedValue,userEnteredValue)))';

function metaUrl(spreadsheetId) {
  return SHEETS_API + encodeURIComponent(spreadsheetId) + '?fields=' + encodeURIComponent(META_FIELDS);
}
function gridUrl(spreadsheetId, ranges) {
  return SHEETS_API + encodeURIComponent(spreadsheetId) + '?includeGridData=true&fields=' + encodeURIComponent(GRID_FIELDS) +
    ranges.map(function (r) { return '&ranges=' + encodeURIComponent(r); }).join('');
}

/** GET with a cached token; on 401 drop the token and retry once interactively. */
function apiGet(url, deps) {
  function call(token) { return deps.fetch(url, { headers: { Authorization: 'Bearer ' + token } }); }
  return deps.getToken(false).catch(function () { return deps.getToken(true); })
    .then(function (token) {
      return call(token).then(function (res) {
        if (res.status !== 401) return res;
        return deps.removeToken(token).then(function () { return deps.getToken(true); }).then(call);
      });
    })
    .then(function (res) {
      if (res.ok) return res.json();
      return res.text().then(function (body) { throw new Error('Sheets API ' + res.status + ': ' + String(body).slice(0, 200)); });
    });
}

if (typeof chrome !== 'undefined' && chrome.runtime && chrome.runtime.onMessage) {
  var chromeDeps = {
    fetch: function (u, o) { return fetch(u, o); },
    getToken: function (interactive) {
      return new Promise(function (resolve, reject) {
        chrome.identity.getAuthToken({ interactive: interactive }, function (result) {
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
    var p;
    if (msg.type === 'scopion:meta') p = apiGet(metaUrl(msg.spreadsheetId), chromeDeps);
    else if (msg.type === 'scopion:grid') p = apiGet(gridUrl(msg.spreadsheetId, msg.ranges || []), chromeDeps);
    else return false;
    p.then(function (data) { sendResponse({ ok: true, data: data }); },
           function (e) { sendResponse({ ok: false, error: String(e && e.message || e) }); });
    return true; // keep the channel open for the async response
  });

  chrome.commands.onCommand.addListener(function (command, tab) {
    if (command === 'toggle-scopion' && tab && tab.id !== undefined) {
      chrome.tabs.sendMessage(tab.id, { type: 'scopion:toggle' }, function () { void chrome.runtime.lastError; });
    }
  });
}
