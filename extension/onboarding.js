/**
 * onboarding.js — fills the onboarding page from _locales (the HTML carries
 * the English text as its fallback) and wires the two things a plain link
 * cannot do from an extension page: open chrome://extensions/shortcuts, and
 * show the reload hint when background.js sent us here from a spreadsheet tab
 * whose content script predates the install.
 */
(function () {
  var i18n = typeof chrome !== 'undefined' && chrome.i18n;
  if (i18n) {
    document.documentElement.lang = (i18n.getUILanguage() || 'en').split('-')[0];
    var nodes = document.querySelectorAll('[data-i18n]');
    for (var i = 0; i < nodes.length; i++) {
      var text = i18n.getMessage(nodes[i].getAttribute('data-i18n'));
      if (text) nodes[i].textContent = text;
    }
  }
  if (location.hash === '#reload') document.getElementById('reload').hidden = false;
  document.getElementById('shortcuts').addEventListener('click', function () {
    chrome.tabs.create({ url: 'chrome://extensions/shortcuts' });
  });
})();
