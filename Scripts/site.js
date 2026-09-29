/* Shared page script: light/dark theme switcher plus small accessibility helpers.
   Load in <head> (not deferred) so the saved theme is applied before first paint. */
(function () {
  var root = document.documentElement;
  var KEY = 'theme';
  var media = window.matchMedia ? window.matchMedia('(prefers-color-scheme: dark)') : null;

  function saved() {
    try { return localStorage.getItem(KEY); } catch (e) { return null; }
  }

  function preferred() {
    var s = saved();
    if (s === 'dark' || s === 'light') return s;
    return media && media.matches ? 'dark' : 'light';
  }

  function apply(theme) {
    root.setAttribute('data-theme', theme);
    var btn = document.getElementById('theme-toggle');
    if (btn) {
      var dark = theme === 'dark';
      btn.setAttribute('aria-pressed', dark ? 'true' : 'false');
      btn.setAttribute('aria-label', dark ? 'Switch to light mode' : 'Switch to dark mode');
      btn.title = dark ? 'Light mode' : 'Dark mode';
    }
  }

  apply(preferred());

  // Follow the operating system setting until the visitor picks a theme themselves.
  if (media) {
    var onChange = function () { if (!saved()) apply(preferred()); };
    if (media.addEventListener) media.addEventListener('change', onChange);
    else if (media.addListener) media.addListener(onChange);
  }

  // Always print in light mode (resume and cover letter especially).
  var beforePrint = null;
  window.addEventListener('beforeprint', function () {
    beforePrint = root.getAttribute('data-theme');
    root.setAttribute('data-theme', 'light');
  });
  window.addEventListener('afterprint', function () {
    if (beforePrint) root.setAttribute('data-theme', beforePrint);
  });

  var ICONS =
    '<svg class="icon-moon" viewBox="0 0 24 24" width="20" height="20" aria-hidden="true" focusable="false">' +
      '<path fill="currentColor" d="M21 12.8A9 9 0 1 1 11.2 3a7 7 0 0 0 9.8 9.8z"/></svg>' +
    '<svg class="icon-sun" viewBox="0 0 24 24" width="20" height="20" aria-hidden="true" focusable="false">' +
      '<circle cx="12" cy="12" r="4.5" fill="currentColor"/>' +
      '<g stroke="currentColor" stroke-width="2" stroke-linecap="round">' +
      '<path d="M12 1.5v2.5M12 20v2.5M1.5 12h2.5M20 12h2.5M4.6 4.6l1.8 1.8M17.6 17.6l1.8 1.8M4.6 19.4l1.8-1.8M17.6 6.4l1.8-1.8"/></g></svg>';

  function addButton() {
    if (document.getElementById('theme-toggle')) return;
    var btn = document.createElement('button');
    btn.id = 'theme-toggle';
    btn.type = 'button';
    btn.className = 'theme-toggle';
    btn.innerHTML = ICONS;
    btn.addEventListener('click', function () {
      var next = root.getAttribute('data-theme') === 'dark' ? 'light' : 'dark';
      try { localStorage.setItem(KEY, next); } catch (e) {}
      apply(next);
    });
    document.body.appendChild(btn);
    apply(root.getAttribute('data-theme') || preferred());
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', addButton);
  else addButton();
})();

/* Make click-only controls (e.g. <div onclick>) reachable and usable from the keyboard. */
(function () {
  function enhance() {
    // Icon-only buttons: use their tooltip as the accessible name.
    var btns = document.querySelectorAll('button[title]:not([aria-label])');
    for (var j = 0; j < btns.length; j++) {
      if (!btns[j].textContent.trim()) btns[j].setAttribute('aria-label', btns[j].title);
    }
    var els = document.querySelectorAll('[onclick]');
    for (var i = 0; i < els.length; i++) {
      var el = els[i];
      var tag = el.tagName;
      if (tag === 'A' || tag === 'BUTTON' || tag === 'INPUT' || tag === 'SELECT' ||
          tag === 'TEXTAREA' || tag === 'SUMMARY' || tag === 'LABEL') continue;
      if (el.classList.contains('sidebar-overlay')) continue;
      if (!el.hasAttribute('role')) el.setAttribute('role', 'button');
      if (!el.hasAttribute('tabindex')) el.tabIndex = 0;
      el.addEventListener('keydown', function (e) {
        if (e.target !== this) return;
        if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); this.click(); }
      });
    }
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', enhance);
  else enhance();
})();
