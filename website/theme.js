/* mailai.uk: theme toggle (light by default, shared with the app via the same key) and the pricing toggle.
   Each page's <head> sets the stored theme before paint to avoid a flash. */
(function () {
  var KEY = 'mailai-theme';
  var root = document.documentElement;
  var current = function () { return root.getAttribute('data-theme') === 'dark' ? 'dark' : 'light'; };

  var SUN = '<circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4"/>';
  var MOON = '<path d="M20 14.5A8 8 0 0 1 9.5 4a8 8 0 1 0 10.5 10.5Z"/>';

  function paintToggles() {
    document.querySelectorAll('.theme-toggle').forEach(function (btn) {
      var toDark = current() === 'light';
      btn.setAttribute('aria-label', toDark ? 'Switch to dark theme' : 'Switch to light theme');
      btn.innerHTML = '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' + (toDark ? MOON : SUN) + '</svg>';
    });
  }

  function setTheme(t) {
    root.setAttribute('data-theme', t);
    try { localStorage.setItem(KEY, t); } catch (e) {}
    paintToggles();
  }

  document.addEventListener('DOMContentLoaded', function () {
    paintToggles();
    document.querySelectorAll('.theme-toggle').forEach(function (btn) {
      btn.addEventListener('click', function () { setTheme(current() === 'light' ? 'dark' : 'light'); });
    });

    /* Pricing: monthly / annual */
    var buttons = document.querySelectorAll('[data-billing]');
    function showBilling(period) {
      buttons.forEach(function (b) { b.setAttribute('aria-pressed', String(b.getAttribute('data-billing') === period)); });
      document.querySelectorAll('[data-monthly]').forEach(function (el) {
        el.textContent = el.getAttribute(period === 'annual' ? 'data-annual' : 'data-monthly');
      });
    }
    buttons.forEach(function (b) { b.addEventListener('click', function () { showBilling(b.getAttribute('data-billing')); }); });
    if (buttons.length) showBilling('annual');
  });
})();
