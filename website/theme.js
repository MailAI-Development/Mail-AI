/* Mail AI — shared site behaviour: theme toggle, mobile nav, scroll reveal.
   The initial theme is set by a tiny inline snippet in each page's <head> to
   avoid a flash; this file wires up the toggle and the rest after load. */
(function () {
  var STORAGE_KEY = 'mailai-theme';

  function currentTheme() {
    return document.documentElement.getAttribute('data-theme') === 'dark' ? 'dark' : 'light';
  }

  function applyTheme(theme) {
    if (theme === 'dark') {
      document.documentElement.setAttribute('data-theme', 'dark');
    } else {
      document.documentElement.removeAttribute('data-theme');
    }
    try { localStorage.setItem(STORAGE_KEY, theme); } catch (e) {}
  }

  function init() {
    /* Theme toggle button(s) */
    document.querySelectorAll('.theme-toggle').forEach(function (btn) {
      btn.addEventListener('click', function () {
        applyTheme(currentTheme() === 'dark' ? 'light' : 'dark');
      });
    });

    /* Mobile nav */
    var navToggle = document.getElementById('navToggle');
    var navLinks = document.getElementById('navLinks');
    if (navToggle && navLinks) {
      navToggle.addEventListener('click', function () {
        navToggle.classList.toggle('active');
        navLinks.classList.toggle('open');
      });
      navLinks.querySelectorAll('a').forEach(function (a) {
        a.addEventListener('click', function () {
          navToggle.classList.remove('active');
          navLinks.classList.remove('open');
        });
      });
    }

    /* Scroll reveal */
    var revealEls = document.querySelectorAll('.reveal, .feature');
    if (revealEls.length && 'IntersectionObserver' in window) {
      var obs = new IntersectionObserver(function (entries) {
        entries.forEach(function (e) {
          if (e.isIntersecting) e.target.classList.add('visible');
          else e.target.classList.remove('visible');
        });
      }, { threshold: 0.12 });
      revealEls.forEach(function (el) { obs.observe(el); });
    }

    /* Subtle grid parallax (no smoothing delay) */
    var ticking = false;
    window.addEventListener('scroll', function () {
      if (!ticking) {
        requestAnimationFrame(function () {
          document.body.style.setProperty('--grid-shift', (-window.scrollY * 0.08) + 'px');
          ticking = false;
        });
        ticking = true;
      }
    }, { passive: true });
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
