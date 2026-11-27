/* mailai.uk: smooth wheel scrolling by linear interpolation (lerp), in the style of Lenis.
   Each wheel event moves a target; every frame the page moves 10% of the way there.
   Off for reduced motion and touch/coarse pointers. Keyboard, scrollbar and nested scroll areas stay native. */
(function () {
  var reduced = window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  var coarse = window.matchMedia('(pointer: coarse)').matches;
  if (reduced || coarse) return;

  var LERP = 0.07;         // share of the remaining distance covered per frame
  var HEADER = 68;         // sticky header height, so anchors land below it
  var root = document.documentElement;
  root.style.scrollBehavior = 'auto'; // we animate ourselves; native smooth would fight each frame

  var target = window.scrollY;
  var current = window.scrollY;
  var running = false;
  var ours = false;        // true while the scroll event comes from our own scrollTo

  var maxScroll = function () { return root.scrollHeight - window.innerHeight; };
  var clamp = function (v) { return Math.max(0, Math.min(maxScroll(), v)); };

  function frame() {
    current += (target - current) * LERP;
    if (Math.abs(target - current) < 0.5) current = target;
    ours = true;
    window.scrollTo(0, current);
    if (current !== target) requestAnimationFrame(frame);
    else running = false;
  }
  function start() {
    if (!running) { running = true; requestAnimationFrame(frame); }
  }

  // Let nested areas that can scroll in the wheel's direction handle it themselves.
  function canScroll(el, dy, dx) {
    for (; el && el !== document.body && el !== root; el = el.parentElement) {
      var s = getComputedStyle(el);
      if (dy && /(auto|scroll)/.test(s.overflowY) && el.scrollHeight > el.clientHeight) {
        if ((dy < 0 && el.scrollTop > 0) || (dy > 0 && el.scrollTop + el.clientHeight < el.scrollHeight)) return true;
      }
      if (dx && /(auto|scroll)/.test(s.overflowX) && el.scrollWidth > el.clientWidth) return true;
    }
    return false;
  }

  window.addEventListener('wheel', function (e) {
    if (e.ctrlKey || e.metaKey) return;                       // pinch / browser zoom
    if (Math.abs(e.deltaX) > Math.abs(e.deltaY)) return;      // sideways: native
    if (canScroll(e.target, e.deltaY, 0)) return;
    var dy = e.deltaY * (e.deltaMode === 1 ? 16 : e.deltaMode === 2 ? window.innerHeight : 1);
    e.preventDefault();
    if (!running) current = window.scrollY;
    target = clamp(target + dy);
    start();
  }, { passive: false });

  // Keyboard, scrollbar drag, find-in-page: follow wherever the page went.
  window.addEventListener('scroll', function () {
    if (ours) { ours = false; return; }
    if (!running) target = current = window.scrollY;
  }, { passive: true });

  window.addEventListener('resize', function () { target = clamp(target); });

  // In-page links glide too.
  document.addEventListener('click', function (e) {
    var a = e.target.closest && e.target.closest('a[href*="#"]');
    if (!a || a.target || e.ctrlKey || e.metaKey || e.shiftKey) return;
    var url = new URL(a.href, location.href);
    if (url.pathname !== location.pathname || !url.hash) return;
    var el = document.getElementById(decodeURIComponent(url.hash.slice(1)));
    if (!el) return;
    e.preventDefault();
    if (!running) current = window.scrollY;
    target = clamp(el.getBoundingClientRect().top + window.scrollY - HEADER);
    history.pushState(null, '', url.hash);
    start();
  });
})();
