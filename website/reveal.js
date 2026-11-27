/* mailai.uk: fade-and-rise reveal, once per element. The hero plays on open; the rest as it scrolls into view.
   Items in the same group (steps, plans, FAQ rows…) arrive one after another. The <head> adds html.reveal only
   when motion is allowed, and removes it after 3s if this script never runs, so content can't stay hidden. */
(function () {
  window.__revealReady = true;
  var root = document.documentElement;
  if (!root.classList.contains('reveal')) return;

  var SELECTOR = '.hero h1, .hero .lede, .hero .ctas, .hero .facts, .preview, .section-head, .steps li, ' +
                 '.split .title, .points li, .billing, .plan, .pricing-note, .faq details, .cta-band, .doc > *';
  var STAGGER = 90;   // ms between items in the same group
  var MAX_DELAY = 450;

  var items = Array.prototype.slice.call(document.querySelectorAll(SELECTOR));
  // Stagger by position among revealed siblings (the three steps, the plan cards, …).
  items.forEach(function (el) {
    var siblings = Array.prototype.filter.call(el.parentElement.children, function (c) { return items.indexOf(c) !== -1; });
    var i = siblings.indexOf(el);
    el.style.setProperty('--reveal-delay', Math.min(i * STAGGER, MAX_DELAY) + 'ms');
  });

  var show = function (el) { el.classList.add('in'); };
  if (!('IntersectionObserver' in window)) { items.forEach(show); return; }

  var io = new IntersectionObserver(function (entries) {
    entries.forEach(function (e) {
      if (e.isIntersecting) { show(e.target); io.unobserve(e.target); }
    });
  }, { rootMargin: '0px 0px -8% 0px', threshold: 0.12 });

  items.forEach(function (el) { io.observe(el); });
})();
