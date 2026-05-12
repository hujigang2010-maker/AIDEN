/* Semir Global HQ — interactive behaviors */

(function () {
  'use strict';

  const $ = (sel, ctx = document) => ctx.querySelector(sel);
  const $$ = (sel, ctx = document) => Array.from(ctx.querySelectorAll(sel));

  /* ----------  Nav: scrolled state + mobile toggle  ---------- */
  const nav = $('#nav');
  const navMenu = $('#navMenu');
  const navLinks = $$('.nav__links a');

  const onScroll = () => {
    if (window.scrollY > 30) nav.classList.add('is-scrolled');
    else nav.classList.remove('is-scrolled');
  };
  window.addEventListener('scroll', onScroll, { passive: true });
  onScroll();

  if (navMenu) {
    navMenu.addEventListener('click', () => {
      nav.classList.toggle('is-mobile-open');
    });
  }
  navLinks.forEach((a) => a.addEventListener('click', () => nav.classList.remove('is-mobile-open')));

  /* ----------  Chapter Rail visibility + active chapter  ---------- */
  const rail = $('#chapterRail');
  const railLinks = $$('.chapter-rail a');
  const hero = $('#hero');

  const heroObserver = new IntersectionObserver(
    (entries) => {
      entries.forEach((entry) => {
        if (entry.isIntersecting) rail.classList.remove('is-visible');
        else rail.classList.add('is-visible');
      });
    },
    { threshold: 0.1 }
  );
  if (hero) heroObserver.observe(hero);

  /* Track active chapter for both nav and rail */
  const chapters = ['chapter-1', 'chapter-2', 'chapter-3', 'chapter-4']
    .map((id) => document.getElementById(id))
    .filter(Boolean);

  const setActiveChapter = (id) => {
    railLinks.forEach((l) => l.classList.toggle('is-active', l.dataset.target === id));
    navLinks.forEach((l) => {
      const href = l.getAttribute('href');
      l.classList.toggle('is-active', href === `#${id}`);
    });
  };

  const sectionObserver = new IntersectionObserver(
    (entries) => {
      entries.forEach((entry) => {
        if (entry.isIntersecting) setActiveChapter(entry.target.id);
      });
    },
    { rootMargin: '-40% 0px -55% 0px', threshold: 0 }
  );
  chapters.forEach((c) => sectionObserver.observe(c));

  // also observe end-of-chapter regions for cleaner activation
  const chapterAnchors = $$('section.chapter-divider, section.section');
  // (Above observer is sufficient; keep simple.)

  /* ----------  Reveal-on-scroll  ---------- */
  const revealTargets = $$(
    '.section__head, .strategy-card, .zone, .traffic-card, .stat, .brand-cloud, ' +
      '.totals__item, .building, .goal, .driver, .pyramid__row, .license-card, ' +
      '.service-card, .strategy-col, .big-pitch, .floors-table, .anchor, .event, .partners__col'
  );
  revealTargets.forEach((el) => el.classList.add('reveal'));

  const revealObserver = new IntersectionObserver(
    (entries) => {
      entries.forEach((entry, i) => {
        if (entry.isIntersecting) {
          // small stagger when many siblings come in together
          const delay = Math.min(i * 40, 200);
          setTimeout(() => entry.target.classList.add('is-visible'), delay);
          revealObserver.unobserve(entry.target);
        }
      });
    },
    { threshold: 0.12, rootMargin: '0px 0px -60px 0px' }
  );
  revealTargets.forEach((el) => revealObserver.observe(el));

  /* ----------  Animated counters  ---------- */
  const counters = $$('[data-count]');
  const animateCount = (el) => {
    const target = parseFloat(el.dataset.count);
    if (Number.isNaN(target)) return;
    const isFloat = !Number.isInteger(target);
    const duration = 1400;
    const start = performance.now();

    const tick = (now) => {
      const t = Math.min((now - start) / duration, 1);
      // ease-out cubic
      const eased = 1 - Math.pow(1 - t, 3);
      const val = target * eased;
      el.textContent = isFloat ? val.toFixed(1) : Math.round(val).toString();
      if (t < 1) requestAnimationFrame(tick);
      else el.textContent = isFloat ? target.toFixed(1) : Math.round(target).toString();
    };
    requestAnimationFrame(tick);
  };

  const counterObserver = new IntersectionObserver(
    (entries) => {
      entries.forEach((entry) => {
        if (entry.isIntersecting) {
          animateCount(entry.target);
          counterObserver.unobserve(entry.target);
        }
      });
    },
    { threshold: 0.5 }
  );
  counters.forEach((el) => counterObserver.observe(el));

  /* ----------  Subtle parallax on hero glows  ---------- */
  const glows = $$('.hero__glow');
  let parallaxFrame = null;
  window.addEventListener(
    'mousemove',
    (e) => {
      if (parallaxFrame) return;
      parallaxFrame = requestAnimationFrame(() => {
        const x = (e.clientX / window.innerWidth - 0.5) * 30;
        const y = (e.clientY / window.innerHeight - 0.5) * 30;
        glows.forEach((g, i) => {
          const k = (i + 1) * 0.6;
          g.style.transform = `translate(${x * k}px, ${y * k}px)`;
        });
        parallaxFrame = null;
      });
    },
    { passive: true }
  );
})();
