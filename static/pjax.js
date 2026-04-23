/**
 * pjax.js
 * ───────
 * Vanilla PushState + Ajax router.
 *
 * Loaded with `defer` in <head> so document.body exists at execution time.
 * Event listeners are registered on `document` and `window` (not body nodes)
 * so they survive body-innerHTML swaps on every PJAX navigation.
 *
 * What it does
 * ────────────
 * 1. Intercepts <a> clicks that point to the same origin.
 * 2. Fetches the target URL, parses the response with DOMParser.
 * 3. Syncs <head> inline styles and the CSRF meta tag from the new page.
 * 4. Swaps document.body.innerHTML with the new page's body content.
 * 5. Re-executes inline <script> tags found in the new body.
 * 6. Updates document.title and pushes a history entry.
 * 7. Handles the Back / Forward buttons via popstate.
 *
 * Opt-out
 * ───────
 * Add  data-pjax="false"  to any <a> that must trigger a full navigation
 * (e.g., links that download files, external links, etc.)
 *
 * Worker persistence across navigations
 * ──────────────────────────────────────
 * window._parserWorker is created once and reused.  After each body swap the
 * worker's onmessage is updated to reference the new grid instances via
 * window._updateWorkerGridMap(), which is called by the editor's own scripts.
 */

/* ── Guard: run only once per true page load ────────────────────────────── */
if (window._pjaxInit) {
  // Script was somehow re-executed (shouldn't happen when loaded in <head>).
  // Bail so we don't double-register event listeners.
  throw new Error('pjax already initialised');
}
window._pjaxInit = true;

(function () {
  'use strict';

  /* ── Skip PJAX for file-download URLs ────────────────────────────────── */
  const SKIP_EXT = /\.(xlsx|xls|pdf|zip|csv|docx)(\?|$)/i;

  /* ── Progress bar ────────────────────────────────────────────────────── */
  let _bar = null;

  function _createBar() {
    if (document.getElementById('pjax-bar')) return;
    _bar = document.createElement('div');
    _bar.id = 'pjax-bar';
    document.body.appendChild(_bar);
  }

  function _barStart() {
    if (!_bar) return;
    _bar.style.opacity = '1';
    _bar.style.width   = '35%';
  }

  function _barProgress() {
    if (!_bar) return;
    _bar.style.width = '70%';
  }

  function _barFinish() {
    if (!_bar) return;
    _bar.style.width = '100%';
    setTimeout(() => {
      if (_bar) { _bar.style.opacity = '0'; _bar.style.width = '0'; }
    }, 280);
  }

  function _barError() {
    if (!_bar) return;
    _bar.style.background = 'var(--red, #ff5f5f)';
    _barFinish();
    setTimeout(() => {
      if (_bar) _bar.style.background = '';
    }, 600);
  }

  /* ── Head synchronisation ────────────────────────────────────────────── */

  /**
   * Replace all inline <style data-pjax> tags in <head> with the new page's
   * inline styles, and update the CSRF meta tag if present.
   */
  function _syncHead(newDoc) {
    // ── Inline styles ──────────────────────────────────────────────────────
    document.querySelectorAll('style[data-pjax]').forEach(s => s.remove());
    newDoc.querySelectorAll('head style').forEach(oldStyle => {
      const s           = document.createElement('style');
      s.setAttribute('data-pjax', '');
      s.textContent     = oldStyle.textContent;
      document.head.appendChild(s);
    });

    // ── CSRF meta (used by index.html's getCsrf()) ─────────────────────────
    const newMeta = newDoc.querySelector('meta[name="csrf-token"]');
    if (newMeta) {
      let oldMeta = document.querySelector('meta[name="csrf-token"]');
      if (oldMeta) {
        oldMeta.setAttribute('content', newMeta.getAttribute('content'));
      } else {
        document.head.appendChild(newMeta.cloneNode(true));
      }
    }
  }

  /* ── Body swap ───────────────────────────────────────────────────────── */

  /**
   * Replace body content and re-execute every inline <script>.
   * External <script src="..."> tags that are already loaded (pjax.js,
   * style.css companion scripts, etc.) are intentionally skipped to
   * avoid double-registration of global event handlers.
   */
  function _swapBody(newDoc) {
    // Collect external src values already loaded in this session.
    const alreadyLoaded = new Set(
      Array.from(document.querySelectorAll('script[src]'))
        .map(s => s.getAttribute('src'))
    );

    document.body.innerHTML = newDoc.body.innerHTML;

    // Re-append progress bar (it lived in old body, now destroyed).
    if (_bar) {
      _bar.style.width   = '0';
      _bar.style.opacity = '1';
      document.body.appendChild(_bar);
    }

    // Re-execute scripts — innerHTML does NOT run them automatically.
    document.body.querySelectorAll('script').forEach(stale => {
      const src = stale.getAttribute('src');

      // Skip already-loaded external scripts to prevent double-init.
      if (src && alreadyLoaded.has(src)) {
        stale.remove();
        return;
      }

      const fresh = document.createElement('script');
      Array.from(stale.attributes).forEach(a =>
        fresh.setAttribute(a.name, a.value)
      );
      fresh.textContent = stale.textContent;
      stale.parentNode.replaceChild(fresh, stale);
    });
  }

  /* ── Core navigate() ─────────────────────────────────────────────────── */

  let _navigating = false;   // Prevent concurrent navigations

  async function navigate(url, pushState = true) {
    if (_navigating) return;

    // Respect the unsaved-changes guard set by index.html.
    if (window._dirty === true) {
      if (!confirm('You have unsaved changes. Leave this page?')) return;
      window._dirty = false;   // User accepted — clear flag before nav
    }

    _navigating = true;
    _barStart();

    try {
      const res = await fetch(url, {
        credentials: 'same-origin',
        headers:     { 'X-PJAX': 'true' },
      });

      // Server redirected (e.g. session expired → login page).
      if (res.redirected) {
        window.location.href = res.url;
        return;
      }

      if (!res.ok) {
        // 404 / 500 etc — fall through to full navigation below.
        throw new Error(`HTTP ${res.status}`);
      }

      _barProgress();

      const html = await res.text();
      const doc  = new DOMParser().parseFromString(html, 'text/html');

      // Update page title.
      document.title = doc.title;

      _syncHead(doc);
      _swapBody(doc);

      if (pushState) {
        history.pushState({ pjax: true, url }, '', url);
      }

      window.scrollTo({ top: 0, behavior: 'instant' });
      _barFinish();

    } catch (err) {
      // Any error — degrade to a full browser navigation.
      console.warn('[PJAX] navigation failed, falling back:', err.message);
      _barError();
      window.location.href = url;
    } finally {
      _navigating = false;
    }
  }

  /* ── Click interception ──────────────────────────────────────────────── */

  document.addEventListener('click', e => {
    // Only plain left-clicks with no modifier keys.
    if (e.button !== 0 || e.ctrlKey || e.metaKey || e.shiftKey || e.altKey) return;

    const a = e.target.closest('a[href]');
    if (!a) return;
    if (a.target === '_blank')           return;   // new tab — let it go
    if (a.dataset.pjax === 'false')      return;   // explicit opt-out
    if (a.getAttribute('href').startsWith('#')) return; // anchor — let it go

    const url = a.href;
    if (!url.startsWith(window.location.origin)) return; // external
    if (SKIP_EXT.test(url))                      return; // file download

    e.preventDefault();
    navigate(url);
  });

  /* ── Back / Forward ──────────────────────────────────────────────────── */

  window.addEventListener('popstate', e => {
    if (e.state && e.state.pjax) {
      navigate(e.state.url || location.href, false);
    } else {
      // Unknown state (e.g. page loaded without PJAX) — reload cleanly.
      window.location.reload();
    }
  });

  /* ── Initialise ──────────────────────────────────────────────────────── */
  // `defer` guarantees document.body exists here.
  _createBar();
  history.replaceState({ pjax: true, url: location.href }, '', location.href);

  console.info('[PJAX] router active.');
})();