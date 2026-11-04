(function (root, factory) {
  'use strict';

  const ads = factory();
  if (typeof module === 'object' && module.exports) module.exports = ads;
  if (root) root.MergeItAds = ads;

  if (root && root.document && root.MERGE_IT_ADSENSE) {
    ads.initialize(root.document, root, root.MERGE_IT_ADSENSE);
  }
})(typeof globalThis === 'undefined' ? this : globalThis, function () {
  'use strict';

  const ADSENSE_URL = 'https://pagead2.googlesyndication.com/pagead/js/adsbygoogle.js';
  const PUBLISHER_ID_PATTERN = /^ca-pub-\d{16}$/;
  const SLOT_ID_PATTERN = /^\d{10}$/;

  function showAdPlaceholders(document, placements) {
    placements.forEach(element => {
      const card = document.createElement('div');
      card.className = 'ad-placeholder';

      const label = document.createElement('span');
      label.className = 'ad-preview-label';
      label.textContent = 'ADVERTISEMENT';

      const copy = document.createElement('span');
      copy.className = 'ad-preview-copy';
      copy.textContent = 'Ad space reserved';

      card.append(label, copy);
      element.replaceChildren(card);
      element.hidden = false;
    });
    return { enabled: false, placeholder: true, count: placements.length, reason: 'adsense-not-configured' };
  }

  function showDemoAds(document, placements) {
    placements.forEach(element => {
      const card = document.createElement('div');
      card.className = 'ad-preview';

      const label = document.createElement('span');
      label.className = 'ad-preview-label';
      label.textContent = 'ADVERTISEMENT · DEMO PREVIEW';

      const copy = document.createElement('span');
      copy.className = 'ad-preview-copy';
      copy.textContent = 'Try Merge It offline: download the Windows desktop version.';

      const link = document.createElement('a');
      link.href = '../downloads/';
      link.textContent = 'View download options';

      card.append(label, copy, link);
      element.replaceChildren(card);
      element.hidden = false;
    });
    return { enabled: false, demo: true, count: placements.length, reason: 'preview-only-not-live-ads' };
  }

  function initialize(document, window, config) {
    if (!/^https?:$/.test(window.location.protocol)) return { enabled: false, reason: 'ads-require-a-hosted-page' };
    const placements = Array.from(document.querySelectorAll('[data-adsense-slot]'));
    if (new URLSearchParams(window.location.search).get('ads') === 'preview') {
      return showDemoAds(document, placements);
    }
    if (!config || !PUBLISHER_ID_PATTERN.test(config.publisherId || '')) {
      return showAdPlaceholders(document, placements);
    }

    const configuredPlacements = placements
      .map(element => ({ element, slot: config.slots && config.slots[element.dataset.adsenseSlot] }))
      .filter(placement => SLOT_ID_PATTERN.test(placement.slot || ''));
    if (!configuredPlacements.length) return showAdPlaceholders(document, placements);

    let script = document.querySelector('script[data-merge-it-adsense]');
    if (!script) {
      script = document.createElement('script');
      script.async = true;
      script.crossOrigin = 'anonymous';
      script.dataset.mergeItAdsense = 'true';
      script.src = `${ADSENSE_URL}?client=${encodeURIComponent(config.publisherId)}`;
    }

    const units = configuredPlacements.map(({ element, slot }) => {
      const unit = document.createElement('ins');
      unit.className = 'adsbygoogle';
      unit.style.display = 'block';
      unit.dataset.adClient = config.publisherId;
      unit.dataset.adSlot = slot;
      unit.dataset.adFormat = 'auto';
      unit.dataset.fullWidthResponsive = 'true';
      element.replaceChildren(unit);
      element.hidden = false;
      return element;
    });

    let queued = false;
    const queueAds = () => {
      if (queued) return;
      queued = true;
      window.adsbygoogle = window.adsbygoogle || [];
      units.forEach(() => window.adsbygoogle.push({}));
    };
    const hidePlacements = () => units.forEach(element => { element.hidden = true; });

    if (script.dataset.mergeItLoaded === 'true') {
      queueAds();
    } else {
      script.addEventListener('load', () => {
        script.dataset.mergeItLoaded = 'true';
        queueAds();
      }, { once: true });
      script.addEventListener('error', () => {
        hidePlacements();
        console.warn('Merge It ads could not load. Check the AdSense setup and network connection.');
      }, { once: true });
      if (!script.isConnected) document.head.append(script);
    }

    return { enabled: true, count: units.length };
  }

  return { initialize };
});
