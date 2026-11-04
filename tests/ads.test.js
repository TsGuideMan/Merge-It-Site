'use strict';

const assert = require('node:assert/strict');
const test = require('node:test');
const ads = require('../app/ads.js');

function makeDocument(placements) {
  const scripts = [];
  const document = {
    head: { append: script => scripts.push(script) },
    createElement(tagName) {
      return {
        tagName,
        dataset: {},
        style: {},
        events: {},
        append(...children) { this.children = children; },
        addEventListener(name, callback) { this.events[name] = callback; },
        isConnected: false
      };
    },
    querySelector(selector) {
      return selector === 'script[data-merge-it-adsense]' ? scripts[0] || null : null;
    },
    querySelectorAll() { return placements; }
  };
  return { document, scripts };
}

test('shows reserved ad spaces without requesting ads when no publisher is configured', () => {
  const element = {
    dataset: { adsenseSlot: 'appTop' },
    hidden: true,
    replaceChildren(...children) { this.children = children; }
  };
  const { document, scripts } = makeDocument([element]);
  const result = ads.initialize(document, { location: { protocol: 'https:' } }, {
    publisherId: '',
    slots: { appTop: '' }
  });
  assert.deepEqual(result, {
    enabled: false,
    placeholder: true,
    count: 1,
    reason: 'adsense-not-configured'
  });
  assert.equal(scripts.length, 0);
  assert.equal(element.hidden, false);
  assert.equal(element.children[0].children[0].textContent, 'ADVERTISEMENT');
  assert.equal(element.children[0].children[1].textContent, 'Ad space reserved');
});

test('keeps the downloadable offline app free from third-party ad requests', () => {
  const { document, scripts } = makeDocument([]);
  const result = ads.initialize(document, { location: { protocol: 'file:' } }, {
    publisherId: 'ca-pub-1234567890123456',
    slots: { appTop: '1234567890' }
  });
  assert.deepEqual(result, { enabled: false, reason: 'ads-require-a-hosted-page' });
  assert.equal(scripts.length, 0);
});

test('shows clearly labeled demo ad previews without requesting AdSense', () => {
  const placement = {
    dataset: { adsenseSlot: 'appTop' },
    hidden: true,
    replaceChildren(...children) { this.children = children; }
  };
  const { document, scripts } = makeDocument([placement]);
  const window = {
    location: { protocol: 'http:', search: '?ads=preview' },
    URLSearchParams
  };
  const result = ads.initialize(document, window, { publisherId: '', slots: {} });
  assert.deepEqual(result, {
    enabled: false,
    demo: true,
    count: 1,
    reason: 'preview-only-not-live-ads'
  });
  assert.equal(placement.hidden, false);
  assert.equal(placement.children[0].children[0].textContent, 'ADVERTISEMENT · DEMO PREVIEW');
  assert.equal(placement.children[0].children[2].href, '../downloads/');
  assert.equal(scripts.length, 0);
});

test('loads a configured publisher once and creates responsive ad units', () => {
  const placements = [
    { dataset: { adsenseSlot: 'appTop' }, hidden: true, replaceChildren(unit) { this.unit = unit; } },
    { dataset: { adsenseSlot: 'appBottom' }, hidden: true, replaceChildren(unit) { this.unit = unit; } }
  ];
  const { document, scripts } = makeDocument(placements);
  const window = { location: { protocol: 'https:' } };
  const config = {
    publisherId: 'ca-pub-1234567890123456',
    slots: { appTop: '1234567890', appBottom: '9876543210' }
  };
  const result = ads.initialize(document, window, config);
  assert.deepEqual(result, { enabled: true, count: 2 });
  assert.equal(scripts.length, 1);
  assert.equal(scripts[0].src, 'https://pagead2.googlesyndication.com/pagead/js/adsbygoogle.js?client=ca-pub-1234567890123456');
  assert.equal(placements[0].hidden, false);
  assert.equal(placements[0].unit.dataset.adSlot, '1234567890');
  assert.equal(placements[1].unit.dataset.fullWidthResponsive, 'true');
  scripts[0].events.load();
  assert.equal(window.adsbygoogle.length, 2);
});
