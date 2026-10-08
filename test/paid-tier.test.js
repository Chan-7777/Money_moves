'use strict';

/**
 * P0-2 — paid options stay hidden until the paid tier is switched on and, in
 * production, PayPal is live.
 * Run: node --test test/paid-tier.test.js
 */

const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const config = require('../api/config');

const read = (f) => fs.readFileSync(path.join(__dirname, '..', f), 'utf8');

function callConfig(env) {
  const keys = ['PAYPAL_CLIENT_ID', 'PAID_TIER_ENABLED', 'VERCEL_ENV', 'PAYPAL_ENV'];
  for (const k of keys) delete process.env[k];
  Object.assign(process.env, { PAYPAL_CLIENT_ID: 'test-client' }, env);
  const res = {
    statusCode: 200, body: undefined,
    status(c) { this.statusCode = c; return this; },
    json(b) { this.body = b; return this; },
    setHeader() {},
  };
  config({ method: 'GET' }, res);
  return res;
}

// ── /api/config decides ───────────────────────────────────────────────────────

test('paid tier is off when PAID_TIER_ENABLED is not set', () => {
  assert.equal(callConfig({}).body.paidEnabled, false);
});

test('paid tier is off unless PAID_TIER_ENABLED is exactly "true"', () => {
  for (const v of ['1', 'yes', 'TRUE', 'false', '']) {
    assert.equal(callConfig({ PAID_TIER_ENABLED: v }).body.paidEnabled, false, `value ${JSON.stringify(v)}`);
  }
});

test('paid tier can be switched on outside production with sandbox PayPal (preview testing)', () => {
  assert.equal(callConfig({ PAID_TIER_ENABLED: 'true', VERCEL_ENV: 'preview' }).body.paidEnabled, true);
  assert.equal(callConfig({ PAID_TIER_ENABLED: 'true' }).body.paidEnabled, true); // local dev
});

test('production never shows the paid tier while PayPal is in sandbox', () => {
  assert.equal(callConfig({ PAID_TIER_ENABLED: 'true', VERCEL_ENV: 'production' }).body.paidEnabled, false);
  assert.equal(callConfig({ PAID_TIER_ENABLED: 'true', VERCEL_ENV: 'production', PAYPAL_ENV: 'sandbox' }).body.paidEnabled, false);
});

test('production shows the paid tier once switched on with live PayPal', () => {
  assert.equal(callConfig({ PAID_TIER_ENABLED: 'true', VERCEL_ENV: 'production', PAYPAL_ENV: 'live' }).body.paidEnabled, true);
});

test('live PayPal alone does not switch the paid tier on', () => {
  assert.equal(callConfig({ VERCEL_ENV: 'production', PAYPAL_ENV: 'live' }).body.paidEnabled, false);
});

// ── pages hide paid content until the server says it is live ──────────────────

// Index range [start, end) of the element whose opening tag starts at `start`.
function elementRange(html, start) {
  const tag = html.slice(start).match(/^<([a-z0-9]+)/i)[1];
  const re = new RegExp(`<${tag}\\b|</${tag}>`, 'gi');
  re.lastIndex = start;
  let depth = 0, m;
  while ((m = re.exec(html))) {
    depth += m[0].startsWith('</') ? -1 : 1;
    if (depth === 0) return [start, m.index + m[0].length];
  }
  throw new Error(`unclosed <${tag}> at ${start}`);
}

function hiddenPaidRanges(html) {
  const ranges = [];
  for (const m of html.matchAll(/<[a-z0-9]+\b[^>]*\bdata-paid-only\b[^>]*>/gi)) {
    assert.match(m[0], /\shidden[\s>]/, `paid-only element must start hidden: ${m[0].slice(0, 80)}`);
    ranges.push(elementRange(html, m.index));
  }
  return ranges;
}

const inside = (ranges, i) => ranges.some(([a, b]) => i >= a && i < b);

test('index.html: every price sits inside a paid-only block that starts hidden', () => {
  const html = read('index.html');
  const ranges = hiddenPaidRanges(html);
  const prices = [...html.matchAll(/data-tier-price=/g)];
  assert.ok(prices.length >= 2);
  for (const p of prices) assert.ok(inside(ranges, p.index), `price at ${p.index} is visible before the server allows it`);
  for (const m of html.matchAll(/one-off/gi)) assert.ok(inside(ranges, m.index), `"one-off" at ${m.index} is visible before the server allows it`);
});

test('index.html: paid blocks are revealed only when /api/config says paidEnabled === true', () => {
  const html = read('index.html');
  assert.match(html, /cfg\.paidEnabled !== true\) return;[\s\S]*?\[data-paid-only\][\s\S]*?hidden = false/);
  assert.match(html, /\[data-paid-only\]\[hidden\]\s*\{\s*display:\s*none\s*!important;?\s*\}/);
});

test('app.html: the one-off options teaser starts hidden and is revealed only by paidEnabled === true', () => {
  const html = read('app.html');
  const ranges = hiddenPaidRanges(html);
  const teaser = html.indexOf('showPaidOptions()');
  assert.ok(teaser > 0);
  assert.ok(inside(ranges, teaser), 'the button that opens paid options must be inside a hidden paid-only block');
  assert.match(html, /cfg\.paidEnabled !== true\) return;[\s\S]*?\[data-paid-only\][\s\S]*?hidden = false/);
  assert.match(html, /\[data-paid-only\]\[hidden\]\s*\{\s*display:\s*none\s*!important;?\s*\}/);
});
