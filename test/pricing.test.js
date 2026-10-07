'use strict';

/**
 * P0-1 — one one-off price list everywhere.
 * Run: node --test test/pricing.test.js
 *
 * Expected prices are the owner's decision (2026-10-07), written here by hand:
 *   plan      A$49.00 one-off  → plan PDF without the action toolkits
 *   complete  A$159.00 one-off → plan PDF with the action toolkits
 * They are NOT read from lib/validators.js, so a wrong price in the code fails here.
 */

const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const ROOT = path.join(__dirname, '..');
const read = (f) => fs.readFileSync(path.join(ROOT, f), 'utf8');

const EXPECTED = { plan: '49.00', complete: '159.00' };
const EXPECTED_LABEL = { plan: 'A$49', complete: 'A$159' };

// Spy on the report builder before capture-order destructures it.
const reportModule = require('../build_pdf_report');
const realGenerateReport = reportModule.generateReport;
const generateCalls = [];
reportModule.generateReport = async (state, opts) => {
  generateCalls.push(opts);
  return realGenerateReport(state, opts);
};

const { TIERS, validateCapturedAmount, tierForAmount } = require('../lib/validators');
const createOrder = require('../api/create-order');
const captureOrder = require('../api/capture-order');
const config = require('../api/config');

// ── helpers ───────────────────────────────────────────────────────────────────

function mockRes() {
  return {
    statusCode: 200, body: undefined, headers: {},
    status(c) { this.statusCode = c; return this; },
    json(b) { this.body = b; return this; },
    setHeader(k, v) { this.headers[k] = v; },
    end() { return this; },
  };
}

function setEnv() {
  process.env.PAYPAL_CLIENT_ID = 'test-client';
  process.env.PAYPAL_SECRET = 'test-secret';
  process.env.CONVERTAPI_SECRET = 'test-convert';
  process.env.RESEND_API_KEY = 'test-resend';
  delete process.env.UPSTASH_REDIS_REST_URL;
  delete process.env.UPSTASH_REDIS_REST_TOKEN;
}

const json = (b, ok = true) => ({ ok, status: ok ? 200 : 500, json: async () => b, text: async () => JSON.stringify(b) });

// Routes every external call the payment handlers make; records each request.
function stubFetch({ capturedValue = '49.00', currency = 'AUD' } = {}) {
  const calls = [];
  global.fetch = async (url, init = {}) => {
    calls.push({ url: String(url), init });
    if (url.includes('/v1/oauth2/token')) return json({ access_token: 'tok' });
    if (url.endsWith('/v2/checkout/orders')) return json({ id: 'ORDER-1' });
    if (url.includes('/capture')) {
      return json({
        status: 'COMPLETED',
        purchase_units: [{ payments: { captures: [{ amount: { value: capturedValue, currency_code: currency } }] } }],
      });
    }
    if (url.includes('convertapi.com')) return json({ Files: [{ FileData: Buffer.from('%PDF-test').toString('base64') }] });
    if (url.includes('api.resend.com')) return json({ id: 'email-1' });
    throw new Error(`unexpected fetch ${url}`);
  };
  return calls;
}

const STATE = {
  income: 5200, expenses: 3400, savings: 1500, goal: 'balance', email: 'buyer@example.com',
  debts: [{ type: 'credit_card', balance: 4800, rate: 20.99, min: 145 }],
  car: { considering: 'no' },
};

let ipCounter = 0;
const req = (body) => ({ method: 'POST', body, headers: { 'x-forwarded-for': `10.0.0.${++ipCounter}` }, socket: {} });

// ── the price list itself ─────────────────────────────────────────────────────

test('server price list is exactly plan 49.00 (no toolkits) and complete 159.00 (toolkits)', () => {
  assert.deepEqual(Object.keys(TIERS).sort(), ['complete', 'plan']);
  assert.equal(TIERS.plan.price, EXPECTED.plan);
  assert.equal(TIERS.complete.price, EXPECTED.complete);
  assert.equal(TIERS.plan.includeToolkits, false);
  assert.equal(TIERS.complete.includeToolkits, true);
});

test('validator accepts only 49.00 and 159.00 AUD', () => {
  const cap = (value, currency_code = 'AUD') => ({ purchase_units: [{ payments: { captures: [{ amount: { value, currency_code } }] } }] });
  assert.equal(validateCapturedAmount(cap('49.00')), true);
  assert.equal(validateCapturedAmount(cap('159.00')), true);
  for (const old of ['149.00', '39.00', '59.00', '14.00', '0.01', '158.99', '49', '1590.00']) {
    assert.equal(validateCapturedAmount(cap(old)), false, `${old} must be rejected`);
  }
  assert.equal(validateCapturedAmount(cap('159.00', 'USD')), false);
});

test('captured amount maps to the tier it paid for', () => {
  assert.equal(tierForAmount('49.00'), 'plan');
  assert.equal(tierForAmount('159.00'), 'complete');
  assert.equal(tierForAmount('149.00'), null);
  assert.equal(tierForAmount(undefined), null);
});

// ── /api/config serves the same prices ────────────────────────────────────────

test('/api/config serves the server prices', () => {
  setEnv();
  const res = mockRes();
  config({ method: 'GET' }, res);
  assert.equal(res.statusCode, 200);
  assert.deepEqual(res.body.prices, EXPECTED);
});

test('/api/config still returns 503 when PayPal is not configured', () => {
  setEnv();
  delete process.env.PAYPAL_CLIENT_ID;
  const res = mockRes();
  config({ method: 'GET' }, res);
  assert.equal(res.statusCode, 503);
});

// ── create-order charges the server price for the chosen tier ────────────────

for (const tier of ['plan', 'complete']) {
  test(`create-order charges ${EXPECTED[tier]} AUD for tier "${tier}", ignoring any client amount`, async () => {
    setEnv();
    const calls = stubFetch();
    const res = mockRes();
    await createOrder(req({ tier, amount: '0.01', price: '0.01' }), res);
    assert.equal(res.statusCode, 200);
    const order = calls.find(c => c.url.endsWith('/v2/checkout/orders'));
    const unit = JSON.parse(order.init.body).purchase_units[0];
    assert.deepEqual(unit.amount, { currency_code: 'AUD', value: EXPECTED[tier] });
    assert.doesNotMatch(unit.description, /month|co-?pilot/i);
  });
}

for (const tier of ['copilot', 'audit', 'core', '', undefined, 'PLAN']) {
  test(`create-order rejects unknown tier ${JSON.stringify(tier)} without calling PayPal`, async () => {
    setEnv();
    const calls = stubFetch();
    const res = mockRes();
    await createOrder(req({ tier }), res);
    assert.equal(res.statusCode, 400);
    assert.equal(calls.length, 0);
  });
}

// ── capture-order delivers the report that was paid for ──────────────────────

test('capture-order: A$49 capture → report without toolkits', async () => {
  setEnv();
  generateCalls.length = 0;
  const calls = stubFetch({ capturedValue: '49.00' });
  const res = mockRes();
  await captureOrder(req({ orderID: 'CAP-49', state: STATE }), res);
  assert.equal(res.statusCode, 200);
  assert.equal(res.body.success, true);
  assert.deepEqual(generateCalls, [{ includeBonuses: false }]);
  const email = JSON.parse(calls.find(c => c.url.includes('resend')).init.body);
  assert.doesNotMatch(email.html, /toolkit/i);
});

test('capture-order: A$159 capture → report with toolkits', async () => {
  setEnv();
  generateCalls.length = 0;
  const calls = stubFetch({ capturedValue: '159.00' });
  const res = mockRes();
  await captureOrder(req({ orderID: 'CAP-159', state: STATE }), res);
  assert.equal(res.statusCode, 200);
  assert.deepEqual(generateCalls, [{ includeBonuses: true }]);
  const email = JSON.parse(calls.find(c => c.url.includes('resend')).init.body);
  assert.match(email.html, /toolkit/i);
});

test('capture-order: old A$149 capture is refused and nothing is delivered', async () => {
  setEnv();
  generateCalls.length = 0;
  const calls = stubFetch({ capturedValue: '149.00' });
  const res = mockRes();
  await captureOrder(req({ orderID: 'CAP-149', state: STATE }), res);
  assert.equal(res.statusCode, 400);
  assert.equal(generateCalls.length, 0);
  assert.equal(calls.some(c => c.url.includes('convertapi') || c.url.includes('resend')), false);
});

test('capture-order emails carry no invented guarantee figures', () => {
  const src = read('api/capture-order.js');
  assert.doesNotMatch(src, /A\$1,400|100x/i);
});

// ── every page shows the server price ─────────────────────────────────────────

function shownPrices(html) {
  const out = {};
  for (const m of html.matchAll(/data-tier-price="(\w+)"[^>]*>([^<]*)</g)) {
    (out[m[1]] ||= new Set()).add(m[2].trim());
  }
  return out;
}

for (const page of ['index.html', 'terms.html']) {
  test(`${page}: every price shown equals the server price, and both tiers are shown`, () => {
    const shown = shownPrices(read(page));
    assert.deepEqual(Object.keys(shown).sort(), ['complete', 'plan']);
    for (const tier of ['plan', 'complete']) {
      const want = page === 'terms.html' ? `A$${EXPECTED[tier]}` : EXPECTED_LABEL[tier];
      assert.deepEqual([...shown[tier]], [want], `${page} ${tier}`);
    }
  });
}

test('app.html: built-in prices equal the server prices and every price is rendered from them', () => {
  const html = read('app.html');
  const m = html.match(/const TIER_PRICES = \{ plan: '([\d.]+)', complete: '([\d.]+)' \};/);
  assert.ok(m, 'TIER_PRICES declaration not found');
  assert.deepEqual({ plan: m[1], complete: m[2] }, EXPECTED);
  // No typed-in price literals: every price comes from TIER_PRICES (updated from /api/config).
  assert.doesNotMatch(html, /A\$(49|149|159)\b/);
  assert.match(html, /data-tier-price="plan"/);
  assert.match(html, /data-tier-price="complete"/);
  assert.match(html, /cfg\.prices/, 'app.html must read prices from /api/config');
});

// ── removed wording stays removed ─────────────────────────────────────────────

const BANNED = [
  [/A\$\s?\d[\d.,]*\s*(<[^>]*>\s*)*\/\s*mo(nth)?\b/i, 'price per month'],
  [/\/\s*month\b/i, '/month'],
  [/cancel[- ]anytime/i, 'cancel anytime'],
  [/rate defen[cs]e/i, 'rate defense'],
  [/save A\$100/i, 'Save A$100'],
  [/A\$(39|59|149)\b/, 'old A$39 / A$59 / A$149 prices'],
  [/co-?pilot/i, 'Co-Pilot (monthly tier name)'],
  [/100x|A\$1,400/i, 'unsourced "100x" / A$1,400 guarantee figure'],
];

for (const page of ['index.html', 'app.html', 'terms.html']) {
  test(`${page}: no monthly / old-price wording`, () => {
    const html = read(page);
    for (const [re, label] of BANNED) assert.doesNotMatch(html, re, `${page} still has ${label}`);
  });
}
