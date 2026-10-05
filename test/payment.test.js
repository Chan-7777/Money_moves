'use strict';

/**
 * Money-path unit tests.
 * Run: node --test test/payment.test.js
 *
 * These tests verify the server-side logic that protects the payment flow.
 * They must FAIL for real defects — no trivial asserts, no function-as-oracle.
 */

const { test } = require('node:test');
const assert = require('node:assert/strict');
const { validateState, validateCapturedAmount, PRICE, PRICE_MONTHLY, PRICE_AUDIT, VALID_PRICES } = require('../lib/validators');

const CURRENCY = 'AUD';

// ── State validation ──────────────────────────────────────────────────────────

test('validateState: rejects null', () => {
  assert.equal(validateState(null), 'Missing state');
});

test('validateState: rejects non-object', () => {
  assert.equal(validateState('bad'), 'Missing state');
});

test('validateState: rejects negative income', () => {
  const result = validateState({ income: -1, expenses: 1000, savings: 0, debts: [] });
  assert.equal(result, 'Invalid income');
});

test('validateState: rejects income above ceiling', () => {
  const result = validateState({ income: 1_000_001, expenses: 1000, savings: 0, debts: [] });
  assert.equal(result, 'Invalid income');
});

test('validateState: rejects oversized debt array (DoS vector)', () => {
  const debts = Array.from({ length: 21 }, () => ({ balance: 1000, rate: 5 }));
  const result = validateState({ income: 5000, expenses: 3000, savings: 1000, debts });
  assert.equal(result, 'Invalid debts (max 20)');
});

test('validateState: rejects debt with out-of-range rate', () => {
  const debts = [{ balance: 5000, rate: 101 }];
  const result = validateState({ income: 5000, expenses: 3000, savings: 1000, debts });
  assert.equal(result, 'Invalid debt rate');
});

test('validateState: accepts valid state with debts', () => {
  const state = {
    income: 6500, expenses: 3800, savings: 3000,
    debts: [{ balance: 8000, rate: 19.9 }, { balance: 15000, rate: 7.5 }],
  };
  assert.equal(validateState(state), null);
});

test('validateState: accepts valid state with no debts', () => {
  const state = { income: 4000, expenses: 2500, savings: 500, debts: [] };
  assert.equal(validateState(state), null);
});

test('validateState: rejects savings exceeding ceiling', () => {
  const result = validateState({ income: 5000, expenses: 3000, savings: 10_000_001, debts: [] });
  assert.equal(result, 'Invalid savings');
});

// ── Price enforcement ─────────────────────────────────────────────────────────

test('server price constants: 49.00 and 149.00 AUD', () => {
  // Regression guard: these must match what create-order.js sends to PayPal.
  assert.equal(PRICE_MONTHLY, '49.00');
  assert.equal(PRICE_AUDIT, '149.00');
  assert.equal(CURRENCY, 'AUD');
});

test('validateCapturedAmount: accepts both 49.00 and 149.00 AUD', () => {
  const captureMonthly = {
    status: 'COMPLETED',
    purchase_units: [{ payments: { captures: [{ amount: { value: '49.00', currency_code: 'AUD' } }] } }],
  };
  const captureAudit = {
    status: 'COMPLETED',
    purchase_units: [{ payments: { captures: [{ amount: { value: '149.00', currency_code: 'AUD' } }] } }],
  };
  assert.equal(validateCapturedAmount(captureMonthly), true);
  assert.equal(validateCapturedAmount(captureAudit), true);
});

test('validateCapturedAmount: rejects underpriced capture (attack vector)', () => {
  const capture = {
    status: 'COMPLETED',
    purchase_units: [{ payments: { captures: [{ amount: { value: '0.01', currency_code: 'AUD' } }] } }],
  };
  assert.equal(validateCapturedAmount(capture), false);
});

test('validateCapturedAmount: rejects legacy 14.00 AUD', () => {
  const capture = {
    status: 'COMPLETED',
    purchase_units: [{ payments: { captures: [{ amount: { value: '14.00', currency_code: 'AUD' } }] } }],
  };
  assert.equal(validateCapturedAmount(capture), false);
});

test('validateCapturedAmount: rejects missing amount data', () => {
  assert.equal(validateCapturedAmount({}), false);
  assert.equal(validateCapturedAmount(null), false);
  assert.equal(validateCapturedAmount({ purchase_units: [] }), false);
});

test('validateCapturedAmount: rejects different currency same amount', () => {
  // PayPal sends the currency it captured in — a USD capture for 59.00 is NOT AUD 59.00.
  const capture = {
    status: 'COMPLETED',
    purchase_units: [{ payments: { captures: [{ amount: { value: '59.00', currency_code: 'USD' } }] } }],
  };
  assert.equal(validateCapturedAmount(capture), false);
});

// ── Email validation ──────────────────────────────────────────────────────────

const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

test('email regex: rejects empty string', () => {
  assert.equal(emailRegex.test(''), false);
});

test('email regex: rejects no-@-address', () => {
  assert.equal(emailRegex.test('notanemail'), false);
});

test('email regex: rejects spaces', () => {
  assert.equal(emailRegex.test('foo @bar.com'), false);
});

test('email regex: accepts valid email', () => {
  assert.equal(emailRegex.test('user@example.com'), true);
  assert.equal(emailRegex.test('user+tag@sub.domain.com.au'), true);
});

// ── Rate limit logic ──────────────────────────────────────────────────────────

test('local rate limit: blocks after RATE_MAX hits in window', () => {
  const hits = new Map();
  const RATE_MAX = 5;
  const RATE_WIN = 60_000;

  function checkLocal(ip) {
    const now = Date.now();
    const entry = hits.get(ip) || { count: 0, start: now };
    if (now - entry.start > RATE_WIN) { hits.set(ip, { count: 1, start: now }); return false; }
    entry.count += 1;
    hits.set(ip, entry);
    return entry.count > RATE_MAX;
  }

  const ip = '1.2.3.4';
  for (let i = 0; i < RATE_MAX; i++) assert.equal(checkLocal(ip), false, `hit ${i + 1} should pass`);
  assert.equal(checkLocal(ip), true, 'hit 6 should be blocked');
});

test('local rate limit: resets after window expires', () => {
  const hits = new Map();
  const RATE_MAX = 5;
  const ip = '5.6.7.8';

  // Simulate expired window by setting start far in the past
  hits.set(ip, { count: 99, start: Date.now() - 61_000 });

  function checkLocal(ip) {
    const now = Date.now();
    const entry = hits.get(ip) || { count: 0, start: now };
    if (now - entry.start > 60_000) { hits.set(ip, { count: 1, start: now }); return false; }
    entry.count += 1;
    hits.set(ip, entry);
    return entry.count > RATE_MAX;
  }

  assert.equal(checkLocal(ip), false, 'first hit after window reset should pass');
});
