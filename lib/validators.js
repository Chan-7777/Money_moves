'use strict';

// One-off prices (AUD) — the only place a price is set. create-order charges these,
// capture-order accepts only these, and /api/config serves them to the pages.
const TIERS = Object.freeze({
  plan: Object.freeze({
    price: '49.00',
    includeToolkits: false,
    desc: 'MoneyMoves AU — Money Plan PDF (one-off)',
  }),
  complete: Object.freeze({
    price: '159.00',
    includeToolkits: true,
    desc: 'MoneyMoves AU — Money Plan PDF + toolkits (one-off)',
  }),
});
const VALID_PRICES = Object.values(TIERS).map(t => t.price);
const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

function validateState(state) {
  if (!state || typeof state !== 'object') return 'Missing state';
  const n = (v) => typeof v === 'number' && isFinite(v);
  if (!n(state.income) || state.income < 0 || state.income > 1_000_000) return 'Invalid income';
  if (!n(state.expenses) || state.expenses < 0 || state.expenses > 1_000_000) return 'Invalid expenses';
  if (!n(state.savings) || state.savings < 0 || state.savings > 10_000_000) return 'Invalid savings';
  if (!Array.isArray(state.debts) || state.debts.length > 20) return 'Invalid debts (max 20)';
  for (const d of state.debts) {
    if (!n(d.balance) || d.balance < 0 || d.balance > 10_000_000) return 'Invalid debt balance';
    if (!n(d.rate) || d.rate < 0 || d.rate > 100) return 'Invalid debt rate';
  }
  return null;
}

function validateCapturedAmount(capture) {
  const value = capture?.purchase_units?.[0]?.payments?.captures?.[0]?.amount?.value;
  const currency = capture?.purchase_units?.[0]?.payments?.captures?.[0]?.amount?.currency_code;
  if (currency && currency !== 'AUD') return false;
  return VALID_PRICES.includes(value);
}

// The tier a captured amount paid for, or null if it matches no price.
function tierForAmount(value) {
  return Object.keys(TIERS).find(k => TIERS[k].price === value) || null;
}

module.exports = {
  validateState,
  validateCapturedAmount,
  tierForAmount,
  TIERS,
  VALID_PRICES,
  EMAIL_RE,
};
