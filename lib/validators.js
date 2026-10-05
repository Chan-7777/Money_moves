'use strict';

const PRICE_CORE = '39.00';
const PRICE_COMPLETE = '59.00';
const VALID_PRICES = [PRICE_CORE, PRICE_COMPLETE];
const PRICE = PRICE_COMPLETE; // Default / top tier reference
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

module.exports = {
  validateState,
  validateCapturedAmount,
  PRICE,
  PRICE_CORE,
  PRICE_COMPLETE,
  VALID_PRICES,
  EMAIL_RE,
};
