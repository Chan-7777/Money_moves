'use strict';

const { log } = require('../lib/logger');

const PAYPAL_BASE = process.env.PAYPAL_ENV === 'live'
  ? 'https://api-m.paypal.com'
  : 'https://api-m.sandbox.paypal.com';

async function getAccessToken() {
  const creds = Buffer.from(
    `${process.env.PAYPAL_CLIENT_ID}:${process.env.PAYPAL_SECRET}`
  ).toString('base64');

  const res = await fetch(`${PAYPAL_BASE}/v1/oauth2/token`, {
    method: 'POST',
    headers: {
      Authorization: `Basic ${creds}`,
      'Content-Type': 'application/x-www-form-urlencoded',
    },
    body: 'grant_type=client_credentials',
  });
  const data = await res.json();
  return data.access_token;
}

module.exports = async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).end();

  const paypalClientId = process.env.PAYPAL_CLIENT_ID;
  const paypalSecret = process.env.PAYPAL_SECRET;
  const convertapiSecret = process.env.CONVERTAPI_SECRET;
  const resendApiKey = process.env.RESEND_API_KEY;

  if (
    !paypalClientId || paypalClientId.includes('REPLACE_WITH') ||
    !paypalSecret || paypalSecret.includes('REPLACE_WITH') ||
    !convertapiSecret || convertapiSecret.includes('REPLACE_WITH') ||
    !resendApiKey || resendApiKey.includes('REPLACE_WITH')
  ) {
    log('error', '[create-order] Server is not fully configured');
    return res.status(503).json({ error: 'Server is not fully configured (missing API credentials).' });
  }

  // Price is authoritative on the server — never trust client-supplied amount.
  const TIERS = {
    core: { price: '39.00', desc: 'MoneyMoves AU — Core Decision Report' },
    complete: { price: '59.00', desc: 'MoneyMoves AU — 12-Month Cashflow & Debt Blueprint (+ Bonuses)' },
  };
  const tierKey = (req.body?.tier === 'core') ? 'core' : 'complete';
  const selectedTier = TIERS[tierKey];
  const PRICE = selectedTier.price;
  const CURRENCY = 'AUD';
  const { email } = req.body || {};

  if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    return res.status(400).json({ error: 'Invalid email address' });
  }

  try {
    const token = await getAccessToken();

    const orderRes = await fetch(`${PAYPAL_BASE}/v2/checkout/orders`, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        intent: 'CAPTURE',
        purchase_units: [{
          amount: { currency_code: CURRENCY, value: PRICE },
          description: selectedTier.desc,
        }],
        ...(email && { payer: { email_address: email } }),
      }),
    });

    const order = await orderRes.json();
    if (!order.id) {
      log('error', '[create-order] PayPal order creation failed', { response: JSON.stringify(order) });
      return res.status(500).json({ error: 'Failed to create PayPal order' });
    }

    res.json({ orderID: order.id });
  } catch (err) {
    log('error', '[create-order] unhandled exception', { error: err.message });
    res.status(500).json({ error: 'Internal error' });
  }
};
