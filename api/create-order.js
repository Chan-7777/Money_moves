'use strict';

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

  const { amount = '14.00', currency = 'AUD', email } = req.body || {};

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
          amount: { currency_code: currency, value: amount },
          description: 'MoneyMoves AU — Personal Money Plan PDF',
        }],
        ...(email && { payer: { email_address: email } }),
      }),
    });

    const order = await orderRes.json();
    if (!order.id) {
      console.error('PayPal create-order error:', order);
      return res.status(500).json({ error: 'Failed to create PayPal order' });
    }

    res.json({ orderID: order.id });
  } catch (err) {
    console.error('create-order exception:', err);
    res.status(500).json({ error: 'Internal error' });
  }
};
