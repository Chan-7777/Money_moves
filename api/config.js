'use strict';

const { log } = require('../lib/logger');
const { TIERS } = require('../lib/validators');

module.exports = (req, res) => {
  const clientId = process.env.PAYPAL_CLIENT_ID;
  if (!clientId || clientId.includes('REPLACE_WITH')) {
    log('error', '[config] PAYPAL_CLIENT_ID is not configured');
    return res.status(503).json({ error: 'PayPal client ID is not configured on the server.' });
  }

  // Prices come from the same list create-order charges, so the page cannot drift from the server.
  const prices = Object.fromEntries(Object.entries(TIERS).map(([key, t]) => [key, t.price]));

  // Paid options are shown only when switched on (PAID_TIER_ENABLED=true), and production
  // additionally needs live PayPal, so real visitors are never offered a sandbox checkout.
  // Previews and local dev can switch it on with sandbox PayPal for testing.
  const paidEnabled = process.env.PAID_TIER_ENABLED === 'true'
    && (process.env.VERCEL_ENV !== 'production' || process.env.PAYPAL_ENV === 'live');

  res.setHeader('Cache-Control', 's-maxage=3600, stale-while-revalidate');
  res.json({ paypalClientId: clientId, prices, paidEnabled });
};
