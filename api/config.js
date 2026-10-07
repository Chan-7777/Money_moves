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

  res.setHeader('Cache-Control', 's-maxage=3600, stale-while-revalidate');
  res.json({ paypalClientId: clientId, prices });
};
