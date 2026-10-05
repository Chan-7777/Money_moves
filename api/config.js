'use strict';

const { log } = require('../lib/logger');

module.exports = (req, res) => {
  const clientId = process.env.PAYPAL_CLIENT_ID;
  if (!clientId || clientId.includes('REPLACE_WITH')) {
    log('error', '[config] PAYPAL_CLIENT_ID is not configured');
    return res.status(503).json({ error: 'PayPal client ID is not configured on the server.' });
  }

  res.setHeader('Cache-Control', 's-maxage=3600, stale-while-revalidate');
  res.json({ paypalClientId: clientId });
};
