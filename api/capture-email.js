'use strict';

const { log } = require('../lib/logger');

function escapeHtml(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

const _ipHits  = new Map();
const RATE_MAX = 3;
const RATE_WIN = 60_000; // ms

function checkLocalRateLimit(ip) {
  const now   = Date.now();
  const entry = _ipHits.get(ip) || { count: 0, start: now };
  if (now - entry.start > RATE_WIN) {
    _ipHits.set(ip, { count: 1, start: now });
    return false;
  }
  entry.count += 1;
  _ipHits.set(ip, entry);
  return entry.count > RATE_MAX;
}

async function checkUpstashRateLimit(ip) {
  const url = process.env.UPSTASH_REDIS_REST_URL;
  const token = process.env.UPSTASH_REDIS_REST_TOKEN;
  if (!url || !token) {
    return checkLocalRateLimit(ip);
  }

  const key = `ratelimit:capture-email:${ip}`;
  try {
    const res = await fetch(`${url}/pipeline`, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify([
        ['INCR', key],
        ['EXPIRE', key, '60', 'NX']
      ]),
    });
    if (!res.ok) {
      log('warn', '[ratelimit] Upstash returned status', { status: res.status });
      return checkLocalRateLimit(ip);
    }
    const results = await res.json();
    const count = results[0]?.result || 1;
    return count > RATE_MAX;
  } catch (err) {
    log('error', '[ratelimit] Upstash error', { error: err.message });
    return checkLocalRateLimit(ip);
  }
}

module.exports = async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).end();

  const ip = ((req.headers['x-forwarded-for'] || '').split(',')[0].trim())
    || req.socket?.remoteAddress
    || 'unknown';
  if (await checkUpstashRateLimit(ip)) {
    return res.status(429).json({ error: 'Too many requests — try again in a minute.' });
  }

  const { email, plan } = req.body || {};
  if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    return res.status(400).json({ error: 'Invalid email' });
  }

  const resendApiKey = process.env.RESEND_API_KEY;
  if (!resendApiKey || resendApiKey.includes('REPLACE_WITH')) {
    log('error', '[capture-email] RESEND_API_KEY is not configured');
    return res.status(503).json({ error: 'Email service is not configured on the server.' });
  }

  const resp = await fetch('https://api.resend.com/emails', {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${resendApiKey}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      from: 'MoneyMoves AU <onboarding@resend.dev>',
      to: email,
      subject: 'Your MoneyMoves AU plan',
      html: `<p>Hi there,</p>
<p>Here's a snapshot from your MoneyMoves AU session:</p>
<p style="background:#f0f7f3;padding:14px;border-radius:8px;font-family:monospace;font-size:14px;">${escapeHtml(plan || '')}</p>
<p>Sign back in any time to re-run your numbers — the plan updates live as you change inputs.</p>
<p>— MoneyMoves AU team</p>`,
    }),
  });

  if (!resp.ok) {
    const err = await resp.text();
    log('error', '[capture-email] Resend error', { error: err });
    return res.status(500).json({ error: 'Email send failed' });
  }

  res.json({ ok: true });
};
