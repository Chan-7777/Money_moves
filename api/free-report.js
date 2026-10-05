'use strict';

const { verifyToken } = require('@clerk/backend');
const { generateReport } = require('../build_pdf_report.js');
const { validateState } = require('../lib/validators');
const { log } = require('../lib/logger');

const _ipHits  = new Map();
const RATE_MAX = 5;
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

  const key = `ratelimit:free-report:${ip}`;
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

async function convertToPdf(docxBuffer) {
  const form = new FormData();
  form.append(
    'File',
    new Blob([docxBuffer], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }),
    'report.docx'
  );
  const res = await fetch('https://v2.convertapi.com/convert/docx/to/pdf', {
    method: 'POST',
    headers: { Authorization: `Bearer ${process.env.CONVERTAPI_SECRET}` },
    body: form,
  });
  if (!res.ok) throw new Error(`ConvertAPI error: ${res.status}`);
  const data = await res.json();
  return Buffer.from(data.Files[0].FileData, 'base64');
}

async function verifyClerkToken(token) {
  const secretKey = process.env.CLERK_SECRET_KEY;
  if (!secretKey || secretKey.includes('REPLACE_WITH')) return null;
  try {
    const verified = await verifyToken(token, {
      secretKey,
      jwtKey: process.env.CLERK_JWT_KEY,
    });
    return verified?.sub ? { uid: verified.sub } : null;
  } catch {
    return null;
  }
}

module.exports = async (req, res) => {
  if (req.method !== 'POST') return res.status(405).end();

  // Verify Clerk session token — prevents unauthenticated PDF generation
  const authHeader = req.headers['authorization'] || '';
  const idToken = authHeader.startsWith('Bearer ') ? authHeader.slice(7) : null;
  if (!idToken) {
    return res.status(401).json({ error: 'Authentication required' });
  }
  const clerkUser = await verifyClerkToken(idToken);
  if (!clerkUser) {
    return res.status(401).json({ error: 'Invalid or expired token' });
  }

  const convertapiSecret = process.env.CONVERTAPI_SECRET;
  if (!convertapiSecret || convertapiSecret.includes('REPLACE_WITH')) {
    log('error', '[free-report] CONVERTAPI_SECRET is not configured');
    return res.status(503).json({ error: 'PDF generation service is not configured on the server.' });
  }

  const ip = ((req.headers['x-forwarded-for'] || '').split(',')[0].trim())
    || req.socket?.remoteAddress
    || 'unknown';
  if (await checkUpstashRateLimit(ip)) {
    return res.status(429).json({ error: 'Too many requests — try again in a minute.' });
  }

  const { state } = req.body || {};
  const stateErr = validateState(state);
  if (stateErr) return res.status(400).json({ error: stateErr });

  try {
    const docxBuffer = await generateReport(state);
    const pdfBuffer  = await convertToPdf(docxBuffer);

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', 'attachment; filename="MoneyMoves_AU_Plan.pdf"');
    res.setHeader('Content-Length', pdfBuffer.length);
    res.status(200).send(pdfBuffer);
  } catch (err) {
    log('error', '[free-report] unhandled exception', { error: err.message });
    res.status(500).json({ error: err.message });
  }
};
