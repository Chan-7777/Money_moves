'use strict';

const { generateReport } = require('../build_pdf_report');
const { validateState, validateCapturedAmount } = require('../lib/validators');
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

async function convertToPdf(docxBuffer) {
  // FormData and Blob are global in Node 22
  const form = new FormData();
  form.append(
    'File',
    new Blob([docxBuffer], {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    }),
    'report.docx'
  );

  const res = await fetch('https://v2.convertapi.com/convert/docx/to/pdf', {
    method: 'POST',
    headers: { Authorization: `Bearer ${process.env.CONVERTAPI_SECRET}` },
    body: form,
  });

  if (!res.ok) {
    const err = await res.text();
    throw new Error(`ConvertAPI error ${res.status}: ${err}`);
  }

  const data = await res.json();
  return Buffer.from(data.Files[0].FileData, 'base64');
}

// ── Rate limiting ─────────────────────────────────────────────────────────────

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

  const key = `ratelimit:capture-order:${ip}`;
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

// ── Idempotency + purchase record ─────────────────────────────────────────────
// Prevents double-captures when the same orderID is submitted twice (double-click,
// network retry). Stores a record in Upstash (30-day TTL) with in-memory fallback.

const _processedOrders = new Map();
const ORDER_TTL = 2_592_000; // 30 days in seconds

async function checkIdempotency(orderID) {
  const url = process.env.UPSTASH_REDIS_REST_URL;
  const token = process.env.UPSTASH_REDIS_REST_TOKEN;
  if (url && token) {
    try {
      const res = await fetch(`${url}/pipeline`, {
        method: 'POST',
        headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
        body: JSON.stringify([['GET', `order:${orderID}`]]),
      });
      if (res.ok) {
        const results = await res.json();
        const raw = results[0]?.result;
        if (raw) return JSON.parse(raw);
      }
    } catch (err) {
      log('warn', '[idempotency] Upstash GET failed, falling back to in-memory', { error: err.message });
    }
  }
  return _processedOrders.get(orderID) || null;
}

async function storeOrderRecord(orderID, record) {
  _processedOrders.set(orderID, record); // always update in-memory
  const url = process.env.UPSTASH_REDIS_REST_URL;
  const token = process.env.UPSTASH_REDIS_REST_TOKEN;
  if (url && token) {
    try {
      await fetch(`${url}/pipeline`, {
        method: 'POST',
        headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
        body: JSON.stringify([
          ['SET', `order:${orderID}`, JSON.stringify(record), 'EX', String(ORDER_TTL)]
        ]),
      });
    } catch (err) {
      log('warn', '[order-record] Failed to persist to Upstash', { orderID, error: err.message });
    }
  }
}

// ── Email delivery ────────────────────────────────────────────────────────────

async function sendEmail(to, pdfBuffer, isComplete = true) {
  const resendApiKey = process.env.RESEND_API_KEY;
  if (!resendApiKey || resendApiKey.includes('REPLACE_WITH')) {
    throw new Error('Email service is not configured on the server.');
  }

  const subject = isComplete
    ? 'Your MoneyMoves AU 12-Month Cashflow & Debt Blueprint (+ Bonuses)'
    : 'Your MoneyMoves AU Core Decision Report';

  const bonusHtml = isComplete
    ? `<ul>
  <li><strong>Sections 1–7:</strong> Your personalised 12-month debt elimination roadmap, car purchase scenarios, 5-year running costs, and inflation stress-tests.</li>
  <li><strong>Bonus Toolkit 1:</strong> Aussie Car Dealer Negotiation Script & Finance Checklist</li>
  <li><strong>Bonus Toolkit 2:</strong> 5-Minute Aussie Bank Rate-Cut Cheatsheet</li>
  <li><strong>Bonus Toolkit 3:</strong> Set-and-Forget Payday Automation Setup</li>
</ul>
<p><strong>Our 30-Day "100x Value" Guarantee:</strong> If this plan does not uncover at least A$1,400 in potential interest savings, lower loan costs, or cashflow improvements over the next 12 months, simply reply directly to this email within 30 days for a prompt, courteous 100% refund.</p>`
    : `<p>Your personalized 7-section financial report is attached, featuring your 12-month debt roadmap, car scenarios, and cashflow map.</p>`;

  const filename = isComplete
    ? 'MoneyMoves_AU_12Month_Blueprint.pdf'
    : 'MoneyMoves_AU_Core_Report.pdf';

  const res = await fetch('https://api.resend.com/emails', {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${resendApiKey}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      from: 'MoneyMoves AU <onboarding@resend.dev>',
      to,
      subject,
      html: `<p>Hi there,</p>
<p>Thanks for ordering! Your personalised financial decision pack is attached as a PDF.</p>
${bonusHtml}
<p>— The MoneyMoves AU Team</p>`,
      attachments: [{
        filename,
        content: pdfBuffer.toString('base64'),
      }],
    }),
  });
  if (!res.ok) {
    const err = await res.text();
    throw new Error(`Resend error: ${err}`);
  }
}

// ── Handler ───────────────────────────────────────────────────────────────────

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
    log('error', '[capture-order] Server is not fully configured');
    return res.status(503).json({ error: 'Server is not fully configured (missing API credentials).' });
  }

  const ip = ((req.headers['x-forwarded-for'] || '').split(',')[0].trim())
    || req.socket?.remoteAddress
    || 'unknown';
  if (await checkUpstashRateLimit(ip)) {
    return res.status(429).json({ error: 'Too many requests — try again in a minute.' });
  }

  const { orderID, state } = req.body || {};
  if (!orderID) return res.status(400).json({ error: 'Missing orderID' });

  // Idempotency guard — return immediately for already-processed orders
  const existing = await checkIdempotency(orderID);
  if (existing) {
    log('info', '[capture-order] duplicate orderID — returning cached result', { orderID });
    return res.json({ success: true, alreadyCaptured: true, emailSent: existing.emailSent });
  }

  const stateErr = validateState(state);
  if (stateErr) return res.status(400).json({ error: stateErr });

  if (state.email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(state.email)) {
    return res.status(400).json({ error: 'Invalid email in state' });
  }

  try {
    const token = await getAccessToken();

    // Capture the PayPal order
    const captureRes = await fetch(`${PAYPAL_BASE}/v2/checkout/orders/${orderID}/capture`, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
    });
    const capture = await captureRes.json();

    if (capture.status !== 'COMPLETED') {
      log('error', '[capture-order] PayPal capture not completed', { orderID, status: capture.status });
      return res.status(400).json({ success: false, error: 'Payment not completed' });
    }

    // Verify the captured amount matches the server-authoritative price.
    if (!validateCapturedAmount(capture)) {
      const capturedAmount = capture.purchase_units?.[0]?.payments?.captures?.[0]?.amount?.value;
      log('error', '[capture-order] amount mismatch', { orderID, capturedAmount });
      return res.status(400).json({ success: false, error: 'Payment amount mismatch' });
    }

    const capturedAmount = capture.purchase_units?.[0]?.payments?.captures?.[0]?.amount?.value;
    const isComplete = (capturedAmount === '49.00' || capturedAmount === '149.00' || capturedAmount === '59.00');
    const tierName = (capturedAmount === '149.00') ? 'audit' : 'copilot';

    // Store capture record BEFORE expensive ops — so a retry after PDF/email
    // failure returns alreadyCaptured instead of hitting PayPal a second time.
    await storeOrderRecord(orderID, {
      email: state.email || null,
      timestamp: Date.now(),
      captured: true,
      tier: tierName,
      amount: capturedAmount,
      emailSent: false,
      emailError: null,
    });

    // Generate DOCX → convert to PDF → email
    const docxBuffer = await generateReport(state, { includeBonuses: isComplete });
    const pdfBuffer  = await convertToPdf(docxBuffer);

    let emailSent = false;
    let emailError = null;

    if (state.email) {
      try {
        await sendEmail(state.email, pdfBuffer, isComplete);
        emailSent = true;
      } catch (err) {
        log('error', '[capture-order] email delivery failed', { orderID, error: err.message });
        emailError = err.message;
      }
    }

    // Update record with final email outcome
    await storeOrderRecord(orderID, {
      email: state.email || null,
      timestamp: Date.now(),
      captured: true,
      tier: tierName,
      amount: capturedAmount,
      emailSent,
      emailError: emailError || null,
    });

    res.json({
      success: true,
      emailSent,
      emailError,
      pdfData: pdfBuffer.toString('base64'),
    });
  } catch (err) {
    log('error', '[capture-order] unhandled exception', { orderID, error: err.message });
    res.status(500).json({ success: false, error: 'Internal error' });
  }
};
