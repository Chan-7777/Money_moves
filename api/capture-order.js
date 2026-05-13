'use strict';

const { generateReport } = require('../build_pdf_report');

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

async function sendEmail(to, pdfBuffer) {
  if (!process.env.RESEND_API_KEY) {
    console.log('[email] RESEND_API_KEY not set — skipping send');
    return;
  }
  const res = await fetch('https://api.resend.com/emails', {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${process.env.RESEND_API_KEY}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      from: 'MoneyMoves AU <onboarding@resend.dev>',
      to,
      subject: 'Your MoneyMoves AU Personal Money Plan',
      html: `<p>Hi there,</p>
<p>Thanks for your purchase! Your personalised money plan PDF is attached.</p>
<p>Open it in any PDF viewer — Adobe Reader, Preview, or your browser.</p>
<p>— MoneyMoves AU team</p>`,
      attachments: [{
        filename: 'MoneyMoves_AU_Plan.pdf',
        content: pdfBuffer.toString('base64'),
      }],
    }),
  });
  if (!res.ok) {
    const err = await res.text();
    console.error('Resend error:', err);
  }
}

module.exports = async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).end();

  const { orderID, state } = req.body || {};
  if (!orderID || !state) return res.status(400).json({ error: 'Missing orderID or state' });

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
      console.error('PayPal capture not completed:', capture);
      return res.status(400).json({ success: false, error: 'Payment not completed' });
    }

    // Generate DOCX → convert to PDF → email
    const docxBuffer = await generateReport(state);
    const pdfBuffer  = await convertToPdf(docxBuffer);

    if (state.email) {
      await sendEmail(state.email, pdfBuffer);
    }

    res.json({ success: true });
  } catch (err) {
    console.error('capture-order exception:', err);
    res.status(500).json({ success: false, error: 'Internal error' });
  }
};
