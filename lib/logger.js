'use strict';

/**
 * Structured logger. Every call writes JSON to stderr/stdout so Vercel's
 * Log Explorer can filter by level/field. Errors are also forwarded to
 * Sentry via HTTP (no SDK) when SENTRY_DSN is configured.
 */

function log(level, message, context = {}) {
  const entry = JSON.stringify({
    ts: new Date().toISOString(),
    level,
    msg: message,
    env: process.env.VERCEL_ENV || 'development',
    ...context,
  });

  if (level === 'error') console.error(entry);
  else if (level === 'warn') console.warn(entry);
  else console.log(entry);

  if (level === 'error' && process.env.SENTRY_DSN) {
    _sendToSentry(message, context).catch(() => {});
  }
}

async function _sendToSentry(message, context) {
  const dsn = new URL(process.env.SENTRY_DSN);
  const projectId = dsn.pathname.slice(1);
  const endpoint = `${dsn.protocol}//${dsn.host}/api/${projectId}/store/`;

  await fetch(endpoint, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'X-Sentry-Auth': [
        'Sentry sentry_version=7',
        `sentry_key=${dsn.username}`,
        'sentry_client=moneymoves-au/1.0',
      ].join(', '),
    },
    body: JSON.stringify({
      timestamp: new Date().toISOString(),
      platform: 'node',
      level: 'error',
      message,
      extra: context,
      environment: process.env.VERCEL_ENV || 'development',
    }),
  });
}

module.exports = { log };
