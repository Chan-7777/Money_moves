'use strict';

/**
 * CLAUDE.md rule 6 — no invented numbers or promises in user-facing copy.
 * Run: node --test test/copy-claims.test.js
 */

const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const read = (f) => fs.readFileSync(path.join(__dirname, '..', f), 'utf8');

test('after-payment rate-cut script: no floored or invented target rate, no assumed cut', () => {
  const html = read('app.html');
  assert.doesNotMatch(html, /Math\.max\(\s*5\.5/, 'floored target rate');
  assert.doesNotMatch(html, /comparable prime/i, 'unsourced claim about other lenders\' rates');
  assert.doesNotMatch(html, /targetRateStr/, 'invented target rate');
  assert.doesNotMatch(html, /If they cut 1\.5 points/i, 'assumed rate cut presented as a saving');
});

test('after-payment rate-cut script only targets a debt you can ask a lender to reprice', () => {
  const html = read('app.html');
  // HECS (indexed, not negotiable) and ATO debt (not a lender) are excluded; 0% debts too.
  assert.match(html, /const callDebt = \(state\.debts \|\| \[\]\)\s*\.filter\(d => d\.rate > 0 && d\.balance > 0 && !\['hecs', 'ato'\]\.includes\(d\.type\)\)/);
  assert.match(html, /if \(callDebt\) \{[\s\S]*?retrieval-script-text/);
});

test('results email copy does not promise monthly check-ins that do not exist', () => {
  const html = read('app.html');
  assert.doesNotMatch(html, /monthly check-?in/i);
});

test('promo video: no stale price, no unowned domain, no features the PDF does not have', () => {
  const src = read('remotion/src/MoneyMovesFlow.jsx');
  assert.doesNotMatch(src, /A\$\s?\d/, 'hard-coded price');
  assert.doesNotMatch(src, /moneymoves\.com\.au/, 'domain the site does not use');
  assert.doesNotMatch(src, /mortgage stress test/i, 'the PDF stress test is 3 months without income');
  assert.doesNotMatch(src, /7-day action checklist/i, 'the checklist is dated, not 7-day');
  assert.doesNotMatch(src, /7-Section/i, 'section count depends on whether a car is planned');
  assert.match(src, /moneymoves-au\.vercel\.app/);
});
