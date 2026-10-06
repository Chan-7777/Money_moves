'use strict';

/**
 * PDF report tests.
 * Run: node --test test/report.test.js
 *
 * Expected figures are worked out independently (see comments), not read back
 * from the generator, so a maths regression fails here.
 */

const { test } = require('node:test');
const assert = require('node:assert/strict');
const JSZip = require('jszip');
const { deriveReportData, generateReport } = require('../build_pdf_report');

const NO_CAR = { considering: 'no', price: 0, deposit: 0, dealerRate: 0, bankRate: 0, term: 5, balloon: 'no', balloonAmount: 0, novated: 'no' };

// Take-home 5,200; costs 3,400; savings 1,500; card 4,800 @ 20.99% (min 145); HECS 24,000.
const CARD_AND_HECS = {
  income: 5200, expenses: 3400, savings: 1500, goal: 'balance',
  email: 'test@example.com', name: 'Test Person', car: NO_CAR,
  debts: [
    { type: 'credit_card', balance: 4800, rate: 20.99, min: 145 },
    { type: 'hecs', balance: 24000, rate: 0, min: 0 },
  ],
};

async function documentText(state, opts) {
  const zip = await JSZip.loadAsync(await generateReport(state, opts));
  const xml = await zip.file('word/document.xml').async('string');
  return xml.replace(/<w:tab\/>/g, ' ').replace(/<[^>]+>/g, '').replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&apos;/g, "'");
}

// ── Figures ───────────────────────────────────────────────────────────────────

test('essentials include debt minimums (matches the app: 3,400 + 145)', () => {
  const d = deriveReportData(CARD_AND_HECS);
  assert.equal(d.surplus, 1655);                 // 5,200 − 3,400 − 145
  assert.equal(d.plan.oneMonth, 3545);           // 3,400 + 145
  assert.equal(d.plan.threeMonth, 10635);        // 3 × 3,545
});

test('buffer comes first: no extra debt money until the 1-month buffer is full', () => {
  const d = deriveReportData(CARD_AND_HECS);
  const m1 = d.plan.rows[0];
  assert.equal(Math.round(m1.toBuffer), 1655);   // whole surplus to buffer
  assert.equal(m1.toDebt, 0);
  assert.equal(m1.toSave, 0);
  assert.equal(d.plan.milestones.buffer1, 2);    // 1,500 + 1,655 = 3,155 < 3,545 → reached in month 2
  assert.equal(Math.round(d.plan.rows[1].toBuffer), 390); // 3,545 − 3,155
});

test('credit card interest is charged monthly, not "payment × months − balance"', () => {
  const d = deriveReportData(CARD_AND_HECS);
  const card = d.debtOrder.find(x => x.type === 'credit_card');
  // Month 1: 4,800 × 20.99%/12 = 83.96 interest, then the 145 minimum → 4,738.96.
  assert.equal(Math.round(d.plan.rows[0].debtLeft), 4739);
  assert.equal(card.clearedMonth, 4);
  assert.ok(card.interestPaid > 200 && card.interestPaid < 320, `plan interest ${card.interestPaid}`);
  // Minimums only: 145/month at 20.99% takes 50 months and ~A$2,435 of interest
  // (standard amortisation, checked separately).
  assert.equal(card.minimumsOnlyClearedMonth, 50);
  assert.ok(Math.abs(card.minimumsOnlyInterest - 2435) <= 2, `min-only interest ${card.minimumsOnlyInterest}`);
});

test('HECS is never charged interest or given extra repayments', () => {
  const d = deriveReportData(CARD_AND_HECS);
  const hecs = d.debtOrder.find(x => x.type === 'hecs');
  assert.equal(hecs.interestPaid, 0);
  assert.equal(hecs.bal, 24000);
  assert.ok(d.plan.rows.every(r => r.debtLeft >= 0));
});

test('a cleared debt\'s minimum rolls into the surplus', () => {
  const d = deriveReportData(CARD_AND_HECS);
  assert.equal(Math.round(d.plan.rows[4].surplus), 1800); // month 5: card gone, 5,200 − 3,400
});

test('cost of inaction = minimums-only interest − plan interest over 5 years', () => {
  const d = deriveReportData(CARD_AND_HECS);
  assert.ok(Math.abs(d.baseInterest5y - 2435) <= 2);
  assert.equal(Math.round(d.costOfInaction), Math.round(d.baseInterest5y - d.planInterest5y));
});

test('avalanche-vs-equal line only exists with 2+ interest-bearing debts', () => {
  assert.equal(deriveReportData(CARD_AND_HECS).avalancheSaving, null);
  const two = { ...CARD_AND_HECS, savings: 20000, debts: [
    { type: 'credit_card', balance: 5000, rate: 21, min: 150 },
    { type: 'personal_loan', balance: 10000, rate: 8, min: 250 },
  ] };
  const d = deriveReportData(two);
  assert.ok(d.avalancheSaving > 0, `avalanche saving ${d.avalancheSaving}`);
});

test('stress test: 3 months without income uses costs + minimums', async () => {
  const text = await documentText(CARD_AND_HECS);
  assert.ok(text.includes('You need A$10,635 (costs + minimums)')); // 3 × 3,545
  assert.ok(text.includes('A$9,135 short'));                         // 10,635 − 1,500
  assert.ok(text.includes('credit card'));                          // card counted as variable-rate
  assert.ok(!text.includes('No current variable loan'));
});

// ── Document content ──────────────────────────────────────────────────────────

test('no-car plan has no placeholder pages and no car toolkit', async () => {
  const text = await documentText(CARD_AND_HECS);
  assert.ok(!/placeholder/i.test(text));
  assert.ok(!text.includes('Car decision'));
  assert.ok(!text.includes('Car dealer negotiation'));
  assert.ok(text.includes('Toolkit 1: The rate-cut phone call for your credit card'));
});

test('sections are numbered from 1 with no gaps', async () => {
  const text = await documentText(CARD_AND_HECS);
  for (const heading of ['1. How your plan works', '2. Your 12-month cash map', '3. Debt roadmap', '4. Stress tests', '5. Your action checklist']) {
    assert.ok(text.includes(heading), `missing "${heading}"`);
  }
});

test('cover uses the person\'s name and shows key dates', async () => {
  const text = await documentText(CARD_AND_HECS);
  assert.ok(text.includes('Prepared for Test Person'));
  assert.ok(!text.includes('Prepared for test@example.com'));
  assert.match(text, /key dates on your numbers/i);
  assert.match(text, /do these this week/i);
});

test('an all-lower-case name is capitalised; other names are left as typed', async () => {
  assert.ok((await documentText({ ...CARD_AND_HECS, name: 'chandan reddy' })).includes('Prepared for Chandan Reddy'));
  assert.ok((await documentText({ ...CARD_AND_HECS, name: "mary-jane o'brien" })).includes("Prepared for Mary-Jane O'Brien"));
  assert.ok((await documentText({ ...CARD_AND_HECS, name: 'Ian McDonald' })).includes('Prepared for Ian McDonald'));
  assert.ok((await documentText({ ...CARD_AND_HECS, name: 'Ana de Silva' })).includes('Prepared for Ana de Silva'));
});

test('month labels always use three-letter months', async () => {
  // The HECS paragraph mentions "June 2026" in a sentence; only date labels must be short.
  const text = (await documentText(CARD_AND_HECS)).replace(/for June 2026/g, '');
  assert.ok(!/\b(June|July|Sept)\b \d{4}/.test(text), 'found a long month label');
  assert.match(text, /\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec) 20\d\d\b/);
});

test('debt table shows the exact rate and HECS as repaid through tax', async () => {
  const text = await documentText(CARD_AND_HECS);
  assert.ok(text.includes('20.99%'));
  assert.ok(!text.includes('21.0%'));
  assert.ok(text.includes('Through tax'));
  assert.ok(text.includes('None (indexation only)'));
});

test('checklist is in date order', async () => {
  const text = await documentText(CARD_AND_HECS);
  const at = (s) => { const i = text.indexOf(s); assert.ok(i >= 0, `missing "${s}"`); return i; };
  // this week → within 2 weeks → buffer full (month 2) → re-run (90 days) → 3-month buffer (month 8)
  assert.ok(at('This week: open a separate') < at('Within 2 weeks: call'));
  assert.ok(at('Within 2 weeks: call') < at('once the buffer is there'));
  assert.ok(at('once the buffer is there') < at('re-run the free tool at'));
  assert.ok(at('re-run the free tool at') < at('your buffer reaches 3 months (A$10,635)'));
});

test('removed claims stay removed', async () => {
  const text = await documentText(CARD_AND_HECS);
  for (const phrase of [
    'Within 90 days of purchase',
    'value typically 10x',
    'A$3,000–$8,000',
    '1.5% below what I am',
    'saves you approximately A$0',
    'Financial Dossier',
  ]) {
    assert.ok(!text.includes(phrase), `found "${phrase}"`);
  }
});

test('car plan includes the car sections, numbered in order', async () => {
  const withCar = { ...CARD_AND_HECS, car: { ...NO_CAR, considering: 'yes', price: 35000, deposit: 5000, dealerRate: 9.99, bankRate: 7.25 } };
  const text = await documentText(withCar);
  assert.ok(text.includes('4. Car decision: four scenarios'));
  assert.ok(text.includes('5. What the car really costs'));
  assert.ok(text.includes('6. Stress tests'));
  assert.ok(text.includes('Toolkit 1: Car dealer negotiation'));
  assert.ok(text.includes('planned car loan'));
});

test('negative surplus: plan says close the gap, buffer is drawn down, nothing goes to debt', () => {
  const d = deriveReportData({ ...CARD_AND_HECS, income: 3000 }); // 3,000 − 3,400 − 145 = −545
  assert.equal(d.surplus, -545);
  assert.match(d.topMove.title, /spending more than you earn/);
  assert.equal(Math.round(d.plan.rows[0].fromBuffer), 545);
  assert.equal(d.plan.rows[0].toDebt, 0);
});

test('no debts and a free (no-toolkit) report still render', async () => {
  const text = await documentText({ ...CARD_AND_HECS, debts: [] }, { includeBonuses: false });
  assert.ok(text.includes('You have no debts to pay off'));
  assert.ok(!text.includes('Toolkit'));
});
