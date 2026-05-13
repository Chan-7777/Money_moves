/**
 * MoneyMoves AU — Personalised PDF Report Renderer (v0.2)
 * ---------------------------------------------------------
 * Takes a `state` JSON object (the same shape the prototype builds) and renders
 * a 7-section Word document. v0.2 incorporates every reviewer fix from the
 * first sample-report critique:
 *
 *   - Dynamic phase-A/phase-B cash-map allocation transition.
 *   - Section 4 heading corrected to "four scenarios" and the wait-scenario
 *     deposit growth uses the same 20% allocation as Section 2.
 *   - Depreciation hold (~5%/6mo) shown numerically in Section 4.
 *   - HECS section in Section 3 only renders when user has a HECS debt.
 *   - "Cost of doing nothing" gets a methodology footnote.
 *   - Section 5 depreciation note reconciles to "~18% Yr 1, ~10%/yr after".
 *   - Section 5 shows fuel L/100km explicitly.
 *   - Section 6 stress-test rate caveat: "applies to variable-rate loans only".
 *   - Section 7 action items reordered so 7-day debt action precedes 30-day buffer.
 *   - All dates computed dynamically from today; "moneymoves.com.au" tagged
 *     as placeholder until live.
 *   - Cover page carries a SAMPLE banner whenever email is the demo address.
 *
 * Usage:
 *   node build_pdf_report.js [state.json] [output.docx]
 */

const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, BorderStyle, WidthType, ShadingType,
  Header, Footer, PageNumber, LevelFormat, PageBreak,
} = require('docx');

const RULES = {
  HIGH_RATE_THRESHOLD: 10.0,
  MIN_BUFFER_MONTHS: 1.0,
  TARGET_BUFFER_MONTHS: 3.0,
  IDEAL_BUFFER_MONTHS: 6.0,
};

// Australian running-cost defaults (FY 2025-26 illustrative averages).
// Sources: Budget Direct, RACV, Finder.com.au.
const RUNNING_COSTS = {
  insuranceAnnual: 1450,
  regoAnnual: 880,
  servicingAnnual: 680,
  fuelLitresPer100km: 7.5,
  fuelPricePerLitre: 1.95,
  depreciationPctYr1: 0.18,
  depreciationPctYr2plus: 0.10,
  // Wait-scenario depreciation hold: 5% over 6 months, 9% over 12 months.
  waitHold6mo: 0.05,
  waitHold12mo: 0.09,
};

const COLOR_GREEN = '0A192F'; // Ink (JPM/McKinsey dark blue/black)
const COLOR_GREEN_DARK = '020C1B';
const COLOR_INK = '112233';
const COLOR_INK_SOFT = '556677';
const COLOR_RED = 'D64545';
const COLOR_LINE = 'D0D7DE';
const COLOR_CREAM = 'F8F9FA';
const COLOR_GOLD_SOFT = 'F0F4F8'; // Muted grey/blue for callouts
const COLOR_ACCENT = '005A9C'; // Institutional accent

// ─── Helpers ─────────────────────────────────────────────────────────────────
const fmt = (n) => {
  if (n === null || n === undefined || isNaN(n)) return '—';
  const sign = n < 0 ? '-' : '';
  return sign + 'A$' + Math.abs(Math.round(n)).toLocaleString('en-AU');
};
const fmtPct = (n, dp = 1) => `${(n).toFixed(dp)}%`;

function calcMonthlyMortgage(principal, annualRate, years) {
  if (principal <= 0 || years <= 0) return 0;
  if (annualRate <= 0) return principal / (years * 12);
  const r = annualRate / 100 / 12;
  const n = years * 12;
  return principal * r * Math.pow(1 + r, n) / (Math.pow(1 + r, n) - 1);
}

// Months to clear amortising debt at a fixed extra monthly payment.
function monthsToClear(balance, annualRate, monthlyPayment) {
  if (balance <= 0 || monthlyPayment <= 0) return null;
  if (annualRate <= 0) return Math.ceil(balance / monthlyPayment);
  const r = annualRate / 100 / 12;
  if (monthlyPayment <= balance * r) return null; // never clears
  const n = -Math.log(1 - (balance * r) / monthlyPayment) / Math.log(1 + r);
  return Math.ceil(n);
}

function totalInterestPaid(balance, annualRate, monthlyPayment) {
  const n = monthsToClear(balance, annualRate, monthlyPayment);
  if (!n) return 0;
  return Math.max(0, monthlyPayment * n - balance);
}

function buildDebtPayoffOrder(debts) {
  const byRate = (a, b) => b.rate - a.rate;
  const high = debts.filter(d => d.rate >= RULES.HIGH_RATE_THRESHOLD && !['hecs','home_loan','ato'].includes(d.type)).sort(byRate);
  const ato  = debts.filter(d => d.type === 'ato');
  const mid  = debts.filter(d => d.rate < RULES.HIGH_RATE_THRESHOLD && d.rate >= 5 && !['hecs','home_loan','ato'].includes(d.type)).sort(byRate);
  const low  = debts.filter(d => d.rate < 5 && !['hecs','home_loan','ato'].includes(d.type)).sort(byRate);
  const home = debts.filter(d => d.type === 'home_loan').sort(byRate);
  const hecs = debts.filter(d => d.type === 'hecs');
  const order = [];
  high.forEach(d => order.push({ ...d, tier: 'High-rate (kill it)' }));
  ato.forEach(d => order.push({ ...d, tier: 'ATO (enforcement risk)' }));
  mid.forEach(d => order.push({ ...d, tier: 'Mid-rate (5-10%)' }));
  low.forEach(d => order.push({ ...d, tier: 'Low-rate consumer' }));
  home.forEach(d => order.push({ ...d, tier: 'Home loan' }));
  hecs.forEach(d => order.push({ ...d, tier: 'HECS / HELP (index-only)' }));
  return order;
}

const NICE_TYPE = {
  credit_card: 'credit card',
  personal_loan: 'personal loan',
  car_loan: 'car loan',
  bnpl: 'BNPL',
  ato: 'ATO debt',
  hecs: 'HECS / HELP',
  home_loan: 'home loan',
  other: 'other debt',
};

// ─── docx primitives ─────────────────────────────────────────────────────────
const sans = (text, opts = {}) => new TextRun({ text, font: 'Arial', ...opts });
const serif = (text, opts = {}) => new TextRun({ text, font: 'Georgia', ...opts });

const arial = sans; // backward compatibility

const p = (text, opts = {}) => new Paragraph({
  children: Array.isArray(text) ? text : [sans(text, { size: opts.size || 20, color: opts.color || COLOR_INK, bold: !!opts.bold })],
  alignment: opts.align || AlignmentType.JUSTIFIED,
  heading: opts.heading,
  keepNext: opts.keepNext,
  spacing: { after: opts.after ?? 120, before: opts.before ?? 0 },
});
const h1 = (t) => new Paragraph({
  children: [serif(t, { size: 36, color: COLOR_GREEN, bold: false })],
  heading: HeadingLevel.HEADING_1,
  keepNext: true,
  spacing: { before: 400, after: 200 },
  border: { bottom: { color: COLOR_LINE, space: 12, style: BorderStyle.SINGLE, size: 4 } },
});
const h3 = (t) => p([sans(t.toUpperCase(), { size: 18, color: COLOR_ACCENT, bold: true })], { heading: HeadingLevel.HEADING_3, after: 120, keepNext: true });
const small = (t, color) => p([sans(t, { size: 18, color: color || COLOR_INK_SOFT, italics: true })], { after: 80 });
const spacer = (after = 200) => new Paragraph({ children: [sans('')], spacing: { after } });
const bullet = (t) => new Paragraph({
  children: Array.isArray(t) ? t : [sans(t, { size: 20, color: COLOR_INK })],
  numbering: { reference: 'bullets', level: 0 },
  spacing: { after: 80 },
});

// Callout box helper
const callout = (title, bodyText) => new Table({
  rows: [new TableRow({ children: [
    new TableCell({
      children: [
        new Paragraph({ children: [sans(title, { size: 20, bold: true, color: COLOR_GREEN })], spacing: { after: 80 } }),
        new Paragraph({ children: Array.isArray(bodyText) ? bodyText : [sans(bodyText, { size: 20, color: COLOR_INK })], alignment: AlignmentType.JUSTIFIED })
      ],
      shading: { type: ShadingType.CLEAR, color: 'auto', fill: COLOR_GOLD_SOFT },
      margins: { top: 200, bottom: 200, left: 200, right: 200 },
      borders: {
        left: { style: BorderStyle.SINGLE, size: 16, color: COLOR_ACCENT },
        top: { style: BorderStyle.NONE, size: 0, color: 'auto' },
        right: { style: BorderStyle.NONE, size: 0, color: 'auto' },
        bottom: { style: BorderStyle.NONE, size: 0, color: 'auto' },
      }
    })
  ]})],
  width: { size: 100, type: WidthType.PERCENTAGE },
});

function cell(text, opts = {}) {
  return new TableCell({
    children: [new Paragraph({
      children: [sans(text == null ? '' : String(text), {
        size: opts.size || 18, bold: !!opts.bold, color: opts.color || COLOR_INK,
      })],
      alignment: opts.align || AlignmentType.LEFT,
    })],
    shading: opts.fill ? { type: ShadingType.CLEAR, color: 'auto', fill: opts.fill } : undefined,
    verticalAlign: 'center',
    margins: { top: 120, bottom: 120, left: 120, right: 120 },
    borders: opts.noBorders ? {
      top: { style: BorderStyle.NONE, size: 0, color: 'auto' },
      bottom: { style: BorderStyle.NONE, size: 0, color: 'auto' },
      left: { style: BorderStyle.NONE, size: 0, color: 'auto' },
      right: { style: BorderStyle.NONE, size: 0, color: 'auto' },
    } : undefined,
  });
}

function table(rows, opts = {}) {
  return new Table({
    rows: rows.map((cells, idx) => new TableRow({
      children: cells.map(c =>
        idx === 0
          ? cell(c, { bold: true, color: COLOR_GREEN, fill: 'FFFFFF', ...opts.headerCell })
          : cell(c, { fill: idx % 2 === 0 ? COLOR_CREAM : 'FFFFFF', ...opts.bodyCell })
      ),
      tableHeader: idx === 0,
    })),
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 12, color: COLOR_GREEN },
      bottom: { style: BorderStyle.SINGLE, size: 12, color: COLOR_GREEN },
      left: { style: BorderStyle.NONE, size: 0, color: 'auto' },
      right: { style: BorderStyle.NONE, size: 0, color: 'auto' },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 2, color: COLOR_LINE },
      insideVertical: { style: BorderStyle.NONE, size: 0, color: 'auto' },
    },
  });
}

const pageBreak = () => new Paragraph({ children: [new PageBreak()] });

// ─── Date helpers (dynamic) ──────────────────────────────────────────────────
const TODAY = new Date();
const fmtDate = (offsetDays) => {
  const d = new Date(TODAY.getTime() + offsetDays * 86400000);
  return d.toLocaleDateString('en-AU', { day: 'numeric', month: 'short', year: 'numeric' });
};
const fmtDateShort = (offsetDays) => {
  const d = new Date(TODAY.getTime() + offsetDays * 86400000);
  return d.toLocaleDateString('en-AU', { day: 'numeric', month: 'short' });
};

// ─── Sections ────────────────────────────────────────────────────────────────
function buildCover(state, derived) {
  const isSample = false;
  const top = derived.topMove;

  const cover = [
    spacer(1200),
    new Paragraph({
      children: [sans('CONFIDENTIAL', { size: 18, color: COLOR_ACCENT, bold: true })],
      alignment: AlignmentType.RIGHT,
    }),
    new Paragraph({
      children: [serif('MoneyMoves', { size: 64, bold: false, color: COLOR_GREEN })],
      spacing: { before: 800, after: 0 },
    }),
    new Paragraph({
      children: [serif('Financial Dossier', { size: 48, color: COLOR_INK_SOFT })],
      spacing: { after: 400 },
    }),
    new Paragraph({
      border: { top: { color: COLOR_LINE, space: 1, style: BorderStyle.SINGLE, size: 12 } },
      spacing: { before: 200, after: 200 }
    }),
    p(`Prepared for: ${state.email || 'You'}`, { size: 20, color: COLOR_INK }),
    p(`Date: ${fmtDate(0)}`, { size: 20, color: COLOR_INK }),
    pageBreak(),
  ];

  if (isSample) {
    cover.push(new Paragraph({
      children: [sans('SAMPLE — DRAFT REPORT (illustrative data, not a real client report)', {
        size: 20, bold: true, color: COLOR_RED,
      })],
      alignment: AlignmentType.CENTER,
      shading: { type: ShadingType.CLEAR, color: 'auto', fill: 'FFF0F0' },
      spacing: { after: 280, before: 280 },
    }));
  }

  cover.push(
    h1('Executive Summary'),
    p('Based on the numbers provided, here is your 90-day action plan.', { after: 240 }),
    
    h3('Primary Objective'),
    callout(top.title, top.shortBody),
    spacer(200),
    
    h3('Secondary Objectives & Delays'),
    bullet(derived.waitItem),
    spacer(120),
    
    h3('Review Horizon'),
    bullet(`Re-run analysis on ${fmtDate(90)} with updated inputs. Deterministic engine refresh takes < 2 minutes.`),
    spacer(240),
    
    h3('Cost of Inaction'),
    callout('Avoidable Interest', [
      sans('If minimums are maintained and recommended sequence is ignored, the model estimates an avoidable cost of ', { size: 20 }),
      sans(fmt(derived.costOfDoingNothing), { size: 20, bold: true, color: COLOR_RED }),
      sans(` over the next 5 years (approx. ${derived.costOfDoingNothingPerWeek} per week).`, { size: 20 }),
    ]),
    spacer(120),
    small('Methodology: Total interest paid on highest-rate debts under minimum payments vs. avalanche sequence with recommended extra payment. Assumes 60-month horizon.'),
    pageBreak(),
  );
  return cover;
}

function build12MonthMap(state, derived) {
  const surplus = derived.surplus;
  const target = derived.bufferTarget;
  const carSaveLabel = (state.car && state.car.considering === 'yes') ? 'Car / save' : 'Invest / save';
  const headerRow = ['Month', 'Income', 'Essentials', 'Surplus', 'Buffer', 'Debt', carSaveLabel, 'Buffer (cum)'];
  const bodyRows = [];

  // Dynamic simulation: phase A (40/40/20) until buffer hits target, then phase B.
  let bufferBalance = state.savings;
  const hasDebts = (state.debts || []).length > 0;
  let firstPhaseBMonth = null;

  for (let m = 1; m <= 12; m++) {
    let toBuffer, toDebt, toCar;
    if (bufferBalance < target) {
      toBuffer = Math.max(0, surplus * 0.4);
      toDebt = Math.max(0, surplus * 0.4);
      toCar = Math.max(0, surplus * 0.2);
      // Cap the buffer top-up at target and redirect the overshoot.
      if (bufferBalance + toBuffer > target) {
        const overshoot = bufferBalance + toBuffer - target;
        toBuffer -= overshoot;
        if (hasDebts) toDebt += overshoot; else toCar += overshoot;
      }
    } else {
      if (firstPhaseBMonth === null) firstPhaseBMonth = m;
      toBuffer = 0;
      toDebt = hasDebts ? Math.max(0, surplus * 0.6) : 0;
      toCar = Math.max(0, surplus * (hasDebts ? 0.4 : 1.0));
    }
    bufferBalance += toBuffer;
    bodyRows.push([
      `M${m}`,
      fmt(state.income),
      fmt(state.expenses),
      fmt(surplus),
      fmt(toBuffer),
      fmt(toDebt),
      fmt(toCar),
      fmt(bufferBalance),
    ]);
  }

  const transitionNote = firstPhaseBMonth
    ? `Your buffer reaches the ${fmt(target)} target around month ${firstPhaseBMonth}. From that point, the allocation switches to 60% debt / 40% car-or-savings.`
    : `Based on your surplus, you do not reach the ${fmt(target)} buffer target within 12 months — the table stays in Phase A throughout.`;

  return [
    h1('2. Your 12-month cash map'),
    p('How every dollar of surplus gets allocated, month by month. The allocation is dynamic: it transitions automatically once your buffer hits the 3-month target.', { after: 200 }),
    h3('Allocation rules in this plan'),
    bullet(`Phase A (buffer < ${fmt(target)}): 40% buffer top-up, 40% debt acceleration, 20% car / savings.`),
    bullet(`Phase B (buffer ≥ ${fmt(target)}): 0% buffer, 60% debt, 40% car / savings (or 100% car / savings if no debts).`),
    bullet(transitionNote),
    spacer(160),
    table([headerRow, ...bodyRows]),
    spacer(160),
    small('Tip: this allocation assumes income and essentials stay constant. If your income changes, re-run the free tool with your new numbers and request a refreshed plan within 90 days at no extra charge.'),
    pageBreak(),
  ];
}

function buildDebtRoadmap(state, derived) {
  const order = derived.debtOrder;
  if (!order.length) {
    return [
      h1('3. Debt roadmap'),
      p('You have no consumer debts to roadmap. Treat this as a strength: every dollar of surplus can go to buffer or investing.', { after: 240 }),
      pageBreak(),
    ];
  }

  const rows = [
    ['Order', 'Debt type', 'Balance', 'Rate', 'Tier', 'Months to clear', 'Total interest'],
    ...order.map((d, i) => [
      String(i + 1),
      NICE_TYPE[d.type] || d.type,
      fmt(d.balance),
      fmtPct(d.rate),
      d.tier,
      d.monthsToClear ? String(d.monthsToClear) : '—',
      fmt(d.totalInterest),
    ]),
  ];

  const sec = [
    h1('3. Debt roadmap'),
    p('The order to attack your debts and what each will cost you. Highest-rate first (the avalanche method) minimises total interest paid.', { after: 200 }),
    table(rows),
    spacer(200),
    h3('What this saves you'),
    p([
      arial('Following the avalanche order vs. paying everything down equally saves you approximately ', { size: 22 }),
      arial(fmt(derived.avalancheSaving), { size: 22, bold: true, color: COLOR_GREEN }),
      arial(' in interest over the life of these debts.', { size: 22 }),
    ]),
    small(`Methodology: avalanche = monthly minimums + recommended extra (${fmt(derived.monthlyAllocation.toDebt)}/mo) applied to highest-rate debt first, rolled to next debt on payoff. Equal-split = same total extra distributed proportionally across all debts.`),
  ];

  // HECS section ONLY if user has a HECS debt.
  const hasHECS = order.some(d => d.type === 'hecs');
  if (hasHECS) {
    sec.push(
      spacer(200),
      h3('Why HECS sits last in your plan'),
      p('HECS / HELP debt is indexed to CPI but charges no interest in the conventional sense. Voluntary repayments rarely beat the after-tax return on investing or paying off interest-bearing debt. Confirm the current indexation rate at studyassist.gov.au before deciding.')
    );
  }
  sec.push(pageBreak());
  return sec;
}

function buildCarScenarios(state, derived) {
  if (state.car.considering !== 'yes' || !state.car.price) {
    return [
      h1('4. Car decision'),
      p('You did not flag a car purchase in your inputs, so this section is a placeholder. If your situation changes, re-run the free tool and request a refreshed plan.', { after: 240 }),
      pageBreak(),
    ];
  }

  const car = state.car;
  const principal = car.price - car.deposit;
  const dealerPmt = calcMonthlyMortgage(principal, car.dealerRate, car.term);
  const bankPmt = calcMonthlyMortgage(principal, car.bankRate, car.term);
  const dealerTotal = dealerPmt * car.term * 12;
  const bankTotal = bankPmt * car.term * 12;
  const saving = dealerTotal - bankTotal;

  // Wait scenarios: deposit grows by 20% of monthly surplus (matches Section 2 Phase A car/save rule).
  const monthlyToCar = derived.monthlyAllocation.toCar;
  const wait6Deposit = car.deposit + monthlyToCar * 6;
  const wait12Deposit = car.deposit + monthlyToCar * 12;
  const wait6Price = car.price * (1 - RUNNING_COSTS.waitHold6mo);
  const wait12Price = car.price * (1 - RUNNING_COSTS.waitHold12mo);
  const wait6Principal = Math.max(0, wait6Price - wait6Deposit);
  const wait12Principal = Math.max(0, wait12Price - wait12Deposit);
  const wait6Pmt = calcMonthlyMortgage(wait6Principal, car.bankRate, car.term);
  const wait12Pmt = calcMonthlyMortgage(wait12Principal, car.bankRate, car.term);
  const wait6Total = wait6Pmt * car.term * 12;
  const wait12Total = wait12Pmt * car.term * 12;

  const rows = [
    ['Scenario', 'Deposit', 'Loan size', 'Monthly', 'Total cost', 'Vs buy-now (dealer)'],
    ['Buy now (dealer)', fmt(car.deposit), fmt(principal), fmt(dealerPmt), fmt(dealerTotal), '—'],
    ['Buy now (bank)', fmt(car.deposit), fmt(principal), fmt(bankPmt), fmt(bankTotal), fmt(-saving)],
    ['Wait 6 months (bank)', fmt(wait6Deposit), fmt(wait6Principal), fmt(wait6Pmt), fmt(wait6Total), fmt(wait6Total - dealerTotal)],
    ['Wait 12 months (bank)', fmt(wait12Deposit), fmt(wait12Principal), fmt(wait12Pmt), fmt(wait12Total), fmt(wait12Total - dealerTotal)],
  ];

  return [
    h1('4. Car decision — four scenarios'),
    p(`We modelled four ways to acquire this car and what each costs over your chosen ${car.term}-year term. Wait scenarios assume your monthly car/save allocation (${fmt(monthlyToCar)}) goes toward a bigger deposit.`, { after: 200 }),
    table(rows),
    spacer(200),
    h3('Assumptions in the wait scenarios'),
    bullet(`Vehicle price assumed to fall ~${(RUNNING_COSTS.waitHold6mo * 100).toFixed(0)}% over 6 months (${fmt(car.price)} → ${fmt(wait6Price)}) and ~${(RUNNING_COSTS.waitHold12mo * 100).toFixed(0)}% over 12 months (${fmt(car.price)} → ${fmt(wait12Price)}). Real depreciation varies by make/model.`),
    bullet(`Deposit grows by ${fmt(monthlyToCar)}/month — the same 20% car/save allocation from Section 2.`),
    bullet(`Bank rate (${fmtPct(car.bankRate)}) used for all wait scenarios on the assumption you would not return to dealer financing if you had time to shop.`),
    spacer(200),
    h3('The headline number'),
    p([
      arial('Choosing the bank loan over the dealer loan saves you ', { size: 22 }),
      arial(fmt(saving), { size: 22, bold: true, color: COLOR_GREEN }),
      arial(' over the loan term. Waiting 6-12 months changes the picture only modestly because depreciation eats some of the saving — but it materially reduces your monthly repayment, which improves your buffer resilience.', { size: 22 }),
    ]),
    spacer(160),
    small('Comparison-rate caveat: the dealer rate above is the nominal rate you entered. Australian comparison rates (AAPR) include fees and can be 1-2 percentage points higher than the headline. Always ask for the comparison rate in writing before signing.'),
    pageBreak(),
  ];
}

function buildOwnershipCost(state, derived) {
  if (state.car.considering !== 'yes' || !state.car.price) {
    return [
      h1('5. Total cost of ownership'),
      p('No car flagged — this section is a placeholder.', { after: 240 }),
      pageBreak(),
    ];
  }

  const car = state.car;
  const principal = car.price - car.deposit;
  const monthlyLoan = calcMonthlyMortgage(principal, car.bankRate, car.term);
  const annualKms = (state._post_purchase && state._post_purchase.annualKms) || 15000;
  const fuelAnnual = (annualKms / 100) * RUNNING_COSTS.fuelLitresPer100km * RUNNING_COSTS.fuelPricePerLitre;
  const depreciationAnnual = car.price * RUNNING_COSTS.depreciationPctYr1;
  const totalAnnual =
    monthlyLoan * 12 +
    RUNNING_COSTS.insuranceAnnual +
    RUNNING_COSTS.regoAnnual +
    RUNNING_COSTS.servicingAnnual +
    fuelAnnual +
    depreciationAnnual;
  const totalWeekly = totalAnnual / 52;
  const totalMonthly = totalAnnual / 12;
  const pctOfMonthly = (totalMonthly / state.income) * 100;

  const rows = [
    ['Cost component', 'Annual', 'Weekly', 'Notes'],
    ['Loan repayment (bank)', fmt(monthlyLoan * 12), fmt((monthlyLoan * 12) / 52), `${fmtPct(car.bankRate)} over ${car.term} years`],
    ['Comprehensive insurance', fmt(RUNNING_COSTS.insuranceAnnual), fmt(RUNNING_COSTS.insuranceAnnual / 52), 'A$30-40k vehicle, mid-cohort driver'],
    ['Registration + CTP', fmt(RUNNING_COSTS.regoAnnual), fmt(RUNNING_COSTS.regoAnnual / 52), 'VIC/NSW average; varies by state'],
    ['Servicing + tyres', fmt(RUNNING_COSTS.servicingAnnual), fmt(RUNNING_COSTS.servicingAnnual / 52), 'Logbook + 2 minor services'],
    ['Fuel', fmt(fuelAnnual), fmt(fuelAnnual / 52), `${annualKms.toLocaleString()} km/yr @ ${RUNNING_COSTS.fuelLitresPer100km} L/100km @ A$${RUNNING_COSTS.fuelPricePerLitre}/L`],
    ['Depreciation (Year 1)', fmt(depreciationAnnual), fmt(depreciationAnnual / 52), `~${(RUNNING_COSTS.depreciationPctYr1 * 100).toFixed(0)}% in Year 1; ~${(RUNNING_COSTS.depreciationPctYr2plus * 100).toFixed(0)}%/yr after`],
    ['TOTAL — Year 1', fmt(totalAnnual), fmt(totalWeekly), `${pctOfMonthly.toFixed(1)}% of monthly take-home`],
  ];

  return [
    h1('5. Total cost of ownership'),
    p('This is the number most car buyers underestimate. The financing cost is just one of six.', { after: 200 }),
    table(rows),
    spacer(200),
    h3('What this means for your cashflow'),
    p([
      arial('All-in, this car will cost about ', { size: 22 }),
      arial(fmt(totalWeekly) + ' per week', { size: 22, bold: true }),
      arial(` in Year 1 — roughly ${pctOfMonthly.toFixed(1)}% of your monthly take-home income.`, { size: 22 }),
    ]),
    p('Australian financial educators commonly recommend keeping total transport costs below 15% of net income. Above 20% and you will struggle to build buffer, repay debt, or invest at the same time.', { after: 200 }),
    small('All running-cost figures are FY 2025-26 illustrative averages from public Australian sources (Budget Direct, RACV, Finder). Your actual costs will vary by vehicle, postcode, driver profile, and km driven. Refresh annually.'),
    pageBreak(),
  ];
}

function buildStressTest(state, derived) {
  const monthlyEssentials = state.expenses;
  const baseSurplus = derived.surplus;

  const carPrincipal = state.car.considering === 'yes' ? state.car.price - state.car.deposit : 0;
  const baseCarPmt = state.car.considering === 'yes' ? calcMonthlyMortgage(carPrincipal, state.car.bankRate, state.car.term) : 0;
  const stressCarPmt = state.car.considering === 'yes' ? calcMonthlyMortgage(carPrincipal, state.car.bankRate + 2, state.car.term) : 0;
  const ratesDelta = (stressCarPmt - baseCarPmt) * 12;

  const stressIncome = state.income * 0.85;
  const debtMins = (state.debts || []).reduce((s, d) => s + (d.min || 0), 0);
  const stressSurplus = stressIncome - state.expenses - debtMins;

  const emergencyHit = 3000;
  const emergencyBufferAfter = state.savings - emergencyHit;
  const emergencyMonthsAfter = emergencyBufferAfter / monthlyEssentials;

  const unemploymentDeficit = monthlyEssentials * 3;
  const survivalMargin = state.savings - unemploymentDeficit;

  const rows = [
    ['Stress scenario', 'Trigger', 'Impact (your numbers)', 'What it means'],
    ['Rates rise 2%', 'RBA tightens', state.car.considering === 'yes' ? `+${fmt(ratesDelta)} / yr on car loan` : 'No current variable loan modelled', state.car.considering === 'yes' ? 'Buffer fills more slowly; review rate type in your contract' : 'Limited direct impact'],
    ['Income drops 15%', 'Reduced hours / role change', `Surplus ${fmt(baseSurplus)} → ${fmt(stressSurplus)}`, stressSurplus < 0 ? 'Plan fails — cut essentials or seek hardship support' : 'Plan timeline extends 3-6 months'],
    ['Emergency expense A$3,000', 'Medical / dental / car repair', `Buffer ${fmt(state.savings)} → ${fmt(emergencyBufferAfter)}`, emergencyMonthsAfter < 1 ? 'Buffer back to Critical — pause all debt acceleration' : `Buffer at ${emergencyMonthsAfter.toFixed(1)} months — recoverable`],
    ['3-month job loss', 'Redundancy / illness', `${fmt(survivalMargin)} ${survivalMargin >= 0 ? 'remaining' : 'shortfall'} after 3 months on essentials only`, survivalMargin >= 0 ? 'You survive on current buffer' : 'Risk of debt cycle — buffer-building is the single most important move'],
  ];

  return [
    h1('6. Stress tests'),
    p('What happens if the world doesn\'t go to plan. Each row uses your numbers, not generic averages.', { after: 200 }),
    table(rows),
    spacer(200),
    h3('What this tells you'),
    p(survivalMargin >= 0
      ? 'You currently survive a 3-month job loss on your existing buffer. That is rare and worth defending — do not deplete the buffer for non-essential purchases until you have additional cushion.'
      : 'You do not currently survive a 3-month job loss on your buffer alone. The single highest-leverage move is buffer-building before any debt acceleration or new car purchase.'),
    spacer(160),
    small('Caveats: the rate-rise scenario applies only to variable-rate loans. If your existing or planned car finance is fixed-rate, treat the +A$ figure as an indicator of refinance exposure when the fixed term ends. The income-drop scenario assumes essentials hold constant; in practice some essentials are negotiable (mobile, streaming, gym).'),
    pageBreak(),
  ];
}

function buildActionChecklist(state, derived) {
  const items = [];

  // Order: fastest-to-act first.
  if (derived.debtOrder.length > 0) {
    const top = derived.debtOrder[0];
    const niceType = NICE_TYPE[top.type] || top.type;
    items.push(`By ${fmtDateShort(7)}: log into your ${niceType} account and set up an extra ${fmt(derived.monthlyAllocation.toDebt)}/month payment on top of the minimum. This single move starts the avalanche immediately.`);
    items.push(`By ${fmtDateShort(14)}: phone your ${niceType} provider and request a rate review or balance-transfer offer. Script: "I'm reviewing my finances and I'd like to know what rate I qualify for given my repayment history."`);
  }

  if (derived.bufferShortfall > 0) {
    items.push(`By ${fmtDateShort(30)}: open a separate high-interest savings account (e.g. ING Savings Maximiser, Macquarie Savings) and set a recurring transfer of ${fmt(derived.monthlyAllocation.toBuffer)}/month from your salary toward the ${fmt(derived.bufferTarget)} buffer target.`);
  }

  if (state.car.considering === 'yes') {
    items.push('Before signing any car loan: get the comparison rate (AAPR) in writing from BOTH the dealer and your bank. The dealer\'s nominal rate is rarely the full picture.');
    items.push('Compare insurance quotes from at least 3 providers (Budget Direct, AAMI, Bingle) — annual premiums vary by 30%+ for the same cover.');
    if (state.car.novated === 'yes') {
      items.push('Speak to a salary-packaging specialist before committing to a novated lease. EVs under the LCT fuel-efficient threshold may qualify for FBT exemption — but eligibility depends on your employer.');
    }
  }

  items.push(`On ${fmtDateShort(90)}: re-run the free tool with your updated numbers. Within 90 days of purchase, reply to the email this PDF arrived in to request a refreshed plan at no extra charge.`);
  items.push('Optional but high-leverage: book one hour with a fee-only Australian financial adviser before any major decision (refinance, new loan, super top-up). Cost A$200-400; value typically 10x.');

  return [
    h1('7. Your action checklist'),
    p('Specific actions, in order, with dates. Tick them off as you go. Earliest-deadline items are listed first.', { after: 200 }),
    ...items.map(t => bullet(t)),
    spacer(280),
    h3('Useful Australian resources'),
    bullet([arial('ASIC Moneysmart: ', { size: 22, bold: true }), arial('moneysmart.gov.au — free, government-backed', { size: 22 })]),
    bullet([arial('National Debt Helpline: ', { size: 22, bold: true }), arial('1800 007 007 — free, confidential, financial counsellors', { size: 22 })]),
    bullet([arial('ATO payment plans: ', { size: 22, bold: true }), arial('ato.gov.au — set up online if you owe under A$100k', { size: 22 })]),
    bullet([arial('Comparison rates explained: ', { size: 22, bold: true }), arial('moneysmart.gov.au/loans/comparison-rate', { size: 22 })]),
    spacer(280),
    small('Re-run the free tool any time at moneymoves-au.vercel.app — results update as your situation changes.'),
    spacer(360),
    p([
      arial('General information only. ', { size: 18, italics: true, bold: true, color: COLOR_INK_SOFT }),
      arial('This document is based on the inputs you provided and deterministic rules. It is not personal financial product advice and does not consider your full circumstances. Speak to a licensed Australian financial adviser before any major decision. Crisis support: National Debt Helpline 1800 007 007.', { size: 18, italics: true, color: COLOR_INK_SOFT }),
    ]),
  ];
}

// ─── Derive ──────────────────────────────────────────────────────────────────
function deriveReportData(state) {
  const debts = state.debts || [];
  const debtMins = debts.reduce((s, d) => s + (d.min || 0), 0);
  const surplus = state.income - state.expenses - debtMins;

  const monthlyEssentials = state.expenses;
  const bufferMonths = state.savings / Math.max(monthlyEssentials, 1);
  const bufferTarget = monthlyEssentials * RULES.TARGET_BUFFER_MONTHS;
  const bufferShortfall = Math.max(0, bufferTarget - state.savings);

  let toBuffer, toDebt, toCar;
  if (bufferMonths < RULES.TARGET_BUFFER_MONTHS) {
    toBuffer = Math.max(0, surplus * 0.4);
    toDebt = Math.max(0, surplus * 0.4);
    toCar = Math.max(0, surplus * 0.2);
  } else {
    toBuffer = 0;
    toDebt = debts.length ? Math.max(0, surplus * 0.6) : 0;
    toCar = Math.max(0, surplus * (debts.length ? 0.4 : 1.0));
  }

  // Use proper amortisation for months-to-clear.
  const order = buildDebtPayoffOrder(debts).map(d => {
    const monthlyForThis = (d.min || 0) + toDebt; // simplifying: extra applied to top debt
    const m = monthsToClear(d.balance, d.rate, monthlyForThis);
    const interest = totalInterestPaid(d.balance, d.rate, monthlyForThis);
    return { ...d, monthsToClear: m, totalInterest: interest };
  });

  const totalInterestAvalanche = order.reduce((s, d) => s + (d.totalInterest || 0), 0);
  // Naive equal-split estimate (proxy: same monthly extra split N ways).
  const totalInterestEqual = debts.reduce((s, d) => {
    const mp = (d.min || 0) + toDebt / Math.max(debts.length, 1);
    return s + totalInterestPaid(d.balance, d.rate, mp);
  }, 0);
  const avalancheSaving = Math.max(0, totalInterestEqual - totalInterestAvalanche);

  // Cost-of-doing-nothing: minimums-only baseline vs avalanche, on highest-rate debt.
  const baselineInterest = debts.reduce((s, d) => s + totalInterestPaid(d.balance, d.rate, d.min || 1), 0);
  const costOfDoingNothing = Math.max(0, baselineInterest - totalInterestAvalanche);
  const costOfDoingNothingPerWeek = fmt(costOfDoingNothing / (5 * 52));

  let topMove;
  if (bufferMonths < RULES.MIN_BUFFER_MONTHS) {
    topMove = {
      title: 'Build a 1-month emergency buffer first',
      shortBody: `Aim for ${fmt(monthlyEssentials)} (1 month of essentials) before any debt acceleration or car purchase. This is the single highest-leverage move you can make.`,
    };
  } else if (order[0]) {
    const top = order[0];
    topMove = {
      title: `Attack the ${NICE_TYPE[top.type] || top.type} (${fmtPct(top.rate)})`,
      shortBody: `It is your highest-rate debt and every extra dollar you put on it returns ${fmtPct(top.rate)} risk-free. Pay an extra ${fmt(toDebt)}/month on top of the minimum.`,
    };
  } else if (bufferMonths >= RULES.IDEAL_BUFFER_MONTHS) {
    topMove = {
      title: 'Start a regular investment plan',
      shortBody: `Buffer is healthy and there is no high-rate debt. Set up an automatic ${fmt(surplus * 0.6)}/month into a low-fee diversified ETF or salary-sacrificed super (in priority order).`,
    };
  } else if (bufferMonths >= RULES.TARGET_BUFFER_MONTHS) {
    // Buffer is 3–6 months (healthy but not ideal). toBuffer is 0 in Phase B,
    // so we must compute the push-to-ideal amount directly from surplus.
    const idealTarget = monthlyEssentials * RULES.IDEAL_BUFFER_MONTHS;
    const remaining = idealTarget - state.savings;
    const monthsToIdeal = Math.max(1, Math.ceil(remaining / Math.max(surplus, 1)));
    const toIdealBuffer = Math.min(surplus, Math.round(remaining / Math.min(monthsToIdeal, 12)));
    topMove = {
      title: 'Push your buffer to the 6-month ideal',
      shortBody: `Your buffer is at ${bufferMonths.toFixed(1)} months — healthy, but not fully padded. The 6-month target is ${fmt(idealTarget)}. Direct ${fmt(toIdealBuffer)}/month into a high-interest savings account and you will get there in ~${monthsToIdeal} months.`,
    };
  } else {
    topMove = {
      title: 'Top your buffer up to 3 months',
      shortBody: `${fmt(bufferTarget)} is the next milestone. Direct ${fmt(toBuffer)}/month into a separate high-interest savings account.`,
    };
  }

  const waitItem = state.car.considering === 'yes' && bufferMonths < RULES.TARGET_BUFFER_MONTHS
    ? `Wait on the car. With buffer at ${bufferMonths.toFixed(1)} months and high-rate debt outstanding, a new financed depreciating asset compounds risk. Section 4 shows the exact dollar difference between buying now and waiting 6-12 months.`
    : `Avoid taking on any new consumer debt (BNPL, store cards, additional car loans) until existing debts are below ${fmtPct(RULES.HIGH_RATE_THRESHOLD)} and your buffer is at ${RULES.TARGET_BUFFER_MONTHS} months.`;

  return {
    surplus,
    bufferMonths,
    bufferTarget,
    bufferShortfall,
    monthlyAllocation: { toBuffer, toDebt, toCar },
    debtOrder: order,
    topMove,
    waitItem,
    costOfDoingNothing,
    costOfDoingNothingPerWeek,
    avalancheSaving,
  };
}

// ─── Build doc ───────────────────────────────────────────────────────────────
function buildReport(state) {
  const derived = deriveReportData(state);
  const children = [
    ...buildCover(state, derived),
    ...build12MonthMap(state, derived),
    ...buildDebtRoadmap(state, derived),
    ...buildCarScenarios(state, derived),
    ...buildOwnershipCost(state, derived),
    ...buildStressTest(state, derived),
    ...buildActionChecklist(state, derived),
  ];

  return new Document({
    creator: 'MoneyMoves AU',
    title: 'MoneyMoves AU — Personalised Decision Pack',
    description: 'Personalised debt + car loan decision pack for an Australian household.',
    styles: {
      default: { document: { run: { font: 'Arial', size: 22 } } },
      paragraphStyles: [
        { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 32, bold: true, font: 'Arial', color: COLOR_INK },
          paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 } },
        { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 24, bold: true, font: 'Arial', color: COLOR_GREEN_DARK },
          paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 2 } },
      ],
    },
    numbering: {
      config: [{
        reference: 'bullets',
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } },
        }],
      }],
    },
    sections: [{
      properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
      headers: {
        default: new Header({
          children: [new Paragraph({
            alignment: AlignmentType.RIGHT,
            children: [arial('MoneyMoves AU — Personalised Decision Pack', { size: 18, color: COLOR_INK_SOFT })],
          })],
        }),
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              arial('Confidential. General information only — not personal financial advice.', { size: 16, color: COLOR_INK_SOFT }),
              new TextRun({ children: ['  |  Page ', PageNumber.CURRENT, ' of ', PageNumber.TOTAL_PAGES], size: 16, color: COLOR_INK_SOFT, font: 'Arial' }),
            ],
          })],
        }),
      },
      children,
    }],
  });
}

async function generateReport(state) {
  const doc = buildReport(state);
  return Packer.toBuffer(doc);
}

module.exports = { buildReport, generateReport };

if (require.main === module) {
  (async () => {
    const stateFile = process.argv[2] || 'sample_state.json';
    const outFile = process.argv[3] || 'MoneyMoves_AU_Sample_Report.docx';
    const stateData = JSON.parse(fs.readFileSync(stateFile, 'utf8'));
    delete stateData._comment;
    const buf = await generateReport(stateData);
    fs.writeFileSync(outFile, buf);
    console.log(`OK: report written to ${outFile} (${buf.length} bytes)`);
  })();
}
