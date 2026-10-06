/**
 * MoneyMoves AU — Personalised Money Plan (PDF renderer, v0.3)
 * ---------------------------------------------------------
 * Takes the app's `state` object and renders a Word document (converted to PDF
 * by the API routes).
 *
 * Every number in the report comes from ONE month-by-month simulation
 * (`simulatePlan`) that follows the same priority order the plan recommends:
 *   1. fill a 1-month emergency buffer
 *   2. clear high-rate (≥10%) and ATO debt, highest rate first
 *   3. fill a 3-month buffer
 *   4. clear the remaining consumer debt, highest rate first
 *   5. everything else → savings (or a car deposit when a car is planned)
 * Interest accrues monthly; a cleared debt's minimum rolls into the next step.
 * HECS/HELP is repaid through tax and home loans run on their own schedule, so
 * neither receives extra repayments. Because the cash map, debt dates, interest
 * figures, cost of inaction and checklist all read the same simulation, they
 * cannot contradict each other.
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

const SIM_MONTHS = 360;           // long enough to find payoff dates for most debts
const COST_OF_INACTION_MONTHS = 60;
const HECS_INDEXATION_PCT = 2.8;  // June 2026 indexation (lower of CPI / WPI) — same figure the app shows
const RATE_RISE_PCT = 2;
// Debt types whose rate normally moves with the market.
const VARIABLE_RATE_TYPES = ['credit_card', 'personal_loan', 'home_loan', 'other'];
// Never accelerated: HECS is repaid through tax; home loans are better served by an offset account.
const NO_EXTRA_REPAYMENT_TYPES = ['hecs', 'home_loan'];

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

const REPORT_NAME = 'Personalised Money Plan';

const COLOR_NAVY = '0A192F';
const COLOR_NAVY_DARK = '020C1B';
const COLOR_INK = '112233';
const COLOR_INK_SOFT = '556677';
const COLOR_RED = 'D64545';
const COLOR_LINE = 'D0D7DE';
const COLOR_CREAM = 'F8F9FA';
const COLOR_GOLD_SOFT = 'F0F4F8';
const COLOR_ACCENT = '005A9C';

// ─── Helpers ─────────────────────────────────────────────────────────────────
const fmt = (n) => {
  if (n === null || n === undefined || isNaN(n)) return '—';
  const rounded = Math.round(n);
  const sign = rounded < 0 ? '-' : '';
  return sign + 'A$' + Math.abs(rounded).toLocaleString('en-AU');
};
// Rates are shown exactly as entered (20.99%, 9.5%), never rounded to 21.0%.
const fmtPct = (n) => `${Number(n.toFixed(2))}%`;

const NICE_TYPE = {
  credit_card: 'credit card',
  personal_loan: 'personal loan',
  car_loan: 'car loan',
  bnpl: 'buy-now-pay-later',
  ato: 'ATO debt',
  hecs: 'HECS / HELP',
  home_loan: 'home loan',
  other: 'other debt',
};
const niceType = (t) => NICE_TYPE[t] || t;
const cap = (s) => s.charAt(0).toUpperCase() + s.slice(1);

function calcMonthlyMortgage(principal, annualRate, years) {
  if (principal <= 0 || years <= 0) return 0;
  if (annualRate <= 0) return principal / (years * 12);
  const r = annualRate / 100 / 12;
  const n = years * 12;
  return principal * r * Math.pow(1 + r, n) / (Math.pow(1 + r, n) - 1);
}

function buildDebtPayoffOrder(debts) {
  const byRate = (a, b) => b.rate - a.rate;
  const isConsumer = d => !['hecs', 'home_loan', 'ato'].includes(d.type);
  const high = debts.filter(d => isConsumer(d) && d.rate >= RULES.HIGH_RATE_THRESHOLD).sort(byRate);
  const ato  = debts.filter(d => d.type === 'ato');
  const mid  = debts.filter(d => isConsumer(d) && d.rate < RULES.HIGH_RATE_THRESHOLD && d.rate >= 5).sort(byRate);
  const low  = debts.filter(d => isConsumer(d) && d.rate < 5).sort(byRate);
  const home = debts.filter(d => d.type === 'home_loan').sort(byRate);
  const hecs = debts.filter(d => d.type === 'hecs');
  return [
    ...high.map(d => ({ ...d, tier: 'High-rate', priority: true })),
    ...ato.map(d => ({ ...d, tier: 'ATO (enforcement risk)', priority: true })),
    ...mid.map(d => ({ ...d, tier: 'Mid-rate', priority: false })),
    ...low.map(d => ({ ...d, tier: 'Low-rate', priority: false })),
    ...home.map(d => ({ ...d, tier: 'Home loan', priority: false })),
    ...hecs.map(d => ({ ...d, tier: 'HECS / HELP', priority: false })),
  ];
}

// ─── Dates ───────────────────────────────────────────────────────────────────
const TODAY = new Date();
// Fixed three-letter months: the en-AU locale mixes "Dec" with "June" and "Sept".
const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const fmtDate = (offsetDays) => {
  const d = new Date(TODAY.getTime() + offsetDays * 86400000);
  return `${d.getDate()} ${MONTHS[d.getMonth()]} ${d.getFullYear()}`;
};
// Plan month 1 is the next calendar month.
const monthLabel = (m) => {
  if (m === null || m === undefined) return '—';
  if (m === 0) return 'Already there';
  const d = new Date(TODAY.getFullYear(), TODAY.getMonth() + m, 1);
  return `${MONTHS[d.getMonth()]} ${d.getFullYear()}`;
};

// A name typed all in lower case gets each word capitalised; anything else is left
// exactly as typed, so names like "McDonald" or "de Silva" keep their own casing.
function displayName(raw) {
  const name = String(raw || '').trim().replace(/\s+/g, ' ');
  if (!name || name !== name.toLowerCase()) return name;
  return name.replace(/(^|[\s'-])(\p{L})/gu, (_, sep, ch) => sep + ch.toUpperCase());
}

// ─── The simulation ──────────────────────────────────────────────────────────
/**
 * strategy:
 *   'plan'     — the staged plan described at the top of this file
 *   'equal'    — same stages, but each month's extra debt money is split evenly
 *                across the open debts (used only to show what avalanche saves)
 *   'minimums' — minimum repayments only; surplus is never used on debt
 */
function simulatePlan(state, { strategy = 'plan', months = SIM_MONTHS } = {}) {
  const ordered = buildDebtPayoffOrder(state.debts || []);
  const debts = ordered.map(d => ({
    ...d,
    bal: Math.max(0, d.balance),
    interestPaid: 0,
    clearedMonth: d.balance <= 0 ? 0 : null,
  }));
  const allMins = debts.reduce((s, d) => s + (d.min || 0), 0);
  const oneMonth = state.expenses + allMins;          // living costs + today's minimums
  const threeMonth = oneMonth * RULES.TARGET_BUFFER_MONTHS;
  const savingLabel = state.car && state.car.considering === 'yes' ? 'Car deposit / savings' : 'Savings / investing';

  let buffer = state.savings;
  let saved = 0;
  const milestones = {
    buffer1: buffer >= oneMonth ? 0 : null,
    buffer3: buffer >= threeMonth ? 0 : null,
    savingsStart: null,
  };
  const rows = [];

  for (let m = 1; m <= months; m++) {
    const row = { m, income: state.income, living: state.expenses, mins: 0, interest: 0,
      toBuffer: 0, fromBuffer: 0, toDebt: 0, toSave: 0, focus: [] };

    // 1. Interest, then minimum repayments (HECS is collected through tax, not from this budget).
    for (const d of debts) {
      if (d.type === 'hecs' || d.bal <= 0) continue;
      const interest = d.bal * d.rate / 1200;
      d.interestPaid += interest;
      row.interest += interest;
      d.bal += interest;
      const pay = Math.min(d.min || 0, d.bal);
      d.bal -= pay;
      row.mins += pay;
      if (d.bal <= 0.005) { d.bal = 0; if (d.clearedMonth === null) d.clearedMonth = m; }
    }

    let available = state.income - state.expenses - row.mins;
    row.surplus = available;

    // 2. A shortfall comes out of the buffer while it lasts.
    if (available < 0) {
      row.fromBuffer = Math.min(buffer, -available);
      buffer -= row.fromBuffer;
      row.focus.push('Covering a shortfall');
      available = 0;
    }

    const payExtra = (list) => {
      let open = list.filter(d => d.bal > 0);
      if (strategy === 'equal') {
        while (available > 0.005 && open.length) {
          const share = available / open.length;
          for (const d of open) {
            const p = Math.min(d.bal, share);
            d.bal -= p; available -= p; row.toDebt += p;
            if (d.bal <= 0.005) { d.bal = 0; if (d.clearedMonth === null) d.clearedMonth = m; }
          }
          open = open.filter(d => d.bal > 0);
        }
      } else {
        for (const d of open) {
          if (available <= 0.005) break;
          const p = Math.min(d.bal, available);
          d.bal -= p; available -= p; row.toDebt += p;
          if (d.bal <= 0.005) { d.bal = 0; if (d.clearedMonth === null) d.clearedMonth = m; }
        }
      }
    };
    const topUpBuffer = (target, label) => {
      if (available <= 0 || buffer >= target) return;
      const t = Math.min(available, target - buffer);
      buffer += t; available -= t; row.toBuffer += t;
      row.focus.push(label);
    };
    const accelerable = debts.filter(d => !NO_EXTRA_REPAYMENT_TYPES.includes(d.type));

    if (strategy !== 'minimums') {
      topUpBuffer(oneMonth, '1-month buffer');
      // 'equal' is the comparison case: no debt gets priority, so it skips straight to splitting.
      const priorityOpen = strategy === 'equal' ? [] : accelerable.filter(d => d.priority && d.bal > 0);
      if (available > 0 && priorityOpen.length) {
        row.focus.push(cap(niceType(priorityOpen[0].type)));
        payExtra(priorityOpen);
      }
      topUpBuffer(threeMonth, '3-month buffer');
      const restOpen = accelerable.filter(d => d.bal > 0);
      if (available > 0 && restOpen.length) {
        row.focus.push(cap(niceType(restOpen[0].type)));
        payExtra(restOpen);
      }
    }
    if (available > 0.005) {
      row.toSave = available; saved += available;
      row.focus.push(savingLabel);
      if (milestones.savingsStart === null) milestones.savingsStart = m;
    }

    if (milestones.buffer1 === null && buffer >= oneMonth - 0.5) milestones.buffer1 = m;
    if (milestones.buffer3 === null && buffer >= threeMonth - 0.5) milestones.buffer3 = m;
    row.buffer = buffer;
    row.saved = saved;
    row.debtLeft = debts.filter(d => d.type !== 'hecs').reduce((s, d) => s + d.bal, 0);
    rows.push(row);
  }

  const consumer = debts.filter(d => !NO_EXTRA_REPAYMENT_TYPES.includes(d.type));
  const consumerCleared = consumer.length && consumer.every(d => d.clearedMonth !== null)
    ? Math.max(...consumer.map(d => d.clearedMonth)) : null;
  return { rows, debts, milestones: { ...milestones, consumerDebtFree: consumerCleared }, oneMonth, threeMonth };
}

const interestOver = (sim, months) => sim.rows.slice(0, months).reduce((s, r) => s + r.interest, 0);

// ─── Derive everything the sections need ─────────────────────────────────────
function deriveReportData(state) {
  const debts = state.debts || [];
  const debtMins = debts.reduce((s, d) => s + (d.min || 0), 0);
  const surplus = state.income - state.expenses - debtMins;
  const monthlyEssentials = state.expenses + debtMins;
  const bufferMonths = state.savings / Math.max(monthlyEssentials, 1);

  const plan = simulatePlan(state, { strategy: 'plan' });
  const minimumsOnly = simulatePlan(state, { strategy: 'minimums' });

  // What avalanche saves only means something with 2+ interest-bearing debts to choose between.
  const interestBearing = debts.filter(d => d.rate > 0 && !NO_EXTRA_REPAYMENT_TYPES.includes(d.type));
  let avalancheSaving = null;
  if (interestBearing.length >= 2) {
    const equal = simulatePlan(state, { strategy: 'equal' });
    avalancheSaving = Math.max(0, interestOver(equal, SIM_MONTHS) - interestOver(plan, SIM_MONTHS));
  }

  const planInterest5y = interestOver(plan, COST_OF_INACTION_MONTHS);
  const baseInterest5y = interestOver(minimumsOnly, COST_OF_INACTION_MONTHS);
  const costOfInaction = Math.max(0, baseInterest5y - planInterest5y);

  // Per-debt results from both runs, in plan order.
  const debtOrder = plan.debts.map((d, i) => ({
    ...d,
    minimumsOnlyClearedMonth: minimumsOnly.debts[i].clearedMonth,
    minimumsOnlyInterest: minimumsOnly.debts[i].interestPaid,
  }));

  const firstRow = plan.rows[0];
  const firstPriority = debtOrder.find(d => d.priority && d.bal !== undefined && !NO_EXTRA_REPAYMENT_TYPES.includes(d.type));
  const hasCar = state.car && state.car.considering === 'yes' && state.car.price > 0;

  let topMove;
  if (surplus < 100) {
    topMove = surplus < 0
      ? { title: "Close the gap: you're spending more than you earn",
          shortBody: `Take-home ${fmt(state.income)} − living costs ${fmt(state.expenses)} − debt minimums ${fmt(debtMins)} = ${fmt(surplus)} a month. No plan works until this is positive: cut a fixed cost (insurance, subscriptions, refinancing) or lift income.` }
      : { title: 'Create a monthly surplus first',
          shortBody: `After living costs and minimums you have ${fmt(surplus)} a month left. That isn't enough for any other step to make progress. Cut one fixed cost or lift income first.` };
  } else if (bufferMonths < RULES.MIN_BUFFER_MONTHS) {
    const then = firstPriority ? ` Then switch the same amount to your ${niceType(firstPriority.type)}.` : '';
    topMove = {
      title: `Build a 1-month emergency buffer (${fmt(plan.oneMonth)})`,
      shortBody: `Put your whole surplus (${fmt(surplus)}/month) into a separate high-interest savings account until it reaches ${fmt(plan.oneMonth)}, which on these numbers happens in ${monthLabel(plan.milestones.buffer1)}.${then} A flat tyre or vet bill on a credit card costs more than any interest you'd save by paying debt first.`,
    };
  } else if (firstPriority) {
    topMove = {
      title: `Pay off your ${niceType(firstPriority.type)} (${fmtPct(firstPriority.rate)})`,
      shortBody: `It is your most expensive debt. Pay your whole surplus (${fmt(surplus)}/month) on top of the minimum and it is gone by ${monthLabel(firstPriority.clearedMonth)}.`,
    };
  } else if (bufferMonths < RULES.TARGET_BUFFER_MONTHS) {
    topMove = {
      title: `Top your buffer up to 3 months (${fmt(plan.threeMonth)})`,
      shortBody: `You have no high-rate debt, so the next step is a 3-month buffer. At ${fmt(surplus)}/month you reach it in ${monthLabel(plan.milestones.buffer3)}.`,
    };
  } else if (debtOrder.some(d => !NO_EXTRA_REPAYMENT_TYPES.includes(d.type) && d.balance > 0)) {
    const next = debtOrder.find(d => !NO_EXTRA_REPAYMENT_TYPES.includes(d.type) && d.balance > 0);
    topMove = {
      title: `Make extra repayments on your ${niceType(next.type)} (${fmtPct(next.rate)})`,
      shortBody: `Your buffer is healthy. Putting ${fmt(surplus)}/month on top of the minimum clears it by ${monthLabel(next.clearedMonth)}. Check for early-repayment fees first.`,
    };
  } else {
    topMove = {
      title: 'Start a regular investment plan',
      shortBody: `Buffer is healthy and there is no consumer debt. Automate ${fmt(surplus)}/month into a low-fee diversified ETF or salary-sacrificed super, in that order of convenience for you.`,
    };
  }

  const waitItem = hasCar && bufferMonths < RULES.TARGET_BUFFER_MONTHS
    ? `Wait on the car until your buffer is at 3 months (${monthLabel(plan.milestones.buffer3)} on these numbers). The car section shows the dollar difference between buying now and waiting.`
    : `Avoid any new consumer debt (buy-now-pay-later, store cards, a financed car) until your high-rate debt is gone and your buffer reaches 3 months.`;

  return {
    surplus,
    debtMins,
    monthlyEssentials,
    bufferMonths,
    plan,
    minimumsOnly,
    firstRow,
    debtOrder,
    hasCar,
    topMove,
    waitItem,
    costOfInaction,
    planInterest5y,
    baseInterest5y,
    avalancheSaving,
  };
}

// ─── docx primitives ─────────────────────────────────────────────────────────
const sans = (text, opts = {}) => new TextRun({ text, font: 'Arial', ...opts });
const serif = (text, opts = {}) => new TextRun({ text, font: 'Georgia', ...opts });
const arial = sans;

const p = (text, opts = {}) => new Paragraph({
  children: Array.isArray(text) ? text : [sans(text, { size: opts.size || 20, color: opts.color || COLOR_INK, bold: !!opts.bold })],
  alignment: opts.align || AlignmentType.LEFT,
  heading: opts.heading,
  keepNext: opts.keepNext,
  spacing: { after: opts.after ?? 120, before: opts.before ?? 0 },
});
const h1 = (t) => new Paragraph({
  children: [serif(t, { size: 36, color: COLOR_NAVY, bold: false })],
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
// Each numbered list passes its own `instance` so numbering restarts at 1.
const numbered = (t, instance) => new Paragraph({
  children: Array.isArray(t) ? t : [sans(t, { size: 20, color: COLOR_INK })],
  numbering: { reference: 'steps', level: 0, instance },
  spacing: { after: 100 },
});

const callout = (title, bodyText) => new Table({
  rows: [new TableRow({ children: [
    new TableCell({
      children: [
        new Paragraph({ children: [sans(title, { size: 22, bold: true, color: COLOR_NAVY })], spacing: { after: 80 } }),
        new Paragraph({ children: Array.isArray(bodyText) ? bodyText : [sans(bodyText, { size: 20, color: COLOR_INK })] }),
      ],
      shading: { type: ShadingType.CLEAR, color: 'auto', fill: COLOR_GOLD_SOFT },
      margins: { top: 200, bottom: 200, left: 200, right: 200 },
      borders: {
        left: { style: BorderStyle.SINGLE, size: 16, color: COLOR_ACCENT },
        top: { style: BorderStyle.NONE, size: 0, color: 'auto' },
        right: { style: BorderStyle.NONE, size: 0, color: 'auto' },
        bottom: { style: BorderStyle.NONE, size: 0, color: 'auto' },
      },
    }),
  ] })],
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
    margins: { top: 100, bottom: 100, left: 100, right: 100 },
  });
}

function table(rows, opts = {}) {
  const size = opts.size;
  return new Table({
    rows: rows.map((cells, idx) => new TableRow({
      children: cells.map(c =>
        idx === 0
          ? cell(c, { bold: true, color: COLOR_NAVY, fill: 'FFFFFF', size })
          : cell(c, { fill: idx % 2 === 0 ? COLOR_CREAM : 'FFFFFF', size })
      ),
      tableHeader: idx === 0,
    })),
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 12, color: COLOR_NAVY },
      bottom: { style: BorderStyle.SINGLE, size: 12, color: COLOR_NAVY },
      left: { style: BorderStyle.NONE, size: 0, color: 'auto' },
      right: { style: BorderStyle.NONE, size: 0, color: 'auto' },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 2, color: COLOR_LINE },
      insideVertical: { style: BorderStyle.NONE, size: 0, color: 'auto' },
    },
  });
}

const pageBreak = () => new Paragraph({ children: [new PageBreak()] });

// Debt-row helpers shared by several sections.
const clearedText = (d) => {
  if (d.type === 'hecs') return 'Through tax';
  if (d.clearedMonth === null) return 'After 30+ years';
  return monthLabel(d.clearedMonth);
};
const rateText = (d) => d.type === 'hecs' ? `Indexed (~${HECS_INDEXATION_PCT}%)` : fmtPct(d.rate);

// ─── Page 1: cover + plan at a glance ────────────────────────────────────────
function buildThisWeek(state, derived) {
  const items = [];
  const firstPriority = derived.debtOrder.find(d => d.priority);
  if (derived.surplus < 100) {
    items.push('List every fixed cost (insurance, phone, subscriptions, loans) and cancel or renegotiate at least one this week.');
    items.push('Call the National Debt Helpline on 1800 007 007 if minimum repayments are hard to meet. It is free and confidential.');
  } else if (derived.bufferMonths < RULES.MIN_BUFFER_MONTHS) {
    items.push(`Open a separate high-interest savings account and set an automatic transfer of ${fmt(derived.surplus)} each month (or the equivalent each payday).`);
  } else if (firstPriority) {
    items.push(`Set up an automatic extra payment of ${fmt(derived.surplus)}/month on your ${niceType(firstPriority.type)}, on top of the minimum.`);
  } else {
    items.push(`Set up an automatic transfer of ${fmt(derived.surplus > 0 ? derived.surplus : 0)}/month for the step shown above.`);
  }
  if (firstPriority && firstPriority.rate > 0) {
    items.push(`Call your ${niceType(firstPriority.type)} provider and ask for a lower rate or a balance-transfer offer (script in the toolkit at the end of this plan).`);
  }
  if (derived.debtMins > 0) {
    items.push(`Keep every minimum repayment on autopay (${fmt(derived.debtMins)}/month in total). A missed payment costs more in fees and interest than this plan saves.`);
  }
  items.push(`Put ${fmtDate(90)} in your calendar to re-run the free tool with updated numbers.`);
  return items.slice(0, 4);
}

function buildCover(state, derived) {
  const plan = derived.plan;
  const who = displayName(state.name) || state.email || 'You';

  const keyDates = [['Milestone', 'When', 'Detail']];
  keyDates.push(['1-month emergency buffer', monthLabel(plan.milestones.buffer1), fmt(plan.oneMonth)]);
  for (const d of derived.debtOrder) {
    if (d.type === 'hecs') {
      keyDates.push(['HECS / HELP', 'Through tax', `${fmt(d.balance)}, no extra repayments needed`]);
    } else {
      keyDates.push([`${cap(niceType(d.type))} cleared`, clearedText(d), `${fmt(d.balance)} at ${fmtPct(d.rate)}`]);
    }
  }
  keyDates.push(['3-month emergency buffer', monthLabel(plan.milestones.buffer3), fmt(plan.threeMonth)]);

  const out = [
    new Paragraph({ children: [sans('CONFIDENTIAL', { size: 16, color: COLOR_ACCENT, bold: true })], alignment: AlignmentType.RIGHT }),
    new Paragraph({ children: [serif('Your Money Plan', { size: 52, color: COLOR_NAVY })], spacing: { before: 200, after: 60 } }),
    p([sans(`Prepared for ${who}  ·  ${fmtDate(0)}`, { size: 20, color: COLOR_INK_SOFT })], { after: 280 }),

    h3('Your first move'),
    callout(derived.topMove.title, derived.topMove.shortBody),
    spacer(160),

    h3('Key dates on your numbers'),
    table(keyDates),
    small('Dates assume your income and costs stay as entered and you follow the steps in order. "Through tax" means HECS is deducted from your pay by the ATO.'),
    spacer(120),

    h3('Do these this week'),
    ...buildThisWeek(state, derived).map(t => numbered(t, 1)),
    pageBreak(),
  ];
  return out;
}

// ─── How the plan works + cost of inaction ───────────────────────────────────
function buildPlanLogic(state, derived, n) {
  const sec = [
    h1(`${n}. How your plan works`),
    p(`Every month: take-home ${fmt(state.income)} − living costs ${fmt(state.expenses)} − debt minimums ${fmt(derived.debtMins)} = `, { after: 0 }),
    p([sans(`${fmt(derived.surplus)} a month to put to work.`, { size: 24, bold: true, color: COLOR_NAVY })], { after: 200 }),
    h3('The order your surplus goes in'),
    numbered(`Emergency buffer up to 1 month of costs (${fmt(derived.plan.oneMonth)}). This stops a surprise bill landing on a credit card.`, 2),
    numbered(`High-rate debt (${RULES.HIGH_RATE_THRESHOLD}%+) and ATO debt, most expensive first. Every dollar here earns that interest rate back, guaranteed.`, 2),
    numbered(`Emergency buffer up to 3 months (${fmt(derived.plan.threeMonth)}).`, 2),
    numbered('Any remaining consumer debt, most expensive first. When a debt is cleared, its minimum repayment joins the next step.', 2),
    numbered(derived.hasCar ? 'Everything after that: your car deposit and savings.' : 'Everything after that: savings and investing.', 2),
    small('HECS / HELP is repaid automatically through your tax once you earn over the threshold, so it never gets extra repayments here. Home loans keep their normal schedule; extra cash for a home loan usually works best in an offset account.'),
    spacer(120),
    h3('The secondary rule'),
    bullet(derived.waitItem),
  ];

  const hasInterestDebt = derived.debtOrder.some(d => d.type !== 'hecs' && d.rate > 0 && d.balance > 0);
  if (hasInterestDebt && derived.surplus >= 100) {
    sec.push(
      spacer(160),
      h3('The cost of doing nothing'),
      callout('Interest you avoid by following this plan', [
        sans('Paying only the minimums, you would pay ', { size: 20 }),
        sans(fmt(derived.baseInterest5y), { size: 20, bold: true }),
        sans(' in interest over the next 5 years. Following this plan, you pay ', { size: 20 }),
        sans(fmt(derived.planInterest5y), { size: 20, bold: true }),
        sans('. That is ', { size: 20 }),
        sans(fmt(derived.costOfInaction), { size: 20, bold: true, color: COLOR_RED }),
        sans(` (about ${fmt(derived.costOfInaction / (5 * 52))} a week) you keep.`, { size: 20 }),
      ]),
      small('How we worked this out: both figures come from a month-by-month calculation of your actual balances, rates and minimums, with interest charged monthly. HECS indexation is not counted as interest.'),
    );
  }
  sec.push(pageBreak());
  return sec;
}

// ─── 12-month cash map ───────────────────────────────────────────────────────
function build12MonthMap(state, derived, n) {
  const saveLabel = derived.hasCar ? 'To car / savings' : 'To savings';
  const header = ['Month', 'Surplus', 'To buffer', 'To debt', saveLabel, 'Buffer', 'Debt left', 'Focus'];
  const rows = derived.plan.rows.slice(0, 12).map(r => [
    monthLabel(r.m),
    fmt(r.surplus),
    r.fromBuffer > 0 ? `−${fmt(r.fromBuffer)}` : fmt(r.toBuffer),
    fmt(r.toDebt),
    fmt(r.toSave),
    fmt(r.buffer),
    fmt(r.debtLeft),
    r.focus.join(' → ') || '—',
  ]);
  const freed = derived.plan.rows.slice(0, 12).some(r => r.surplus > derived.surplus + 0.5);

  return [
    h1(`${n}. Your 12-month cash map`),
    p(`Where each month's surplus goes. Surplus = take-home ${fmt(state.income)} − living costs ${fmt(state.expenses)} − debt minimums actually due that month.`, { after: 160 }),
    table([header, ...rows], { size: 16 }),
    spacer(120),
    ...(freed ? [small('Your surplus rises in later months because a cleared debt no longer needs its minimum repayment; that money moves to the next step.')] : []),
    small('"Debt left" excludes HECS / HELP. Figures assume income and costs stay as entered. If they change, re-run the free tool for an updated plan.'),
    pageBreak(),
  ];
}

// ─── Debt roadmap ────────────────────────────────────────────────────────────
function buildDebtRoadmap(state, derived, n) {
  const order = derived.debtOrder;
  if (!order.length) {
    return [
      h1(`${n}. Debt roadmap`),
      p('You have no debts to pay off. Every dollar of surplus can go to your buffer and then savings, which is a strong position.', { after: 240 }),
      pageBreak(),
    ];
  }

  const rows = [
    ['Order', 'Debt', 'Balance', 'Rate', 'Cleared by', 'Interest you pay', 'Minimums only'],
    ...order.map((d, i) => {
      if (d.type === 'hecs') {
        return [String(i + 1), 'HECS / HELP', fmt(d.balance), rateText(d), 'Through tax', 'None (indexation only)', 'Same'];
      }
      const minOnly = d.minimumsOnlyClearedMonth === null
        ? (d.min > 0 ? 'Never at current minimum' : 'No minimum entered')
        : `${monthLabel(d.minimumsOnlyClearedMonth)}, ${fmt(d.minimumsOnlyInterest)} interest`;
      return [String(i + 1), cap(niceType(d.type)), fmt(d.balance), rateText(d), clearedText(d), fmt(d.interestPaid), minOnly];
    }),
  ];

  const sec = [
    h1(`${n}. Debt roadmap`),
    p('The order to pay your debts off, when each one is gone, and what it costs you, compared with paying only the minimums.', { after: 200 }),
    table(rows, { size: 16 }),
    spacer(160),
  ];

  if (derived.avalancheSaving !== null) {
    sec.push(
      h3('Why this order'),
      p([
        sans('Paying the most expensive debt first, instead of splitting extra money evenly across your debts, saves you ', { size: 20 }),
        sans(fmt(derived.avalancheSaving), { size: 20, bold: true, color: COLOR_NAVY }),
        sans(' in interest.', { size: 20 }),
      ]),
    );
  }

  const noMin = order.filter(d => d.type !== 'hecs' && d.rate > 0 && !(d.min > 0));
  if (noMin.length) {
    sec.push(small(`No minimum repayment was entered for your ${noMin.map(d => niceType(d.type)).join(', ')}, so interest builds on it until the plan reaches it. Check the minimum on your statement and re-run the tool.`));
  }

  if (order.some(d => d.type === 'hecs')) {
    sec.push(
      spacer(160),
      h3('Why HECS / HELP gets no extra repayments'),
      p(`HECS / HELP is indexed once a year (${HECS_INDEXATION_PCT}% for June 2026, the lower of CPI and wage growth) but charges no interest. It is repaid automatically through your tax once your income passes the repayment threshold. Paying it early rarely beats paying off interest-bearing debt or building savings. Confirm the current rate at studyassist.gov.au.`),
    );
  }
  sec.push(
    spacer(120),
    small('How we worked this out: month by month, each debt is charged interest at its rate, then its minimum is paid, then the plan\'s extra money goes to the debt listed first. "Minimums only" runs the same calculation with no extra money.'),
    pageBreak(),
  );
  return sec;
}

// ─── Car (only when a car is planned) ────────────────────────────────────────
function buildCarScenarios(state, derived, n) {
  const car = state.car;
  const principal = car.price - car.deposit;
  const dealerPmt = calcMonthlyMortgage(principal, car.dealerRate, car.term);
  const bankPmt = calcMonthlyMortgage(principal, car.bankRate, car.term);
  const dealerTotal = dealerPmt * car.term * 12;
  const bankTotal = bankPmt * car.term * 12;
  const saving = dealerTotal - bankTotal;

  // Deposit growth while waiting = what the plan actually sets aside for the car in those months.
  const set6 = derived.plan.rows.slice(0, 6).reduce((s, r) => s + r.toSave, 0);
  const set12 = derived.plan.rows.slice(0, 12).reduce((s, r) => s + r.toSave, 0);
  const wait6Deposit = car.deposit + set6;
  const wait12Deposit = car.deposit + set12;
  const wait6Price = car.price * (1 - RUNNING_COSTS.waitHold6mo);
  const wait12Price = car.price * (1 - RUNNING_COSTS.waitHold12mo);
  const wait6Principal = Math.max(0, wait6Price - wait6Deposit);
  const wait12Principal = Math.max(0, wait12Price - wait12Deposit);
  const wait6Pmt = calcMonthlyMortgage(wait6Principal, car.bankRate, car.term);
  const wait12Pmt = calcMonthlyMortgage(wait12Principal, car.bankRate, car.term);
  const wait6Total = wait6Pmt * car.term * 12;
  const wait12Total = wait12Pmt * car.term * 12;

  const rows = [
    ['Scenario', 'Deposit', 'Loan size', 'Monthly', 'Total repaid', 'Vs dealer now'],
    ['Buy now (dealer finance)', fmt(car.deposit), fmt(principal), fmt(dealerPmt), fmt(dealerTotal), '—'],
    ['Buy now (bank loan)', fmt(car.deposit), fmt(principal), fmt(bankPmt), fmt(bankTotal), fmt(-saving)],
    ['Wait 6 months (bank)', fmt(wait6Deposit), fmt(wait6Principal), fmt(wait6Pmt), fmt(wait6Total), fmt(wait6Total - dealerTotal)],
    ['Wait 12 months (bank)', fmt(wait12Deposit), fmt(wait12Principal), fmt(wait12Pmt), fmt(wait12Total), fmt(wait12Total - dealerTotal)],
  ];

  return [
    h1(`${n}. Car decision: four scenarios`),
    p(`Four ways to buy this car and what each costs over your ${car.term}-year term.`, { after: 200 }),
    table(rows),
    spacer(200),
    h3('Assumptions'),
    bullet(`Price assumed to fall about ${(RUNNING_COSTS.waitHold6mo * 100).toFixed(0)}% over 6 months (${fmt(car.price)} → ${fmt(wait6Price)}) and about ${(RUNNING_COSTS.waitHold12mo * 100).toFixed(0)}% over 12 months (→ ${fmt(wait12Price)}). Real depreciation varies by make and model.`),
    bullet(set12 > 0
      ? `While waiting, the deposit grows by what your plan sets aside for the car: ${fmt(set6)} in 6 months and ${fmt(set12)} in 12 months.`
      : 'Your plan sends all surplus to your buffer and debts for the next 12 months, so the deposit does not grow while you wait.'),
    bullet(`The bank rate (${fmtPct(car.bankRate)}) is used for the wait scenarios, on the assumption you would shop around with more time.`),
    spacer(200),
    h3('The headline number'),
    p([
      sans('Choosing the bank loan over dealer finance saves ', { size: 22 }),
      sans(fmt(saving), { size: 22, bold: true, color: COLOR_NAVY }),
      sans(' over the loan term.', { size: 22 }),
    ]),
    small('The dealer rate is the rate you entered. The comparison rate, which includes fees, can be noticeably higher. Always ask for the comparison rate in writing before signing.'),
    pageBreak(),
  ];
}

function buildOwnershipCost(state, derived, n) {
  const car = state.car;
  const principal = car.price - car.deposit;
  const monthlyLoan = calcMonthlyMortgage(principal, car.bankRate, car.term);
  const annualKms = (state._post_purchase && state._post_purchase.annualKms) || 15000;
  const fuelAnnual = (annualKms / 100) * RUNNING_COSTS.fuelLitresPer100km * RUNNING_COSTS.fuelPricePerLitre;
  const depreciationAnnual = car.price * RUNNING_COSTS.depreciationPctYr1;
  const totalAnnual = monthlyLoan * 12 + RUNNING_COSTS.insuranceAnnual + RUNNING_COSTS.regoAnnual
    + RUNNING_COSTS.servicingAnnual + fuelAnnual + depreciationAnnual;
  const totalWeekly = totalAnnual / 52;
  const pctOfIncome = ((totalAnnual / 12) / state.income) * 100;

  const rows = [
    ['Cost', 'Per year', 'Per week', 'Notes'],
    ['Loan repayment (bank)', fmt(monthlyLoan * 12), fmt((monthlyLoan * 12) / 52), `${fmtPct(car.bankRate)} over ${car.term} years`],
    ['Comprehensive insurance', fmt(RUNNING_COSTS.insuranceAnnual), fmt(RUNNING_COSTS.insuranceAnnual / 52), 'A$30–40k car, mid-range driver'],
    ['Registration + CTP', fmt(RUNNING_COSTS.regoAnnual), fmt(RUNNING_COSTS.regoAnnual / 52), 'Varies by state'],
    ['Servicing + tyres', fmt(RUNNING_COSTS.servicingAnnual), fmt(RUNNING_COSTS.servicingAnnual / 52), 'Logbook + minor services'],
    ['Fuel', fmt(fuelAnnual), fmt(fuelAnnual / 52), `${annualKms.toLocaleString()} km/yr at ${RUNNING_COSTS.fuelLitresPer100km} L/100km, A$${RUNNING_COSTS.fuelPricePerLitre}/L`],
    ['Depreciation (year 1)', fmt(depreciationAnnual), fmt(depreciationAnnual / 52), `~${(RUNNING_COSTS.depreciationPctYr1 * 100).toFixed(0)}% in year 1, ~${(RUNNING_COSTS.depreciationPctYr2plus * 100).toFixed(0)}%/yr after`],
    ['Total, year 1', fmt(totalAnnual), fmt(totalWeekly), `${pctOfIncome.toFixed(1)}% of take-home pay`],
  ];

  return [
    h1(`${n}. What the car really costs`),
    p('The repayment is only one of six costs. This is the figure most buyers underestimate.', { after: 200 }),
    table(rows),
    spacer(200),
    p([
      sans('All-in, this car costs about ', { size: 22 }),
      sans(`${fmt(totalWeekly)} a week`, { size: 22, bold: true }),
      sans(` in year 1, roughly ${pctOfIncome.toFixed(1)}% of your take-home pay.`, { size: 22 }),
    ]),
    small('Running costs are FY2025-26 illustrative averages from public Australian sources (Budget Direct, RACV, Finder). Yours will vary by car, postcode, driver and kilometres.'),
    pageBreak(),
  ];
}

// ─── Stress tests ────────────────────────────────────────────────────────────
function buildStressTest(state, derived, n) {
  const essentials = derived.monthlyEssentials;          // living costs + minimums
  const debts = state.debts || [];

  const variable = debts.filter(d => VARIABLE_RATE_TYPES.includes(d.type) && d.balance > 0 && d.rate > 0);
  const rateRiseMonthly = variable.reduce((s, d) => s + d.balance * RATE_RISE_PCT / 1200, 0);
  let carRise = 0;
  if (derived.hasCar) {
    const principal = state.car.price - state.car.deposit;
    carRise = calcMonthlyMortgage(principal, state.car.bankRate + RATE_RISE_PCT, state.car.term)
            - calcMonthlyMortgage(principal, state.car.bankRate, state.car.term);
  }
  const rateImpact = variable.length || carRise > 0
    ? `About ${fmt(rateRiseMonthly + carRise)} more a month (${[...variable.map(d => niceType(d.type)), ...(carRise > 0 ? ['planned car loan'] : [])].join(', ')})`
    : 'None of your debts are variable-rate';
  const rateMeaning = rateRiseMonthly + carRise > 0
    ? `Surplus falls to ${fmt(derived.surplus - rateRiseMonthly - carRise)}; payoff dates move later`
    : 'Little direct impact';

  const debtMins = derived.debtMins;
  const stressSurplus = state.income * 0.85 - state.expenses - debtMins;
  const emergencyAfter = state.savings - 3000;
  const emergencyMonths = emergencyAfter / Math.max(essentials, 1);
  const jobLossNeed = essentials * 3;
  const jobLossMargin = state.savings - jobLossNeed;

  const rows = [
    ['What happens', 'Impact on your numbers', 'What it means'],
    [`Rates rise ${RATE_RISE_PCT}%`, rateImpact, rateMeaning],
    ['Income drops 15%', `Surplus ${fmt(derived.surplus)} → ${fmt(stressSurplus)}`, stressSurplus < 0 ? 'Plan stops working: cut costs or seek hardship help' : 'Same steps, every date moves later'],
    ['A$3,000 emergency bill', `Buffer ${fmt(state.savings)} → ${fmt(emergencyAfter)}`, emergencyAfter < 0 ? `Savings cover ${fmt(state.savings)}; the other ${fmt(-emergencyAfter)} would likely go on a card` : emergencyMonths < 1 ? 'Buffer back below 1 month: rebuild it first' : `Buffer still covers ${emergencyMonths.toFixed(1)} months`],
    ['3 months without income', `You need ${fmt(jobLossNeed)} (costs + minimums)`, jobLossMargin >= 0 ? `Covered, with ${fmt(jobLossMargin)} to spare` : `${fmt(-jobLossMargin)} short`],
  ];

  return [
    h1(`${n}. Stress tests`),
    p('What happens if things don\'t go to plan, using your numbers.', { after: 200 }),
    table(rows),
    spacer(200),
    h3('What this tells you'),
    p(jobLossMargin >= 0
      ? 'Your savings would carry you through three months without income. That is uncommon. Protect the buffer and avoid dipping into it for non-essentials.'
      : `Your savings would not carry you through three months without income. Reaching the 3-month buffer (${monthLabel(derived.plan.milestones.buffer3)} on these numbers) fixes that.`),
    spacer(120),
    small('Rate rise: applies to debts whose rates usually move with the market (credit cards, personal loans, variable home loans). Fixed-rate loans only feel it when the fixed term ends. Income drop: assumes living costs stay the same; in practice some can be cut.'),
    pageBreak(),
  ];
}

// ─── Action checklist ────────────────────────────────────────────────────────
// Days from today to the start of plan month m (for sorting dated checklist items).
const daysUntilMonth = (m) => Math.round((new Date(TODAY.getFullYear(), TODAY.getMonth() + m, 1) - TODAY) / 86400000);

function buildActionChecklist(state, derived, n) {
  const plan = derived.plan;
  const items = [];   // { day, text } — sorted by day; undated items go last in insertion order
  const add = (day, text) => items.push({ day, text });
  const firstPriority = derived.debtOrder.find(d => d.priority);
  const UNDATED = Infinity;

  if (derived.surplus >= 100) {
    if (derived.bufferMonths < RULES.MIN_BUFFER_MONTHS) {
      add(0, `This week: open a separate high-interest savings account (for example ING Savings Maximiser or Macquarie Savings) and automate ${fmt(derived.surplus)}/month into it until it holds ${fmt(plan.oneMonth)} (${monthLabel(plan.milestones.buffer1)}).`);
      if (firstPriority) {
        add(daysUntilMonth(plan.milestones.buffer1), `${monthLabel(plan.milestones.buffer1)}: once the buffer is there, move the same automatic payment to your ${niceType(firstPriority.type)}, on top of the minimum. It is cleared by ${clearedText(firstPriority)}.`);
      }
    } else if (firstPriority) {
      add(0, `This week: set up an automatic extra ${fmt(derived.surplus)}/month on your ${niceType(firstPriority.type)}, on top of the minimum. It is cleared by ${clearedText(firstPriority)}.`);
    }
    if (firstPriority && firstPriority.rate > 0) {
      add(14, `Within 2 weeks: call your ${niceType(firstPriority.type)} provider and ask for a lower rate or a 0% balance-transfer offer. Every point off the rate brings the payoff date closer.`);
    }
    if (plan.milestones.buffer3 !== null && plan.milestones.buffer3 > 0) {
      const next = derived.debtOrder.some(d => !NO_EXTRA_REPAYMENT_TYPES.includes(d.type) && d.clearedMonth > plan.milestones.buffer3)
        ? 'your remaining debts' : (derived.hasCar ? 'your car deposit and savings' : 'savings and investing');
      add(daysUntilMonth(plan.milestones.buffer3), `${monthLabel(plan.milestones.buffer3)}: your buffer reaches 3 months (${fmt(plan.threeMonth)}). From then on, surplus goes to ${next}.`);
    }
  } else {
    add(0, 'This week: list every fixed cost and cancel or renegotiate at least one (insurance, phone, subscriptions, refinancing).');
    add(0, 'If minimum repayments are hard to meet, call the National Debt Helpline (1800 007 007). It is free and confidential.');
  }

  add(90, `${fmtDate(90)}: re-run the free tool at moneymoves-au.vercel.app with your updated numbers to get a fresh plan.`);

  if (derived.hasCar) {
    add(UNDATED, 'Before signing any car loan: get the comparison rate in writing from both the dealer and your bank.');
    add(UNDATED, 'Compare insurance quotes from at least 3 providers before you buy.');
    if (state.car.novated === 'yes') {
      add(UNDATED, 'Speak to a salary-packaging specialist before committing to a novated lease. Eligibility and benefits depend on your employer and the vehicle.');
    }
  }
  add(UNDATED, 'Before any big decision (refinancing, a new loan, a super top-up), consider one session with a fee-only licensed financial adviser, typically A$200–400.');

  const ordered = items.map((it, i) => ({ ...it, i })).sort((a, b) => (a.day - b.day) || (a.i - b.i));

  return [
    h1(`${n}. Your action checklist`),
    p('In date order. Tick them off as you go.', { after: 200 }),
    ...ordered.map(it => bullet(it.text)),
    spacer(240),
    h3('Useful Australian resources'),
    bullet([sans('ASIC Moneysmart: ', { size: 20, bold: true }), sans('moneysmart.gov.au (free, government-backed)', { size: 20 })]),
    bullet([sans('National Debt Helpline: ', { size: 20, bold: true }), sans('1800 007 007 (free, confidential financial counsellors)', { size: 20 })]),
    bullet([sans('ATO payment plans: ', { size: 20, bold: true }), sans('ato.gov.au (set up online if you owe under A$100k)', { size: 20 })]),
    bullet([sans('Comparison rates explained: ', { size: 20, bold: true }), sans('moneysmart.gov.au/loans/comparison-rate', { size: 20 })]),
    spacer(320),
    p([
      sans('General information only. ', { size: 18, italics: true, bold: true, color: COLOR_INK_SOFT }),
      sans('This plan is based on the numbers you entered and fixed rules. It is not personal financial product advice and does not consider your full circumstances. Speak to a licensed Australian financial adviser before any major decision.', { size: 18, italics: true, color: COLOR_INK_SOFT }),
    ]),
  ];
}

// ─── Toolkits ────────────────────────────────────────────────────────────────
function buildBonusToolkits(state, derived) {
  const toolkits = [];

  if (derived.hasCar) {
    toolkits.push((k) => [
      pageBreak(),
      h1(`Toolkit ${k}: Car dealer negotiation`),
      p('Dealer finance often carries a higher rate, extra fees and add-ons that are easy to miss in the paperwork. Use these when you talk to the dealer.', { after: 200 }),
      callout('Rule 1: agree the price before you talk about finance', 'Negotiate the drive-away price first. Discussing price and finance together makes it easy to give back a discount through a higher rate, a bigger balloon or added fees.'),
      spacer(180),
      h3('What to say'),
      bullet([sans('When asked "What monthly repayment are you after?": ', { size: 20, bold: true }), sans('"I\'m only negotiating the drive-away price today. I have finance pre-approved with my bank."', { size: 20, italics: true })]),
      bullet([sans('When offered dealer finance: ', { size: 20, bold: true }), sans(`"I'll consider it if your comparison rate, in writing, beats my bank's ${fmtPct(state.car.bankRate)}, with no early payout fees or bundled add-ons."`, { size: 20, italics: true })]),
      bullet([sans('When add-ons appear (paint protection, extended warranty, gap insurance): ', { size: 20, bold: true }), sans('"Please remove the optional add-ons from the drive-away price. I\'ll decide on any of them separately."', { size: 20, italics: true })]),
      spacer(200),
      h3('Red flags'),
      bullet([sans('Large balloon payment: ', { size: 20, bold: true }), sans('keeps repayments low but leaves a big lump sum at the end, often more than the car is worth.', { size: 20 })]),
      bullet([sans('Headline rate only: ', { size: 20, bold: true }), sans('the comparison rate includes establishment and monthly fees. Compare comparison rates.', { size: 20 })]),
      bullet([sans('Used car: ', { size: 20, bold: true }), sans('run a PPSR search at ppsr.gov.au (A$2) to check it isn\'t under finance, written off or stolen.', { size: 20 })]),
    ]);
  }

  const rateDebt = derived.debtOrder.find(d => !NO_EXTRA_REPAYMENT_TYPES.includes(d.type) && d.rate > 0 && d.balance > 0);
  if (rateDebt) {
    const debtName = niceType(rateDebt.type);
    toolkits.push((k) => [
      pageBreak(),
      h1(`Toolkit ${k}: The rate-cut phone call for your ${debtName}`),
      p(`Your ${debtName} costs ${fmtPct(rateDebt.rate)} on ${fmt(rateDebt.balance)}, about ${fmt(rateDebt.balance * rateDebt.rate / 1200)} in interest this month. Lenders often give better rates to new customers than existing ones, and a short call can be enough to get yours reviewed.`, { after: 200 }),
      callout('Before you call', `Look up two or three real offers for a ${debtName} (rates or balance-transfer deals) on Finder, Mozo or Canstar. Write down the best one. Have your account number, balance (${fmt(rateDebt.balance)}) and current rate (${fmtPct(rateDebt.rate)}) ready.`),
      spacer(180),
      h3('The call, step by step'),
      bullet([sans('1. Ask for the right team: ', { size: 20, bold: true }), sans('ask for "retentions" or say you are thinking of closing or transferring the account. That team can usually offer more than general support.', { size: 20 })]),
      bullet([sans('2. Open with the facts: ', { size: 20, bold: true }), sans(`"I'm paying ${fmtPct(rateDebt.rate)} on a ${fmt(rateDebt.balance)} balance and I've always paid on time. [Lender] is offering [the rate you found]."`, { size: 20, italics: true })]),
      bullet([sans('3. Ask directly: ', { size: 20, bold: true }), sans('"What can you do on the rate to keep my account?"', { size: 20, italics: true })]),
      bullet([sans('4. If the answer is no: ', { size: 20, bold: true }), sans('"Thanks. I\'ll go ahead with the transfer then. Can you tell me the payout figure?" Then compare the balance-transfer offer properly, including the rate after the promotional period ends.', { size: 20, italics: true })]),
      bullet([sans('5. Ask about the annual fee: ', { size: 20, bold: true }), sans('"Can you also waive this year\'s annual fee?"', { size: 20, italics: true })]),
      small('Only quote offers you have actually seen. A balance transfer only helps if you stop adding to the card and clear the balance before the promotional rate ends.'),
    ]);
  }

  const first = derived.firstRow || { toBuffer: 0, toDebt: 0, toSave: 0 };
  toolkits.push((k) => [
    pageBreak(),
    h1(`Toolkit ${k}: Automate your payday`),
    p('Good decisions are easier when they happen automatically. This four-account setup runs your plan on payday with scheduled transfers (PayID / Osko).', { after: 200 }),
    callout('The rule', 'Don\'t spend from the account your pay lands in. The day after payday, scheduled transfers move every dollar to the account it belongs in.'),
    spacer(180),
    h3('Your four accounts'),
    table([
      ['Account', 'What it is for', 'Card?', 'Your amount each month'],
      ['1. Pay arrives', 'Receives your take-home pay', 'No', `${fmt(state.income)} in, all moved out the next day`],
      ['2. Bills', 'Rent or mortgage, utilities, food, rego, debt minimums', 'No', fmt(state.expenses + derived.debtMins)],
      ['3. Buffer / debt', 'High-interest savings, plus your extra debt payment', 'No', first.toBuffer || first.toDebt ? `${fmt(first.toBuffer + first.toDebt)} this month (${(first.focus || []).join(' → ') || 'per your plan'})` : '—'],
      ['4. Spending', 'Everything else', 'Yes', 'Whatever is left after the above'],
    ]),
    spacer(200),
    h3('Setting it up'),
    bullet([sans('Low-fee banks: ', { size: 20, bold: true }), sans('many Australian banks offer fee-free accounts with instant PayID/Osko transfers. Compare on Canstar or Finder.', { size: 20 })]),
    bullet([sans('Timing: ', { size: 20, bold: true }), sans('schedule transfers for the business day after payday, so a late pay run doesn\'t cause an overdraw.', { size: 20 })]),
    bullet([sans('No cards on accounts 2 and 3: ', { size: 20, bold: true }), sans('keeping them off Apple Pay and Google Wallet stops impulse spending from them.', { size: 20 })]),
  ]);

  return toolkits.flatMap((build, i) => build(i + 1));
}

// ─── Build doc ───────────────────────────────────────────────────────────────
function buildReport(state, opts = {}) {
  const derived = deriveReportData(state);
  const includeBonuses = opts.includeBonuses !== false;

  // Sections are numbered in the order they appear; car sections only exist when a car is planned.
  const builders = [
    buildPlanLogic,
    build12MonthMap,
    buildDebtRoadmap,
    ...(derived.hasCar ? [buildCarScenarios, buildOwnershipCost] : []),
    buildStressTest,
    buildActionChecklist,
  ];
  const children = [
    ...buildCover(state, derived),
    ...builders.flatMap((build, i) => build(state, derived, i + 1)),
    ...(includeBonuses ? buildBonusToolkits(state, derived) : []),
  ];

  return new Document({
    creator: 'MoneyMoves AU',
    title: `MoneyMoves AU — ${REPORT_NAME}`,
    description: 'Personalised money plan for an Australian household.',
    styles: {
      default: { document: { run: { font: 'Arial', size: 22 } } },
      paragraphStyles: [
        { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 32, bold: true, font: 'Arial', color: COLOR_INK },
          paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 } },
        { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 24, bold: true, font: 'Arial', color: COLOR_NAVY_DARK },
          paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 2 } },
      ],
    },
    numbering: {
      config: [
        {
          reference: 'bullets',
          levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } } }],
        },
        {
          reference: 'steps',
          levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } } }],
        },
      ],
    },
    sections: [{
      properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1200, right: 1200, bottom: 1200, left: 1200 } } },
      headers: {
        default: new Header({
          children: [new Paragraph({
            alignment: AlignmentType.RIGHT,
            children: [sans(`MoneyMoves AU — ${REPORT_NAME}`, { size: 16, color: COLOR_INK_SOFT })],
          })],
        }),
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              sans('General information only — not personal financial advice.', { size: 16, color: COLOR_INK_SOFT }),
              new TextRun({ children: ['  |  Page ', PageNumber.CURRENT, ' of ', PageNumber.TOTAL_PAGES], size: 16, color: COLOR_INK_SOFT, font: 'Arial' }),
            ],
          })],
        }),
      },
      children,
    }],
  });
}

async function generateReport(state, opts = {}) {
  const doc = buildReport(state, opts);
  return Packer.toBuffer(doc);
}

module.exports = { buildReport, generateReport, deriveReportData, simulatePlan };

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
