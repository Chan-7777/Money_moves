# Roadmap prompts

Copy one prompt into Claude Code per session. Do items in order within a phase.
Every prompt relies on the rules in `CLAUDE.md` (loaded automatically) and ends by
updating `docs/ROADMAP.md`.

---

## Start of any session

```
Read CLAUDE.md and docs/ROADMAP.md. Tell me: what is live, what is in progress or in review,
and which item is next. Run npm test and confirm the golden figures still pass. Don't change
anything yet.
```

## Check nothing broke (run after every merge)

```
Regression check after the last merge:
1. npm test, design-system/lint-tokens.sh, scripts/dead-code-scan.sh.
2. Open the live site at 1440px and 390px. Run the reference case (take-home 5,200; costs 3,400;
   savings 1,500; credit card 4,800 at 20.99% min 145; HECS 24,000) through the app.
   Confirm: surplus A$1,655, first move "Build a 1-month emergency buffer (A$3,545)",
   HECS shows "Indexed / Via tax", stickers animate, keyboard can operate every option.
3. Confirm the PDF download returns a PDF (I'll sign in if needed).
4. Check the browser console for errors.
Report pass/fail per step and update the ROADMAP change log. Don't fix anything without asking.
```

---

## Phase 0

### P0-1 — One price everywhere
```
Roadmap item P0-1. The one-off price is A$[PRICE] for [WHAT IS INCLUDED].
Branch p0-1-one-price. Make every price mention match: index.html pricing section and FAQ,
app.html paid section and post-payment copy, terms.html, and the server price in
api/create-order.js + the accepted amount in lib/validators.js. Remove all "/month",
"cancel anytime", "monthly rate defense", "Save A$100" and A$39/A$59 wording.
Serve the price from /api/config so the page can't drift from the server.
I approve editing the payment files for this item only.
Tests: price shown == server price; validator accepts only the new amount.
Follow CLAUDE.md "Finishing an item", update docs/ROADMAP.md, don't merge.
```

### P0-2 — Hide paid tier until PayPal is live
```
Roadmap item P0-2. Branch p0-2-hide-paid. Show the paid tier (landing pricing cards and the
in-app "View paid options" teaser) only when /api/config reports PayPal is in live mode.
Otherwise the free plan and free PDF stay. No payment-file edits.
Follow CLAUDE.md "Finishing an item", update docs/ROADMAP.md, don't merge.
```

### P0-6 — Upstash rate limiting
```
Roadmap item P0-6. I've added UPSTASH_REDIS_REST_URL and UPSTASH_REDIS_REST_TOKEN to Vercel
production. Verify (read-only) that api/free-report.js uses them, then test against a preview:
the 6th request within a minute from one IP must get 429. Don't change payment files.
Update docs/ROADMAP.md.
```

### P0-7 — Privacy policy
```
Roadmap item P0-7. Branch p0-7-privacy. Update privacy.html so it names every service that
receives user data and why: Clerk (sign-in), ConvertAPI (receives plan numbers to build the PDF),
Resend (email), PayPal (payment), Firebase (anonymous page-visit count). Keep "we do not store your
financial inputs" only if still true. Update the "Last updated" date.
Follow CLAUDE.md "Finishing an item", update docs/ROADMAP.md, don't merge.
```

### P0-9 — CI failures
```
Roadmap item P0-9. Branch p0-9-ci. Fix the CI failures that predate the roadmap:
npm audit high (nanoid), test/e2e/smoke.spec.js looking for the old "14" price, and Stryker
failing its initial run. Fix causes, don't loosen or skip assertions. Show the CI run passing.
Update docs/ROADMAP.md, don't merge.
```

---

## Phase 1

### P1-1 — Shared money engine (do first)
```
Roadmap item P1-1. Branch p1-1-plan-engine.
Move the month-by-month simulation (simulatePlan, debt ordering, buffer targets, interest)
out of build_pdf_report.js into lib/plan-engine.js. It must work both in Node (module.exports)
and in the browser (window.PlanEngine, loaded with <script src="lib/plan-engine.js">).
Then make app.html's results (surplus, buffer target, debt order, payoff info) come from the
engine instead of its own formulas, and make build_pdf_report.js import it.
Do NOT change any output: all golden figures in CLAUDE.md and every existing test must pass
unchanged. Add a test that the app's figures equal the PDF's for the reference case and for
3 more cases (no debts, negative surplus, car planned).
Follow CLAUDE.md "Finishing an item", update docs/ROADMAP.md, don't merge.
```

### P1-2 — CSV / Excel download
```
Roadmap item P1-2. Branch p1-2-csv-export. Add a "Download as spreadsheet (CSV)" button on the
results page that exports the 12-month cash map and the debt schedule from lib/plan-engine.js.
Columns must match the PDF's cash map. It opens correctly in Excel and Google Sheets
(UTF-8 with BOM, A$ amounts as plain numbers). No server call needed.
Tests: CSV rows equal the engine's first 12 months for the reference case.
Follow CLAUDE.md "Finishing an item", update docs/ROADMAP.md, don't merge.
```

### P1-3 — "Is my debt a lot?" health check
```
Roadmap item P1-3. Branch p1-3-debt-health. Add a card near the top of the results:
total debt (excluding HECS, shown separately), debt as % of yearly take-home, months until
consumer debt is cleared on the plan, and a green/amber/red verdict.
Define the thresholds in lib/plan-engine.js with a comment explaining them. Don't compare to
"the average Australian" unless the figure is cited (ABS source + date).
Show the same card in the PDF summary page.
Tests: verdict boundaries, HECS excluded, no-debt case, app == PDF.
Follow CLAUDE.md "Finishing an item", update docs/ROADMAP.md, don't merge.
```

### P1-4 — Snowball vs avalanche
```
Roadmap item P1-4. Branch p1-4-snowball. Add a 'snowball' strategy (smallest balance first)
to lib/plan-engine.js beside the existing avalanche plan. On the results page show both:
debt-free date and total interest for each, with one plain sentence on the trade-off
(snowball = quicker early wins, avalanche = less interest). Only show it with 2+ debts.
The recommended plan stays avalanche.
Tests: hand-worked 2-debt case where the orders differ; single-debt case hides the comparison.
Follow CLAUDE.md "Finishing an item", update docs/ROADMAP.md, don't merge.
```

### P1-5 — Debt-free by a date
```
Roadmap item P1-5. Branch p1-5-target-date. Add an optional "I want to be debt-free by
[month/year]" input on the results page. Use lib/plan-engine.js to find the monthly amount
needed (search over the extra payment until the last consumer debt clears by that month).
Show: needed amount, current surplus, achievable yes/no, and what it means for the buffer.
HECS and home loans are excluded, as in the plan.
Tests: hand-checked single-debt amortisation case; unreachable date; date already met.
Follow CLAUDE.md "Finishing an item", update docs/ROADMAP.md, don't merge.
```

### P1-6 — Lump sum
```
Roadmap item P1-6. Branch p1-6-lump-sum. Add an optional "one-off amount" (bonus, tax refund)
input. Apply it in lib/plan-engine.js at month 0 in the plan's order (buffer to 1 month,
priority debt, 3-month buffer, other debt, savings) and show before/after: dates and interest.
Tests: lump sum smaller than buffer gap; lump sum that clears a debt; no debts.
Follow CLAUDE.md "Finishing an item", update docs/ROADMAP.md, don't merge.
```

---

## Phase 2

### P2-1 — Balance transfer / consolidation calculator
```
Roadmap item P2-1. Branch p2-1-balance-transfer. Add a calculator: user enters an offer
(promo rate, promo months, transfer fee %, revert rate) or a consolidation loan (rate, fees, term).
Using lib/plan-engine.js, compare total interest + fees and payoff date against staying put,
at the user's planned repayment. Show the revert-rate risk if the balance isn't cleared in the
promo period. General information wording only; no lender names or recommendations.
Tests: hand-worked transfer case; case where the transfer costs more.
Follow CLAUDE.md "Finishing an item", update docs/ROADMAP.md, don't merge.
```

### P2-2 — HECS repayment estimator
```
Roadmap item P2-2. Branch p2-2-hecs. Create lib/au-rates.js holding HECS indexation and the
current ATO repayment thresholds/rates with source URL, financial year and validUntil, plus a
test that fails after validUntil. Add an optional gross-income input and estimate yearly
compulsory repayment and years to clear (including indexation). Replace the hard-coded 2.8%
in build_pdf_report.js with the au-rates value.
Tests: thresholds from the ATO table for 3 incomes (worked by hand), stale-data test.
Follow CLAUDE.md "Finishing an item", update docs/ROADMAP.md, don't merge.
```

### P2-3 — Low-income / shortfall mode
```
Roadmap item P2-3. Branch p2-3-shortfall. When surplus < A$100 or negative, replace the thin
message with a practical page: biggest fixed costs to review (from the user's own inputs where
possible), a hardship-request script for lenders, the National Debt Helpline (1800 007 007),
and a re-run prompt. No invented savings figures. Same content in the PDF.
Tests: negative, zero and A$50 surplus cases render the mode; positive surplus does not.
Follow CLAUDE.md "Finishing an item", update docs/ROADMAP.md, don't merge.
```

---

## Phase 3

### P3-1 — Mortgage extra repayments + offset
```
Roadmap item P3-1. Branch p3-1-mortgage. Add a home-loan tool: balance, rate, years left,
repayment; show time and interest saved by extra repayments or an offset balance, and
"pay off a 30-year loan in 15 years" (required extra). Use lib/plan-engine.js amortisation.
Home loans stay excluded from the main debt plan's extra repayments.
Tests: hand-checked amortisation; offset equals extra-repayment interest saving in the simple case.
Follow CLAUDE.md "Finishing an item", update docs/ROADMAP.md, don't merge.
```

### P3-2 — Opt-in progress tracker
```
Roadmap item P3-2. Branch p3-2-tracker. Propose (don't build yet) an opt-in tracker: what is
stored, where, retention, how users delete it, how the monthly check-in email works, and every
copy change needed to "nothing is stored". Wait for my approval before writing code.
Update docs/ROADMAP.md with the proposal link.
```

---

## Phase 4

### P4-1 — Landing pages
```
Roadmap item P4-1. Branch p4-1-landing-pages. From docs/research/*.csv pick the 5 questions with
"App answers it? = Yes" and the most search variants. Build one page each (e.g.
/what-debt-to-pay-off-first-australia) with a direct answer, a worked example from
lib/plan-engine.js, and a link into the tool. Plain English, no invented figures, design tokens
only, add to sitemap. Follow CLAUDE.md "Finishing an item", update docs/ROADMAP.md, don't merge.
```

### P4-2 — Sourced articles
```
Roadmap item P4-2. Branch p4-2-articles. Draft (as files, not published) 3 articles for the
"stats" and "credit score" questions in docs/research. Every statistic must cite its source
(ABS, RBA, ASIC Moneysmart, credit bureaus) with a date. List every source at the end.
Show me drafts before any page goes live. Update docs/ROADMAP.md.
```
