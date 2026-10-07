# MoneyMoves AU — Claude Code instructions

@AGENTS.md

Everything in AGENTS.md applies (definition of done, green gate, money-path protection,
banned shortcuts). This file adds the rules for building the feature roadmap without
breaking what is live.

## Roadmap
- Plan and live status: `docs/ROADMAP.md`. Read it at the start of any roadmap task.
- Task prompts: `docs/roadmap-prompts.md`.
- Research behind it: `docs/research/` (real Australian Google questions, tagged by coverage).

## Build rules (nothing breaks)

1. **One item at a time, one branch per item.** Branch name = the item ID in ROADMAP.md
   (e.g. `p1-2-csv-export`). Never mix two roadmap items in one commit.
2. **Never merge to `master` or deploy without the user saying so in the current message.**
   `master` auto-deploys to production (moneymoves-au.vercel.app).
3. **One money engine.** Once item P1-1 is done, every money calculation (surplus, buffer
   targets, interest, payoff dates, allocation order) lives in `lib/plan-engine.js` and is
   used by both `app.html` and `build_pdf_report.js`. Never re-implement a formula in a page,
   the PDF or an API route — add it to the engine with a test.
4. **Golden figures must not move.** The reference case in `test/report.test.js`
   (take-home 5,200; costs 3,400; savings 1,500; card 4,800 @ 20.99% min 145; HECS 24,000)
   must keep producing: surplus 1,655 · 1-month buffer 3,545 · 3-month buffer 10,635 ·
   buffer full in month 2 · card cleared month 4 · minimums-only 50 months / ~A$2,435.
   If a change moves one of these, stop and explain why before changing the test.
5. **Tests first.** Every new calculation gets tests with expected values worked out
   independently (by hand or a separate formula) — never by reading the code's own output.
6. **No invented numbers.** Every figure shown to a user is computed from their inputs or
   comes from a cited source (source + date in a code comment). No floors, multipliers or
   "estimated" values that are not calculated. Unsourced statistics are banned in copy,
   PDFs and articles.
7. **Yearly data lives in one place.** HECS indexation, ATO repayment thresholds and any
   other rate that changes each financial year go in `lib/au-rates.js` with `validUntil`.
   A test fails once `validUntil` has passed, so stale rates cannot ship silently.
8. **App and PDF must agree.** Any figure shown in both places is asserted equal in a test.
9. **Keep the existing UX fixes.** Pills stay `<button>`s with `aria-pressed`; inputs keep
   labels and `inputmode="decimal"`; errors keep `aria-invalid` + `aria-describedby`;
   reduced-motion users get no animation; colours come from `design-system/tokens.css`.
10. **Payment and auth stay untouched** (`api/create-order.js`, `api/capture-order.js`,
    `api/free-report.js`, Clerk setup) unless the item says otherwise and the user approves
    in the current message.

## Finishing an item (every time)

1. `npm test`, `bash design-system/lint-tokens.sh`, `bash scripts/dead-code-scan.sh` — all pass.
2. Check the change in a real browser (local server or Vercel preview), desktop and 390px mobile.
3. Update the item's row in `docs/ROADMAP.md`: status, branch, commit, tests added, date,
   and anything not tested.
4. Report: what changed, exact commands run, what was NOT tested. Never call it done
   if any of that is missing.
