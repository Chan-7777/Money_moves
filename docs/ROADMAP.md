# MoneyMoves AU — Roadmap tracker

Update the row for an item whenever its status changes (see CLAUDE.md → "Finishing an item").
Status values: `todo` · `in progress` · `in review` (branch pushed, not merged) · `live` · `blocked (reason)`.

Questions answered = how many of the ~220 real Australian Google questions in
`docs/research/` the item turns from "partly/no" into "yes".

Baseline (2026-10-07): live at moneymoves-au.vercel.app, `master` = `ad0356d`, `npm test` 40/40.

## Phase 0 — Ready to launch (blocks paid marketing)

| ID | Item | Owner | Status | Branch | Commit | Tests added | Updated | Notes / not tested |
|---|---|---|---|---|---|---|---|---|
| P0-1 | One one-off price everywhere (landing, app, terms, FAQ, `api/config.js`) | Claude, after price decision | in progress (local branch, not pushed) | `p0-1-one-price` | `b6020cd` | `test/pricing.test.js` (23); 2 price tests in `payment.test.js` updated to new prices. `npm test` 63/63 | 2026-10-07 | Prices: plan A$49 (no toolkits), complete A$159 (toolkits), both one-off; payment-file edits approved 2026-10-07. Browser-checked on local server (real `api/config.js`, PayPal sandbox `sb`) at 1280px + 390px: index, app paid section, terms. **Not tested:** a real sandbox/live PayPal payment; post-payment screen in a browser (covered only by the capture-order handler test); Vercel preview; Resend email delivery. Free beta PDF (`api/free-report.js`) still includes the toolkits. |
| P0-2 | Hide paid tier until live PayPal works | Claude | todo | | | | | |
| P0-3 | Production Clerk instance + keys | User | todo | | | | | Removes "Development mode" |
| P0-4 | Live PayPal credentials (`PAYPAL_ENV=live`) | User | todo | | | | | Then one real purchase test |
| P0-5 | ConvertAPI credit above 250 conversions | User | todo | | | | | |
| P0-6 | Upstash rate limiting in production | User (account) + Claude (wire-up check) | todo | | | | | |
| P0-7 | Privacy policy lists Clerk, ConvertAPI, Firebase | Claude | todo | | | | | |
| P0-8 | Lawyer check: personal vs general advice (AFSL) | User | todo | | | | | |
| P0-9 | Fix pre-existing CI failures (nanoid, e2e "14", Stryker) | Claude | todo | | | | | |

## Phase 1 — Quick wins (~58 questions)

| ID | Item | Questions | Status | Branch | Commit | Tests added | Updated | Notes / not tested |
|---|---|---|---|---|---|---|---|---|
| P1-1 | Shared money engine `lib/plan-engine.js` used by app + PDF | foundation | todo | | | | | Must not move golden figures |
| P1-2 | CSV / Excel download of cash map + debt schedule | ~5 | todo | | | | | |
| P1-3 | "Is my debt a lot?" health check card | ~27 | todo | | | | | |
| P1-4 | Snowball vs avalanche side by side | ~8 | todo | | | | | |
| P1-5 | Debt-free by a chosen date | ~13 | todo | | | | | |
| P1-6 | Lump sum: where does a one-off amount go | ~5 | todo | | | | | |

## Phase 2 — Deeper Australian tools (~21 questions)

| ID | Item | Questions | Status | Branch | Commit | Tests added | Updated | Notes / not tested |
|---|---|---|---|---|---|---|---|---|
| P2-1 | Balance transfer + consolidation calculator | ~10 | todo | | | | | |
| P2-2 | HECS repayment estimator (`lib/au-rates.js`) | ~5 | todo | | | | | Yearly data |
| P2-3 | Low-income / shortfall mode | ~6 | todo | | | | | |

## Phase 3 — Growth & retention (~17 questions)

| ID | Item | Questions | Status | Branch | Commit | Tests added | Updated | Notes / not tested |
|---|---|---|---|---|---|---|---|---|
| P3-1 | Mortgage extra repayments + offset calculator | ~13 | todo | | | | | |
| P3-2 | Opt-in progress tracker + monthly check-in email | ~4 | todo | | | | | Changes "nothing is stored" copy — needs approval |

## Phase 4 — Content & SEO (ongoing)

| ID | Item | Status | Branch | Commit | Updated | Notes |
|---|---|---|---|---|---|---|
| P4-1 | Landing pages for questions already answered | todo | | | | |
| P4-2 | Articles: stats (27 Qs) and credit scores (12 Qs), sourced | todo | | | | Every figure needs a source |

## Change log
| Date | Item | What changed | By |
|---|---|---|---|
| 2026-10-07 | — | Roadmap created from `docs/research/` | Claude |
| 2026-10-07 | P0-1 | One-off prices A$49 / A$159 from one server list; monthly wording removed (`b6020cd`) | Claude |
