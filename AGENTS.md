# MoneyMoves AU — Agent Standing Rules

## What this is
Australian personal finance decision tool. Vanilla HTML/CSS/JS frontend, Vercel serverless functions (Node 22), Firebase Realtime DB + Auth, PayPal JS SDK for one-time PDF payments, Resend for email, ConvertAPI for DOCX→PDF.

## Definition of done
A task is DONE only when it works end-to-end with real data, real services, and real failure handling — proven by the repo's green gate (`npm test`), not by claims. Report anything less as INCOMPLETE with what remains. **Overstating completion is the worst possible failure.**

## Green gate
Before handing off any work: `npm test`. No subsets, no shortcuts. All tests must pass.

Also run before reporting done on any frontend change:
- `bash design-system/lint-tokens.sh` — fails on raw hex/px values
- `bash scripts/dead-code-scan.sh` — fails on unreferenced API routes

The Stop hook enforces `npm test` + token lint automatically at turn end. The CI pipeline (`github/workflows/ci.yml`) runs all gates on every push.

## Money path — highest protection
The payment flow (`api/create-order.js` → `api/capture-order.js`) must never be touched without explicit approval in the current message. Rules:
- Price is authoritative **on the server** — one list, `TIERS` in `lib/validators.js` (`plan` 49.00, `complete` 159.00, both one-off AUD). `create-order.js` charges from it and refuses unknown tiers; never read an amount from `req.body`.
- Captured amount **must be validated** against `TIERS` (`validateCapturedAmount`) in `capture-order.js` before delivering the PDF; the paid tier decides whether toolkits are included.
- Pages get prices from `/api/config`; `test/pricing.test.js` fails if any page shows a different price.
- Do not add fallbacks, stubs, or `console.log` replacements to payment or email handlers.
- Any change to `api/capture-order.js`, `api/create-order.js`, or `api/free-report.js` requires running `npm test` and showing pass output before reporting done.

## Banned shortcuts (no exceptions without explicit per-message approval)
- Hardcoded return values where computation or DB queries belong
- Mock/placeholder data in production paths
- Empty or log-only catch blocks — failures must surface clearly
- Skipping, commenting out, or loosening test assertions to force green
- Duplicate files with V2/New/Fixed suffixes — edit the canonical file
- Features beyond what was asked; error handling for impossible scenarios
- `--no-verify` on commits

## Required behaviour
- If blocked (missing credentials, unclear requirement): STOP and state exactly what's blocking.
- Every completion report states: what was tested, exact command, what was NOT tested.
- When uncertain about a library or pattern: search GitHub issues and official docs; compare ≥2 alternatives.

## Key files
| File | Role |
|---|---|
| `api/create-order.js` | Creates PayPal order — price hardcoded here |
| `api/capture-order.js` | Captures payment, generates PDF, emails it |
| `api/free-report.js` | Generates free PDF for beta users |
| `api/capture-email.js` | Email capture for waitlist |
| `api/config.js` | Serves PayPal client ID to frontend |
| `build_pdf_report.js` | DOCX report generator (all 7 sections) |
| `design-system/tokens.css` | Single source of truth for all design tokens |
| `design-system/lint-tokens.sh` | Fails if any frontend file uses raw hex/px |
| `design-system/UI-PATTERNS.md` | Component usage conventions |

## Self-check before reporting done
1. Did I hardcode anything to make output look right?
2. Would this survive a new user with zero data, a network failure, and an invalid input?
3. If all env vars were missing right now, does the server return a clear 503 (not silently succeed)?
4. Did `npm test` pass — not just "I think it would pass"?
