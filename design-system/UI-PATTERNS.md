# MoneyMoves AU — UI Patterns

All tokens live in `design-system/tokens.css`. All components live in `design-system/components.css`. No raw hex colours, no hardcoded `px` font-sizes, no arbitrary spacing outside these files. Run `bash design-system/lint-tokens.sh` before merging any frontend change.

---

## Buttons

| Variant | When to use |
|---|---|
| `.btn` (primary, accent fill) | The single primary action on a page or modal — "Get my plan", "Pay now", "Continue" |
| `.btn.btn-secondary` (surface fill, border) | Secondary or cancel actions alongside a primary — "Back", "Edit", "Not now" |
| `.btn.btn-ghost` (no fill, text only) | Tertiary actions with no visual weight — "Learn more", inline links that look like buttons |
| `.btn.btn-destructive` (danger fill) | Irreversible destructive actions only — "Delete account", "Clear all data" |
| `.btn.btn-icon` | Icon-only controls (close modal, copy, etc.) — always pair with `aria-label` |

**Modifiers:**
- `.btn--sm` — compact rows, tag lists, inline table actions
- `.btn--full` — full-width on mobile CTAs and form cards

**States:**
- Every async action must set `disabled` and show a spinner or changed label while in flight.
- After success, update the label ("✓ Saved") and optionally revert after 2 s.
- On error, show the `.toast.toast-danger` component — never silently swallow errors.

**Never:**
- Use `<a>` styled as a button for an action that mutates state — use `<button>`.
- Apply `cursor: not-allowed` without also setting `disabled`.
- Use a custom hex or inline `background` on a button — all variants are in components.css.

---

## Inputs

```html
<div class="input-wrap">
  <span class="input-prefix">$</span>
  <input class="input" type="number" placeholder="0" />
</div>
<p class="input-error">Must be greater than zero</p>
```

- Prefix (`$`, `%`, icons) lives in `.input-prefix` inside `.input-wrap`.
- Validation errors use `.input-error` immediately below the field — never an alert box.
- Focus ring comes from `.input:focus` in components.css — do not override it inline.
- Radio pill groups use `.radio-pills` + `<label class="radio-pill">` — not custom `<div>` toggles.

---

## Cards

| Variant | When to use |
|---|---|
| `.card` (default) | Most content containers — section panels, form cards, result tiles |
| `.card.card-accent` | Highlight the primary recommendation or the paid CTA |
| `.card.card-success` | Confirmed / paid / completed states |
| `.card.card-warning` | Caution — close to a limit, soft-deadline content |
| `.card.card-danger` | Critical errors or destructive-action confirmations |

Cards must never set their own `background`, `border`, or `border-radius` — use variants.

---

## Page layout

Every page follows this shell:

```
<nav>            sticky, z-index var(--z-overlay)
<header>         page title, breadcrumb, or hero
<main>           all interactive content
<footer>         attribution, disclaimer, links
```

- Container: `<div class="container">` — max-width `var(--container-max)`, centered, padded `0 var(--space-6)`.
- Section padding: `var(--space-18) 0` desktop → `var(--space-14) 0` tablet → `var(--space-10) 0` mobile.
- Never nest `.container` inside another `.container`.

---

## Loading states

Every async action (API call, Firebase read, PDF generation) must show a loading state:

1. Disable the triggering button and replace its label with `<span class="spinner spinner--sm"></span> Loading…`
2. If the wait exceeds ~300 ms, show a full `.empty-state` with a large spinner on the content area.
3. On completion, remove the spinner and render the result — or show a toast on error.

**Never** leave the UI frozen without feedback.

---

## Error states

- **Field-level errors** → `.input-error` below the input.
- **Action errors** (API failure, network offline) → `.toast.toast-danger` — auto-dismiss after 5 s, but include a retry path.
- **Page-level errors** (content can't load) → `.empty-state` with an error icon, a plain-English message, and a "Try again" button.
- **Never** use `alert()` or `console.error()` as user-facing error handling.

---

## Empty states

Use `.empty-state` whenever a list, grid, or data section has no items to show:

```html
<div class="empty-state">
  <div class="empty-state-icon">📭</div>
  <h3 class="empty-state-title">No results yet</h3>
  <p class="empty-state-desc">Fill in your details above to see your plan.</p>
</div>
```

Always provide a title + description. Icon is optional but improves scannability.

---

## Toasts

```javascript
showToast('Plan saved!', 'success');   // success | warning | danger | info
showToast('Something went wrong', 'danger');
```

- Toasts stack at the bottom-right (`var(--z-toast)`).
- Auto-dismiss after 4 s.
- `danger` toasts do not auto-dismiss — require explicit close so the user can read the error.
- The `#toast-container` div must be present in the `<body>` of every page that calls `showToast`.

---

## Modals

Use `.modal-overlay` + `.modal` for confirmations and gated actions only — not for forms longer than 3 fields (use a new step/page instead).

Always:
- Trap focus inside the open modal.
- Close on `Escape` and on overlay click.
- Set `aria-modal="true"` and `role="dialog"` on `.modal`.

---

## Touch targets

All interactive elements must have a minimum touch target of `44px × 44px` (WCAG 2.5.5). Apply `min-height: 44px` on compact controls (icon buttons, small badges) that are visually smaller.

---

## Responsive breakpoints

| Breakpoint | Value | What changes |
|---|---|---|
| Tablet | `max-width: 768px` | Single-column grids, reduced section padding |
| Mobile | `max-width: 640px` | Full-width CTAs, stacked form rows, smaller hero |

Use `@media (max-width: NNNpx)` inside the page's `<style>` block. Never use inline `style=` for responsive overrides.

---

## Google Sign-In button

Always use the `.google-cta` (landing) or `.google-btn` (app/vote) class — never custom-style a Google button. The Google SVG colours (`#4285F4`, `#34A853`, `#FBBC05`, `#EA4335`) are explicitly allowed by the lint script.
