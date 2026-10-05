'use strict';

/**
 * Smoke tests — page loads + critical UI visible.
 * These run against a static file server (no Firebase, no API).
 * They verify the HTML skeleton is intact and auth gates are present.
 */

const { test, expect } = require('@playwright/test');

// ── Landing page ──────────────────────────────────────────────────────────────

test('landing: page loads with correct title', async ({ page }) => {
  await page.goto('/');
  await expect(page).toHaveTitle(/MoneyMoves/);
});

test('landing: Google CTA is visible', async ({ page }) => {
  await page.goto('/');
  // At least one .google-cta link with "Sign in with Google" text
  const cta = page.locator('.google-cta').first();
  await expect(cta).toBeVisible();
  await expect(cta).toContainText('Sign in with Google');
});

test('landing: pricing section exists', async ({ page }) => {
  await page.goto('/');
  // The A$14 paid plan is present in the DOM
  await expect(page.locator('text=14')).toBeVisible();
});

// ── App page ──────────────────────────────────────────────────────────────────

test('app: auth overlay is shown when not signed in', async ({ page }) => {
  await page.goto('/app.html');
  // #auth-overlay has display:flex by default; JS hides it on sign-in
  const overlay = page.locator('#auth-overlay');
  await expect(overlay).toBeVisible();
});

test('app: Google sign-in button is present in auth overlay', async ({ page }) => {
  await page.goto('/app.html');
  await expect(page.locator('#auth-google-btn')).toBeVisible();
  await expect(page.locator('#auth-google-btn')).toContainText('Continue with Google');
});

// ── Vote page ─────────────────────────────────────────────────────────────────

test('vote: page loads with correct title', async ({ page }) => {
  await page.goto('/vote.html');
  await expect(page).toHaveTitle(/What should we build next/);
});

test('vote: heading is visible', async ({ page }) => {
  await page.goto('/vote.html');
  await expect(page.locator('h1')).toContainText('What should we build next');
});
