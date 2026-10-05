#!/usr/bin/env bash
# lint-tokens.sh — fails if any frontend file uses raw hex colors or
# arbitrary px font-sizes outside of design-system/tokens.css.
#
# Usage:
#   bash design-system/lint-tokens.sh          # check all frontend files
#   bash design-system/lint-tokens.sh app.html # check one file
#
# Exit code: 0 = clean, 1 = violations found

set -euo pipefail

PASS=0
FAIL=1
violations=0

# ── Config ────────────────────────────────────────────────────────────────
# Files to scan (relative to project root)
if [[ $# -gt 0 ]]; then
  FILES=("$@")
else
  FILES=(index.html app.html vote.html)
fi

# Hex colors that are ALLOWED (Google brand colours, not ours to token-ise)
ALLOWED_HEX="#dadce0|#bbb|#4285F4|#34A853|#FBBC05|#EA4335"

# Font sizes that are ALLOWED raw (inside clamp(), or truly unique to one place)
# We allow clamp() expressions and sizes that have no token equivalent (10px sticker labels)
ALLOWED_FONT_SIZES="10px"

# ── Helpers ───────────────────────────────────────────────────────────────
red()   { printf '\033[0;31m%s\033[0m\n' "$*"; }
green() { printf '\033[0;32m%s\033[0m\n' "$*"; }
yellow(){ printf '\033[0;33m%s\033[0m\n' "$*"; }
bold()  { printf '\033[1m%s\033[0m\n' "$*"; }

check_file() {
  local file="$1"
  local file_violations=0

  if [[ ! -f "$file" ]]; then
    yellow "  SKIP  $file (not found)"
    return
  fi

  bold "Checking $file…"

  # ── 1. Raw hex colors in CSS context ────────────────────────────────
  # Match #xxx or #xxxxxx that are NOT in the allowed list and NOT inside
  # base64 data URIs, NOT in HTML attribute values for non-style attributes,
  # NOT inside <!-- comments -->.
  #
  # Strategy: extract only <style> blocks and style="..." attributes, then
  # grep those. For simplicity we grep the whole file but exclude known-ok lines.

  local hex_pattern='#([0-9a-fA-F]{3}){1,2}\b'
  local hex_violations
  hex_violations=$(grep -nEo "$hex_pattern" "$file" \
    | grep -vE "$ALLOWED_HEX" \
    | grep -vE "base64|data:image|<!--" \
    | grep -vE "^[^:]+:[0-9]+:#[0-9a-fA-F]{3,6}" \
    || true)

  # Filter to lines actually inside <style> or style= contexts
  local style_hex
  style_hex=$(grep -nE "$hex_pattern" "$file" \
    | grep -vE "($ALLOWED_HEX)" \
    | grep -vE "(base64|data:image|<!--)" \
    | grep -E "(<style|style=|background:|color:|border:|box-shadow:|fill:)" \
    | grep -vE "($ALLOWED_HEX)" \
    || true)

  if [[ -n "$style_hex" ]]; then
    red "  [FAIL] Raw hex colors found:"
    echo "$style_hex" | head -20 | sed 's/^/         /'
    ((file_violations += $(echo "$style_hex" | wc -l)))
  fi

  # ── 2. Hardcoded px font-sizes ───────────────────────────────────────
  # Match font-size: NNpx (not inside clamp, not 10px which is allowed for tiny labels)
  local fontsize_pattern='font-size\s*:\s*[0-9]+px'
  local px_fonts
  px_fonts=$(grep -nE "$fontsize_pattern" "$file" \
    | grep -vE "clamp\(" \
    | grep -vE "($ALLOWED_FONT_SIZES)" \
    | grep -vE "<!-- " \
    || true)

  if [[ -n "$px_fonts" ]]; then
    red "  [FAIL] Hardcoded px font-sizes found:"
    echo "$px_fonts" | head -20 | sed 's/^/         /'
    ((file_violations += $(echo "$px_fonts" | wc -l)))
  fi

  # ── 3. Old token variable names ──────────────────────────────────────
  local old_vars_pattern='var\(--(green|green-dark|green-soft|gold|ink|ink-soft|bg|card|line|red|red-soft|amber|amber-soft)\)'
  local old_vars
  old_vars=$(grep -nE "$old_vars_pattern" "$file" || true)

  if [[ -n "$old_vars" ]]; then
    red "  [FAIL] Old design token variable references found:"
    echo "$old_vars" | head -20 | sed 's/^/         /'
    ((file_violations += $(echo "$old_vars" | wc -l)))
  fi

  # ── Report ────────────────────────────────────────────────────────────
  if [[ $file_violations -eq 0 ]]; then
    green "  [PASS] $file — no violations"
  else
    red "  $file — $file_violations violation(s)"
    ((violations += file_violations))
  fi

  echo ""
}

# ── Run ───────────────────────────────────────────────────────────────────
bold "MoneyMoves AU — Design Token Lint"
bold "=================================="
echo ""

for f in "${FILES[@]}"; do
  check_file "$f"
done

# ── Summary ───────────────────────────────────────────────────────────────
if [[ $violations -eq 0 ]]; then
  green "All files clean. Zero token violations."
  exit $PASS
else
  red "Total violations: $violations"
  echo ""
  echo "Fix all violations before merging. Each raw hex or px font-size"
  echo "should be replaced with a var(--color-*), var(--text-*), or"
  echo "var(--space-*) token from design-system/tokens.css."
  exit $FAIL
fi
