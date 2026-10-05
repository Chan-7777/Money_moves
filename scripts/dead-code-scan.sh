#!/usr/bin/env bash
# Checks every api/*.js route is referenced from at least one source file,
# and flags dev-tool artifacts that have been tracked by git.
set -euo pipefail

ROOT="$(git rev-parse --show-toplevel 2>/dev/null || pwd)"
FAIL=0

echo "=== Dead-code scan ==="

# 1. API route reachability — each route must appear in HTML or non-api JS
for f in "$ROOT"/api/*.js; do
  route="$(basename "$f" .js)"
  if grep -rq "/api/$route" "$ROOT" \
       --include="*.html" --include="*.js" \
       --exclude-dir=node_modules --exclude-dir=".git" \
       --exclude-dir=api 2>/dev/null; then
    echo "OK    api/$route.js"
  else
    echo "DEAD  api/$route.js — not referenced outside api/"
    FAIL=1
  fi
done

# 2. Dev artifacts must not be tracked by git
DEV_ARTIFACTS=(
  build_spec.js
  build_test_matrix.py
  convert_report.py
  report_preview.html
)
for f in "${DEV_ARTIFACTS[@]}"; do
  if git -C "$ROOT" ls-files --error-unmatch "$f" 2>/dev/null; then
    echo "WARN  $f is git-tracked — add to .gitignore (dev artifact, not production code)"
    FAIL=1
  fi
done

if [ "$FAIL" -eq 0 ]; then
  echo "PASS"
fi
exit "$FAIL"
