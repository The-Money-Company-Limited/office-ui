#!/bin/sh
set -eu
export LD_LIBRARY_PATH=/work/office-ui/.cod1460/libs/usr/lib/x86_64-linux-gnu:/work/office-ui/.cod1460/libs/lib/x86_64-linux-gnu
export TMPDIR=/work/office-ui/.cod1460/tmp
export FONTCONFIG_FILE=/work/office-ui/.cod1460/fonts.conf
export PLAYWRIGHT_BROWSERS_PATH=/work/office-ui/.cod1460/browsers
results=apps/v4/components/extend/slides/tests/results
mkdir -p "$results"
(cd apps/v4 && node node_modules/eslint/bin/eslint.js components/extend/slides/*.js components/extend/slides/*.ts components/extend/slides/*.tsx components/extend/slides/tests/*.mjs components/extend/slides/tests/*.ts && node node_modules/typescript/bin/tsc --noEmit -p .)
node apps/v4/components/extend/slides/tests/charts.mjs
node apps/v4/components/extend/slides/tests/chart-checker.mjs
node apps/v4/components/extend/slides/tests/checker.mjs > "$results/chart-checker.log"
(cd apps/v4 && pnpm artifact-template:build) > "$results/chart-artifact-build.log" 2>&1
node apps/v4/components/extend/slides/tests/stage.mjs
SLIDES_BUNDLE="$results/staging/valid" sh apps/v4/components/extend/slides/tests/run-browser.sh
cp "$results/results.json" "$results/chart-positive.json"
node apps/v4/components/extend/slides/tests/baseline.mjs --chart-review
cp "$results/chart-positive.json" "$results/results.json"
unzip -t "$results/editable-chart.pptx" > "$results/chart-archive.log"
printf '%s\n' 'PASS chart verification: lint, TypeScript, units, checker, artifact build, staging, browser, negative controls and PPTX archive'
