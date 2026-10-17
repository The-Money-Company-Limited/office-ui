#!/bin/sh
set -eu
node --test apps/v4/experiments/artifact-template/dev/autosave.test.mjs apps/v4/experiments/artifact-template/dev/bridge.test.mjs
cd apps/v4
node node_modules/eslint/bin/eslint.js \
  experiments/artifact-template/src/*.ts experiments/artifact-template/src/*.tsx \
  experiments/artifact-template/dev/*.mjs \
  components/extend/docx-editor.tsx components/extend/xlsx-editor.tsx \
  components/extend/pdf-editor.tsx components/extend/pdf-editor-dialogs.tsx \
  components/extend/pdf-editor-panels.tsx \
  components/extend/slides/slides-editor.tsx components/extend/slides/slide-canvas.tsx \
  components/extend/slides/export.ts components/extend/slides/validate.js
printf '%s\n' 'PASS ESLint'
node node_modules/typescript/bin/tsc --noEmit -p .
printf '%s\n' 'PASS TypeScript'
pnpm artifact-template:build > experiments/artifact-template/dev/results/build.log 2>&1
printf '%s\n' 'PASS artifact-template:build'
