# Slides verification

Run builds, typechecks and browser tests on Modal from
`~/Desktop/Code/office-remote`. The browser runner uses the existing `.cod1460`
Playwright, Chromium and shared-library setup inside `/work/office-ui`.
JSZip resolves from `apps/v4/package.json`; there is no nested slides package.

```sh
./r up
./r sync office-ui
./r run office-ui 'cd apps/v4 && node node_modules/eslint/bin/eslint.js components/extend/slides/*.js components/extend/slides/*.ts components/extend/slides/*.tsx components/extend/slides/tests/*.mjs components/extend/slides/tests/*.ts experiments/artifact-template/src/main.tsx && node node_modules/typescript/bin/tsc --noEmit -p .'
./r run office-ui 'cd apps/v4 && pnpm artifact-template:build > components/extend/slides/tests/results/artifact-build.log 2>&1; result=$?; tail -n 18 components/extend/slides/tests/results/artifact-build.log; exit "$result"'
./r run office-ui 'cd apps/v4 && node experiments/artifact-template/stage.mjs experiments/artifact-template/samples/deck components/extend/slides/tests/results/polish-release-2'
./r run office-ui 'SLIDES_BUNDLE=apps/v4/components/extend/slides/tests/results/polish-release-2 sh apps/v4/components/extend/slides/tests/run-browser.sh'
./r run office-ui 'node apps/v4/components/extend/slides/tests/stage.mjs'
./r run office-ui 'unzip -t apps/v4/components/extend/slides/tests/results/deck.pptx'
```

Staging refuses an existing output directory. Use a fresh release path for a
subsequent run. `tests/results/` contains ignored screenshots, ZIP/PPTX downloads,
the chart PNG extracted from the PPTX and JSON test results. The browser runner
starts and closes its own HTTP server and Chromium process.

The eleven browser cases cover:

- Drag, resize and keyboard nudge preserve every flow sibling's box. Nested flow
  groups and rotated titles are covered. Selection leaves flow untouched; one
  undo restores the original HTML and one redo restores the complete edit.
- The header contains the deck title, file/theme/download icons and Present.
  Undo/Redo are disabled icon buttons until available. The download menu opens
  with the keyboard, offers both formats, and restores focus on Escape.
- Native fullscreen exit returns to the displayed slide and updates its anchor
  and thumbnail. Fallback presentation blocks background wheel/key scrolling,
  supports arrows/space/PageUp/PageDown, and restores page scrolling on Escape.
- Sanitization, fixed canvas dimensions, hidden notes and unique element IDs.
- Rotated title drag, snap, nudge and proportional resize; text editing,
  undo/redo, duplication/deletion and toolbar history shortcuts.
- Window anchors, theme changes without scrolling, native fullscreen navigation
  and Escape returning to the displayed slide.
- ZIP edits survive reload. PPTX contains native text/shapes/image/table/rule
  objects, a rendered SVG raster and speaker notes; every XML part parses.
  This checks the archive, not desktop PowerPoint.

For focused iterations, build the real artifact entry point with unrelated
Office imports excluded:

```sh
./r run office-ui 'cd apps/v4 && node node_modules/vite/bin/vite.js build --config components/extend/slides/tests/vite.config.ts && node components/extend/slides/tests/prepare-bundle.mjs'
./r run office-ui 'SLIDES_BUNDLE=apps/v4/components/extend/slides/tests/results/bundle sh apps/v4/components/extend/slides/tests/run-browser.sh'
```

The negative control builds the committed slides implementation from
`acb9dfce7fd6e9d0395364785753efbc763fe3bf` in an isolated fixture, then requires
all seven `polish` regressions to fail. It never changes working source files,
branches or commits. Run after preparing the focused bundle above:

```sh
./r run office-ui 'node apps/v4/components/extend/slides/tests/baseline.mjs'
```

Negative-control build logs, browser logs and results are saved separately.
Run the positive browser command afterward to restore `results.json` to passing
results. Do not run the browser and negative-control scripts concurrently:
they share the results directory.

Verified on Modal on 2026-10-17: ESLint and whole-app TypeScript passed;
`artifact-template:build` passed (Vite reported directive/sourcemap and large-chunk
warnings). All 11 browser cases passed against `polish-release-2`. All seven
polish regressions failed against the committed editor. Staging/path rejection
passed, and `unzip -t` found no PPTX archive errors. The header was visually
inspected in both themes.
