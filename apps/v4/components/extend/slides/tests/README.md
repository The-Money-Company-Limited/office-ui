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

## Agent checker (COD-1474)

The CLI uses Chromium's DOM with `validate.js` and the artifact's `slides.css`.
The shared canvas reset makes the standalone render independent of app styles.
Contrast compares text paint against a second screenshot with text paint hidden,
including sibling backgrounds, SVG, embedded images and group opacity. It uses
the weakest glyph's strongest painted pixel (antialiased edges do not define the
color); highly patterned backgrounds still warrant inspecting the PNG. Text
geometry uses ink metrics rather than the font's unused ascent/descent space.

On a fresh Linux machine, install Playwright beside `slides-check.mjs` with
`npm install --no-save playwright`, then `npx playwright install --with-deps chromium`.
Keep `validate.js` and `slides.css` beside the command. In the shared Modal sandbox,
reuse the already-installed browser and libraries, without touching other repos:

```sh
cd ~/Desktop/Code/office-remote
./r up
./r sync office-ui
./r run office-ui 'mkdir -p apps/v4/components/extend/slides/node_modules && ln -sfn /work/office-ui/.cod1460/node_modules/playwright apps/v4/components/extend/slides/node_modules/playwright'
./r run office-ui 'export LD_LIBRARY_PATH=/work/office-ui/.cod1460/libs/usr/lib/x86_64-linux-gnu:/work/office-ui/.cod1460/libs/lib/x86_64-linux-gnu TMPDIR=/work/office-ui/.cod1460/tmp FONTCONFIG_FILE=/work/office-ui/.cod1460/fonts.conf PLAYWRIGHT_BROWSERS_PATH=/work/office-ui/.cod1460/browsers && node apps/v4/components/extend/slides/tests/checker.mjs --negative-control'
./r run office-ui 'cd apps/v4 && node node_modules/eslint/bin/eslint.js components/extend/slides/slides-check.mjs components/extend/slides/validate.js components/extend/slides/validate.d.ts components/extend/slides/tests/checker.mjs'
```

The boundary test invokes the actual CLI with the existing sample and
`tests/fixtures/checker` in both human and JSON modes, checks exit codes and
1920×1080 PNG dimensions, and saves their outputs under `tests/results/checker`.
The fixture covers dropped tags, attributes and CSS, overflow, ancestor clipping,
mid-word wrapping (including inline formatting), bounds/margins, overlapping text
boxes, small/large/SVG/layered/translucent text contrast and failed images.
Controls cover ordinary wrapping, nested inline text, notes, valid images,
shorthand CSS, and large text above 3:1. It also checks invalid manifests,
missing files, symlink escapes, continuation, and warning-only exit 0.
The negative control disables each of the nine checks in an isolated copy and
requires the corresponding assertion to fail; it never modifies working source.

Verified on Modal on 2026-10-17: sample deck — 4 slides, 0 errors, 31 margin
warnings, exit 0; fixture deck — 10 slides, 18 errors, 3 warnings, exit 1.
All requested diagnostics were reported, including numeric and right-to-left
mid-word wraps. Clean controls had no findings. All nine disabled-check controls
failed their corresponding assertion; treating canvas clipping as an error also
failed the warning-only assertion. ESLint, the focused artifact build and all 11
existing artifact browser cases passed (the build retained its existing Vite
warnings). Four sample PNGs were checked as 1920×1080; sample and contrast-fixture
screenshots were visually inspected. JSON, human summaries, PNGs and logs are in
`tests/results/checker/`.

### Margin review

Margin findings use the element's border box and the existing strict >1px
crossing tolerance. Only visible text-bearing groups participate; an outermost
finding suppresses findings for its descendants. Backgrounds, rules, SVG metadata
and hidden text do not create margin findings. `margin.html` now covers nested
cards/tables and decoration while retaining exactly one expected finding.
The checker test also covers all four edges at 0px, 1px and 1.25px crossings.

On Modal, the checker command above with `--negative-control` passed all detection,
boundary and negative-control cases, including removing each new filter. ESLint
passed. The final zero-warning sample assertion failed: the unchanged sample
still has 13 genuine margin warnings (previously 31). Its HTML uses 110–120px
padding and footers below y=952; those sample files are outside the permitted
`components/extend/slides/**` edit scope. The defect fixture reports 18 errors,
3 warnings, including exactly one margin finding. Both CLI outputs retain their
expected exit statuses (sample 0, fixture 1); the test suite exits 1 because the
sample layout still needs correction. See `results/checker/margin-verification.log`,
`sample.txt`, `fixture.txt` and `margin-boundaries.json` for the latest run.

A proposed sample-only layout patch is in
`results/checker-margin/sample-layout.patch`; it has not been applied to the
out-of-scope sample files. The preview copy passed the same CLI on Modal with
4 slides, 0 errors and 0 warnings (`results/checker/sample-preview.txt`). It changes
padding to 128px, fits the grid/table to 1664px and moves the two footers inside
the safe area. The preview HTML lives in `results/checker-margin/sample-preview`.
