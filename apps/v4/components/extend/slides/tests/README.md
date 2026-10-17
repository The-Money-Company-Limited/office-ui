# Slides verification

Run from `~/Desktop/Code/office-remote`. All builds, typechecks and browser tests
run on Modal. The browser runner uses the existing `.cod1460` Playwright,
Chromium and shared-library setup inside `/work/office-ui`.

```sh
./r up
./r sync office-ui
./r run office-ui 'cd apps/v4 && node node_modules/eslint/bin/eslint.js components/extend/slides/*.js components/extend/slides/*.ts components/extend/slides/*.tsx components/extend/slides/tests/*.mjs components/extend/slides/tests/*.ts experiments/artifact-template/src/main.tsx experiments/artifact-template/stage.mjs'
./r run office-ui 'cd apps/v4 && node node_modules/typescript/bin/tsc --noEmit -p .'
./r run office-ui 'cd apps/v4 && pnpm artifact-template:build'
./r run office-ui 'cd apps/v4 && node experiments/artifact-template/stage.mjs experiments/artifact-template/samples/deck components/extend/slides/tests/results/node20-release'
./r run office-ui 'SLIDES_BUNDLE=apps/v4/components/extend/slides/tests/results/node20-release sh apps/v4/components/extend/slides/tests/run-browser.sh'
./r run office-ui 'node apps/v4/components/extend/slides/tests/stage.mjs'
./r run office-ui 'unzip -t apps/v4/components/extend/slides/tests/results/deck.pptx'
```

Staging deliberately refuses an existing output directory. Use a fresh release
path on subsequent runs. `tests/results/` contains ignored screenshots, the
ZIP/PPTX downloads, the chart PNG extracted from the PPTX and JSON test results.
The browser runner starts and closes its own HTTP server and Chromium process.

The four browser cases check:

1. The artifact loads four slides, normalizes the fixed canvas, rejects script,
   iframe, event attributes and invalid CSS, logs slide/element warnings, hides
   notes and avoids duplicate IDs in the thumbnails.
2. A rotated title can be dragged, pinned, snapped, nudged and resized with its
   ratio intact; text can be edited, undone/redone, duplicated and deleted.
   History shortcuts also work with focus in the toolbar.
3. Slide anchors use window scrolling. Theme controls preserve scroll. Native
   fullscreen opens the selected slide; arrows, PageDown, space and Escape work.
4. Downloaded HTML reloads with edits intact. The PPTX has four well-formed slide
   XML files, native text/shapes/image/table/rule objects, an SVG raster with
   rendered content, and speaker notes. This checks the archive and XML; it does
   not launch desktop PowerPoint.

For quicker slides-only iterations, build the real artifact entry point while
excluding its unused DOCX/XLSX/PPTX/PDF imports:

```sh
./r run office-ui 'cd apps/v4 && node node_modules/vite/bin/vite.js build --config components/extend/slides/tests/vite.config.ts && node components/extend/slides/tests/prepare-bundle.mjs'
./r run office-ui 'SLIDES_BUNDLE=apps/v4/components/extend/slides/tests/results/bundle sh apps/v4/components/extend/slides/tests/run-browser.sh'
```

The negative control builds `HEAD`'s original artifact entry point in an isolated
fixture and runs the same four browser cases against it. It also runs the
original staging script against the sample deck. It never edits working source
files, switches branches or creates commits. Run before committing this feature,
after preparing the focused bundle above:

```sh
./r run office-ui 'node apps/v4/components/extend/slides/tests/baseline.mjs'
```

The baseline for this change is `f14e3720b344041a3de53771a98475cc00d27d9c`.
Negative-control logs/results are saved separately. Run the positive browser
command afterward to restore `results.json` to the passing results.

Verified on Modal on 2026-10-17: ESLint, the whole-app TypeScript check and the
full artifact-template build all passed. Staging/path rejection passed. All
four browser cases passed against the full staged release; all four failed
against the original entry point, and original staging rejected the deck.
`unzip -t` reported no compressed-data errors. Every PPTX XML part parsed
successfully. The presentation screenshot and exported chart raster were also
visually inspected.

The shared validator is plain JavaScript with TypeScript declarations. Staging
also passed with the repository's declared Node 20.5.1, using these exact Modal
commands (the output directory was fresh):

```sh
./r run office-ui 'cd apps/v4 && npm exec --yes --cache /work/office-ui/.slides-npm-cache --package=node@20.5.1 -- node --version'
./r run office-ui 'cd apps/v4 && npm exec --yes --cache /work/office-ui/.slides-npm-cache --package=node@20.5.1 -- node experiments/artifact-template/stage.mjs experiments/artifact-template/samples/deck components/extend/slides/tests/results/node20-release'
```

The final full-release browser run and ZIP integrity check passed again after
this compatibility change.
