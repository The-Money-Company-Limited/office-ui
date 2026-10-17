import assert from "node:assert/strict"
import { execFileSync, spawnSync } from "node:child_process"
import { cp, mkdir, readFile, writeFile } from "node:fs/promises"
import path from "node:path"
import { pathToFileURL } from "node:url"

const app = path.resolve("apps/v4")
const results = path.join(app, "components/extend/slides/tests/results")
const fixture = path.join(results, "baseline-source")
const baseline = path.join(results, "baseline-bundle")
await mkdir(fixture, { recursive: true })
const original = execFileSync(
  "git",
  ["show", "HEAD:apps/v4/experiments/artifact-template/src/main.tsx"],
  { encoding: "utf8" }
)
const { default: ts } = await import(
  pathToFileURL(path.join(app, "node_modules/typescript/lib/typescript.js"))
)
const entry = original
  .replace(
    '"../../../app/globals.css"',
    JSON.stringify(path.join(app, "app/globals.css"))
  )
  .replace(
    '"./theme.css"',
    JSON.stringify(
      path.join(app, "experiments/artifact-template/src/theme.css")
    )
  )
await writeFile(
  path.join(fixture, "main.js"),
  ts.transpileModule(entry, {
    compilerOptions: {
      jsx: ts.JsxEmit.ReactJSX,
      target: ts.ScriptTarget.ES2022,
      module: ts.ModuleKind.ESNext,
    },
  }).outputText
)

const index = await readFile(
  path.join(app, "experiments/artifact-template/index.html"),
  "utf8"
)
await writeFile(
  path.join(fixture, "index.html"),
  index.replace("/src/main.tsx", "./main.js")
)
const build = spawnSync(
  "node",
  [
    "node_modules/vite/bin/vite.js",
    "build",
    fixture,
    "--config",
    "components/extend/slides/tests/vite.config.ts",
    "--outDir",
    baseline,
  ],
  { cwd: app, encoding: "utf8" }
)
await writeFile(
  path.join(results, "baseline-build.log"),
  build.stdout + build.stderr
)
assert.equal(build.status, 0, "baseline build")
const stage = path.join(fixture, "stage.mjs")
await writeFile(
  stage,
  execFileSync("git", [
    "show",
    "HEAD:apps/v4/experiments/artifact-template/stage.mjs",
  ])
)
const staging = spawnSync(
  "node",
  [
    stage,
    path.join(app, "experiments/artifact-template/samples/deck"),
    path.join(results, "baseline-stage"),
  ],
  { encoding: "utf8" }
)
assert.notEqual(staging.status, 0)
assert.match(
  staging.stderr,
  /Only DOCX, XLSX, PPTX and PDF files are supported/
)
await writeFile(path.join(results, "baseline-stage.log"), staging.stderr)
for (const entry of ["artifact.json", "deck"])
  await cp(path.join(results, "bundle", entry), path.join(baseline, entry), {
    recursive: true,
  })
const browser = spawnSync(
  "sh",
  ["apps/v4/components/extend/slides/tests/run-browser.sh"],
  { env: { ...process.env, SLIDES_BUNDLE: baseline }, encoding: "utf8" }
)
await writeFile(
  path.join(results, "baseline-browser.log"),
  browser.stdout + browser.stderr
)
assert.equal(
  browser.status,
  1,
  "browser suite fails without slides integration"
)
const cases = JSON.parse(
  await readFile(path.join(results, "results.json"), "utf8")
)
assert.equal(cases.length, 4)
assert.ok(
  cases.every((test) => !test.pass),
  "every new test fails without the change"
)
await writeFile(
  path.join(results, "baseline-results.json"),
  JSON.stringify(cases, null, 2)
)
console.log(
  "PASS negative control: 4/4 browser cases fail without slides integration; original staging rejects the deck"
)
