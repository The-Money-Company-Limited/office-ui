import assert from "node:assert/strict"
import { execFileSync, spawnSync } from "node:child_process"
import { cp, mkdir, readFile, rm, writeFile } from "node:fs/promises"
import path from "node:path"
import { pathToFileURL } from "node:url"

const app = path.resolve("apps/v4")
const slides = "apps/v4/components/extend/slides"
const results = path.join(app, "components/extend/slides/tests/results")
const fixture = path.join(results, "baseline-source")
const baseline = path.join(results, "baseline-bundle")
const revision = "acb9dfce7fd6e9d0395364785753efbc763fe3bf"
const { default: ts } = await import(
  pathToFileURL(path.join(app, "node_modules/typescript/lib/typescript.js"))
)
await rm(fixture, { recursive: true, force: true })
await mkdir(fixture, { recursive: true })
for (const file of [
  "slides-editor.tsx",
  "slide-canvas.tsx",
  "geometry.ts",
  "slides.css",
  "export.ts",
  "validate.js",
]) {
  const source = execFileSync(
    "git",
    ["show", `${revision}:${slides}/${file}`],
    {
      encoding: "utf8",
    }
  )
  await writeFile(
    path.join(fixture, file.replace(/\.tsx?$/, ".js")),
    /\.tsx?$/.test(file)
      ? ts.transpileModule(source, {
          compilerOptions: {
            jsx: ts.JsxEmit.ReactJSX,
            target: ts.ScriptTarget.ES2022,
            module: ts.ModuleKind.ESNext,
          },
        }).outputText
      : source
  )
}
const config = path.join(fixture, "vite.config.mjs")
await writeFile(
  config,
  `
import focused from ${JSON.stringify(path.join(app, "components/extend/slides/tests/vite.config.ts"))}
export default {
  ...focused,
  resolve: {
    ...focused.resolve,
    alias: [
      { find: "@/components/extend/slides/slides-editor", replacement: ${JSON.stringify(path.join(fixture, "slides-editor.js"))} },
      ...focused.resolve.alias,
    ],
  },
  build: { ...focused.build, outDir: ${JSON.stringify(baseline)} },
}
`
)
const build = spawnSync(
  "node",
  ["node_modules/vite/bin/vite.js", "build", "--config", config],
  { cwd: app, encoding: "utf8" }
)
await writeFile(
  path.join(results, "baseline-build.log"),
  build.stdout + build.stderr
)
assert.equal(build.status, 0, "baseline build")
for (const entry of ["artifact.json", "deck"])
  await cp(path.join(results, "bundle", entry), path.join(baseline, entry), {
    recursive: true,
  })
const browser = spawnSync(
  "sh",
  ["apps/v4/components/extend/slides/tests/run-browser.sh"],
  {
    env: { ...process.env, SLIDES_BUNDLE: baseline, SLIDES_CASE: "polish " },
    encoding: "utf8",
  }
)
await writeFile(
  path.join(results, "baseline-browser.log"),
  browser.stdout + browser.stderr
)
console.log(browser.stdout)
assert.equal(
  browser.status,
  1,
  "regressions fail against the committed slides editor"
)
const cases = JSON.parse(
  await readFile(path.join(results, "results.json"), "utf8")
)
assert.equal(cases.length, 7)
assert.ok(
  cases.every((test) => !test.pass),
  "every regression fails without the fixes"
)
await writeFile(
  path.join(results, "baseline-results.json"),
  JSON.stringify(cases, null, 2)
)
console.log(
  "PASS negative control: all 7 polish regressions fail against the committed editor"
)
