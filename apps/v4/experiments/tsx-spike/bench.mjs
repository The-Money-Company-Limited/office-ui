// Transpiler choice, measured in headless Chrome on the same 70 source files:
//   node apps/v4/experiments/tsx-spike/bench.mjs
// For each transpiler: download + start-up time, then the time to transform
// every source file, in a fresh browser context, three times.
import { cp, mkdir, readFile, rm, stat, writeFile } from "node:fs/promises"
import { gzipSync } from "node:zlib"
import path from "node:path"
import { createRequire } from "node:module"
import { fileURLToPath, pathToFileURL } from "node:url"
import * as esbuild from "esbuild"

import { serve } from "./serve.mjs"

const here = path.dirname(fileURLToPath(import.meta.url))
const require = createRequire(import.meta.url)
const bench = path.join(here, "out/bench")
await rm(bench, { recursive: true, force: true })
await mkdir(bench, { recursive: true })
await cp(path.join(here, "out/optionb/src"), path.join(bench, "src"), { recursive: true })
await cp(path.join(here, "out/optionb/loader-manifest.json"), path.join(bench, "loader-manifest.json"))
await writeFile(path.join(bench, "sucrase-entry.js"), 'export { transform } from "sucrase"\n')
await esbuild.build({
  entryPoints: [path.join(bench, "sucrase-entry.js")],
  bundle: true,
  format: "esm",
  minify: true,
  outfile: path.join(bench, "sucrase.js"),
  nodePaths: [path.join(here, "node_modules")],
})
const esbuildWasm = path.dirname(require.resolve("esbuild-wasm/package.json"))
await cp(path.join(esbuildWasm, "esm/browser.min.js"), path.join(bench, "esbuild.js"))
await cp(path.join(esbuildWasm, "esbuild.wasm"), path.join(bench, "esbuild.wasm"))
const sizes = {}
for (const f of ["sucrase.js", "esbuild.js", "esbuild.wasm"]) {
  const body = await readFile(path.join(bench, f))
  sizes[f] = { raw: body.length, gzip: gzipSync(body).length }
}
await writeFile(
  path.join(bench, "index.html"),
  `<!doctype html><script type="module">
const manifest = await (await fetch("./loader-manifest.json")).json()
const files = manifest.files.filter((f) => /\\.(tsx?|jsx)$/.test(f))
const sources = await Promise.all(files.map(async (f) => [f, await (await fetch("./" + f)).text()]))
const which = location.hash.slice(1)
const t0 = performance.now()
let transform
if (which === "sucrase") {
  const m = await import("./sucrase.js")
  transform = (f, s) => m.transform(s, { transforms: f.endsWith("x") ? ["typescript", "jsx"] : ["typescript"], jsxRuntime: "automatic", production: true, disableESTransforms: true, filePath: f }).code
} else {
  const m = await import("./esbuild.js")
  await m.initialize({ wasmURL: "./esbuild.wasm", worker: false })
  transform = async (f, s) => (await m.transform(s, { loader: f.endsWith("x") ? "tsx" : "ts", jsx: "automatic", format: "esm", target: "es2022" })).code
}
const t1 = performance.now()
let bytes = 0
for (const [f, s] of sources) bytes += (await transform(f, s)).length
const t2 = performance.now()
window.result = { startup: t1 - t0, transform: t2 - t1, files: sources.length, outBytes: bytes }
</script>`
)
const { chromium } = await import(pathToFileURL("/opt/cod1460/node_modules/playwright/index.mjs"))
const browser = await chromium.launch({ headless: true, args: ["--no-sandbox", "--disable-dev-shm-usage"] })
const { server, port } = await serve(bench)
const results = { sizes }
for (const which of ["sucrase", "esbuild", "sucrase", "esbuild", "sucrase", "esbuild"]) {
  const context = await browser.newContext()
  const page = await context.newPage()
  await page.goto(`http://127.0.0.1:${port}/#${which}`)
  await page.waitForFunction(() => window.result, null, { timeout: 60_000 })
  ;(results[which] ??= []).push(await page.evaluate(() => window.result))
  await context.close()
}
const median = (v) => [...v].sort((a, b) => a - b)[1]
for (const which of ["sucrase", "esbuild"])
  results[which + "Median"] = {
    startupMs: Math.round(median(results[which].map((r) => r.startup))),
    transformMs: Math.round(median(results[which].map((r) => r.transform))),
  }
console.log(JSON.stringify(results, null, 1))
await writeFile(path.join(here, "out/bench-results.json"), JSON.stringify(results, null, 1) + "\n")
server.close()
await browser.close()
