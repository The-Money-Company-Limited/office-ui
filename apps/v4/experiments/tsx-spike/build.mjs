// Spike build, run from the repository root on the build box after
// `pnpm build` and `pnpm --filter v4 artifact-template:build`:
//   node apps/v4/experiments/tsx-spike/build.mjs
// Writes out/fixture/budget.xlsx, out/baseline (today's bundle, staged) and
// out/optionb (real TSX source + prebuilt vendor ESM + service-worker loader).
import { execFileSync } from "node:child_process"
import { cp, mkdir, readdir, readFile, rm, stat, writeFile } from "node:fs/promises"
import { realpathSync, statSync } from "node:fs"
import { createRequire } from "node:module"
import path from "node:path"
import { fileURLToPath } from "node:url"
import * as esbuild from "esbuild"
import { compile, ready } from "./compile.js"

await ready
const here = path.dirname(fileURLToPath(import.meta.url))
const repo = path.resolve(here, "../../../..")
const v4 = path.join(repo, "apps/v4")
const template = path.join(v4, "experiments/artifact-template")
const xlsxPackage = path.join(repo, "engines/react-xlsx/packages/react-xlsx")
const out = path.join(here, "out")
const release = path.join(out, "optionb-sw") // ESM modules through a service worker
const registryRelease = path.join(out, "optionb") // one cached script (loader.js)
const rel = (p) => path.relative(repo, p).split(path.sep).join("/")
const src = (p) => "src/" + rel(p)

// 1. A workbook with two sheets and formulas.
const JSZip = createRequire(path.join(v4, "package.json"))("jszip")
await mkdir(path.join(out, "fixture"), { recursive: true })
{
  const zip = new JSZip()
  const sheets = ["Budget", "Notes"]
  zip.file(
    "[Content_Types].xml",
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("")}</Types>`
  )
  zip.file(
    "_rels/.rels",
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
  )
  zip.file(
    "xl/workbook.xml",
    `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${sheets.map((name, i) => `<sheet name="${name}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("")}</sheets></workbook>`
  )
  zip.file(
    "xl/_rels/workbook.xml.rels",
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join("")}</Relationships>`
  )
  // Budget: A=item, B=quantity, C=price, D=B*C; D202 = SUM(D2:D201).
  let rows = `<row r="1"><c r="A1" t="inlineStr"><is><t>Item</t></is></c><c r="B1" t="inlineStr"><is><t>Qty</t></is></c><c r="C1" t="inlineStr"><is><t>Price</t></is></c><c r="D1" t="inlineStr"><is><t>Total</t></is></c></row>`
  for (let r = 2; r <= 201; r++)
    rows += `<row r="${r}"><c r="A${r}" t="inlineStr"><is><t>Item ${r - 1}</t></is></c><c r="B${r}"><v>${r % 7 + 1}</v></c><c r="C${r}"><v>${(r * 3) % 50 + 1}</v></c><c r="D${r}"><f>B${r}*C${r}</f></c></row>`
  rows += `<row r="202"><c r="A202" t="inlineStr"><is><t>Sum</t></is></c><c r="D202"><f>SUM(D2:D201)</f></c></row>`
  zip.file(
    "xl/worksheets/sheet1.xml",
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:D202"/><sheetData>${rows}</sheetData></worksheet>`
  )
  zip.file(
    "xl/worksheets/sheet2.xml",
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:A1"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Second sheet</t></is></c></row></sheetData></worksheet>`
  )
  await writeFile(path.join(out, "fixture/budget.xlsx"), await zip.generateAsync({ type: "nodebuffer" }))
}

// 2. Baseline: today's bundle staged with the same workbook.
await rm(path.join(out, "baseline"), { recursive: true, force: true })
execFileSync("node", [path.join(template, "stage.mjs"), path.join(out, "fixture/budget.xlsx"), path.join(out, "baseline")], {
  cwd: v4,
  stdio: "inherit",
})

// 3. Walk the spreadsheet module graph with the same compile step the browser
// uses, collecting served files and every bare import with the names it uses.
const wasm = path.join(xlsxPackage, "vendor/sheets-wasm/duke_sheets_wasm_bg.wasm")
const manifest = {
  entry: src(path.join(template, "src/main.tsx")),
  aliases: {
    "@/lib/pdf-thumbnail-utils": src(path.join(template, "src/pdf-engine.ts")),
    "@extend-ai/react-xlsx": src(path.join(xlsxPackage, "src/index.ts")),
    "@extend-ai/react-xlsx/agent": src(path.join(xlsxPackage, "src/agent.ts")),
    "@extend-ai/react-xlsx/duke_sheets_wasm_bg.wasm": src(wasm),
  },
  prefixes: [["@/", src(v4) + "/"]],
  vendor: {},
  files: [],
  preload: [],
  css: "vendor/app.css",
}
// Modules of other document kinds, loaded lazily by main.tsx; not served.
const otherKinds = new Set(
  ["components/extend/docx-editor", "components/extend/pptx-viewer", "components/extend/slides/slides-editor", "components/extend/pdf-editor"].map(
    (p) => "@/" + p
  )
)
otherKinds.add("@extend-ai/react-docx/agent")
const exists = (p) => {
  if (!p.startsWith("src/")) return false
  try {
    return statSync(path.join(repo, p.slice(4))).isFile()
  } catch {
    return false
  }
}
async function walk(queue) {
const bare = new Map() // spec -> {names:Set, all:boolean, default:boolean, from:string}
const files = new Set()
let lines = 0
let sourceBytes = 0
while (queue.length) {
  const file = queue.pop()
  if (files.has(file)) continue
  files.add(file)
  if (file.endsWith(".css")) {
    // Follow relative @import so the browser Tailwind build can read them.
    const css = await readFile(path.join(repo, file.slice(4)), "utf8")
    for (const [, spec] of css.matchAll(/@import\s+["']([^"']+)["']/g))
      if (spec.startsWith(".")) queue.push(src(path.join(repo, file.slice(4), "..", spec)))
    continue
  }
  if (!/\.(tsx?|jsx|mjs|js)$/.test(file)) continue
  const text = await readFile(path.join(repo, file.slice(4)), "utf8")
  lines += text.split("\n").length
  sourceBytes += Buffer.byteLength(text)
  compile(file, text, manifest, exists, (spec, statement, dynamic, resolved) => {
    if (otherKinds.has(spec)) return
    if (resolved?.path) return queue.push(resolved.path)
    if (resolved) return
    if (spec.startsWith(".") || spec.startsWith("@/")) throw new Error(`${file}: cannot resolve ${spec}`)
    const entry = bare.get(spec) ?? { names: new Set(), all: false, default: false, from: file }
    bare.set(spec, entry)
    if (dynamic) return (entry.all = entry.default = true)
    const clause = statement
      .replace(/^\s*(import|export)\s*/, "")
      .replace(/\s*from\s*["'][^"']*["']\s*;?\s*$/s, "")
      .trim()
    if (/^["']/.test(clause) || clause === "") return // side effect
    if (clause.startsWith("*")) {
      entry.all = true
      entry.default = true
      return
    }
    const braces = clause.match(/\{([^}]*)\}/s)
    const head = clause.replace(/\{[^}]*\}/s, "").replace(/,/g, "").trim()
    if (head) entry.default = true
    if (braces)
      for (const part of braces[1].split(",")) {
        const name = part.trim().split(/\s+as\s+/)[0].trim()
        if (!name) continue
        if (name === "default") entry.default = true
        else entry.names.add(name)
      }
  })
}
return { bare, files, lines, sourceBytes }
}
const workerEntry = src(path.join(xlsxPackage, "src/xlsx-worker.ts"))
const { bare, files, lines, sourceBytes } = await walk([manifest.entry, workerEntry])
const workerGraph = await walk([workerEntry])
const sourceFiles = [...files].filter((f) => /\.(tsx?|jsx)$/.test(f))

// 4. Vendor for the service-worker variant: one ESM entry per bare specifier, split so shared code (React) is
// one copy. Explicit names keep the icon packs tree-shaken.
const entries = path.join(v4, ".tsx-spike-vendor")
await rm(entries, { recursive: true, force: true })
await mkdir(entries, { recursive: true })
const entryPoints = {}
for (const [spec, entry] of bare) {
  const name = spec.replace(/[^a-zA-Z0-9]+/g, "_")
  let names = [...entry.names]
  if (entry.all) {
    try {
      const require = createRequire(path.join(repo, entry.from.slice(4)))
      names = Object.keys(require(spec))
    } catch (error) {
      console.warn(`introspection failed for ${spec}: ${error.message}`)
      names = null
    }
  }
  const lines = []
  if (names === null) lines.push(`export * from ${JSON.stringify(spec)}`)
  else {
    const valid = names.filter((n) => /^[A-Za-z_$][\w$]*$/.test(n) && n !== "default" && n !== "__esModule")
    if (valid.length) lines.push(`export { ${valid.join(", ")} } from ${JSON.stringify(spec)}`)
  }
  if (entry.default) lines.push(`import * as ns from ${JSON.stringify(spec)}`, `export default ns.default ?? ns`)
  await writeFile(path.join(entries, name + ".js"), lines.join("\n") + "\n")
  entryPoints[name] = path.join(entries, name + ".js")
  manifest.vendor[spec] = `vendor/${name}.js`
}
await rm(release, { recursive: true, force: true })
const vendorResult = await esbuild.build({
  entryPoints,
  bundle: true,
  splitting: true,
  format: "esm",
  platform: "browser",
  minify: true,
  outdir: path.join(release, "vendor"),
  chunkNames: "chunk-[hash]",
  define: { "process.env.NODE_ENV": '"production"' },
  nodePaths: [path.join(v4, "node_modules"), path.join(xlsxPackage, "node_modules"), path.join(repo, "node_modules/.pnpm/node_modules")],
  loader: { ".json": "json" },
  metafile: true,
  logLevel: "warning",
})
manifest.preload = Object.keys(vendorResult.metafile.outputs).map((p) => rel(p).replace(rel(release) + "/", ""))

// 5. Source files, the calc core (prebuilt wasm + its JS glue) and CSS.
for (const file of files) {
  await mkdir(path.dirname(path.join(release, file)), { recursive: true })
  await cp(path.join(repo, file.slice(4)), path.join(release, file))
}
await mkdir(path.dirname(path.join(release, src(wasm))), { recursive: true })
await cp(wasm, path.join(release, src(wasm)))
manifest.files = [...files, src(wasm)].sort()
manifest.modules = [...manifest.files.filter((f) => /\.(tsx?|jsx|js)$/.test(f)), ...manifest.preload]
const assets = await readdir(path.join(template, "dist/assets"))
const css = assets.filter((a) => a.startsWith("index-") && a.endsWith(".css"))
if (css.length !== 1) throw new Error(`expected one CSS asset, found ${css}`)
await cp(path.join(template, "dist/assets", css[0]), path.join(release, "vendor/app.css"))

// Tailwind's own CSS and the compiler, for rebuilding after a source edit.
const tailwindDir = realpathSync(path.join(v4, "node_modules/tailwindcss"))
const animateDir = realpathSync(path.join(v4, "node_modules/tw-animate-css"))
await mkdir(path.join(release, "vendor/tw/tailwindcss"), { recursive: true })
for (const f of await readdir(tailwindDir))
  if (f.endsWith(".css")) await cp(path.join(tailwindDir, f), path.join(release, "vendor/tw/tailwindcss", f))
const animatePackage = JSON.parse(await readFile(path.join(animateDir, "package.json"), "utf8"))
const animateCss = animatePackage.style ?? animatePackage.exports?.["."]?.style ?? animatePackage.exports?.["."]?.default
await mkdir(path.join(release, "vendor/tw/tw-animate-css"), { recursive: true })
await cp(path.join(animateDir, animateCss), path.join(release, "vendor/tw/tw-animate-css/tw-animate.css"))
await esbuild.build({
  entryPoints: [path.join(here, "tailwind.js")],
  bundle: true,
  format: "esm",
  minify: true,
  outfile: path.join(release, "vendor/tailwind-compiler.js"),
  nodePaths: [path.join(v4, "node_modules")],
  logLevel: "warning",
})
{
  const { createHash } = await import("node:crypto")
  const paths = manifest.files.filter((f) => /\.(tsx?|jsx)$/.test(f))
  const texts = await Promise.all(paths.map((f) => readFile(path.join(repo, f.slice(4)), "utf8")))
  manifest.tailwind = {
    entry: src(path.join(v4, "app/globals.css")),
    packages: { tailwindcss: "vendor/tw/tailwindcss/index.css", "tw-animate-css": "vendor/tw/tw-animate-css/tw-animate.css" },
    plain: [src(path.join(template, "src/theme.css"))],
    sourceHash: createHash("sha1").update(paths.map((f, i) => f + "\0" + texts[i]).join("\0")).digest("hex"),
  }
}

// 6. Loader, service worker and the per-document files.
await esbuild.build({
  entryPoints: [path.join(here, "sw.js")],
  bundle: true,
  format: "iife",
  minify: true,
  outfile: path.join(release, "sw.js"),
  logLevel: "warning",
})
await cp(path.join(here, "index-sw.html"), path.join(release, "index.html"))
await writeFile(path.join(release, "loader-manifest.json"), JSON.stringify(manifest, null, 1) + "\n")
for (const f of ["artifact.json", "cool.json", "template.json", "document"])
  await cp(path.join(out, "baseline", f), path.join(release, f), { recursive: true })

// 7. The registry variant: same source, CSS and document; vendor as one
// script that sets self.__vendor, the compiler as one script, loader.js.
await rm(registryRelease, { recursive: true, force: true })
await cp(release, registryRelease, { recursive: true })
for (const f of await readdir(path.join(registryRelease, "vendor")))
  if (f.endsWith(".js") && f !== "tailwind-compiler.js") await rm(path.join(registryRelease, "vendor", f))
await rm(path.join(registryRelease, "sw.js"))
const vendorLines = []
const vendorObject = []
let index = 0
for (const [spec, entry] of bare) {
  const id = `m${index++}`
  if (spec.endsWith(".json")) {
    // Parsed on first use, not when the vendor script runs (the map atlases).
    vendorLines.push(`import ${id} from ${JSON.stringify(spec + "?text")}`)
    vendorObject.push(`${JSON.stringify(spec)}: { get default() { return (this._v ??= JSON.parse(${id})) } }`)
  } else if (entry.all || entry.default) {
    vendorLines.push(`import * as ${id} from ${JSON.stringify(spec)}`)
    vendorObject.push(`${JSON.stringify(spec)}: ${id}`)
  } else {
    const names = [...entry.names]
    vendorLines.push(`import { ${names.map((n, j) => `${n} as ${id}_${j}`).join(", ")} } from ${JSON.stringify(spec)}`)
    vendorObject.push(`${JSON.stringify(spec)}: { ${names.map((n, j) => `${n}: ${id}_${j}`).join(", ")} }`)
  }
}
await writeFile(path.join(entries, "registry-vendor.js"), vendorLines.join("\n") + `\nself.__vendor = { ${vendorObject.join(", ")} }\n`)
// The xlsx worker gets only the dependencies of its own graph.
{
  const keep = new Set(workerGraph.bare.keys())
  const specs = [...bare.keys()]
  const lines = vendorLines.filter((_, i) => keep.has(specs[i]))
  const object = vendorObject.filter((_, i) => keep.has(specs[i]))
  await writeFile(path.join(entries, "registry-vendor-worker.js"), lines.join("\n") + `\nself.__vendor = { ${object.join(", ")} }\n`)
}
await writeFile(path.join(entries, "registry-compiler.js"), `import { toRegistry } from ${JSON.stringify(path.join(here, "compile.js"))}\nself.tsxCompiler = { toRegistry }\n`)
const { createHash } = await import("node:crypto")
async function contentNamed(entry, name, options = {}) {
  const result = await esbuild.build({
    entryPoints: [entry],
    bundle: true,
    format: "iife",
    platform: "browser",
    minify: true,
    write: false,
    define: { "process.env.NODE_ENV": '"production"' },
    loader: { ".json": "json" },
    logLevel: "warning",
    plugins: [
      {
        name: "json-text",
        setup(build) {
          build.onResolve({ filter: /\.json\?text$/ }, async (args) => {
            const resolved = await build.resolve(args.path.slice(0, -5), { resolveDir: args.resolveDir, kind: args.kind })
            return { path: resolved.path, namespace: "json-text" }
          })
          build.onLoad({ filter: /.*/, namespace: "json-text" }, async (args) => ({
            contents: JSON.stringify(JSON.parse(await readFile(args.path, "utf8"))),
            loader: "text",
          }))
        },
      },
    ],
    ...options,
  })
  const body = result.outputFiles[0].contents
  const file = `vendor/${name}-${createHash("sha1").update(body).digest("hex").slice(0, 10)}.js`
  await writeFile(path.join(registryRelease, file), body)
  return file
}
const registryManifest = {
  entry: manifest.entry,
  aliases: manifest.aliases,
  prefixes: manifest.prefixes,
  vendor: Object.fromEntries([...bare.keys()].map((spec) => [spec, "~"])),
  files: manifest.files,
  vendorFile: await contentNamed(path.join(entries, "registry-vendor.js"), "vendor", {
    nodePaths: [path.join(v4, "node_modules"), path.join(xlsxPackage, "node_modules"), path.join(repo, "node_modules/.pnpm/node_modules")],
  }),
  workerVendorFile: await contentNamed(path.join(entries, "registry-vendor-worker.js"), "vendor-worker", {
    nodePaths: [path.join(v4, "node_modules"), path.join(xlsxPackage, "node_modules"), path.join(repo, "node_modules/.pnpm/node_modules")],
  }),
  workerModules: { [workerEntry]: [...workerGraph.files].filter((f) => /\.(tsx?|jsx|js)$/.test(f)).sort() },
  compilerFile: await contentNamed(path.join(entries, "registry-compiler.js"), "compiler", { nodePaths: [path.join(here, "node_modules")] }),
  css: manifest.css,
  tailwind: { ...manifest.tailwind, compiler: "vendor/tailwind-compiler.js" },
}
await cp(path.join(here, "index.html"), path.join(registryRelease, "index.html"))
await cp(path.join(here, "loader.js"), path.join(registryRelease, "loader.js"))
await writeFile(path.join(registryRelease, "loader-manifest.json"), JSON.stringify(registryManifest, null, 1) + "\n")

async function size(dir) {
  let total = 0
  let count = 0
  for (const entry of await readdir(dir, { withFileTypes: true, recursive: true }))
    if (entry.isFile()) {
      total += (await stat(path.join(entry.parentPath, entry.name))).size
      count++
    }
  return { count, total }
}
const summary = {
  sourceFiles: sourceFiles.length,
  sourceLines: lines,
  sourceBytes,
  servedFiles: files.size,
  bareSpecifiers: [...bare.keys()].sort(),
  vendorSw: await size(path.join(release, "vendor")),
  optionbSw: await size(release),
  optionb: await size(registryRelease),
  vendorFile: registryManifest.vendorFile,
  compilerFile: registryManifest.compilerFile,
  workerVendorFile: registryManifest.workerVendorFile,
  workerModules: Object.values(registryManifest.workerModules)[0].length,
  workerSpecifiers: [...workerGraph.bare.keys()],
  baseline: await size(path.join(out, "baseline")),
  baselineAssets: await size(path.join(out, "baseline/assets")),
}
await writeFile(path.join(out, "build-summary.json"), JSON.stringify(summary, null, 1) + "\n")
console.log(JSON.stringify(summary, null, 1))
