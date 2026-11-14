// Builds the spreadsheet template as source: dist-xlsx/ holds the page, the
// loader, the real TSX/TS files under src/ (at their repository paths), and
// what an agent never edits: the npm packages prebuilt in vendor/, the calc
// core's wasm, and the CSS for first paint. stage.mjs adds the document.
// Run by `pnpm --filter v4 artifact-template:build`.
import { createHash } from "node:crypto"
import { existsSync, realpathSync, statSync } from "node:fs"
import {
  copyFile,
  mkdir,
  readdir,
  readFile,
  rm,
  writeFile,
} from "node:fs/promises"
import path from "node:path"
import { fileURLToPath } from "node:url"
import * as esbuild from "esbuild"

import { templateJson } from "../template-version.mjs"
import { candidates, compileModule, resolve } from "./compile.js"
import { buildCss } from "./tailwind.js"

const here = path.dirname(fileURLToPath(import.meta.url))
const template = path.dirname(here)
const v4 = path.resolve(template, "../..")
const repo = path.resolve(v4, "../..")
const xlsx = path.join(repo, "engines/react-xlsx/packages/react-xlsx")
const out = path.join(template, "dist-xlsx")
const served = (file) =>
  "src/" + path.relative(repo, file).split(path.sep).join("/")
const onDisk = (file) => path.join(repo, file.slice("src/".length))

const entry = served(path.join(template, "src/main.tsx"))
const worker = served(path.join(xlsx, "src/xlsx-worker.ts"))
const config = {
  aliases: {
    "@extend-ai/react-xlsx": served(path.join(xlsx, "src/index.ts")),
    "@extend-ai/react-xlsx/agent": served(path.join(xlsx, "src/agent.ts")),
    "@extend-ai/react-xlsx/duke_sheets_wasm_bg.wasm": served(
      path.join(xlsx, "vendor/sheets-wasm/duke_sheets_wasm_bg.wasm")
    ),
  },
  prefixes: [["@/", served(v4) + "/"]],
}
// main.tsx and webmcp.ts load the other document kinds lazily; a spreadsheet
// site leaves them out.
const otherKinds = [
  "@/components/extend/docx-editor",
  "@/components/extend/pptx-viewer",
  "@/components/extend/slides/slides-editor",
  "@/components/extend/pdf-editor",
  "@extend-ai/react-docx/agent",
]
const left = otherKinds.map((spec) => resolve(spec, entry, config, () => false))
const absent = left.filter((r) => r.missing).map((r) => r.missing)
const absentFiles = new Set(absent.flatMap(candidates))
const skippedPackages = new Set(
  left.filter((r) => r.vendor).map((r) => r.vendor)
)
const exists = (file) =>
  !absentFiles.has(file) &&
  existsSync(onDisk(file)) &&
  statSync(onDisk(file)).isFile()

// The icon pack has thousands of icons: vendor/ holds the ones the source
// imports by name. Every other package is whole.
const iconPack = "@hugeicons/core-free-icons"
const iconImport =
  /import\s*\{([^}]*)\}\s*from\s*["']@hugeicons\/core-free-icons["']/g
const icons = new Set()
const typeImport = /(?:import|export)\s+type\b[^;]*?\bfrom\s*["']([^"']+)["']/g

// 1. Walk the graphs with the browser's compile step: the whole site's from
// the page and the worker it starts, and the worker's own.
const codeFile = /\.(tsx?|jsx?)$/
async function walk(start) {
  const files = new Set()
  const codes = new Map()
  const packages = new Map() // name -> a file that imports it
  const queue = [...start]
  while (queue.length) {
    const file = queue.pop()
    if (files.has(file)) continue
    files.add(file)
    if (!codeFile.test(file)) continue
    const source = await readFile(onDisk(file), "utf8")
    const result = compileModule(file, source, config, exists)
    for (const [, names] of source.matchAll(iconImport))
      for (const name of names.split(","))
        if (name.trim()) icons.add(name.trim().split(/\s+as\s+/)[0])
    for (const base of result.missing)
      if (!absent.includes(base))
        throw new Error(`${file} imports ${base}, which is not a file`)
    for (const name of result.packages)
      if (!skippedPackages.has(name) && !packages.has(name))
        packages.set(name, file)
    codes.set(file, result.code)
    queue.push(...result.files)
    // sucrase drops type-only imports; ship those files too, for an agent to read.
    for (const [, spec] of source.matchAll(typeImport)) {
      const resolved = resolve(spec, file, config, exists)
      if (resolved.path) queue.push(resolved.path)
    }
  }
  return { files, codes, packages }
}
const page = await walk([entry, worker])
const workerGraph = await walk([worker])

await rm(out, { recursive: true, force: true })
await mkdir(path.join(out, "vendor"), { recursive: true })
const write = async (file, body) => {
  await mkdir(path.dirname(path.join(out, file)), { recursive: true })
  await writeFile(path.join(out, file), body)
}
// Content-named, so the loader caches it once for good.
const named = async (name, extension, body) => {
  const file = `vendor/${name}-${createHash("sha256").update(body).digest("hex").slice(0, 12)}${extension}`
  await write(file, body)
  return file
}

// 2. Source files and assets at their served paths; stylesheets with the
// relative stylesheets they import, packages' (tailwindcss) from vendor/css.
const css = { entry: null, plain: [], packages: {} }
const sourceFiles = new Set(page.files)
const stylesheets = [...page.files].filter((file) => file.endsWith(".css"))
for (const sheet of stylesheets) {
  const text = await readFile(onDisk(sheet), "utf8")
  if (/@import\s+["']tailwindcss["']/.test(text)) {
    if (css.entry)
      throw new Error(`Two Tailwind entries: ${css.entry} and ${sheet}`)
    css.entry = sheet
  } else if (page.files.has(sheet)) css.plain.push(sheet)
  for (const [, spec] of text.matchAll(/@import\s+["']([^"']+)["']/g)) {
    if (spec.startsWith(".")) {
      const file = served(path.join(onDisk(sheet), "..", spec))
      if (!sourceFiles.has(file)) {
        sourceFiles.add(file)
        stylesheets.push(file)
      }
    } else if (!css.packages[spec])
      css.packages[spec] = await copyStylePackage(spec)
  }
}
if (!css.entry)
  throw new Error("No Tailwind entry stylesheet in the page's graph")
for (const file of sourceFiles) {
  await mkdir(path.dirname(path.join(out, file)), { recursive: true })
  await copyFile(onDisk(file), path.join(out, file))
}

// A CSS package's stylesheets, copied whole so its own relative imports work.
async function copyStylePackage(name) {
  const directory = realpathSync(path.join(v4, "node_modules", name))
  const pkg = JSON.parse(
    await readFile(path.join(directory, "package.json"), "utf8")
  )
  const style = pkg.style ?? pkg.exports?.["."]?.style ?? "index.css"
  const styleDir = path.dirname(path.join(directory, style))
  for (const file of await readdir(styleDir))
    if (file.endsWith(".css")) {
      await mkdir(path.join(out, "vendor/css", name), { recursive: true })
      await copyFile(
        path.join(styleDir, file),
        path.join(out, "vendor/css", name, file)
      )
    }
  return `vendor/css/${name}/${path.basename(style)}`
}

// 3. npm packages, prebuilt: one script for the page and one for the worker.
// Each sets self.__vendor to every package's namespace, so source edits can
// use any export (of the icon pack, the icons above). A package runs the
// first time a module requires it, so the chart libraries never run for a
// workbook without charts.
async function vendor(packages, name) {
  const lines = [
    "const namespace = (m) => (m && m.__esModule ? m : { ...m, default: m })",
  ]
  const fields = []
  for (const spec of packages.keys()) {
    const key = JSON.stringify(spec)
    if (spec === iconPack) {
      const list = [...icons]
      lines.push(
        `import { ${list.map((icon, i) => `${icon} as icon${i}`).join(", ")} } from ${key}`
      )
      fields.push(
        `${key}: { ${list.map((icon, i) => `${icon}: icon${i}`).join(", ")} }`
      )
    } else fields.push(`get ${key}() { return namespace(require(${key})) }`)
  }
  const result = await esbuild.build({
    stdin: {
      contents: `${lines.join("\n")}\nself.__vendor = { ${fields.join(", ")} }\n`,
      resolveDir: here,
      loader: "js",
    },
    bundle: true,
    format: "iife",
    platform: "browser",
    minify: true,
    write: false,
    define: { "process.env.NODE_ENV": '"production"' },
    logLevel: "warning",
    plugins: [
      {
        // Each package resolves from a file that imports it, as in the repository.
        name: "importer",
        setup(build) {
          build.onResolve({ filter: /.*/ }, (args) =>
            args.importer === "<stdin>"
              ? build.resolve(args.path, {
                  resolveDir: path.dirname(onDisk(packages.get(args.path))),
                  kind: args.kind,
                })
              : undefined
          )
        },
      },
    ],
  })
  return named(name, ".js", result.outputFiles[0].contents)
}

// 4. The two compilers the page loads only when it needs them.
async function bundle(contents, name, format) {
  const result = await esbuild.build({
    stdin: { contents, resolveDir: here, loader: "js" },
    bundle: true,
    format,
    platform: "browser",
    minify: true,
    write: false,
    logLevel: "warning",
  })
  return named(name, ".js", result.outputFiles[0].contents)
}

// 5. The CSS for first paint, built the way the page rebuilds it.
const text = (file) => readFile(path.join(out, file), "utf8")
const shippedCss = await buildCss(css, [...page.codes.values()], text)
const modules = [...page.files].filter((file) => codeFile.test(file)).sort()
const sources = await Promise.all(
  modules.map(
    async (file) => file + "\0" + (await readFile(onDisk(file), "utf8"))
  )
)

const manifest = {
  entry,
  workers: { [worker]: await vendor(workerGraph.packages, "vendor-worker") },
  config,
  files: [...sourceFiles].sort(),
  absent,
  vendor: await vendor(page.packages, "vendor"),
  compiler: await bundle(
    `import { candidates, compileModule } from "./compile.js"\nself.compiler = { candidates, compileModule }\n`,
    "compiler",
    "iife"
  ),
  css: {
    ...css,
    shipped: await named("app", ".css", shippedCss),
    compiler: await bundle(
      `export { buildCss } from "./tailwind.js"\n`,
      "tailwind",
      "esm"
    ),
    sourceHash: createHash("sha1").update(sources.join("\0")).digest("hex"),
  },
}
manifest.version = createHash("sha256")
  .update(
    JSON.stringify({ ...manifest, css: { ...manifest.css, sourceHash: "" } })
  )
  .digest("hex")
  .slice(0, 12)
await write("loader-manifest.json", JSON.stringify(manifest, null, 1) + "\n")

// 6. The page: the template's index.html, running the loader.
const html = await readFile(path.join(template, "index.html"), "utf8")
const script = '<script type="module" src="/src/main.tsx"></script>'
if (!html.includes(script))
  throw new Error(`index.html no longer has ${script}`)
await write(
  "index.html",
  html.replace(script, '<script type="module" src="./loader.js"></script>')
)
await copyFile(path.join(here, "loader.js"), path.join(out, "loader.js"))
await copyFile(path.join(here, "template.md"), path.join(out, "template.md"))
await write("template.json", templateJson(["@extend-ai/react-xlsx"]))
console.log(
  `dist-xlsx: ${modules.length} modules, ${page.packages.size} packages, version ${manifest.version}`
)
