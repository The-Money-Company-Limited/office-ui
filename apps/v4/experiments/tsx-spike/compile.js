// One compile step, shared by the build script (Node) and the service worker:
// TSX/TS source in, browser ESM out, with every import specifier rewritten to
// a relative URL the browser can fetch. Paths are relative to the release root.
import { transform } from "sucrase"
import { parse } from "es-module-lexer/js"

// The asm.js build of the lexer needs no init.
export const ready = Promise.resolve()

const extensions = [".ts", ".tsx", ".js", ".jsx", ".mjs", "/index.ts", "/index.tsx", "/index.js"]

function join(dir, spec) {
  const parts = dir ? dir.split("/") : []
  for (const part of spec.split("/")) {
    if (part === "..") parts.pop()
    else if (part !== "." && part !== "") parts.push(part)
  }
  return parts.join("/")
}

function dirname(path) {
  const i = path.lastIndexOf("/")
  return i < 0 ? "" : path.slice(0, i)
}

function relative(fromDir, to) {
  const from = fromDir ? fromDir.split("/") : []
  const target = to.split("/")
  let i = 0
  while (i < from.length && i < target.length - 1 && from[i] === target[i]) i++
  const up = from.length - i
  return (up ? "../".repeat(up) : "./") + target.slice(i).join("/")
}

function probe(base, exists) {
  if (exists(base)) return base
  for (const extension of extensions)
    if (exists(base + extension)) return base + extension
  if (base.endsWith(".js"))
    for (const extension of [".ts", ".tsx"])
      if (exists(base.slice(0, -3) + extension)) return base.slice(0, -3) + extension
  return null
}

// Resolve one specifier written in `from`. Returns {path, query} for a served
// file, {vendor} for a prebuilt dependency, or null.
export function resolve(spec, from, manifest, exists) {
  const q = spec.indexOf("?")
  const bare = q < 0 ? spec : spec.slice(0, q)
  const query = q < 0 ? "" : spec.slice(q + 1)
  if (manifest.aliases[bare]) return { path: manifest.aliases[bare], query }
  for (const [prefix, target] of manifest.prefixes)
    if (bare.startsWith(prefix)) {
      const path = probe(target + bare.slice(prefix.length), exists)
      return path ? { path, query } : null
    }
  if (bare.startsWith("./") || bare.startsWith("../")) {
    const path = probe(join(dirname(from), bare), exists)
    return path ? { path, query } : null
  }
  if (manifest.vendor[bare]) return { vendor: manifest.vendor[bare] }
  return null
}

export function transpile(path, source) {
  if (!/\.(tsx?|jsx)$/.test(path)) return source
  const transforms = []
  if (/\.tsx?$/.test(path)) transforms.push("typescript")
  if (/\.(tsx|jsx)$/.test(path)) transforms.push("jsx")
  return transform(source, {
    transforms,
    jsxRuntime: "automatic",
    production: true,
    disableESTransforms: true,
    filePath: path,
  }).code
}

// Rewrite specifiers; `onImport(spec, statement, dynamic, resolved)` lets the
// build script see every import.
export function rewrite(path, code, manifest, exists, onImport) {
  const [imports] = parse(code)
  let out = ""
  let last = 0
  for (const entry of imports) {
    if (entry.d === -2 || entry.n === undefined) continue // import.meta or computed
    const dynamic = entry.d > -1
    const resolved = resolve(entry.n, path, manifest, exists)
    onImport?.(entry.n, dynamic ? "" : code.slice(entry.ss, entry.se), dynamic, resolved)
    if (!resolved) continue // left as is; fails only if it runs
    let target
    if (resolved.vendor) target = relative(dirname(path), resolved.vendor)
    else {
      target = relative(dirname(path), resolved.path)
      const query = resolved.query || (resolved.path.endsWith(".css") ? "css" : resolved.path.endsWith(".json") ? "json" : "")
      if (query) target += "?" + query
    }
    out += code.slice(last, entry.s) + (dynamic ? JSON.stringify(target) : target)
    last = entry.e
  }
  return out + code.slice(last)
}

export function compile(path, source, manifest, exists, onImport) {
  return rewrite(path, transpile(path, source), manifest, exists, onImport)
}

// Registry form, for the loader that runs one cached script instead of one
// request per module: ESM -> CommonJS (sucrase "imports"), every require()
// rewritten to the resolved module id, wrapped in __define(id, factory).
// Ids: a served path (plus "?url"/"?css"), or "~" + a vendor specifier.
export function toRegistry(path, source, manifest, exists) {
  const transforms = ["imports"]
  if (/\.tsx?$/.test(path)) transforms.push("typescript")
  if (/\.(tsx|jsx)$/.test(path)) transforms.push("jsx")
  let code = transform(source, {
    transforms,
    jsxRuntime: "automatic",
    production: true,
    disableESTransforms: true,
    filePath: path,
  }).code
  code = code.replace(/\brequire\((['"])([^'"]+)\1\)/g, (match, quote, spec) => {
    const resolved = resolve(spec, path, manifest, exists)
    if (!resolved) return match
    if (resolved.vendor) return `require(${JSON.stringify("~" + spec.split("?")[0])})`
    const query = resolved.query || (resolved.path.endsWith(".css") ? "css" : "")
    return `require(${JSON.stringify(resolved.path + (query ? "?" + query : ""))})`
  })
  code = code.replace(/\bimport\.meta\.url\b/g, "__importMetaUrl")
  return `__define(${JSON.stringify(path)}, function (require, module, exports, __importMetaUrl) {\n${code}\n});\n`
}
