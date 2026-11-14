// The one compile step, shared by the build (Node) and the page's loader (a
// browser worker): a TSX/TS file in, a module the loader's runtime can run
// out. Paths are relative to the site root, such as
// src/apps/v4/components/extend/xlsx-editor.tsx.
//
// sucrase strips types and JSX and turns ESM into CommonJS; every require()
// is then rewritten to the module id it resolves to, and the code is wrapped
// in __define(id, factory). An id is a served path (plus "?url" or "?css"),
// or "~" plus an npm package name served from vendor/.
import { transform } from "sucrase"

const extensions = [
  ".ts",
  ".tsx",
  ".js",
  ".jsx",
  "/index.ts",
  "/index.tsx",
  "/index.js",
]

function join(dir, spec) {
  const parts = dir ? dir.split("/") : []
  for (const part of spec.split("/")) {
    if (part === "..") parts.pop()
    else if (part !== "." && part !== "") parts.push(part)
  }
  return parts.join("/")
}

const dirname = (path) => path.slice(0, Math.max(0, path.lastIndexOf("/")))

// The files a specifier without an extension may name, in the order they are
// tried. "./a.js" may also be "./a.ts", as TypeScript allows.
export function candidates(base) {
  const list = [base, ...extensions.map((extension) => base + extension)]
  if (base.endsWith(".js"))
    list.push(base.slice(0, -3) + ".ts", base.slice(0, -3) + ".tsx")
  return list
}

// Resolves one specifier written in `from`: {path, query} for a served file,
// {missing} for a local path that names no known file yet, {vendor} for an
// npm package. config: {aliases, prefixes}; exists(path) says whether a file
// is known.
export function resolve(spec, from, config, exists) {
  const q = spec.indexOf("?")
  const bare = q < 0 ? spec : spec.slice(0, q)
  const query = q < 0 ? "" : spec.slice(q + 1)
  if (config.aliases[bare]) return { path: config.aliases[bare], query }
  const prefix = config.prefixes.find(([start]) => bare.startsWith(start))
  const base = prefix
    ? prefix[1] + bare.slice(prefix[0].length)
    : bare.startsWith("./") || bare.startsWith("../")
      ? join(dirname(from), bare)
      : null
  if (base !== null) {
    const path = candidates(base).find(exists)
    return path ? { path, query } : { missing: base }
  }
  return { vendor: bare }
}

// Returns {code, files, missing, packages}: the served files it imports
// (modules, stylesheets, assets), the local paths it names that are not known
// files, and the npm packages it imports.
export function compileModule(path, source, config, exists) {
  const transforms = ["imports"]
  if (/\.tsx?$/.test(path)) transforms.push("typescript")
  if (/\.(tsx|jsx)$/.test(path)) transforms.push("jsx")
  const files = new Set()
  const missing = new Set()
  const packages = new Set()
  let code = transform(source, {
    transforms,
    jsxRuntime: "automatic",
    production: true,
    disableESTransforms: true,
    filePath: path,
  }).code
  code = code.replace(
    /\brequire\((['"])([^'"]+)\1\)/g,
    (match, quote, spec) => {
      const resolved = resolve(spec, path, config, exists)
      if (resolved.missing) {
        missing.add(resolved.missing)
        return match
      }
      if (resolved.vendor) {
        packages.add(resolved.vendor)
        return `require(${JSON.stringify("~" + resolved.vendor)})`
      }
      files.add(resolved.path)
      const query =
        resolved.query || (resolved.path.endsWith(".css") ? "css" : "")
      return `require(${JSON.stringify(resolved.path + (query ? "?" + query : ""))})`
    }
  )
  code = code.replace(/\bimport\.meta\.url\b/g, "__importMetaUrl")
  return {
    code: `__define(${JSON.stringify(path)}, function (require, module, exports, __importMetaUrl) {\n${code}\n});\n`,
    files: [...files],
    missing: [...missing],
    packages: [...packages],
  }
}
