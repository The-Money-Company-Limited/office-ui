// Import resolution and linking, shared by the build (Node) and the loader
// (the page). Paths are relative to the site root, such as
// src/apps/v4/components/extend/xlsx-editor.tsx.

const extensions = [
  ".ts",
  ".tsx",
  ".js",
  ".jsx",
  "/index.ts",
  "/index.tsx",
  "/index.js",
]
export const codeFile = /\.(tsx?|jsx?)$/

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

// One specifier written in `from`: {path, query} for a served file,
// {missing} for a local path that names no file in `exists`, {vendor} for an
// npm package. config: {aliases, prefixes}.
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
  // In a stylesheet, @import "extra.css" is a URL relative to it, found or
  // not yet (the loader then looks for it). A name with a slash
  // (tailwindcss/theme.css) is relative only when that file exists; anything
  // else (tailwindcss) is a package.
  if (
    base === null &&
    from.endsWith(".css") &&
    bare.endsWith(".css") &&
    !bare.startsWith("/")
  ) {
    const local = join(dirname(from), bare)
    if (!bare.includes("/") || exists(local))
      return exists(local) ? { path: local, query } : { missing: local }
  }
  if (base === null) {
    // vendor/ holds package modules only; a file URL (?url) needs an alias.
    if (query) throw new Error(`${spec}: an npm import with a query needs an alias`)
    return { vendor: bare }
  }
  const path = candidates(base).find(exists)
  return path ? { path, query } : { missing: base }
}

// Links the graph from `entry`. specsOf(path) is a file's import specifiers,
// or undefined while it is not compiled. Returns the modules in import order,
// each module's specifier-to-id map (an id is a path, "path?url",
// "path?css" or "~package"), the stylesheets modules import (in order), every
// stylesheet reached, the files still to compile (`needed`) and the local
// paths that name no file (`missing`, and `missingFrom`: one importer of each).
export function link(entry, specsOf, config, exists) {
  const modules = []
  const maps = {}
  const sheets = []
  const css = []
  const needed = new Set()
  const missing = new Set()
  const missingFrom = {}
  const seen = new Set()
  const visit = (path) => {
    if (seen.has(path)) return
    seen.add(path)
    const specs = specsOf(path)
    if (!specs) return needed.add(path)
    const map = {}
    for (const spec of specs) {
      const target = resolve(spec, path, config, exists)
      if (target.missing) {
        missing.add(target.missing)
        missingFrom[target.missing] ??= path
      }
      else if (target.vendor) map[spec] = "~" + target.vendor
      else {
        const isCss = target.path.endsWith(".css")
        map[spec] =
          target.path +
          (target.query ? "?" + target.query : isCss ? "?css" : "")
        if (
          isCss &&
          !target.query &&
          codeFile.test(path) &&
          !sheets.includes(target.path)
        )
          sheets.push(target.path)
        if (!target.query && (isCss || codeFile.test(target.path)))
          visit(target.path)
      }
    }
    if (path.endsWith(".css")) css.push(path)
    else {
      maps[path] = map
      modules.push(path)
    }
  }
  visit(entry)
  return {
    modules,
    maps,
    sheets,
    css,
    needed: [...needed],
    missing: [...missing],
    missingFrom,
  }
}

// A stylesheet's @import specifiers.
export const cssImports = (text) =>
  [...text.matchAll(/@import\s+["']([^"']+)["']/g)].map((match) => match[1])
