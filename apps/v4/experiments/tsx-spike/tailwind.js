// Tailwind in the browser, used only when served source no longer matches the
// source the shipped vendor/app.css was built from (an agent edited a file).
// Bundled to vendor/tailwind-compiler.js and loaded lazily by index.html.
import { compile } from "tailwindcss"

const dirname = (p) => p.slice(0, p.lastIndexOf("/"))
function join(dir, spec) {
  const parts = dir.split("/")
  for (const part of spec.split("/")) {
    if (part === "..") parts.pop()
    else if (part !== "." && part !== "") parts.push(part)
  }
  return parts.join("/")
}

// entry: a served CSS path such as src/apps/v4/app/globals.css.
// packages: bare CSS imports mapped to served paths.
export async function buildCss(entry, sources, packages, text) {
  const compiler = await compile(await text(entry), {
    base: dirname(entry),
    loadStylesheet: async (id, base) => {
      const path = packages[id] ?? join(base, id)
      return { path, base: dirname(path), content: await text(path) }
    },
    loadModule: async (id) => {
      throw new Error(`Tailwind plugin ${id} is not supported in the browser`)
    },
  })
  // Like Tailwind's own scanner, any token in the source may be a class;
  // build() ignores tokens that are not.
  const candidates = new Set()
  for (const source of sources) for (const token of source.split(/[\s"'`]+/)) if (token) candidates.add(token)
  return compiler.build([...candidates])
}
