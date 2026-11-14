// Tailwind CSS from the site's modules, the same way in the build (Node) and
// in the page. The build ships its result as vendor/app-*.css; the loader
// runs it again, from vendor/tailwind-*.js, when the CSS's inputs changed.
import postcss from "postcss"
import nested from "postcss-nested"
import { compile } from "tailwindcss"

const dirname = (path) => path.slice(0, path.lastIndexOf("/"))
function join(dir, spec) {
  const parts = dir ? dir.split("/") : []
  for (const part of spec.split("/")) {
    if (part === "..") parts.pop()
    else if (part !== "." && part !== "") parts.push(part)
  }
  return parts.join("/")
}

// sheets: the stylesheets the modules import, in order; packages: bare CSS
// imports (tailwindcss) mapped to served paths; codes: the compiled modules;
// text(path) reads a served file. Every @import, in any sheet, is resolved
// here.
export async function buildCss(sheets, packages, codes, text) {
  const compiler = await compile(
    sheets.map((sheet) => `@import "${sheet}";`).join("\n"),
    {
      base: "",
      loadStylesheet: async (id, base) => {
        const path = packages[id] ?? join(base, id)
        return { path, base: dirname(path), content: await text(path) }
      },
      loadModule: async (id) => {
        throw new Error(`Tailwind plugin ${id} is not supported in the browser`)
      },
    }
  )
  // Like Tailwind's own scanner, any token may be a class; build() ignores
  // tokens that are not.
  const candidates = new Set()
  for (const code of codes)
    for (const token of code.split(/[\s"'`]+/)) if (token) candidates.add(token)
  // Tailwind writes variants as nested rules, which Chrome matches far more
  // slowly than flat ones (layout took about 7 times as long); flatten them.
  return postcss([nested]).process(compiler.build([...candidates]), {
    from: undefined,
  }).css
}
