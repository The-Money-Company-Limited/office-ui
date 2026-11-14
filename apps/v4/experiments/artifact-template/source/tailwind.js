// Tailwind CSS from the site's modules, the same way in the build (Node) and
// in the page. The build ships its result as vendor/app-*.css; the loader
// runs it again, from vendor/tailwind-*.js, only after a source file changed.
import postcss from "postcss"
import nested from "postcss-nested"
import { compile } from "tailwindcss"

const dirname = (path) => path.slice(0, path.lastIndexOf("/"))
function join(dir, spec) {
  const parts = dir.split("/")
  for (const part of spec.split("/")) {
    if (part === "..") parts.pop()
    else if (part !== "." && part !== "") parts.push(part)
  }
  return parts.join("/")
}

// css: {entry, packages, plain} from the loader manifest, all served paths.
// modules: the compiled modules' code; text(path) reads a served file.
export async function buildCss(css, modules, text) {
  const compiler = await compile(await text(css.entry), {
    base: dirname(css.entry),
    loadStylesheet: async (id, base) => {
      const path = css.packages[id] ?? join(base, id)
      return { path, base: dirname(path), content: await text(path) }
    },
    loadModule: async (id) => {
      throw new Error(`Tailwind plugin ${id} is not supported in the browser`)
    },
  })
  // Like Tailwind's own scanner, any token may be a class; build() ignores
  // tokens that are not.
  const candidates = new Set()
  for (const code of modules)
    for (const token of code.split(/[\s"'`]+/)) if (token) candidates.add(token)
  const plain = await Promise.all(css.plain.map(text))
  // Tailwind writes variants as nested rules, which Chrome matches far more
  // slowly than flat ones (layout took about 7 times as long); flatten them as
  // Vite's build does.
  const result = postcss([nested]).process(
    [compiler.build([...candidates]), ...plain].join("\n"),
    { from: undefined }
  )
  return result.css
}
