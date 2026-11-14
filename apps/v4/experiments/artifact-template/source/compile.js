// The one compile step, shared by the build (Node) and the page's loader (a
// browser worker). It is a pure function of the file's path and source, so
// its result can be cached under the file's ETag: sucrase strips types and
// JSX and turns ESM into CommonJS, and the result is a module factory plus the
// specifiers it requires. resolve.js links the specifiers on every load.
import { transform } from "sucrase"

import { codeFile, cssImports } from "./resolve.js"

export { candidates } from "./resolve.js"

export function compileFile(path, source) {
  // A stylesheet's entry keeps its text, so a CSS rebuild uses exactly the
  // revision its key was computed from.
  if (!codeFile.test(path)) return { code: source, specs: cssImports(source) }
  const transforms = ["imports"]
  if (/\.tsx?$/.test(path)) transforms.push("typescript")
  if (/\.(tsx|jsx)$/.test(path)) transforms.push("jsx")
  const code = transform(source, {
    transforms,
    jsxRuntime: "automatic",
    production: true,
    disableESTransforms: true,
    filePath: path,
  }).code.replace(/\bimport\.meta\.url\b/g, "__importMetaUrl")
  const specs = [
    ...new Set(
      [...code.matchAll(/\brequire\((['"])([^'"]+)\1\)/g)].map(
        (match) => match[2]
      )
    ),
  ]
  return {
    code: `function (require, module, exports, __importMetaUrl) {\n${code}\n}`,
    specs,
  }
}
