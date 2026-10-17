import { execFileSync } from "node:child_process"
import { readFileSync, realpathSync } from "node:fs"
import path from "node:path"
import { fileURLToPath } from "node:url"
import { defineConfig } from "vite"

const root = path.dirname(fileURLToPath(import.meta.url))
const appRoot = path.resolve(root, "../..")

// template.json names the bundle a release was staged from, so a platform that
// keeps several bundles can tell which one serves an artifact: this
// repository's commit and the commit of each engine it bundles.
const describe = (cwd: string) =>
  execFileSync("git", ["describe", "--always", "--dirty"], {
    cwd,
    encoding: "utf8",
  }).trim()

// An engine linked from a fork checkout is named by its commit; one installed
// from npm by its version.
function engineVersion(name: string) {
  const directory = realpathSync(path.join(appRoot, "node_modules", name))
  if (!directory.includes(`${path.sep}node_modules${path.sep}`)) {
    return describe(directory)
  }
  return JSON.parse(readFileSync(path.join(directory, "package.json"), "utf8"))
    .version as string
}

const version = describe(root)
const engines = Object.fromEntries(
  [
    "@extend-ai/react-docx",
    "@extend-ai/react-xlsx",
    "@extend-ai/react-pptx",
  ].map((name) => [name, engineVersion(name)])
)

export default defineConfig({
  root,
  base: "./",
  publicDir: false,
  plugins: [
    {
      name: "template-version",
      generateBundle() {
        this.emitFile({
          type: "asset",
          fileName: "template.json",
          source:
            JSON.stringify({ template: "office", version, engines }) + "\n",
        })
      },
    },
  ],
  resolve: {
    // The engine workspace packages share one React instance with the app.
    dedupe: ["react", "react-dom"],
    alias: [
      {
        find: "@/lib/pdf-thumbnail-utils",
        replacement: path.join(root, "src/pdf-engine.ts"),
      },
      { find: "@", replacement: appRoot },
    ],
  },
  worker: {
    format: "es",
  },
  build: {
    outDir: path.join(root, "dist"),
    emptyOutDir: true,
  },
})
