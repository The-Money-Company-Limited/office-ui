import { execFileSync } from "node:child_process"
import path from "node:path"
import { fileURLToPath } from "node:url"
import { defineConfig } from "vite"

const root = path.dirname(fileURLToPath(import.meta.url))
const appRoot = path.resolve(root, "../..")

// template.json names the bundle a release was staged from, so a platform that
// keeps several bundles can tell which one serves an artifact.
const version = execFileSync("git", ["describe", "--always", "--dirty"], {
  cwd: root,
  encoding: "utf8",
}).trim()

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
          source: JSON.stringify({ template: "office", version }) + "\n",
        })
      },
    },
  ],
  resolve: {
    // The engine forks are linked from sibling checkouts; one React for all.
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
