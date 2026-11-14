import path from "node:path"
import { fileURLToPath } from "node:url"
import { defineConfig } from "vite"

import { templateJson } from "./template-version.mjs"

const root = path.dirname(fileURLToPath(import.meta.url))
const appRoot = path.resolve(root, "../..")

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
          source: templateJson([
            "@extend-ai/react-docx",
            "@extend-ai/react-xlsx",
            "@extend-ai/react-pptx",
          ]),
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
