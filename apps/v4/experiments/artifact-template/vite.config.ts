import path from "node:path"
import { fileURLToPath } from "node:url"
import { defineConfig } from "vite"

const root = path.dirname(fileURLToPath(import.meta.url))
const appRoot = path.resolve(root, "../..")

export default defineConfig({
  root,
  base: "./",
  publicDir: false,
  resolve: {
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
