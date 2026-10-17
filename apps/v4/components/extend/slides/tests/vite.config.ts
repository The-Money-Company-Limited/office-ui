import { fileURLToPath } from "node:url"
import { mergeConfig } from "vite"

import template from "../../../../experiments/artifact-template/vite.config"

// Exercise the real artifact entry point without building unrelated Office engines.
export default mergeConfig(template, {
  build: {
    outDir: fileURLToPath(new URL("./results/bundle", import.meta.url)),
    rollupOptions: {
      external: [
        "@/components/extend/docx-editor",
        "@/components/extend/pptx-viewer",
        "@/components/extend/pdf-editor",
        "./xlsx-artifact",
      ],
    },
  },
})
