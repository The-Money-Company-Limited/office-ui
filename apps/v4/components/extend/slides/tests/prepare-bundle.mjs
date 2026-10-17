import { cp, readFile, writeFile } from "node:fs/promises"
import { fileURLToPath } from "node:url"

import { validateManifest } from "../validate.js"

const bundle = new URL("./results/bundle/", import.meta.url)
const sample = new URL(
  "../../../../experiments/artifact-template/samples/deck/",
  import.meta.url
)
const manifest = validateManifest(
  JSON.parse(await readFile(new URL("deck.json", sample), "utf8"))
)
await cp(fileURLToPath(sample), fileURLToPath(new URL("deck", bundle)), {
  recursive: true,
})
await writeFile(
  new URL("artifact.json", bundle),
  JSON.stringify({
    kind: "slides",
    title: manifest.title,
    file: "deck/deck.json",
  })
)
