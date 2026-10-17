import assert from "node:assert/strict"
import { execFileSync, spawnSync } from "node:child_process"
import { cp, mkdir, readFile, rm, writeFile } from "node:fs/promises"
import path from "node:path"

const app = path.resolve("apps/v4")
const work = path.join(app, "components/extend/slides/tests/results/staging")
const source = path.join(app, "experiments/artifact-template/samples/deck")
const stage = path.join(app, "experiments/artifact-template/stage.mjs")
await rm(work, { recursive: true, force: true })
await mkdir(work, { recursive: true })
const valid = path.join(work, "valid")
execFileSync("node", [stage, source, valid], { stdio: "inherit" })
assert.deepEqual(
  JSON.parse(await readFile(path.join(valid, "artifact.json"))),
  {
    kind: "slides",
    title: "Cool Computers · A shared canvas",
    file: "deck/deck.json",
  }
)
await cp(source, path.join(work, "bad"), { recursive: true })
await writeFile(
  path.join(work, "bad/deck.json"),
  JSON.stringify({ title: "Unsafe", order: ["../../escape"] })
)
let result = spawnSync(
  "node",
  [stage, path.join(work, "bad"), path.join(work, "bad-output")],
  { encoding: "utf8" }
)
assert.notEqual(result.status, 0)
assert.match(result.stderr, /Invalid slide order or IDs/)
console.log("PASS staging and path rejection")
