// Build the pre-autosave entry point without switching branches or touching git state.
import assert from "node:assert/strict"
import { execFileSync, spawnSync } from "node:child_process"
import { readFile, writeFile } from "node:fs/promises"

const main = "apps/v4/experiments/artifact-template/src/main.tsx"
const original = await readFile(main)
try {
  await writeFile(main, execFileSync("git", ["show", `HEAD:${main}`]))
  execFileSync("pnpm", ["artifact-template:build"], {
    cwd: "apps/v4",
    stdio: "ignore",
  })
  execFileSync(
    "sh",
    ["apps/v4/experiments/artifact-template/dev/browser.sh", "--negative"],
    { stdio: "inherit" }
  )
} finally {
  await writeFile(main, original)
  execFileSync("sh", ["apps/v4/experiments/artifact-template/dev/verify.sh"], {
    stdio: "inherit",
  })
}
const unit = spawnSync(
  "node",
  ["--test", "apps/v4/experiments/artifact-template/dev/autosave.test.mjs"],
  {
    env: { ...process.env, AUTOSAVE_NEGATIVE: "1" },
    encoding: "utf8",
  }
)
console.log(unit.stdout)
assert.notEqual(unit.status, 0)
assert.match(unit.stdout, /# fail 5/)
console.log(
  "PASS negative controls: every browser and contract case fails without autosave wiring"
)
