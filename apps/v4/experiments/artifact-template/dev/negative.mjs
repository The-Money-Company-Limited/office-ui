// Build the pre-autosave entry point without switching branches or touching git state.
import assert from "node:assert/strict"
import { execFileSync, spawnSync } from "node:child_process"
import { readFile, writeFile } from "node:fs/promises"

const main = "apps/v4/experiments/artifact-template/src/main.tsx"
const original = await readFile(main)
try {
  await writeFile(main, execFileSync("git", ["show", `59895b6^:${main}`]))
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
assert.match(unit.stdout, /# fail 13/)
console.log(
  "PASS negative controls: every browser and contract case fails without autosave wiring"
)

const baseline = spawnSync(
  "node",
  ["--test", "apps/v4/experiments/artifact-template/dev/autosave.test.mjs"],
  {
    env: { ...process.env, AUTOSAVE_BASELINE: "1" },
    encoding: "utf8",
  }
)
console.log(baseline.stdout)
assert.notEqual(baseline.status, 0)
assert.match(baseline.stdout, /# fail 8/)
console.log(
  "PASS round-2 negative control: all eight new controller cases fail on 59895b6"
)

const bridge = "apps/v4/experiments/artifact-template/dev/autosave-harness.mjs"
const bridgeSource = await readFile(bridge, "utf8")
for (const [name, change] of [
  [
    "receipts precede",
    (source) =>
      source.replace(
        "if (receipts.has(data.id))",
        "if (false && receipts.has(data.id))"
      ),
  ],
  [
    "receipt retention",
    (source) =>
      source
        .replace("24 * 60 * 60_000", "48 * 60 * 60_000")
        .replace("receipts.size > 32", "receipts.size > 64"),
  ],
  [
    "accepts only|slides enforce",
    (source) =>
      source.replace(
        "if (!(await allowedPaths(artifact, files, folder)))",
        "if (false)"
      ),
  ],
  [
    "25 MiB",
    (source) =>
      source
        .replace("25 * 1024 * 1024", "100 * 1024 * 1024")
        .replace("64 * 1024 * 1024", "100 * 1024 * 1024"),
  ],
  [
    "malformed requests",
    (source) =>
      source.replace('return fail("error")', 'return fail("path_not_allowed")'),
  ],
  ["offline includes", (source) => source.replace("{ retryAfter }", "{}")],
]) {
  try {
    const changed = change(bridgeSource)
    assert.notEqual(changed, bridgeSource, name)
    await writeFile(bridge, changed)
    const result = spawnSync(
      "node",
      [
        "--test",
        `--test-name-pattern=${name}`,
        "apps/v4/experiments/artifact-template/dev/bridge.test.mjs",
      ],
      { encoding: "utf8" }
    )
    assert.notEqual(result.status, 0, `${name}: mutation unexpectedly passed`)
    assert.match(result.stdout, /not ok/, result.stderr)
    console.log(`PASS bridge negative control: ${name}`)
  } finally {
    await writeFile(bridge, bridgeSource)
  }
}
try {
  await writeFile(
    bridge,
    bridgeSource.replace(
      "if (!response.ok) return",
      "if (!response.ok) { send({ type: 'access', canSave: false, version: version || '1' }); return }"
    )
  )
  execFileSync(
    "sh",
    ["apps/v4/experiments/artifact-template/dev/browser.sh", "--negative"],
    {
      env: { ...process.env, AUTOSAVE_SCENARIOS: "failed-poll" },
      stdio: "inherit",
    }
  )
} finally {
  await writeFile(bridge, bridgeSource)
}
const validate = "apps/v4/components/extend/slides/validate.js"
const validateSource = await readFile(validate, "utf8")
try {
  await writeFile(
    validate,
    validateSource.replace(
      "/^[A-Za-z0-9_-]{1,64}$/",
      "/^[a-zA-Z][a-zA-Z0-9_-]{0,79}$/"
    )
  )
  const result = spawnSync(
    "node",
    [
      "--test",
      "--test-name-pattern=slide manifests",
      "apps/v4/experiments/artifact-template/dev/bridge.test.mjs",
    ],
    { encoding: "utf8" }
  )
  assert.notEqual(result.status, 0)
  assert.match(result.stdout, /not ok/)
  console.log("PASS slide ID negative control")
} finally {
  await writeFile(validate, validateSource)
}
