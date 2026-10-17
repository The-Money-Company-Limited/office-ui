import assert from "node:assert/strict"
import { spawnSync } from "node:child_process"
import { cp, mkdir, readFile, rm, symlink, writeFile } from "node:fs/promises"
import path from "node:path"
import { fileURLToPath } from "node:url"

const slides = fileURLToPath(new URL("../", import.meta.url))
const output = path.join(slides, "tests/results/checker")
const cli = path.join(slides, "slides-check.mjs")
const fixture = path.join(slides, "tests/fixtures/checker")
const sample = fileURLToPath(
  new URL(
    "../../../../experiments/artifact-template/samples/deck",
    import.meta.url
  )
)
await mkdir(output, { recursive: true })

function run(command, deck, ...args) {
  const result = spawnSync(process.execPath, [command, deck, ...args], {
    encoding: "utf8",
    timeout: 120_000,
  })
  assert.equal(result.error, undefined)
  assert.equal(result.signal, null)
  return {
    ...result,
    report: args.includes("--json") ? JSON.parse(result.stdout) : undefined,
  }
}
const sampleResult = run(
  cli,
  sample,
  "--json",
  "--png",
  path.join(output, "sample-png")
)
await writeFile(path.join(output, "sample.json"), sampleResult.stdout)
assert.equal(sampleResult.status, 0, sampleResult.stdout + sampleResult.stderr)
assert.equal(sampleResult.report.errors, 0)
assert.equal(sampleResult.report.slides.length, 4)
for (const { id } of sampleResult.report.slides) {
  const png = await readFile(path.join(output, "sample-png", `${id}.png`))
  assert.equal(png.readUInt32BE(16), 1920)
  assert.equal(png.readUInt32BE(20), 1080)
}
console.log(
  `Sample: ${sampleResult.report.errors} errors, ${sampleResult.report.warnings} warnings; four 1920×1080 PNGs`
)

const result = run(
  cli,
  fixture,
  "--json",
  "--png",
  path.join(output, "fixture-png")
)
await writeFile(path.join(output, "fixture.json"), result.stdout)
assert.equal(result.status, 1)
const expected = [
  ["markup", "markup", "markup-video"],
  ["overflow", "text-overflow", "overflow-text"],
  ["clipped", "text-clipped", "clipped-text"],
  ["wrap", "mid-word-wrap", "wrap-text"],
  ["outside", "outside-slide", "outside-shape"],
  ["margin", "margin", "margin-text"],
  ["overlap", "text-overlap", "overlap-first"],
  ["contrast", "contrast", "contrast-small"],
  ["image", "image-load", "image-broken"],
]
function assertFinding(report, [id, code, element]) {
  const slide = report.slides.find((slide) => slide.id === id)
  const issue = slide?.issues.find(
    (issue) => issue.code === code && issue.element === element
  )
  assert.ok(issue, `${id}: expected ${code} on ${element}`)
  assert.ok(issue.hint.length > 20)
  assert.equal(
    issue.severity,
    ["outside-slide", "margin"].includes(code) ? "warning" : "error"
  )
  if (id !== "outside") assert.ok(issue.text.length > 0)
}
for (const entry of expected) {
  assertFinding(result.report, entry)
  console.log(`PASS fixture: ${entry[1]} on ${entry[2]}`)
}
for (const element of [
  "markup-attribute",
  "markup-style",
  "markup-invalid",
  "markup-outside",
])
  assertFinding(result.report, ["markup", "markup", element])
assertFinding(result.report, ["wrap", "mid-word-wrap", "wrap-inline"])
assertFinding(result.report, ["wrap", "mid-word-wrap", "wrap-number"])
assertFinding(result.report, ["wrap", "mid-word-wrap", "wrap-rtl"])
for (const element of [
  "contrast-large",
  "contrast-layer",
  "contrast-svg",
  "contrast-opacity",
])
  assertFinding(result.report, ["contrast", "contrast", element])
const contrast = result.report.slides.find(
  (slide) => slide.id === "contrast"
).issues
assert.equal(
  contrast.find((issue) => issue.element === "contrast-small").threshold,
  4.5
)
assert.equal(
  contrast.find((issue) => issue.element === "contrast-large").threshold,
  3
)
assert.deepEqual(
  result.report.slides.find((slide) => slide.id === "clean").issues,
  []
)
const marginIssues = result.report.slides.find(
  (slide) => slide.id === "margin"
).issues
assert.deepEqual(
  marginIssues.map(({ code, element }) => [code, element]),
  [["margin", "margin-text"]],
  "only the outermost content group warns; boundary content and decoration do not"
)
console.log(
  `PASS fixture total: ${result.report.errors} errors, ${result.report.warnings} warnings; clean controls have no findings`
)

for (const [name, deck] of [
  ["sample", sample],
  ["fixture", fixture],
]) {
  const human = run(cli, deck)
  await writeFile(path.join(output, `${name}.txt`), human.stdout)
  assert.match(human.stdout, /slides, \d+ errors, \d+ warnings/)
  if (name === "fixture") assert.match(human.stdout, /Fix:/)
  console.log(human.stdout)
}

const warningDeck = path.join(output, "warnings")
await mkdir(path.join(warningDeck, "slides"), { recursive: true })
await writeFile(
  path.join(warningDeck, "deck.json"),
  JSON.stringify({ title: "Warnings only", order: ["margin", "outside"] })
)
for (const id of ["margin", "outside"])
  await cp(
    path.join(fixture, "slides", `${id}.html`),
    path.join(warningDeck, "slides", `${id}.html`)
  )
const warnings = run(cli, warningDeck, "--json")
assert.equal(warnings.status, 0)
assert.equal(warnings.report.errors, 0)
assert.equal(warnings.report.warnings, 3)
console.log("PASS warnings alone exit 0")

const boundaryDeck = path.join(output, "margin-boundaries")
await mkdir(path.join(boundaryDeck, "slides"), { recursive: true })
const boundaryOrder = []
const boundaryOffenders = []
for (const edge of ["left", "top", "right", "bottom"]) {
  for (const [label, crossing] of [
    ["exact", 0],
    ["tolerated", 1],
    ["outside", 1.25],
  ]) {
    const id = `${edge}-${label}`
    boundaryOrder.push(id)
    if (crossing > 1) boundaryOffenders.push(id)
    const x =
      edge === "left"
        ? 128 - crossing
        : edge === "right"
          ? 1792 - 400 + crossing
          : 400
    const y =
      edge === "top"
        ? 128 - crossing
        : edge === "bottom"
          ? 952 - 80 + crossing
          : 400
    await writeFile(
      path.join(boundaryDeck, "slides", `${id}.html`),
      `<section id="${id}"><p id="${id}-text" style="position: absolute; left: ${x}px; top: ${y}px; width: 400px; height: 80px; padding: 8px; border: 2px solid black;">Boundary content</p></section>`
    )
  }
}
await writeFile(
  path.join(boundaryDeck, "deck.json"),
  JSON.stringify({ title: "Margin boundaries", order: boundaryOrder })
)
const boundaries = run(cli, boundaryDeck, "--json")
await writeFile(path.join(output, "margin-boundaries.json"), boundaries.stdout)
assert.equal(boundaries.status, 0)
assert.equal(boundaries.report.errors, 0)
assert.equal(boundaries.report.warnings, 4)
assert.deepEqual(
  boundaries.report.slides
    .filter((slide) => slide.issues.length)
    .map((slide) => slide.id),
  boundaryOffenders
)
console.log(
  "PASS all four margin edges: exact and 1px crossings pass; 1.25px crossings warn"
)

const invalidDeck = path.join(output, "invalid")
await mkdir(path.join(invalidDeck, "slides"), { recursive: true })
await writeFile(
  path.join(invalidDeck, "deck.json"),
  JSON.stringify({ title: "Invalid", order: ["../escape"] })
)
const invalid = run(cli, invalidDeck, "--json")
assert.equal(invalid.status, 1)
assert.match(invalid.report.error, /Invalid slide order/)
await writeFile(
  path.join(invalidDeck, "deck.json"),
  JSON.stringify({
    title: "Missing and escaped slides",
    order: ["missing", "escaped", "clean"],
  })
)
await rm(path.join(invalidDeck, "slides/escaped.html"), { force: true })
await symlink(
  path.join(fixture, "slides/clean.html"),
  path.join(invalidDeck, "slides/escaped.html")
)
await cp(
  path.join(fixture, "slides/clean.html"),
  path.join(invalidDeck, "slides/clean.html")
)
const missing = run(cli, invalidDeck, "--json")
assert.equal(missing.status, 1)
assert.equal(missing.report.errors, 2)
assert.match(missing.report.slides[1].issues[0].message, /escapes/)
assert.deepEqual(missing.report.slides[2].issues, [])
console.log("PASS invalid manifest, unreadable/escaped slides and continuation")

if (process.argv.includes("--negative-control")) {
  const original = await readFile(cli, "utf8")
  const mutants = [
    ["issues.push({", "void ({"],
    ["!svg &&", "false && !svg &&"],
    ["if (!clipX && !clipY) continue", "if (true) continue"],
    ["!segment.isWordLike", "true || !segment.isWordLike"],
    ["if (outside(bounds, canvas))", "if (false && outside(bounds, canvas))"],
    ["outside(bounds, {", "false && outside(bounds, {"],
    ["Math.min(a.right, b.right)", "false && Math.min(a.right, b.right)"],
    ["if (ratio < threshold)", "if (false && ratio < threshold)"],
    [
      "element instanceof HTMLImageElement",
      "false && element instanceof HTMLImageElement",
    ],
  ]
  const directory = path.join(output, "negative-control")
  await mkdir(directory, { recursive: true })
  for (const file of ["validate.js", "slides.css"])
    await cp(path.join(slides, file), path.join(directory, file))
  const mutant = path.join(directory, "slides-check.mjs")
  for (const [i, [from, to]] of mutants.entries()) {
    assert.ok(original.includes(from), `mutation anchor ${i}`)
    await writeFile(mutant, original.replace(from, to))
    const negative = run(mutant, fixture, "--json")
    assert.equal(
      negative.report.error,
      undefined,
      "mutant must run, not fail to launch"
    )
    assert.throws(
      () => assertFinding(negative.report, expected[i]),
      assert.AssertionError
    )
    console.log(
      `PASS negative control: ${expected[i][1]} assertion fails with that check removed`
    )
  }
  const absent = run(path.join(directory, "absent-checker.mjs"), sample)
  assert.notEqual(absent.status, 0, "sample command fails without the new CLI")
  await writeFile(
    mutant,
    original.replace(
      "ancestor && ancestor !== root;",
      "ancestor && root.contains(ancestor);"
    )
  )
  assert.notEqual(run(mutant, warningDeck, "--json").status, 0)
  console.log(
    "PASS negative control: warning-only assertion fails when canvas clipping is an error"
  )
  for (const [from, to, label] of [
    ["textElements.has(element)", "true", "decoration"],
    [
      "!marginOffenders.some((parent) => parent.contains(element))",
      "true",
      "descendant suppression",
    ],
  ]) {
    assert.ok(original.includes(from))
    await writeFile(mutant, original.replace(from, to))
    const negative = run(mutant, fixture, "--json")
    const issues = negative.report.slides.find(
      (slide) => slide.id === "margin"
    ).issues
    assert.notDeepEqual(
      issues.map(({ code, element }) => [code, element]),
      [["margin", "margin-text"]]
    )
    console.log(
      `PASS negative control: margin regression fails without ${label} filtering`
    )
  }
}
assert.equal(
  sampleResult.report.warnings,
  0,
  "sample must respect the 128px safe margins"
)
console.log("PASS sample has no warnings")
