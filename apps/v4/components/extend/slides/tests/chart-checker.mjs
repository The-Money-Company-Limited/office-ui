import assert from "node:assert/strict"
import { spawnSync } from "node:child_process"
import { cp, mkdir, readFile, writeFile } from "node:fs/promises"
import path from "node:path"
import { fileURLToPath } from "node:url"

const directory = fileURLToPath(
  new URL("./results/chart-checker/", import.meta.url)
)
const slides = fileURLToPath(new URL("../", import.meta.url))
await mkdir(path.join(directory, "deck/slides"), { recursive: true })
await writeFile(
  path.join(directory, "deck/deck.json"),
  JSON.stringify({ title: "Chart validation", order: ["invalid"] })
)
await writeFile(
  path.join(directory, "deck/slides/invalid.html"),
  `<section id="invalid">
<figure id="invalid-ragged" data-chart="column" style="width: 600px; height: 300px;"><table><tr><th>X</th><th>A</th><th>B</th></tr><tr><td>One</td><td>2</td></tr></table></figure>
<figure id="invalid-numeric" data-chart="line" style="width: 600px; height: 300px;"><table><tr><th>X</th><th>A</th></tr><tr><td>One</td><td>bad</td></tr></table></figure>
<figure id="invalid-empty" data-chart="area" style="width: 600px; height: 300px;"><table><tr><th>X</th><th></th></tr></table></figure>
</section>`
)
function run(cli) {
  const result = spawnSync(
    process.execPath,
    [cli, path.join(directory, "deck"), "--json"],
    { encoding: "utf8" }
  )
  assert.equal(result.error, undefined)
  const report = JSON.parse(result.stdout)
  assert.equal(report.error, undefined)
  return report
}
function verify(report) {
  for (const name of ["ragged", "numeric", "empty"]) {
    const issues = report.slides[0].issues.filter(
      (issue) =>
        issue.element === `invalid-${name}` && issue.code === "chart-data"
    )
    assert.ok(issues.length, name)
    assert.ok(
      issues.every(
        (issue) => issue.severity === "error" && issue.hint.length > 30
      )
    )
  }
}
const report = run(path.join(slides, "slides-check.mjs"))
verify(report)
await writeFile(
  path.join(directory, "report.json"),
  JSON.stringify(report, null, 2)
)
for (const file of ["validate.js", "charts.js", "slides.css"])
  await cp(path.join(slides, file), path.join(directory, file))
const source = await readFile(path.join(slides, "slides-check.mjs"), "utf8")
const disabled = source.replace(
  /window\.slidesCharts\s*\.readChart\(\s*figure\s*\)\s*\.errors/,
  "[]"
)
assert.notEqual(disabled, source, "chart-check mutation must apply")
await writeFile(path.join(directory, "slides-check.mjs"), disabled)
assert.throws(
  () => verify(run(path.join(directory, "slides-check.mjs"))),
  assert.AssertionError
)
console.log(
  "PASS chart checker reports ragged, non-numeric and empty series with fixes; assertions fail with chart checks removed"
)
