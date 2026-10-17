import assert from "node:assert/strict"
import { mkdir, readFile, writeFile } from "node:fs/promises"
import path from "node:path"
import { pathToFileURL } from "node:url"

const { chromium } = await import(
  pathToFileURL(path.resolve(".cod1460/node_modules/playwright/index.mjs"))
)
const source = await readFile(new URL("../charts.js", import.meta.url), "utf8")
const validator = await readFile(
  new URL("../validate.js", import.meta.url),
  "utf8"
)
const browser = await chromium.launch({
  headless: true,
  args: ["--no-sandbox"],
})
const page = await browser.newPage()
const results = []
const tests = [
  {
    name: "parse series and options",
    mutation: ["Number(row[index + 1])", "0"],
    check: async (chart) => {
      assert.deepEqual(chart.series, [
        { name: "Sales", values: [18, 32] },
        { name: "Costs", values: [4, -2] },
      ])
      assert.equal(chart.legend, "right")
      assert.equal(chart.colors[0], "a4e1d1")
    },
  },
  {
    name: "ragged rows",
    bad: "<tr><td>B</td><td>3</td></tr>",
    match: /same width/,
    mutation: ["row.length !== width", "false"],
  },
  {
    name: "non numeric values",
    bad: "<tr><td>B</td><td>nope</td><td>4</td></tr>",
    match: /finite number/,
    mutation: ["!numeric(value)", "false"],
  },
  {
    name: "blank values",
    bad: "<tr><td>B</td><td></td><td>4</td></tr>",
    match: /finite number/,
    mutation: ['value !== ""', "true"],
  },
  {
    name: "empty series",
    header: "<th></th>",
    match: /non-empty name/,
    mutation: ["some((name) => !name)", "some(() => false)"],
  },
  {
    name: "header only",
    noRows: true,
    match: /at least one/,
    mutation: ["rows.length < 2", "false"],
  },
  {
    name: "scatter needs numeric X",
    type: "scatter",
    match: /numeric X/,
    mutation: ['type === "scatter" &&', "false &&"],
  },
  {
    name: "pie single series",
    type: "pie",
    match: /one series/,
    mutation: ['["pie", "doughnut"].includes(type) &&', "false &&"],
  },
]
async function parse(test, text) {
  return page.evaluate(
    async ({ text, test }) => {
      const charts = await import(`data:text/javascript;base64,${btoa(text)}`)
      const figure = document.createElement("figure")
      figure.setAttribute("data-chart", test.type || "column")
      figure.setAttribute("data-legend", "right")
      figure.setAttribute("data-colors", "teal,blue")
      figure.innerHTML = `<table><tr><th>Category</th><th>Sales</th>${test.header || "<th>Costs</th>"}</tr>${test.noRows ? "" : `<tr><td>A</td><td>18</td><td>4</td></tr>${test.bad || "<tr><td>B</td><td>32</td><td>-2</td></tr>"}`}</table>`
      return charts.readChart(figure)
    },
    { text, test: { ...test, check: undefined, match: undefined } }
  )
}
try {
  for (const test of tests) {
    const verify = async (source) => {
      const chart = await parse(test, source)
      if (test.check) await test.check(chart)
      else assert.match(chart.errors.join(" "), test.match)
    }
    await verify(source)
    const disabled = source.replace(...test.mutation)
    assert.notEqual(disabled, source)
    await assert.rejects(() => verify(disabled), { name: "AssertionError" })
    results.push(test.name)
    console.log(`PASS ${test.name}; fails with its behavior removed`)
  }
  async function scales(text) {
    const actual = await page.evaluate(async (text) => {
      const { niceScale, chartValueScale } = await import(
        `data:text/javascript;base64,${btoa(text)}`
      )
      return {
        ticks: [
          [0, 90],
          [-90, 0],
          [-30, 90],
          [0, 0],
          [0, 0.009],
        ].map(([min, max]) => niceScale(min, max).ticks),
        stacked: chartValueScale({
          categories: ["A", "B"],
          series: [{ values: [40, -30] }, { values: [50, -40] }],
          stacked: true,
        }),
      }
    }, text)
    assert.deepEqual(actual.ticks, [
      [0, 25, 50, 75, 100],
      [-100, -75, -50, -25, 0],
      [-50, 0, 50, 100],
      [0, 0.25, 0.5, 0.75, 1],
      [0, 0.0025, 0.005, 0.0075, 0.01],
    ])
    assert.deepEqual(actual.stacked.ticks, [-100, -50, 0, 50, 100])
  }
  await scales(source)
  await assert.rejects(
    () =>
      scales(source.replace("value >= roughStep / magnitude", "value >= 10")),
    { name: "AssertionError" }
  )
  await assert.rejects(
    () =>
      scales(
        source
          .replace("chart.stacked ? negative : values", "values")
          .replace("chart.stacked ? positive : values", "values")
      ),
    { name: "AssertionError" }
  )
  console.log(
    "PASS nice axis scales, decimals, zero, negatives and stacked bounds; fails without each behavior"
  )
  async function renderer(text) {
    const output = await page.evaluate(async (text) => {
      const charts = await import(`data:text/javascript;base64,${btoa(text)}`)
      return charts.CHART_TYPES.map((type) => {
        const root = document.createElement("div")
        root.innerHTML = `<figure data-chart="${type}" data-labels="true" data-title="A &lt; B" data-colors="rose"><table><tr><th>X</th><th>Sales</th></tr><tr><td>1</td><td>18</td></tr><tr><td>2</td><td>32</td></tr></table></figure>`
        charts.renderCharts(root)
        const svg = root.querySelector("svg")
        return {
          type,
          svg: svg?.outerHTML,
          text: svg?.textContent,
          rows: root.querySelectorAll("tr").length,
        }
      })
    }, text)
    for (const { type, svg, text, rows } of output) {
      assert.ok(svg?.includes('viewBox="0 0 1000 600"'), type)
      assert.ok(svg.includes("#fda4af"), type)
      assert.ok(
        text.includes("18") && text.includes("32") && text.includes("A < B"),
        type
      )
      assert.ok(!/NaN|Infinity|undefined/.test(svg), type)
      assert.equal(rows, 3)
      assert.ok(
        svg.includes(
          ["pie", "doughnut", "area", "line"].includes(type)
            ? "<path"
            : type === "scatter"
              ? "<circle"
              : "<rect"
        ),
        type
      )
    }
  }
  await renderer(source)
  await assert.rejects(
    () =>
      renderer(
        source.replace(
          "export function renderCharts(root, interactive = false) {",
          "export function renderCharts(root, interactive = false) { return;"
        )
      ),
    { name: "AssertionError" }
  )
  console.log("PASS all seven SVG renderers; fails without rendering")
  async function narrow(text) {
    const widths = await page.evaluate(async (text) => {
      const charts = await import(`data:text/javascript;base64,${btoa(text)}`)
      const root = document.createElement("div")
      root.innerHTML =
        '<figure data-chart="column" style="width: 40px; height: 700px"><table><tr><th>X</th><th>Y</th></tr><tr><td>A</td><td>2</td></tr></table></figure>'
      document.body.append(root)
      charts.renderCharts(root)
      const widths = [...root.querySelectorAll("svg rect")].map((rect) =>
        Number(rect.getAttribute("width"))
      )
      root.remove()
      return widths
    }, text)
    assert.ok(
      widths.length && widths.every((width) => width > 0),
      "narrow resize keeps SVG dimensions positive"
    )
  }
  await narrow(source)
  await assert.rejects(
    () =>
      narrow(
        source.replace(
          "Math.max(600, (600 * figure.clientWidth) / figure.clientHeight)",
          "(600 * figure.clientWidth) / figure.clientHeight"
        )
      ),
    { name: "AssertionError" }
  )
  console.log("PASS narrow chart resize; fails without minimum view width")
  const url = `data:text/javascript;base64,${Buffer.from(source).toString("base64")}`
  async function subset(text) {
    const html = await page.evaluate(
      async (text) => {
        const { sanitizeSlide } = await import(
          `data:text/javascript;base64,${btoa(text)}`
        )
        return sanitizeSlide(
          '<section id="test"><figure data-chart="line" data-legend="right" data-colors="teal,blue" data-labels="true" onclick="alert(1)"><table><tr><th>X</th><th>Y</th></tr><tr><td>1</td><td>2</td></tr></table><svg><rect /></svg></figure></section>',
          "test",
          document,
          () => {}
        )
      },
      text.replace("./charts.js", url)
    )
    assert.match(html, /data-chart="line"/)
    assert.match(html, /data-legend="right"/)
    assert.ok(!html.includes("onclick") && !html.includes("<svg"))
  }
  await subset(validator)
  await assert.rejects(
    () =>
      subset(
        validator.replace(
          'tag === "figure" && chartAttribute(name, value)',
          "false"
        )
      ),
    { name: "AssertionError" }
  )
  console.log(
    "PASS sanitizer preserves chart source only; fails without chart attributes"
  )
  await mkdir(new URL("./results/", import.meta.url), { recursive: true })
  await writeFile(
    new URL("./results/chart-units.json", import.meta.url),
    JSON.stringify(
      {
        pass: true,
        negativeControls: true,
        cases: [...results, "seven renderers", "narrow resize", "sanitizer"],
      },
      null,
      2
    )
  )
} finally {
  await browser.close()
}
