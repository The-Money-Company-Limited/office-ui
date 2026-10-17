export const CHART_TYPES = [
  "column",
  "bar",
  "line",
  "area",
  "pie",
  "doughnut",
  "scatter",
]
export const CHART_PALETTE = {
  blue: "7dd3fc",
  teal: "a4e1d1",
  violet: "c4b5fd",
  rose: "fda4af",
  amber: "fcd34d",
  slate: "a9bed8",
}
export const CHART_OPTIONS = {
  "data-chart": CHART_TYPES,
  "data-legend": ["none", "top", "bottom", "left", "right"],
  "data-labels": ["true", "false"],
  "data-gridlines": ["true", "false"],
  "data-stacked": ["true", "false"],
  "data-format": ["number", "percent", "currency"],
}
export function chartAttribute(name, value) {
  if (CHART_OPTIONS[name]) return CHART_OPTIONS[name].includes(value)
  if (["data-title", "data-x-title", "data-y-title"].includes(name))
    return value.length <= 200
  return (
    name === "data-colors" &&
    value.split(",").every((name) => Object.hasOwn(CHART_PALETTE, name))
  )
}
export function readChart(figure) {
  const rows = [...(figure.querySelector(":scope > table")?.rows || [])].map(
    (row) => [...row.cells].map((cell) => cell.textContent.trim())
  )
  const type = figure.getAttribute("data-chart")
  const errors = []
  if (!CHART_TYPES.includes(type))
    errors.push("Choose a supported data-chart type.")
  if (rows.length < 2 || (rows[0]?.length || 0) < 2)
    errors.push(
      "Add a header with series names and at least one category/value row."
    )
  const width = rows[0]?.length || 0
  if (rows.some((row) => row.length !== width))
    errors.push(
      "Make every row the same width as the header; fill each missing value."
    )
  if (rows[0]?.slice(1).some((name) => !name))
    errors.push("Give every series a non-empty name in the header row.")
  const numeric = (value) =>
    value !== undefined && value !== "" && Number.isFinite(Number(value))
  if (
    rows
      .slice(1)
      .some((row) => !row[0] || row.slice(1).some((value) => !numeric(value)))
  )
    errors.push(
      "Give each category a name and each value a finite number (no blanks, commas or units)."
    )
  if (figure.querySelector("[colspan], [rowspan]"))
    errors.push("Use unmerged chart cells; remove colspan and rowspan.")
  if (
    ["pie", "doughnut"].includes(type) &&
    (width !== 2 ||
      rows.slice(1).some((row) => Number(row[1]) < 0) ||
      !rows.slice(1).some((row) => Number(row[1]) > 0))
  )
    errors.push(
      "Pie and doughnut need one series of non-negative values with a positive total."
    )
  if (type === "scatter" && rows.slice(1).some((row) => !numeric(row[0])))
    errors.push("For scatter, use numeric X values in the first column.")
  const get = (name, fallback = "") =>
    figure.getAttribute(`data-${name}`) ?? fallback
  return {
    type,
    rows,
    errors,
    categories: rows.slice(1).map((row) => row[0]),
    series: (rows[0]?.slice(1) || []).map((name, index) => ({
      name,
      values: rows.slice(1).map((row) => Number(row[index + 1])),
    })),
    title: get("title"),
    legend: get("legend", "bottom"),
    labels: get("labels") === "true",
    gridlines: get("gridlines", "true") === "true",
    stacked:
      ["column", "bar", "area"].includes(type) && get("stacked") === "true",
    xTitle: get("x-title"),
    yTitle: get("y-title"),
    format: get("format", "number"),
    colors: get("colors", Object.keys(CHART_PALETTE).join(","))
      .split(",")
      .map((name) => CHART_PALETTE[name] || CHART_PALETTE.blue),
  }
}
export function niceScale(min, max) {
  if (min === max) {
    min = Math.min(0, min)
    max = Math.max(1, max)
  }
  const roughStep = (max - min) / 4
  const magnitude = 10 ** Math.floor(Math.log10(roughStep))
  const step =
    [1, 2, 2.5, 5, 10].find((value) => value >= roughStep / magnitude) *
    magnitude
  const start = Math.floor(min / step),
    end = Math.ceil(max / step)
  const ticks = Array.from({ length: end - start + 1 }, (_, i) =>
    Number(((start + i) * step).toPrecision(12))
  )
  return { min: ticks[0], max: ticks.at(-1), step, ticks }
}
export function chartValueScale(chart) {
  const values = chart.series.flatMap((series) => series.values)
  const positive = chart.categories.map((_, i) =>
    chart.series.reduce((sum, series) => sum + Math.max(0, series.values[i]), 0)
  )
  const negative = chart.categories.map((_, i) =>
    chart.series.reduce((sum, series) => sum + Math.min(0, series.values[i]), 0)
  )
  return niceScale(
    Math.min(0, ...(chart.stacked ? negative : values)),
    Math.max(0, ...(chart.stacked ? positive : values))
  )
}
export function writeChartRows(figure, rows) {
  const table = figure.ownerDocument.createElement("table")
  for (const [index, values] of rows.entries()) {
    const row = table.insertRow()
    for (const value of values) {
      const cell = figure.ownerDocument.createElement(index ? "td" : "th")
      cell.textContent = value
      row.append(cell)
    }
  }
  figure.querySelector(":scope > table")?.remove()
  figure.prepend(table)
}
export function renderCharts(root, interactive = false) {
  for (const figure of root.querySelectorAll("figure[data-chart]")) {
    figure.querySelector(":scope > svg")?.remove()
    const chart = readChart(figure)
    const dom = figure.ownerDocument
    const svg = dom.createElementNS("http://www.w3.org/2000/svg", "svg")
    const viewWidth =
      figure.clientWidth && figure.clientHeight
        ? Math.max(600, (600 * figure.clientWidth) / figure.clientHeight)
        : 1000
    svg.setAttribute("viewBox", `0 0 ${viewWidth} 600`)
    svg.setAttribute("role", "img")
    svg.setAttribute("aria-label", chart.title || `${chart.type} chart`)
    if (interactive) {
      figure.tabIndex = 0
      figure.setAttribute(
        "aria-label",
        `${chart.title || chart.type} chart. Enter to edit data.`
      )
    }
    function add(tag, attrs, text) {
      const node = dom.createElementNS(svg.namespaceURI, tag)
      for (const [name, value] of Object.entries(attrs))
        node.setAttribute(name, String(value))
      if (text !== undefined) node.textContent = text
      svg.append(node)
      return node
    }
    const label = (x, y, text, anchor = "middle", size = 20) =>
      add(
        "text",
        {
          x,
          y,
          fill: "currentColor",
          "text-anchor": anchor,
          "font-size": size,
          "font-family": "Arial",
        },
        text
      )
    const paint = (i) => `#${chart.colors[i % chart.colors.length]}`
    const format = (value) =>
      new Intl.NumberFormat("en-US", {
        style:
          chart.format === "currency"
            ? "currency"
            : chart.format === "percent"
              ? "percent"
              : "decimal",
        currency: "USD",
        maximumFractionDigits: 2,
      }).format(value)
    if (chart.errors.length) {
      label(viewWidth / 2, 280, "Fix chart data", "middle", 28)
      label(viewWidth / 2, 320, "Open Data to edit the source table.")
      figure.append(svg)
      continue
    }
    if (chart.title) label(viewWidth / 2, 35, chart.title, "middle", 28)
    const round = ["pie", "doughnut"].includes(chart.type)
    const legend = round
      ? chart.categories
      : chart.series.map((series) => series.name)
    if (chart.legend !== "none")
      legend.forEach((name, i) => {
        const vertical = ["left", "right"].includes(chart.legend)
        const x = vertical
          ? chart.legend === "left"
            ? 12
            : viewWidth - 175
          : viewWidth / 2 - legend.length * 70 + i * 140
        const y = vertical ? 90 + i * 30 : chart.legend === "top" ? 70 : 582
        add("rect", { x, y: y - 14, width: 14, height: 14, fill: paint(i) })
        label(x + 21, y, name, "start", 17)
      })
    const left = chart.legend === "left" ? 225 : 100
    const right = chart.legend === "right" ? viewWidth - 210 : viewWidth - 45
    const top = chart.legend === "top" ? 110 : 65
    const bottom = 495
    const width = right - left,
      height = bottom - top
    if (round) {
      const values = chart.series[0].values
      const total = values.reduce((a, b) => a + b, 0)
      const cx = (left + right) / 2,
        cy = (top + bottom) / 2,
        r = Math.min(width, height) / 2 - 20
      let angle = -Math.PI / 2
      values.forEach((value, i) => {
        const next = angle + (value / total) * Math.PI * 2
        const point = (a, radius) =>
          `${cx + Math.cos(a) * radius},${cy + Math.sin(a) * radius}`
        // Two arcs also handle a single slice occupying the entire circle.
        const mid = (angle + next) / 2
        const inner = chart.type === "doughnut" ? r * 0.55 : 0
        const d = `M${point(angle, r)} A${r},${r} 0 0 1 ${point(mid, r)} A${r},${r} 0 0 1 ${point(next, r)} L${point(next, inner)} ${inner ? `A${inner},${inner} 0 0 0 ${point(mid, inner)} A${inner},${inner} 0 0 0 ${point(angle, inner)}` : ""} Z`
        add("path", { d, fill: paint(i) })
        if (chart.labels && value)
          label(
            cx + Math.cos(mid) * r * 0.8,
            cy + Math.sin(mid) * r * 0.8,
            format(value)
          )
        angle = next
      })
    } else {
      const axis = chartValueScale(chart)
      const { min, max } = axis
      const bar = chart.type === "bar"
      const valuePosition = (value) =>
        bar
          ? left + ((value - min) / (max - min)) * width
          : bottom - ((value - min) / (max - min)) * height
      const count = chart.categories.length
      const step = (bar ? height : width) / count
      const xs = chart.categories.map(Number)
      const xAxis =
        chart.type === "scatter"
          ? niceScale(Math.min(...xs), Math.max(...xs))
          : null
      const categoryPosition = (i) =>
        chart.type === "scatter"
          ? left + ((xs[i] - xAxis.min) / (xAxis.max - xAxis.min)) * width
          : (bar ? top : left) + step * (i + 0.5)
      for (const value of axis.ticks) {
        const pos = valuePosition(value)
        if (chart.gridlines)
          add("line", {
            x1: bar ? pos : left,
            x2: bar ? pos : right,
            y1: bar ? top : pos,
            y2: bar ? bottom : pos,
            stroke: "currentColor",
            opacity: 0.18,
          })
        label(
          bar ? pos : left - 12,
          bar ? bottom + 24 : pos + 6,
          format(value),
          bar ? "middle" : "end",
          17
        )
      }
      if (xAxis)
        xAxis.ticks.forEach((value) =>
          label(
            left + ((value - xAxis.min) / (xAxis.max - xAxis.min)) * width,
            bottom + 28,
            String(value),
            "middle",
            18
          )
        )
      else
        chart.categories.forEach((name, i) =>
          label(
            bar ? left - 12 : categoryPosition(i),
            bar ? categoryPosition(i) + 6 : bottom + 28,
            name,
            bar ? "end" : "middle",
            18
          )
        )
      const positive = Array(count).fill(0),
        negative = Array(count).fill(0)
      chart.series.forEach((series, s) => {
        const base = series.values.map((value, i) =>
          chart.stacked ? (value >= 0 ? positive[i] : negative[i]) : 0
        )
        const end = series.values.map((value, i) => base[i] + value)
        const points = end.map((value, i) => [
          categoryPosition(i),
          valuePosition(value),
        ])
        if (["line", "area"].includes(chart.type)) {
          const lower = base
            .map((value, i) => [categoryPosition(i), valuePosition(value)])
            .reverse()
          add("path", {
            d: `M${points.map((p) => p.join(",")).join(" L")}${chart.type === "area" ? ` L${lower.map((p) => p.join(",")).join(" L")} Z` : ""}`,
            fill: chart.type === "area" ? paint(s) : "none",
            "fill-opacity": 0.5,
            stroke: paint(s),
            "stroke-width": 3,
          })
        }
        series.values.forEach((value, i) => {
          const c = categoryPosition(i),
            v = valuePosition(end[i]),
            b = valuePosition(base[i])
          let labelCategory = c
          if (["bar", "column"].includes(chart.type)) {
            const thickness =
              (step * 0.7) / (chart.stacked ? 1 : chart.series.length)
            const offset = c - step * 0.35 + (chart.stacked ? 0 : s * thickness)
            labelCategory = offset + thickness / 2
            add("rect", {
              x: bar ? Math.min(v, b) : offset,
              y: bar ? offset : Math.min(v, b),
              width: bar ? Math.abs(v - b) : thickness,
              height: bar ? thickness : Math.abs(v - b),
              fill: paint(s),
            })
          } else add("circle", { cx: c, cy: v, r: 5, fill: paint(s) })
          if (chart.labels)
            label(
              bar ? v + 10 : labelCategory,
              bar ? labelCategory + 6 : v - 12,
              format(value),
              bar ? "start" : "middle",
              18
            )
          if (value >= 0) positive[i] = end[i]
          else negative[i] = end[i]
        })
      })
      if (chart.xTitle) label((left + right) / 2, 550, chart.xTitle)
      if (chart.yTitle) {
        const node = label(25, (top + bottom) / 2, chart.yTitle)
        node.setAttribute("transform", `rotate(-90 25 ${(top + bottom) / 2})`)
      }
    }
    figure.append(svg)
  }
}
