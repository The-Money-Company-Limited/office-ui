import * as React from "react"

import { Input } from "@/components/ui/input"
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select"
import { Switch } from "@/components/ui/switch"

import {
  CHART_PALETTE,
  CHART_TYPES,
  readChart,
  renderCharts,
  writeChartRows,
} from "./charts.js"

const ICONS = [
  "M5 20V10h3v10m4 0V4h3v16m4 0v-8h3v8",
  "M4 5h16v3H4m0 4h10v3H4m0 4h19v3H4",
  "m3 19 6-8 5 3 7-11",
  "M3 20v-3l6-6 5 3 7-11v17Z",
  "M12 3a9 9 0 1 0 9 9h-9Z M15 2v7h7a9 9 0 0 0-7-7",
  "M21 12a9 9 0 1 1-9-9v5a4 4 0 1 0 4 4Z",
  "M5 17h2m4-9h2m4 5h2m-4-9h2",
]

export function ChartPanel({
  figure,
  docked,
  tab,
  onTab,
  onCommit,
  onClose,
  onUndo,
  onRedo,
  style,
}: {
  docked: boolean
  figure: Element
  tab: "Type" | "Data" | "Style"
  onTab: (tab: "Type" | "Data" | "Style") => void
  onCommit: () => void
  onClose: () => void
  onUndo: () => void
  onRedo: () => void
  style: React.CSSProperties
}) {
  const [errors, setErrors] = React.useState<string[]>([])
  const chart = readChart(figure)
  const round = ["pie", "doughnut"].includes(chart.type)
  function update(change: () => void, commit = true) {
    change()
    const next = readChart(figure)
    setErrors(next.errors)
    renderCharts(figure.parentElement!, true)
    if (commit && !next.errors.length) onCommit()
  }
  function attribute(name: string, value: string, commit = true) {
    update(() => figure.setAttribute(`data-${name}`, value), commit)
  }
  function rows(next: string[][], commit = true) {
    update(() => writeChartRows(figure, next), commit)
  }
  function finish() {
    if (!readChart(figure).errors.length) onCommit()
  }
  return (
    <div
      className="slides-chart-panel"
      data-docked={docked || undefined}
      role="dialog"
      aria-label="Chart controls"
      style={style}
      onPointerDown={(event) => event.stopPropagation()}
      onDoubleClick={(event) => event.stopPropagation()}
      onKeyDown={(event) => {
        event.stopPropagation()
        if (
          (event.metaKey || event.ctrlKey) &&
          event.key.toLowerCase() === "z"
        ) {
          event.preventDefault()
          finish()
          if (event.shiftKey) onRedo()
          else onUndo()
        }
        if (event.key === "Escape" && !event.defaultPrevented) {
          event.preventDefault()
          onClose()
        }
      }}
    >
      <header>
        <strong>Chart</strong>
        <button
          type="button"
          aria-label="Close chart controls"
          onClick={onClose}
        >
          ×
        </button>
      </header>
      <nav aria-label="Chart controls section">
        {(["Type", "Data", "Style"] as const).map((name) => (
          <button
            type="button"
            key={name}
            aria-pressed={tab === name}
            onClick={() => onTab(name)}
          >
            {name}
          </button>
        ))}
      </nav>
      {tab === "Type" && (
        <div className="slides-chart-types">
          {CHART_TYPES.map((type, i) => (
            <button
              type="button"
              key={type}
              aria-label={`${type} chart`}
              aria-pressed={chart.type === type}
              onClick={() => {
                const previous = figure.getAttribute("data-chart")!
                figure.setAttribute("data-chart", type)
                const problems = readChart(figure).errors
                figure.setAttribute("data-chart", previous)
                if (problems.length) setErrors(problems)
                else attribute("chart", type)
              }}
            >
              <svg viewBox="0 0 24 24" aria-hidden="true">
                <path
                  d={ICONS[i]}
                  fill="none"
                  stroke="currentColor"
                  strokeWidth="1.5"
                />
              </svg>
              <span>{type}</span>
            </button>
          ))}
        </div>
      )}
      {tab === "Data" && (
        <>
          <div className="slides-chart-grid">
            <table aria-label="Chart data">
              <tbody>
                {chart.rows.map((row, r) => (
                  <tr key={r}>
                    {row.map((value, c) => (
                      <td
                        key={c}
                        className={r === 0 ? "slides-chart-series" : undefined}
                      >
                        <input
                          autoFocus={r === 0 && c === 0}
                          aria-label={`Row ${r + 1} column ${c + 1}`}
                          value={value}
                          onChange={(event) => {
                            const next = readChart(figure).rows
                            next[r][c] = event.target.value
                            rows(next, false)
                          }}
                          onBlur={finish}
                          onKeyDown={(event) => {
                            if (event.key === "Enter") {
                              event.currentTarget.blur()
                            }
                          }}
                          onPaste={(event) => {
                            const text =
                              event.clipboardData.getData("text/plain")
                            if (!/[\t\n]/.test(text)) return
                            event.preventDefault()
                            const pasted = text
                              .replace(/\r/g, "")
                              .replace(/\n$/, "")
                              .split("\n")
                              .map((row) => row.split("\t"))
                            const next = readChart(figure).rows
                            const width = Math.max(
                              next[0].length,
                              c + Math.max(...pasted.map((row) => row.length))
                            )
                            while (next.length < r + pasted.length)
                              next.push(Array(width).fill("0"))
                            for (const row of next)
                              while (row.length < width) row.push("0")
                            pasted.forEach((row, y) =>
                              row.forEach((value, x) => {
                                next[r + y][c + x] = value
                              })
                            )
                            rows(next)
                          }}
                        />
                        {r === 0 && c > 0 && chart.series.length > 1 && (
                          <button
                            type="button"
                            className="slides-chart-remove-series"
                            aria-label={`Remove series ${c}`}
                            onClick={() =>
                              rows(
                                readChart(figure).rows.map((row) =>
                                  row.filter((_, i) => i !== c)
                                )
                              )
                            }
                          >
                            ×
                          </button>
                        )}
                      </td>
                    ))}
                    {r > 0 && (
                      <td>
                        <button
                          type="button"
                          aria-label={`Remove row ${r + 1}`}
                          disabled={chart.rows.length <= 2}
                          onClick={() =>
                            rows(
                              readChart(figure).rows.filter((_, i) => i !== r)
                            )
                          }
                        >
                          −
                        </button>
                      </td>
                    )}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div className="slides-chart-actions">
            <button
              type="button"
              onClick={() =>
                rows([
                  ...readChart(figure).rows,
                  [String(chart.rows.length), ...chart.series.map(() => "0")],
                ])
              }
            >
              Add row
            </button>
            <button
              type="button"
              disabled={round}
              onClick={() =>
                rows(
                  readChart(figure).rows.map((row, i) => [
                    ...row,
                    i ? "0" : `Series ${chart.series.length + 1}`,
                  ])
                )
              }
            >
              Add series
            </button>
          </div>
        </>
      )}
      {tab === "Style" && (
        <div className="slides-chart-style">
          {[
            ["title", "Title", chart.title],
            ...(!round
              ? [
                  ["x-title", "X axis title", chart.xTitle],
                  ["y-title", "Y axis title", chart.yTitle],
                ]
              : []),
          ].map(([name, label, value]) => (
            <label key={name}>
              {label}
              <Input
                value={value}
                maxLength={200}
                onChange={(event) => attribute(name, event.target.value, false)}
                onBlur={finish}
              />
            </label>
          ))}
          <ChartSelect
            label="Legend"
            value={chart.legend}
            options={["none", "top", "bottom", "left", "right"]}
            onChange={(value) => attribute("legend", value)}
          />
          <ChartSelect
            label="Number format"
            value={chart.format}
            options={["number", "percent", "currency"]}
            onChange={(value) => attribute("format", value)}
          />
          {[
            ["labels", "Data labels", chart.labels],
            ...(!round ? [["gridlines", "Gridlines", chart.gridlines]] : []),
            ...(["column", "bar", "area"].includes(chart.type)
              ? [["stacked", "Stacked", chart.stacked]]
              : []),
          ].map(([name, label, checked]) => (
            <label className="slides-chart-check" key={String(name)}>
              <Switch
                checked={Boolean(checked)}
                onCheckedChange={(checked) =>
                  attribute(String(name), String(checked))
                }
              />
              {label}
            </label>
          ))}
          {(round
            ? chart.categories
            : chart.series.map((series) => series.name)
          ).map((name, i) => (
            <ChartSelect
              key={i}
              label={`${name} color`}
              value={
                Object.keys(CHART_PALETTE).find(
                  (key) =>
                    CHART_PALETTE[key] === chart.colors[i % chart.colors.length]
                )!
              }
              options={Object.keys(CHART_PALETTE)}
              onChange={(value) => {
                const colors = Array.from(
                  { length: Math.max(i + 1, chart.colors.length) },
                  (_, index) =>
                    Object.keys(CHART_PALETTE).find(
                      (key) =>
                        CHART_PALETTE[key] ===
                        chart.colors[index % chart.colors.length]
                    )!
                )
                colors[i] = value
                attribute("colors", colors.join(","))
              }}
            />
          ))}
        </div>
      )}
      {errors.length > 0 && (
        <p role="alert" className="slides-chart-help">
          {errors.join(" ")}
        </p>
      )}
    </div>
  )
}

function ChartSelect({
  label,
  value,
  options,
  onChange,
}: {
  label: string
  value: string
  options: string[]
  onChange: (value: string) => void
}) {
  const [open, setOpen] = React.useState(false)
  const caption = (value: string) => value[0].toUpperCase() + value.slice(1)
  return (
    <div
      className="slides-chart-field"
      onKeyDown={(event) => {
        if (event.key === "Escape" && open) {
          event.stopPropagation()
          event.preventDefault()
          setOpen(false)
        }
      }}
    >
      <span>{label}</span>
      <Select
        open={open}
        onOpenChange={setOpen}
        value={value}
        onValueChange={(value) => {
          if (value) onChange(value)
        }}
        modal={false}
      >
        <SelectTrigger size="sm" aria-label={label}>
          <SelectValue>{caption(value)}</SelectValue>
        </SelectTrigger>
        <SelectContent className="z-40" alignItemWithTrigger={false}>
          {options.map((option) => (
            <SelectItem key={option} value={option}>
              {caption(option)}
            </SelectItem>
          ))}
        </SelectContent>
      </Select>
    </div>
  )
}
