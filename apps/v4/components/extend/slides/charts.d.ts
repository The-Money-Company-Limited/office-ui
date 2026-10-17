export type ChartType =
  | "column"
  | "bar"
  | "line"
  | "area"
  | "pie"
  | "doughnut"
  | "scatter"
export const CHART_TYPES: ChartType[]
export const CHART_PALETTE: Record<string, string>
export const CHART_OPTIONS: Record<string, string[]>
export function chartAttribute(name: string, value: string): boolean
export type Chart = {
  type: ChartType
  rows: string[][]
  errors: string[]
  categories: string[]
  series: { name: string; values: number[] }[]
  title: string
  legend: string
  labels: boolean
  gridlines: boolean
  stacked: boolean
  xTitle: string
  yTitle: string
  format: string
  colors: string[]
}
export function readChart(figure: Element): Chart
export function writeChartRows(figure: Element, rows: string[][]): void
export function renderCharts(root: ParentNode, interactive?: boolean): void
