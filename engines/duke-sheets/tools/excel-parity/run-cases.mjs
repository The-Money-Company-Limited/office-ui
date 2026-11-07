// Runs cases.mjs on an engine build. Usage: node run-cases.mjs <wasm-dir> <out.json>
import { readFile, writeFile } from "node:fs/promises"
import path from "node:path"
import { cases, known } from "./cases.mjs"

const [wasmDir, out] = process.argv.slice(2)
const wasm = await import(path.resolve(wasmDir, "duke_sheets_wasm.js"))
wasm.initSync({ module: await readFile(path.resolve(wasmDir, "duke_sheets_wasm_bg.wasm")) })
const value = (v) => {
  const j = v.toJs()
  if (j === null || j === undefined) return null
  return typeof j === "object" ? (j.error ?? j.value ?? JSON.stringify(j)) : j
}
const same = (got, want) =>
  typeof want === "number" && typeof got === "number"
    ? Math.abs(got - want) <= 1e-9 * Math.max(1, Math.abs(want))
    : want === "" ? got === "" || got === null : got === want
const results = []
for (const [name, cells, f, want, at = "Z1", fmt, locale = "en-US"] of cases) {
  const r = { name, f, want }
  try {
    wasm.setLocale(locale)
    const wb = new wasm.Workbook()
    const sheet = (n) => wb.sheetIndex(n) ?? wb.addSheet(n)
    for (const [key, v] of Object.entries(cells)) {
      if (key.startsWith("name:")) { wb.defineName(key.slice(5), v); continue }
      let ws = wb.getSheet(0), addr = key
      if (key.startsWith("sheet:")) { const [s, a] = key.slice(6).split("!"); ws = wb.getSheet(sheet(s)); addr = a }
      if (Array.isArray(v)) ws.setCell(addr, v[0])
      else if (typeof v === "string" && v.startsWith("=")) ws.setFormula(addr, v)
      else ws.setCell(addr, v)
      if (Array.isArray(v)) r.textCellStoredAs = typeof value(ws.getCell(addr))
    }
    const ws = wb.getSheet(0)
    if (f) ws.setFormula(at, f)
    if (fmt !== undefined) {
      const numberFormat = typeof fmt === "number" ? { formatType: "builtin", id: fmt } : { formatType: "custom", formatString: fmt }
      ws.setCellStyle(at, { numberFormat })
    }
    wb.calculate()
    r.got = fmt !== undefined ? ws.getFormattedValue(at) : value(ws.getCalculatedValue(at))
    wb.free()
  } catch (error) {
    r.got = "THREW: " + String(error).slice(0, 200)
  }
  r.ok = same(r.got, want)
  results.push(r)
}
await writeFile(out, JSON.stringify(results, null, 1))
const bad = results.filter((r) => !r.ok)
console.log(`${results.length} cases, ${bad.length} wrong (${bad.filter((r) => known.has(r.name)).length} known)`)
for (const r of bad) console.log(`- ${r.name}: ${r.f ?? ""} want ${JSON.stringify(r.want)} got ${JSON.stringify(r.got)}${r.textCellStoredAs ? " (text cell stored as " + r.textCellStoredAs + ")" : ""}`)
process.exitCode = bad.some((r) => !known.has(r.name)) ? 1 : 0
