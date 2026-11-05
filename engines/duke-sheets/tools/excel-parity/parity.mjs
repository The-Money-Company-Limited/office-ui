// Recalculates real workbooks with the shipped engine and compares every formula with the value Excel saved.
// Usage: node parity.mjs <wasm-dir> <xlsx-dir> <out.json> [--max <mismatches>]
import { readFile, readdir, writeFile } from "node:fs/promises"
import path from "node:path"

const [wasmDir, dir, out] = process.argv.slice(2)
const maxAt = process.argv.indexOf("--max")
const max = maxAt > 0 ? Number(process.argv[maxAt + 1]) : Infinity
const wasm = await import(path.resolve(wasmDir, "duke_sheets_wasm.js"))
wasm.initSync({ module: await readFile(path.resolve(wasmDir, "duke_sheets_wasm_bg.wasm")) })

async function* walk(d) {
  for (const e of await readdir(d, { withFileTypes: true })) {
    const p = path.join(d, e.name)
    if (e.isDirectory()) yield* walk(p)
    else if (/\.xlsx$/i.test(e.name) && !e.name.startsWith("~$")) yield p
  }
}
const typed = (v) => {
  const j = v.toJs()
  if (j === null || j === undefined) return null
  if (typeof j === "object") return j.error ?? j.value ?? JSON.stringify(j)
  return j
}
const col = (c) => { let s = ""; for (c++; c; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + ((c - 1) % 26)) + s; return s }
function same(a, b) {
  if ((a === null || a === "") && (b === null || b === "")) return true
  if (typeof a === "number" && typeof b === "number") return Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(b)) || Math.abs(a - b) < 1e-12
  if (typeof a === "boolean" || typeof b === "boolean") return String(a).toUpperCase() === String(b).toUpperCase()
  // Excel saves a formula's text result without trailing spaces (CONCATENATE(A2," ") is stored as "AA"),
  // so the engine's value matches when it differs only there.
  return String(a) === String(b) || (typeof a === "string" && a.trimEnd() === b)
}
// Functions whose result legitimately differs from the saved one: the clock, randomness, the file's own location.
const volatile = /\b(NOW|TODAY|RAND|RANDBETWEEN|RANDARRAY|CELL|INFO|OFFSET\(.*TODAY)\s*\(/i

// The regional settings each workbook was computed with. Excel takes them from the computer and
// the file doesn't record them, so these come from the results Excel saved: day-first date text,
// Spanish month names, Swedish TEXT codes. Every other workbook is en-US.
const locales = { 49196: "en-GB", 33157: "en-GB", 31202: "es-ES", 45896: "sv-SE" }

const report = { files: 0, failedToOpen: [], formulas: 0, compared: 0, mismatches: [] }
for await (const file of walk(dir)) {
  report.files++
  let wb
  try {
    wasm.setLocale(locales[path.basename(path.dirname(file))] ?? "en-US")
    wb = wasm.Workbook.fromBytes(new Uint8Array(await readFile(file)))
  } catch (error) {
    report.failedToOpen.push({ file: path.relative(dir, file), error: String(error).slice(0, 200) })
    continue
  }
  const cells = []
  for (let s = 0; s < wb.sheetCount; s++) {
    const ws = wb.getSheet(s)
    for (let start = 0, rows; (rows = ws.getRowsBatch(start, 2000, { includeFormulas: true }) ?? []).length; start = rows.at(-1).index + 1)
    for (const row of rows) {
      for (const cell of row.cells) {
        if (!cell.formula) continue
        report.formulas++
        const saved = typed(ws.getCalculatedValueAt(row.index, cell.col))
        if (saved === null) continue // no saved value to compare with
        cells.push({ s, sheet: ws.name, r: row.index, c: cell.col, formula: cell.formula, saved })
      }
    }
  }
  try {
    wb.calculate()
  } catch (error) {
    report.failedToOpen.push({ file: path.relative(dir, file), error: "calculate: " + String(error).slice(0, 200) })
    continue
  }
  for (const x of cells) {
    if (volatile.test(x.formula)) continue
    report.compared++
    const got = typed(wb.getSheet(x.s).getCalculatedValueAt(x.r, x.c))
    if (!same(got, x.saved))
      report.mismatches.push({ file: path.relative(dir, file), cell: `${x.sheet}!${col(x.c)}${x.r + 1}`, formula: x.formula.slice(0, 200), excel: x.saved, engine: got })
  }
  wb.free()
}
await writeFile(out, JSON.stringify(report, null, 1))
console.log(`${report.files} files (${report.failedToOpen.length} failed), ${report.formulas} formulas, ${report.compared} compared, ${report.mismatches.length} mismatches`)
if (report.mismatches.length > max) {
  console.log(`more mismatches than the baseline (${max})`)
  process.exitCode = 1
}
// No workbook fails now, so one that does is a regression even if it lowers the mismatch count.
if (report.failedToOpen.length) {
  console.log(`failed: ${report.failedToOpen.map((f) => f.file).join(", ")}`)
  process.exitCode = 1
}
