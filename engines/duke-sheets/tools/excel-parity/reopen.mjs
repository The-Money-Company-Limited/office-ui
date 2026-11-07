// Saves every workbook in a folder as .xlsx and opens it again: each formula must read the same.
// Usage: node reopen.mjs <wasm-dir> <xlsx-dir> <out.json>
import { readFile, readdir, writeFile } from "node:fs/promises"
import path from "node:path"

const [wasmDir, dir, out] = process.argv.slice(2)
const wasm = await import(path.resolve(wasmDir, "duke_sheets_wasm.js"))
wasm.initSync({ module: await readFile(path.resolve(wasmDir, "duke_sheets_wasm_bg.wasm")) })

async function* walk(d) {
  for (const e of await readdir(d, { withFileTypes: true })) {
    const p = path.join(d, e.name)
    if (e.isDirectory()) yield* walk(p)
    else if (/\.xlsx$/i.test(e.name) && !e.name.startsWith("~$")) yield p
  }
}
const formulas = (wb) => {
  const all = new Map()
  for (let s = 0; s < wb.sheetCount; s++)
    for (const { row, col, formula } of wb.getSheet(s).formulaCells) all.set(`${s}!${row},${col}`, formula)
  return all
}

const report = { files: 0, formulas: 0, failed: [], changed: [] }
for await (const file of walk(dir)) {
  report.files++
  const name = path.relative(dir, file)
  let before, after
  try {
    const wb = wasm.Workbook.fromBytes(new Uint8Array(await readFile(file)))
    before = formulas(wb)
    const saved = wb.saveXlsxBytes()
    wb.free()
    const reopened = wasm.Workbook.fromBytes(saved)
    after = formulas(reopened)
    reopened.free()
  } catch (error) {
    report.failed.push({ file: name, error: String(error).slice(0, 200) })
    continue
  }
  report.formulas += before.size
  for (const [cell, formula] of before)
    if (after.get(cell) !== formula) report.changed.push({ file: name, cell, before: formula, after: after.get(cell) ?? null })
  for (const cell of after.keys())
    if (!before.has(cell)) report.changed.push({ file: name, cell, before: null, after: after.get(cell) })
}
await writeFile(out, JSON.stringify(report, null, 1))
console.log(`${report.files} files, ${report.formulas} formulas: ${report.changed.length} changed, ${report.failed.length} failed to open or save`)
if (report.changed.length || report.failed.length) process.exit(1)
