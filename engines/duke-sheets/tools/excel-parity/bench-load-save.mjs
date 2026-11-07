// Times opening and saving the corpus's most formula-heavy workbooks with engine builds run back
// to back. Usage: node bench-load-save.mjs <xlsx-dir> <top> <runs> <label>=<wasm-dir>...
import { readFile, readdir } from "node:fs/promises"
import path from "node:path"

const [dir, top, runs, ...builds] = process.argv.slice(2)
const engines = []
for (const build of builds) {
  const [label, d] = build.split("=")
  const wasm = await import(path.resolve(d, "duke_sheets_wasm.js"))
  wasm.initSync({ module: await readFile(path.resolve(d, "duke_sheets_wasm_bg.wasm")) })
  engines.push({ label, wasm })
}
async function* walk(d) {
  for (const e of await readdir(d, { withFileTypes: true })) {
    const p = path.join(d, e.name)
    if (e.isDirectory()) yield* walk(p)
    else if (/\.xlsx$/i.test(e.name) && !e.name.startsWith("~$")) yield p
  }
}
const files = []
for await (const file of walk(dir)) {
  const bytes = new Uint8Array(await readFile(file))
  try {
    const wb = engines[0].wasm.Workbook.fromBytes(bytes)
    let n = 0
    for (let s = 0; s < wb.sheetCount; s++) n += wb.getSheet(s).formulaCount
    wb.free()
    files.push({ file: path.relative(dir, file), bytes, n })
  } catch {}
}
files.sort((a, b) => b.n - a.n)
const median = (xs) => xs.sort((a, b) => a - b)[xs.length >> 1].toFixed(0)
console.log(`file | formulas | ${engines.map((e) => `load ${e.label}`).join(" | ")} | ${engines.map((e) => `save ${e.label}`).join(" | ")}`)
for (const { file, bytes, n } of files.slice(0, Number(top))) {
  const t = engines.map(() => ({ load: [], save: [] }))
  for (let r = 0; r < Number(runs); r++) {
    engines.forEach(({ wasm }, i) => {
      let start = performance.now()
      const wb = wasm.Workbook.fromBytes(bytes)
      t[i].load.push(performance.now() - start)
      start = performance.now()
      wb.saveXlsxBytes()
      t[i].save.push(performance.now() - start)
      wb.free()
    })
  }
  console.log(`${file} | ${n} | ${t.map((x) => median(x.load)).join(" | ")} | ${t.map((x) => median(x.save)).join(" | ")}`)
}
