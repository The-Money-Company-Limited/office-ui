// Times opening and saving the corpus's most formula-heavy workbooks with engine builds run back
// to back. Usage: node bench-load-save.mjs <xlsx-dir> <top> <runs> <label>=<wasm-dir>...
import { readFile, readdir } from "node:fs/promises"
import path from "node:path"

const [dir, topArg, runsArg, ...builds] = process.argv.slice(2)
const [top, runs] = [Number(topArg), Number(runsArg)]
const count = (x) => Number.isInteger(x) && x > 0
if (!dir || !count(top) || !count(runs) || !builds.length || builds.some((b) => b.indexOf("=") < 1)) {
  console.error("Usage: node bench-load-save.mjs <xlsx-dir> <top> <runs> <label>=<wasm-dir>...")
  process.exit(2)
}
const engines = []
for (const build of builds) {
  const at = build.indexOf("=")
  const [label, d] = [build.slice(0, at), build.slice(at + 1)]
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
let unreadable = 0
for await (const file of walk(dir)) {
  const bytes = new Uint8Array(await readFile(file))
  try {
    const wb = engines[0].wasm.Workbook.fromBytes(bytes)
    let n = 0
    for (let s = 0; s < wb.sheetCount; s++) n += wb.getSheet(s).formulaCount
    wb.free()
    files.push({ file: path.relative(dir, file), bytes, n })
  } catch {
    unreadable++
  }
}
if (unreadable) console.log(`${unreadable} workbooks didn't open and are left out`)
files.sort((a, b) => b.n - a.n)
const median = (xs) => {
  const sorted = [...xs].sort((a, b) => a - b)
  const mid = sorted.length >> 1
  return (sorted.length % 2 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2).toFixed(0)
}
// Opens and saves a workbook, returning the times taken.
const time = (wasm, bytes) => {
  let start = performance.now()
  const wb = wasm.Workbook.fromBytes(bytes)
  const load = performance.now() - start
  start = performance.now()
  wb.saveXlsxBytes()
  const save = performance.now() - start
  wb.free()
  return { load, save }
}
console.log(`file | formulas | ${engines.map((e) => `load ${e.label}`).join(" | ")} | ${engines.map((e) => `save ${e.label}`).join(" | ")}`)
for (const { file, bytes, n } of files.slice(0, top)) {
  const t = engines.map(() => ({ load: [], save: [] }))
  // One untimed pass each, then each run starts with the next build, so no build always goes first.
  engines.forEach(({ wasm }) => time(wasm, bytes))
  for (let r = 0; r < runs; r++) {
    for (let k = 0; k < engines.length; k++) {
      const i = (r + k) % engines.length
      const { load, save } = time(engines[i].wasm, bytes)
      t[i].load.push(load)
      t[i].save.push(save)
    }
  }
  console.log(`${file} | ${n} | ${t.map((x) => median(x.load)).join(" | ")} | ${t.map((x) => median(x.save)).join(" | ")}`)
}
