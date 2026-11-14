// node apps/v4/experiments/tsx-spike/profile.mjs <variant>: CPU profile of a
// warm load (main thread), self time summed by script, plus the top functions.
import path from "node:path"
import { fileURLToPath, pathToFileURL } from "node:url"

import { serve } from "./serve.mjs"

const here = path.dirname(fileURLToPath(import.meta.url))
const variant = process.argv[2] || "optionb"
const { chromium } = await import(pathToFileURL("/opt/cod1460/node_modules/playwright/index.mjs"))
const browser = await chromium.launch({ headless: true, args: ["--no-sandbox", "--disable-dev-shm-usage"] })
const { server, port } = await serve(path.join(here, "out", variant))
const context = await browser.newContext()
const page = await context.newPage()
const ready = () =>
  page.waitForFunction(
    () => document.querySelector('[aria-label="Formula bar"]') && document.getElementById("c-A1")?.textContent?.trim() === "Item",
    null,
    { polling: 16 }
  )
await page.goto(`http://127.0.0.1:${port}/`)
await ready()
await page.waitForTimeout(1500)
const cdp = await context.newCDPSession(page)
await cdp.send("Profiler.enable")
await cdp.send("Profiler.setSamplingInterval", { interval: 200 })
await cdp.send("Profiler.start")
await page.goto(`http://127.0.0.1:${port}/`)
await ready()
const ms = await page.evaluate(() => performance.now())
const { profile } = await cdp.send("Profiler.stop")
const byId = new Map(profile.nodes.map((n) => [n.id, n]))
const self = new Map()
const deltas = profile.timeDeltas
for (let i = 0; i < profile.samples.length; i++) {
  const node = byId.get(profile.samples[i])
  const key = node.callFrame.url.split("/").pop() || `(${node.callFrame.functionName || "program"})`
  self.set(key, (self.get(key) || 0) + (deltas[i] || 0) / 1000)
}
const fns = new Map()
for (let i = 0; i < profile.samples.length; i++) {
  const f = byId.get(profile.samples[i]).callFrame
  const key = `${f.functionName || "(anon)"} ${f.url.split("/").pop()}:${f.lineNumber}`
  fns.set(key, (fns.get(key) || 0) + (deltas[i] || 0) / 1000)
}
const top = (m, n) => [...m].sort((a, b) => b[1] - a[1]).slice(0, n).map(([k, v]) => `${Math.round(v)}ms ${k}`)
console.log(variant, "ready", Math.round(ms))
console.log(top(self, 15).join("\n"))
console.log("--")
console.log(top(fns, 15).join("\n"))
await browser.close()
server.close()
