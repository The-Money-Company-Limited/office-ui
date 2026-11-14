// node apps/v4/experiments/tsx-spike/debug.mjs <variant>: load once, print
// console output, errors, failed requests and the first cell anchors.
import path from "node:path"
import { fileURLToPath, pathToFileURL } from "node:url"

import { serve } from "./serve.mjs"

const here = path.dirname(fileURLToPath(import.meta.url))
const variant = process.argv[2] || "optionb"
const { chromium } = await import(pathToFileURL("/opt/cod1460/node_modules/playwright/index.mjs"))
const browser = await chromium.launch({ headless: true, args: ["--no-sandbox", "--disable-dev-shm-usage"] })
const { server, port } = await serve(path.join(here, "out", variant))
const context = await browser.newContext()
context.on("serviceworker", (worker) => console.log("service worker", worker.url()))
const page = await context.newPage()
page.on("console", (m) => console.log("console", m.type(), m.text().slice(0, 500)))
page.on("pageerror", (e) => console.log("pageerror", e.message, e.stack?.slice(0, 800)))
page.on("requestfailed", (r) => console.log("failed", r.url(), r.failure()?.errorText))
page.on("response", (r) => r.status() >= 400 && console.log("status", r.status(), r.url()))
await page.goto(`http://127.0.0.1:${port}/`)
if (process.argv[3] === "ready")
  await page.waitForFunction(
    () => document.querySelector('[aria-label="Formula bar"]') && document.getElementById("c-A1")?.textContent?.trim() === "Item",
    null,
    { polling: 16 }
  )
else await page.waitForTimeout(Number(process.argv[3] || 8000))
console.log(
  await page.evaluate(() => ({
    anchors: [...document.querySelectorAll("[id^=c-]")].slice(0, 6).map((e) => [e.id, e.textContent]),
    formula: !!document.querySelector('[aria-label="Formula bar"]'),
    ids: [...document.querySelectorAll("[id]")].slice(0, 30).map((e) => e.id),
  }))
)
if (process.argv[4] === "edit") {
  const cellText = (ref) => page.evaluate((id) => document.getElementById(id)?.textContent, `c-${ref}`)
  const name = page.getByRole("textbox", { name: "Name box", exact: true })
  const formula = page.getByRole("textbox", { name: "Formula bar", exact: true })
  console.log("before", await cellText("B2"), await cellText("C2"), await cellText("D2"))
  await name.fill("B2")
  await name.press("Enter")
  console.log("formula B2", await formula.inputValue())
  await formula.fill("100")
  await formula.press("Enter")
  for (let i = 0; i < 20; i++) {
    console.log("after", i * 250, await cellText("B2"), await cellText("D2"))
    await page.waitForTimeout(250)
  }
  console.log("after", await cellText("B2"), await cellText("C2"), await cellText("D2"))
  await name.fill("D2")
  await name.press("Enter")
  console.log("formula D2", await formula.inputValue())
  await page.getByRole("tab", { name: "Notes", exact: true }).click()
  await page.waitForTimeout(1000)
  console.log("notes A1", await cellText("A1"), page.url())
  await page.getByRole("tab", { name: "Budget", exact: true }).click()
  await page.waitForTimeout(1000)
  console.log("budget A1", await cellText("A1"), await cellText("D2"), page.url())
  console.log("tabs", await page.evaluate(() => [...document.querySelectorAll('[role=tab], button')].map((b) => b.textContent.trim()).filter((t) => /Budget|Notes/.test(t))))
}
if (process.argv[5] === "reload") {
  await page.goto(page.url().split("?")[0])
  await page.waitForFunction(() => document.querySelector('[aria-label="Formula bar"]') && document.getElementById("c-A1")?.textContent?.trim() === "Item", null, { polling: 16 })
  await page.waitForTimeout(1000)
}
console.log(
  "resources",
  await page.evaluate(() =>
    performance
      .getEntriesByType("resource")
      .map((r) => [Math.round(r.startTime), Math.round(r.responseEnd), r.name.split("/").slice(-1)[0].slice(0, 50), r.initiatorType])
      .sort((a, b) => b[1] - a[1])
      .slice(0, 25)
  )
)
console.log("marks", await page.evaluate(() => performance.getEntriesByType("mark").map((m) => [m.name, Math.round(m.startTime)])))
await page.screenshot({ path: path.join(here, "out", `debug-${variant}.png`) })
await browser.close()
server.close()
