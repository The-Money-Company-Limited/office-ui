// Headless Chrome measurements, run from the repository root on the box:
//   node apps/v4/experiments/tsx-spike/measure.mjs [runs]
// For each release (baseline, optionb) and header policy (public, private):
// cold = new browser context (empty HTTP cache, no service worker), warm =
// the same context loading the page again. Ready = the formula bar exists and
// cell A1 of the first sheet shows its text on the anchor layer.
import assert from "node:assert/strict"
import { readFile, writeFile } from "node:fs/promises"
import path from "node:path"
import { fileURLToPath, pathToFileURL } from "node:url"

import { serve } from "./serve.mjs"

const here = path.dirname(fileURLToPath(import.meta.url))
const out = path.join(here, "out")
const runs = Number(process.argv[2] || 3)
// A latency run serves HTTP/2 with that many ms before each answer, and
// skips the service-worker variant (no service worker on a self-signed origin).
const latency = Number(process.argv[3] || 0)
const tls = latency > 0
const scheme = tls ? "https" : "http"
if (tls) process.env.NODE_TLS_REJECT_UNAUTHORIZED = "0"
const { chromium } = await import(pathToFileURL("/opt/cod1460/node_modules/playwright/index.mjs"))
const browser = await chromium.launch({ headless: true, args: ["--no-sandbox", "--disable-dev-shm-usage"] })

const median = (values) => [...values].sort((a, b) => a - b)[Math.floor(values.length / 2)]

async function ready(page) {
  await page.waitForFunction(
    () =>
      document.querySelector('[aria-label="Formula bar"]') &&
      document.getElementById("c-A1")?.textContent?.trim() === "Item",
    null,
    { timeout: 60_000, polling: 16 }
  )
  return page.evaluate(() => performance.now())
}

async function stats(port) {
  return (await fetch(`${scheme}://127.0.0.1:${port}/__stats`)).json()
}

async function load(page, port) {
  await stats(port)
  await page.goto(`${scheme}://127.0.0.1:${port}/`)
  const ms = await ready(page)
  // Let lazy work started by the load finish before counting bytes.
  await page.waitForTimeout(1500)
  const marks = await page.evaluate(() => Object.fromEntries(performance.getEntriesByType("mark").map((m) => [m.name, Math.round(m.startTime)])))
  const loader = await page.evaluate(() => window.__loaderStats)
  const serviceWorker = await page.evaluate(
    () =>
      navigator.serviceWorker?.controller &&
      new Promise((done) => {
        navigator.serviceWorker.addEventListener("message", (e) => done(e.data.stats), { once: true })
        navigator.serviceWorker.controller.postMessage("stats")
      })
  )
  return { ms, ...(await stats(port)), marks, serviceWorker, loader }
}

const results = {}
for (const variant of tls ? ["baseline", "optionb"] : ["baseline", "optionb-sw", "optionb"]) {
  for (const policy of ["public", "private"]) {
    const { server, port } = await serve(path.join(out, variant), { policy, latency, tls })
    // One discarded load fills the server's gzip memo.
    {
      const context = await browser.newContext({ ignoreHTTPSErrors: true })
      await load(await context.newPage(), port)
      await context.close()
    }
    const cold = []
    const warm = []
    for (let i = 0; i < runs; i++) {
      const context = await browser.newContext({ ignoreHTTPSErrors: true })
      const page = await context.newPage()
      const errors = []
      page.on("pageerror", (e) => errors.push(e.message))
      cold.push(await load(page, port))
      warm.push(await load(page, port))
      if (errors.length) console.warn(variant, policy, errors)
      await context.close()
    }
    server.close()
    const summary = (list) => ({
      ms: Math.round(median(list.map((r) => r.ms))),
      bytes: median(list.map((r) => r.bytes)),
      requests: median(list.map((r) => r.requests)),
      notModified: median(list.map((r) => r.notModified)),
      all: list.map((r) => Math.round(r.ms)),
      marks: list.map((r) => r.marks),
      serviceWorker: list.map((r) => r.serviceWorker),
      loader: list.map((r) => r.loader),
    })
    results[`${variant}/${policy}`] = { cold: summary(cold), warm: summary(warm) }
    console.log(variant, policy, JSON.stringify(results[`${variant}/${policy}`]))
  }
}

if (tls) {
  await writeFile(path.join(out, `results-latency-${latency}.json`), JSON.stringify({ results }, null, 1) + "\n")
  await browser.close()
  process.exit(0)
}

// Functional checks on option B: formula recalculation, sheet tabs, and an
// edit to one source file on the served site showing after a reload.
const editFile = path.join(out, "optionb/src/apps/v4/components/extend/xlsx-editor.tsx")
const original = await readFile(editFile, "utf8")
const checks = {}
for (const variant of ["baseline", "optionb"]) {
  const { server, port } = await serve(path.join(out, variant), { policy: "public" })
  const context = await browser.newContext()
  const page = await context.newPage()
  await page.goto(`http://127.0.0.1:${port}/`)
  await ready(page)
  const cellText = (ref) => page.evaluate((id) => document.getElementById(id)?.textContent?.trim(), `c-${ref}`)
  const name = page.getByRole("textbox", { name: "Name box", exact: true })
  const formula = page.getByRole("textbox", { name: "Formula bar", exact: true })
  const before = { D2: await cellText("D2") }
  await name.fill("D2")
  await name.press("Enter")
  before.formulaD2 = await formula.inputValue()
  await name.fill("B2")
  await name.press("Enter")
  await formula.fill("100")
  await formula.press("Enter")
  await page.waitForFunction(() => document.getElementById("c-D2")?.textContent?.trim() === "700", null, { timeout: 10_000 })
  const after = { D2: await cellText("D2") }
  await name.fill("D202")
  await name.press("Enter")
  after.formulaD202 = await formula.inputValue()
  await page.getByRole("tab", { name: "Notes", exact: true }).click()
  await page.waitForFunction(() => document.getElementById("c-A1")?.textContent?.trim() === "Second sheet", null, { timeout: 10_000 })
  const sheetUrl = page.url()
  await page.getByRole("tab", { name: "Budget", exact: true }).click()
  await page.waitForFunction(() => new URL(location.href).searchParams.get("sheet") === "Budget", null, { timeout: 10_000 })
  await name.fill("D2")
  await name.press("Enter")
  after.budgetD2AfterTabs = await formula.inputValue()
  checks[variant] = { before, after, sheetUrl }
  await page.screenshot({ path: path.join(out, `${variant}.png`) })

  if (variant === "optionb") {
    // One agent edit: a toolbar label and a Tailwind class the shipped CSS lacks.
    assert.ok(original.includes('aria-label="Zoom in"'))
    await writeFile(
      editFile,
      original.replace('aria-label="Zoom in"', 'aria-label="Zoom in (edited on the site)" className="tracking-[0.37em]"')
    )
    const spacing = () =>
      page.evaluate(() => getComputedStyle(document.querySelector('[aria-label="Zoom in (edited on the site)"]')).letterSpacing)
    await stats(port)
    await page.goto(`http://127.0.0.1:${port}/`)
    const ms = await ready(page)
    await page.getByRole("button", { name: "Zoom in (edited on the site)", exact: true }).waitFor({ timeout: 5000 })
    const spacingBeforeCss = await spacing()
    await page.waitForFunction(() => window.__loaderStats.css === "rebuilt", null, { timeout: 30_000 })
    checks.edit = {
      reloadMs: Math.round(ms),
      server: await stats(port),
      loader: await page.evaluate(() => window.__loaderStats),
      labelShown: true,
      letterSpacing: { beforeTailwindRebuild: spacingBeforeCss, after: await spacing() },
    }
    // The next load uses the rebuilt script and CSS from the cache.
    await page.goto(`http://127.0.0.1:${port}/`)
    checks.editWarm = { ms: Math.round(await ready(page)), loader: await page.evaluate(() => window.__loaderStats), letterSpacing: await spacing(), server: await stats(port) }
    await writeFile(editFile, original)
  }
  await context.close()
  server.close()
}
console.log(JSON.stringify(checks, null, 1))
await writeFile(path.join(out, "results.json"), JSON.stringify({ results, checks }, null, 1) + "\n")
await browser.close()
