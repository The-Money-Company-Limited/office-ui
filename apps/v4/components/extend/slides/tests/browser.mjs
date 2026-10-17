import assert from "node:assert/strict"
import { mkdir, readFile, writeFile } from "node:fs/promises"
import { createServer } from "node:http"
import path from "node:path"
import { pathToFileURL } from "node:url"

const root = process.cwd()
const { chromium } = await import(
  pathToFileURL(path.join(root, ".cod1460/node_modules/playwright/index.mjs"))
)
const { default: JSZip } = await import("../node_modules/jszip/lib/index.js")
const bundle =
  process.env.SLIDES_BUNDLE || "apps/v4/experiments/artifact-template/dist"
const output = "apps/v4/components/extend/slides/tests/results"
await mkdir(output, { recursive: true })
const server = createServer(async (request, response) => {
  try {
    const url = new URL(request.url, "http://localhost")
    const file = url.pathname === "/" ? "index.html" : url.pathname.slice(1)
    if (file.includes("..")) throw new Error("Invalid path")
    let data = await readFile(path.join(bundle, file))
    if (
      file === "deck/slides/welcome.html" &&
      request.headers.cookie?.includes("unsafe=1")
    ) {
      data = Buffer.from(
        data
          .toString()
          .replace(
            "<aside>",
            '<script>window.slideScriptRan = true</script><iframe src="about:blank"></iframe><div id="welcome-unsafe" onclick="window.slideScriptRan=true" style="/* hidden */position: fixed; width: 10dvw; background: url(https://invalid.example/a); color: red;">Safe text</div><p id="welcome-unknown" style="made-up: 10px; color: blue; color: invalid;">Valid text</p><aside>'
          )
      )
    }
    if (
      file === "deck/slides/welcome.html" &&
      request.headers.cookie?.includes("image=1")
    ) {
      data = Buffer.from(
        data
          .toString()
          .replace(
            "<aside>",
            '<img id="welcome-test-image" alt="Embedded test image" src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/l9sAAAAASUVORK5CYII=" style="position: absolute; left: 1700px; top: 100px; width: 80px; height: 80px;"><hr style="position: absolute; left: 100px; top: 1000px; width: 900px; border-top: 2px solid #7dd3fc;"><aside>'
          )
      )
    }
    if (
      file === "deck/slides/welcome.html" &&
      request.headers.cookie?.includes("rotated=1")
    ) {
      data = Buffer.from(
        data
          .toString()
          .replace(
            'id="welcome-title" style="',
            'id="welcome-title" style="transform: rotate(3deg); '
          )
      )
    }
    const types = {
      ".html": "text/html",
      ".js": "text/javascript",
      ".css": "text/css",
      ".json": "application/json",
      ".wasm": "application/wasm",
    }
    response.setHeader(
      "Content-Type",
      types[path.extname(file)] || "application/octet-stream"
    )
    response.end(data)
  } catch {
    response.writeHead(404)
    response.end("Not found")
  }
})
await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve))
const browser = await chromium.launch({
  headless: true,
  args: ["--no-sandbox"],
  env: { ...process.env, HOME: path.join(root, ".cod1460") },
})
const base = `http://127.0.0.1:${server.address().port}`
const results = []
async function run(name, test) {
  if (process.env.SLIDES_CASE && !name.startsWith(process.env.SLIDES_CASE))
    return
  const page = await browser.newPage({
    viewport: { width: 1440, height: 1000 },
    acceptDownloads: true,
  })
  page.setDefaultTimeout(8000)
  const errors = []
  page.on("pageerror", (error) => errors.push(error.message))
  try {
    await test(page)
    assert.deepEqual(errors, [], "no uncaught browser errors")
    console.log(`PASS ${name}`)
    results.push({ name, pass: true })
  } catch (error) {
    console.log(
      `FAIL ${name}: ${error.stack}\nBrowser errors: ${errors.join("; ")}`
    )
    await page.screenshot({
      path: `${output}/${name.replaceAll(" ", "-")}-failure.png`,
    })
    results.push({ name, pass: false, error: error.message })
  } finally {
    await page.close()
  }
}
async function open(page) {
  await page.goto(base)
  await page.locator(".slides-stage #welcome-title").waitFor()
}
async function drag(page, locator, dx, dy, shift = false) {
  const box = await locator.boundingBox()
  assert.ok(box)
  const x = box.x + box.width / 2
  const y = box.y + box.height / 2
  if (shift) await page.keyboard.down("Shift")
  await page.mouse.move(x, y)
  await page.mouse.down()
  await page.mouse.move(x + dx, y + dy, { steps: 8 })
  await page.mouse.up()
  if (shift) await page.keyboard.up("Shift")
}
async function near(actual, expected, message) {
  assert.ok(
    Math.abs(actual - expected) < 3,
    `${message}: ${actual} ≈ ${expected}`
  )
}

await run("sanitizer and deck contract", async (page) => {
  const warnings = []
  page.on("console", (message) => {
    if (message.type() === "warning") warnings.push(message.text())
  })
  await page.context().addCookies([{ name: "unsafe", value: "1", url: base }])
  await open(page)
  assert.equal(await page.locator(".slides-stage").count(), 4)
  assert.equal(
    await page
      .locator(".slides-content script, .slides-content iframe")
      .count(),
    0
  )
  const unsafe = await page
    .locator(".slides-stage #welcome-unsafe")
    .evaluate((element) => ({
      click: element.getAttribute("onclick"),
      position: element.style.position,
      width: element.style.width,
      background: element.style.background,
      color: element.style.color,
    }))
  assert.deepEqual(unsafe, {
    click: null,
    position: "",
    width: "",
    background: "",
    color: "red",
  })
  assert.equal(await page.evaluate(() => window.slideScriptRan), undefined)
  assert.equal(
    await page.locator("#welcome-unknown").getAttribute("style"),
    "color: blue;"
  )
  assert.ok(
    warnings.some(
      (warning) =>
        warning.includes("welcome-unknown") &&
        warning.includes("invalid style color")
    )
  )
  assert.ok(
    warnings.some((warning) => warning.includes("Slide welcome: <script>"))
  )
  assert.ok(
    warnings.some(
      (warning) =>
        warning.includes("welcome-unsafe") && warning.includes("position")
    )
  )
  assert.equal(await page.locator(".slides-strip [id]").count(), 0)
  assert.equal(await page.locator("#welcome > aside").isVisible(), false)
  await near(
    await page
      .locator("#welcome")
      .evaluate((element) => parseFloat(getComputedStyle(element).width)),
    1920,
    "fixed canvas"
  )
})

await run("canvas edits and history", async (page) => {
  await page.context().addCookies([{ name: "rotated", value: "1", url: base }])
  await open(page)
  let title = page.locator(".slides-stage #welcome-title")
  const start = await title.boundingBox()
  await drag(page, title, 70, 40)
  assert.equal(
    await title.evaluate((element) => element.style.position),
    "absolute"
  )
  assert.match(
    await title.evaluate((element) => element.style.transform),
    /^matrix\(/,
    "drag preserves rotation"
  )
  const moved = await title.boundingBox()
  assert.ok(moved.x > start.x + 40 && moved.y > start.y + 20)
  assert.equal(
    await page.locator("#welcome > aside").count(),
    1,
    "pinning preserves final notes"
  )
  assert.equal(
    await page
      .locator("#welcome")
      .evaluate((element) => element.lastElementChild.tagName),
    "ASIDE"
  )
  const canvasBounds = await page.locator("#welcome").boundingBox()
  const titleBounds = await title.boundingBox()
  const centre = canvasBounds.x + canvasBounds.width / 2
  await page.mouse.move(
    titleBounds.x + titleBounds.width / 2,
    titleBounds.y + titleBounds.height / 2
  )
  await page.mouse.down()
  await page.mouse.move(centre + 2, titleBounds.y + titleBounds.height / 2, {
    steps: 6,
  })
  assert.ok(
    await page.locator(".slides-guide-x").count(),
    "centre snapping guide"
  )
  await page.mouse.up()
  const snapped = await title.boundingBox()
  await near(snapped.x + snapped.width / 2, centre, "snaps to slide centre")
  await page.keyboard.press("Shift+ArrowRight")
  const nudged = await title.boundingBox()
  const scale = await page
    .locator("#welcome")
    .evaluate((element) => element.getBoundingClientRect().width / 1920)
  await near(nudged.x - snapped.x, 10 * scale, "shift nudge")
  await drag(page, page.locator('[data-resize="se"]').first(), 60, 25, true)
  const resized = await title.boundingBox()
  assert.ok(resized.width > moved.width + 40)
  await near(
    resized.width / resized.height,
    moved.width / moved.height,
    "aspect ratio"
  )
  await title.dblclick()
  await page.keyboard.insertText("A canvas we can edit")
  await page.keyboard.press("Escape")
  assert.equal(await title.innerText(), "A canvas we can edit")
  await page.keyboard.press("Meta+z")
  assert.equal(await title.innerText(), "A shared canvas")
  await page.keyboard.press("Meta+Shift+z")
  assert.equal(await title.innerText(), "A canvas we can edit")
  await page.getByRole("button", { name: "Use light appearance" }).focus()
  await page.keyboard.press("Meta+z")
  assert.equal(
    await title.innerText(),
    "A shared canvas",
    "history shortcuts work from the toolbar"
  )
  await page.getByRole("button", { name: "Redo", exact: true }).click()
  await title.click()
  await page.keyboard.press("Meta+d")
  assert.equal(await page.locator("#welcome h1").count(), 2)
  await page.keyboard.press("Delete")
  assert.equal(await page.locator("#welcome h1").count(), 1)
  await page.keyboard.press("Meta+z")
  assert.equal(await page.locator("#welcome h1").count(), 2)
  await page.screenshot({ path: `${output}/editing.png` })
})

await run("window anchors theme and presentation", async (page) => {
  await open(page)
  await page.getByRole("button", { name: "Slide 3", exact: true }).click()
  assert.equal(new URL(page.url()).hash, "#delivery")
  assert.ok(
    await page
      .getByRole("button", { name: "Slide 3", exact: true })
      .evaluate((element) =>
        element.classList.contains("slides-thumbnail-active")
      )
  )
  assert.ok(await page.evaluate(() => window.scrollY > 500))
  assert.equal(
    await page.locator('[data-slide-id="delivery"] #delivery').count(),
    1
  )
  const scrollBeforeTheme = await page.evaluate(() => window.scrollY)
  await page.getByRole("button", { name: "Use light appearance" }).click()
  await near(
    await page.evaluate(() => window.scrollY),
    scrollBeforeTheme,
    "toolbar preserves document scroll"
  )
  assert.equal(
    await page
      .locator("html")
      .evaluate((element) => element.classList.contains("dark")),
    false
  )
  await page.getByRole("button", { name: "Present", exact: true }).click()
  await page.locator(".slides-presenting").waitFor()
  await page.waitForFunction(() => Boolean(document.fullscreenElement))
  assert.equal(await page.locator(".slides-stage").count(), 1)
  assert.equal(await page.locator("#delivery").count(), 1)
  assert.equal(await page.locator("#delivery > aside").isVisible(), false)
  await page.keyboard.press("PageDown")
  await page.locator("#adoption").waitFor()
  await page.keyboard.press("ArrowLeft")
  await page.locator("#delivery").waitFor()
  await page.keyboard.press("Space")
  await page.locator("#adoption").waitFor()
  await page.screenshot({ path: `${output}/presenting.png` })
  await page.keyboard.press("Escape")
  await page.locator(".slides-toolbar").waitFor()
  assert.equal(await page.locator(".slides-stage").count(), 4)
})

await run("zip and editable PowerPoint export", async (page) => {
  await page.context().addCookies([{ name: "image", value: "1", url: base }])
  await open(page)
  await page.locator("#welcome-title").dblclick()
  await page.keyboard.insertText("Exported canvas")
  await page.keyboard.press("Escape")
  let pending = page.waitForEvent("download")
  await page.getByRole("button", { name: "Download .zip" }).click()
  let download = await pending
  await download.saveAs(`${output}/deck.zip`)
  const zip = await JSZip.loadAsync(await readFile(`${output}/deck.zip`))
  const manifest = JSON.parse(await zip.file("deck/deck.json").async("string"))
  assert.equal(manifest.order.length, 4)
  assert.ok(
    (await zip.file("deck/slides/welcome.html").async("string")).includes(
      "Exported canvas"
    )
  )
  await page.route("**/deck/slides/*.html", async (route) => {
    const id = new URL(route.request().url()).pathname.split("/").at(-1)
    await route.fulfill({
      contentType: "text/html",
      body: await zip.file(`deck/slides/${id}`).async("string"),
    })
  })
  await page.reload()
  await page.locator("#welcome-title").waitFor()
  assert.equal(
    await page.locator("#welcome-title").innerText(),
    "Exported canvas"
  )
  pending = page.waitForEvent("download", { timeout: 45000 })
  await page.getByRole("button", { name: "Export .pptx" }).click()
  download = await pending
  await download.saveAs(`${output}/deck.pptx`)
  const pptx = await JSZip.loadAsync(await readFile(`${output}/deck.pptx`))
  assert.equal(
    Object.keys(pptx.files).filter((name) =>
      /^ppt\/slides\/slide\d+.xml$/.test(name)
    ).length,
    4
  )
  const first = await pptx.file("ppt/slides/slide1.xml").async("string")
  assert.ok(
    first.includes("Exported canvas") && first.includes("<p:sp>"),
    "native text boxes"
  )
  assert.ok(first.includes("<p:pic>"), "native embedded image")
  assert.ok(first.includes('prst="line"'), "native horizontal rule")
  const second = await pptx.file("ppt/slides/slide2.xml").async("string")
  assert.ok(second.includes('prst="roundRect"'), "native card shapes")
  const third = await pptx.file("ppt/slides/slide3.xml").async("string")
  assert.ok(
    third.includes("<a:tbl>") && third.includes("Authoring"),
    "native table"
  )
  const fourth = await pptx.file("ppt/slides/slide4.xml").async("string")
  assert.ok(fourth.includes("<p:pic>"), "SVG raster fallback")
  assert.ok(
    (
      await pptx.file("ppt/notesSlides/notesSlide1.xml").async("string")
    ).includes("Introduce a deck")
  )
  const media = Object.entries(pptx.files).filter(([name]) =>
    /^ppt\/media\/.*\.png$/.test(name)
  )
  const chart = await media.at(-1)[1].async("nodebuffer")
  assert.ok(chart.length > 5000, "chart raster contains rendered content")
  await writeFile(`${output}/chart.png`, chart)
  for (const [name, file] of Object.entries(pptx.files)) {
    if (!name.endsWith(".xml")) continue
    const xml = await file.async("string")
    assert.equal(
      await page.evaluate(
        (xml) =>
          new DOMParser()
            .parseFromString(xml, "application/xml")
            .querySelector("parsererror")?.textContent || null,
        xml
      ),
      null,
      name
    )
  }
})
await browser.close()
await new Promise((resolve) => server.close(resolve))
await writeFile(`${output}/results.json`, JSON.stringify(results, null, 2))
if (results.some((result) => !result.pass)) process.exitCode = 1
