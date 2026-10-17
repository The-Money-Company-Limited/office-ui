import assert from "node:assert/strict"
import { mkdir, readFile, writeFile } from "node:fs/promises"
import { createServer } from "node:http"
import path from "node:path"
import { pathToFileURL } from "node:url"
import JSZip from "jszip"

const root = process.cwd()
const { chromium } = await import(
  pathToFileURL(path.join(root, ".cod1460/node_modules/playwright/index.mjs"))
)
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

for (const action of ["drag", "resize", "nudge", "nested drag"]) {
  await run(`polish flow ${action}`, async (page) => {
    await page
      .context()
      .addCookies([{ name: "rotated", value: "1", url: base }])
    if (action === "nested drag") {
      await page.route("**/deck/slides/welcome.html", async (route) => {
        const response = await route.fetch()
        const html = (await response.text())
          .replace(
            '<h1 id="welcome-title"',
            '<div style="display: flex; flex-direction: column; gap: 50px;">' +
              '<h1 id="welcome-title"'
          )
          .replace(
            '</p>\n  <div style="position:',
            '</p></div>\n  <div style="position:'
          )
        await route.fulfill({ response, body: html })
      })
    }
    await open(page)
    const title = page.locator("#welcome-title")
    const beforeHTML = await page
      .locator("#welcome")
      .evaluate((el) => el.outerHTML)
    const boxes = () =>
      page.locator("#welcome").evaluate((canvas) => {
        const origin = canvas.getBoundingClientRect()
        return [...canvas.querySelectorAll("p, p span")]
          .sort((a, b) => a.id.localeCompare(b.id))
          .map((el) => {
            const rect = el.getBoundingClientRect()
            return [
              rect.x - origin.x,
              rect.y - origin.y,
              rect.width,
              rect.height,
            ]
          })
      })
    const before = await boxes()
    await title.click()
    assert.equal(
      await page.locator("#welcome").evaluate((el) => el.outerHTML),
      beforeHTML,
      "selecting does not pin flow"
    )
    const start = await title.boundingBox()
    if (action === "resize") {
      await drag(page, page.locator('[data-resize="se"]').first(), 60, 25)
    } else if (action === "nudge") {
      await page.keyboard.press("Shift+ArrowRight")
    } else {
      await drag(page, title, 70, 40)
    }
    const changed = await title.boundingBox()
    assert.ok(
      action === "resize"
        ? changed.width > start.width + 30
        : changed.x > start.x + 3
    )
    const after = await boxes()
    for (let i = 0; i < before.length; i++) {
      for (let j = 0; j < 4; j++) {
        assert.ok(
          Math.abs(after[i][j] - before[i][j]) < 0.5,
          `sibling ${i} coordinate ${j} stays at its original box: ${after[i][j]} ≈ ${before[i][j]}`
        )
      }
    }
    const editedHTML = await page
      .locator("#welcome")
      .evaluate((el) => el.outerHTML)
    await page.getByRole("button", { name: "Undo", exact: true }).click()
    assert.equal(
      await page.locator("#welcome").evaluate((el) => el.outerHTML),
      beforeHTML,
      "one undo restores all flow siblings"
    )
    assert.equal(
      await page
        .getByRole("button", { name: "Undo", exact: true })
        .isDisabled(),
      true
    )
    await page.getByRole("button", { name: "Redo", exact: true }).click()
    assert.equal(
      await page.locator("#welcome").evaluate((el) => el.outerHTML),
      editedHTML
    )
  })
}

await run("polish header and download menu", async (page) => {
  await open(page)
  const header = page.locator(".slides-header")
  assert.equal(
    await header.locator("strong").innerText(),
    "Cool Computers · A shared canvas"
  )
  assert.equal((await header.innerText()).includes("deck.json"), false)
  assert.equal(
    await header.locator("svg").count(),
    3,
    "file, theme and download icons"
  )
  assert.equal(await header.getByRole("button").count(), 3)
  assert.equal(
    await page.getByText("Click to select", { exact: false }).count(),
    0
  )
  assert.equal(
    await page.getByText("Edits stay in this tab.", { exact: false }).count(),
    0
  )
  for (const name of ["Undo", "Redo"]) {
    const button = page.getByRole("button", { name, exact: true })
    assert.equal(await button.isDisabled(), true)
    assert.equal(await button.innerText(), "")
    assert.equal(await button.locator("svg").count(), 1)
  }
  const download = page.getByRole("button", { name: "Download", exact: true })
  assert.equal(await download.innerText(), "")
  await download.focus()
  await page.keyboard.press("Enter")
  await page.getByRole("menuitem", { name: "Deck as .zip" }).waitFor()
  assert.equal(
    await page.getByRole("menuitem", { name: "PowerPoint .pptx" }).count(),
    1
  )
  await page.keyboard.press("Escape")
  await page.waitForFunction(
    () => document.activeElement?.getAttribute("aria-label") === "Download"
  )
  await page.getByRole("menuitem").first().waitFor({ state: "hidden" })
  assert.equal(await page.getByRole("menuitem").count(), 0)
  await page.screenshot({ path: `${output}/header.png` })
})

async function assertReturnedSlide(page, id, number) {
  await page.locator(".slides-toolbar").waitFor()
  await page
    .waitForFunction(
      ({ id, number }) => {
        const top = document.getElementById(id)?.getBoundingClientRect().top
        return (
          top >= 100 &&
          top < window.innerHeight / 3 &&
          document
            .querySelector(`[aria-label="Slide ${number}"]`)
            ?.getAttribute("aria-current") === "page"
        )
      },
      { id, number }
    )
    .catch(async (error) => {
      console.log(
        await page.evaluate(
          (id) => ({
            id,
            top: document.getElementById(id)?.getBoundingClientRect().top,
            active: document
              .querySelector('[aria-current="page"]')
              ?.getAttribute("aria-label"),
            scroll: window.scrollY,
            height: window.innerHeight,
            hash: window.location.hash,
          }),
          id
        )
      )
      throw error
    })
  assert.equal(new URL(page.url()).hash, `#${id}`)
  assert.equal(await page.locator(".slides-stage").count(), 4)
}

await run("polish native fullscreen exit", async (page) => {
  await open(page)
  await page.getByRole("button", { name: "Present", exact: true }).click()
  await page.waitForFunction(() => Boolean(document.fullscreenElement))
  await page.keyboard.press("ArrowRight")
  await page.keyboard.press("Space")
  await page.locator("#delivery").waitFor()
  // Chrome handles native Escape itself before dispatching a DOM key event.
  await page.evaluate(() => document.exitFullscreen())
  await assertReturnedSlide(page, "delivery", 3)
})

await run("polish fallback presentation scroll and Escape", async (page) => {
  await page.addInitScript(() => {
    Element.prototype.requestFullscreen = () =>
      Promise.reject(new Error("Unavailable"))
  })
  await open(page)
  await page.evaluate(() => {
    const background = document.createElement("div")
    background.style.height = "4000px"
    document.body.append(background)
  })
  await page.getByRole("button", { name: "Present", exact: true }).click()
  await page.locator(".slides-presenting").waitFor()
  const scroll = await page.evaluate(() => window.scrollY)
  await page.mouse.move(800, 500)
  await page.mouse.wheel(0, 700)
  await page.waitForTimeout(250)
  assert.equal(
    await page.evaluate(() => window.scrollY),
    scroll,
    "wheel cannot scroll the page behind presentation"
  )
  for (const [key, id] of [
    ["ArrowRight", "workflow"],
    ["ArrowDown", "delivery"],
    ["ArrowUp", "workflow"],
    ["Space", "delivery"],
    ["PageDown", "adoption"],
    ["PageUp", "delivery"],
  ]) {
    await page.keyboard.press(key)
    await page.locator(`#${id}`).waitFor()
    assert.equal(await page.evaluate(() => window.scrollY), scroll)
  }
  await page.keyboard.press("Escape")
  await assertReturnedSlide(page, "delivery", 3)
  const restoredScroll = await page.evaluate(() => window.scrollY)
  await page.mouse.wheel(0, 400)
  await page.waitForTimeout(250)
  assert.ok(
    (await page.evaluate(() => window.scrollY)) > restoredScroll,
    "page scroll restored after presentation"
  )
})

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
  await assertReturnedSlide(page, "adoption", 4)
})

await run("zip and editable PowerPoint export", async (page) => {
  await page.context().addCookies([{ name: "image", value: "1", url: base }])
  await open(page)
  await page.locator("#welcome-title").dblclick()
  await page.keyboard.insertText("Exported canvas")
  await page.keyboard.press("Escape")
  let pending = page.waitForEvent("download")
  await page.getByRole("button", { name: "Download", exact: true }).click()
  await page.getByRole("menuitem", { name: "Deck as .zip" }).click()
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
  await page.getByRole("button", { name: "Download", exact: true }).click()
  await page.getByRole("menuitem", { name: "PowerPoint .pptx" }).click()
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
  assert.ok(fourth.includes("<c:chart"), "native editable chart")
  assert.ok(
    (
      await pptx.file("ppt/notesSlides/notesSlide1.xml").async("string")
    ).includes("Introduce a deck")
  )
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
await run("charts editing and native export", async (page) => {
  await open(page)
  await page.getByRole("button", { name: "Slide 4", exact: true }).click()
  const chart = page.locator("#adoption-chart")
  await chart.click()
  const panel = page.getByRole("dialog", { name: "Chart controls" })
  await panel.getByRole("button", { name: "line chart", exact: true }).click()
  assert.equal(await chart.getAttribute("data-chart"), "line")
  await chart.dblclick({ position: { x: 40, y: 60 } })
  const value = panel.getByRole("textbox", {
    name: "Row 2 column 2",
    exact: true,
  })
  await value.fill("27")
  assert.ok(
    (await chart.locator("svg").textContent()).includes("27"),
    "live preview"
  )
  await value.press("Tab")
  await page.getByRole("button", { name: "Undo", exact: true }).click()
  assert.equal(
    await chart.locator("tr").nth(1).locator("td").nth(1).textContent(),
    "18",
    "one undo per cell edit"
  )
  await page.getByRole("button", { name: "Redo", exact: true }).click()
  await panel.getByRole("button", { name: "Add series", exact: true }).click()
  assert.equal(await chart.locator("tr").first().locator("th").count(), 3)
  await panel.getByRole("button", { name: "Style", exact: true }).click()
  await panel.getByLabel("Legend", { exact: true }).selectOption("right")
  assert.equal(await chart.getAttribute("data-legend"), "right")
  await page.getByRole("button", { name: "Undo", exact: true }).click()
  assert.equal(await chart.getAttribute("data-legend"), "none")
  assert.equal(await chart.getAttribute("data-chart"), "line")
  await panel.getByLabel("Title", { exact: true }).fill("Editable adoption")
  await panel.getByLabel("Title", { exact: true }).press("Tab")
  await page.screenshot({ path: `${output}/chart-dark.png` })
  await page.getByRole("button", { name: "Use light appearance" }).click()
  await page.screenshot({ path: `${output}/chart-light.png` })
  await panel.getByRole("button", { name: "Data", exact: true }).click()
  const first = panel.getByRole("textbox", {
    name: "Row 2 column 2",
    exact: true,
  })
  await first.fill("bad")
  await first.press("Tab")
  assert.ok(await panel.getByRole("alert").isVisible())
  await first.press("Escape")
  assert.equal(await panel.count(), 0)
  assert.equal(
    await chart.locator("tr").nth(1).locator("td").nth(1).textContent(),
    "27"
  )
  await chart.focus()
  await chart.press("Enter")
  assert.equal(
    await panel
      .getByRole("textbox", { name: "Row 1 column 1", exact: true })
      .evaluate((element) => document.activeElement === element),
    true
  )
  await panel
    .getByRole("textbox", { name: "Row 2 column 2", exact: true })
    .evaluate((element) => {
      const data = new DataTransfer()
      data.setData("text/plain", "42\t7\n32\t8")
      element.dispatchEvent(
        new ClipboardEvent("paste", {
          clipboardData: data,
          bubbles: true,
          cancelable: true,
        })
      )
    })
  assert.equal(
    await chart.locator("tr").nth(1).locator("td").nth(1).textContent(),
    "42"
  )
  await page.getByRole("button", { name: "Undo", exact: true }).click()
  assert.equal(
    await chart.locator("tr").nth(1).locator("td").nth(1).textContent(),
    "27",
    "one undo per TSV paste"
  )
  await panel.getByRole("button", { name: "Close chart controls" }).click()
  let pending = page.waitForEvent("download", { timeout: 45000 })
  await page.getByRole("button", { name: "Download", exact: true }).click()
  await page.getByRole("menuitem", { name: "PowerPoint .pptx" }).click()
  let download = await pending
  await download.saveAs(`${output}/editable-chart.pptx`)
  const pptx = await JSZip.loadAsync(
    await readFile(`${output}/editable-chart.pptx`)
  )
  const chartFile = Object.keys(pptx.files).find((name) =>
    /^ppt\/charts\/chart\d+.xml$/.test(name)
  )
  assert.ok(chartFile)
  const xml = await pptx.file(chartFile).async("string")
  assert.match(xml, /<c:lineChart>/)
  assert.match(xml, /<c:v>27<\/c:v>/)
  assert.match(xml, /<c:v>Series 2<\/c:v>/)
  assert.match(xml, /a4e1d1/i)
  assert.match(xml, /Editable adoption/)
  assert.ok(!xml.includes("<c:legend>"))
  assert.equal((xml.match(/<c:ser>/g) || []).length, 2)
  assert.ok(
    Object.keys(pptx.files).some((name) => name.endsWith(".xlsx")),
    "embedded editable workbook"
  )
  pending = page.waitForEvent("download")
  await page.getByRole("button", { name: "Download", exact: true }).click()
  await page.getByRole("menuitem", { name: "Deck as .zip" }).click()
  download = await pending
  await download.saveAs(`${output}/editable-chart.zip`)
  const zip = await JSZip.loadAsync(
    await readFile(`${output}/editable-chart.zip`)
  )
  const html = await zip.file("deck/slides/adoption.html").async("string")
  assert.match(html, /data-chart="line"/)
  assert.match(html, />27<\/td>/)
  assert.ok(!html.includes("<svg"), "ZIP keeps only chart source")
  await page.route("**/deck/slides/adoption.html", (route) =>
    route.fulfill({ contentType: "text/html", body: html })
  )
  await page.reload()
  assert.equal(await chart.getAttribute("data-chart"), "line")
  assert.ok((await chart.locator("svg").textContent()).includes("27"))
})
await run("charts seven native PowerPoint types", async (page) => {
  const types = ["column", "bar", "line", "area", "pie", "doughnut", "scatter"]
  await page.route("**/deck/slides/adoption.html", (route) =>
    route.fulfill({
      contentType: "text/html",
      body: `<section id="adoption"><div style="position: relative; width: 1920px; height: 1080px; box-shadow: 0px 0px 4px #000000;">${types.map((type, i) => `<figure data-chart="${type}" data-title="${type}" data-colors="rose,teal" data-legend="right" data-labels="true" data-stacked="true" data-format="currency" data-x-title="Horizontal" data-y-title="Vertical" style="position: absolute; left: ${128 + (i % 3) * 550}px; top: ${128 + Math.floor(i / 3) * 280}px; width: 500px; height: 260px;"><table><tr><th>X</th><th>Revenue</th></tr><tr><td>1</td><td>18</td></tr><tr><td>3</td><td>32</td></tr></table></figure>`).join("")}</div></section>`,
    })
  )
  await open(page)
  assert.equal(await page.locator("#adoption figure svg").count(), 7)
  const pending = page.waitForEvent("download", { timeout: 45000 })
  await page.getByRole("button", { name: "Download", exact: true }).click()
  await page.getByRole("menuitem", { name: "PowerPoint .pptx" }).click()
  const download = await pending
  await download.saveAs(`${output}/chart-types.pptx`)
  const pptx = await JSZip.loadAsync(
    await readFile(`${output}/chart-types.pptx`)
  )
  const files = Object.keys(pptx.files).filter((name) =>
    /^ppt\/charts\/chart\d+.xml$/.test(name)
  )
  assert.equal(files.length, 7)
  for (const [i, file] of files.entries()) {
    const xml = await pptx.file(file).async("string")
    assert.match(
      xml,
      new RegExp(`<c:${types[i] === "column" ? "bar" : types[i]}Chart>`)
    )
    assert.match(xml, /<c:v>18<\/c:v>/)
    assert.match(xml, /<c:v>32<\/c:v>/)
    assert.match(xml, /fda4af/i)
    assert.match(xml, /<c:legendPos val="r"/)
    assert.match(xml, /<c:showVal val="1"/)
    if (["column", "bar", "area"].includes(types[i]))
      assert.match(xml, /<c:grouping val="stacked"/)
    if (types[i] === "scatter") {
      assert.match(xml, /<c:xVal>/)
      assert.match(xml, /<c:yVal>/)
      assert.match(xml, /<c:v>3<\/c:v>/)
    }
  }
})
await browser.close()
await new Promise((resolve) => server.close(resolve))
await writeFile(`${output}/results.json`, JSON.stringify(results, null, 2))
if (results.some((result) => !result.pass)) process.exitCode = 1
