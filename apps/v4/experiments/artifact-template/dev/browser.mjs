import assert from "node:assert/strict"
import { execFileSync } from "node:child_process"
import { mkdir, readFile, rm, writeFile } from "node:fs/promises"
import path from "node:path"
import { pathToFileURL } from "node:url"
import JSZip from "jszip"
import { PDFDocument } from "pdf-lib"

import { startHarness } from "./autosave-harness.mjs"
import { fixtures } from "./fixtures.mjs"

const root = path.resolve("apps/v4/experiments/artifact-template")
const output = `${root}/dev/results`
await mkdir(output, { recursive: true })
await fixtures(`${output}/fixtures`)
const { chromium } = await import(
  pathToFileURL(path.resolve(".cod1460/node_modules/playwright/index.mjs"))
)
const browser = await chromium.launch({
  headless: true,
  args: ["--no-sandbox", "--disable-dev-shm-usage"],
})
const outcomes = []
const negative = process.argv.includes("--negative")
const kinds = process.env.AUTOSAVE_KINDS?.split(",") ?? [
  "docx",
  "xlsx",
  "slides",
  "pdf",
  "pptx",
]

async function properties(page) {
  await page.getByRole("button", { name: "More actions", exact: true }).click()
  await page
    .getByRole("menuitem", { name: "Document properties…", exact: true })
    .click()
  await page.getByLabel("Title", { exact: true }).waitFor()
}
async function ready(page, kind, readOnly = false) {
  if (kind === "docx")
    await page
      .getByText("Autosave original", { exact: false })
      .first()
      .waitFor()
  if (kind === "xlsx" && !readOnly)
    await page
      .getByRole("textbox", { name: "Formula bar", exact: true })
      .waitFor()
  if (kind === "xlsx")
    await page.locator('[data-cell-ref="A1"], #c-A1').first().waitFor()
  if (kind === "slides") await page.locator("#welcome-title").waitFor()
  if (kind === "pptx") await page.getByText("of 1", { exact: true }).waitFor()
  if (kind === "pdf")
    await page
      .getByRole("button", { name: "More actions", exact: true })
      .waitFor()
  if (readOnly) await page.getByText("View only", { exact: true }).waitFor()
}
async function edit(page, kind, text) {
  if (kind === "docx") {
    const mode = page.getByRole("button", { name: /Mode: Viewing/ })
    if (await mode.count()) {
      await mode.click()
      await page.getByRole("menuitem", { name: "Editing", exact: true }).click()
    }
    const editor = page.locator('[data-docx-paragraph-node-index="0"]').first()
    await editor.click()
    await page.keyboard.press("Home")
    await page.keyboard.insertText(text + " ")
    await page
      .getByRole("button", { name: /Use light appearance|Use dark appearance/ })
      .click()
  } else if (kind === "xlsx") {
    const name = page.getByRole("textbox", { name: "Name box", exact: true })
    await name.fill("A1")
    await name.press("Enter")
    const formula = page.getByRole("textbox", {
      name: "Formula bar",
      exact: true,
    })
    await formula.fill(text)
    await formula.press("Enter")
  } else if (kind === "slides") {
    const title = page.locator("#welcome-title")
    await title.dblclick()
    await title.fill(text)
    await page
      .getByRole("button", { name: "Use light appearance", exact: true })
      .or(
        page.getByRole("button", { name: "Use dark appearance", exact: true })
      )
      .click()
  } else {
    await properties(page)
    await page.getByLabel("Title", { exact: true }).fill(text)
    await page.getByRole("button", { name: /^(Apply|Save)$/ }).click()
  }
}
async function assertEdit(page, kind, text) {
  if (kind === "docx")
    await page.getByText(text, { exact: false }).first().waitFor()
  if (kind === "xlsx") {
    await page
      .getByRole("textbox", { name: "Name box", exact: true })
      .fill("A1")
    await page
      .getByRole("textbox", { name: "Name box", exact: true })
      .press("Enter")
    assert.equal(
      await page
        .getByRole("textbox", { name: "Formula bar", exact: true })
        .inputValue(),
      text
    )
  }
  if (kind === "slides")
    assert.equal(await page.locator("#welcome-title").innerText(), text)
  if (kind === "pdf") {
    await properties(page)
    assert.equal(
      await page.getByLabel("Title", { exact: true }).inputValue(),
      text
    )
    await page.getByRole("button", { name: "Cancel", exact: true }).click()
  }
}
async function checkBytes(kind, buffer, text) {
  if (kind === "pdf")
    assert.equal((await PDFDocument.load(buffer)).getTitle(), text)
  else {
    const zip = await JSZip.loadAsync(buffer)
    if (kind === "docx")
      assert.match(
        await zip.file("word/document.xml").async("string"),
        new RegExp(text)
      )
    if (kind === "xlsx") {
      const xml = await Promise.all(
        Object.keys(zip.files)
          .filter((name) => name.endsWith(".xml"))
          .map((name) => zip.file(name).async("string"))
      )
      assert.ok(xml.some((value) => value.includes(text)))
    }
    if (kind === "slides")
      assert.match(
        await zip.file("deck/slides/welcome.html").async("string"),
        new RegExp(text)
      )
  }
}
async function run(kind, scenario) {
  const label = `${kind}-${scenario}`
  const release = `${output}/${label}`
  await rm(release, { recursive: true, force: true })
  execFileSync("node", [
    `${root}/stage.mjs`,
    kind === "slides"
      ? `${root}/samples/deck`
      : `${output}/fixtures/sample.${kind}`,
    release,
  ])
  if (scenario === "slide-ids") {
    await writeFile(
      `${release}/deck/deck.json`,
      JSON.stringify({ title: "Numeric IDs", order: ["9"] })
    )
    await writeFile(
      `${release}/deck/slides/9.html`,
      '<section id="9"><h1 id="9-title">Numeric slide ID</h1></section>'
    )
  }
  const server = await startHarness(release)
  const url = `http://127.0.0.1:${server.address().port}`
  const control = async (update) =>
    (
      await fetch(
        `${url}/__doc/control`,
        update ? { method: "POST", body: JSON.stringify(update) } : undefined
      )
    ).json()
  const context = await browser.newContext({
    viewport: { width: 1440, height: 1000 },
    acceptDownloads: true,
  })
  const page = await context.newPage()
  page.setDefaultTimeout(negative ? 5000 : 20000)
  if (["timeout-retry", "late-saved"].includes(scenario)) {
    await page.addInitScript(() => {
      const timeout = window.setTimeout.bind(window)
      window.setTimeout = (fn, ms, ...args) =>
        timeout(
          fn,
          ms === 19 * 60_000 ? 5000 : ms === 20 * 60_000 ? 6000 : ms,
          ...args
        )
    })
  }
  if (scenario === "webmcp")
    await page.addInitScript(() => {
      // Stands in for the browser's WebMCP host: it keeps what the page registers.
      const tools = new Map()
      window.webmcpTools = tools
      Object.defineProperty(Document.prototype, "modelContext", {
        configurable: true,
        value: {
          registerTool: async (tool, { signal }) => {
            tools.set(tool.name, tool)
            signal.addEventListener("abort", () => tools.delete(tool.name))
          },
        },
      })
    })
  await page.addInitScript(() => {
    window.cellDraws = []
    const fillText = CanvasRenderingContext2D.prototype.fillText
    CanvasRenderingContext2D.prototype.fillText = function (text, ...args) {
      if (
        ["7319", "0042", "42"].includes(String(text)) &&
        this.textAlign !== "center"
      )
        window.cellDraws.push({ text: String(text), align: this.textAlign })
      return fillText.call(this, text, ...args)
    }
    window.saveMessages = []
    addEventListener("message", (event) => {
      if (event.source === window && event.data?.cool === "doc")
        window.saveMessages.push(event.data)
    })
  })
  const errors = []
  page.on("pageerror", (error) => errors.push(error.message))
  const file =
    kind === "slides"
      ? `${release}/deck/slides/${scenario === "slide-ids" ? "9" : "welcome"}.html`
      : `${release}/document/sample.${kind}`
  try {
    if (scenario === "not-document-access")
      await control({ mode: "not_document" })
    if (scenario === "read-only") await control({ mode: "read_only" })
    await page.goto(url)
    if (scenario === "slide-ids") await page.locator('[id="9-title"]').waitFor()
    else if (scenario !== "not-document-access")
      await ready(page, kind, scenario === "read-only")
    const before = await readFile(file)
    const originalTitle =
      kind === "slides" && scenario !== "slide-ids"
        ? await page.locator("#welcome-title").innerText()
        : "Autosave original"
    if (scenario === "webmcp") {
      // An agent in the browser reads the open document, unsaved edits included; nothing writes.
      const call = (name, input = {}) =>
        page.evaluate(
          async ([name, input]) => window.webmcpTools.get(name).execute(input),
          [name, input]
        )
      const tools = () =>
        page.evaluate(() =>
          [...window.webmcpTools.values()].map((tool) => ({
            name: tool.name,
            annotations: tool.annotations,
          }))
        )
      const readOnly = { readOnlyHint: true, untrustedContentHint: true }
      if (kind === "docx") {
        await page.waitForFunction(() => window.webmcpTools.size === 2)
        assert.deepEqual(await tools(), [
          { name: "doc_read", annotations: readOnly },
          { name: "doc_find", annotations: readOnly },
        ])
        const found = await call("doc_find", { text: "original" })
        assert.deepEqual(
          found.matches.map(({ start, end }) => [start, end]),
          [[9, 17]]
        )
        await edit(page, kind, "Unsaved")
        const { paragraphs } = await call("doc_read")
        assert.equal(paragraphs[0].text, "Unsaved Autosave original")
      } else if (kind === "xlsx") {
        await page.waitForFunction(() => window.webmcpTools.size === 1)
        assert.deepEqual(await tools(), [
          { name: "sheet_read", annotations: readOnly },
        ])
        const all = await call("sheet_read")
        assert.deepEqual(all.sheets, ["Sheet1"])
        assert.deepEqual(
          all.cells.find((cell) => cell.ref === "A1")?.value,
          "Autosave original"
        )
        assert.deepEqual((await call("sheet_read", { range: "B1" })).cells, [
          { ref: "B1", value: 42 },
        ])
      } else {
        await page.waitForTimeout(1000)
        assert.deepEqual(await tools(), [])
      }
    } else if (scenario === "undo" && kind === "docx") {
      // Undo takes back the last word typed, as in Google Docs, not one key or the whole burst.
      const mode = page.getByRole("button", { name: /Mode: Viewing/ })
      if (await mode.count()) {
        await mode.click()
        await page.getByRole("menuitem", { name: "Editing", exact: true }).click()
      }
      const paragraph = page
        .locator('[data-docx-paragraph-node-index="0"]')
        .first()
      await paragraph.click()
      await page.keyboard.press("End")
      await page.keyboard.type(" alpha beta", { delay: 40 })
      const text = async () => (await paragraph.innerText()).trimEnd()
      await page.waitForFunction(
        () => document.body.innerText.includes("alpha beta")
      )
      await page.keyboard.press("ControlOrMeta+z")
      assert.equal(await text(), "Autosave original alpha")
      await page.keyboard.press("ControlOrMeta+z")
      assert.equal(await text(), "Autosave original")
      await page.keyboard.press("ControlOrMeta+Shift+z")
      assert.equal(await text(), "Autosave original alpha")
    } else if (scenario === "undo" && kind === "xlsx") {
      // One undo takes back one committed cell edit.
      const name = page.getByRole("textbox", { name: "Name box", exact: true })
      const formula = page.getByRole("textbox", {
        name: "Formula bar",
        exact: true,
      })
      const cell = async (ref) => {
        await name.fill(ref)
        await name.press("Enter")
        return formula.inputValue()
      }
      for (const [ref, value] of [
        ["A3", "first"],
        ["A4", "second"],
      ]) {
        await name.fill(ref)
        await name.press("Enter")
        await page.keyboard.type(value, { delay: 30 })
        await page.keyboard.press("Enter")
      }
      await page.keyboard.press("ControlOrMeta+z")
      assert.equal(await cell("A4"), "")
      assert.equal(await cell("A3"), "first")
      await page.locator('[data-cell-ref="B3"], #c-B3').first().click()
      await page.keyboard.press("ControlOrMeta+z")
      assert.equal(await cell("A3"), "")
    } else if (scenario === "header") {
      // The header names the file and switches the page's theme, which a reload keeps.
      await page.getByText(`sample.${kind}`, { exact: true }).first().waitFor()
      const dark = () =>
        page.evaluate(() => document.documentElement.classList.contains("dark"))
      const wasDark = await dark()
      await page
        .getByRole("button", {
          name: wasDark ? "Use light appearance" : "Use dark appearance",
          exact: true,
        })
        .click()
      assert.equal(await dark(), !wasDark)
      await page.reload()
      await ready(page, kind)
      assert.equal(await dark(), !wasDark)
    } else if (scenario === "slide-ids") {
      const title = page.locator('[id="9-title"]')
      await title.dblclick()
      await title.fill("Numeric ID edited")
      await page
        .getByRole("button", {
          name: /Use light appearance|Use dark appearance/,
        })
        .click()
      await page
        .locator("[data-autosave-status]")
        .filter({ hasText: /^Saved$/ })
        .waitFor()
      assert.match(await readFile(file, "utf8"), /Numeric ID edited/)
      await page.reload()
      await page
        .getByText("Numeric ID edited", { exact: true })
        .first()
        .waitFor()
    } else if (scenario === "not-document-access") {
      await page
        .getByText("This page is no longer a document", { exact: false })
        .first()
        .waitFor()
      assert.equal((await control()).saves.length, 0)
    } else if (["path_not_allowed", "not_document"].includes(scenario)) {
      await control({ mode: scenario, pollFailed: true })
      await edit(page, kind, "Retained local edit")
      await page
        .locator("[data-autosave-status]")
        .filter({
          hasText:
            scenario === "not_document"
              ? /This page is no longer a document/
              : /path outside the document/,
        })
        .waitFor()
      assert.deepEqual(await readFile(file), before)
      assert.equal((await control()).saves.length, 0)
      if (scenario === "path_not_allowed")
        assert.equal(
          await page.getByText("View only", { exact: true }).count(),
          0
        )
    } else if (scenario === "beforeunload") {
      await control({ mode: "offline", retryAfter: 60 })
      await edit(page, kind, "Unsaved edit")
      await page.getByText("Offline, will retry", { exact: true }).waitFor()
      const prompt = page.waitForEvent("dialog", { timeout: 5000 })
      const navigation = page.reload().catch(() => {})
      const dialog = await prompt
      assert.equal(dialog.type(), "beforeunload")
      await dialog.dismiss()
      await navigation
      await assertEdit(page, kind, "Unsaved edit")
    } else if (scenario === "read-only") {
      if (kind === "docx") {
        assert.equal(await page.locator('[contenteditable="true"]').count(), 0)
        assert.equal(
          await page.getByRole("button", { name: /Mode:/ }).count(),
          0
        )
      }
      if (kind === "xlsx") {
        assert.equal(
          await page.getByRole("textbox", { name: "Formula bar" }).count(),
          0
        )
        assert.equal(
          await page.getByRole("button", { name: "Undo", exact: true }).count(),
          0
        )
      }
      if (kind === "slides") {
        await page.locator("#welcome-title").dblclick()
        assert.equal(await page.locator("[contenteditable]").count(), 0)
        assert.equal(
          await page.getByRole("button", { name: "Undo", exact: true }).count(),
          0
        )
      }
      if (kind === "pdf") {
        assert.equal(
          await page.getByRole("button", { name: /Mode:/ }).count(),
          0
        )
        await properties(page)
        assert.ok(await page.getByLabel("Title", { exact: true }).isDisabled())
      }
      assert.deepEqual(await readFile(file), before)
      assert.equal((await control()).saves.length, 0)
    } else {
      if (scenario === "conflict") await control({ mode: "conflict" })
      if (scenario === "offline") await control({ mode: "offline" })
      if (scenario === "retry-after")
        await control({ mode: "offline", retryAfter: 5 })
      if (scenario === "failed-poll") await control({ pollFailed: true })
      if (["timeout-retry", "late-saved"].includes(scenario))
        await control({ receipt: "late" })
      let text =
        scenario === "alignment" ? "7319" : `Persisted-${kind}-${scenario}`
      await edit(page, kind, text)
      if (["timeout-retry", "late-saved"].includes(scenario)) {
        await page.waitForFunction(() =>
          window.saveMessages.some((message) => message.type === "save")
        )
        if (scenario === "late-saved") {
          text = `Newer-${kind}`
          await edit(page, kind, text)
        }
        await page.getByText("Offline, will retry", { exact: true }).waitFor()
        if (scenario === "late-saved") await control({ deliver: true })
      }
      if (scenario === "retry-after") {
        await page.getByText("Offline, will retry", { exact: true }).waitFor()
        await control({ mode: "online" })
        await page.evaluate(() => {
          Object.defineProperty(document, "visibilityState", {
            configurable: true,
            value: "hidden",
          })
          document.dispatchEvent(new Event("visibilitychange"))
        })
        await page.waitForTimeout(2500)
        assert.equal(
          await page.evaluate(
            () =>
              window.saveMessages.filter((message) => message.type === "save")
                .length
          ),
          1
        )
      }
      if (scenario === "failed-poll") {
        assert.equal(
          await page.getByText("View only", { exact: true }).count(),
          0
        )
        assert.equal(
          await page.evaluate(() =>
            window.saveMessages.some(
              (message) => message.type === "access" && !message.canSave
            )
          ),
          false
        )
      }
      if (scenario === "conflict") {
        await page
          .getByText("This was changed while you were editing.", {
            exact: true,
          })
          .waitFor()
        const pending = page.waitForEvent("download")
        await page
          .getByRole("button", { name: "Keep mine as a copy", exact: true })
          .click()
        const download = await pending
        const target = `${output}/${label}-copy.${kind === "slides" ? "zip" : kind}`
        await download.saveAs(target)
        await checkBytes(kind, await readFile(target), text)
        assert.deepEqual(await readFile(file), before)
        assert.equal((await control()).saves.length, 0)
        await page.screenshot({ path: `${output}/${label}.png` })
        await page
          .getByRole("button", { name: "Reload theirs", exact: true })
          .click()
        await ready(page, kind)
        assert.equal(
          await page
            .getByText("This was changed while you were editing.", {
              exact: true,
            })
            .count(),
          0
        )
      } else {
        if (scenario === "offline") {
          await page.getByText("Offline, will retry", { exact: true }).waitFor()
          assert.deepEqual(await readFile(file), before)
          await control({ mode: "online" })
        }
        await page
          .locator("[data-autosave-status]")
          .filter({ hasText: /^Saved$/ })
          .waitFor()
        assert.ok((await control()).saves.length >= 1)
        if (scenario === "timeout-retry") {
          const result = await control()
          assert.equal(result.saves.length, 1)
          assert.ok(result.attempts.length >= 2)
          assert.deepEqual(result.attempts[1].files, result.attempts[0].files)
          assert.equal(result.attempts[1].id, result.attempts[0].id)
          assert.equal(result.attempts[1].base, result.attempts[0].base)
          await control({ deliver: true })
        }
        if (scenario === "late-saved") {
          const result = await control()
          assert.equal(result.saves.length, 2)
          assert.notEqual(result.saves[1].id, result.saves[0].id)
          assert.equal(result.saves[1].base, "2")
        }
        if (scenario === "retry-after") {
          const result = await control()
          assert.ok(result.attempts[1].time - result.attempts[0].time >= 4900)
        }
        if (kind === "slides")
          assert.deepEqual((await control()).saves[0].paths.sort(), [
            "deck/deck.json",
            "deck/slides/welcome.html",
          ])
        if (kind !== "pdf")
          assert.ok(
            await page
              .getByRole("button", { name: "Undo", exact: true })
              .isEnabled()
          )
        if (scenario === "failed-poll") await control({ pollFailed: false })
        await page.reload()
        await assertEdit(page, kind, text)
        if (scenario === "alignment") {
          const alignment = async () => {
            await page.waitForFunction(() =>
              ["7319", "0042", "42"].every((text) =>
                window.cellDraws.some((draw) => draw.text === text)
              )
            )
            return page.evaluate(() =>
              Object.fromEntries(
                window.cellDraws.map((draw) => [draw.text, draw.align])
              )
            )
          }
          assert.deepEqual(await alignment(), {
            7319: "right",
            42: "right",
            "0042": "left",
          })
          await control({ mode: "read_only" })
          await page.reload()
          await ready(page, kind, true)
          assert.deepEqual(await alignment(), {
            7319: "right",
            42: "right",
            "0042": "left",
          })
        }
        if (scenario === "saved" && kind === "pdf") {
          await page
            .getByRole("button", { name: "Mode: View", exact: true })
            .click()
          await page
            .getByRole("menuitemradio", { name: "Annotate", exact: true })
            .click()
          await page
            .getByRole("button", { name: "Rectangle", exact: true })
            .click()
          const box = await page.locator("#page-1").boundingBox()
          assert.ok(box)
          await page.mouse.move(box.x + 100, box.y + 150)
          await page.mouse.down()
          await page.mouse.move(box.x + 250, box.y + 230, { steps: 8 })
          await page.mouse.up()
          await page
            .locator("[data-autosave-status]")
            .filter({ hasText: /^Saved$/ })
            .waitFor()
          const pdf = await PDFDocument.load(await readFile(file))
          assert.equal(pdf.getPage(0).node.Annots()?.size(), 1)
          assert.ok(
            await page
              .getByRole("button", { name: "Undo", exact: true })
              .isEnabled()
          )
          await page.reload()
          await assertEdit(page, kind, text)
        }
        if (scenario === "saved") {
          await page.evaluate(() => {
            window.autosaveReloadSentinel = true
          })
          await writeFile(file, before)
          await control({ changed: true })
          if (kind === "docx") {
            await page
              .getByText(text, { exact: false })
              .first()
              .waitFor({ state: "hidden" })
            await page
              .getByText("Autosave original", { exact: false })
              .first()
              .waitFor()
          } else {
            // Wait for the new document to finish loading before opening a dialog or typing a cell address.
            await page.waitForFunction(
              () =>
                document.querySelector("[data-autosave-status]") &&
                !document.body.innerText.includes("Opening document")
            )
            if (kind === "slides")
              await page.waitForFunction(
                (title) =>
                  document.getElementById("welcome-title")?.textContent ===
                  title,
                originalTitle
              )
            else await page.waitForTimeout(1500)
            await assertEdit(page, kind, originalTitle)
          }
          assert.equal(
            await page.evaluate(() => window.autosaveReloadSentinel),
            true
          )
        }
      }
    }
    await page.screenshot({ path: `${output}/${label}-final.png` })
    assert.deepEqual(errors, [])
    outcomes.push({ label, passed: true })
    console.log(`PASS ${label}`)
  } catch (error) {
    await page
      .screenshot({ path: `${output}/${label}-failure.png` })
      .catch(() => {})
    await writeFile(
      `${output}/${label}-failure.txt`,
      `${error.stack}\n${errors.join("\n")}\n${await page.locator("body").innerText()}`
    )
    outcomes.push({ label, passed: false, error: error.message })
    console.log(`FAIL ${label}: ${error.message.split("\n")[0]}`)
  } finally {
    await context.close()
    server.closeAllConnections()
    await new Promise((resolve) => server.close(resolve))
  }
}
try {
  for (const kind of kinds)
    for (const scenario of process.env.AUTOSAVE_SCENARIOS?.split(",") ?? [
      "saved",
      "conflict",
      "read-only",
      "offline",
      "timeout-retry",
      "late-saved",
      "beforeunload",
      "retry-after",
      "failed-poll",
      "path_not_allowed",
      "not_document",
      "not-document-access",
      "alignment",
      "slide-ids",
      "header",
      "undo",
      "webmcp",
    ])
      if (
        (scenario !== "alignment" || kind === "xlsx") &&
        (scenario !== "slide-ids" || kind === "slides") &&
        (scenario !== "header" || kind === "pptx" || kind === "pdf") &&
        (scenario !== "undo" || kind === "docx" || kind === "xlsx") &&
        (scenario !== "webmcp" || ["docx", "xlsx", "pdf"].includes(kind)) &&
        (kind !== "pptx" || scenario === "header")
      )
        await run(kind, scenario)
} finally {
  await browser.close()
}
await writeFile(
  `${output}/${negative ? "negative" : "browser"}.json`,
  JSON.stringify(outcomes, null, 2)
)
assert.ok(
  outcomes.every((result) => (negative ? !result.passed : result.passed)),
  JSON.stringify(
    outcomes.filter((result) => (negative ? result.passed : !result.passed)),
    null,
    2
  )
)
