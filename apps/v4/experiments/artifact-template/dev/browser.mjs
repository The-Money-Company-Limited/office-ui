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
  if (kind === "xlsx")
    await page.locator('[data-cell-ref="A1"], #c-A1').first().waitFor()
  if (kind === "slides") await page.locator("#welcome-title").waitFor()
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
  const errors = []
  page.on("pageerror", (error) => errors.push(error.message))
  const file =
    kind === "slides"
      ? `${release}/deck/slides/welcome.html`
      : `${release}/document/sample.${kind}`
  try {
    if (scenario === "read-only") await control({ mode: "read_only" })
    await page.goto(url)
    await ready(page, kind, scenario === "read-only")
    const before = await readFile(file)
    const originalTitle =
      kind === "slides"
        ? await page.locator("#welcome-title").innerText()
        : "Autosave original"
    if (scenario === "read-only") {
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
      const text = `Persisted-${kind}-${scenario}`
      await edit(page, kind, text)
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
        await page.reload()
        await assertEdit(page, kind, text)
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
    ])
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
