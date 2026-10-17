#!/usr/bin/env node
import { mkdir, readFile, realpath } from "node:fs/promises"
import path from "node:path"

const usage = "node slides-check.mjs <deck-dir> [--json] [--png <dir>]"

async function main() {
  const args = process.argv.slice(2)
  const json = args.includes("--json")
  try {
    const deckDir = args.shift()
    if (!deckDir || deckDir.startsWith("--")) throw new Error(usage)
    let png
    while (args.length) {
      const flag = args.shift()
      if (flag === "--json") continue
      if (flag === "--png" && args[0] && !args[0].startsWith("--"))
        png = args.shift()
      else throw new Error(usage)
    }
    const directory = await realpath(deckDir)
    const validator = await readFile(
      new URL("./validate.js", import.meta.url),
      "utf8"
    )
    const charts = await readFile(
      new URL("./charts.js", import.meta.url),
      "utf8"
    )
    const chartsURL = `data:text/javascript;base64,${Buffer.from(charts).toString("base64")}`
    const validatorURL = `data:text/javascript;base64,${Buffer.from(validator.replace("./charts.js", chartsURL)).toString("base64")}`
    const { SLIDE_HEIGHT, SLIDE_WIDTH, validateManifest } = await import(
      validatorURL
    )
    const manifest = validateManifest(
      JSON.parse(await readFile(path.join(directory, "deck.json"), "utf8"))
    )
    const css = await readFile(new URL("./slides.css", import.meta.url), "utf8")
    const { chromium } = await import("playwright").catch(() => {
      throw new Error(
        "Install beside slides-check.mjs: npm install --no-save playwright; npx playwright install --with-deps chromium"
      )
    })
    const browser = await chromium.launch({ headless: true })
    const slides = []
    try {
      const page = await browser.newPage({
        viewport: { width: SLIDE_WIDTH, height: SLIDE_HEIGHT },
        deviceScaleFactor: 1,
      })
      await page.route("**/*", (route) => route.abort())
      await page.setContent(
        '<!doctype html><html><body style="margin:0"><div class="slides-content"></div></body></html>'
      )
      await page.addStyleTag({ content: css })
      await page.evaluate(
        async ({ validatorURL, chartsURL }) => {
          window.slidesValidator = await import(validatorURL)
          window.slidesCharts = await import(chartsURL)
        },
        { validatorURL, chartsURL }
      )
      if (png) await mkdir(png, { recursive: true })
      for (const id of manifest.order) {
        try {
          const file = await realpath(
            path.join(directory, "slides", `${id}.html`)
          )
          if (!file.startsWith(path.join(directory, "slides") + path.sep))
            throw new Error("Slide path escapes deck/slides")
          const source = await readFile(file, "utf8")
          const issues = await page.evaluate(
            async ({ source, id }) => {
              const issues = []
              const html = window.slidesValidator.sanitizeSlide(
                source,
                id,
                document,
                (message, diagnostic) => {
                  issues.push({
                    code: "markup",
                    severity: "error",
                    ...diagnostic,
                    message,
                    hint: `Remove or replace ${diagnostic.detail}; use the closed subset in FORMAT.md.`,
                  })
                }
              )
              document.querySelector(".slides-content").innerHTML = html
              for (const figure of document.querySelectorAll(
                "figure[data-chart]"
              )) {
                for (const message of window.slidesCharts.readChart(figure)
                  .errors)
                  issues.push({
                    code: "chart-data",
                    severity: "error",
                    element: figure.id,
                    text: figure.textContent.trim().slice(0, 80),
                    message,
                    hint: message,
                  })
              }
              window.slidesCharts.renderCharts(document)
              await document.fonts.ready
              await Promise.all(
                [...document.images].map((image) =>
                  image.decode().catch(() => {})
                )
              )
              return issues
            },
            { source, id }
          )
          const layout = await page.evaluate(inspectLayout)
          issues.push(...layout.issues)
          const visible = await page.screenshot({
            animations: "disabled",
            ...(png ? { path: path.join(png, `${id}.png`) } : {}),
          })
          // Hide only text paint, preserving geometry, backgrounds and currentColor.
          const hiddenStyle = await page.addStyleTag({
            content:
              ".slides-content, .slides-content * { -webkit-text-fill-color: transparent !important; text-decoration-color: transparent !important; } .slides-content :is(text, tspan) { fill: transparent !important; stroke: transparent !important; }",
          })
          const background = await page.screenshot({ animations: "disabled" })
          await hiddenStyle.evaluate((element) => element.remove())
          issues.push(
            ...(await page.evaluate(inspectContrast, {
              runs: layout.runs,
              visible: visible.toString("base64"),
              background: background.toString("base64"),
            }))
          )
          slides.push({ id, issues })
        } catch (error) {
          slides.push({
            id,
            issues: [
              {
                code: "slide-load",
                severity: "error",
                element: id,
                text: "",
                message: error.message,
                hint: `Supply a readable slides/${id}.html with one section whose ID is ${id}.`,
              },
            ],
          })
        }
      }
    } finally {
      await browser.close()
    }
    const issues = slides.flatMap((slide) => slide.issues)
    const errors = issues.filter((issue) => issue.severity === "error").length
    const warnings = issues.length - errors
    const report = { title: manifest.title, errors, warnings, slides }
    if (json) console.log(JSON.stringify(report, null, 2))
    else {
      console.log(
        `${manifest.title}: ${slides.length} slides, ${errors} errors, ${warnings} warnings`
      )
      for (const { id, issues } of slides) {
        console.log(
          `${id}: ${issues.length ? `${issues.filter((issue) => issue.severity === "error").length} errors, ${issues.filter((issue) => issue.severity === "warning").length} warnings` : "OK"}`
        )
        for (const issue of issues)
          console.log(
            `  ${issue.severity.toUpperCase()} ${issue.code} [${issue.element}] ${JSON.stringify(issue.text)}: ${issue.message} Fix: ${issue.hint}`
          )
      }
    }
    process.exitCode = errors ? 1 : 0
  } catch (error) {
    if (json)
      console.log(
        JSON.stringify({
          errors: 1,
          warnings: 0,
          slides: [],
          error: error.message,
        })
      )
    else console.error(`slides-check: ${error.message}`)
    process.exitCode = 1
  }
}

function inspectLayout() {
  const root = document.querySelector(".slides-content > section")
  const canvas = root.getBoundingClientRect()
  const issues = []
  const runs = []
  const boxes = new Map()
  const seen = new Set()
  const metrics = document.createElement("canvas").getContext("2d")
  const textStart = (element) =>
    (element.textContent || element.getAttribute("alt") || "")
      .replace(/\s+/g, " ")
      .trim()
      .slice(0, 80)
  const rect = (r) => ({
    left: r.left,
    top: r.top,
    right: r.right,
    bottom: r.bottom,
    width: r.width,
    height: r.height,
  })
  const outside = (a, b, x = true, y = true) =>
    (x && (a.left < b.left - 1 || a.right > b.right + 1)) ||
    (y && (a.top < b.top - 1 || a.bottom > b.bottom + 1))
  const add = (
    element,
    code,
    message,
    hint,
    severity = "error",
    extra = {}
  ) => {
    const key = `${element.id}:${code}:${extra.relatedElement || ""}`
    if (seen.has(key)) return
    seen.add(key)
    issues.push({
      code,
      severity,
      element: element.id,
      text: textStart(element),
      message,
      hint,
      ...extra,
    })
  }
  const visible = (element) => {
    if (
      element.closest("aside, title, desc") ||
      !element.getClientRects().length
    )
      return false
    for (
      let node = element;
      node && node !== root.parentElement;
      node = node.parentElement
    ) {
      const style = getComputedStyle(node)
      if (
        style.display === "none" ||
        style.visibility === "hidden" ||
        Number(style.opacity) === 0
      )
        return false
    }
    return true
  }
  const textNodes = []
  const textElements = new Set()
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT)
  while (walker.nextNode()) {
    const node = walker.currentNode
    if (!node.textContent.trim() || !visible(node.parentElement)) continue
    textNodes.push(node)
    for (
      let element = node.parentElement;
      element !== root;
      element = element.parentElement
    )
      textElements.add(element)
  }
  const marginOffenders = []
  for (const element of root.querySelectorAll("*")) {
    if (!visible(element)) continue
    const bounds = element.getBoundingClientRect()
    if (!bounds.width && !bounds.height) continue
    if (outside(bounds, canvas))
      add(
        element,
        "outside-slide",
        "Element extends beyond the 1920×1080 canvas.",
        "Move or resize the element to fit inside x=0–1920 and y=0–1080.",
        "warning",
        { bounds: rect(bounds) }
      )
    else if (
      textElements.has(element) &&
      !marginOffenders.some((parent) => parent.contains(element)) &&
      outside(bounds, { left: 128, top: 128, right: 1792, bottom: 952 })
    ) {
      marginOffenders.push(element)
      add(
        element,
        "margin",
        "Element enters the 128px safe margin.",
        "Move or resize this content group inside x=128–1792 and y=128–952.",
        "warning",
        { bounds: rect(bounds) }
      )
    }
    if (
      element instanceof HTMLImageElement &&
      (!element.complete || !element.naturalWidth)
    )
      add(
        element,
        "image-load",
        "Embedded image failed to decode.",
        "Replace src with a valid base64 PNG, JPEG, GIF or WebP image."
      )
  }
  for (const node of textNodes) {
    const element = node.parentElement
    let box = element
    while (
      box !== root &&
      getComputedStyle(box).display === "inline" &&
      !box.matches("text")
    )
      box = box.parentElement
    const range = document.createRange()
    const chars = []
    const style = getComputedStyle(element)
    metrics.font = `${style.fontStyle} ${style.fontWeight} ${style.fontSize} ${style.fontFamily}`
    for (let offset = 0; offset < node.length; ) {
      const char = String.fromCodePoint(node.textContent.codePointAt(offset))
      range.setStart(node, offset)
      range.setEnd(node, offset + char.length)
      const bounds = range.getBoundingClientRect()
      const measure = metrics.measureText(char)
      const scale =
        bounds.height /
        (measure.fontBoundingBoxAscent + measure.fontBoundingBoxDescent)
      const baseline = bounds.top + measure.fontBoundingBoxAscent * scale
      chars.push({
        char,
        baseline,
        ...rect(bounds),
        ink: {
          ...rect(bounds),
          top: baseline - measure.actualBoundingBoxAscent * scale,
          bottom: baseline + measure.actualBoundingBoxDescent * scale,
        },
      })
      offset += char.length
    }
    const glyphs = chars.filter(
      (char) => char.char.trim() && char.width && char.height
    )
    if (!glyphs.length) continue
    if (!boxes.has(box)) boxes.set(box, [])
    boxes.get(box).push(...chars)
    const svg = element instanceof SVGElement
    const bounds = box.getBoundingClientRect()
    if (!svg && glyphs.some((glyph) => outside(glyph.ink, bounds)))
      add(
        box,
        "text-overflow",
        "Text extends beyond its text box.",
        "Increase the box width/height, reduce font-size, or shorten the text."
      )
    // Canvas clipping belongs to the warning-only outside-slide check.
    for (
      let ancestor = element;
      ancestor && ancestor !== root;
      ancestor = ancestor.parentElement
    ) {
      const style = getComputedStyle(ancestor)
      const clipX = style.overflowX !== "visible"
      const clipY = style.overflowY !== "visible"
      if (!clipX && !clipY) continue
      const bounds = ancestor.getBoundingClientRect()
      if (
        glyphs.some((glyph) =>
          outside(svg ? glyph : glyph.ink, bounds, clipX, clipY)
        )
      )
        add(
          element,
          "text-clipped",
          `Text is clipped by ${ancestor.id}.`,
          `Enlarge ${ancestor.id}, move the text inside it, or remove its overflow clipping.`
        )
    }
    runs.push({
      element: element.id,
      text: textStart(element),
      size: parseFloat(getComputedStyle(element).fontSize),
      glyphs,
    })
  }
  for (const [box, chars] of boxes) {
    // Word segmentation keeps intentional spaces/hyphens and CJK line breaks legal.
    const text = chars.map((char) => char.char).join("")
    const segments = new Intl.Segmenter("en", { granularity: "word" }).segment(
      text
    )
    for (const segment of segments) {
      if (
        !segment.isWordLike ||
        /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}\p{Script=Thai}]/u.test(
          segment.segment
        )
      )
        continue
      const start = [...text.slice(0, segment.index)].length
      const word = chars.slice(start, start + [...segment.segment].length)
      if (
        word.some(
          (char, i) =>
            i &&
            Math.abs(char.baseline - word[i - 1].baseline) >
              Math.min(char.height, word[i - 1].height) / 2
        )
      )
        add(
          box,
          "mid-word-wrap",
          `Word ${JSON.stringify(segment.segment)} wraps across lines.`,
          "Widen the text box, shorten the word, or remove word-break/overflow-wrap forcing word splits."
        )
    }
  }
  const entries = [...boxes.keys()].filter((box) => box !== root)
  for (let i = 0; i < entries.length; i++) {
    for (const other of entries.slice(i + 1)) {
      const element = entries[i]
      if (element.contains(other) || other.contains(element)) continue
      const a = element.getBoundingClientRect()
      const b = other.getBoundingClientRect()
      if (
        Math.min(a.right, b.right) - Math.max(a.left, b.left) > 1 &&
        Math.min(a.bottom, b.bottom) - Math.max(a.top, b.top) > 1
      )
        add(
          element,
          "text-overlap",
          `Text box overlaps ${other.id} (${JSON.stringify(textStart(other))}).`,
          `Move or resize ${element.id} or ${other.id} so their text boxes do not overlap.`,
          "error",
          { relatedElement: other.id }
        )
    }
  }
  return { issues, runs }
}

async function inspectContrast({ runs, visible, background }) {
  const pixels = async (base64) => {
    const image = new Image()
    image.src = `data:image/png;base64,${base64}`
    await image.decode()
    const canvas = document.createElement("canvas")
    canvas.width = image.width
    canvas.height = image.height
    const context = canvas.getContext("2d", { willReadFrequently: true })
    context.drawImage(image, 0, 0)
    return context.getImageData(0, 0, canvas.width, canvas.height)
  }
  const [paint, back] = await Promise.all([pixels(visible), pixels(background)])
  const linear = Array.from({ length: 256 }, (_, n) => {
    const c = n / 255
    return c <= 0.04045 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4
  })
  const luminance = (data, i) =>
    linear[data[i]] * 0.2126 +
    linear[data[i + 1]] * 0.7152 +
    linear[data[i + 2]] * 0.0722
  const issues = new Map()
  for (const run of runs) {
    const threshold = run.size >= 44 ? 3 : 4.5
    let ratio = 21
    for (const glyph of run.glyphs) {
      if (
        glyph.left < 0 ||
        glyph.top < 0 ||
        glyph.right > paint.width ||
        glyph.bottom > paint.height
      )
        continue
      let strongest = 1
      for (let y = Math.ceil(glyph.top); y < Math.floor(glyph.bottom); y++) {
        for (let x = Math.ceil(glyph.left); x < Math.floor(glyph.right); x++) {
          const index = (y * paint.width + x) * 4
          const a = luminance(paint.data, index)
          const b = luminance(back.data, index)
          strongest = Math.max(
            strongest,
            (Math.max(a, b) + 0.05) / (Math.min(a, b) + 0.05)
          )
        }
      }
      ratio = Math.min(ratio, strongest)
    }
    if (ratio < threshold) {
      const previous = issues.get(run.element)
      if (previous && previous.ratio <= ratio) continue
      issues.set(run.element, {
        code: "contrast",
        severity: "error",
        element: run.element,
        text: run.text,
        ratio: Number(ratio.toFixed(2)),
        threshold,
        message: `Text contrast is ${ratio.toFixed(2)}:1; requires ${threshold}:1.`,
        hint: `Change the text/background colors or add an opaque backing to reach ${threshold}:1; check opacity too.`,
      })
    }
  }
  return [...issues.values()]
}

await main()
