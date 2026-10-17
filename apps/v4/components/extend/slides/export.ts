import { toPng } from "html-to-image"
import JSZip from "jszip"
import pptxgen from "pptxgenjs"

import { SLIDE_HEIGHT, SLIDE_WIDTH, type Deck } from "./validate.js"

function saveBlob(blob: Blob, name: string) {
  const url = URL.createObjectURL(blob)
  const anchor = document.createElement("a")
  anchor.href = url
  anchor.download = name
  document.body.append(anchor)
  anchor.click()
  anchor.remove()
  setTimeout(() => URL.revokeObjectURL(url), 60_000)
}

function fileName(title: string) {
  return title.replace(/[^a-zA-Z0-9 _-]/g, "").trim() || "deck"
}

export async function downloadDeck(deck: Deck) {
  const zip = new JSZip()
  zip.file("deck/deck.json", JSON.stringify(deck.manifest, null, 2) + "\n")
  for (const id of deck.manifest.order)
    zip.file(`deck/slides/${id}.html`, deck.slides[id] + "\n")
  saveBlob(
    await zip.generateAsync({ type: "blob" }),
    `${fileName(deck.manifest.title)}.zip`
  )
}

const PX_PER_INCH = 144
const inches = (pixels: number) => pixels / PX_PER_INCH
const points = (pixels: string) => parseFloat(pixels) / 2

function color(value: string): string {
  const channels = value.match(/[\d.]+/g)
  if (!channels || channels.length < 3) return "000000"
  return channels
    .slice(0, 3)
    .map((channel) => Math.round(Number(channel)).toString(16).padStart(2, "0"))
    .join("")
}
function alpha(value: string) {
  const channels = value.match(/[\d.]+/g)
  return channels && channels.length > 3 ? Number(channels[3]) : 1
}

function textOptions(element: Element): pptxgen.TextPropsOptions {
  const style = getComputedStyle(element)
  return {
    fontFace: style.fontFamily.split(",")[0].replace(/["']/g, ""),
    fontSize: points(style.fontSize),
    bold: Number(style.fontWeight) >= 600 || style.fontWeight === "bold",
    italic: style.fontStyle === "italic",
    underline: style.textDecorationLine.includes("underline")
      ? { style: "sng" }
      : undefined,
    color: color(style.color),
    breakLine: false,
    ...(element instanceof HTMLAnchorElement && element.href
      ? { hyperlink: { url: element.href } }
      : {}),
  }
}

function textRuns(element: Element): pptxgen.TextProps[] {
  const runs: pptxgen.TextProps[] = []
  function walk(node: Node) {
    if (node.nodeType === Node.TEXT_NODE) {
      if (node.textContent)
        runs.push({
          text: node.textContent,
          options: textOptions(node.parentElement!),
        })
      return
    }
    if (!(node instanceof Element)) return
    const block =
      node !== element &&
      ["block", "list-item"].includes(getComputedStyle(node).display)
    if (block && runs.length)
      runs.push({ text: "\n", options: textOptions(node) })
    for (const child of node.childNodes) walk(child)
  }
  walk(element)
  return runs
}

export async function exportPowerPoint(deck: Deck) {
  const pptx = new pptxgen()
  pptx.defineLayout({
    name: "HTML_SLIDES",
    width: inches(SLIDE_WIDTH),
    height: inches(SLIDE_HEIGHT),
  })
  pptx.layout = "HTML_SLIDES"
  pptx.title = deck.manifest.title
  pptx.author = "Cool Computers"
  pptx.subject = "HTML slides"
  const host = document.createElement("div")
  host.className = "slides-content"
  host.inert = true
  host.setAttribute("aria-hidden", "true")
  Object.assign(host.style, {
    position: "fixed",
    left: "-20000px",
    top: "0px",
    width: `${SLIDE_WIDTH}px`,
    pointerEvents: "none",
  })
  document.body.append(host)
  try {
    await document.fonts.ready
    for (const id of deck.manifest.order) {
      host.innerHTML = deck.slides[id]
      const root = host.firstElementChild as HTMLElement
      root.removeAttribute("id")
      for (const element of root.querySelectorAll("[id]"))
        element.removeAttribute("id")
      await Promise.all(
        [...root.querySelectorAll("img")].map((image) => image.decode())
      )
      const slide = pptx.addSlide()
      slide.background = {
        color: color(getComputedStyle(root).backgroundColor),
      }
      const notes = root.querySelector(":scope > aside")?.textContent
      if (notes) slide.addNotes(notes)
      const origin = root.getBoundingClientRect()
      function bounds(element: Element) {
        const rect = element.getBoundingClientRect()
        return {
          x: inches(rect.x - origin.x),
          y: inches(rect.y - origin.y),
          w: inches(rect.width),
          h: inches(rect.height),
        }
      }
      async function raster(element: Element) {
        const box = bounds(element)
        // A wrapper gives SVG roots the HTMLElement surface expected by html-to-image.
        const wrapper = document.createElement("div")
        const clone = element.cloneNode(true) as HTMLElement
        const computed = getComputedStyle(element)
        for (let index = 0; index < computed.length; index++) {
          const property = computed[index]
          clone.style.setProperty(property, computed.getPropertyValue(property))
        }
        Object.assign(clone.style, {
          position: "absolute",
          left: "0px",
          top: "0px",
          margin: "0px",
        })
        Object.assign(wrapper.style, {
          position: "relative",
          width: `${box.w * PX_PER_INCH}px`,
          height: `${box.h * PX_PER_INCH}px`,
        })
        wrapper.append(clone)
        host.append(wrapper)
        const outer = wrapper.getBoundingClientRect()
        const inner = clone.getBoundingClientRect()
        clone.style.left = `${outer.left - inner.left}px`
        clone.style.top = `${outer.top - inner.top}px`
        try {
          const data = await toPng(wrapper, {
            pixelRatio: 1,
            skipFonts: true,
            cacheBust: false,
          })
          slide.addImage({ data, ...box })
        } finally {
          wrapper.remove()
        }
      }
      async function visit(element: Element) {
        const style = getComputedStyle(element)
        if (element.localName === "aside" || style.display === "none") return
        const box = bounds(element)
        if (element.localName === "hr") {
          slide.addShape(pptx.ShapeType.line, {
            ...box,
            h: 0,
            line: {
              color: color(style.borderTopColor),
              width: Math.max(0.5, points(style.borderTopWidth)),
            },
          })
          return
        }
        if (box.w <= 0 || box.h <= 0) return
        const complex =
          style.transform !== "none" ||
          style.boxShadow !== "none" ||
          Number(style.opacity) !== 1 ||
          style.backgroundImage !== "none" ||
          style.borderStyle === "dashed" ||
          style.borderStyle === "dotted"
        if (
          complex ||
          element.localName === "svg" ||
          element.localName === "ul" ||
          element.localName === "ol"
        ) {
          await raster(element)
          return
        }
        if (element instanceof HTMLImageElement) {
          slide.addImage({ data: element.src, ...box })
          return
        }
        if (element instanceof HTMLTableElement) {
          if (element.querySelector("[colspan], [rowspan]")) {
            await raster(element)
            return
          }
          const rows: pptxgen.TableRow[] = [...element.rows].map((row) =>
            [...row.cells].map((cell) => {
              const cellStyle = getComputedStyle(cell)
              const background = alpha(cellStyle.backgroundColor)
                ? cellStyle.backgroundColor
                : getComputedStyle(row).backgroundColor
              return {
                text: cell.innerText,
                options: {
                  ...textOptions(cell),
                  fill: {
                    color: color(background),
                    transparency: (1 - alpha(background)) * 100,
                  },
                  margin: [
                    points(cellStyle.paddingTop),
                    points(cellStyle.paddingRight),
                    points(cellStyle.paddingBottom),
                    points(cellStyle.paddingLeft),
                  ],
                  align:
                    cellStyle.textAlign === "center"
                      ? ("center" as const)
                      : cellStyle.textAlign === "right"
                        ? ("right" as const)
                        : ("left" as const),
                  border: {
                    type: "solid" as const,
                    pt: points(cellStyle.borderTopWidth),
                    color: color(cellStyle.borderTopColor),
                  },
                },
              }
            })
          )
          const first = element.rows[0]
          slide.addTable(rows, {
            ...box,
            colW: first
              ? [...first.cells].map((cell) => bounds(cell).w)
              : undefined,
            rowH: [...element.rows].map((row) => bounds(row).h),
            autoPage: false,
            margin: 0,
          })
          return
        }
        const backgroundAlpha = alpha(style.backgroundColor)
        const border = parseFloat(style.borderTopWidth)
        if (backgroundAlpha > 0 || border > 0 || element.localName === "hr") {
          slide.addShape(
            parseFloat(style.borderRadius)
              ? pptx.ShapeType.roundRect
              : pptx.ShapeType.rect,
            {
              ...box,
              fill: {
                color: color(style.backgroundColor),
                transparency: (1 - backgroundAlpha) * 100,
              },
              line: {
                color: color(style.borderTopColor),
                width: points(style.borderTopWidth),
                transparency: border ? 0 : 100,
              },
            }
          )
        }
        const hasBlocks = [...element.children].some(
          (child) =>
            !["inline", "contents"].includes(getComputedStyle(child).display) &&
            child.localName !== "span"
        )
        if (
          element.textContent?.trim() &&
          !hasBlocks &&
          !element.querySelector("img,svg,table")
        ) {
          const padding = [
            style.paddingTop,
            style.paddingRight,
            style.paddingBottom,
            style.paddingLeft,
          ].map(points) as [number, number, number, number]
          slide.addText(textRuns(element), {
            ...box,
            ...textOptions(element),
            margin: padding,
            align:
              style.textAlign === "center"
                ? "center"
                : style.textAlign === "right"
                  ? "right"
                  : "left",
            valign: "top",
            breakLine: false,
            paraSpaceAfter: 0,
          })
        } else {
          // A mixed container cannot omit direct text when walking its child objects.
          if (
            [...element.childNodes].some(
              (node) => node.nodeType === 3 && node.textContent?.trim()
            )
          ) {
            await raster(element)
            return
          }
          for (const child of element.children) await visit(child)
        }
      }
      for (const element of root.children) await visit(element)
    }
    const data = await pptx.write({ outputType: "blob" })
    saveBlob(data as Blob, `${fileName(deck.manifest.title)}.pptx`)
  } finally {
    host.remove()
  }
}
