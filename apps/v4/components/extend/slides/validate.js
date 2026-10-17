import { chartAttribute } from "./charts.js"

/** The deck contract lives here; callers supply a DOM (a browser or CLI adapter). */
export const SLIDE_WIDTH = 1920
export const SLIDE_HEIGHT = 1080
const ID = /^[A-Za-z0-9_-]{1,64}$/
export function validateManifest(value) {
  if (!value || typeof value !== "object") throw new Error("Invalid deck.json")
  const { title, order, sections } = value
  if (typeof title !== "string" || !title.trim())
    throw new Error("Missing deck title")
  if (
    !Array.isArray(order) ||
    !order.length ||
    order.length > 500 ||
    order.some((id) => typeof id !== "string" || !ID.test(id)) ||
    new Set(order).size !== order.length
  )
    throw new Error("Invalid slide order or IDs")
  if (
    sections !== undefined &&
    (!Array.isArray(sections) ||
      sections.some(
        (section) =>
          !section ||
          typeof section.title !== "string" ||
          !Array.isArray(section.slides) ||
          section.slides.some((id) => !order.includes(id))
      ))
  )
    throw new Error("Invalid deck sections")
  return { title, order, ...(sections ? { sections } : {}) }
}
const HTML_TAGS = new Set(
  "h1 h2 h3 p ul ol li b i u a span div figure img table thead tbody tfoot tr th td hr".split(
    " "
  )
)
const SVG_TAGS = new Set(
  "svg g path rect circle ellipse line polyline polygon text tspan title desc".split(
    " "
  )
)
const SVG_ATTRS = new Set(
  "viewBox x y x1 y1 x2 y2 cx cy r rx ry width height d points fill stroke stroke-width stroke-linecap stroke-linejoin opacity fill-opacity stroke-opacity transform text-anchor dominant-baseline font-size font-family font-weight".split(
    " "
  )
)
const CSS_PROPERTIES = new Set(
  (
    "width height min-width min-height max-width max-height margin margin-top margin-right margin-bottom margin-left " +
    "padding padding-top padding-right padding-bottom padding-left gap row-gap column-gap " +
    "color background background-color font-family font-size font-weight font-style line-height letter-spacing text-align text-decoration text-transform white-space white-space-collapse text-wrap-mode word-break overflow-wrap " +
    "display flex flex-direction flex-wrap flex-grow flex-shrink flex-basis align-items align-self align-content justify-content justify-items justify-self order " +
    "grid-template-columns grid-template-rows grid-column grid-row grid-auto-flow grid-auto-columns grid-auto-rows " +
    "position top right bottom left z-index border border-top border-right border-bottom border-left border-width border-style border-color border-radius " +
    "box-shadow transform transform-origin opacity overflow object-fit object-position border-collapse border-spacing table-layout list-style-type box-sizing"
  ).split(" ")
)
function safeValue(value, css = false) {
  const lengths = value
    .replace(/(?:rgba?|hsla?)\([^)]*\)/gi, "")
    .replace(/(["']).*?\1/g, "")
  const units = [
    ...lengths.matchAll(/(?:^|[\s,(])[-+]?(?:\d*\.)?\d+([a-z%]+)/gi),
  ]
  return (
    !/[\\@!<>]|url\s*\(|var\s*\(|expression\s*\(/i.test(value) &&
    (!css ||
      units.every((match) => /^(px|fr|deg|rad|turn)$/i.test(match[1]))) &&
    [...value.matchAll(/([\w-]+)\s*\(/g)].every((match) =>
      /^(rgb|rgba|hsl|hsla|translate|translateX|translateY|scale|scaleX|scaleY|rotate|skew|skewX|skewY|matrix|repeat|minmax)$/i.test(
        match[1]
      )
    )
  )
}
function safeImage(value) {
  return /^data:image\/(png|jpeg|gif|webp);base64,[a-zA-Z0-9+/=]+$/.test(value)
}
export function sanitizeSlide(source, slideId, dom, warn = console.warn) {
  if (!ID.test(slideId)) throw new Error("Invalid slide ID")
  const template = dom.createElement("template")
  template.innerHTML = source
  const roots = [...template.content.children]
  const root = roots.find((element) => element.localName === "section")
  if (!root)
    throw new Error(`Slide ${slideId}: expected one <section id="${slideId}">`)
  const drop = (element, detail) =>
    warn(
      `Slide ${slideId}: <${element.localName || "text"}${element.id ? ` id="${element.id}"` : ""}> dropped ${detail}`,
      {
        element: element.id || element.localName || "text",
        text: (element.textContent || element.getAttribute?.("alt") || "")
          .replace(/\s+/g, " ")
          .trim()
          .slice(0, 80),
        detail,
      }
    )
  for (const node of [...template.content.childNodes]) {
    if (node !== root && (node.nodeType !== 3 || node.textContent?.trim())) {
      drop(node, "outside section")
    }
  }
  if (root.id !== slideId) drop(root, "mismatched section ID")
  const seen = new Set([slideId])
  let nextId = 1
  const children = [...root.children]
  function clean(element, inSvg) {
    const tag = element.localName
    const isRoot = element === root
    const isNotes =
      tag === "aside" &&
      element.parentElement === root &&
      element === children.at(-1)
    const svg = inSvg || tag === "svg"
    if (!isRoot && !isNotes && !(svg ? SVG_TAGS : HTML_TAGS).has(tag)) {
      drop(element, "unsupported element")
      element.remove()
      return
    }
    for (const attribute of [...element.attributes]) {
      const { name, value } = attribute
      const allowed =
        name === "style" ||
        name === "id" ||
        name === "title" ||
        (tag === "figure" && chartAttribute(name, value)) ||
        (tag === "img" &&
          ((name === "src" && safeImage(value)) || name === "alt")) ||
        (tag === "a" &&
          name === "href" &&
          /^(https?:\/\/|mailto:)/i.test(value.trim())) ||
        ((tag === "td" || tag === "th") &&
          /^(colspan|rowspan)$/.test(name) &&
          /^[1-9]\d?$/.test(value)) ||
        (svg &&
          SVG_ATTRS.has(name) &&
          safeValue(value) &&
          !/url|javascript|data:/i.test(value))
      if (!allowed) {
        drop(element, `attribute ${name}`)
        element.removeAttribute(name)
      }
    }
    if (tag === "img" && !element.hasAttribute("src")) {
      drop(element, "image without an embedded source")
      element.remove()
      return
    }
    const style = element.style
    const originalStyle = element.getAttribute("style") || ""
    const declaredProperties = new Set()
    for (const declaration of originalStyle.split(";")) {
      const separator = declaration.indexOf(":")
      if (separator < 0) continue
      const property = declaration
        .slice(0, separator)
        .replace(/\/\*.*?\*\//g, "")
        .trim()
        .toLowerCase()
      declaredProperties.add(property)
      const probe = dom.createElement("div").style
      probe.cssText = declaration
      if (!probe.getPropertyValue(property))
        drop(element, `invalid style ${property}`)
    }
    for (const property of [...style]) {
      const value = style.getPropertyValue(property)
      // Chromium expands an ordinary border into unsupported border-image resets.
      if (
        property.startsWith("border-image-") &&
        declaredProperties.has("border") &&
        ![...declaredProperties].some((name) => name.startsWith("border-image"))
      ) {
        style.removeProperty(property)
        continue
      }
      if (
        !(
          CSS_PROPERTIES.has(property) ||
          /^(grid-(column|row)-(start|end)|overflow-(x|y)|border-(top|right|bottom|left)-(width|style|color)|border-(top|bottom)-(left|right)-radius|background-(image|position-x|position-y|size|repeat|attachment|origin|clip)|text-decoration-(line|style|color|thickness))$/.test(
            property
          )
        ) ||
        !safeValue(value, true) ||
        style.getPropertyPriority(property) ||
        (property === "position" &&
          !/^(static|relative|absolute)$/.test(value)) ||
        (property === "display" &&
          !/^(block|inline|inline-block|flex|inline-flex|grid|none|table|table-row|table-cell|table-row-group|table-header-group|table-footer-group|list-item)$/.test(
            value
          ))
      ) {
        drop(element, `style ${property}`)
        style.removeProperty(property)
      }
    }
    if (style.cssText) element.setAttribute("style", style.cssText)
    else element.removeAttribute("style")
    if (!isRoot) {
      if (
        !/^[A-Za-z0-9_-]{1,160}$/.test(element.id) ||
        seen.has(element.id) ||
        !element.id.startsWith(`${slideId}-`)
      ) {
        if (element.id) drop(element, "invalid or duplicate element ID")
        do {
          element.id = `${slideId}-e${nextId++}`
        } while (seen.has(element.id))
      }
      seen.add(element.id)
    }
    if (tag === "figure" && element.hasAttribute("data-chart")) {
      for (const child of [...element.children]) {
        if (child.localName !== "table") child.remove()
      }
    }
    for (const child of [...element.children]) clean(child, svg)
    for (const node of [...element.childNodes])
      if (node.nodeType === 8) node.remove()
  }
  clean(root, false)
  root.id = slideId
  const style = root.style
  const canvasStyles = {
    width: `${SLIDE_WIDTH}px`,
    height: `${SLIDE_HEIGHT}px`,
    position: "relative",
    overflow: "hidden",
    "box-sizing": "border-box",
    margin: "0px",
    top: "0px",
    left: "0px",
    right: "auto",
    bottom: "auto",
    transform: "none",
    "max-width": "none",
    "max-height": "none",
    "min-width": "0px",
    "min-height": "0px",
  }
  for (const [property, value] of Object.entries(canvasStyles)) {
    if (
      style.getPropertyValue(property) &&
      style.getPropertyValue(property) !== value
    )
      drop(root, `canvas style ${property}`)
    style.setProperty(property, value)
  }
  if (style.display === "none") {
    drop(root, "hidden canvas")
    style.display = "block"
  }
  return root.outerHTML
}
export async function loadDeck(src, signal) {
  const response = await fetch(src, { signal, cache: "no-store" })
  if (!response.ok) throw new Error(`Deck metadata: HTTP ${response.status}`)
  const manifest = validateManifest(await response.json())
  const base = new URL(".", new URL(src, window.location.href))
  const slides = Object.fromEntries(
    await Promise.all(
      manifest.order.map(async (id) => {
        const response = await fetch(new URL(`slides/${id}.html`, base), {
          signal,
          cache: "no-store",
        })
        if (!response.ok)
          throw new Error(`Slide ${id}: HTTP ${response.status}`)
        return [id, sanitizeSlide(await response.text(), id, document)]
      })
    )
  )
  return { manifest, slides }
}
