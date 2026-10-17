export type Box = { x: number; y: number; width: number; height: number }
export type Guide = { axis: "x" | "y"; value: number }

export function elementBox(
  element: Element,
  canvas: Element,
  scale: number
): Box {
  const rect = element.getBoundingClientRect()
  const origin = canvas.getBoundingClientRect()
  return {
    x: (rect.x - origin.x) / scale,
    y: (rect.y - origin.y) / scale,
    width: rect.width / scale,
    height: rect.height / scale,
  }
}

export function snapBox(
  box: Box,
  targets: Box[],
  threshold: number
): { box: Box; guides: Guide[] } {
  const result = { ...box }
  const guides: Guide[] = []
  for (const axis of ["x", "y"] as const) {
    const size = axis === "x" ? "width" : "height"
    let distance = threshold
    let correction = 0
    let guide: number | undefined
    for (const target of targets) {
      for (const point of [
        target[axis],
        target[axis] + target[size] / 2,
        target[axis] + target[size],
      ]) {
        for (const edge of [
          box[axis],
          box[axis] + box[size] / 2,
          box[axis] + box[size],
        ]) {
          if (Math.abs(point - edge) < distance) {
            distance = Math.abs(point - edge)
            correction = point - edge
            guide = point
          }
        }
      }
    }
    result[axis] += correction
    if (guide !== undefined) guides.push({ axis, value: guide })
  }
  return { box: result, guides }
}

export function resizeBox(
  start: Box,
  handle: string,
  dx: number,
  dy: number,
  keepRatio: boolean
): Box {
  let width = Math.max(
    16,
    start.width + (handle.includes("w") ? -dx : handle.includes("e") ? dx : 0)
  )
  let height = Math.max(
    16,
    start.height + (handle.includes("n") ? -dy : handle.includes("s") ? dy : 0)
  )
  if (keepRatio) {
    const ratio = start.width / start.height
    if (
      handle === "n" ||
      handle === "s" ||
      Math.abs(dy / start.height) > Math.abs(dx / start.width)
    )
      width = height * ratio
    else height = width / ratio
  }
  return {
    x: start.x + (handle.includes("w") ? start.width - width : 0),
    y: start.y + (handle.includes("n") ? start.height - height : 0),
    width,
    height,
  }
}

export function copyTypography(
  source: Element,
  target: HTMLElement | SVGElement
) {
  const computed = getComputedStyle(source)
  for (const property of [
    "color",
    "font-family",
    "font-size",
    "font-weight",
    "font-style",
    "line-height",
    "letter-spacing",
    "text-align",
    "text-transform",
    "white-space",
    "word-break",
    "overflow-wrap",
  ]) {
    target.style.setProperty(property, computed.getPropertyValue(property))
  }
}

export function pinFlowSiblings(
  element: HTMLElement | SVGElement,
  canvas: HTMLElement,
  scale: number
) {
  if (getComputedStyle(element).position === "absolute") return
  const parents: Element[] = []
  for (
    let parent = element.parentElement;
    parent;
    parent = parent.parentElement
  ) {
    parents.unshift(parent)
    if (parent === canvas) break
  }
  for (const parent of parents) {
    // Measure and capture styles before removing anything from this flow.
    const pins = [...parent.children]
      .filter(
        (child): child is HTMLElement | SVGElement =>
          (child instanceof HTMLElement || child instanceof SVGElement) &&
          getComputedStyle(child).display !== "none" &&
          getComputedStyle(child).position !== "absolute"
      )
      .map((child) =>
        preparePin(child, canvas, elementBox(child, canvas, scale))
      )
    for (const pin of pins) pin()
  }
}

export function pinElement(
  element: HTMLElement | SVGElement,
  canvas: HTMLElement,
  box: Box
) {
  preparePin(element, canvas, box)()
}

function preparePin(
  element: HTMLElement | SVGElement,
  canvas: HTMLElement,
  box: Box
) {
  // Compose ancestor transforms before moving out of flow; the measured box
  // already accounts for translation, so only retain the linear transform.
  let matrix = new DOMMatrix()
  for (
    let node: Element | null = element;
    node && node !== canvas;
    node = node.parentElement
  ) {
    const transform = getComputedStyle(node).transform
    if (transform !== "none") matrix = new DOMMatrix(transform).multiply(matrix)
  }
  const width =
    element instanceof HTMLElement
      ? element.offsetWidth
      : parseFloat(getComputedStyle(element).width)
  const height =
    element instanceof HTMLElement
      ? element.offsetHeight
      : parseFloat(getComputedStyle(element).height)
  copyTypography(element, element)

  return () => {
    if (element.parentElement !== canvas)
      canvas.insertBefore(element, canvas.querySelector(":scope > aside"))
    Object.assign(element.style, {
      position: "absolute",
      margin: "0px",
      transform:
        matrix.a === 1 && matrix.b === 0 && matrix.c === 0 && matrix.d === 1
          ? "none"
          : `matrix(${matrix.a}, ${matrix.b}, ${matrix.c}, ${matrix.d}, 0, 0)`,
      transformOrigin: "0px 0px",
      width: `${width}px`,
      height: `${height}px`,
      right: "auto",
      bottom: "auto",
      boxSizing: "border-box",
      minWidth: "0px",
      minHeight: "0px",
      maxWidth: "none",
      maxHeight: "none",
    })
    applyBox(element, box)
  }
}

export function applyBox(element: HTMLElement | SVGElement, box: Box) {
  if (element.style.transform && element.style.transform !== "none") {
    const matrix = new DOMMatrix(element.style.transform)
    const width = parseFloat(element.style.width)
    const height = parseFloat(element.style.height)
    const scaleX =
      box.width / (Math.abs(matrix.a) * width + Math.abs(matrix.c) * height)
    const scaleY =
      box.height / (Math.abs(matrix.b) * width + Math.abs(matrix.d) * height)
    matrix.a *= scaleX
    matrix.c *= scaleX
    matrix.b *= scaleY
    matrix.d *= scaleY
    const offsetX = Math.min(
      0,
      matrix.a * width,
      matrix.c * height,
      matrix.a * width + matrix.c * height
    )
    const offsetY = Math.min(
      0,
      matrix.b * width,
      matrix.d * height,
      matrix.b * width + matrix.d * height
    )
    Object.assign(element.style, {
      left: `${box.x - offsetX}px`,
      top: `${box.y - offsetY}px`,
      transform: matrix.toString(),
    })
  } else {
    Object.assign(element.style, {
      left: `${box.x}px`,
      top: `${box.y}px`,
      width: `${box.width}px`,
      height: `${box.height}px`,
    })
  }
}
