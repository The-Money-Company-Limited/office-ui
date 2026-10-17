import * as React from "react"

import {
  applyBox,
  copyTypography,
  elementBox,
  pinElement,
  resizeBox,
  snapBox,
  type Box,
  type Guide,
} from "./geometry"
import { sanitizeSlide, SLIDE_HEIGHT, SLIDE_WIDTH } from "./validate.js"

const HANDLES = ["nw", "n", "ne", "e", "se", "s", "sw", "w"]
type SlideElement = HTMLElement | SVGElement

type Gesture = {
  pointer: number
  x: number
  y: number
  box: Box
  element: SlideElement
  handle: string | null
  before: string
  moved: boolean
  targets: Box[]
}

export function SlideCanvas({
  id,
  html,
  scale,
  editable,
  onChange,
  onUndo,
  onRedo,
}: {
  id: string
  html: string
  scale: number
  editable: boolean
  onChange: (html: string) => void
  onUndo: () => void
  onRedo: () => void
}) {
  const host = React.useRef<HTMLDivElement>(null)
  const content = React.useRef<HTMLDivElement>(null)
  const selected = React.useRef<SlideElement | null>(null)
  const gesture = React.useRef<Gesture | null>(null)
  const editing = React.useRef<HTMLElement | null>(null)
  const [box, setBox] = React.useState<Box | null>(null)
  const [guides, setGuides] = React.useState<Guide[]>([])

  React.useLayoutEffect(() => {
    const container = content.current!
    const selectedId = selected.current?.id
    container.innerHTML = html
    selected.current = selectedId
      ? container.querySelector(`#${selectedId}`)
      : null
    editing.current = null
    if (!editable) selected.current = null
    setBox(
      selected.current
        ? elementBox(selected.current, container.firstElementChild!, scale)
        : null
    )
  }, [html, editable, scale])

  function canvas() {
    return content.current!.firstElementChild as HTMLElement
  }
  function select(element: SlideElement | null) {
    selected.current = element
    setBox(element ? elementBox(element, canvas(), scale) : null)
  }
  function commit() {
    const clean = sanitizeSlide(canvas().outerHTML, id, document)
    if (clean !== html) onChange(clean)
  }
  function endText() {
    const element = editing.current
    if (!element) return
    editing.current = null
    element.removeAttribute("contenteditable")
    if (element.textContent?.includes("\n"))
      element.style.whiteSpace = "pre-wrap"
    commit()
    select(element)
  }
  function targetElement(target: EventTarget | null): SlideElement | null {
    if (!(target instanceof Element)) return null
    const element = target.closest("svg, table") || target
    if (
      element === canvas() ||
      !canvas().contains(element) ||
      element.closest("aside")
    )
      return null
    return element as SlideElement
  }
  function pointerDown(event: React.PointerEvent) {
    if (!editable || event.button !== 0 || editing.current) return
    const handle =
      (event.target as Element).closest<HTMLElement>("[data-resize]")?.dataset
        .resize || null
    const element = handle ? selected.current : targetElement(event.target)
    select(element)
    host.current?.focus({ preventScroll: true })
    if (!element) return
    event.preventDefault()
    const start = elementBox(element, canvas(), scale)
    gesture.current = {
      pointer: event.pointerId,
      x: event.clientX,
      y: event.clientY,
      box: start,
      element,
      handle,
      before: canvas().outerHTML,
      moved: false,
      targets: [
        { x: 0, y: 0, width: SLIDE_WIDTH, height: SLIDE_HEIGHT },
        ...[...canvas().querySelectorAll("[id]")]
          .filter(
            (other) =>
              other !== element &&
              !other.contains(element) &&
              !element.contains(other) &&
              !other.closest("aside")
          )
          .map((other) => elementBox(other, canvas(), scale)),
      ],
    }
  }
  function pointerMove(event: React.PointerEvent) {
    const drag = gesture.current
    if (!drag) return
    const dx = (event.clientX - drag.x) / scale
    const dy = (event.clientY - drag.y) / scale
    if (!drag.moved && Math.hypot(dx, dy) < 3 / scale) return
    if (!drag.moved) {
      host.current?.setPointerCapture(event.pointerId)
      pinElement(drag.element, canvas(), drag.box)
    }
    drag.moved = true
    let next: Box
    if (drag.handle) {
      next = resizeBox(drag.box, drag.handle, dx, dy, event.shiftKey)
      setGuides([])
    } else {
      const snapped = snapBox(
        { ...drag.box, x: drag.box.x + dx, y: drag.box.y + dy },
        drag.targets,
        6 / scale
      )
      next = snapped.box
      setGuides(snapped.guides)
    }
    applyBox(drag.element, next)
    setBox(next)
  }
  function pointerEnd(event: React.PointerEvent, cancel = false) {
    const drag = gesture.current
    if (!drag) return
    gesture.current = null
    if (host.current?.hasPointerCapture(event.pointerId))
      host.current.releasePointerCapture(event.pointerId)
    setGuides([])
    if (cancel) {
      content.current!.innerHTML = drag.before
      select(null)
    } else if (drag.moved) commit()
  }
  function keyDown(event: React.KeyboardEvent) {
    if (!editable) return
    if (editing.current) {
      if (event.key === "Escape") {
        endText()
        host.current?.focus()
      }
      return
    }
    const command = event.metaKey || event.ctrlKey
    const key = event.key.toLowerCase()
    if (command && key === "z") {
      event.preventDefault()
      if (event.shiftKey) onRedo()
      else onUndo()
      return
    }
    const element = selected.current
    if (event.key === "Escape") {
      select(null)
      return
    }
    if (!element) return
    if (event.key === "Enter") {
      event.preventDefault()
      editText(element)
      return
    }
    if (command && key === "d") {
      event.preventDefault()
      const clone = element.cloneNode(true) as SlideElement
      clone.id = `${id}-copy-${crypto.randomUUID().slice(0, 8)}`
      for (const child of clone.querySelectorAll("[id]"))
        child.removeAttribute("id")
      const oldBox = elementBox(element, canvas(), scale)
      copyTypography(element, clone)
      canvas().insertBefore(clone, canvas().querySelector(":scope > aside"))
      pinElement(clone, canvas(), {
        ...oldBox,
        x: oldBox.x + 24,
        y: oldBox.y + 24,
      })
      selected.current = clone
      commit()
    } else if (event.key === "Delete" || event.key === "Backspace") {
      event.preventDefault()
      element.remove()
      select(null)
      commit()
    } else if (
      ["ArrowUp", "ArrowDown", "ArrowLeft", "ArrowRight"].includes(event.key)
    ) {
      event.preventDefault()
      const current = elementBox(element, canvas(), scale)
      const amount = event.shiftKey ? 10 : 1
      pinElement(element, canvas(), {
        ...current,
        x:
          current.x +
          (event.key === "ArrowRight"
            ? amount
            : event.key === "ArrowLeft"
              ? -amount
              : 0),
        y:
          current.y +
          (event.key === "ArrowDown"
            ? amount
            : event.key === "ArrowUp"
              ? -amount
              : 0),
      })
      commit()
    }
  }
  function editText(element: SlideElement | null) {
    if (
      !(element instanceof HTMLElement) ||
      !/^(H[1-3]|P|LI|SPAN|B|I|U|A|TD|TH|DIV)$/.test(element.tagName) ||
      element.querySelector("div,table,img,svg,ul,ol,h1,h2,h3,p")
    )
      return
    editing.current = element
    element.setAttribute("contenteditable", "plaintext-only")
    element.focus()
    setBox(null)
    const range = document.createRange()
    range.selectNodeContents(element)
    const selection = window.getSelection()
    selection?.removeAllRanges()
    selection?.addRange(range)
  }

  return (
    <div
      ref={host}
      className="slides-stage"
      tabIndex={editable ? 0 : -1}
      role="group"
      aria-label={`Slide ${id} canvas`}
      style={{ width: SLIDE_WIDTH * scale, height: SLIDE_HEIGHT * scale }}
      onPointerDown={pointerDown}
      onPointerMove={pointerMove}
      onPointerUp={pointerEnd}
      onPointerCancel={(event) => pointerEnd(event, true)}
      onKeyDown={keyDown}
      onDoubleClick={(event) => {
        if (editable)
          editText(
            (event.target as Element).closest<HTMLElement>("td, th") ||
              targetElement(event.target)
          )
      }}
      onBlur={(event) => {
        if (event.target === editing.current) endText()
      }}
      onClick={(event) => {
        if (editable && (event.target as Element).closest("a"))
          event.preventDefault()
      }}
    >
      <div className="slides-scale" style={{ transform: `scale(${scale})` }}>
        <div ref={content} className="slides-content" data-slide-id={id} />
        {editable && box && (
          <div
            className="slides-selection"
            style={{
              left: box.x,
              top: box.y,
              width: box.width,
              height: box.height,
              borderWidth: 1.5 / scale,
            }}
          >
            {HANDLES.map((handle) => (
              <span
                key={handle}
                data-resize={handle}
                className={`slides-handle slides-handle-${handle}`}
                style={{
                  width: 8 / scale,
                  height: 8 / scale,
                  borderWidth: 1 / scale,
                }}
              />
            ))}
          </div>
        )}
        {guides.map((guide) => (
          <div
            key={guide.axis}
            className={`slides-guide slides-guide-${guide.axis}`}
            style={
              guide.axis === "x"
                ? { left: guide.value, width: 1 / scale }
                : { top: guide.value, height: 1 / scale }
            }
          />
        ))}
      </div>
    </div>
  )
}
