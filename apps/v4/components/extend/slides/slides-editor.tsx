import * as React from "react"

import { SlideCanvas } from "./slide-canvas"
import { loadDeck, SLIDE_HEIGHT, SLIDE_WIDTH, type Deck } from "./validate.js"

import "./slides.css"

type History = { past: Deck[]; current: Deck; future: Deck[] }

export function SlidesEditor({
  src,
  fileName = "deck.json",
  isDark,
  onIsDarkChange,
}: {
  src: string
  fileName?: string
  isDark: boolean
  onIsDarkChange: (dark: boolean) => void
}) {
  const [history, setHistory] = React.useState<History | null>(null)
  const [active, setActive] = React.useState(0)
  const [presenting, setPresenting] = React.useState(false)
  const [size, setSize] = React.useState({ width: 1200, height: 800 })
  const [error, setError] = React.useState<string | null>(null)
  const [busy, setBusy] = React.useState(false)
  const surface = React.useRef<HTMLDivElement>(null)
  const presentation = React.useRef(false)
  const deck = history?.current

  const exitPresentation = React.useCallback(() => {
    presentation.current = false
    setPresenting(false)
    if (document.fullscreenElement) void document.exitFullscreen()
    requestAnimationFrame(() => {
      if (deck)
        document.getElementById(deck.manifest.order[active])?.scrollIntoView()
    })
  }, [active, deck])

  React.useEffect(() => {
    const controller = new AbortController()
    loadDeck(src, controller.signal)
      .then((current) => {
        setError(null)
        setHistory({ past: [], current, future: [] })
        const index = current.manifest.order.indexOf(
          window.location.hash.slice(1)
        )
        setActive(Math.max(0, index))
        if (index >= 0)
          requestAnimationFrame(() =>
            document
              .getElementById(current.manifest.order[index])
              ?.scrollIntoView()
          )
      })
      .catch((reason: unknown) => {
        if (!controller.signal.aborted)
          setError(
            reason instanceof Error ? reason.message : "Could not open deck"
          )
      })
    return () => controller.abort()
  }, [src])

  React.useLayoutEffect(() => {
    function resize() {
      setSize({ width: window.innerWidth, height: window.innerHeight })
    }
    resize()
    window.addEventListener("resize", resize)
    return () => window.removeEventListener("resize", resize)
  }, [])

  React.useEffect(() => {
    if (!deck || presenting) return
    const observer = new IntersectionObserver(
      () => {
        if (presentation.current) return
        const pages = [
          ...(surface.current?.querySelectorAll<HTMLElement>(
            "[data-slide-index]"
          ) || []),
        ]
        const visible = pages
          .map((page) => {
            const box = page.getBoundingClientRect()
            return {
              index: Number(page.dataset.slideIndex),
              height: Math.max(
                0,
                Math.min(box.bottom, window.innerHeight) -
                  Math.max(box.top, 112)
              ),
            }
          })
          .sort((a, b) => b.height - a.height)[0]
        if (visible && visible.height > 0) setActive(visible.index)
      },
      { rootMargin: "-110px 0px 0px", threshold: [0, 0.25, 0.5, 0.75, 1] }
    )
    surface.current
      ?.querySelectorAll("[data-slide-index]")
      .forEach((element) => observer.observe(element))
    function hashChange() {
      const index = deck!.manifest.order.indexOf(window.location.hash.slice(1))
      if (index >= 0) {
        setActive(index)
        document.getElementById(deck!.manifest.order[index])?.scrollIntoView()
      }
    }
    window.addEventListener("hashchange", hashChange)
    return () => {
      observer.disconnect()
      window.removeEventListener("hashchange", hashChange)
    }
  }, [deck, presenting])

  React.useEffect(() => {
    function fullscreenChange() {
      if (!document.fullscreenElement) {
        presentation.current = false
        setPresenting(false)
      }
    }
    document.addEventListener("fullscreenchange", fullscreenChange)
    return () =>
      document.removeEventListener("fullscreenchange", fullscreenChange)
  }, [])

  React.useEffect(() => {
    if (!presenting || !deck) return
    function keyDown(event: KeyboardEvent) {
      if (event.key === "Escape") {
        event.preventDefault()
        exitPresentation()
      } else if (
        ["ArrowRight", "ArrowDown", " ", "PageDown"].includes(event.key)
      ) {
        event.preventDefault()
        setActive((index) =>
          Math.min(index + 1, deck!.manifest.order.length - 1)
        )
      } else if (["ArrowLeft", "ArrowUp", "PageUp"].includes(event.key)) {
        event.preventDefault()
        setActive((index) => Math.max(0, index - 1))
      }
    }
    window.addEventListener("keydown", keyDown)
    return () => window.removeEventListener("keydown", keyDown)
  }, [presenting, deck, exitPresentation])

  function updateSlide(id: string, html: string) {
    setHistory(
      (value) =>
        value && {
          past: [...value.past, value.current].slice(-100),
          current: {
            ...value.current,
            slides: { ...value.current.slides, [id]: html },
          },
          future: [],
        }
    )
  }
  function undo() {
    setHistory((value) =>
      !value?.past.length
        ? value
        : {
            past: value.past.slice(0, -1),
            current: value.past.at(-1)!,
            future: [value.current, ...value.future],
          }
    )
  }
  function redo() {
    setHistory((value) =>
      !value?.future.length
        ? value
        : {
            past: [...value.past, value.current],
            current: value.future[0],
            future: value.future.slice(1),
          }
    )
  }
  function goTo(index: number) {
    setActive(index)
    const id = deck!.manifest.order[index]
    window.history.replaceState(null, "", `#${id}`)
    document
      .getElementById(id)
      ?.scrollIntoView({ behavior: "instant", block: "start" })
  }
  async function present() {
    presentation.current = true
    setPresenting(true)
    try {
      await surface.current?.requestFullscreen()
    } catch {
      /* Presentation mode also works where fullscreen is unavailable. */
    }
  }
  async function download(format: "zip" | "pptx") {
    if (!deck) return
    setBusy(true)
    setError(null)
    try {
      const exports = await import("./export")
      if (format === "zip") await exports.downloadDeck(deck)
      else await exports.exportPowerPoint(deck)
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : "Export failed")
    } finally {
      setBusy(false)
    }
  }

  const scale = Math.max(
    0.05,
    Math.min(
      (size.width - (presenting ? 32 : 208)) / SLIDE_WIDTH,
      (size.height - (presenting ? 32 : 160)) / SLIDE_HEIGHT,
      1
    )
  )

  return (
    <div
      ref={surface}
      onKeyDown={(event) => {
        if (
          presenting ||
          event.defaultPrevented ||
          (event.target as Element).closest("[contenteditable]")
        )
          return
        if (
          (event.metaKey || event.ctrlKey) &&
          event.key.toLowerCase() === "z"
        ) {
          event.preventDefault()
          if (event.shiftKey) redo()
          else undo()
        }
      }}
      className={
        presenting ? "artifact-slides slides-presenting" : "artifact-slides"
      }
    >
      {!presenting && (
        <header className="slides-toolbar">
          <div className="slides-header">
            <span className="slides-file-icon" aria-hidden="true">
              ▤
            </span>
            <strong title={fileName}>{deck?.manifest.title || fileName}</strong>
            <span className="slides-filename">{fileName}</span>
            <button
              type="button"
              onClick={() => onIsDarkChange(!isDark)}
              aria-label={
                isDark ? "Use light appearance" : "Use dark appearance"
              }
            >
              {isDark ? "☀" : "☾"}
            </button>
            <button
              type="button"
              disabled={!deck || busy}
              onClick={() => void download("zip")}
            >
              Download .zip
            </button>
            <button
              type="button"
              disabled={!deck || busy}
              onClick={() => void download("pptx")}
            >
              {busy ? "Exporting…" : "Export .pptx"}
            </button>
            <button
              type="button"
              disabled={!deck}
              onClick={() => void present()}
            >
              Present
            </button>
          </div>
          <div className="slides-ribbon">
            <button
              type="button"
              disabled={!history?.past.length}
              onClick={undo}
            >
              Undo
            </button>
            <button
              type="button"
              disabled={!history?.future.length}
              onClick={redo}
            >
              Redo
            </button>
            <span>
              Click to select · Double-click to edit · Drag handles to resize
            </span>
            <span className="slides-memory" role="status">
              Edits stay in this tab. Download to keep them.
            </span>
          </div>
        </header>
      )}
      {error && (
        <p className="slides-error" role="alert">
          {error}
        </p>
      )}
      {!deck && !error && <div className="artifact-state">Opening deck…</div>}
      {deck && (
        <>
          {!presenting && (
            <nav
              className="slides-strip"
              aria-label="Slides"
              data-document-thumbnail-sidebar
            >
              {deck.manifest.order.map((id, index) => (
                <button
                  type="button"
                  key={id}
                  className={
                    index === active
                      ? "slides-thumbnail slides-thumbnail-active"
                      : "slides-thumbnail"
                  }
                  aria-label={`Slide ${index + 1}`}
                  aria-current={index === active ? "page" : undefined}
                  onClick={() => goTo(index)}
                >
                  <span className="slides-thumbnail-frame">
                    <span
                      className="slides-thumbnail-content slides-content"
                      aria-hidden="true"
                      inert
                      dangerouslySetInnerHTML={{
                        __html: deck.slides[id].replace(/ id="[^"]*"/g, ""),
                      }}
                    />
                  </span>
                  <span>{index + 1}</span>
                </button>
              ))}
            </nav>
          )}
          <div className="slides-pages">
            {deck.manifest.order.map(
              (id, index) =>
                (!presenting || index === active) && (
                  <div
                    className="slides-page"
                    key={id}
                    data-slide-index={index}
                    data-slide-anchor={id}
                  >
                    <SlideCanvas
                      id={id}
                      html={deck.slides[id]}
                      scale={scale}
                      editable={!presenting}
                      onChange={(html) => updateSlide(id, html)}
                      onUndo={undo}
                      onRedo={redo}
                    />
                    {!presenting && (
                      <span className="slides-page-label">
                        {index + 1} / {deck.manifest.order.length}
                      </span>
                    )}
                  </div>
                )
            )}
          </div>
          {presenting && (
            <button
              type="button"
              className="slides-exit"
              onClick={exitPresentation}
            >
              Exit presentation
            </button>
          )}
        </>
      )}
    </div>
  )
}
