import * as React from "react"
import {
  SaveAccess,
  SaveStatus,
  useDocumentAdapter,
  useDocumentRevision,
  useSaveState,
} from "@/experiments/artifact-template/src/autosave-context"

import { Button } from "@/components/ui/button"
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu"
import { Separator } from "@/components/ui/separator"
import { IconPlaceholder } from "@/components/icon-placeholder"
import { AppearanceToggle } from "@/components/extend/appearance-toggle"

import { chartPanelLayout } from "./chart-layout"
import { renderCharts } from "./charts.js"
import type { Box } from "./geometry"
import { SlideCanvas } from "./slide-canvas"
import { loadDeck, SLIDE_HEIGHT, SLIDE_WIDTH, type Deck } from "./validate.js"

import "./slides.css"

type History = { past: Deck[]; current: Deck; future: Deck[] }

export function SlidesEditor({
  src,
  title,
  isDark,
  onIsDarkChange,
}: {
  src: string
  title: string
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
  const returnSlide = React.useRef<string | null>(null)
  const deck = history?.current
  const { canSave } = useSaveState()
  const savedDeck = React.useRef<Deck | null>(null)
  useDocumentRevision(deck, Boolean(deck))
  useDocumentAdapter({
    serialize: async () => {
      if (!deck) throw new Error("The deck is not ready.")
      const { serializeDeck } = await import("./export")
      return {
        files: serializeDeck(deck, savedDeck.current),
        saved: () => {
          savedDeck.current = deck
        },
      }
    },
    download: async () => {
      if (deck) await (await import("./export")).downloadDeck(deck)
    },
  })

  const [chartSelection, setChartSelection] = React.useState<{
    id: string
    box: Box
  } | null>(null)
  const onChartSelection = React.useCallback((id: string, box: Box | null) => {
    setChartSelection((current) => {
      if (!box) return current?.id === id ? null : current
      if (
        current?.id === id &&
        Object.keys(box).every(
          (key) => box[key as keyof Box] === current.box[key as keyof Box]
        )
      )
        return current
      return { id, box }
    })
  }, [])

  const exitPresentation = React.useCallback(async () => {
    if (!presentation.current) return
    presentation.current = false
    if (document.fullscreenElement) await document.exitFullscreen()
    returnSlide.current = deck?.manifest.order[active] || null
    setPresenting(false)
  }, [active, deck])

  React.useLayoutEffect(() => {
    if (!presenting) {
      const id = returnSlide.current
      if (id) {
        returnSlide.current = null
        window.history.replaceState(null, "", `#${id}`)
        document
          .getElementById(id)
          ?.scrollIntoView({ behavior: "instant", block: "start" })
      }
      return
    }
    const root = document.documentElement
    const body = document.body
    const rootOverflow = root.style.overflow
    const bodyOverflow = body.style.overflow
    root.style.overflow = "hidden"
    body.style.overflow = "hidden"
    return () => {
      root.style.overflow = rootOverflow
      body.style.overflow = bodyOverflow
    }
  }, [presenting])

  React.useEffect(() => {
    const controller = new AbortController()
    loadDeck(src, controller.signal)
      .then((current) => {
        setError(null)
        savedDeck.current = current
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
      if (!document.fullscreenElement) void exitPresentation()
    }
    document.addEventListener("fullscreenchange", fullscreenChange)
    return () =>
      document.removeEventListener("fullscreenchange", fullscreenChange)
  }, [exitPresentation])

  React.useEffect(() => {
    if (!presenting || !deck) return
    function keyDown(event: KeyboardEvent) {
      if (event.key === "Escape") {
        event.preventDefault()
        void exitPresentation()
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
    if (!canSave) return
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
    if (!canSave) return
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
    if (!canSave) return
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

  const naturalScale = Math.max(
    0.05,
    Math.min(
      (size.width - (presenting ? 32 : 208)) / SLIDE_WIDTH,
      (size.height - (presenting ? 32 : 160)) / SLIDE_HEIGHT,
      1
    )
  )

  const chartLayout =
    chartSelection && !presenting
      ? chartPanelLayout(chartSelection.box, naturalScale, size.width)
      : null
  const reserve = chartLayout?.sidebar ? chartLayout.width + 24 : 0
  const stripWidth = chartLayout?.compact ? 0 : 152
  const scale = reserve
    ? Math.max(
        0.05,
        Math.min(
          naturalScale,
          (size.width - stripWidth - 48 - reserve) / SLIDE_WIDTH
        )
      )
    : naturalScale

  return (
    <div
      ref={surface}
      data-chart-sidebar={chartLayout?.sidebar || undefined}
      data-chart-compact={chartLayout?.compact || undefined}
      style={
        {
          "--slides-chart-reserve": `${reserve}px`,
          "--slides-strip-width": `${stripWidth}px`,
        } as React.CSSProperties
      }
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
            <div className="flex min-w-0 flex-1 items-center gap-2">
              <span className="flex size-7 shrink-0 items-center justify-center rounded-md bg-primary/10 text-primary">
                <IconPlaceholder
                  lucide="FileText"
                  tabler="IconFileText"
                  hugeicons="File01Icon"
                  phosphor="FileDocIcon"
                  remixicon="RiFileTextLine"
                  className="size-4"
                />
              </span>
              <strong
                className="truncate text-sm font-semibold"
                title={deck?.manifest.title || title}
              >
                {deck?.manifest.title || title}
              </strong>
              <SaveStatus />
              <SaveAccess />
            </div>
            <AppearanceToggle isDark={isDark} onIsDarkChange={onIsDarkChange} />
            <Separator
              orientation="vertical"
              className="mx-1 h-4 self-center"
            />
            <DropdownMenu>
              <DropdownMenuTrigger
                disabled={!deck || busy}
                render={
                  <Button
                    variant="ghost"
                    size="icon-sm"
                    aria-label="Download"
                    title="Download"
                  />
                }
              >
                <IconPlaceholder
                  lucide="Download"
                  tabler="IconDownload"
                  hugeicons="Download01Icon"
                  phosphor="DownloadSimpleIcon"
                  remixicon="RiDownload2Line"
                  className="size-4"
                />
              </DropdownMenuTrigger>
              <DropdownMenuContent align="end">
                <DropdownMenuItem onClick={() => void download("zip")}>
                  Deck as .zip
                </DropdownMenuItem>
                <DropdownMenuItem onClick={() => void download("pptx")}>
                  PowerPoint .pptx
                </DropdownMenuItem>
              </DropdownMenuContent>
            </DropdownMenu>
            <Button
              className="slides-present"
              size="sm"
              disabled={!deck}
              onClick={() => void present()}
            >
              Present
            </Button>
          </div>
          {canSave && (
            <div className="slides-ribbon">
              <Button
                variant="ghost"
                size="icon-sm"
                aria-label="Undo"
                title="Undo"
                disabled={!history?.past.length}
                onClick={undo}
              >
                <IconPlaceholder
                  lucide="Undo2"
                  tabler="IconArrowBackUp"
                  hugeicons="Undo02Icon"
                  phosphor="ArrowUUpLeftIcon"
                  remixicon="RiArrowGoBackLine"
                  className="size-4"
                />
              </Button>
              <Button
                variant="ghost"
                size="icon-sm"
                aria-label="Redo"
                title="Redo"
                disabled={!history?.future.length}
                onClick={redo}
              >
                <IconPlaceholder
                  lucide="Redo2"
                  tabler="IconArrowForwardUp"
                  hugeicons="Redo02Icon"
                  phosphor="ArrowUUpRightIcon"
                  remixicon="RiArrowGoForwardLine"
                  className="size-4"
                />
              </Button>
            </div>
          )}
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
                    <SlideThumbnailContent html={deck.slides[id]} />
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
                      editable={canSave && !presenting}
                      onChange={(html) => updateSlide(id, html)}
                      onUndo={undo}
                      onRedo={redo}
                      chartOpen={chartSelection?.id === id}
                      chartLayout={
                        chartSelection?.id === id ? chartLayout : null
                      }
                      onChartSelection={onChartSelection}
                      viewportWidth={size.width}
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

function SlideThumbnailContent({ html }: { html: string }) {
  const content = React.useRef<HTMLSpanElement>(null)
  React.useLayoutEffect(() => {
    content.current!.innerHTML = html.replace(/ id="[^"]*"/g, "")
    renderCharts(content.current!)
  }, [html])
  return (
    <span
      ref={content}
      className="slides-thumbnail-content slides-content"
      aria-hidden="true"
      inert
    />
  )
}
