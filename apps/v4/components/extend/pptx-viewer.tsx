"use client"

import * as React from "react"
import {
  ReactPptxViewer,
  usePptxViewerThumbnails,
  type ParsedPresentation,
  type PptxSlideThumbnailItem,
  type PptxSlideThumbnailRenderWindow,
  type PptxViewerController,
  type PptxViewerError,
  type PresentationSource,
  type ViewerZoomLevel,
  type ViewerZoomState,
} from "@extend-ai/react-pptx"
import { flushSync } from "react-dom"

import "@extend-ai/react-pptx/styles.css"

import { useVirtualizer } from "@tanstack/react-virtual"

import { cn } from "@/lib/utils"
import { Button } from "@/components/ui/button"
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu"
import { Input } from "@/components/ui/input"
import { ScrollArea } from "@/components/ui/scroll-area"
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select"
import { Separator } from "@/components/ui/separator"
import { Spinner } from "@/components/ui/spinner"
import {
  Tooltip,
  TooltipContent,
  TooltipProvider,
  TooltipTrigger,
} from "@/components/ui/tooltip"
import {
  DocumentViewerThumbnailSidebar,
  useElementWidth,
  useInlineThumbnailSidebar,
} from "@/components/extend/document-viewer-sidebar"
import { AppearanceToggle } from "@/components/extend/appearance-toggle"
import { FileThumbnail } from "@/components/extend/file-thumbnail"
import { IconPlaceholder } from "@/components/icon-placeholder"

const PPTX_MIME_TYPE =
  "application/vnd.openxmlformats-officedocument.presentationml.presentation"
const PPT_MIME_TYPE = "application/vnd.ms-powerpoint"
const DEFAULT_ZOOM = 100
const PPTX_LOADING_INDICATOR_DELAY_MS = 300
const PPTX_THUMBNAIL_WIDTH = 112
const PPTX_THUMBNAIL_LIST_PADDING = 12
const PPTX_THUMBNAIL_ROW_ESTIMATE = 112
const PPTX_THUMBNAIL_PREFETCH_ROWS = 0
const PPTX_SCROLL_TOP_EPSILON_PX = 24
const PPTX_INSTANT_NAVIGATION_TIMEOUT_MS = 250
const PPTX_SMOOTH_NAVIGATION_TIMEOUT_MS = 1_000
const ZOOM_OPTIONS = [25, 50, 75, 100, 125, 150, 175, 200, 300, 400] as const
const ZOOM_MODE_LABELS = {
  "fit-page": "Fit page",
  "fit-width": "Fit width",
  automatic: "Automatic",
} satisfies Record<Exclude<ViewerZoomLevel, number>, string>

type UploadedPresentation = {
  file: File
  identity: string
  sourceUrl: string | undefined
}

function areNumberArraysEqual(
  left: readonly number[],
  right: readonly number[]
) {
  return (
    left.length === right.length &&
    left.every((value, index) => value === right[index])
  )
}

function formatPresentationName(fileName: string | undefined, url: string) {
  if (fileName?.trim()) return fileName

  const pathname = url.split("?")[0] ?? ""
  const rawName = pathname.split("/").pop() ?? "presentation.pptx"

  try {
    return decodeURIComponent(rawName)
  } catch {
    return rawName
  }
}

function ensurePresentationExtension(fileName: string) {
  const lowerFileName = fileName.toLowerCase()

  return lowerFileName.endsWith(".pptx") || lowerFileName.endsWith(".ppt")
    ? fileName
    : `${fileName}.pptx`
}

function downloadBlob(blob: Blob, fileName: string) {
  const url = URL.createObjectURL(blob)
  const anchor = document.createElement("a")

  anchor.href = url
  anchor.download = fileName
  anchor.rel = "noopener"
  document.body.append(anchor)
  anchor.click()
  anchor.remove()
  window.setTimeout(() => URL.revokeObjectURL(url), 0)
}

async function downloadPresentation({
  file,
  fileName,
  url,
}: {
  file?: File
  fileName: string
  url?: string
}) {
  if (file) {
    downloadBlob(file, ensurePresentationExtension(fileName))
    return
  }

  if (!url) return

  const response = await fetch(url)
  if (!response.ok) {
    throw new Error(`Failed to download presentation (${response.status})`)
  }

  downloadBlob(await response.blob(), ensurePresentationExtension(fileName))
}

function getNextZoom(currentZoom: number, direction: 1 | -1) {
  if (direction > 0) {
    return ZOOM_OPTIONS.find((value) => value > currentZoom) ?? currentZoom
  }

  for (let index = ZOOM_OPTIONS.length - 1; index >= 0; index -= 1) {
    const value = ZOOM_OPTIONS[index]
    if (value < currentZoom) return value
  }

  return currentZoom
}

function isZoomMode(value: string): value is Exclude<ViewerZoomLevel, number> {
  return value in ZOOM_MODE_LABELS
}

function useDelayedLoadingIndicator(isLoading: boolean, delayMs: number) {
  const [showSpinner, setShowSpinner] = React.useState(false)
  const [previousIsLoading, setPreviousIsLoading] = React.useState(isLoading)

  if (previousIsLoading !== isLoading) {
    setPreviousIsLoading(isLoading)
    setShowSpinner(false)
  }

  React.useEffect(() => {
    if (!isLoading) return

    const timeoutId = window.setTimeout(() => {
      setShowSpinner(true)
    }, delayMs)

    return () => window.clearTimeout(timeoutId)
  }, [delayMs, isLoading])

  return isLoading && showSpinner
}

function ViewerLoadingSurface({
  showSpinner = true,
}: {
  showSpinner?: boolean
}) {
  return (
    <div className="grid h-full min-h-96 w-full place-items-center bg-background">
      {showSpinner ? <Spinner className="size-4" /> : null}
    </div>
  )
}

function ToolbarTooltip({
  children,
  label,
}: {
  children: React.ReactNode
  label: string
}) {
  return (
    <Tooltip>
      <TooltipTrigger asChild>
        <span className="inline-flex">{children}</span>
      </TooltipTrigger>
      <TooltipContent side="bottom">{label}</TooltipContent>
    </Tooltip>
  )
}

function PptxFileActionsMenu({
  controlsDisabled,
  downloadDisabled,
  isPreparingDownload,
  onDownload,
  onUploadClick,
  showDownloadButton,
  showUploadButton,
}: {
  controlsDisabled: boolean
  downloadDisabled: boolean
  isPreparingDownload: boolean
  onDownload: () => void
  onUploadClick: () => void
  showDownloadButton: boolean
  showUploadButton: boolean
}) {
  if (!showDownloadButton && !showUploadButton) return null

  return (
    <DropdownMenu>
      <DropdownMenuTrigger asChild>
        <Button
          type="button"
          variant="ghost"
          size="icon-sm"
          aria-label="Open PowerPoint actions"
        >
          <IconPlaceholder
            lucide="Ellipsis"
            tabler="IconDots"
            hugeicons="MoreHorizontalIcon"
            phosphor="DotsThreeIcon"
            remixicon="RiMoreLine"
            className="size-4"
          />
        </Button>
      </DropdownMenuTrigger>
      <DropdownMenuContent align="end" className="w-44">
        {showDownloadButton ? (
          <DropdownMenuItem disabled={downloadDisabled} onClick={onDownload}>
            {isPreparingDownload ? (
              <Spinner className="size-4" />
            ) : (
              <IconPlaceholder
                lucide="Download"
                tabler="IconDownload"
                hugeicons="Download01Icon"
                phosphor="DownloadSimpleIcon"
                remixicon="RiDownload2Line"
                className="size-4"
              />
            )}
            Download
          </DropdownMenuItem>
        ) : null}
        {showUploadButton ? (
          <DropdownMenuItem disabled={controlsDisabled} onClick={onUploadClick}>
            <IconPlaceholder
              lucide="Upload"
              tabler="IconUpload"
              hugeicons="Upload01Icon"
              phosphor="UploadSimpleIcon"
              remixicon="RiUpload2Line"
              className="size-4"
            />
            Upload
          </DropdownMenuItem>
        ) : null}
      </DropdownMenuContent>
    </DropdownMenu>
  )
}

function PptxSlideNumberControl({
  activeSlideIndex,
  controlsDisabled,
  onSlideChange,
  slideCount,
}: {
  activeSlideIndex: number
  controlsDisabled: boolean
  onSlideChange: (slideIndex: number) => void
  slideCount: number
}) {
  const inputRef = React.useRef<HTMLInputElement>(null)
  const displaySlide = slideCount ? activeSlideIndex + 1 : 1
  const [isEditing, setIsEditing] = React.useState(false)
  const [draftSlide, setDraftSlide] = React.useState(() => String(displaySlide))

  React.useEffect(() => {
    if (!isEditing) return

    inputRef.current?.focus()
    inputRef.current?.select()
  }, [isEditing])

  const applySlideDraft = React.useCallback(
    (value: string) => {
      const parsedSlide = Number(value.trim())
      if (!Number.isInteger(parsedSlide)) return

      const nextSlide = Math.min(
        Math.max(parsedSlide, 1),
        Math.max(slideCount, 1)
      )
      onSlideChange(nextSlide - 1)
    },
    [onSlideChange, slideCount]
  )

  return (
    <div className="flex items-center text-sm whitespace-nowrap text-muted-foreground">
      <span>Slide</span>
      {isEditing ? (
        <Input
          ref={inputRef}
          aria-label="Slide number"
          inputMode="numeric"
          pattern="[0-9]*"
          size="sm"
          value={draftSlide}
          className="mx-1 w-14 min-w-14 rounded-md [&_[data-slot=input]]:text-center"
          onBlur={() => setIsEditing(false)}
          onChange={(event: React.ChangeEvent<HTMLInputElement>) => {
            setDraftSlide(event.target.value)
            applySlideDraft(event.target.value)
          }}
          onKeyDown={(event: React.KeyboardEvent<HTMLInputElement>) => {
            if (event.key === "Enter" || event.key === "Escape") {
              event.currentTarget.blur()
            }
          }}
        />
      ) : (
        <Button
          type="button"
          variant="ghost"
          size="sm"
          className="font-normal"
          aria-label={`Current slide ${displaySlide}. Edit slide number`}
          disabled={controlsDisabled || !slideCount}
          onClick={() => {
            setDraftSlide(String(displaySlide))
            setIsEditing(true)
          }}
        >
          {displaySlide}
        </Button>
      )}
      <span>of {slideCount || "-"}</span>
    </div>
  )
}

function PptxToolbar({
  activeSlideIndex,
  controlsDisabled,
  fileName,
  isDark,
  isPreparingDownload,
  onDownload,
  onIsDarkChange,
  onPresent,
  onSlideChange,
  onToggleSidebar,
  onUploadClick,
  onZoomChange,
  resolvedZoom,
  showDownloadButton,
  showUploadButton,
  slideCount,
  toolbarActions,
  zoomLevel,
}: {
  activeSlideIndex: number
  controlsDisabled: boolean
  fileName?: string
  isDark?: boolean
  isPreparingDownload: boolean
  onDownload: () => void
  onIsDarkChange?: (isDark: boolean) => void
  onPresent: () => void
  onSlideChange: (slideIndex: number) => void
  onToggleSidebar: () => void
  onUploadClick: () => void
  onZoomChange: (zoomLevel: ViewerZoomLevel) => void
  resolvedZoom: number
  showDownloadButton: boolean
  showUploadButton: boolean
  slideCount: number
  toolbarActions?: React.ReactNode
  zoomLevel: ViewerZoomLevel
}) {
  const canGoPrevious = !controlsDisabled && activeSlideIndex > 0
  const canGoNext =
    !controlsDisabled && slideCount > 0 && activeSlideIndex < slideCount - 1
  const canZoomOut = !controlsDisabled && resolvedZoom > ZOOM_OPTIONS[0]
  const canZoomIn =
    !controlsDisabled && resolvedZoom < ZOOM_OPTIONS[ZOOM_OPTIONS.length - 1]
  const selectValue =
    typeof zoomLevel === "number" ? zoomLevel.toString() : zoomLevel
  const roundedZoom = Number(resolvedZoom.toFixed(2))
  const zoomOptions = ZOOM_OPTIONS.includes(
    roundedZoom as (typeof ZOOM_OPTIONS)[number]
  )
    ? ZOOM_OPTIONS
    : [...ZOOM_OPTIONS, roundedZoom].sort((left, right) => left - right)

  return (
    <div
      data-slot="pptx-viewer-toolbar"
      className="flex min-h-12 flex-wrap items-center justify-between gap-2 border-b bg-background px-3 py-2"
    >
      <TooltipProvider>
        <div className="flex min-w-0 flex-wrap items-center gap-1">
          <ToolbarTooltip label="Toggle thumbnails">
            <Button
              type="button"
              variant="ghost"
              size="icon-sm"
              aria-label="Toggle thumbnails"
              disabled={controlsDisabled}
              onClick={onToggleSidebar}
            >
              <IconPlaceholder
                lucide="PanelLeft"
                tabler="IconLayoutSidebar"
                hugeicons="SidebarLeftIcon"
                phosphor="SidebarIcon"
                remixicon="RiLayoutLeftLine"
                className="size-4"
              />
            </Button>
          </ToolbarTooltip>
          {fileName ? (
            <span
              className="min-w-0 truncate px-1 text-sm font-semibold"
              title={fileName}
            >
              {fileName}
            </span>
          ) : null}
          <Separator orientation="vertical" className="mx-1 h-4 self-center" />
          <ToolbarTooltip label="Previous slide">
            <Button
              type="button"
              variant="ghost"
              size="icon-sm"
              aria-label="Previous slide"
              disabled={!canGoPrevious}
              onClick={() => onSlideChange(activeSlideIndex - 1)}
            >
              <IconPlaceholder
                lucide="ChevronLeft"
                tabler="IconChevronLeft"
                hugeicons="ArrowLeft01Icon"
                phosphor="CaretLeftIcon"
                remixicon="RiArrowLeftSLine"
                className="size-4"
              />
            </Button>
          </ToolbarTooltip>
          <PptxSlideNumberControl
            activeSlideIndex={activeSlideIndex}
            controlsDisabled={controlsDisabled}
            onSlideChange={onSlideChange}
            slideCount={slideCount}
          />
          <ToolbarTooltip label="Next slide">
            <Button
              type="button"
              variant="ghost"
              size="icon-sm"
              aria-label="Next slide"
              disabled={!canGoNext}
              onClick={() => onSlideChange(activeSlideIndex + 1)}
            >
              <IconPlaceholder
                lucide="ArrowRight"
                tabler="IconArrowRight"
                hugeicons="ArrowRight01Icon"
                phosphor="ArrowRightIcon"
                remixicon="RiArrowRightLine"
                className="size-4"
              />
            </Button>
          </ToolbarTooltip>
        </div>
        <div className="ml-auto flex min-w-0 flex-wrap items-center justify-end gap-1">
          <div className="flex flex-none items-center gap-1">
            <ToolbarTooltip label="Zoom out">
              <Button
                type="button"
                variant="ghost"
                size="icon-sm"
                aria-label="Zoom out"
                disabled={!canZoomOut}
                onClick={() => onZoomChange(getNextZoom(resolvedZoom, -1))}
              >
                <IconPlaceholder
                  lucide="CircleMinus"
                  tabler="IconCircleMinus"
                  hugeicons="MinusSignCircleIcon"
                  phosphor="MinusCircleIcon"
                  remixicon="RiIndeterminateCircleLine"
                  className="size-4"
                />
              </Button>
            </ToolbarTooltip>
            <Select
              value={selectValue}
              onValueChange={(value) => {
                if (value === null) return
                onZoomChange(isZoomMode(value) ? value : Number(value))
              }}
              disabled={controlsDisabled}
              modal={false}
            >
              <SelectTrigger
                size="sm"
                className="w-[104px] min-w-[104px]"
                aria-label="Zoom level"
              >
                <SelectValue>{Math.round(resolvedZoom)}%</SelectValue>
              </SelectTrigger>
              <SelectContent align="end" alignItemWithTrigger={false}>
                {Object.entries(ZOOM_MODE_LABELS).map(([value, label]) => (
                  <SelectItem key={value} value={value}>
                    {label}
                  </SelectItem>
                ))}
                {zoomOptions.map((value) => (
                  <SelectItem key={value} value={value.toString()}>
                    {value}%
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
            <ToolbarTooltip label="Zoom in">
              <Button
                type="button"
                variant="ghost"
                size="icon-sm"
                aria-label="Zoom in"
                disabled={!canZoomIn}
                onClick={() => onZoomChange(getNextZoom(resolvedZoom, 1))}
              >
                <IconPlaceholder
                  lucide="CirclePlusIcon"
                  tabler="IconCirclePlusFilled"
                  hugeicons="PlusSignCircleIcon"
                  phosphor="PlusCircleIcon"
                  remixicon="RiAddCircleFill"
                  className="size-4"
                />
              </Button>
            </ToolbarTooltip>
          </div>
          {onIsDarkChange ? (
            <AppearanceToggle
              isDark={isDark ?? false}
              onIsDarkChange={onIsDarkChange}
            />
          ) : null}
          <Button
            type="button"
            variant="ghost"
            size="sm"
            disabled={controlsDisabled || !slideCount}
            onClick={onPresent}
          >
            <IconPlaceholder
              lucide="Play"
              tabler="IconPlayerPlay"
              hugeicons="PlayIcon"
              phosphor="PlayIcon"
              remixicon="RiPlayLine"
              className="size-4"
            />
            Present
          </Button>
          {toolbarActions ? (
            <>
              <Separator
                orientation="vertical"
                className="mx-1 h-4 self-center"
              />
              {toolbarActions}
            </>
          ) : null}
          {showDownloadButton || showUploadButton ? (
            <Separator
              orientation="vertical"
              className="mx-1 h-4 self-center"
            />
          ) : null}
          <PptxFileActionsMenu
            controlsDisabled={false}
            downloadDisabled={controlsDisabled || isPreparingDownload}
            isPreparingDownload={isPreparingDownload}
            onDownload={onDownload}
            onUploadClick={onUploadClick}
            showDownloadButton={showDownloadButton}
            showUploadButton={showUploadButton}
          />
        </div>
      </TooltipProvider>
    </div>
  )
}

function PptxSidebarThumbnail({
  aspectRatio,
  containerRef,
  displayFileName,
  isActive,
  slideNumber,
  status,
}: {
  aspectRatio: number
  containerRef: PptxSlideThumbnailItem["containerRef"]
  displayFileName: string
  isActive: boolean
  slideNumber: number
  status: PptxSlideThumbnailItem["status"]
}) {
  return (
    <FileThumbnail
      file={{
        name: `${displayFileName} slide ${slideNumber}`,
        type: PPTX_MIME_TYPE,
      }}
      previewAspectRatio={aspectRatio}
      previewClassName="rounded-sm bg-white"
      previewContent={
        <div
          ref={containerRef}
          className="size-full overflow-hidden bg-white [&_[data-rpv-slide-wrapper]]:!m-0"
        />
      }
      isLoading={status !== "ready" && status !== "error"}
      hasError={status === "error"}
      className={cn(
        "w-28 max-w-full rounded-sm border-0 shadow-none transition-shadow duration-150",
        isActive
          ? "ring-2 ring-ring"
          : "ring-0 group-hover/pptx-thumbnail:ring-1 group-hover/pptx-thumbnail:ring-border"
      )}
    />
  )
}

function PptxThumbnailSidebarList({
  activeSlideIndex,
  displayFileName,
  isLoading,
  onRenderWindowChange,
  onSelectSlide,
  sidebarOpen,
  slideCount,
  thumbnails,
}: {
  activeSlideIndex: number
  displayFileName: string
  isLoading: boolean
  onRenderWindowChange: (window: PptxSlideThumbnailRenderWindow) => void
  onSelectSlide: (slideIndex: number) => void
  sidebarOpen: boolean
  slideCount: number
  thumbnails: PptxSlideThumbnailItem[]
}) {
  const viewportRef = React.useRef<HTMLDivElement | null>(null)
  const thumbnailListboxId = React.useId()
  const visibleThumbnails = React.useMemo(
    () => thumbnails.slice(0, slideCount || 0),
    [slideCount, thumbnails]
  )
  const activeDescendantId = visibleThumbnails.length
    ? `${thumbnailListboxId}-slide-${activeSlideIndex + 1}`
    : undefined
  const virtualizer = useVirtualizer({
    count: visibleThumbnails.length,
    estimateSize: () => PPTX_THUMBNAIL_ROW_ESTIMATE,
    getItemKey: (index) => visibleThumbnails[index]?.slide.id ?? index,
    getScrollElement: () => viewportRef.current,
    overscan: 0,
    paddingStart: PPTX_THUMBNAIL_LIST_PADDING,
    paddingEnd: PPTX_THUMBNAIL_LIST_PADDING,
  })
  const virtualItems = virtualizer.getVirtualItems()
  const renderWindowSignature = virtualItems
    .map((virtualRow) => virtualRow.index)
    .join(",")
  const virtualSlideIndexes = React.useMemo(
    () =>
      renderWindowSignature
        ? renderWindowSignature.split(",").map((index) => Number(index))
        : [],
    [renderWindowSignature]
  )

  React.useEffect(() => {
    if (!sidebarOpen || isLoading || !visibleThumbnails.length) {
      onRenderWindowChange({
        prefetchSlideIndexes: [],
        visibleSlideIndexes: [],
      })
      return
    }

    const visibleSlideIndexes = virtualSlideIndexes
    const firstVirtualIndex = virtualSlideIndexes[0] ?? 0
    const lastVirtualIndex =
      virtualSlideIndexes[virtualSlideIndexes.length - 1] ?? firstVirtualIndex
    const firstPrefetchIndex = Math.max(
      0,
      firstVirtualIndex - PPTX_THUMBNAIL_PREFETCH_ROWS
    )
    const lastPrefetchIndex = Math.min(
      visibleThumbnails.length - 1,
      lastVirtualIndex + PPTX_THUMBNAIL_PREFETCH_ROWS
    )
    const visibleSlideIndexSet = new Set(visibleSlideIndexes)
    const prefetchSlideIndexes: number[] = []

    for (
      let index = firstPrefetchIndex;
      index <= lastPrefetchIndex;
      index += 1
    ) {
      if (!visibleSlideIndexSet.has(index)) {
        prefetchSlideIndexes.push(index)
      }
    }

    onRenderWindowChange({
      prefetchSlideIndexes,
      visibleSlideIndexes,
    })
  }, [
    isLoading,
    onRenderWindowChange,
    sidebarOpen,
    virtualSlideIndexes,
    visibleThumbnails.length,
  ])

  React.useEffect(() => {
    if (!sidebarOpen || isLoading || !visibleThumbnails.length) return

    const frameId = window.requestAnimationFrame(() => {
      virtualizer.scrollToIndex(
        Math.min(activeSlideIndex, visibleThumbnails.length - 1),
        { align: "auto" }
      )
    })

    return () => window.cancelAnimationFrame(frameId)
  }, [
    activeSlideIndex,
    isLoading,
    sidebarOpen,
    virtualizer,
    visibleThumbnails.length,
  ])

  const handleKeyDown = React.useCallback(
    (event: React.KeyboardEvent<HTMLDivElement>) => {
      if (slideCount < 1) return

      if (event.altKey || event.ctrlKey || event.metaKey) return
      const nextSlideIndex = slideIndexForKey(
        event.key,
        activeSlideIndex,
        slideCount
      )

      if (nextSlideIndex === null) return

      event.preventDefault()
      onSelectSlide(nextSlideIndex)
    },
    [activeSlideIndex, onSelectSlide, slideCount]
  )

  return (
    <ScrollArea
      className="h-full [&_[data-slot=scroll-area-scrollbar]]:hidden"
      orientation="vertical"
      viewportClassName="focus-visible:ring-0 focus-visible:ring-offset-0"
      viewportProps={{
        "aria-activedescendant": activeDescendantId,
        "aria-busy": isLoading || undefined,
        "aria-label": "PowerPoint slides",
        onKeyDown: handleKeyDown,
        onMouseDown: (event) => {
          event.preventDefault()
          event.currentTarget.focus({ preventScroll: true })
        },
        role: "listbox",
        tabIndex: 0,
      }}
      viewportRef={viewportRef}
    >
      {isLoading ? (
        <div className="p-4">
          <div className="mx-auto aspect-video w-28 overflow-hidden rounded-sm bg-background shadow-xs">
            <div className="h-full animate-pulse bg-muted" />
          </div>
          <div className="mx-auto mt-3 h-3 w-10 rounded-full bg-muted" />
        </div>
      ) : visibleThumbnails.length ? (
        <div
          className="relative"
          style={{
            height: virtualizer.getTotalSize(),
          }}
        >
          {virtualItems.map((virtualRow) => {
            const thumbnail = visibleThumbnails[virtualRow.index]
            if (!thumbnail) return null

            const isActive = thumbnail.slideIndex === activeSlideIndex

            return (
              <div
                key={virtualRow.key}
                ref={virtualizer.measureElement}
                data-index={virtualRow.index}
                className={cn(
                  "absolute top-0 right-2 left-2 pb-2 [contain:layout_paint_style]",
                  isActive && "z-10"
                )}
                style={{
                  transform: `translateY(${virtualRow.start}px)`,
                }}
              >
                <div
                  id={`${thumbnailListboxId}-slide-${thumbnail.slideNumber}`}
                  role="option"
                  aria-current={isActive ? "page" : undefined}
                  aria-label={`Slide ${thumbnail.slideNumber}`}
                  aria-posinset={thumbnail.slideNumber}
                  aria-selected={isActive}
                  aria-setsize={slideCount}
                  className={cn(
                    "group/pptx-thumbnail flex h-auto w-full cursor-default flex-col items-center gap-2 p-2 text-xs outline-none select-none",
                    isActive
                      ? "text-foreground"
                      : "text-muted-foreground"
                  )}
                  onClick={() => onSelectSlide(thumbnail.slideIndex)}
                >
                  <PptxSidebarThumbnail
                    aspectRatio={thumbnail.aspectRatio}
                    containerRef={thumbnail.containerRef}
                    displayFileName={displayFileName}
                    isActive={isActive}
                    slideNumber={thumbnail.slideNumber}
                    status={thumbnail.status}
                  />
                  {thumbnail.slideNumber}
                </div>
              </div>
            )
          })}
        </div>
      ) : null}
    </ScrollArea>
  )
}

function PptxThumbnailSidebarContent({
  activeSlideIndex,
  controller,
  displayFileName,
  isLoading,
  onSelectSlide,
  sidebarOpen,
  slideCount,
}: {
  activeSlideIndex: number
  controller: PptxViewerController | null
  displayFileName: string
  isLoading: boolean
  onSelectSlide: (slideIndex: number) => void
  sidebarOpen: boolean
  slideCount: number
}) {
  const [renderWindow, setRenderWindow] =
    React.useState<PptxSlideThumbnailRenderWindow>({
      prefetchSlideIndexes: [],
      visibleSlideIndexes: [],
    })
  const thumbnailOptions = React.useMemo(
    () => ({
      renderWindow,
      resolution: {
        maxHeight: Math.round(PPTX_THUMBNAIL_WIDTH * 0.75),
        maxWidth: PPTX_THUMBNAIL_WIDTH,
      },
    }),
    [renderWindow]
  )
  const { thumbnails } = usePptxViewerThumbnails(controller, thumbnailOptions)
  const handleRenderWindowChange = React.useCallback(
    (nextWindow: PptxSlideThumbnailRenderWindow) => {
      setRenderWindow((currentWindow) => {
        const currentVisible = currentWindow.visibleSlideIndexes ?? []
        const nextVisible = nextWindow.visibleSlideIndexes ?? []
        const currentPrefetch = currentWindow.prefetchSlideIndexes ?? []
        const nextPrefetch = nextWindow.prefetchSlideIndexes ?? []

        if (
          areNumberArraysEqual(currentVisible, nextVisible) &&
          areNumberArraysEqual(currentPrefetch, nextPrefetch)
        ) {
          return currentWindow
        }

        return nextWindow
      })
    },
    []
  )

  return (
    <PptxThumbnailSidebarList
      activeSlideIndex={activeSlideIndex}
      displayFileName={displayFileName}
      isLoading={isLoading}
      onRenderWindowChange={handleRenderWindowChange}
      onSelectSlide={onSelectSlide}
      sidebarOpen={sidebarOpen}
      slideCount={slideCount}
      thumbnails={thumbnails}
    />
  )
}

function slideIndexForKey(key: string, index: number, count: number) {
  switch (key) {
    case "ArrowDown":
    case "ArrowRight":
    case "PageDown":
    case " ":
      return Math.min(count - 1, index + 1)
    case "ArrowUp":
    case "ArrowLeft":
    case "PageUp":
      return Math.max(0, index - 1)
    case "Home":
      return 0
    case "End":
      return count - 1
    default:
      return null
  }
}

function isInteractiveElement(target: EventTarget | null) {
  return (
    target instanceof Element &&
    Boolean(
      target.closest(
        'input, textarea, select, button, a, [contenteditable]:not([contenteditable="false"]), [role="combobox"], [role="menu"], [role="slider"]'
      )
    )
  )
}

function PptxSlideshow({
  containerRef,
  initialSlide,
  onExit,
  presentation,
}: {
  containerRef: React.RefObject<HTMLDivElement | null>
  initialSlide: number
  onExit: () => void
  presentation: ParsedPresentation
}) {
  const [slideIndex, setSlideIndex] = React.useState(initialSlide)
  const slideCount = presentation.document.slides.length

  return (
    <div
      ref={containerRef}
      role="dialog"
      aria-modal="true"
      aria-label="PowerPoint slideshow"
      tabIndex={-1}
      className="fixed inset-0 z-50 bg-black text-white outline-none"
      onKeyDown={(event) => {
        if (event.key === "Escape") {
          event.preventDefault()
          onExit()
          return
        }
        if (
          event.altKey ||
          event.ctrlKey ||
          event.metaKey ||
          isInteractiveElement(event.target)
        )
          return
        const nextSlide = slideIndexForKey(event.key, slideIndex, slideCount)
        if (nextSlide === null) return
        event.preventDefault()
        event.stopPropagation()
        setSlideIndex(nextSlide)
      }}
      onClick={(event) => {
        if (isInteractiveElement(event.target)) return
        setSlideIndex((index) => Math.min(slideCount - 1, index + 1))
      }}
    >
      <ReactPptxViewer
        source={presentation}
        mode="slide"
        slideIndex={slideIndex}
        zoom="fit-page"
        height="100dvh"
        slideIdPrefix="slideshow-"
        className="!border-0 !bg-black [&_.rpv-stage]:!bg-black [&_.rpv-status]:!bg-black [&_.rpv-status]:text-white"
        viewportClassName="!flex !items-center !justify-center !overflow-hidden !p-0"
      />
      <button
        type="button"
        className="absolute top-3 right-3 rounded bg-black/70 px-3 py-2 text-sm text-white focus-visible:outline-2"
        onClick={onExit}
      >
        Exit presentation
      </button>
    </div>
  )
}

export type PptxViewerPreviewProps = {
  className?: string
  defaultThumbnailSidebarOpen?: boolean
  defaultZoom?: ViewerZoomLevel
  fileName?: string
  initialSlide?: number
  /** The page's theme; with onIsDarkChange the toolbar shows a light/dark switch. */
  isDark?: boolean
  onIsDarkChange?: (isDark: boolean) => void
  showDownload?: boolean
  showToolbar?: boolean
  showUpload?: boolean
  /**
   * Let the page scroll the slides: no fixed height, a sticky toolbar, and
   * the window as the scroll container. Use it when the viewer is the page.
   */
  scrollsWindow?: boolean
  /** Prefix for the slide wrapper ids (`${slideIdPrefix}slide-${n}`). */
  slideIdPrefix?: string
  src?: string
  toolbarActions?: React.ReactNode
}

export function PptxViewerPreview({
  className,
  defaultThumbnailSidebarOpen = false,
  defaultZoom = DEFAULT_ZOOM,
  fileName,
  initialSlide = 1,
  isDark,
  onIsDarkChange,
  showDownload = true,
  showToolbar = true,
  showUpload = true,
  scrollsWindow = false,
  slideIdPrefix,
  src,
  toolbarActions,
}: PptxViewerPreviewProps) {
  const fileInputRef = React.useRef<HTMLInputElement>(null)
  const rootRef = React.useRef<HTMLDivElement>(null)
  const toolbarRef = React.useRef<HTMLDivElement>(null)
  const fullscreenRef = React.useRef<HTMLDivElement>(null)
  const [presentation, setPresentation] =
    React.useState<ParsedPresentation | null>(null)
  const [isPresenting, setIsPresenting] = React.useState(false)
  const [presentationError, setPresentationError] = React.useState<string>()
  const returnPositionRef = React.useRef<{
    focus: HTMLElement | null
    windowX: number
    windowY: number
    viewportX: number
    viewportY: number
  } | null>(null)
  const [viewerShellRef, viewerShellWidth] = useElementWidth<HTMLDivElement>()
  const requestedInitialSlideIndex = Math.max(0, Math.round(initialSlide) - 1)
  const [viewportElement, setViewportElement] =
    React.useState<HTMLDivElement | null>(null)
  const [uploadedPresentation, setUploadedPresentation] =
    React.useState<UploadedPresentation | null>(null)
  // Null until the reader toggles the slide strip; until then it is open by
  // default only where it fits beside the slides, as in Google Slides.
  const [sidebarChoice, setSidebarChoice] = React.useState<boolean | null>(null)
  const [thumbnailSidebarMounted, setThumbnailSidebarMounted] =
    React.useState(false)
  const [activeSlideIndex, setActiveSlideIndex] = React.useState(
    requestedInitialSlideIndex
  )
  const [zoomLevel, setZoomLevel] = React.useState(defaultZoom)
  const [resolvedZoom, setResolvedZoom] = React.useState(() =>
    typeof defaultZoom === "number" ? defaultZoom : DEFAULT_ZOOM
  )
  const [slideCount, setSlideCount] = React.useState(0)
  const [isLoading, setIsLoading] = React.useState(Boolean(src))
  const [loadError, setLoadError] = React.useState<string>()
  const [isPreparingDownload, setIsPreparingDownload] = React.useState(false)
  const [controller, setController] =
    React.useState<PptxViewerController | null>(null)
  const sidebarInline = useInlineThumbnailSidebar(viewerShellWidth)
  const sidebarOpen =
    sidebarChoice ?? (defaultThumbnailSidebarOpen && sidebarInline)
  const activeUploadedPresentation =
    uploadedPresentation?.sourceUrl === src ? uploadedPresentation : null
  const source: PresentationSource | undefined =
    activeUploadedPresentation?.file ?? src
  const sourceIdentity = activeUploadedPresentation?.identity ?? src ?? ""
  const hasPresentation = Boolean(source)
  const displayFileName = React.useMemo(
    () =>
      activeUploadedPresentation?.file.name ??
      (src
        ? formatPresentationName(fileName, src)
        : (fileName ?? "presentation.pptx")),
    [activeUploadedPresentation?.file.name, fileName, src]
  )
  const thumbnailSidebarVisible = Boolean(sidebarOpen && hasPresentation)
  const controlsDisabled = !hasPresentation || isLoading || Boolean(loadError)
  const shouldShowLoadingSpinner = useDelayedLoadingIndicator(
    isLoading,
    PPTX_LOADING_INDICATOR_DELAY_MS
  )
  const virtualization = React.useMemo(
    () => ({
      enabled: true,
      overscanViewport: 0.75,
      scrollElement:
        scrollsWindow && viewportElement ? window : viewportElement,
    }),
    [scrollsWindow, viewportElement]
  )
  // How far the slide list has scrolled past the top of its visible area. In
  // window mode that area starts below the sticky toolbar.
  const getViewportScrollTop = React.useCallback(
    (viewport: HTMLDivElement) =>
      scrollsWindow
        ? (toolbarRef.current?.offsetHeight ?? 0) -
          viewport.getBoundingClientRect().top
        : viewport.scrollTop,
    [scrollsWindow]
  )
  const pendingSlideNavigationRef = React.useRef<{
    hasObservedIntermediateSlide: boolean
    targetIndex: number
    timeoutId: number
  } | null>(null)
  const cancelPendingSlideNavigation = React.useCallback(() => {
    const pendingNavigation = pendingSlideNavigationRef.current

    if (pendingNavigation) {
      window.clearTimeout(pendingNavigation.timeoutId)
      pendingSlideNavigationRef.current = null
    }
  }, [])

  if (thumbnailSidebarVisible && !thumbnailSidebarMounted) {
    setThumbnailSidebarMounted(true)
  }

  const [previousSource, setPreviousSource] = React.useState({
    sourceIdentity,
    defaultZoom,
  })
  if (
    previousSource.sourceIdentity !== sourceIdentity ||
    previousSource.defaultZoom !== defaultZoom
  ) {
    setPreviousSource({ sourceIdentity, defaultZoom })
    setSlideCount(0)
    setIsPresenting(false)
    setPresentationError(undefined)
    setZoomLevel(defaultZoom)
    setResolvedZoom(
      typeof defaultZoom === "number" ? defaultZoom : DEFAULT_ZOOM
    )
    setLoadError(undefined)
    setIsLoading(Boolean(sourceIdentity))
  }

  React.useEffect(() => {
    cancelPendingSlideNavigation()
    if (!viewportElement) return
    if (!scrollsWindow) {
      viewportElement.scrollTo({ top: 0, left: 0 })
      return
    }
    const scrollTop = getViewportScrollTop(viewportElement)
    if (scrollTop > 0) {
      window.scrollTo({ top: window.scrollY - scrollTop, left: 0 })
    }
  }, [
    cancelPendingSlideNavigation,
    defaultZoom,
    getViewportScrollTop,
    scrollsWindow,
    sourceIdentity,
    viewportElement,
  ])

  React.useEffect(() => {
    const root = rootRef.current
    const toolbarElement = toolbarRef.current
    if (!scrollsWindow || !root || !toolbarElement) return

    const observer = new ResizeObserver(() => {
      root.style.setProperty(
        "--pptx-toolbar-height",
        `${toolbarElement.offsetHeight}px`
      )
    })
    observer.observe(toolbarElement)
    return () => observer.disconnect()
  }, [scrollsWindow, showToolbar])

  React.useEffect(
    () => cancelPendingSlideNavigation,
    [cancelPendingSlideNavigation]
  )

  const [initialSlideSource, setInitialSlideSource] = React.useState({
    sourceIdentity,
    requestedInitialSlideIndex,
  })
  if (
    initialSlideSource.sourceIdentity !== sourceIdentity ||
    initialSlideSource.requestedInitialSlideIndex !== requestedInitialSlideIndex
  ) {
    setInitialSlideSource({ sourceIdentity, requestedInitialSlideIndex })
    setActiveSlideIndex(requestedInitialSlideIndex)
  }

  React.useEffect(() => {
    if (controller?.isReady()) {
      void controller.goToSlide(requestedInitialSlideIndex, {
        behavior: "instant",
        block: "center",
      })
    }
  }, [requestedInitialSlideIndex, sourceIdentity, controller])

  const navigateToSlide = React.useCallback(
    (nextSlideIndex: number, behavior: ScrollBehavior) => {
      const normalizedSlideIndex = Math.min(
        Math.max(0, Math.round(nextSlideIndex)),
        Math.max(0, slideCount - 1)
      )
      const resolvedBehavior =
        behavior === "smooth" &&
        window.matchMedia("(prefers-reduced-motion: reduce)").matches
          ? "instant"
          : behavior

      cancelPendingSlideNavigation()
      const timeoutId = window.setTimeout(
        () => {
          const pendingNavigation = pendingSlideNavigationRef.current

          if (pendingNavigation?.targetIndex !== normalizedSlideIndex) return

          pendingSlideNavigationRef.current = null
          setActiveSlideIndex(normalizedSlideIndex)
        },
        resolvedBehavior === "smooth"
          ? PPTX_SMOOTH_NAVIGATION_TIMEOUT_MS
          : PPTX_INSTANT_NAVIGATION_TIMEOUT_MS
      )
      pendingSlideNavigationRef.current = {
        hasObservedIntermediateSlide: false,
        targetIndex: normalizedSlideIndex,
        timeoutId,
      }

      setActiveSlideIndex(normalizedSlideIndex)
      void controller?.goToSlide(normalizedSlideIndex, {
        behavior: resolvedBehavior,
        block: "center",
      })
    },
    [cancelPendingSlideNavigation, slideCount, controller]
  )
  const handleSlideChange = React.useCallback(
    (nextSlideIndex: number) => navigateToSlide(nextSlideIndex, "smooth"),
    [navigateToSlide]
  )
  const handleThumbnailSlideChange = React.useCallback(
    (nextSlideIndex: number) => navigateToSlide(nextSlideIndex, "instant"),
    [navigateToSlide]
  )
  const handleViewerSlideChange = React.useCallback(
    (nextSlideIndex: number) => {
      if (returnPositionRef.current) return
      const pendingNavigation = pendingSlideNavigationRef.current

      if (pendingNavigation) {
        if (nextSlideIndex !== pendingNavigation.targetIndex) {
          pendingNavigation.hasObservedIntermediateSlide = true
          return
        }

        // The controller reports the destination immediately, before the smooth
        // scroll starts. Keep the destination authoritative until the scroll
        // observer has crossed an intermediate slide and reports it again.
        if (!pendingNavigation.hasObservedIntermediateSlide) return

        window.clearTimeout(pendingNavigation.timeoutId)
        pendingSlideNavigationRef.current = null
      }

      React.startTransition(() => {
        setActiveSlideIndex((currentSlideIndex) =>
          currentSlideIndex === nextSlideIndex
            ? currentSlideIndex
            : nextSlideIndex
        )
      })
    },
    []
  )
  const syncFirstSlideAtViewportTop = React.useCallback(
    (viewport: HTMLDivElement) => {
      if (
        returnPositionRef.current ||
        getViewportScrollTop(viewport) > PPTX_SCROLL_TOP_EPSILON_PX
      )
        return

      setActiveSlideIndex((currentSlideIndex) =>
        currentSlideIndex === 0 ? currentSlideIndex : 0
      )

      if (!isLoading && controller && controller.getSlideIndex() !== 0) {
        void controller.goToSlide(0, {
          behavior: "instant",
          block: "start",
        })
      }
    },
    [getViewportScrollTop, isLoading, controller]
  )
  const interruptViewerScroll = React.useCallback(
    (viewport: HTMLDivElement) => {
      if (returnPositionRef.current) return
      const scroller = scrollsWindow ? window : viewport

      // Stop an in-flight native smooth scroll before manual wheel, touch,
      // pointer, or keyboard input takes ownership of the viewport.
      scroller.scrollTo({
        behavior: "instant",
        left: scrollsWindow ? window.scrollX : viewport.scrollLeft,
        top: scrollsWindow ? window.scrollY : viewport.scrollTop,
      })
      cancelPendingSlideNavigation()
      syncFirstSlideAtViewportTop(viewport)
    },
    [cancelPendingSlideNavigation, scrollsWindow, syncFirstSlideAtViewportTop]
  )
  const handleViewerUserScrollIntent = React.useCallback(
    (event: React.SyntheticEvent<HTMLDivElement>) =>
      interruptViewerScroll(event.currentTarget),
    [interruptViewerScroll]
  )
  const handleViewerScroll = React.useCallback(
    (event: React.UIEvent<HTMLDivElement>) => {
      if (pendingSlideNavigationRef.current) return
      syncFirstSlideAtViewportTop(event.currentTarget)
    },
    [syncFirstSlideAtViewportTop]
  )
  const handleNavigationKey = React.useCallback(
    (event: KeyboardEvent | React.KeyboardEvent<HTMLDivElement>) => {
      if (
        controlsDisabled ||
        !slideCount ||
        returnPositionRef.current ||
        event.defaultPrevented ||
        event.altKey ||
        event.ctrlKey ||
        event.metaKey ||
        isInteractiveElement(event.target)
      )
        return
      const nextSlideIndex = slideIndexForKey(
        event.key,
        activeSlideIndex,
        slideCount
      )
      if (nextSlideIndex === null) return
      event.preventDefault()
      navigateToSlide(nextSlideIndex, "instant")
    },
    [activeSlideIndex, controlsDisabled, navigateToSlide, slideCount]
  )

  React.useEffect(() => {
    if (!scrollsWindow || !viewportElement) return

    const handleIntent = (event: Event) => {
      if (returnPositionRef.current) return
      // Scrolling the thumbnail strip must not interrupt slide navigation.
      if (
        event.target instanceof Node &&
        viewportElement.contains(event.target)
      ) {
        interruptViewerScroll(viewportElement)
      }
    }
    const handleScroll = () => {
      if (returnPositionRef.current || pendingSlideNavigationRef.current) return
      syncFirstSlideAtViewportTop(viewportElement)
    }
    const handleKeyDown = (event: KeyboardEvent) => {
      if (
        event.target instanceof Node &&
        event.target !== document.body &&
        event.target !== document.documentElement &&
        !rootRef.current?.contains(event.target)
      )
        return
      handleNavigationKey(event)
    }

    window.addEventListener("keydown", handleKeyDown)
    window.addEventListener("scroll", handleScroll, { passive: true })
    window.addEventListener("touchstart", handleIntent, { passive: true })
    window.addEventListener("wheel", handleIntent, { passive: true })
    return () => {
      window.removeEventListener("keydown", handleKeyDown)
      window.removeEventListener("scroll", handleScroll)
      window.removeEventListener("touchstart", handleIntent)
      window.removeEventListener("wheel", handleIntent)
    }
  }, [
    handleNavigationKey,
    interruptViewerScroll,
    scrollsWindow,
    syncFirstSlideAtViewportTop,
    viewportElement,
  ])

  React.useEffect(() => {
    if (!isPresenting) return
    const handleFullscreenChange = () => {
      if (document.fullscreenElement !== fullscreenRef.current)
        setIsPresenting(false)
    }
    document.addEventListener("fullscreenchange", handleFullscreenChange)
    return () => {
      document.removeEventListener("fullscreenchange", handleFullscreenChange)
      const position = returnPositionRef.current
      if (!position) return
      position.focus?.focus({ preventScroll: true })
      window.scrollTo({
        left: position.windowX,
        top: position.windowY,
        behavior: "instant",
      })
      viewportElement?.scrollTo({
        left: position.viewportX,
        top: position.viewportY,
        behavior: "instant",
      })
      returnPositionRef.current = null
    }
  }, [isPresenting, viewportElement])

  const handlePresent = async () => {
    if (!presentation || controlsDisabled || !slideCount) return
    setPresentationError(undefined)
    cancelPendingSlideNavigation()
    returnPositionRef.current = {
      focus:
        document.activeElement instanceof HTMLElement
          ? document.activeElement
          : null,
      windowX: window.scrollX,
      windowY: window.scrollY,
      viewportX: viewportElement?.scrollLeft ?? 0,
      viewportY: viewportElement?.scrollTop ?? 0,
    }
    // Mount the fullscreen target within the click's user activation.
    flushSync(() => setIsPresenting(true))
    try {
      await fullscreenRef.current!.requestFullscreen()
      fullscreenRef.current?.focus({ preventScroll: true })
    } catch {
      setIsPresenting(false)
      setPresentationError(
        "Unable to enter full screen. Please try presenting again."
      )
    }
  }
  const handleExitPresentation = () => {
    if (document.fullscreenElement === fullscreenRef.current) {
      void document.exitFullscreen().catch(() => {
        setPresentationError(
          "Unable to exit full screen. Press Escape to exit."
        )
      })
    } else {
      setIsPresenting(false)
    }
  }

  const handleLoad = React.useCallback((presentation: ParsedPresentation) => {
    setPresentation(presentation)
    const nextSlideCount = presentation.document.slides.length

    setSlideCount(nextSlideCount)
    setActiveSlideIndex((currentSlideIndex) =>
      Math.min(currentSlideIndex, Math.max(0, nextSlideCount - 1))
    )
    setLoadError(undefined)
  }, [])

  const handleReady = React.useCallback(() => {
    setIsLoading(false)
  }, [])

  const handleZoomChange = React.useCallback((state: ViewerZoomState) => {
    setZoomLevel(state.level)
    setResolvedZoom(state.resolvedZoom)
  }, [])

  const handleError = React.useCallback((error: PptxViewerError) => {
    setLoadError(error.message)
    setIsLoading(false)
  }, [])

  const handleDownload = React.useCallback(async () => {
    if (isPreparingDownload || !source) return

    setIsPreparingDownload(true)

    try {
      await downloadPresentation({
        file: activeUploadedPresentation?.file,
        fileName: displayFileName,
        url: activeUploadedPresentation ? undefined : src,
      })
    } catch (error) {
      console.error(error)
    } finally {
      setIsPreparingDownload(false)
    }
  }, [
    activeUploadedPresentation,
    displayFileName,
    isPreparingDownload,
    source,
    src,
  ])

  async function handleUpload(event: React.ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0]
    event.target.value = ""

    if (!file) return

    setUploadedPresentation({
      file,
      identity: `${file.name}-${file.size}-${file.lastModified}`,
      sourceUrl: src,
    })
  }

  const toolbar = showToolbar ? (
    <PptxToolbar
      activeSlideIndex={activeSlideIndex}
      controlsDisabled={controlsDisabled}
      fileName={fileName ? displayFileName : undefined}
      isDark={isDark}
      isPreparingDownload={isPreparingDownload}
      onDownload={handleDownload}
      onIsDarkChange={onIsDarkChange}
      onPresent={handlePresent}
      onSlideChange={handleSlideChange}
      onToggleSidebar={() => setSidebarChoice(!sidebarOpen)}
      onUploadClick={() => fileInputRef.current?.click()}
      onZoomChange={setZoomLevel}
      resolvedZoom={resolvedZoom}
      showDownloadButton={showDownload}
      showUploadButton={showUpload}
      slideCount={slideCount}
      toolbarActions={toolbarActions}
      zoomLevel={zoomLevel}
    />
  ) : null
  const slides = !source ? (
    <div className="grid h-full min-h-96 place-items-center p-6 text-center">
      <div className="max-w-md rounded-lg border bg-background p-4 text-sm shadow-xs">
        <div className="font-medium">
          Upload a PowerPoint presentation to preview
        </div>
        <div className="mt-1 text-muted-foreground">
          Pass a PPTX URL with the <code>src</code> prop or upload a PPTX or
          legacy PPT file.
        </div>
        <Button
          type="button"
          variant="outline"
          size="sm"
          className="mt-4"
          onClick={() => fileInputRef.current?.click()}
        >
          <IconPlaceholder
            lucide="Upload"
            tabler="IconUpload"
            hugeicons="Upload01Icon"
            phosphor="UploadSimpleIcon"
            remixicon="RiUpload2Line"
            className="size-4"
          />
          Upload PowerPoint
        </Button>
      </div>
    </div>
  ) : (
    <ReactPptxViewer
      key={sourceIdentity}
      ref={setController}
      source={source}
      mode="continuous"
      initialSlide={requestedInitialSlideIndex}
      zoom={zoomLevel}
      height="100%"
      showToolbar={false}
      showThumbnails={false}
      virtualization={virtualization}
      slideIdPrefix={slideIdPrefix}
      className="min-h-full !border-0 !bg-transparent [&_.rpv-stage]:min-h-full [&_.rpv-stage]:!bg-transparent [&_.rpv-status]:!bg-transparent [&_.rpv-workspace]:min-h-full [&_[data-rpv-list-item]]:[contain:layout_paint_style]"
      viewportClassName="!min-h-full !px-4 !py-6"
      renderLoading={() => (
        <ViewerLoadingSurface showSpinner={shouldShowLoadingSpinner} />
      )}
      renderError={(error) => (
        <div className="grid h-full min-h-96 place-items-center p-6 text-center">
          <div className="max-w-md rounded-lg border bg-background p-4 text-sm text-destructive shadow-xs">
            <div className="font-medium">Unable to display PowerPoint</div>
            <div className="mt-1 text-muted-foreground">{error.message}</div>
          </div>
        </div>
      )}
      emptyState={
        <div className="text-sm text-muted-foreground">
          This presentation has no slides.
        </div>
      }
      onLoad={handleLoad}
      onReady={handleReady}
      onError={handleError}
      onSlideChange={handleViewerSlideChange}
      onZoomChange={handleZoomChange}
    />
  )

  return (
    <div
      ref={rootRef}
      className={cn(
        scrollsWindow
          ? "flex min-h-dvh flex-col bg-background [--pptx-toolbar-height:0px]"
          : "flex h-[640px] min-h-0 flex-col overflow-hidden bg-background",
        className
      )}
    >
      {isPresenting && presentation ? (
        <PptxSlideshow
          key={sourceIdentity}
          containerRef={fullscreenRef}
          initialSlide={activeSlideIndex}
          onExit={handleExitPresentation}
          presentation={presentation}
        />
      ) : null}
      {presentationError ? (
        <div role="alert" className="p-3 text-sm text-destructive">
          {presentationError}
        </div>
      ) : null}
      <input
        ref={fileInputRef}
        type="file"
        accept={`.ppt,.pptx,${PPT_MIME_TYPE},${PPTX_MIME_TYPE}`}
        className="hidden"
        onChange={handleUpload}
      />
      {scrollsWindow && toolbar ? (
        <div ref={toolbarRef} className="sticky top-0 z-40">
          {toolbar}
        </div>
      ) : (
        toolbar
      )}
      <div
        ref={viewerShellRef}
        className={
          scrollsWindow
            ? "relative flex flex-1 bg-muted/30"
            : "relative flex min-h-0 flex-1 overflow-hidden bg-muted/30"
        }
      >
        <DocumentViewerThumbnailSidebar
          className={
            scrollsWindow
              ? sidebarInline
                ? "sticky top-(--pptx-toolbar-height)! bottom-auto! h-[calc(100dvh-var(--pptx-toolbar-height))] self-start"
                : "fixed top-(--pptx-toolbar-height)!"
              : undefined
          }
          inline={sidebarInline}
          open={thumbnailSidebarVisible}
        >
          {thumbnailSidebarMounted && hasPresentation ? (
            <PptxThumbnailSidebarContent
              activeSlideIndex={activeSlideIndex}
              controller={controller}
              displayFileName={displayFileName}
              isLoading={isLoading}
              onSelectSlide={handleThumbnailSlideChange}
              sidebarOpen={thumbnailSidebarVisible}
              slideCount={slideCount}
            />
          ) : null}
        </DocumentViewerThumbnailSidebar>
        {scrollsWindow ? (
          <div
            ref={setViewportElement}
            aria-label="PowerPoint presentation"
            className="min-w-0 flex-1"
            onPointerDown={handleViewerUserScrollIntent}
          >
            {slides}
          </div>
        ) : (
          <ScrollArea
            className="min-h-0 flex-1"
            viewportClassName="p-0"
            viewportProps={{
              "aria-label": "PowerPoint presentation",
              onKeyDown: handleNavigationKey,
              onPointerDown: handleViewerUserScrollIntent,
              onScroll: handleViewerScroll,
              onTouchStart: handleViewerUserScrollIntent,
              onWheel: handleViewerUserScrollIntent,
              tabIndex: 0,
            }}
            viewportRef={setViewportElement}
          >
            {slides}
          </ScrollArea>
        )}
      </div>
    </div>
  )
}
