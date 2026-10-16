"use client"

import * as React from "react"
import {
  DocxEditorViewer,
  paragraphLetterheadFloatSideAtNodeIndex,
  useDocxBorders,
  useDocxComments,
  useDocxEditor,
  useDocxLineSpacing,
  useDocxPageLayout,
  useDocxParagraphStyles,
  useDocxTrackChanges,
  useDocxViewerThumbnails,
  type DocxBorderContext,
  type DocxBorderPreset,
  type DocxCommentCardRenderProps,
  type DocxEditorController,
  type DocxEditorViewerMode,
  type DocxTrackedChangeCardRenderProps,
  type ParagraphStyleDefinition,
} from "@extend-ai/react-docx"

import type { RegistryIconProps } from "@/lib/registry-icon-props"
import { cn } from "@/lib/utils"
import { Button } from "@/components/ui/button"
import { ColorPicker } from "@/components/ui/color-picker"
import {
  DropdownMenu,
  DropdownMenuCheckboxItem,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuSeparator,
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
import { ToggleGroup, ToggleGroupItem } from "@/components/ui/toggle"
import {
  Tooltip,
  TooltipContent,
  TooltipProvider,
  TooltipTrigger,
} from "@/components/ui/tooltip"
import {
  PreviewCard,
  PreviewCardContent,
  PreviewCardTrigger,
} from "@/components/extend/document-preview-card"
import {
  DocumentViewerThumbnailSidebar,
  useElementWidth,
  useInlineThumbnailSidebar,
} from "@/components/extend/document-viewer-sidebar"
import {
  createDocxCommentCardRenderer,
  createDocxTrackedChangeCardRenderer,
} from "@/components/extend/docx-annotation-card"
import { FileThumbnail } from "@/components/extend/file-thumbnail"
import { ToolbarOverflow } from "@/components/extend/toolbar-overflow"
import { IconPlaceholder } from "@/components/icon-placeholder"

function ArrowExpandDiagonal01Glyph(props: RegistryIconProps) {
  return (
    <IconPlaceholder
      lucide="ArrowDownRight"
      tabler="IconArrowDownRight"
      hugeicons="ArrowExpandDiagonal01Icon"
      phosphor="ArrowDownRightIcon"
      remixicon="RiArrowRightDownLine"
      {...props}
    />
  )
}

function ArrowExpandDiagonal02Glyph(props: RegistryIconProps) {
  return (
    <IconPlaceholder
      lucide="ArrowUpRight"
      tabler="IconArrowUpRight"
      hugeicons="ArrowExpandDiagonal02Icon"
      phosphor="ArrowUpRightIcon"
      remixicon="RiArrowRightUpLine"
      {...props}
    />
  )
}

function BorderAll01Glyph(props: RegistryIconProps) {
  return (
    <IconPlaceholder
      lucide="Grid2x2"
      tabler="IconBorderAll"
      hugeicons="BorderAll01Icon"
      phosphor="GridFourIcon"
      remixicon="RiGridLine"
      {...props}
    />
  )
}

function BorderBottom01Glyph(props: RegistryIconProps) {
  return (
    <IconPlaceholder
      lucide="PanelBottom"
      tabler="IconBorderBottom"
      hugeicons="BorderBottom01Icon"
      phosphor="SquareHalfBottomIcon"
      remixicon="RiLayoutBottomLine"
      {...props}
    />
  )
}

function BorderHorizontalGlyph(props: RegistryIconProps) {
  return (
    <IconPlaceholder
      lucide="SeparatorHorizontal"
      tabler="IconBorderHorizontal"
      hugeicons="BorderHorizontalIcon"
      phosphor="MinusIcon"
      remixicon="RiSeparator"
      {...props}
    />
  )
}

function BorderInnerGlyph(props: RegistryIconProps) {
  return (
    <IconPlaceholder
      lucide="SquarePlus"
      tabler="IconBorderInner"
      hugeicons="BorderInnerIcon"
      phosphor="PlusSquareIcon"
      remixicon="RiAddBoxLine"
      {...props}
    />
  )
}

function BorderLeft01Glyph(props: RegistryIconProps) {
  return (
    <IconPlaceholder
      lucide="PanelLeft"
      tabler="IconBorderLeft"
      hugeicons="BorderLeft01Icon"
      phosphor="SquareHalfIcon"
      remixicon="RiLayoutLeftLine"
      {...props}
    />
  )
}

function BorderNone01Glyph(props: RegistryIconProps) {
  return (
    <IconPlaceholder
      lucide="SquareDashed"
      tabler="IconBorderNone"
      hugeicons="BorderNone01Icon"
      phosphor="SelectionIcon"
      remixicon="RiCheckboxBlankLine"
      {...props}
    />
  )
}

function BorderRight01Glyph(props: RegistryIconProps) {
  return (
    <IconPlaceholder
      lucide="PanelRight"
      tabler="IconBorderRight"
      hugeicons="BorderRight01Icon"
      phosphor="SquareHalfIcon"
      remixicon="RiLayoutRightLine"
      {...props}
    />
  )
}

function BorderTop01Glyph(props: RegistryIconProps) {
  return (
    <IconPlaceholder
      lucide="PanelTop"
      tabler="IconBorderTop"
      hugeicons="BorderTop01Icon"
      phosphor="AlignTopSimpleIcon"
      remixicon="RiLayoutTopLine"
      {...props}
    />
  )
}

function BorderVerticalGlyph(props: RegistryIconProps) {
  return (
    <IconPlaceholder
      lucide="SeparatorVertical"
      tabler="IconBorderVertical"
      hugeicons="BorderVerticalIcon"
      phosphor="LineVerticalIcon"
      remixicon="RiLayoutColumnLine"
      {...props}
    />
  )
}

function LineGlyph(props: RegistryIconProps) {
  return (
    <IconPlaceholder
      lucide="Minus"
      tabler="IconMinus"
      hugeicons="LineIcon"
      phosphor="MinusIcon"
      remixicon="RiSubtractLine"
      {...props}
    />
  )
}

function Moon02Glyph(props: RegistryIconProps) {
  return (
    <IconPlaceholder
      lucide="Moon"
      tabler="IconMoon"
      hugeicons="Moon02Icon"
      phosphor="MoonIcon"
      remixicon="RiMoonLine"
      {...props}
    />
  )
}

function Sun03Glyph(props: RegistryIconProps) {
  return (
    <IconPlaceholder
      lucide="Sun"
      tabler="IconSun"
      hugeicons="Sun03Icon"
      phosphor="SunIcon"
      remixicon="RiSunLine"
      {...props}
    />
  )
}

function TextColorGlyph(props: RegistryIconProps) {
  return (
    <IconPlaceholder
      lucide="Baseline"
      tabler="IconTextColor"
      hugeicons="TextColorIcon"
      phosphor="TextAaIcon"
      remixicon="RiFontColor"
      {...props}
    />
  )
}

const DOCX_MIME_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
const DOCX_LOADING_INDICATOR_DELAY_MS = 300
const DOCX_THUMBNAIL_WIDTH = 92
const DOCX_EDITOR_DEFAULT_ZOOM_SCALE = 100
const ZOOM_OPTIONS = [50, 75, 90, 100, 110, 125, 150, 175, 200] as const
const FONT_FAMILIES = [
  "Calibri",
  "Arial",
  "Times New Roman",
  "Georgia",
  "Helvetica",
  "Courier New",
] as const
const FONT_SIZE_OPTIONS = [
  8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48,
] as const
const LINE_SPACING_OPTIONS = [1, 1.15, 1.2, 1.5, 2, 2.5, 3] as const
const FALLBACK_PARAGRAPH_STYLE_OPTIONS = [
  { id: "Normal", name: "Body" },
  { id: "Heading1", name: "Heading 1" },
  { id: "Heading2", name: "Heading 2" },
  { id: "Heading3", name: "Heading 3" },
] as const
const HIGHLIGHT_COLORS = [
  { label: "Yellow", value: "yellow", color: "#fff59d" },
  { label: "Green", value: "green", color: "#bbf7d0" },
  { label: "Cyan", value: "cyan", color: "#a5f3fc" },
  { label: "Magenta", value: "magenta", color: "#f5d0fe" },
  { label: "Red", value: "red", color: "#fecaca" },
  { label: "Blue", value: "blue", color: "#bfdbfe" },
] as const
const HIGHLIGHT_PREVIEW_COLORS: Record<string, string> = {
  black: "#111827",
  white: "#ffffff",
  ...Object.fromEntries(
    HIGHLIGHT_COLORS.map((option) => [option.value, option.color])
  ),
}
const DOCX_PADDING_WARNING_TEXT = "a style property during rerender"
const DEFAULT_HEADING_PREVIEW_RUN_STYLES: Record<
  1 | 2 | 3 | 4 | 5 | 6,
  NonNullable<ParagraphStyleDefinition["runStyle"]>
> = {
  1: {
    fontFamily: "Calibri Light",
    fontSizePt: 16,
    bold: true,
    color: "#2f5496",
  },
  2: {
    fontFamily: "Calibri Light",
    fontSizePt: 13,
    bold: true,
    color: "#2f5496",
  },
  3: {
    fontFamily: "Calibri",
    fontSizePt: 12,
    bold: true,
    color: "#1f3763",
  },
  4: {
    fontFamily: "Calibri",
    fontSizePt: 11,
    bold: true,
    color: "#1f3763",
  },
  5: {
    fontFamily: "Calibri",
    fontSizePt: 11,
    bold: true,
    color: "#1f3763",
  },
  6: {
    fontFamily: "Calibri",
    fontSizePt: 11,
    bold: true,
    color: "#1f3763",
  },
}

type UploadedDocxFile = {
  file: File
  identity: string
}

type BorderControlOption = {
  id: DocxBorderPreset
  label: string
  contexts?: DocxBorderContext[]
  separatorBefore?: boolean
}

type DocxFontToggleValue =
  | "bold"
  | "italic"
  | "underline"
  | "strike"
  | "superscript"
  | "subscript"

const BORDER_CONTROL_OPTIONS: BorderControlOption[] = [
  { id: "bottom", label: "Bottom Border" },
  { id: "top", label: "Top Border" },
  { id: "left", label: "Left Border" },
  { id: "right", label: "Right Border" },
  { id: "none", label: "No Border", separatorBefore: true },
  { id: "all", label: "All Borders" },
  { id: "outside", label: "Outside Borders" },
  { id: "inside", label: "Inside Borders", contexts: ["table"] },
  {
    id: "inside-horizontal",
    label: "Inside Horizontal Border",
    contexts: ["table"],
  },
  {
    id: "inside-vertical",
    label: "Inside Vertical Border",
    contexts: ["table"],
  },
  {
    id: "diagonal-down",
    label: "Diagonal Down Border",
    contexts: ["table"],
    separatorBefore: true,
  },
  { id: "diagonal-up", label: "Diagonal Up Border", contexts: ["table"] },
  { id: "horizontal-line", label: "Horizontal Line", separatorBefore: true },
]

function borderControlOptionIcon(optionId: DocxBorderPreset) {
  switch (optionId) {
    case "bottom":
      return BorderBottom01Glyph
    case "top":
      return BorderTop01Glyph
    case "left":
      return BorderLeft01Glyph
    case "right":
      return BorderRight01Glyph
    case "none":
      return BorderNone01Glyph
    case "inside":
      return BorderInnerGlyph
    case "inside-horizontal":
      return BorderHorizontalGlyph
    case "inside-vertical":
      return BorderVerticalGlyph
    case "diagonal-down":
      return ArrowExpandDiagonal01Glyph
    case "diagonal-up":
      return ArrowExpandDiagonal02Glyph
    case "horizontal-line":
      return LineGlyph
    default:
      return BorderAll01Glyph
  }
}

async function loadDocxFile(
  url: string,
  displayFileName: string
): Promise<File> {
  const response = await fetch(url)
  if (!response.ok) {
    throw new Error(`Failed to fetch DOCX (${response.status})`)
  }

  const blob = await response.blob()
  return new File([blob], displayFileName, {
    type: blob.type || DOCX_MIME_TYPE,
  })
}

function formatDocumentName(fileName: string | undefined, url: string) {
  if (fileName?.trim()) return fileName

  const pathname = url.split("?")[0] ?? ""
  const rawName = pathname.split("/").pop() ?? "document.docx"

  try {
    return decodeURIComponent(rawName)
  } catch {
    return rawName
  }
}

function getNextZoomScale(currentZoomScale: number, direction: 1 | -1) {
  const currentIndex = ZOOM_OPTIONS.indexOf(
    currentZoomScale as (typeof ZOOM_OPTIONS)[number]
  )
  let fallbackIndex = -1

  if (direction > 0) {
    fallbackIndex = ZOOM_OPTIONS.findIndex((value) => value > currentZoomScale)
  } else {
    for (let index = ZOOM_OPTIONS.length - 1; index >= 0; index -= 1) {
      if (ZOOM_OPTIONS[index] < currentZoomScale) {
        fallbackIndex = index
        break
      }
    }
  }

  const resolvedIndex = currentIndex >= 0 ? currentIndex : fallbackIndex
  if (resolvedIndex < 0) return currentZoomScale

  const nextIndex = Math.min(
    Math.max(resolvedIndex + direction, 0),
    ZOOM_OPTIONS.length - 1
  )

  return ZOOM_OPTIONS[nextIndex] ?? currentZoomScale
}

function normalizeDocxZoomScale(value: number | undefined): number {
  return typeof value === "number" &&
    ZOOM_OPTIONS.includes(value as (typeof ZOOM_OPTIONS)[number])
    ? value
    : DOCX_EDITOR_DEFAULT_ZOOM_SCALE
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

function isDocxPaddingWarning(args: unknown[]) {
  return (
    typeof args[0] === "string" &&
    args[0].includes(DOCX_PADDING_WARNING_TEXT) &&
    args.some((arg) => String(arg).includes("padding"))
  )
}

function useSuppressDocxPaddingWarning(enabled: boolean) {
  React.useEffect(() => {
    if (!enabled) return

    const originalConsoleError = console.error

    console.error = (...args: unknown[]) => {
      if (isDocxPaddingWarning(args)) return
      originalConsoleError(...args)
    }

    return () => {
      console.error = originalConsoleError
    }
  }, [enabled])
}

function normalizeHexColor(value?: string, fallback = "#111827") {
  if (!value) return fallback

  const trimmed = value.trim()
  const threeDigit = /^#([0-9a-f]{3})$/i.exec(trimmed)
  if (threeDigit?.[1]) {
    const [red, green, blue] = threeDigit[1].split("")
    return `#${red}${red}${green}${green}${blue}${blue}`.toLowerCase()
  }

  const sixDigit = /^#([0-9a-f]{6})$/i.exec(trimmed)
  if (sixDigit?.[1]) {
    return `#${sixDigit[1].toLowerCase()}`
  }

  return fallback
}

function inferParagraphStyleHeadingLevel(value?: string) {
  if (!value) return undefined

  const match = value.match(/(?:^|[\s_-])(?:heading|h)\s*([1-6])(?:$|[\s_-])/i)
  const level = match?.[1] ? Number(match[1]) : NaN

  return level >= 1 && level <= 6 ? (level as 1 | 2 | 3 | 4 | 5 | 6) : undefined
}

function resolveParagraphStyleHeadingLevel(option?: ParagraphStyleDefinition) {
  if (
    Number.isFinite(option?.headingLevel) &&
    option?.headingLevel &&
    option.headingLevel >= 1 &&
    option.headingLevel <= 6
  ) {
    return option.headingLevel as 1 | 2 | 3 | 4 | 5 | 6
  }

  return (
    inferParagraphStyleHeadingLevel(option?.id) ??
    inferParagraphStyleHeadingLevel(option?.name)
  )
}

function resolveParagraphStyleRunPreview(option?: ParagraphStyleDefinition) {
  if (!option) return undefined

  const headingLevel = resolveParagraphStyleHeadingLevel(option)
  const headingRunStyle = headingLevel
    ? DEFAULT_HEADING_PREVIEW_RUN_STYLES[headingLevel]
    : undefined

  return headingRunStyle
    ? {
        ...headingRunStyle,
        ...(option.runStyle ?? {}),
      }
    : option.runStyle
}

function resolveHighlightPreview(value?: string) {
  if (!value) return undefined

  const normalized = value.trim().toLowerCase()
  if (!normalized) return undefined

  const hex = normalizeHexColor(normalized, "")
  if (hex) return hex

  return HIGHLIGHT_PREVIEW_COLORS[normalized]
}

function paragraphStylePreviewStyle(
  option: ParagraphStyleDefinition
): React.CSSProperties {
  const runStyle = resolveParagraphStyleRunPreview(option)
  const textDecoration = [
    runStyle?.underline ? "underline" : "",
    runStyle?.strike ? "line-through" : "",
  ]
    .filter(Boolean)
    .join(" ")

  return {
    textAlign: option.align ?? "left",
    fontFamily: runStyle?.fontFamily,
    fontSize: runStyle?.fontSizePt ? `${runStyle.fontSizePt}pt` : "11pt",
    fontWeight:
      runStyle?.bold !== undefined ? (runStyle.bold ? 700 : 400) : undefined,
    fontStyle: runStyle?.italic ? "italic" : undefined,
    textDecoration: textDecoration || undefined,
    color:
      runStyle?.color !== undefined
        ? normalizeHexColor(runStyle.color)
        : undefined,
    backgroundColor: resolveHighlightPreview(runStyle?.highlight),
    lineHeight: 1,
    whiteSpace: "pre-wrap",
  }
}

function parseToolbarSectionColumns(
  sectionPropertiesXml?: string
): { count: number; gapPx: number } | undefined {
  if (!sectionPropertiesXml) return undefined

  const columnsTag = sectionPropertiesXml.match(/<w:cols\b[^>]*\/?>/i)?.[0]
  if (!columnsTag) return undefined

  const countRaw = columnsTag.match(/\bw:num="(\d+)"/i)?.[1]
  const count = countRaw ? Number(countRaw) : 1
  if (!Number.isFinite(count) || count <= 1) return undefined

  const gapRaw = columnsTag.match(/\bw:space="(\d+)"/i)?.[1]
  const gapTwips = gapRaw ? Number(gapRaw) : 720
  const gapPx = Math.max(0, Math.round((gapTwips * 96) / 1440))

  return {
    count: Math.max(2, Math.round(count)),
    gapPx,
  }
}

function ToolbarTooltip({
  label,
  children,
}: {
  label: string
  children: React.ReactNode
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

function EditorLoadingSurface({
  showSpinner = true,
}: {
  showSpinner?: boolean
}) {
  return (
    <div className="grid h-full min-h-52 place-items-center bg-transparent">
      {showSpinner ? <Spinner className="size-4" /> : null}
    </div>
  )
}

export type DocxEditorMode = "viewing" | "suggesting" | "editing"

// The Google Docs mode menu: what each mode is called, its icon, and the
// viewer mode that runs it.
const DOCX_EDITOR_MODES = [
  {
    value: "editing",
    label: "Editing",
    viewerMode: "edit",
    icon: (
      <IconPlaceholder
        lucide="Pencil"
        tabler="IconPencil"
        hugeicons="PencilEdit01Icon"
        phosphor="PencilSimpleIcon"
        remixicon="RiPencilLine"
        className="size-4"
      />
    ),
  },
  {
    value: "suggesting",
    label: "Suggesting",
    viewerMode: "suggest",
    icon: (
      <IconPlaceholder
        lucide="FilePen"
        tabler="IconFilePencil"
        hugeicons="FilePenIcon"
        phosphor="NotePencilIcon"
        remixicon="RiFileEditLine"
        className="size-4"
      />
    ),
  },
  {
    value: "viewing",
    label: "Viewing",
    viewerMode: "read-only",
    icon: (
      <IconPlaceholder
        lucide="Eye"
        tabler="IconEye"
        hugeicons="ViewIcon"
        phosphor="EyeIcon"
        remixicon="RiEyeLine"
        className="size-4"
      />
    ),
  },
] as const satisfies readonly {
  value: DocxEditorMode
  label: string
  viewerMode: DocxEditorViewerMode
  icon: React.ReactNode
}[]

function docxEditorModeOption(mode: DocxEditorMode) {
  return (
    DOCX_EDITOR_MODES.find((option) => option.value === mode) ??
    DOCX_EDITOR_MODES[0]
  )
}

function DocxModeMenu({
  disabled,
  mode,
  onModeChange,
}: {
  disabled: boolean
  mode: DocxEditorMode
  onModeChange: (mode: DocxEditorMode) => void
}) {
  const current = docxEditorModeOption(mode)

  return (
    <DropdownMenu>
      <DropdownMenuTrigger asChild>
        <Button
          type="button"
          variant="ghost"
          size="sm"
          disabled={disabled}
          aria-label={`Mode: ${current.label}`}
        >
          {current.icon}
          <span className="hidden sm:inline">{current.label}</span>
        </Button>
      </DropdownMenuTrigger>
      <DropdownMenuContent align="end" className="z-40 w-40">
        {DOCX_EDITOR_MODES.map((option) => (
          <DropdownMenuItem
            key={option.value}
            onClick={() => onModeChange(option.value)}
          >
            {option.icon}
            {option.label}
            {option.value === mode ? (
              <IconPlaceholder
                lucide="Check"
                tabler="IconCheck"
                hugeicons="Tick02Icon"
                phosphor="CheckIcon"
                remixicon="RiCheckLine"
                className="ms-auto size-4"
              />
            ) : null}
          </DropdownMenuItem>
        ))}
      </DropdownMenuContent>
    </DropdownMenu>
  )
}

// The review card's actions, at the bottom of the card.
function TrackedChangeCardActions({
  accept,
  reject,
}: Pick<DocxTrackedChangeCardRenderProps, "accept" | "reject">) {
  return (
    <div className="flex justify-end gap-0.5">
      <ToolbarIconButton label="Accept" size="icon-xs" onClick={accept}>
        <IconPlaceholder
          lucide="Check"
          tabler="IconCheck"
          hugeicons="Tick02Icon"
          phosphor="CheckIcon"
          remixicon="RiCheckLine"
          className="size-4"
        />
      </ToolbarIconButton>
      <ToolbarIconButton label="Reject" size="icon-xs" onClick={reject}>
        <IconPlaceholder
          lucide="X"
          tabler="IconX"
          hugeicons="Cancel01Icon"
          phosphor="XIcon"
          remixicon="RiCloseLine"
          className="size-4"
        />
      </ToolbarIconButton>
    </div>
  )
}

function CommentCardActions({
  comment,
  reply,
  setResolved,
}: Pick<DocxCommentCardRenderProps, "comment" | "reply" | "setResolved">) {
  const [draft, setDraft] = React.useState("")

  if (comment.resolved) {
    return (
      <div className="flex justify-end">
        <Button
          type="button"
          variant="ghost"
          size="xs"
          onClick={() => setResolved(false)}
        >
          Reopen
        </Button>
      </div>
    )
  }
  // Replies belong to the thread: only its root takes a reply or a resolve.
  if (comment.parentId !== undefined) return null

  return (
    <div className="flex items-center gap-1">
      <Input
        size="sm"
        value={draft}
        onChange={(event) => setDraft(event.target.value)}
        onKeyDown={(event) => {
          // Keys typed here belong to the reply, not the document.
          event.stopPropagation()
          if (event.key !== "Enter" || !draft.trim()) return
          event.preventDefault()
          if (reply(draft.trim()).ok) setDraft("")
        }}
        placeholder="Reply"
        aria-label="Reply"
        className="text-xs"
      />
      <Button
        type="button"
        variant="ghost"
        size="xs"
        onClick={() => setResolved(true)}
      >
        Resolve
      </Button>
    </div>
  )
}

function ToolbarIconButton({
  active,
  children,
  label,
  ...props
}: React.ComponentProps<typeof Button> & {
  active?: boolean
  label: string
}) {
  return (
    <ToolbarTooltip label={label}>
      <Button
        type="button"
        variant={active ? "secondary" : "ghost"}
        size="icon-sm"
        aria-label={label}
        data-pressed={active ? "" : undefined}
        {...props}
      >
        {children}
      </Button>
    </ToolbarTooltip>
  )
}

function DocxSidebarThumbnail({
  canvasRef,
  displayFileName,
  hasError,
  isActive,
  isLoading,
  pageNumber,
  pixelHeightPx,
  pixelWidthPx,
  previewAspectRatio,
}: {
  canvasRef: React.RefCallback<HTMLCanvasElement>
  displayFileName: string
  hasError: boolean
  isActive: boolean
  isLoading: boolean
  pageNumber: number
  pixelHeightPx: number
  pixelWidthPx: number
  previewAspectRatio: number
}) {
  return (
    <FileThumbnail
      file={{
        name: `${displayFileName} page ${pageNumber}`,
        type: DOCX_MIME_TYPE,
      }}
      previewAspectRatio={previewAspectRatio}
      previewClassName="rounded-md bg-white"
      previewContent={
        <canvas
          ref={canvasRef}
          width={pixelWidthPx}
          height={pixelHeightPx}
          className="!size-full bg-white object-cover object-top"
        />
      }
      isLoading={isLoading}
      hasError={hasError}
      className={cn(
        "w-[92px] rounded-md border-0 shadow-xs ring-0 transition-shadow duration-150",
        isActive && "shadow-sm"
      )}
    />
  )
}

function ParagraphStylePreviewCard({
  option,
}: {
  option: ParagraphStyleDefinition
}) {
  return (
    <div className="w-[260px] p-3">
      <p className="text-[10px] font-medium tracking-wide text-muted-foreground uppercase">
        Style Preview
      </p>
      <div className="mt-2 rounded-sm border border-zinc-300 bg-white p-2.5 text-zinc-900">
        <p style={paragraphStylePreviewStyle(option)}>{option.name}</p>
        <p
          className="mt-1 text-[11px]"
          style={{
            color: "#6b7280",
            textAlign: option.align ?? "left",
          }}
        >
          The quick brown fox jumps over the lazy dog.
        </p>
      </div>
    </div>
  )
}

function DocxEditorToolbar({
  controlsDisabled,
  editor,
  fileName,
  isDark,
  mode,
  onImageUploadClick,
  onIsDarkChange,
  onModeChange,
  onOpenLinkEditor,
  onToggleSidebar,
  onUploadClick,
  setZoomScale,
  showAppearanceToggle,
  zoomScale,
}: {
  controlsDisabled: boolean
  editor: DocxEditorController
  fileName: string
  isDark: boolean
  mode: DocxEditorMode
  onImageUploadClick: () => void
  onIsDarkChange: (checked: boolean) => void
  onModeChange: (mode: DocxEditorMode) => void
  onOpenLinkEditor: () => void
  onToggleSidebar: () => void
  /** Opens a file picker to replace the document; no button without it. */
  onUploadClick?: () => void
  setZoomScale: React.Dispatch<React.SetStateAction<number>>
  showAppearanceToggle: boolean
  zoomScale: number
}) {
  const { layout: pageLayout } = useDocxPageLayout(editor)
  const { paragraphStyles, selectedParagraphStyleId, setParagraphStyle } =
    useDocxParagraphStyles(editor)
  const { lineSpacing, setLineSpacing } = useDocxLineSpacing(editor)
  const { borderContext, activeBorderPresets, applyBorderPreset } =
    useDocxBorders(editor)
  const { showTrackedChanges, setShowTrackedChanges } =
    useDocxTrackChanges(editor)
  const { showComments, setShowComments } = useDocxComments(editor)
  const selectedRunStyle = editor.selectedRunStyle
  const selectedParagraph = editor.selectedParagraph
  const paragraphStyleOptions: ParagraphStyleDefinition[] =
    paragraphStyles.length > 0
      ? paragraphStyles
      : [...FALLBACK_PARAGRAPH_STYLE_OPTIONS]
  const selectedParagraphStyleValue =
    selectedParagraphStyleId ??
    paragraphStyleOptions.find((option) => "isDefault" in option)?.id ??
    paragraphStyleOptions[0]?.id ??
    "Normal"
  const [paragraphStylePreviewId, setParagraphStylePreviewId] = React.useState<
    string | null
  >(null)
  const selectedLineSpacingValue = React.useMemo(() => {
    const current = Number.isFinite(lineSpacing.multiple)
      ? lineSpacing.multiple
      : 1
    const nearest = LINE_SPACING_OPTIONS.reduce((closest, candidate) => {
      return Math.abs(candidate - current) < Math.abs(closest - current)
        ? candidate
        : closest
    }, LINE_SPACING_OPTIONS[0])
    return String(nearest)
  }, [lineSpacing.multiple])
  const textColorValue = normalizeHexColor(selectedRunStyle?.color)
  const highlightColor =
    HIGHLIGHT_COLORS.find(
      (option) => option.value === selectedRunStyle?.highlight
    )?.color ?? "#fff59d"
  const activeNodeIndex =
    editor.selection.kind === "paragraph"
      ? editor.selection.nodeIndex
      : editor.selection.tableIndex
  const letterheadColumns =
    editor.selection.kind === "paragraph" &&
    paragraphLetterheadFloatSideAtNodeIndex(
      editor.model.nodes,
      editor.selection.nodeIndex
    )
      ? { count: 2, gapPx: 28 }
      : undefined
  const sections = editor.model.metadata.sections ?? []
  const activeSection = sections
    .filter((section) => section.startNodeIndex <= activeNodeIndex)
    .at(-1)
  const activeColumns =
    letterheadColumns ??
    parseToolbarSectionColumns(activeSection?.sectionPropertiesXml) ??
    parseToolbarSectionColumns(editor.model.metadata.sectionPropertiesXml) ??
    pageLayout.columns
  const activeBorderControlOptions = BORDER_CONTROL_OPTIONS.filter(
    (option) => activeBorderPresets[option.id]
  )
  const borderTriggerLabel =
    activeBorderControlOptions.length === 1
      ? activeBorderControlOptions[0].label
      : "Borders"
  const borderTriggerIcon = React.createElement(
    borderControlOptionIcon(
      activeBorderControlOptions.length === 1
        ? activeBorderControlOptions[0].id
        : "all"
    ),
    { className: "size-4" }
  )
  const canEdit = !controlsDisabled && mode === "editing"
  const canReview = !controlsDisabled && mode !== "viewing"
  const canZoomIn = zoomScale < ZOOM_OPTIONS[ZOOM_OPTIONS.length - 1]
  const canZoomOut = zoomScale > ZOOM_OPTIONS[0]
  const fontToggleValues = React.useMemo<DocxFontToggleValue[]>(() => {
    const values: DocxFontToggleValue[] = []

    if (selectedRunStyle?.bold) values.push("bold")
    if (selectedRunStyle?.italic) values.push("italic")
    if (selectedRunStyle?.underline) values.push("underline")
    if (selectedRunStyle?.strike) values.push("strike")
    if (selectedRunStyle?.verticalAlign === "superscript") {
      values.push("superscript")
    }
    if (selectedRunStyle?.verticalAlign === "subscript") {
      values.push("subscript")
    }

    return values
  }, [selectedRunStyle])

  const preserveTextSelection = React.useCallback(
    (event: React.MouseEvent<HTMLElement> | React.PointerEvent<HTMLElement>) =>
      event.preventDefault(),
    []
  )
  const applyFontToggleValues = React.useCallback(
    (nextValues: DocxFontToggleValue[]) => {
      const changedValue =
        nextValues.find((value) => !fontToggleValues.includes(value)) ??
        fontToggleValues.find((value) => !nextValues.includes(value))

      if (changedValue === "bold") {
        editor.toggleBold()
      } else if (changedValue === "italic") {
        editor.toggleItalic()
      } else if (changedValue === "underline") {
        editor.toggleUnderline()
      } else if (changedValue === "strike") {
        editor.toggleStrike()
      } else if (changedValue === "superscript") {
        editor.toggleSuperscript()
      } else if (changedValue === "subscript") {
        editor.toggleSubscript()
      }
    },
    [editor, fontToggleValues]
  )

  const zoomControls = (
    <>
      <ToolbarIconButton
        label="Zoom out"
        disabled={controlsDisabled || !canZoomOut}
        onClick={() =>
          setZoomScale((currentZoomScale) =>
            getNextZoomScale(currentZoomScale, -1)
          )
        }
      >
        <IconPlaceholder
          lucide="CircleMinus"
          tabler="IconCircleMinus"
          hugeicons="MinusSignCircleIcon"
          phosphor="MinusCircleIcon"
          remixicon="RiIndeterminateCircleLine"
          className="size-4"
        />
      </ToolbarIconButton>
      <Select
        value={zoomScale.toString()}
        onValueChange={(value) => setZoomScale(Number(value))}
        disabled={controlsDisabled}
        modal={false}
      >
        <SelectTrigger
          size="sm"
          className="w-[84px] min-w-[84px]"
          aria-label="Zoom level"
        >
          <SelectValue>{Math.round(zoomScale)}%</SelectValue>
        </SelectTrigger>
        <SelectContent
          align="end"
          alignItemWithTrigger={false}
          className="z-40"
        >
          {ZOOM_OPTIONS.map((value) => (
            <SelectItem key={value} value={value.toString()}>
              {value}%
            </SelectItem>
          ))}
        </SelectContent>
      </Select>
      <ToolbarIconButton
        label="Zoom in"
        disabled={controlsDisabled || !canZoomIn}
        onClick={() =>
          setZoomScale((currentZoomScale) =>
            getNextZoomScale(currentZoomScale, 1)
          )
        }
      >
        <IconPlaceholder
          lucide="CirclePlusIcon"
          tabler="IconCirclePlusFilled"
          hugeicons="PlusSignCircleIcon"
          phosphor="PlusCircleIcon"
          remixicon="RiAddCircleFill"
          className="size-4"
        />
      </ToolbarIconButton>
    </>
  )

  return (
    <div className="docx-editor-toolbar bg-background">
      <TooltipProvider>
        <div className="docx-editor-header flex min-h-11 items-center gap-2 overflow-x-auto overflow-y-hidden border-b px-3">
          <ToolbarIconButton
            label="Toggle pages"
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
          </ToolbarIconButton>
          <div className="flex min-w-0 flex-1 items-center gap-2">
            <span className="flex size-7 shrink-0 items-center justify-center rounded-md bg-primary/10 text-primary">
              <IconPlaceholder
                lucide="FileText"
                tabler="IconFileText"
                hugeicons="File01Icon"
                phosphor="FileDocIcon"
                remixicon="RiFileWord2Line"
                className="size-4"
              />
            </span>
            <span className="truncate text-sm font-semibold" title={fileName}>
              {fileName}
            </span>
          </div>
          {mode === "viewing" ? zoomControls : null}
          {showAppearanceToggle ? (
            <>
              <ToolbarIconButton
                label={isDark ? "Use light appearance" : "Use dark appearance"}
                onClick={() => onIsDarkChange(!isDark)}
              >
                {isDark ? (
                  <Sun03Glyph className="size-4" />
                ) : (
                  <Moon02Glyph className="size-4" />
                )}
              </ToolbarIconButton>
              <Separator
                orientation="vertical"
                className="mx-1 h-4 self-center"
              />
            </>
          ) : null}
          <DropdownMenu>
            <DropdownMenuTrigger asChild>
              <Button
                type="button"
                variant={showTrackedChanges ? "secondary" : "ghost"}
                size="icon-sm"
                disabled={controlsDisabled}
                aria-label="Tracked changes"
              >
                <IconPlaceholder
                  lucide="FileDiff"
                  tabler="IconFileDiff"
                  hugeicons="FileDiffIcon"
                  phosphor="GitDiffIcon"
                  remixicon="RiFileHistoryLine"
                  className="size-4"
                />
              </Button>
            </DropdownMenuTrigger>
            <DropdownMenuContent align="end" className="z-40 w-48">
              <DropdownMenuCheckboxItem
                checked={showTrackedChanges}
                onCheckedChange={setShowTrackedChanges}
              >
                Show tracked changes
              </DropdownMenuCheckboxItem>
              {canReview && editor.trackedChanges.length > 0 ? (
                <>
                  <DropdownMenuSeparator />
                  <DropdownMenuItem onClick={editor.acceptAllRevisions}>
                    <IconPlaceholder
                      lucide="Check"
                      tabler="IconCheck"
                      hugeicons="Tick02Icon"
                      phosphor="CheckIcon"
                      remixicon="RiCheckLine"
                      className="size-4"
                    />
                    Accept all
                  </DropdownMenuItem>
                  <DropdownMenuItem onClick={editor.rejectAllRevisions}>
                    <IconPlaceholder
                      lucide="X"
                      tabler="IconX"
                      hugeicons="Cancel01Icon"
                      phosphor="XIcon"
                      remixicon="RiCloseLine"
                      className="size-4"
                    />
                    Reject all
                  </DropdownMenuItem>
                </>
              ) : null}
            </DropdownMenuContent>
          </DropdownMenu>
          <ToolbarIconButton
            label="Show comments"
            active={showComments}
            disabled={controlsDisabled}
            onClick={() => setShowComments(!showComments)}
          >
            <IconPlaceholder
              lucide="MessageSquare"
              tabler="IconMessage"
              hugeicons="Comment01Icon"
              phosphor="ChatIcon"
              remixicon="RiChat1Line"
              className="size-4"
            />
          </ToolbarIconButton>
          <DocxModeMenu
            disabled={controlsDisabled}
            mode={mode}
            onModeChange={onModeChange}
          />
          <Separator orientation="vertical" className="mx-1 h-4 self-center" />
          {onUploadClick ? (
            <ToolbarIconButton
              label="Import DOCX"
              disabled={editor.isImporting}
              onClick={onUploadClick}
            >
              {editor.isImporting ? (
                <Spinner className="size-4" />
              ) : (
                <IconPlaceholder
                  lucide="Upload"
                  tabler="IconUpload"
                  hugeicons="Upload01Icon"
                  phosphor="UploadSimpleIcon"
                  remixicon="RiUpload2Line"
                  className="size-4"
                />
              )}
            </ToolbarIconButton>
          ) : null}
          <ToolbarIconButton
            label="Download DOCX"
            disabled={controlsDisabled || editor.isImporting}
            onClick={editor.exportDocx}
          >
            <IconPlaceholder
              lucide="Download"
              tabler="IconDownload"
              hugeicons="Download01Icon"
              phosphor="DownloadSimpleIcon"
              remixicon="RiDownload2Line"
              className="size-4"
            />
          </ToolbarIconButton>
        </div>

        {mode === "viewing" ? null : (
          <ToolbarOverflow
            className="docx-editor-ribbon min-h-12 border-b px-3 py-2"
            end={zoomControls}
          >
            <div className="flex shrink-0 items-center gap-1">
              <ToolbarIconButton
                label="Undo"
                disabled={!editor.canUndo || mode === "viewing"}
                onClick={editor.undo}
              >
                <IconPlaceholder
                  lucide="Undo2"
                  tabler="IconArrowBackUp"
                  hugeicons="Undo02Icon"
                  phosphor="ArrowUUpLeftIcon"
                  remixicon="RiArrowGoBackLine"
                  className="size-4"
                />
              </ToolbarIconButton>
              <ToolbarIconButton
                label="Redo"
                disabled={!editor.canRedo || mode === "viewing"}
                onClick={editor.redo}
              >
                <IconPlaceholder
                  lucide="Redo2"
                  tabler="IconArrowForwardUp"
                  hugeicons="Redo02Icon"
                  phosphor="ArrowUUpRightIcon"
                  remixicon="RiArrowGoForwardLine"
                  className="size-4"
                />
              </ToolbarIconButton>
            </div>
            <Select
              value={selectedParagraphStyleValue}
              onOpenChange={() => setParagraphStylePreviewId(null)}
              onValueChange={(value) => {
                if (value) {
                  setParagraphStyle(value)
                }
              }}
              disabled={!canEdit}
              modal={false}
            >
              <SelectTrigger
                size="sm"
                className="w-[136px] min-w-[136px]"
                aria-label="Paragraph style"
              >
                <SelectValue placeholder="Style" />
              </SelectTrigger>
              <SelectContent
                align="start"
                alignItemWithTrigger={false}
                className="z-40 min-w-[210px]"
              >
                {paragraphStyleOptions.map((option) => (
                  <PreviewCard
                    key={option.id}
                    open={paragraphStylePreviewId === option.id}
                    onOpenChange={(open) => {
                      setParagraphStylePreviewId((current) =>
                        open
                          ? option.id
                          : current === option.id
                            ? null
                            : current
                      )
                    }}
                    openDelay={0}
                    closeDelay={120}
                  >
                    <SelectItem
                      value={option.id}
                      label={option.name}
                      className="relative min-w-[190px]"
                      onFocus={() => setParagraphStylePreviewId(option.id)}
                      onPointerEnter={() =>
                        setParagraphStylePreviewId(option.id)
                      }
                    >
                      <PreviewCardTrigger asChild>
                        <span className="block w-full truncate">
                          {option.name}
                        </span>
                      </PreviewCardTrigger>
                    </SelectItem>
                    <PreviewCardContent
                      side="right"
                      align="start"
                      sideOffset={10}
                      alignOffset={-4}
                      data-slot="paragraph-style-preview-card"
                      className="origin-(--transform-origin) rounded-lg border bg-popover text-popover-foreground shadow-lg/5 transition-opacity duration-100 outline-none not-dark:bg-clip-padding before:pointer-events-none before:absolute before:inset-0 before:rounded-[calc(var(--radius-lg)-1px)] before:shadow-[0_1px_--theme(--color-black/4%)] data-ending-style:opacity-0 data-starting-style:opacity-0 dark:before:shadow-[0_-1px_--theme(--color-white/6%)]"
                    >
                      <ParagraphStylePreviewCard option={option} />
                    </PreviewCardContent>
                  </PreviewCard>
                ))}
              </SelectContent>
            </Select>

            <Select
              value={selectedRunStyle?.fontFamily ?? "Calibri"}
              onValueChange={(value) => {
                if (value) {
                  editor.setFontFamily(value)
                }
              }}
              disabled={!canEdit}
              modal={false}
            >
              <SelectTrigger
                size="sm"
                className="w-[156px] min-w-[156px]"
                aria-label="Font family"
              >
                <SelectValue placeholder="Font" />
              </SelectTrigger>
              <SelectContent
                align="start"
                alignItemWithTrigger={false}
                className="z-40"
              >
                {FONT_FAMILIES.map((fontFamily) => (
                  <SelectItem key={fontFamily} value={fontFamily}>
                    <span style={{ fontFamily }}>{fontFamily}</span>
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>

            <Select
              value={String(Math.round(selectedRunStyle?.fontSizePt ?? 12))}
              onValueChange={(value) => {
                const nextSize = Number(value)
                if (Number.isFinite(nextSize)) {
                  editor.setFontSize(nextSize)
                }
              }}
              disabled={!canEdit}
              modal={false}
            >
              <SelectTrigger
                size="sm"
                className="w-[86px] min-w-[86px]"
                aria-label="Font size"
              >
                <SelectValue placeholder="Size" />
              </SelectTrigger>
              <SelectContent
                align="start"
                alignItemWithTrigger={false}
                className="z-40"
              >
                {FONT_SIZE_OPTIONS.map((size) => (
                  <SelectItem key={size} value={String(size)}>
                    {size} pt
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>

            <Select
              value={selectedLineSpacingValue}
              onValueChange={(value) => {
                const nextSpacing = Number(value)
                if (Number.isFinite(nextSpacing)) {
                  setLineSpacing(nextSpacing)
                }
              }}
              disabled={!canEdit}
              modal={false}
            >
              <SelectTrigger
                size="sm"
                className="w-[94px] min-w-[94px]"
                aria-label="Line spacing"
              >
                <SelectValue placeholder="Spacing" />
              </SelectTrigger>
              <SelectContent
                align="start"
                alignItemWithTrigger={false}
                className="z-40"
              >
                {LINE_SPACING_OPTIONS.map((spacing) => (
                  <SelectItem key={spacing} value={String(spacing)}>
                    {spacing}x
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
            <ToggleGroup
              className="shrink-0"
              disabled={!canEdit}
              multiple
              spacing="default"
              value={fontToggleValues}
              onMouseDown={preserveTextSelection}
              onPointerDown={preserveTextSelection}
              onValueChange={(value) =>
                applyFontToggleValues(value as DocxFontToggleValue[])
              }
            >
              <ToolbarTooltip label="Bold">
                <ToggleGroupItem aria-label="Bold" size="sm" value="bold">
                  <IconPlaceholder
                    lucide="Bold"
                    tabler="IconBold"
                    hugeicons="TextBoldIcon"
                    phosphor="TextBIcon"
                    remixicon="RiBold"
                    className="size-4"
                  />
                </ToggleGroupItem>
              </ToolbarTooltip>
              <ToolbarTooltip label="Italic">
                <ToggleGroupItem aria-label="Italic" size="sm" value="italic">
                  <IconPlaceholder
                    lucide="Italic"
                    tabler="IconItalic"
                    hugeicons="TextItalicIcon"
                    phosphor="TextItalicIcon"
                    remixicon="RiItalic"
                    className="size-4"
                  />
                </ToggleGroupItem>
              </ToolbarTooltip>
              <ToolbarTooltip label="Underline">
                <ToggleGroupItem
                  aria-label="Underline"
                  size="sm"
                  value="underline"
                >
                  <IconPlaceholder
                    lucide="Underline"
                    tabler="IconUnderline"
                    hugeicons="TextUnderlineIcon"
                    phosphor="TextUnderlineIcon"
                    remixicon="RiUnderline"
                    className="size-4"
                  />
                </ToggleGroupItem>
              </ToolbarTooltip>
              <ToolbarTooltip label="Strikethrough">
                <ToggleGroupItem
                  aria-label="Strikethrough"
                  size="sm"
                  value="strike"
                >
                  <IconPlaceholder
                    lucide="Strikethrough"
                    tabler="IconStrikethrough"
                    hugeicons="TextStrikethroughIcon"
                    phosphor="TextStrikethroughIcon"
                    remixicon="RiStrikethrough"
                    className="size-4"
                  />
                </ToggleGroupItem>
              </ToolbarTooltip>
              <ToolbarTooltip label="Superscript">
                <ToggleGroupItem
                  aria-label="Superscript"
                  size="sm"
                  value="superscript"
                >
                  <IconPlaceholder
                    lucide="Superscript"
                    tabler="IconSuperscript"
                    hugeicons="TextSuperscriptIcon"
                    phosphor="TextSuperscriptIcon"
                    remixicon="RiSuperscript"
                    className="size-4"
                  />
                </ToggleGroupItem>
              </ToolbarTooltip>
              <ToolbarTooltip label="Subscript">
                <ToggleGroupItem
                  aria-label="Subscript"
                  size="sm"
                  value="subscript"
                >
                  <IconPlaceholder
                    lucide="Subscript"
                    tabler="IconSubscript"
                    hugeicons="TextSubscriptIcon"
                    phosphor="TextSubscriptIcon"
                    remixicon="RiSubscript"
                    className="size-4"
                  />
                </ToggleGroupItem>
              </ToolbarTooltip>
            </ToggleGroup>
            <div className="flex shrink-0 items-center gap-1">
              <ColorPicker
                label="Text color"
                icon={TextColorGlyph}
                color={textColorValue}
                disabled={!canEdit}
                onTriggerMouseDown={preserveTextSelection}
                onTriggerPointerDown={preserveTextSelection}
                onChange={(color) =>
                  editor.setTextColor(normalizeHexColor(color))
                }
              />
              <DropdownMenu>
                <DropdownMenuTrigger asChild>
                  <Button
                    type="button"
                    variant="ghost"
                    size="icon-sm"
                    disabled={!canEdit}
                    aria-label="Highlight color"
                    className="relative"
                  >
                    <IconPlaceholder
                      lucide="Highlighter"
                      tabler="IconHighlight"
                      hugeicons="HighlighterIcon"
                      phosphor="HighlighterIcon"
                      remixicon="RiMarkPenLine"
                      className="size-4"
                    />
                    <span
                      className="absolute right-1 bottom-1 h-1 w-4 rounded-full border border-background"
                      style={{ backgroundColor: highlightColor }}
                    />
                  </Button>
                </DropdownMenuTrigger>
                <DropdownMenuContent align="start" className="z-40 w-44">
                  {HIGHLIGHT_COLORS.map((option) => (
                    <DropdownMenuCheckboxItem
                      key={option.value}
                      checked={selectedRunStyle?.highlight === option.value}
                      onCheckedChange={() => editor.setHighlight(option.value)}
                    >
                      <span
                        className="size-3 rounded-full border"
                        style={{ backgroundColor: option.color }}
                      />
                      {option.label}
                    </DropdownMenuCheckboxItem>
                  ))}
                  <DropdownMenuSeparator />
                  <DropdownMenuCheckboxItem
                    checked={!selectedRunStyle?.highlight}
                    onCheckedChange={() => editor.setHighlight(undefined)}
                  >
                    No highlight
                  </DropdownMenuCheckboxItem>
                </DropdownMenuContent>
              </DropdownMenu>
              <ToolbarIconButton
                label="Link"
                disabled={!canEdit}
                onMouseDown={preserveTextSelection}
                onPointerDown={preserveTextSelection}
                onClick={onOpenLinkEditor}
              >
                <IconPlaceholder
                  lucide="Link"
                  tabler="IconLink"
                  hugeicons="Link02Icon"
                  phosphor="LinkIcon"
                  remixicon="RiLink"
                  className="size-4"
                />
              </ToolbarIconButton>
            </div>
            <div className="flex shrink-0 items-center gap-1">
              <ToolbarIconButton
                label="Align left"
                active={
                  selectedParagraph?.style?.align === "left" ||
                  !selectedParagraph?.style?.align
                }
                disabled={!canEdit}
                onMouseDown={preserveTextSelection}
                onPointerDown={preserveTextSelection}
                onClick={() => editor.setAlignment("left")}
              >
                <IconPlaceholder
                  lucide="AlignLeft"
                  tabler="IconAlignLeft"
                  hugeicons="TextAlignLeft01Icon"
                  phosphor="TextAlignLeftIcon"
                  remixicon="RiAlignLeft"
                  className="size-4"
                />
              </ToolbarIconButton>
              <ToolbarIconButton
                label="Align center"
                active={selectedParagraph?.style?.align === "center"}
                disabled={!canEdit}
                onMouseDown={preserveTextSelection}
                onPointerDown={preserveTextSelection}
                onClick={() => editor.setAlignment("center")}
              >
                <IconPlaceholder
                  lucide="AlignCenter"
                  tabler="IconAlignCenter"
                  hugeicons="TextAlignCenterIcon"
                  phosphor="TextAlignCenterIcon"
                  remixicon="RiAlignCenter"
                  className="size-4"
                />
              </ToolbarIconButton>
              <ToolbarIconButton
                label="Align right"
                active={selectedParagraph?.style?.align === "right"}
                disabled={!canEdit}
                onMouseDown={preserveTextSelection}
                onPointerDown={preserveTextSelection}
                onClick={() => editor.setAlignment("right")}
              >
                <IconPlaceholder
                  lucide="AlignRight"
                  tabler="IconAlignRight"
                  hugeicons="TextAlignRight01Icon"
                  phosphor="TextAlignRightIcon"
                  remixicon="RiAlignRight"
                  className="size-4"
                />
              </ToolbarIconButton>
              <ToolbarIconButton
                label="Justify"
                active={selectedParagraph?.style?.align === "justify"}
                disabled={!canEdit}
                onMouseDown={preserveTextSelection}
                onPointerDown={preserveTextSelection}
                onClick={() => editor.setAlignment("justify")}
              >
                <IconPlaceholder
                  lucide="AlignJustify"
                  tabler="IconAlignJustified"
                  hugeicons="TextAlignJustifyLeftIcon"
                  phosphor="TextAlignJustifyIcon"
                  remixicon="RiAlignJustify"
                  className="size-4"
                />
              </ToolbarIconButton>
            </div>
            <div className="flex shrink-0 items-center gap-1">
              <ToolbarIconButton
                label="Bulleted list"
                active={editor.hasUnorderedList}
                disabled={!canEdit}
                onMouseDown={preserveTextSelection}
                onPointerDown={preserveTextSelection}
                onClick={() => editor.toggleList("unordered")}
              >
                <IconPlaceholder
                  lucide="Columns3Icon"
                  tabler="IconLayoutColumns"
                  hugeicons="LeftToRightListBulletIcon"
                  phosphor="ColumnsIcon"
                  remixicon="RiLayoutColumnLine"
                  className="size-4"
                />
              </ToolbarIconButton>
              <ToolbarIconButton
                label="Numbered list"
                active={editor.hasOrderedList}
                disabled={!canEdit}
                onMouseDown={preserveTextSelection}
                onPointerDown={preserveTextSelection}
                onClick={() => editor.toggleList("ordered")}
              >
                <IconPlaceholder
                  lucide="ListOrdered"
                  tabler="IconListNumbers"
                  hugeicons="LeftToRightListNumberIcon"
                  phosphor="ListNumbersIcon"
                  remixicon="RiListOrdered"
                  className="size-4"
                />
              </ToolbarIconButton>
            </div>
            <DropdownMenu>
              <DropdownMenuTrigger asChild>
                <Button
                  type="button"
                  variant="ghost"
                  size="sm"
                  disabled={!canEdit}
                  className="h-7 gap-1 px-2 shadow-none"
                >
                  {borderTriggerIcon}
                  {borderTriggerLabel}
                </Button>
              </DropdownMenuTrigger>
              <DropdownMenuContent align="start" className="z-40 w-56">
                {BORDER_CONTROL_OPTIONS.map((option) => {
                  const enabledForContext =
                    !option.contexts || option.contexts.includes(borderContext)
                  const BorderOptionIcon = borderControlOptionIcon(option.id)

                  return (
                    <React.Fragment key={option.id}>
                      {option.separatorBefore ? (
                        <DropdownMenuSeparator />
                      ) : null}
                      <DropdownMenuCheckboxItem
                        checked={activeBorderPresets[option.id]}
                        disabled={!enabledForContext}
                        onCheckedChange={() => {
                          if (enabledForContext) {
                            applyBorderPreset(option.id)
                          }
                        }}
                      >
                        <BorderOptionIcon className="size-4 text-muted-foreground" />
                        {option.label}
                      </DropdownMenuCheckboxItem>
                    </React.Fragment>
                  )
                })}
              </DropdownMenuContent>
            </DropdownMenu>
            <div className="flex shrink-0 items-center gap-1">
              <ToolbarIconButton
                label="Insert image"
                disabled={!canEdit}
                onClick={onImageUploadClick}
              >
                <IconPlaceholder
                  lucide="ImagePlus"
                  tabler="IconPhotoPlus"
                  hugeicons="ImageAdd01Icon"
                  phosphor="ImageIcon"
                  remixicon="RiImageAddLine"
                  className="size-4"
                />
              </ToolbarIconButton>
              <ToolbarIconButton
                label="Insert table"
                disabled={!canEdit}
                onClick={editor.insertTable}
              >
                <IconPlaceholder
                  lucide="TableIcon"
                  tabler="IconTable"
                  hugeicons="TableIcon"
                  phosphor="TableIcon"
                  remixicon="RiTableLine"
                  className="size-4"
                />
              </ToolbarIconButton>
            </div>
          </ToolbarOverflow>
        )}
      </TooltipProvider>
    </div>
  )
}

export function DocxEditorPreview({
  anchorIdPrefix,
  author,
  className,
  defaultMode = "editing",
  defaultZoomScale = DOCX_EDITOR_DEFAULT_ZOOM_SCALE,
  fileName,
  isDark,
  onIsDarkChange,
  scrollsWindow = false,
  showUpload = true,
  src,
}: {
  /** Gives pages and paragraphs stable DOM ids under this prefix. */
  anchorIdPrefix?: string
  /** Who suggestions and comment replies are by; the engine's default otherwise. */
  author?: string
  className?: string
  /** The mode the document opens in: Viewing, Suggesting or Editing. */
  defaultMode?: DocxEditorMode
  defaultZoomScale?: number
  fileName?: string
  isDark: boolean
  onIsDarkChange: (isDark: boolean) => void
  /**
   * The document scrolls the page instead of a box of its own, and the
   * toolbar sticks to the top: the editor fills the page.
   */
  scrollsWindow?: boolean
  /** Shows a button that replaces the document with a file from the device. */
  showUpload?: boolean
  src?: string
}) {
  return (
    <DocxEditorContent
      key={src}
      anchorIdPrefix={anchorIdPrefix}
      author={author}
      className={className}
      defaultMode={defaultMode}
      defaultZoomScale={defaultZoomScale}
      scrollsWindow={scrollsWindow}
      effectiveIsDark={isDark}
      fileName={fileName}
      setIsDark={onIsDarkChange}
      showAppearanceToggle
      showUpload={showUpload}
      url={src}
    />
  )
}

function DocxEditorContent({
  anchorIdPrefix,
  author,
  className,
  defaultMode,
  defaultZoomScale,
  effectiveIsDark,
  fileName,
  scrollsWindow = false,
  setIsDark,
  showAppearanceToggle,
  showUpload,
  url,
}: {
  anchorIdPrefix?: string
  author?: string
  className?: string
  defaultMode: DocxEditorMode
  defaultZoomScale?: number
  scrollsWindow?: boolean
  effectiveIsDark: boolean
  fileName?: string
  setIsDark: (checked: boolean) => void
  showAppearanceToggle: boolean
  showUpload: boolean
  url?: string
}) {
  const fileInputRef = React.useRef<HTMLInputElement>(null)
  const imageInputRef = React.useRef<HTMLInputElement>(null)
  const viewportRef = React.useRef<HTMLDivElement | null>(null)
  const [viewportElement, setViewportElement] =
    React.useState<HTMLDivElement | null>(null)
  const [viewerShellRef, viewerShellWidth] = useElementWidth<HTMLDivElement>()
  const [uploadedDocxFile, setUploadedDocxFile] =
    React.useState<UploadedDocxFile | null>(null)
  const [activePage, setActivePage] = React.useState(1)
  const [sidebarOpen, setSidebarOpen] = React.useState(false)
  const resolvedDefaultZoomScale = normalizeDocxZoomScale(defaultZoomScale)
  // Null until the reader picks a zoom; until then it is the default, or the
  // zoom that fits the page in a narrower window, as Google Docs opens it.
  const [chosenZoomScale, setChosenZoomScale] = React.useState<number | null>(
    null
  )
  const [loadError, setLoadError] = React.useState<string>()
  const [isLoadingDocument, setIsLoadingDocument] = React.useState(Boolean(url))
  const [mode, setMode] = React.useState(defaultMode)
  const [linkEditorOpen, setLinkEditorOpen] = React.useState(false)
  const [linkDraft, setLinkDraft] = React.useState("")
  const viewerBackgroundColor =
    "color-mix(in oklab, var(--muted) 40%, transparent)"
  const displayFileName = React.useMemo(
    () =>
      uploadedDocxFile?.file.name ??
      (url ? formatDocumentName(fileName, url) : (fileName ?? "document.docx")),
    [fileName, uploadedDocxFile?.file.name, url]
  )
  const editorOptions = React.useMemo(
    () => ({
      author,
      initialFileName: displayFileName,
      // Suggesting shows the changes and comments it is about.
      initialShowComments: defaultMode === "suggesting",
      initialShowTrackedChanges: defaultMode === "suggesting",
    }),
    [author, defaultMode, displayFileName]
  )
  const editor = useDocxEditor(editorOptions)
  const { layout: pageLayout } = useDocxPageLayout(editor)
  const fitZoomScale =
    viewerShellWidth && pageLayout.pageWidthPx
      ? Math.floor(((viewerShellWidth - 32) / pageLayout.pageWidthPx) * 100)
      : resolvedDefaultZoomScale
  const zoomScale =
    chosenZoomScale ??
    Math.max(ZOOM_OPTIONS[0], Math.min(resolvedDefaultZoomScale, fitZoomScale))
  const setZoomScale = React.useCallback<
    React.Dispatch<React.SetStateAction<number>>
  >(
    (value) =>
      setChosenZoomScale(
        typeof value === "function" ? value(zoomScale) : value
      ),
    [zoomScale]
  )
  const { importDocxFile, status } = editor
  const [reportedPageCount, setReportedPageCount] = React.useState(0)
  const thumbnailEditor = React.useMemo<DocxEditorController>(
    () => ({
      ...editor,
      totalPages: Math.max(editor.totalPages, reportedPageCount),
    }),
    [editor, reportedPageCount]
  )
  const thumbnailOptions = React.useMemo(
    () => ({
      pixelRatio: 2,
      resolution: {
        maxHeight: DOCX_THUMBNAIL_WIDTH * 1.35,
        maxWidth: DOCX_THUMBNAIL_WIDTH,
      },
    }),
    []
  )
  const { thumbnails } = useDocxViewerThumbnails(
    thumbnailEditor,
    thumbnailOptions
  )
  const shouldShowDocumentSpinner = useDelayedLoadingIndicator(
    isLoadingDocument,
    DOCX_LOADING_INDICATOR_DELAY_MS
  )
  const loadingState = (
    <EditorLoadingSurface showSpinner={shouldShowDocumentSpinner} />
  )
  const canReview = mode !== "viewing"
  const renderTrackedChangeCard = React.useMemo(
    () =>
      createDocxTrackedChangeCardRenderer(
        editor.documentTheme,
        canReview
          ? ({ accept, reject }) => (
              <TrackedChangeCardActions accept={accept} reject={reject} />
            )
          : undefined
      ),
    [canReview, editor.documentTheme]
  )
  const renderCommentCard = React.useMemo(
    () =>
      createDocxCommentCardRenderer(
        editor.documentTheme,
        canReview
          ? ({ comment, reply, setResolved }) => (
              <CommentCardActions
                comment={comment}
                reply={reply}
                setResolved={setResolved}
              />
            )
          : undefined
      ),
    [canReview, editor.documentTheme]
  )
  const hasDocument = Boolean(url || uploadedDocxFile)
  const sidebarInline = useInlineThumbnailSidebar(viewerShellWidth)
  const pageCount =
    hasDocument && !isLoadingDocument && !loadError
      ? Math.max(1, reportedPageCount || editor.totalPages)
      : 0
  const controlsDisabled =
    !hasDocument || isLoadingDocument || Boolean(loadError)
  const thumbnailSidebarOpen = Boolean(
    sidebarOpen && (pageCount || isLoadingDocument)
  )
  const handlePageCountChange = React.useCallback((nextPageCount: number) => {
    setReportedPageCount(Math.max(1, Math.round(nextPageCount || 1)))
  }, [])
  const setViewportRef = React.useCallback((element: HTMLDivElement | null) => {
    viewportRef.current = element
    setViewportElement(element)
  }, [])
  // The page's scroller exists only in the browser: it is read after mount,
  // so rendering on the server never touches document.
  const [pageScrollElement, setPageScrollElement] =
    React.useState<HTMLElement | null>(null)
  React.useEffect(() => {
    setPageScrollElement(
      scrollsWindow ? (document.scrollingElement as HTMLElement | null) : null
    )
  }, [scrollsWindow])
  const pageVirtualization = React.useMemo(
    () => ({
      enabled: true,
      overscan: 1,
      scrollElement: scrollsWindow ? pageScrollElement : viewportElement,
      zoomScale: zoomScale / 100,
    }),
    [pageScrollElement, scrollsWindow, viewportElement, zoomScale]
  )
  // scrollViewportTo scrolls whatever holds the document: the page, or the
  // viewer's own box.
  const scrollViewportTo = React.useCallback(
    (top: number) => {
      if (scrollsWindow) window.scrollTo({ top, left: 0, behavior: "auto" })
      else viewportRef.current?.scrollTo({ top, left: 0, behavior: "auto" })
    },
    [scrollsWindow]
  )

  useSuppressDocxPaddingWarning(!isLoadingDocument && !loadError)

  const [previousDefaultZoom, setPreviousDefaultZoom] = React.useState(
    resolvedDefaultZoomScale
  )
  if (previousDefaultZoom !== resolvedDefaultZoomScale) {
    setPreviousDefaultZoom(resolvedDefaultZoomScale)
    setChosenZoomScale(null)
    setActivePage(1)
  }

  React.useEffect(() => {
    scrollViewportTo(0)
  }, [resolvedDefaultZoomScale, scrollViewportTo, url])

  // Suggesting refuses edits it cannot track; the status says why, briefly.
  const [seenRefusalNonce, setSeenRefusalNonce] = React.useState(
    editor.suggestionRefusalNonce
  )
  const [refusalMessage, setRefusalMessage] = React.useState<string>()
  if (seenRefusalNonce !== editor.suggestionRefusalNonce) {
    setSeenRefusalNonce(editor.suggestionRefusalNonce)
    setRefusalMessage(status)
  }
  React.useEffect(() => {
    if (!refusalMessage) return
    const timeout = window.setTimeout(() => setRefusalMessage(undefined), 4000)
    return () => window.clearTimeout(timeout)
  }, [refusalMessage, seenRefusalNonce])

  function changeMode(nextMode: DocxEditorMode) {
    setMode(nextMode)
    if (nextMode === "suggesting") {
      editor.setShowTrackedChanges(true)
      editor.setShowComments(true)
    }
  }

  const [previousStatus, setPreviousStatus] = React.useState(status)
  if (previousStatus !== status) {
    setPreviousStatus(status)
    if (
      status.startsWith("Failed to load file") ||
      status === "Only .docx files are supported"
    ) {
      setLoadError(status)
      setIsLoadingDocument(false)
    }
  }

  React.useEffect(() => {
    let isCurrent = true

    async function load() {
      if (!uploadedDocxFile && !url) return

      try {
        const docxFile =
          uploadedDocxFile?.file ??
          (url ? await loadDocxFile(url, displayFileName) : null)
        if (!docxFile) return
        await importDocxFile(docxFile)

        if (isCurrent) {
          setIsLoadingDocument(false)
          setActivePage(1)
          scrollViewportTo(0)
        }
      } catch (error) {
        if (isCurrent) {
          setLoadError(
            error instanceof Error ? error.message : "Unknown DOCX load error"
          )
          setIsLoadingDocument(false)
        }
      }
    }

    void load()

    return () => {
      isCurrent = false
    }
  }, [displayFileName, importDocxFile, scrollViewportTo, uploadedDocxFile, url])

  const updateActivePageFromViewport = React.useCallback(() => {
    const viewport = viewportRef.current
    if (!viewport || !pageCount) return

    const viewportRect = scrollsWindow
      ? { top: 0, height: window.innerHeight }
      : viewport.getBoundingClientRect()
    const viewportCenter = viewportRect.top + viewportRect.height / 2
    let closestPage = 1
    let closestDistance = Number.POSITIVE_INFINITY

    viewport
      .querySelectorAll<HTMLElement>(
        '[data-docx-page-wrapper="true"][data-index]'
      )
      .forEach((page) => {
        const pageIndex = Number(page.dataset.index)
        if (!Number.isFinite(pageIndex)) return

        const pageRect = page.getBoundingClientRect()
        const pageCenter = pageRect.top + pageRect.height / 2
        const distance = Math.abs(pageCenter - viewportCenter)

        if (distance < closestDistance) {
          closestDistance = distance
          closestPage = pageIndex + 1
        }
      })

    setActivePage((currentPage) =>
      currentPage === closestPage ? currentPage : closestPage
    )
  }, [pageCount, scrollsWindow])

  React.useEffect(() => {
    const viewport = viewportRef.current
    if (!viewport || !pageCount) return
    const scroller = scrollsWindow ? window : viewport

    let frameId = 0
    const handleScroll = () => {
      window.cancelAnimationFrame(frameId)
      frameId = window.requestAnimationFrame(updateActivePageFromViewport)
    }

    frameId = window.requestAnimationFrame(updateActivePageFromViewport)
    scroller.addEventListener("scroll", handleScroll, { passive: true })

    return () => {
      window.cancelAnimationFrame(frameId)
      scroller.removeEventListener("scroll", handleScroll)
    }
  }, [pageCount, scrollsWindow, updateActivePageFromViewport])

  const scrollToPage = React.useCallback(
    (pageNumber: number) => {
      const viewport = viewportRef.current
      const targetPageIndex = pageNumber - 1
      const page = viewport?.querySelector<HTMLElement>(
        `[data-docx-page-wrapper="true"][data-index="${targetPageIndex}"]`
      )

      setActivePage(pageNumber)

      if (!viewport) return

      // The document's top, in the scroller's coordinates; on the page the
      // sticky toolbar covers its top, which scroll-padding-top accounts for.
      const scrollerTop = scrollsWindow
        ? -window.scrollY +
          (parseFloat(
            getComputedStyle(document.documentElement).scrollPaddingTop
          ) || 0)
        : viewport.getBoundingClientRect().top - viewport.scrollTop

      if (!page) {
        const pageStridePx =
          (pageLayout.pageHeightPx + pageLayout.viewportDefaults.pageGapPx) *
          (zoomScale / 100)
        const documentTop = scrollsWindow
          ? viewport.getBoundingClientRect().top - scrollerTop
          : 0

        scrollViewportTo(
          Math.max(0, documentTop + targetPageIndex * pageStridePx - 24)
        )
        return
      }

      scrollViewportTo(page.getBoundingClientRect().top - scrollerTop - 24)
    },
    [
      pageLayout.pageHeightPx,
      pageLayout.viewportDefaults.pageGapPx,
      scrollViewportTo,
      scrollsWindow,
      zoomScale,
    ]
  )

  async function handleUpload(event: React.ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0]
    event.target.value = ""

    if (!file) return

    setChosenZoomScale(null)
    setActivePage(1)
    setReportedPageCount(0)
    setIsLoadingDocument(true)
    setLoadError(undefined)
    setUploadedDocxFile({
      file,
      identity: `${file.name}-${file.size}-${file.lastModified}`,
    })
  }

  async function handleImageUpload(event: React.ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0]
    event.target.value = ""

    if (!file) return

    await editor.insertImageFile(file)
  }

  function openLinkEditor() {
    setLinkDraft(editor.selectedLink ?? "")
    setLinkEditorOpen(true)
  }

  // Cmd/Ctrl+K in the page opens the link editor, as in Docs and Word.
  const openLinkEditorFromShortcut = React.useEffectEvent(openLinkEditor)
  React.useEffect(() => {
    if (mode !== "editing") return
    const handleKeyDown = (event: KeyboardEvent) => {
      const target = event.target as Element | null
      if (
        event.key.toLowerCase() !== "k" ||
        !(event.metaKey || event.ctrlKey) ||
        event.altKey ||
        event.shiftKey ||
        !target?.closest("[data-docx-page-wrapper]")
      )
        return
      event.preventDefault()
      openLinkEditorFromShortcut()
    }
    document.addEventListener("keydown", handleKeyDown)
    return () => document.removeEventListener("keydown", handleKeyDown)
  }, [mode])

  function applyLink() {
    editor.setLink(linkDraft.trim() || undefined)
    setLinkEditorOpen(false)
  }

  return (
    <div
      className={cn(
        scrollsWindow
          ? "flex min-h-dvh flex-col bg-background"
          : "flex h-[720px] min-h-0 flex-col overflow-hidden bg-background",
        className
      )}
    >
      <input
        ref={fileInputRef}
        type="file"
        accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        className="hidden"
        onChange={handleUpload}
      />
      <input
        ref={imageInputRef}
        type="file"
        accept="image/*"
        className="hidden"
        onChange={handleImageUpload}
      />
      <div className={scrollsWindow ? "sticky top-0 z-30" : undefined}>
        <DocxEditorToolbar
          controlsDisabled={controlsDisabled}
          editor={editor}
          fileName={displayFileName}
          isDark={effectiveIsDark}
          mode={mode}
          onImageUploadClick={() => imageInputRef.current?.click()}
          onIsDarkChange={setIsDark}
          onModeChange={changeMode}
          onOpenLinkEditor={openLinkEditor}
          onToggleSidebar={() => setSidebarOpen((open) => !open)}
          onUploadClick={
            showUpload ? () => fileInputRef.current?.click() : undefined
          }
          setZoomScale={setZoomScale}
          showAppearanceToggle={showAppearanceToggle}
          zoomScale={zoomScale}
        />
      </div>
      <div
        ref={viewerShellRef}
        className={cn(
          "relative flex bg-muted/30",
          scrollsWindow ? "flex-1" : "min-h-0 flex-1 overflow-hidden"
        )}
      >
        {linkEditorOpen ? (
          <div className="absolute top-3 left-1/2 z-50 flex w-[min(420px,calc(100%-2rem))] -translate-x-1/2 items-center gap-2 rounded-lg border bg-background p-2 shadow-lg">
            <Input
              value={linkDraft}
              onChange={(event) => setLinkDraft(event.target.value)}
              onKeyDown={(event) => {
                if (event.key === "Enter") {
                  event.preventDefault()
                  applyLink()
                }
                if (event.key === "Escape") {
                  setLinkEditorOpen(false)
                }
              }}
              placeholder="https://example.com"
              aria-label="Link URL"
            />
            <Button type="button" size="sm" onClick={applyLink}>
              Apply
            </Button>
            <Button
              type="button"
              size="sm"
              variant="ghost"
              onClick={() => {
                editor.setLink(undefined)
                setLinkEditorOpen(false)
              }}
            >
              Clear
            </Button>
          </div>
        ) : null}
        {refusalMessage ? (
          <div
            role="status"
            className="fixed bottom-4 left-1/2 z-50 w-max max-w-[calc(100%-2rem)] -translate-x-1/2 rounded-lg border bg-background px-3 py-2 text-xs text-muted-foreground shadow-lg"
          >
            {refusalMessage}
          </div>
        ) : null}
        <DocumentViewerThumbnailSidebar
          inline={sidebarInline}
          open={thumbnailSidebarOpen}
        >
          <ScrollArea
            className="h-full"
            scrollFade
            viewportClassName="overscroll-contain"
          >
            <div className="p-4">
              {isLoadingDocument ? (
                <>
                  <div className="mx-auto h-28 w-20 overflow-hidden rounded-md bg-background shadow-xs">
                    <div className="h-full animate-pulse bg-muted" />
                  </div>
                  <div className="mx-auto mt-3 h-3 w-10 rounded-full bg-muted" />
                </>
              ) : (
                <div className="flex flex-col items-center gap-3">
                  {thumbnails.slice(0, pageCount || 0).map((thumbnail) => (
                    <Button
                      key={thumbnail.pageIndex}
                      type="button"
                      variant="ghost"
                      size="sm"
                      className={cn(
                        "!h-auto w-full flex-col items-center gap-2 p-2 text-xs shadow-none hover:bg-sidebar-accent",
                        thumbnail.pageNumber === activePage &&
                          "bg-sidebar-accent text-foreground",
                        thumbnail.pageNumber !== activePage &&
                          "text-muted-foreground"
                      )}
                      onFocus={(event) => event.currentTarget.blur()}
                      onClick={() => scrollToPage(thumbnail.pageNumber)}
                    >
                      <DocxSidebarThumbnail
                        canvasRef={thumbnail.canvasRef}
                        displayFileName={displayFileName}
                        hasError={thumbnail.status === "error"}
                        isActive={thumbnail.pageNumber === activePage}
                        isLoading={
                          !thumbnail.isMounted &&
                          thumbnail.status !== "ready" &&
                          thumbnail.status !== "error"
                        }
                        pageNumber={thumbnail.pageNumber}
                        pixelHeightPx={thumbnail.pixelHeightPx}
                        pixelWidthPx={thumbnail.pixelWidthPx}
                        previewAspectRatio={thumbnail.aspectRatio}
                      />
                      {thumbnail.pageNumber}
                    </Button>
                  ))}
                </div>
              )}
            </div>
          </ScrollArea>
        </DocumentViewerThumbnailSidebar>
        <DocumentScroller
          scrollsWindow={scrollsWindow}
          style={{ backgroundColor: viewerBackgroundColor }}
          viewportRef={setViewportRef}
        >
          {!url && !uploadedDocxFile ? (
            <div className="grid h-full min-h-96 place-items-center p-6 text-center">
              <div className="max-w-md rounded-lg border bg-background p-4 text-sm shadow-xs">
                <div className="font-medium">Upload a DOCX to edit</div>
                <div className="mt-1 text-muted-foreground">
                  Pass a DOCX URL with the <code>src</code> prop or upload a
                  file.
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
                  Upload DOCX
                </Button>
              </div>
            </div>
          ) : loadError ? (
            <div className="grid h-full min-h-96 place-items-center p-6 text-center">
              <div className="max-w-md rounded-lg border bg-background p-4 text-sm text-destructive shadow-xs">
                <div className="font-medium">Unable to edit DOCX</div>
                <div className="mt-1 text-muted-foreground">{loadError}</div>
              </div>
            </div>
          ) : isLoadingDocument ? (
            loadingState
          ) : (
            <div className="mx-auto flex min-h-full justify-center">
              <div className="origin-top" style={{ zoom: zoomScale / 100 }}>
                <DocxEditorViewer
                  editor={editor}
                  anchorIdPrefix={anchorIdPrefix}
                  mode={docxEditorModeOption(mode).viewerMode}
                  showTrackedChanges={editor.showTrackedChanges}
                  renderTrackedChangeCard={renderTrackedChangeCard}
                  showComments={editor.showComments}
                  renderCommentCard={renderCommentCard}
                  loadingState={loadingState}
                  pageGapBackgroundColor={viewerBackgroundColor}
                  pageVirtualization={pageVirtualization}
                  deferInitialPaginationPaint={false}
                  onPageCountChange={handlePageCountChange}
                />
              </div>
            </div>
          )}
        </DocumentScroller>
      </div>
    </div>
  )
}

// DocumentScroller holds the pages: a scroll box of its own, or, when the
// document scrolls the page, a plain block that grows with it.
function DocumentScroller({
  children,
  scrollsWindow,
  style,
  viewportRef,
}: {
  children: React.ReactNode
  scrollsWindow: boolean
  style: React.CSSProperties
  viewportRef: (element: HTMLDivElement | null) => void
}) {
  if (scrollsWindow) {
    return (
      <div ref={viewportRef} className="min-w-0 flex-1 p-4" style={style}>
        {children}
      </div>
    )
  }
  return (
    <ScrollArea
      className="min-h-0 flex-1"
      style={style}
      viewportClassName="overscroll-contain p-4"
      viewportRef={viewportRef}
    >
      {children}
    </ScrollArea>
  )
}
