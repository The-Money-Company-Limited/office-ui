"use client"

import * as React from "react"
import { getRegionalFormatSamples, useXlsxViewer } from "@extend-ai/react-xlsx"

import { Button } from "@/components/ui/button"
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuGroup,
  DropdownMenuLabel,
  DropdownMenuRadioGroup,
  DropdownMenuRadioItem,
  DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu"
import { IconPlaceholder } from "@/components/icon-placeholder"

const STORAGE_KEY = "xlsx-editor-regional-format"
/** Regional formats the engine knows the dates, numbers and currency of. */
const REGIONAL_FORMATS = [
  "en-US",
  "en-GB",
  "en-CA",
  "en-AU",
  "en-IN",
  "en-IE",
  "de-DE",
  "de-CH",
  "fr-FR",
  "fr-CA",
  "es-ES",
  "es-MX",
  "it-IT",
  "pt-BR",
  "pt-PT",
  "nl-NL",
  "sv-SE",
  "da-DK",
  "nb-NO",
  "pl-PL",
  "cs-CZ",
  "ja-JP",
  "zh-CN",
  "ko-KR",
]

function browserRegionalFormat() {
  return (typeof navigator === "undefined" ? "" : navigator.language) || "en-US"
}

/**
 * The viewer's regional format: the one they picked (kept in this browser), or the browser's
 * language. Picking the browser's own forgets the choice, so it follows the browser again.
 */
export function useRegionalFormat() {
  const [regionalFormat, setRegionalFormatState] = React.useState(() => {
    try {
      return window.localStorage.getItem(STORAGE_KEY) ?? browserRegionalFormat()
    } catch {
      return browserRegionalFormat()
    }
  })
  const setRegionalFormat = React.useCallback((tag: string) => {
    setRegionalFormatState(tag)
    try {
      if (tag === browserRegionalFormat())
        window.localStorage.removeItem(STORAGE_KEY)
      else window.localStorage.setItem(STORAGE_KEY, tag)
    } catch {
      // Without storage the choice lasts until the page closes.
    }
  }, [])
  return [regionalFormat, setRegionalFormat] as const
}

const RegionalFormatContext = React.createContext<
  ((tag: string) => void) | null
>(null)

/** Lets the Regional format menu change the format the editor passes to its controller. */
export const RegionalFormatProvider = RegionalFormatContext.Provider

let displayNames: Intl.DisplayNames | null | undefined
function regionalFormatName(tag: string) {
  try {
    displayNames ??= new Intl.DisplayNames(["en"], { type: "language" })
    return displayNames.of(tag) ?? tag
  } catch {
    return tag
  }
}

/**
 * The regional format dates and numbers are typed and shown in, as Excel takes it from the
 * computer, with a menu to pick another. A choice re-applies to the open workbook.
 */
export function EditorRegionalFormat() {
  const setRegionalFormat = React.useContext(RegionalFormatContext)
  const { readOnly, regionalFormat } = useXlsxViewer()
  const browserFormat = browserRegionalFormat()
  const formats = React.useMemo(
    () =>
      // The browser's own first; one picked elsewhere (an older choice) last.
      [browserFormat, ...REGIONAL_FORMATS, regionalFormat].filter(
        (tag, index, tags) =>
          tags.findIndex(
            (other) => other.toLowerCase() === tag.toLowerCase()
          ) === index
      ),
    [browserFormat, regionalFormat]
  )
  const [samples, setSamples] = React.useState<
    Record<string, [string, string]>
  >({})
  if (!setRegionalFormat) return null
  const name = regionalFormatName(regionalFormat)
  return (
    <DropdownMenu
      onOpenChange={(open) => {
        // The samples come from the engine, which a read-only workbook only runs in a worker.
        if (open)
          void getRegionalFormatSamples(formats)
            .then(setSamples)
            .catch(() => {})
      }}
    >
      <DropdownMenuTrigger asChild>
        <Button
          type="button"
          variant="ghost"
          size="sm"
          title="Regional format"
          aria-label={`Regional format: ${name}`}
          className="text-muted-foreground"
        >
          <IconPlaceholder
            lucide="Globe"
            tabler="IconWorld"
            hugeicons="Globe02Icon"
            phosphor="GlobeIcon"
            remixicon="RiGlobalLine"
            className="size-3.5"
          />
          <span className="hidden max-w-40 truncate sm:inline">{name}</span>
        </Button>
      </DropdownMenuTrigger>
      <DropdownMenuContent
        align="end"
        side="top"
        className="z-40 max-h-80 w-80 overflow-y-auto"
      >
        <DropdownMenuGroup>
          <DropdownMenuLabel>
            {readOnly
              ? "Regional format: how dates and numbers are shown"
              : "Regional format: how dates and numbers are typed and shown"}
          </DropdownMenuLabel>
          <DropdownMenuRadioGroup
            value={regionalFormat}
            onValueChange={(tag: string) => setRegionalFormat(tag)}
          >
            {formats.map((tag) => (
              <DropdownMenuRadioItem key={tag} value={tag} closeOnClick>
                <span className="min-w-0 flex-1 truncate">
                  {regionalFormatName(tag)}
                  {tag === browserFormat ? (
                    <span className="text-muted-foreground"> · Browser</span>
                  ) : null}
                </span>
                <span className="shrink-0 text-xs text-muted-foreground tabular-nums">
                  {samples[tag]?.join(" · ")}
                </span>
              </DropdownMenuRadioItem>
            ))}
          </DropdownMenuRadioGroup>
        </DropdownMenuGroup>
      </DropdownMenuContent>
    </DropdownMenu>
  )
}
