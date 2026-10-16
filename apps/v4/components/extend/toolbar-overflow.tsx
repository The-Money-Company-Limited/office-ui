"use client"

import * as React from "react"

import { cn } from "@/lib/utils"
import { Button } from "@/components/ui/button"
import { Separator } from "@/components/ui/separator"
import { IconPlaceholder } from "@/components/icon-placeholder"

const MORE_BUTTON_WIDTH_PX = 32

/**
 * One toolbar row, as in Google Docs: groups that do not fit move, in order,
 * into a More menu at the end of the row instead of wrapping onto more rows.
 * `end` stays at the right edge (zoom, for example).
 */
export function ToolbarOverflow({
  children,
  className,
  end,
}: {
  children: React.ReactNode
  className?: string
  end?: React.ReactNode
}) {
  const groups = React.Children.toArray(children)
  const rowRef = React.useRef<HTMLDivElement>(null)
  const widthsRef = React.useRef<number[]>([])
  const [visibleCount, setVisibleCount] = React.useState(groups.length)
  const [menuOpen, setMenuOpen] = React.useState(false)
  const menuRef = React.useRef<HTMLDivElement>(null)

  React.useLayoutEffect(() => {
    const row = rowRef.current
    if (!row) return

    // Widths are measured while a group is in the row and remembered while it
    // is in the menu, so the row never has to render every group to refit.
    const fit = () => {
      row
        .querySelectorAll<HTMLElement>(":scope > [data-toolbar-group]")
        .forEach((element) => {
          widthsRef.current[Number(element.dataset.toolbarGroup)] =
            element.offsetWidth
        })
      const gap = parseFloat(getComputedStyle(row).columnGap) || 0
      let used = 0
      let count = 0
      for (let index = 0; index < groups.length; index += 1) {
        const width = (widthsRef.current[index] ?? 0) + (index > 0 ? gap : 0)
        const moreWidth =
          index < groups.length - 1 ? gap + MORE_BUTTON_WIDTH_PX : 0
        if (used + width + moreWidth > row.clientWidth) break
        used += width
        count += 1
      }
      setVisibleCount(count)
    }

    fit()
    const observer = new ResizeObserver(fit)
    observer.observe(row)
    return () => observer.disconnect()
  }, [groups.length])

  const hidden = groups.slice(visibleCount)

  React.useEffect(() => {
    if (!menuOpen) return
    // The menu's own pickers open in portals, so only a press on the
    // toolbar's page area outside the menu closes it.
    const close = (event: PointerEvent) => {
      const target = event.target as Element | null
      if (menuRef.current?.contains(target)) return
      if (
        target?.closest(
          "[data-slot$='-positioner'], [role='listbox'], [role='menu']"
        )
      )
        return
      setMenuOpen(false)
    }
    const closeOnEscape = (event: KeyboardEvent) => {
      if (event.key === "Escape") setMenuOpen(false)
    }
    document.addEventListener("pointerdown", close)
    document.addEventListener("keydown", closeOnEscape)
    return () => {
      document.removeEventListener("pointerdown", close)
      document.removeEventListener("keydown", closeOnEscape)
    }
  }, [menuOpen])

  return (
    <div className={cn("relative flex items-center gap-2", className)}>
      <div
        ref={rowRef}
        className="flex min-w-0 flex-1 items-center gap-2 overflow-hidden"
      >
        {groups.slice(0, visibleCount).map((group, index) => (
          <React.Fragment key={index}>
            {index > 0 ? (
              <Separator
                orientation="vertical"
                className="h-4 shrink-0 self-center"
              />
            ) : null}
            <div
              data-toolbar-group={index}
              className="flex shrink-0 items-center gap-1"
            >
              {group}
            </div>
          </React.Fragment>
        ))}
        {hidden.length > 0 ? (
          <div ref={menuRef} className="shrink-0">
            <Button
              type="button"
              variant={menuOpen ? "secondary" : "ghost"}
              size="icon-sm"
              aria-label="More"
              aria-expanded={menuOpen}
              onClick={() => setMenuOpen((open) => !open)}
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
            {menuOpen ? (
              <div className="absolute top-full right-2 z-30 mt-1 flex max-w-[min(32rem,calc(100vw-1rem))] flex-wrap items-center gap-2 rounded-lg border bg-popover p-2 text-popover-foreground shadow-lg">
                {hidden.map((group, index) => (
                  <div key={index} className="flex items-center gap-1">
                    {group}
                  </div>
                ))}
              </div>
            ) : null}
          </div>
        ) : null}
      </div>
      {end ? (
        <div className="flex flex-none items-center gap-1">{end}</div>
      ) : null}
    </div>
  )
}
