"use client"

import * as React from "react"

import { cn } from "@/lib/utils"
import { Button } from "@/components/ui/button"
import { Popover, PopoverPopup, PopoverTrigger } from "@/components/ui/popover"
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
        // Each separator occupies one pixel and has a gap on both sides.
        const width =
          (widthsRef.current[index] ?? 0) + (index > 0 ? 2 * gap + 1 : 0)
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

  return (
    <div className={cn("relative flex items-center gap-2", className)}>
      <div ref={rowRef} className="flex min-w-0 flex-1 items-center gap-2">
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
          <Popover open={menuOpen} onOpenChange={setMenuOpen}>
            <PopoverTrigger
              render={
                <Button
                  type="button"
                  variant={menuOpen ? "secondary" : "ghost"}
                  size="icon-sm"
                  className="shrink-0"
                  aria-label="More"
                />
              }
            >
              <IconPlaceholder
                lucide="EllipsisVertical"
                tabler="IconDotsVertical"
                hugeicons="MoreVerticalIcon"
                phosphor="DotsThreeVerticalIcon"
                remixicon="RiMore2Line"
                className="size-4"
              />
            </PopoverTrigger>
            <PopoverPopup
              align="end"
              aria-label="More tools"
              className="[&>[data-slot=popover-viewport]]:p-2"
            >
              {/* The viewport wraps its children in its own element, so the
                  row is a child of it rather than the viewport itself. */}
              <div className="flex w-max max-w-[min(32rem,calc(100vw-2rem))] flex-wrap items-center gap-2">
                {hidden.map((group, index) => (
                  <div key={index} className="flex items-center gap-1">
                    {group}
                  </div>
                ))}
              </div>
            </PopoverPopup>
          </Popover>
        ) : null}
      </div>
      {end ? (
        <div className="flex flex-none items-center gap-1">{end}</div>
      ) : null}
    </div>
  )
}
