import type { Box } from "./geometry"
import { SLIDE_WIDTH } from "./validate.js"

export function chartPanelLayout(
  box: Box,
  scale: number,
  viewportWidth: number
) {
  const slideLeft = 152 + (viewportWidth - 152 - SLIDE_WIDTH * scale) / 2
  const leftRoom = slideLeft + box.x * scale - 164
  const rightRoom = viewportWidth - 12 - slideLeft - (box.x + box.width) * scale
  const sidebar = Math.max(leftRoom, rightRoom) < 332
  return {
    sidebar,
    side: !sidebar && leftRoom > rightRoom ? "left" : "right",
    width: viewportWidth < 760 ? Math.min(320, viewportWidth * 0.45) : 320,
    compact: sidebar && viewportWidth < 760,
  } as const
}
