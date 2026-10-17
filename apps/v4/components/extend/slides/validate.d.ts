export const SLIDE_WIDTH: 1920
export const SLIDE_HEIGHT: 1080

export type DeckManifest = {
  title: string
  order: string[]
  sections?: { title: string; slides: string[] }[]
}

export type Deck = {
  manifest: DeckManifest
  slides: Record<string, string>
}

export function validateManifest(value: unknown): DeckManifest
export function sanitizeSlide(
  source: string,
  slideId: string,
  dom: Document,
  warn?: (
    message: string,
    diagnostic: { element: string; text: string; detail: string }
  ) => void
): string
export function loadDeck(src: string, signal: AbortSignal): Promise<Deck>
