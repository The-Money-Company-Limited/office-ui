import type { TextStyle } from "@extend-ai/react-docx-doc-model";

export type FontKerningMode = "auto" | "normal" | "none";

export function resolveDocxFontKerning(
  style: TextStyle | undefined,
  fontSizePt: number
): "normal" | "none";
export function resolveDocxFontKerning(
  style: TextStyle | undefined,
  fontSizePt?: number
): "normal" | "none" | undefined;
export function resolveDocxFontKerning(
  style: TextStyle | undefined,
  fontSizePt = style?.fontSizePt
): "normal" | "none" | undefined {
  const thresholdPt = style?.kerningMinFontSizePt;
  if (!Number.isFinite(thresholdPt) || (thresholdPt as number) <= 0) return "none";
  if (!Number.isFinite(fontSizePt)) return undefined;
  return (fontSizePt as number) >= (thresholdPt as number) ? "normal" : "none";
}

export function measuredFontSizePt(font: string | undefined): number | undefined {
  const size = font?.match(/(?:^|\s)([+]?(?:\d+(?:\.\d*)?|\.\d+))(px|pt)(?=\s|\/|$)/i);
  if (!size) return undefined;
  const value = Number(size[1]) * (size[2]?.toLowerCase() === "px" ? 72 / 96 : 1);
  return Number.isFinite(value) && value > 0 ? value : undefined;
}
