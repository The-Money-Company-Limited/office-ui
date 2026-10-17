interface BodyRect {
  left: number;
  top: number;
  right: number;
  bottom: number;
  width: number;
  height: number;
}

type ClipBounds = Pick<BodyRect, "left" | "top" | "right" | "bottom">;
const CLIPPING_OVERFLOW_VALUES = new Set(["hidden", "clip", "scroll", "auto"]);

export function createClippedBodyRectReader(
  body: HTMLElement
): (element: HTMLElement) => BodyRect {
  const rectangles = new WeakMap<HTMLElement, DOMRect>();
  const childClips = new WeakMap<HTMLElement, ClipBounds>();
  const unrestricted: ClipBounds = {
    left: -Infinity, top: -Infinity, right: Infinity, bottom: Infinity,
  };
  const readRect = (element: HTMLElement): DOMRect => {
    let rect = rectangles.get(element);
    if (!rect) {
      rect = element.getBoundingClientRect();
      rectangles.set(element, rect);
    }
    return rect;
  };
  const childClip = (element: HTMLElement | null): ClipBounds => {
    // Stop inside the body so page overflow remains measurable.
    if (!element || element === body) return unrestricted;
    const cached = childClips.get(element);
    if (cached) return cached;
    const inherited = childClip(element.parentElement);
    const style = element.ownerDocument.defaultView?.getComputedStyle(element);
    const clipX = CLIPPING_OVERFLOW_VALUES.has(style?.overflowX ?? "visible");
    const clipY = CLIPPING_OVERFLOW_VALUES.has(style?.overflowY ?? "visible");
    let clip = inherited;
    if (clipX || clipY) {
      const rect = readRect(element);
      const scaleX = element.offsetWidth > 0 ? rect.width / element.offsetWidth : 1;
      const scaleY = element.offsetHeight > 0 ? rect.height / element.offsetHeight : 1;
      const left = rect.left + element.clientLeft * scaleX;
      const top = rect.top + element.clientTop * scaleY;
      const right = left + element.clientWidth * scaleX;
      const bottom = top + element.clientHeight * scaleY;
      clip = {
        left: clipX ? Math.max(inherited.left, left) : inherited.left,
        top: clipY ? Math.max(inherited.top, top) : inherited.top,
        right: clipX ? Math.min(inherited.right, right) : inherited.right,
        bottom: clipY ? Math.min(inherited.bottom, bottom) : inherited.bottom,
      };
    }
    childClips.set(element, clip);
    return clip;
  };
  return (element) => {
    const rect = readRect(element);
    const clip = childClip(element.parentElement);
    const left = Math.max(rect.left, clip.left);
    const top = Math.max(rect.top, clip.top);
    const right = Math.min(rect.right, clip.right);
    const bottom = Math.min(rect.bottom, clip.bottom);
    if (
      right < left || bottom < top ||
      (rect.width > 0 && right === left) ||
      (rect.height > 0 && bottom === top)
    ) {
      return { left, top, right: left, bottom: top, width: 0, height: 0 };
    }
    return {
      left, top, right, bottom,
      width: Math.max(0, right - left),
      height: Math.max(0, bottom - top),
    };
  };
}
