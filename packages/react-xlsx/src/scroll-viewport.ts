import type { Virtualizer } from "@tanstack/react-virtual";

export interface ScrollViewportRect {
  bottom: number;
  height: number;
  left: number;
  right: number;
  top: number;
  width: number;
}

/**
 * The scroll surface the grid reads and writes. An `HTMLElement` satisfies it directly; in window
 * mode {@link createWindowScrollViewport} presents the window's scroll position as if it were the
 * grid element's own, so every reader keeps one formula.
 */
export interface ScrollViewport {
  scrollTop: number;
  scrollLeft: number;
  readonly clientWidth: number;
  readonly clientHeight: number;
  readonly scrollWidth: number;
  readonly scrollHeight: number;
  getBoundingClientRect(): ScrollViewportRect;
}

export interface WindowScrollMetricsInput {
  /** The grid element's border box in client coordinates. */
  gridLeft: number;
  gridTop: number;
  /** The grid's full content size (its scrollWidth / scrollHeight). */
  gridWidth: number;
  gridHeight: number;
  /** The window's layout viewport size, excluding scrollbars. */
  viewportWidth: number;
  viewportHeight: number;
  /** Height of an app toolbar fixed above the grid; sticky headers stick below it. */
  stickyTop: number;
}

export interface WindowScrollMetrics {
  clientHeight: number;
  clientWidth: number;
  /** The part of the viewport the grid is visible in, in client coordinates. */
  rect: ScrollViewportRect;
  scrollLeft: number;
  scrollTop: number;
}

/**
 * Maps the grid's position on a window-scrolled page to element-style scroll metrics: the scroll
 * offset is how far the grid has passed under the viewport's top-left (below `stickyTop`), and the
 * client size is the part of the grid visible in the viewport. `scrollTop + clientHeight` never
 * exceeds the grid height, so canvases sized to the client box never grow the page.
 */
export function resolveWindowScrollMetrics(input: WindowScrollMetricsInput): WindowScrollMetrics {
  const visibleLeft = Math.max(input.gridLeft, 0);
  const visibleTop = Math.max(input.gridTop, input.stickyTop);
  const visibleRight = Math.min(input.gridLeft + input.gridWidth, input.viewportWidth);
  const visibleBottom = Math.min(input.gridTop + input.gridHeight, input.viewportHeight);
  const clientWidth = Math.max(0, visibleRight - visibleLeft);
  const clientHeight = Math.max(0, visibleBottom - visibleTop);
  return {
    clientHeight,
    clientWidth,
    rect: {
      bottom: visibleTop + clientHeight,
      height: clientHeight,
      left: visibleLeft,
      right: visibleLeft + clientWidth,
      top: visibleTop,
      width: clientWidth
    },
    scrollLeft: Math.min(visibleLeft - input.gridLeft, input.gridWidth),
    scrollTop: Math.min(visibleTop - input.gridTop, input.gridHeight)
  };
}

function readWindowScrollMetrics(element: HTMLElement, stickyTop: number) {
  const gridRect = element.getBoundingClientRect();
  return {
    gridRect,
    metrics: resolveWindowScrollMetrics({
      gridHeight: element.scrollHeight,
      gridLeft: gridRect.left,
      gridTop: gridRect.top,
      gridWidth: element.scrollWidth,
      stickyTop,
      viewportHeight: document.documentElement.clientHeight,
      viewportWidth: document.documentElement.clientWidth
    })
  };
}

/**
 * A {@link ScrollViewport} over the window for a grid element that does not scroll itself.
 * Writing `scrollTop`/`scrollLeft` scrolls the window so the grid reaches that offset; writing the
 * current offset is a no-op, so restoring offset 0 never scrolls a grid that sits lower on the page.
 */
export function createWindowScrollViewport(element: HTMLElement, stickyTop: number): ScrollViewport {
  const read = () => readWindowScrollMetrics(element, stickyTop);
  const scrollAxis = (axis: "left" | "top", offset: number) => {
    const { gridRect, metrics } = read();
    const current = axis === "top" ? metrics.scrollTop : metrics.scrollLeft;
    const target = Math.max(0, offset);
    if (Math.abs(target - current) < 0.5) {
      return;
    }
    if (axis === "top") {
      window.scrollTo({ behavior: "instant", top: window.scrollY + gridRect.top - stickyTop + target });
    } else {
      window.scrollTo({ behavior: "instant", left: window.scrollX + gridRect.left + target });
    }
  };

  return {
    get scrollTop() {
      return read().metrics.scrollTop;
    },
    set scrollTop(offset: number) {
      scrollAxis("top", offset);
    },
    get scrollLeft() {
      return read().metrics.scrollLeft;
    },
    set scrollLeft(offset: number) {
      scrollAxis("left", offset);
    },
    get clientWidth() {
      return read().metrics.clientWidth;
    },
    get clientHeight() {
      return read().metrics.clientHeight;
    },
    get scrollWidth() {
      return element.scrollWidth;
    },
    get scrollHeight() {
      return element.scrollHeight;
    },
    getBoundingClientRect() {
      return read().metrics.rect;
    }
  };
}

/**
 * Virtualizer hooks that observe the window instead of the scroll element, so the DOM grid's
 * row and column virtualizers see the same offsets {@link createWindowScrollViewport} reports.
 */
export function createWindowVirtualizerObservers(getViewport: () => ScrollViewport | null) {
  const subscribe = (listener: () => void) => {
    window.addEventListener("scroll", listener, { passive: true });
    window.addEventListener("resize", listener);
    return () => {
      window.removeEventListener("scroll", listener);
      window.removeEventListener("resize", listener);
    };
  };

  return {
    observeElementRect: (
      _instance: Virtualizer<HTMLDivElement, Element>,
      callback: (rect: { height: number; width: number }) => void
    ) => {
      const update = () => {
        const viewport = getViewport();
        if (viewport) {
          callback({ height: viewport.clientHeight, width: viewport.clientWidth });
        }
      };
      update();
      return subscribe(update);
    },
    observeElementOffset: (
      instance: Virtualizer<HTMLDivElement, Element>,
      callback: (offset: number, isScrolling: boolean) => void
    ) => {
      const update = () => {
        const viewport = getViewport();
        if (viewport) {
          callback(instance.options.horizontal ? viewport.scrollLeft : viewport.scrollTop, false);
        }
      };
      update();
      return subscribe(update);
    },
    scrollToFn: (
      offset: number,
      options: { adjustments?: number },
      instance: Virtualizer<HTMLDivElement, Element>
    ) => {
      const viewport = getViewport();
      if (!viewport) {
        return;
      }
      const target = offset + (options.adjustments ?? 0);
      if (instance.options.horizontal) {
        viewport.scrollLeft = target;
      } else {
        viewport.scrollTop = target;
      }
    }
  };
}
