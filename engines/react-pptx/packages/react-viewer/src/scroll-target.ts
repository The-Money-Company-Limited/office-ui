/**
 * One scrolling viewport the viewer follows: an element scroller or the
 * window. Every scroll read, write, listener, and viewport measurement goes
 * through this so the window and element cases share one code path.
 */
export interface ScrollTarget {
  /** The scroller handed to TanStack Virtual. */
  readonly scroller: HTMLElement | Window;
  /** The scrolling element, or `undefined` when the window scrolls. */
  readonly element: HTMLElement | undefined;
  getTop(): number;
  setTop(top: number): void;
  getLeft(): number;
  setLeft(left: number): void;
  scrollTo(top: number, behavior?: ScrollBehavior): void;
  /** Top edge of the visible viewport in client coordinates. */
  clientTop(): number;
  /** Border-box size of the visible viewport; zero when not laid out. */
  rect(): { width: number; height: number };
  /** Inner size excluding scrollbars; zero when not laid out. */
  clientSize(): { width: number; height: number };
  onScroll(listener: () => void): () => void;
  /** Calls `listener` when the viewport or `content` changes size. */
  onResize(content: HTMLElement, listener: () => void): () => void;
}

export function isWindowScroller(scroller: HTMLElement | Window): boolean {
  if ('document' in scroller) return true;
  // The document's root scroller is the window: its rect spans the whole
  // document and it never receives scroll events.
  const ownerDocument = scroller.ownerDocument;
  return scroller === ownerDocument.documentElement || scroller === ownerDocument.scrollingElement;
}

function observeResize(elements: Element[], listener: () => void): () => void {
  if (typeof ResizeObserver === 'undefined') return () => {};
  const observer = new ResizeObserver(listener);
  for (const element of elements) observer.observe(element);
  return () => observer.disconnect();
}

function windowScrollTarget(view: Window): ScrollTarget {
  const root = () => view.document.documentElement;
  return {
    scroller: view,
    element: undefined,
    getTop: () => view.scrollY,
    setTop: (top) => view.scrollTo({ top, left: view.scrollX }),
    getLeft: () => view.scrollX,
    setLeft: (left) => view.scrollTo({ top: view.scrollY, left }),
    scrollTo: (top, behavior) => view.scrollTo(behavior ? { top, behavior } : { top }),
    clientTop: () => 0,
    rect: () => ({ width: view.innerWidth, height: view.innerHeight }),
    clientSize: () => ({ width: root().clientWidth, height: root().clientHeight }),
    onScroll: (listener) => {
      view.addEventListener('scroll', listener, { passive: true });
      return () => view.removeEventListener('scroll', listener);
    },
    onResize: (content, listener) => {
      view.addEventListener('resize', listener);
      const stopObserving = observeResize([content], listener);
      return () => {
        view.removeEventListener('resize', listener);
        stopObserving();
      };
    },
  };
}

function elementScrollTarget(element: HTMLElement): ScrollTarget {
  return {
    scroller: element,
    element,
    getTop: () => element.scrollTop,
    setTop: (top) => {
      element.scrollTop = top;
    },
    getLeft: () => element.scrollLeft,
    setLeft: (left) => {
      element.scrollLeft = left;
    },
    scrollTo: (top, behavior) => {
      if (typeof element.scrollTo === 'function') {
        element.scrollTo(behavior ? { top, behavior } : { top });
      } else {
        element.scrollTop = top;
      }
    },
    clientTop: () => element.getBoundingClientRect().top,
    rect: () => {
      const rect = element.getBoundingClientRect();
      return { width: rect.width, height: rect.height };
    },
    clientSize: () => ({ width: element.clientWidth, height: element.clientHeight }),
    onScroll: (listener) => {
      element.addEventListener('scroll', listener, { passive: true });
      return () => element.removeEventListener('scroll', listener);
    },
    onResize: (content, listener) =>
      observeResize(content === element ? [element] : [element, content], listener),
  };
}

export function createScrollTarget(scroller: HTMLElement | Window): ScrollTarget {
  if ('document' in scroller) return windowScrollTarget(scroller);
  const view = scroller.ownerDocument.defaultView;
  return view && isWindowScroller(scroller)
    ? windowScrollTarget(view)
    : elementScrollTarget(scroller);
}
