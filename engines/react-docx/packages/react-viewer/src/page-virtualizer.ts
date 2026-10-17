import * as React from "react";
import { flushSync } from "react-dom";
import {
  Virtualizer,
  elementScroll,
  observeElementOffset,
  observeElementRect,
  observeWindowOffset,
  observeWindowRect,
  windowScroll,
  type PartialKeys,
  type VirtualizerOptions,
} from "@tanstack/react-virtual";

type PageVirtualizerOptions<TScrollElement extends Element | Window> =
  PartialKeys<
    VirtualizerOptions<TScrollElement, Element>,
    "observeElementRect" | "observeElementOffset" | "scrollToFn"
  >;

/**
 * TanStack's React adapter re-renders its host whenever a scroll starts and
 * again when it settles (`isScrolling`), as well as when the range changes.
 * The page stack renders every visible page, so each of those re-renders costs
 * a full page render in the middle of a scroll. The viewer only reads the
 * range, so this re-renders when the range changes and on nothing else.
 */
function usePageRangeVirtualizer<TScrollElement extends Element | Window>(
  options: VirtualizerOptions<TScrollElement, Element>
): Virtualizer<TScrollElement, Element> {
  const rerender = React.useReducer((count: number) => count + 1, 0)[1];
  const renderedRangeRef = React.useRef<string>("");
  const resolvedOptions: VirtualizerOptions<TScrollElement, Element> = {
    ...options,
    onChange: (instance, sync) => {
      const range = instance.range;
      const renderedRange = range
        ? `${range.startIndex}:${range.endIndex}`
        : "";
      if (renderedRange === renderedRangeRef.current) {
        return;
      }
      renderedRangeRef.current = renderedRange;
      if (sync) {
        flushSync(rerender);
      } else {
        rerender();
      }
    },
  };
  const [instance] = React.useState(
    () => new Virtualizer<TScrollElement, Element>(resolvedOptions)
  );
  instance.setOptions(resolvedOptions);
  React.useLayoutEffect(() => instance._didMount(), [instance]);
  React.useLayoutEffect(() => instance._willUpdate());
  return instance;
}

export function usePageVirtualizer(
  options: PageVirtualizerOptions<Element>
): Virtualizer<Element, Element> {
  return usePageRangeVirtualizer<Element>({
    observeElementRect,
    observeElementOffset,
    scrollToFn: elementScroll,
    ...options,
  });
}

export function useWindowPageVirtualizer(
  options: PartialKeys<PageVirtualizerOptions<Window>, "getScrollElement">
): Virtualizer<Window, Element> {
  return usePageRangeVirtualizer<Window>({
    getScrollElement: () => (typeof document !== "undefined" ? window : null),
    observeElementRect: observeWindowRect,
    observeElementOffset: observeWindowOffset,
    scrollToFn: windowScroll,
    initialOffset: () => (typeof document !== "undefined" ? window.scrollY : 0),
    ...options,
  });
}
