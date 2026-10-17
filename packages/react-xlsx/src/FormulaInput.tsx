import * as React from "react";
import { formulaReferences } from "./formula-assist";

/** A native input keeps text selection, IME and accessibility; its mirror colours references. */
export const XlsxFormulaInput = React.forwardRef<HTMLInputElement, React.InputHTMLAttributes<HTMLInputElement>>(
  function XlsxFormulaInput({ value, style, className, onScroll, onPointerDown, ...props }, forwardedRef) {
    const inputRef = React.useRef<HTMLInputElement | null>(null);
    const [mirrorStyle, setMirrorStyle] = React.useState<React.CSSProperties>({});
    const [scrollLeft, setScrollLeft] = React.useState(0);
    const text = String(value ?? "");
    const references = formulaReferences(text);
    const tinted = references.length > 0;
    React.useLayoutEffect(() => {
      const input = inputRef.current;
      if (!input) return;
      const css = getComputedStyle(input);
      const next = {
        font: css.font,
        letterSpacing: css.letterSpacing,
        padding: css.padding,
        borderWidth: css.borderWidth,
        borderStyle: "solid",
        borderColor: "transparent",
        textAlign: css.textAlign as React.CSSProperties["textAlign"]
      };
      setMirrorStyle((current) => JSON.stringify(current) === JSON.stringify(next) ? current : next);
      setScrollLeft(input.scrollLeft);
    }, [value, style, className]);
    const parts: React.ReactNode[] = [];
    let offset = 0;
    for (const reference of references) {
      parts.push(text.slice(offset, reference.start));
      parts.push(<span key={reference.start} style={{ color: reference.color }}>{reference.text}</span>);
      offset = reference.end;
    }
    parts.push(text.slice(offset));
    return (
      <span style={{ position: "relative", display: "block", flex: "1 1 auto", minWidth: 0, width: "100%", height: style?.height, color: style?.color }}>
        <input
          {...props}
          className={className}
          onPointerDown={(event) => { event.stopPropagation(); onPointerDown?.(event); }}
          ref={(input) => {
            inputRef.current = input;
            if (typeof forwardedRef === "function") forwardedRef(input);
            else if (forwardedRef) forwardedRef.current = input;
          }}
          value={value}
          onScroll={(event) => {
            setScrollLeft(event.currentTarget.scrollLeft);
            onScroll?.(event);
          }}
          style={{ ...style, width: "100%", ...(tinted ? { color: "transparent", caretColor: style?.color ?? "var(--foreground, currentColor)", WebkitTextFillColor: "transparent" } : {}) }}
        />
        {tinted ? (
          <span aria-hidden="true" style={{ ...mirrorStyle, position: "absolute", inset: 0, boxSizing: "border-box", display: "flex", alignItems: "center", overflow: "hidden", pointerEvents: "none", whiteSpace: "pre" }}>
            <span style={{ transform: `translateX(${-scrollLeft}px)`, flexShrink: 0 }}>{parts}</span>
          </span>
        ) : null}
      </span>
    );
  }
);
