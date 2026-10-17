import * as React from "react";
import { formulaReferences } from "./formula-assist";
import { useFormulaCaret } from "./FormulaEditing";

export const XlsxCellInput = React.forwardRef<HTMLTextAreaElement, React.TextareaHTMLAttributes<HTMLTextAreaElement>>(
  function XlsxCellInput({ value, style, onSelect, onFocus, ...props }, forwardedRef) {
    const inputRef = React.useRef<HTMLTextAreaElement | null>(null);
    const wrapperRef = React.useRef<HTMLDivElement>(null);
    const caret = useFormulaCaret(inputRef, value);
    const text = String(value ?? "");
    const references = formulaReferences(text);
    const [size, setSize] = React.useState({ width: 0, height: 0 });
    React.useLayoutEffect(() => {
      const input = inputRef.current;
      const wrapper = wrapperRef.current;
      if (!input || !wrapper) return;
      const resize = () => {
        const anchor = wrapper.parentElement!.getBoundingClientRect();
        const grid = input.closest('[role="grid"]')?.getBoundingClientRect();
        const css = getComputedStyle(input);
        const canvas = document.createElement("canvas");
        const context = canvas.getContext("2d")!;
        context.font = css.font;
        const padding = parseFloat(css.paddingLeft) + parseFloat(css.paddingRight) + 4;
        const measured = Math.max(...text.split("\n").map(line => context.measureText(line).width)) + padding;
        const width = Math.min(Math.max(anchor.width, measured), Math.max(1, Math.min(window.innerWidth, grid?.right ?? window.innerWidth) - anchor.left - 2));
        input.style.width = `${width}px`;
        input.style.height = `${anchor.height}px`;
        const height = Math.max(anchor.height, input.scrollHeight);
        input.style.height = `${height}px`;
        setSize(current => current.width === width && current.height === height ? current : { width, height });
      };
      resize();
      const observer = new ResizeObserver(resize);
      observer.observe(input.closest('[role="grid"]') ?? wrapper.parentElement!);
      window.addEventListener("resize", resize);
      return () => { observer.disconnect(); window.removeEventListener("resize", resize); };
    }, [text, style]);
    const parts: React.ReactNode[] = [];
    let offset = 0;
    for (const reference of references) {
      parts.push(text.slice(offset, reference.start));
      parts.push(<span key={reference.start} style={{ color: reference.color }}>{reference.text}</span>);
      offset = reference.end;
    }
    parts.push(text.slice(offset));
    const textStyle: React.CSSProperties = {
      ...style, boxSizing: "border-box", width: size.width || "100%", height: size.height || "100%",
      resize: "none", overflow: "hidden", whiteSpace: "pre-wrap", overflowWrap: "anywhere", textAlign: "left"
    };
    return <div ref={wrapperRef} style={{ position: "absolute", left: 0, top: 0, width: size.width || "100%", height: size.height || "100%", background: style?.backgroundColor === "transparent" ? "var(--xlsx-menu-surface, white)" : style?.backgroundColor }}>
      <textarea {...props} data-xlsx-editor="" aria-label="Cell editor" rows={1}
        ref={input => {
          inputRef.current = input;
          if (typeof forwardedRef === "function") forwardedRef(input);
          else if (forwardedRef) forwardedRef.current = input;
        }}
        value={value}
        onPointerDown={event => event.stopPropagation()}
        onFocus={event => { caret.onFocus(); onFocus?.(event); }}
        onSelect={event => { caret.onSelect(event); onSelect?.(event); }}
        style={{ ...textStyle, display: "block", ...(references.length ? { color: "transparent", caretColor: style?.color ?? "currentColor", WebkitTextFillColor: "transparent" } : {}) }}
      />
      {references.length ? <span aria-hidden="true" style={{ ...textStyle, position: "absolute", inset: 0, pointerEvents: "none", background: "transparent", boxShadow: "none" }}><span>{parts}</span></span> : null}
    </div>;
  }
);
