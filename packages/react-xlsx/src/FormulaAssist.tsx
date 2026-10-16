import * as React from "react";
import {
  applyFormulaFunction,
  formatFormulaArgument,
  resolveFormulaAssist,
  type FormulaAssist
} from "./formula-assist";
import type { FormulaFunction } from "./formula-functions";
import type { ViewerPalette } from "./viewer-palette";

export interface XlsxFormulaAssistState {
  assist: FormulaAssist | null;
  highlightedIndex: number;
  choose: (fn: FormulaFunction) => void;
  /** Handles the key when the suggestions use it; the editor skips its own handling then. */
  handleKeyDown: (event: React.KeyboardEvent<HTMLInputElement>) => boolean;
  /** Keeps the caret current; spread onto the input. */
  onSelect: (event: React.SyntheticEvent<HTMLInputElement>) => void;
}

/**
 * Function suggestions and argument help for a formula input, as in Google
 * Sheets. The cell editor and the formula bar share it.
 */
export function useXlsxFormulaAssist(
  value: string,
  onChange: (value: string) => void,
  inputRef: React.RefObject<HTMLInputElement | null>
): XlsxFormulaAssistState {
  const [caret, setCaret] = React.useState<number | null>(null);
  const [highlight, setHighlight] = React.useState({ index: 0, value });
  const assist = caret === null ? null : resolveFormulaAssist(value, Math.min(caret, value.length));
  const highlightedIndex =
    highlight.value === value && assist?.kind === "functions"
      ? Math.min(highlight.index, assist.matches.length - 1)
      : 0;

  const choose = React.useCallback(
    (fn: FormulaFunction) => {
      if (assist?.kind !== "functions") {
        return;
      }
      const next = applyFormulaFunction(value, assist, fn);
      onChange(next.text);
      setCaret(next.caret);
      requestAnimationFrame(() => inputRef.current?.setSelectionRange(next.caret, next.caret));
    },
    [assist, inputRef, onChange, value]
  );

  const handleKeyDown = React.useCallback(
    (event: React.KeyboardEvent<HTMLInputElement>) => {
      if (assist?.kind !== "functions") {
        return false;
      }
      if (event.key === "ArrowDown" || event.key === "ArrowUp") {
        event.preventDefault();
        const step = event.key === "ArrowDown" ? 1 : -1;
        const count = assist.matches.length;
        setHighlight({ index: (highlightedIndex + step + count) % count, value });
        return true;
      }
      if (event.key === "Enter" || event.key === "Tab") {
        event.preventDefault();
        choose(assist.matches[highlightedIndex]);
        return true;
      }
      return false;
    },
    [assist, choose, highlightedIndex, value]
  );

  const onSelect = React.useCallback((event: React.SyntheticEvent<HTMLInputElement>) => {
    setCaret(event.currentTarget.selectionStart);
  }, []);

  return { assist, choose, handleKeyDown, highlightedIndex, onSelect };
}

/** The suggestions list or the signature card, placed by the caller. */
export function XlsxFormulaAssistPopup({
  palette,
  state,
  style
}: {
  palette: ViewerPalette;
  state: XlsxFormulaAssistState;
  style?: React.CSSProperties;
}) {
  const { assist } = state;
  if (!assist) {
    return null;
  }
  const surface: React.CSSProperties = {
    backgroundColor: palette.surface,
    border: `1px solid ${palette.strongBorder}`,
    borderRadius: 8,
    boxShadow: "0 8px 24px rgba(0, 0, 0, 0.18)",
    color: palette.text,
    font: "12px/1.4 system-ui, -apple-system, \"Segoe UI\", sans-serif",
    maxWidth: 360,
    minWidth: 240,
    position: "absolute",
    zIndex: 40,
    ...style
  };

  if (assist.kind === "signature") {
    return (
      <div role="tooltip" style={{ ...surface, padding: "8px 10px" }}>
        <FormulaSignature fn={assist.fn} argumentIndex={assist.argumentIndex} />
        <div style={{ color: palette.mutedText, marginTop: 4 }}>{assist.fn.summary}</div>
      </div>
    );
  }

  const highlighted = assist.matches[state.highlightedIndex];
  return (
    <div role="listbox" aria-label="Functions" style={{ ...surface, padding: 4 }}>
      {assist.matches.map((fn, index) => (
        <div
          key={fn.name}
          role="option"
          aria-selected={index === state.highlightedIndex}
          // Keeps focus, and the caret, in the input.
          onMouseDown={(event) => {
            event.preventDefault();
            state.choose(fn);
          }}
          style={{
            // A tint that reads on both palettes' surfaces.
            backgroundColor: index === state.highlightedIndex ? "rgba(127, 127, 127, 0.2)" : "transparent",
            borderRadius: 4,
            cursor: "pointer",
            fontFamily: "ui-monospace, SFMono-Regular, Menlo, monospace",
            padding: "4px 8px"
          }}
        >
          {fn.name}
        </div>
      ))}
      {highlighted ? (
        <div style={{ borderTop: `1px solid ${palette.border}`, marginTop: 4, padding: "6px 8px 4px" }}>
          <FormulaSignature fn={highlighted} />
          <div style={{ color: palette.mutedText, marginTop: 2 }}>{highlighted.summary}</div>
        </div>
      ) : null}
    </div>
  );
}

function FormulaSignature({ argumentIndex, fn }: { argumentIndex?: number; fn: FormulaFunction }) {
  return (
    <div style={{ fontFamily: "ui-monospace, SFMono-Regular, Menlo, monospace" }}>
      {fn.name}(
      {fn.arguments.map((argument, index) => (
        <React.Fragment key={argument.name}>
          {index > 0 ? ", " : null}
          <span style={index === argumentIndex ? { fontWeight: 700, textDecoration: "underline" } : undefined}>
            {formatFormulaArgument(argument)}
          </span>
        </React.Fragment>
      ))}
      )
    </div>
  );
}
