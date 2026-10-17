import * as React from "react";
import type { FormulaPoint } from "./formula-assist";
import type { XlsxCellAddress } from "./types";

export interface FormulaEditorSession {
  input: FormulaTextElement;
  point: FormulaPoint | null;
  canPoint: () => boolean;
  insert: (cell: XlsxCellAddress, extend?: boolean) => boolean;
}

export type FormulaTextElement = HTMLInputElement | HTMLTextAreaElement;

interface CellDraft {
  cell: XlsxCellAddress;
  value: string;
  focusCell: boolean;
}

interface FormulaEditing {
  draft: CellDraft | null;
  setDraft: React.Dispatch<React.SetStateAction<CellDraft | null>>;
  actions: React.MutableRefObject<{ commit: () => void; cancel: () => void } | null>;
  caret: React.MutableRefObject<[number, number]>;
  inputs: React.MutableRefObject<Set<FormulaTextElement>>;
  editor: React.MutableRefObject<FormulaEditorSession | null>;
  navigate: React.MutableRefObject<((key: string, extend: boolean) => boolean) | null>;
  refresh: () => void;
}

export const FormulaEditingContext = React.createContext<FormulaEditing | null>(null);

export function FormulaEditingProvider({ children }: { children: React.ReactNode }) {
  const parent = React.useContext(FormulaEditingContext);
  const [draft, setDraft] = React.useState<CellDraft | null>(null);
  const actions = React.useRef<FormulaEditing["actions"]["current"]>(null);
  const caret = React.useRef<[number, number]>([0, 0]);
  const inputs = React.useRef(new Set<FormulaTextElement>());
  const editor = React.useRef<FormulaEditorSession | null>(null);
  const navigate = React.useRef<FormulaEditing["navigate"]["current"]>(null);
  const [revision, refresh] = React.useReducer((value: number) => value + 1, 0);
  const value = React.useMemo(() => ({ editor, navigate, refresh, draft, setDraft, actions, caret, inputs }), [revision, draft]);
  return <FormulaEditingContext.Provider value={parent ?? value}>{children}</FormulaEditingContext.Provider>;
}

/** One uncommitted draft shared by the cell editor and formula bar. */
export function useXlsxCellEditor() {
  const editing = React.useContext(FormulaEditingContext);
  if (!editing) throw new Error("useXlsxCellEditor requires XlsxViewerProvider.");
  return {
    draft: editing.draft,
    begin: (cell: XlsxCellAddress, value: string) => {
      editing.caret.current = [value.length, value.length];
      editing.setDraft({ cell, value, focusCell: false });
    },
    change: React.useCallback((value: string) => {
      editing.setDraft(current => current ? { ...current, value } : null);
    }, [editing.setDraft]),
    commit: () => editing.actions.current?.commit(),
    cancel: () => editing.actions.current?.cancel(),
    onBlur: (event: React.FocusEvent<FormulaTextElement>) => {
      if (!(event.relatedTarget instanceof Element && event.relatedTarget.hasAttribute("data-xlsx-editor"))) {
        editing.actions.current?.commit();
      }
    }
  };
}

export function useFormulaCaret(inputRef: React.RefObject<FormulaTextElement | null>, value: unknown) {
  const editing = React.useContext(FormulaEditingContext);
  React.useLayoutEffect(() => {
    const input = inputRef.current;
    if (!input || !editing) return;
    editing.inputs.current.add(input);
    return () => { editing.inputs.current.delete(input); };
  }, [editing?.inputs, inputRef]);
  React.useLayoutEffect(() => {
    const input = inputRef.current;
    if (input && editing?.draft && document.activeElement !== input) {
      input.setSelectionRange(...editing.caret.current);
    }
  }, [value, editing?.draft, editing?.caret, inputRef]);
  return {
    onFocus: () => {
      if (editing?.draft) inputRef.current?.setSelectionRange(...editing.caret.current);
    },
    onSelect: (event: React.SyntheticEvent<FormulaTextElement>) => {
      const input = event.currentTarget;
      if (!editing?.draft || document.activeElement !== input) return;
      editing.caret.current = [input.selectionStart ?? 0, input.selectionEnd ?? 0];
      for (const other of editing.inputs.current) {
        if (other !== input) other.setSelectionRange(...editing.caret.current);
      }
    }
  };
}
