import * as React from "react";
import type { FormulaPoint } from "./formula-assist";
import type { XlsxCellAddress } from "./types";

export interface FormulaEditorSession {
  input: HTMLInputElement;
  point: FormulaPoint | null;
  canPoint: () => boolean;
  insert: (cell: XlsxCellAddress, extend?: boolean) => boolean;
}

interface FormulaEditing {
  editor: React.MutableRefObject<FormulaEditorSession | null>;
  navigate: React.MutableRefObject<((key: string, extend: boolean) => boolean) | null>;
  refresh: () => void;
}

export const FormulaEditingContext = React.createContext<FormulaEditing | null>(null);

export function FormulaEditingProvider({ children }: { children: React.ReactNode }) {
  const parent = React.useContext(FormulaEditingContext);
  const editor = React.useRef<FormulaEditorSession | null>(null);
  const navigate = React.useRef<FormulaEditing["navigate"]["current"]>(null);
  const [revision, refresh] = React.useReducer((value: number) => value + 1, 0);
  const value = React.useMemo(() => ({ editor, navigate, refresh }), [revision]);
  return <FormulaEditingContext.Provider value={parent ?? value}>{children}</FormulaEditingContext.Provider>;
}
