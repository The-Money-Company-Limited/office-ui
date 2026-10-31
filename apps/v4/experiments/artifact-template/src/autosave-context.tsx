"use client"

import * as React from "react"

import { Button } from "@/components/ui/button"
import {
  Popover,
  PopoverPopup,
  PopoverTitle,
  PopoverTrigger,
} from "@/components/ui/popover"

import type { Autosave, DocumentAdapter, SaveState } from "./autosave"

export const AutosaveContext = React.createContext<Autosave | null>(null)
const viewing: SaveState = {
  known: true,
  canSave: true,
  status: "",
  conflict: false,
}
const subscribe = () => () => {}
const snapshot = () => viewing

export function useDocumentSave() {
  return React.useContext(AutosaveContext)
}
export function useSaveState() {
  const save = useDocumentSave()
  return React.useSyncExternalStore(
    save?.subscribe ?? subscribe,
    save?.getSnapshot ?? snapshot,
    snapshot
  )
}

export function useDocumentAdapter(adapter: DocumentAdapter) {
  const save = useDocumentSave()
  const latest = React.useRef(adapter)
  React.useLayoutEffect(() => {
    latest.current = adapter
  })
  React.useLayoutEffect(
    () =>
      save?.register({
        serialize: (options) => latest.current.serialize(options),
        download: () => latest.current.download(),
      }),
    [save]
  )
}

export function useDocumentRevision(revision: unknown, ready: boolean) {
  const save = useDocumentSave()
  const previous = React.useRef({ revision, ready })
  React.useLayoutEffect(() => {
    if (
      ready &&
      previous.current.ready &&
      previous.current.revision !== revision
    )
      save?.dirty()
    previous.current = { revision, ready }
    save?.setLoaded(ready)
    if (ready) save?.documentReady()
  }, [ready, revision, save])
}

export function SaveAccess() {
  const { known, canSave, reason } = useSaveState()
  if (canSave || !known) return null
  // A new tab, so edits made before a sign-in expired stay on this page.
  return reason === "signed_out" ? (
    <a
      className="text-xs text-muted-foreground"
      href={`https://www.cool.computer/private-access?redirect_url=${encodeURIComponent(window.location.href)}`}
      target="_blank"
      rel="noopener"
    >
      Sign in to edit
    </a>
  ) : (
    <span className="text-xs text-muted-foreground">View only</span>
  )
}

export function SaveStatus() {
  const save = useDocumentSave()
  const { status, conflict } = useSaveState()
  if (!save) return null
  return (
    <Popover open={conflict}>
      <PopoverTrigger
        render={<button type="button" />}
        className="border-0 bg-transparent p-0 text-xs text-muted-foreground outline-none focus-visible:ring-2 focus-visible:ring-ring"
        tabIndex={conflict ? 0 : -1}
        style={
          status === "Saved"
            ? { animation: "autosave-saved 3s both" }
            : undefined
        }
        aria-live="polite"
        data-autosave-status
      >
        {status}
      </PopoverTrigger>
      <PopoverPopup
        align="start"
        className="w-80 border-white/15 bg-background/80 shadow-xl backdrop-blur-xl"
      >
        <PopoverTitle className="text-sm font-medium">
          A newer version was saved while you were editing.
        </PopoverTitle>
        <p className="mt-1 text-xs text-muted-foreground">
          Load latest replaces your unsaved changes. Download yours first to
          keep them.
        </p>
        <ConflictChoices save={save} />
      </PopoverPopup>
    </Popover>
  )
}

// Mounted while the conflict popover is open, so the download label resets each time.
function ConflictChoices({ save }: { save: Autosave }) {
  const [download, setDownload] = React.useState<"idle" | "done" | "failed">(
    "idle"
  )
  return (
    <div className="mt-5 flex flex-wrap justify-end gap-2">
      <Button
        variant="ghost"
        size="sm"
        onClick={() =>
          void save
            .downloadCopy()
            .then((started) => setDownload(started ? "done" : "failed"))
        }
      >
        {download === "done"
          ? "Copy downloaded"
          : download === "failed"
            ? "Download didn't start. Try again"
            : "Download mine"}
      </Button>
      <Button size="sm" onClick={save.reloadTheirs}>
        Load latest
      </Button>
    </div>
  )
}
