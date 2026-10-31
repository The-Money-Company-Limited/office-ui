export type DocumentFile = { path: string; bytes: ArrayBuffer }
export type DocumentSnapshot = { files: DocumentFile[]; saved?: () => void }
// A read passes `commit: false` and the editor must stay as it is; a save
// commits edits in progress, such as a table cell's text.
export type SerializeOptions = { commit: boolean }
export type DocumentAdapter = {
  serialize: (
    options?: SerializeOptions
  ) => DocumentSnapshot | Promise<DocumentSnapshot>
  download: () => void | Promise<void>
}
// Why the page can't edit: shown instead of edit controls.
export type SaveReason = "read_only" | "signed_out" | "not_document" | "unavailable"
export type SaveState = {
  // False until access is known (an answer, or the bridge timing out), so an
  // editor never flashes "View only".
  known: boolean
  canSave: boolean
  reason?: SaveReason
  status: string
  conflict: boolean
}

// A save the bridge refused for good: what the person reads (what happened,
// then how to keep their work), why editing ends, and whether it comes back
// when the bridge's access answer turns from no to yes.
type Refusal = { message: string; reason: SaveReason; resumes?: true }
const keep = "Download a copy to keep your changes."
const refusals: Record<string, Refusal> = {
  forbidden: {
    message: `You can no longer edit this. ${keep}`,
    reason: "read_only",
    resumes: true,
  },
  read_only: {
    message: `This document can't be edited here any more. ${keep}`,
    reason: "read_only",
    resumes: true,
  },
  signed_out: {
    message: "Your sign-in expired. Sign in to keep saving.",
    reason: "signed_out",
    resumes: true,
  },
  path_not_allowed: {
    message: `Changes to this document can't be saved. ${keep}`,
    reason: "read_only",
  },
  not_document: {
    message: `This page is no longer a document. ${keep}`,
    reason: "not_document",
  },
  too_large: {
    message: `This document is too large to save here. ${keep}`,
    reason: "read_only",
  },
}
const unknownRefusal: Refusal = {
  message: `Couldn't save. ${keep}`,
  reason: "read_only",
}

export class Autosave {
  private state: SaveState = {
    known: false,
    canSave: false,
    status: "",
    conflict: false,
  }
  private listeners = new Set<() => void>()
  private version = ""
  private liveVersion = ""
  private changedDuringSave: string | null = null
  private revision = 0
  private savedRevision = 0
  private pending: {
    id: string
    revision: number
    base: string
    snapshot?: DocumentSnapshot
  } | null = null
  private inFlight = false
  private retryWaiting = false
  private retryNotBefore = 0
  private adapter: DocumentAdapter | null = null
  private timer?: ReturnType<typeof setTimeout>
  private acknowledgementTimer?: ReturnType<typeof setTimeout>
  private helloTimer?: ReturnType<typeof setTimeout>
  private fadeTimer?: ReturnType<typeof setTimeout>
  private retry = 0
  private stopped = false
  private restore: {
    x: number
    y: number
    id?: string
    offset: number
  } | null = null
  private receivedAccess = false
  // The refusal or conflict that stopped saving; undefined while saving works.
  private blockedBy?: Refusal | "conflict"
  // A refusal resumes only after the bridge has said no since it, so a save
  // the bridge's access answer disagrees with isn't retried on every answer.
  private refusedSinceBlock = false

  constructor(
    private page: Window,
    private reload: (version: string) => void
  ) {}

  getSnapshot = () => this.state
  subscribe = (listener: () => void) => {
    this.listeners.add(listener)
    return () => {
      this.listeners.delete(listener)
    }
  }
  private update(patch: Partial<SaveState>) {
    this.state = { ...this.state, ...patch }
    this.listeners.forEach((listener) => listener())
  }
  private send(message: object, transfer: Transferable[] = []) {
    this.page.postMessage(
      { cool: "doc", ...message },
      this.page.location.origin,
      transfer
    )
  }
  start() {
    this.stopped = false
    this.page.addEventListener("message", this.message)
    this.page.addEventListener("pagehide", this.hide)
    this.page.document.addEventListener("visibilitychange", this.visibility)
    this.hello()
    return () => {
      this.stopped = true
      clearTimeout(this.timer)
      clearTimeout(this.acknowledgementTimer)
      clearTimeout(this.helloTimer)
      clearTimeout(this.fadeTimer)
      this.page.removeEventListener("message", this.message)
      this.page.removeEventListener("pagehide", this.hide)
      this.page.removeEventListener("beforeunload", this.beforeUnload)
      this.page.document.removeEventListener(
        "visibilitychange",
        this.visibility
      )
    }
  }
  private hello = () => {
    this.send({ type: "hello" })
    this.helloTimer = setTimeout(() => {
      if (this.receivedAccess) return
      this.update({ known: true, reason: "unavailable" })
      this.hello()
    }, 3000)
  }
  register(adapter: DocumentAdapter) {
    this.adapter = adapter
    return () => {
      if (this.adapter === adapter) this.adapter = null
    }
  }
  private loaded = false
  // useDocumentRevision reports whether the open document has finished loading.
  setLoaded(loaded: boolean) {
    this.loaded = loaded
  }
  // The open document's bytes as they are now, unsaved edits included, read
  // without changing the editor; null while it is still loading, when the
  // editor holds an empty document.
  async documentBytes(): Promise<Uint8Array | null> {
    if (!this.loaded) return null
    const file = (await this.adapter?.serialize({ commit: false }))?.files[0]
    return file ? new Uint8Array(file.bytes) : null
  }
  dirty = () => {
    if (!this.state.canSave) return
    this.revision++
    this.page.addEventListener("beforeunload", this.beforeUnload)
    clearTimeout(this.fadeTimer)
    if (this.state.status === "Saved") this.update({ status: "" })
    if (!this.blockedBy && !this.retry) this.schedule(2000)
  }
  private schedule(delay: number) {
    clearTimeout(this.timer)
    this.timer = setTimeout(() => {
      this.retryWaiting = false
      void this.flush()
    }, delay)
  }
  private beforeUnload = (event: BeforeUnloadEvent) => {
    event.preventDefault()
    event.returnValue = ""
  }
  private hide = () => {
    void this.flush()
  }
  private visibility = () => {
    if (this.page.document.visibilityState === "hidden") this.hide()
    // Coming back, e.g. from signing in in another tab: ask for access again.
    else this.send({ type: "hello" })
  }
  flush = async () => {
    if (
      this.stopped ||
      !this.state.canSave ||
      this.blockedBy ||
      this.inFlight ||
      this.retryWaiting ||
      !this.adapter ||
      this.revision === this.savedRevision
    )
      return
    if (Date.now() < this.retryNotBefore) {
      this.schedule(this.retryNotBefore - Date.now())
      return
    }
    clearTimeout(this.timer)
    const flight = this.pending ?? {
      id: crypto.randomUUID(),
      revision: this.revision,
      base: this.version,
      snapshot: undefined as DocumentSnapshot | undefined,
    }
    this.pending = flight
    this.inFlight = true
    this.update({ status: "Saving…" })
    try {
      flight.snapshot ??= await this.adapter.serialize()
      if (this.stopped || this.pending !== flight) return
      if (!this.state.canSave) {
        this.inFlight = false
        return
      }
      // Keep the snapshot intact for acknowledgements; transferred buffers detach.
      const files = flight.snapshot.files.map(({ path, bytes }) => ({
        path,
        bytes: bytes.slice(0),
      }))
      this.send(
        { type: "save", id: flight.id, files, base: flight.base },
        files.map((file) => file.bytes)
      )
      // Outwait the bridge (19 minutes) and retain the request if its result is unknown.
      this.acknowledgementTimer = setTimeout(() => {
        if (this.pending === flight) this.failed("offline")
      }, 20 * 60_000)
    } catch {
      if (this.pending === flight) {
        if (!flight.snapshot) this.pending = null
        this.failed("error")
      }
    }
  }
  private failed(reason: string, version?: string, retryAfter?: number) {
    clearTimeout(this.timer)
    clearTimeout(this.acknowledgementTimer)
    this.inFlight = false
    if (version) this.liveVersion = version
    if (reason === "conflict") {
      this.blockedBy = "conflict"
      this.update({ conflict: true, status: "Couldn't save" })
    } else if (reason === "offline" || reason === "error") {
      this.update({
        status:
          reason === "offline"
            ? "Offline, will retry"
            : "Couldn't save yet. Retrying…",
      })
      this.retryWaiting = true
      const delay = Math.min(30_000, 2000 * 2 ** this.retry++)
      if (typeof retryAfter === "number" && Number.isFinite(retryAfter))
        this.retryNotBefore = Math.max(
          this.retryNotBefore,
          Date.now() + retryAfter * 1000
        )
      this.schedule(Math.max(delay, this.retryNotBefore - Date.now()))
    } else {
      const refusal = refusals[reason]
      if (!refusal) console.warn(`Save refused: ${reason}`)
      this.blockedBy = refusal ?? unknownRefusal
      this.refusedSinceBlock = false
      this.update({
        canSave: false,
        reason: this.blockedBy.reason,
        status: this.blockedBy.message,
      })
    }
  }
  private message = (event: MessageEvent) => {
    if (
      event.source !== this.page ||
      event.origin !== this.page.location.origin
    )
      return
    const data = event.data
    if (!data || data.cool !== "doc") return
    if (
      data.type === "access" &&
      typeof data.canSave === "boolean" &&
      typeof data.version === "string"
    ) {
      this.receivedAccess = true
      clearTimeout(this.helloTimer)
      if (!this.version) this.version = data.version
      this.liveVersion = data.version
      const refusal =
        this.blockedBy === "conflict" ? undefined : this.blockedBy
      if (refusal && !data.canSave) this.refusedSinceBlock = true
      if (refusal?.resumes && data.canSave && this.refusedSinceBlock) {
        this.blockedBy = undefined
        this.update({ status: "" })
      }
      const stuck = this.blockedBy !== "conflict" && this.blockedBy
      const regained = data.canSave && !this.state.canSave
      this.update({
        known: true,
        canSave: data.canSave && !stuck,
        // Access refusals follow the bridge, which knows the current reason.
        reason:
          stuck && (!stuck.resumes || data.canSave)
            ? stuck.reason
            : data.reason,
        ...(data.reason === "not_document"
          ? {
              status:
                this.revision === this.savedRevision
                  ? "This page is no longer a document."
                  : refusals.not_document.message,
            }
          : {}),
      })
      // Saving can resume with unsaved work waiting: after a refusal, or when
      // access came back during a retry or a save that was cut short.
      if (
        regained &&
        !this.blockedBy &&
        this.revision !== this.savedRevision
      )
        this.schedule(0)
    } else if (
      data.type === "saved" &&
      this.pending &&
      data.id === this.pending.id &&
      typeof data.version === "string"
    ) {
      clearTimeout(this.timer)
      clearTimeout(this.acknowledgementTimer)
      const flight = this.pending
      this.pending = null
      this.inFlight = false
      this.retryWaiting = false
      flight.snapshot?.saved?.()
      this.savedRevision = flight.revision
      this.version = data.version
      this.liveVersion = this.changedDuringSave ?? data.version
      this.changedDuringSave = null
      this.retry = 0
      if (this.revision !== this.savedRevision) {
        this.update({ status: "" })
        this.schedule(0)
      } else {
        this.update({ status: "Saved" })
        this.fadeTimer = setTimeout(() => this.update({ status: "" }), 3000)
        this.page.removeEventListener("beforeunload", this.beforeUnload)
        if (this.liveVersion !== this.version) this.reloadTheirs()
      }
    } else if (
      data.type === "save_failed" &&
      this.inFlight &&
      data.id === this.pending?.id
    ) {
      this.failed(data.reason, data.version, data.retryAfter)
    } else if (
      data.type === "changed" &&
      typeof data.version === "string" &&
      data.version !== this.version
    ) {
      this.liveVersion = data.version
      if (this.pending) this.changedDuringSave = data.version
      if (this.revision === this.savedRevision && !this.pending)
        this.reloadTheirs()
    }
  }
  reloadTheirs = () => {
    clearTimeout(this.timer)
    clearTimeout(this.acknowledgementTimer)
    this.pending = null
    this.inFlight = false
    this.retryWaiting = false
    this.page.removeEventListener("beforeunload", this.beforeUnload)
    this.changedDuringSave = null
    this.savedRevision = this.revision
    this.version = this.liveVersion
    this.blockedBy = undefined
    this.retry = 0
    this.update({ conflict: false, status: "" })
    const selection = this.page.getSelection()
    this.restore = {
      x: this.page.scrollX,
      y: this.page.scrollY,
      id: selection?.anchorNode?.parentElement?.closest("[id]")?.id,
      offset: selection?.anchorOffset ?? 0,
    }
    this.reload(this.version)
  }
  documentReady = () => {
    const position = this.restore
    if (!position) return
    this.restore = null
    this.page.requestAnimationFrame(() =>
      this.page.requestAnimationFrame(() => {
        this.page.scrollTo(position.x, position.y)
        const node = position.id
          ? this.page.document.getElementById(position.id)?.firstChild
          : null
        if (node?.nodeType === 3)
          this.page
            .getSelection()
            ?.collapse(
              node,
              Math.min(position.offset, node.textContent?.length ?? 0)
            )
      })
    )
  }
  // True when the download started; the caller says so where it was asked.
  downloadCopy = async () => {
    try {
      await this.adapter?.download()
      return true
    } catch {
      return false
    }
  }
}
