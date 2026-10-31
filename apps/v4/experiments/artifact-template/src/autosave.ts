export type DocumentFile = { path: string; bytes: ArrayBuffer }
export type DocumentSnapshot = { files: DocumentFile[]; saved?: () => void }
export type DocumentAdapter = {
  serialize: () => DocumentSnapshot | Promise<DocumentSnapshot>
  download: () => void | Promise<void>
}
export type SaveState = {
  canSave: boolean
  reason?: string
  status: string
  conflict: boolean
}

const failures: Record<string, string> = {
  forbidden: "You no longer have permission to edit. Keep a copy or reload.",
  path_not_allowed:
    "This save includes a path outside the document. Keep a copy and reload.",
  not_document: "This page is no longer a document",
  too_large: "The document is too large. Download a copy.",
  error: "Please try again when the connection returns.",
}

export class Autosave {
  private state: SaveState = { canSave: false, status: "", conflict: false }
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
  private blocked = false

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
      this.update({ reason: "unavailable" })
      this.hello()
    }, 3000)
  }
  register(adapter: DocumentAdapter) {
    this.adapter = adapter
    return () => {
      if (this.adapter === adapter) this.adapter = null
    }
  }
  // Whether the open document has finished loading (useDocumentRevision).
  ready = false
  // The open document's bytes as they are now, unsaved edits included; null
  // while it is still loading, when the editor holds an empty document.
  async documentBytes(): Promise<Uint8Array | null> {
    if (!this.ready) return null
    const file = (await this.adapter?.serialize())?.files[0]
    return file ? new Uint8Array(file.bytes) : null
  }
  dirty = () => {
    if (!this.state.canSave) return
    this.revision++
    this.page.addEventListener("beforeunload", this.beforeUnload)
    clearTimeout(this.fadeTimer)
    if (this.state.status === "Saved") this.update({ status: "" })
    if (!this.blocked && !this.retry) this.schedule(2000)
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
  }
  flush = async () => {
    if (
      this.stopped ||
      !this.state.canSave ||
      this.blocked ||
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
      this.blocked = true
      this.update({ conflict: true, status: "Couldn't save" })
    } else if (reason === "offline" || reason === "error") {
      this.update({
        status:
          reason === "offline"
            ? "Offline, will retry"
            : `Couldn't save. ${failures.error}`,
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
      this.blocked = true
      this.update({
        canSave: ["forbidden", "not_document"].includes(reason)
          ? false
          : this.state.canSave,
        reason:
          reason === "not_document"
            ? reason
            : reason === "forbidden"
              ? "read_only"
              : this.state.reason,
        status: `Couldn't save. ${failures[reason] ?? reason}`,
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
      this.update({
        canSave: data.canSave,
        reason: data.reason,
        ...(data.reason === "not_document"
          ? { status: failures.not_document }
          : {}),
      })
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
    this.blocked = false
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
  downloadCopy = async () => {
    try {
      await this.adapter?.download()
    } catch {
      this.update({
        status: "Couldn't save. Download failed; please try again.",
      })
    }
  }
}
