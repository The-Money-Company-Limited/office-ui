// Read-only WebMCP tools (COD-1456): an agent in the person's browser can read
// the open document as it is now, unsaved edits included. Tools register only
// where the browser has `document.modelContext`. They make no edits; a read
// serializes the document as an autosave tick does, which commits a table
// cell's in-progress text. Document text is marked untrusted, since it can
// carry instructions aimed at the agent reading it.

type ModelContext = {
  registerTool(tool: Tool, options: { signal: AbortSignal }): Promise<void>
}

type Tool = {
  name: string
  title: string
  description: string
  inputSchema: object
  annotations: { readOnlyHint: boolean; untrustedContentHint: boolean }
  execute: (input: Record<string, unknown>) => Promise<unknown>
}

export type DocumentKind = "docx" | "xlsx" | "pptx" | "pdf" | "slides"

const annotations = { readOnlyHint: true, untrustedContentHint: true }

const optionalString = (input: Record<string, unknown>, name: string) => {
  const value = input[name]
  if (value !== undefined && typeof value !== "string")
    throw new Error(`${name} must be a string.`)
  return value
}

function wordTools(bytes: () => Promise<Uint8Array>): Tool[] {
  const open = async () =>
    (await import("@extend-ai/react-docx/agent")).openDocx(await bytes())
  return [
    {
      name: "doc_read",
      title: "Read the document",
      description:
        "Returns the body's paragraphs in order, each with its id and text, as they are on the page now. Text in tables, headers and footers is not included.",
      inputSchema: { type: "object", properties: {} },
      annotations,
      execute: async () => ({ paragraphs: (await open()).paragraphs() }),
    },
    {
      name: "doc_find",
      title: "Find text in the document",
      description:
        "Finds every occurrence of the text in the body's paragraphs and returns each as a paragraph id with start and end offsets into that paragraph's text. Tables, headers and footers are not searched.",
      inputSchema: {
        type: "object",
        properties: { text: { type: "string", minLength: 1 } },
        required: ["text"],
      },
      annotations,
      execute: async (input) => {
        const text = optionalString(input, "text")
        if (!text) throw new Error("text must be a non-empty string.")
        return { matches: (await open()).find(text) }
      },
    },
  ]
}

function sheetTools(bytes: () => Promise<Uint8Array>): Tool[] {
  return [
    {
      name: "sheet_read",
      title: "Read the workbook",
      description:
        "Lists the sheets and returns the non-empty cells of one sheet (the first when none is named), optionally limited to an A1 range such as A1:D20. Each cell has its ref, value and formula.",
      inputSchema: {
        type: "object",
        properties: {
          sheet: { type: "string" },
          range: { type: "string", pattern: "^[A-Za-z]+[0-9]+(:[A-Za-z]+[0-9]+)?$" },
        },
      },
      annotations,
      execute: async (input) => {
        const range = optionalString(input, "range")
        if (range !== undefined && !/^[A-Za-z]+[0-9]+(:[A-Za-z]+[0-9]+)?$/.test(range))
          throw new Error("range must be an A1 reference such as A1:D20.")
        const { openXlsx } = await import("@extend-ai/react-xlsx/agent")
        const workbook = openXlsx(await bytes())
        const sheets = workbook.sheets()
        const sheet = optionalString(input, "sheet") ?? sheets[0]
        return { sheets, sheet, cells: workbook.cells(sheet, range) }
      },
    },
  ]
}

const toolsByKind: Partial<
  Record<DocumentKind, (bytes: () => Promise<Uint8Array>) => Tool[]>
> = { docx: wordTools, xlsx: sheetTools }

// Registers the tools for this kind of document and returns their unregistration.
export function registerDocumentTools(
  kind: DocumentKind,
  read: () => Promise<Uint8Array | null>
): () => void {
  const context = (document as Document & { modelContext?: ModelContext })
    .modelContext
  const tools = toolsByKind[kind]
  if (!context || !tools) return () => {}
  const bytes = async () => {
    const current = await read()
    if (!current) throw new Error("The document is still opening.")
    return current
  }
  const controller = new AbortController()
  for (const tool of tools(bytes))
    void context
      .registerTool(tool, { signal: controller.signal })
      .catch((error: unknown) =>
        console.warn(`WebMCP tool ${tool.name} was not registered`, error)
      )
  return () => controller.abort()
}
