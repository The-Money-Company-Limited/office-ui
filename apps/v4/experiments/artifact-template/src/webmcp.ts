// Read-only WebMCP tools (COD-1456, COD-1716): an agent in the person's
// browser can read the open document as it is now, unsaved edits included.
// Tools register only where the browser has `document.modelContext`. A read
// changes nothing on the page, not even a table cell's in-progress text.
// Document text is marked untrusted, since it can carry instructions aimed at
// the agent reading it.

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
const noInput = { type: "object", properties: {} }

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
        "Returns the document as it is on the page now: the body's paragraphs in order, each with its id, body index and text; the body's tables, each with its body index and its rows of cell text; and the headers and footers, each with the pages it is on (default, first, even).",
      inputSchema: noInput,
      annotations,
      execute: async () => {
        const doc = await open()
        return {
          paragraphs: doc.paragraphs(),
          tables: doc.tables(),
          headersFooters: doc.headersFooters(),
        }
      },
    },
    {
      name: "doc_find",
      title: "Find text in the document",
      description:
        "Finds the text everywhere in the document; matching is exact and case-sensitive. Matches in body paragraphs come back as a paragraph id with start and end offsets into its text; matches in tables as the table's body index with the cell's row and column; matches in headers and footers as `part`, the index into doc_read's headersFooters.",
      inputSchema: {
        type: "object",
        properties: { text: { type: "string", minLength: 1 } },
        required: ["text"],
      },
      annotations,
      execute: async (input) => {
        const text = optionalString(input, "text")
        if (!text) throw new Error("text must be a non-empty string.")
        const doc = await open()
        return {
          matches: doc.find(text),
          tableCells: doc.tables().flatMap(({ index, rows }) =>
            rows.flatMap((cells, row) =>
              cells.flatMap((cell, column) =>
                cell.includes(text) ? [{ table: index, row, column }] : []
              )
            )
          ),
          headersFooters: doc
            .headersFooters()
            .flatMap(({ kind, pages, text: partText }, part) =>
              partText.includes(text) ? [{ part, kind, pages }] : []
            ),
        }
      },
    },
    {
      name: "doc_comments",
      title: "Read the document's comments",
      description:
        "Lists the Word comments saved in the document, replies included: each with its id, author, date, text, whether it is resolved, the comment it replies to, and, for a comment on a body paragraph, that paragraph's id and the text it covers. Comments left on the page itself are not included.",
      inputSchema: noInput,
      annotations,
      execute: async () => ({ comments: (await open()).comments() }),
    },
    {
      name: "doc_revisions",
      title: "Read the document's tracked changes",
      description:
        "Lists the tracked changes (suggestions) in the body: each with its id, kind (insertion, deletion, or a suggested paragraph break or join), author, date, text and paragraph id. `unlisted` names the paragraphs and tables holding tracked changes this list cannot show, such as moves and changes inside tables.",
      inputSchema: noInput,
      annotations,
      execute: async () => (await open()).revisions(),
    },
  ]
}

function sheetTools(bytes: () => Promise<Uint8Array>): Tool[] {
  const open = async () =>
    (await import("@extend-ai/react-xlsx/agent")).openXlsx(await bytes())
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
        const workbook = await open()
        const sheets = workbook.sheets()
        const sheet = optionalString(input, "sheet") ?? sheets[0]
        return { sheets, sheet, cells: workbook.cells(sheet, range) }
      },
    },
    {
      name: "sheet_comments",
      title: "Read the workbook's comments",
      description:
        "Lists the comment threads saved in the workbook: each with its id, sheet, cell ref, whether it is resolved, whether it is a legacy note (`legacy`), and its comments with author, date and text. Suggestions are threads too, and sheet_suggestions reads them as proposed values.",
      inputSchema: noInput,
      annotations,
      execute: async () => ({ comments: (await open()).comments() }),
    },
    {
      name: "sheet_suggestions",
      title: "Read the workbook's suggestions",
      description:
        "Lists the suggested cell values saved in the workbook: each with its thread id, sheet, cell ref, the proposed value (a formula starts with =), the value it would replace, its note, author and date, and its status: open, accepted, rejected or resolved.",
      inputSchema: noInput,
      annotations,
      execute: async () => ({ suggestions: (await open()).suggestions() }),
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
