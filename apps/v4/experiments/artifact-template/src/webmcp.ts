// Read-only WebMCP tools (COD-1456): an agent in the person's browser can read
// the open document as it is now, unsaved edits included. Tools register only
// where the browser has `document.modelContext`; they never write, so changes
// stay with the people on the page. Document text is marked untrusted, since it
// can carry instructions aimed at the agent reading it.

type ModelContext = {
  registerTool(
    tool: {
      name: string
      title: string
      description: string
      inputSchema: object
      annotations: { readOnlyHint: boolean; untrustedContentHint: boolean }
      execute: (input: Record<string, unknown>) => Promise<unknown>
    },
    options: { signal: AbortSignal }
  ): Promise<void>
}

type Kind = "docx" | "xlsx" | "pptx" | "pdf" | "slides"

const annotations = { readOnlyHint: true, untrustedContentHint: true }

// Registers the tools for this kind of document and returns their unregistration.
export function registerDocumentTools(
  kind: Kind,
  read: () => Promise<Uint8Array | null>
): () => void {
  const context = (document as Document & { modelContext?: ModelContext })
    .modelContext
  const controller = new AbortController()
  if (!context || (kind !== "docx" && kind !== "xlsx")) {
    return () => controller.abort()
  }
  const bytes = async () => {
    const current = await read()
    if (!current) throw new Error("The document is still opening.")
    return current
  }
  const tools =
    kind === "docx"
      ? [
          {
            name: "doc_read",
            title: "Read the document",
            description:
              "Returns the document's paragraphs in order, each with its id and text, as they are on the page now.",
            inputSchema: { type: "object", properties: {} },
            annotations,
            execute: async () => {
              const { openDocx } = await import("@extend-ai/react-docx/agent")
              return { paragraphs: (await openDocx(await bytes())).paragraphs() }
            },
          },
          {
            name: "doc_find",
            title: "Find text in the document",
            description:
              "Finds every occurrence of the text and returns each as a paragraph id with start and end offsets into that paragraph's text.",
            inputSchema: {
              type: "object",
              properties: { text: { type: "string", minLength: 1 } },
              required: ["text"],
            },
            annotations,
            execute: async (input: Record<string, unknown>) => {
              if (typeof input.text !== "string" || !input.text)
                throw new Error("text must be a non-empty string.")
              const { openDocx } = await import("@extend-ai/react-docx/agent")
              return { matches: (await openDocx(await bytes())).find(input.text) }
            },
          },
        ]
      : [
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
            execute: async (input: Record<string, unknown>) => {
              const { openXlsx } = await import("@extend-ai/react-xlsx/agent")
              const workbook = openXlsx(await bytes())
              const sheets = workbook.sheets()
              const sheet = typeof input.sheet === "string" ? input.sheet : sheets[0]
              const range = typeof input.range === "string" ? input.range : undefined
              return { sheets, sheet, cells: workbook.cells(sheet, range) }
            },
          },
        ]
  for (const tool of tools)
    void context
      .registerTool(tool, { signal: controller.signal })
      .catch((error: unknown) =>
        console.warn(`WebMCP tool ${tool.name} was not registered`, error)
      )
  return () => controller.abort()
}
