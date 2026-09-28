import * as React from "react"
import { createRoot } from "react-dom/client"

import "../../../app/globals.css"
import "./theme.css"

const DocxEditor = React.lazy(() =>
  import("@/components/extend/docx-editor").then((module) => ({
    default: module.DocxEditorPreview,
  }))
)
const XlsxEditor = React.lazy(() => import("./xlsx-artifact"))
const PptxViewer = React.lazy(() =>
  import("@/components/extend/pptx-viewer").then((module) => ({
    default: module.PptxViewerPreview,
  }))
)
const PdfEditor = React.lazy(() =>
  import("@/components/extend/pdf-editor").then((module) => ({
    default: module.PDFEditor,
  }))
)

type Artifact = {
  kind: "docx" | "xlsx" | "pptx" | "pdf"
  title: string
  file: string
}

function parseArtifact(value: unknown): Artifact {
  if (!value || typeof value !== "object")
    throw new Error("Invalid artifact.json")
  const artifact = value as Record<string, unknown>
  if (!["docx", "xlsx", "pptx", "pdf"].includes(String(artifact.kind)))
    throw new Error("Unsupported artifact type")
  if (typeof artifact.title !== "string" || !artifact.title.trim())
    throw new Error("Missing artifact title")
  if (
    typeof artifact.file !== "string" ||
    !/^document\/[a-zA-Z0-9][a-zA-Z0-9._-]*\.(docx|xlsx|pptx|pdf)$/.test(
      artifact.file
    ) ||
    !artifact.file.endsWith(`.${artifact.kind}`)
  )
    throw new Error("Invalid artifact file path")
  return artifact as Artifact
}

function storedTheme(): "dark" | "light" {
  try {
    return localStorage.getItem("cool-computers-theme") === "light"
      ? "light"
      : "dark"
  } catch {
    return "dark"
  }
}

function App() {
  const [artifact, setArtifact] = React.useState<Artifact | null>(null)
  const [error, setError] = React.useState<string | null>(null)
  const [theme, setTheme] = React.useState(storedTheme)
  const isDark = theme === "dark"

  React.useLayoutEffect(() => {
    document.documentElement.classList.toggle("dark", isDark)
    document.documentElement.style.colorScheme = theme
    try {
      localStorage.setItem("cool-computers-theme", theme)
    } catch {
      // Keep the appearance usable when storage is unavailable.
    }
  }, [isDark, theme])

  React.useEffect(() => {
    let current = true
    fetch("./artifact.json", { cache: "no-store" })
      .then(async (response) => {
        if (!response.ok)
          throw new Error(`Artifact metadata: HTTP ${response.status}`)
        return parseArtifact(await response.json())
      })
      .then((value) => {
        if (current) {
          setArtifact(value)
          document.title = `${value.title} · Cool Computers`
        }
      })
      .catch((reason: unknown) => {
        if (current)
          setError(
            reason instanceof Error ? reason.message : "Could not load artifact"
          )
      })
    return () => {
      current = false
    }
  }, [])

  const documentProps = artifact
    ? {
        className: `artifact-${artifact.kind} h-full min-h-0`,
        fileName: artifact.file.split("/").at(-1),
        src: `./${artifact.file}`,
      }
    : null

  return (
    <main className="artifact-editor">
      {error ? (
        <div className="artifact-state" role="alert">
          <strong>Could not open this document</strong>
          <span>{error}</span>
          <button type="button" onClick={() => window.location.reload()}>
            Reload
          </button>
        </div>
      ) : documentProps ? (
        <React.Suspense
          fallback={<div className="artifact-state">Opening document…</div>}
        >
          {artifact?.kind === "docx" ? (
            <DocxEditor
              {...documentProps}
              isDark={isDark}
              onIsDarkChange={(value) => setTheme(value ? "dark" : "light")}
            />
          ) : artifact?.kind === "xlsx" ? (
            <XlsxEditor
              {...documentProps}
              isDark={isDark}
              onIsDarkChange={(value) => setTheme(value ? "dark" : "light")}
            />
          ) : artifact?.kind === "pptx" ? (
            <PptxViewer {...documentProps} showUpload={false} />
          ) : (
            <PdfEditor {...documentProps} showUpload={false} />
          )}
        </React.Suspense>
      ) : (
        <div className="artifact-state">Opening artifact…</div>
      )}
    </main>
  )
}

createRoot(document.getElementById("root")!).render(<App />)
