# Office artifact template: build and edge delivery

This is the implementation reference for the DOCX, XLSX, PPTX, and PDF artifact
experiment. RuntimeVM is a separate repository; this experiment does not
change it. Its artifact service was last reviewed on 2026-10-16 at
[`ff898d1`](https://github.com/The-Money-Company-Limited/runtimevm/tree/ff898d124b08dcfb47a0ca06ddebc63fe1ee5d8b);
[`artifact-computers.md`](https://github.com/The-Money-Company-Limited/runtimevm/blob/ff898d124b08dcfb47a0ca06ddebc63fe1ee5d8b/backend/docs/artifact-computers.md) there is the
source of truth for how a published site is served.

## What is built

One Vite build produces a reusable static browser app. It contains React,
styles, editor code, and the WASM needed by the spreadsheet and PDF editors.
The [Vite config](./vite.config.ts) emits relative URLs (`base: "./"`), so the
same build can be copied into separate releases. The app
[loads `artifact.json`](./src/main.tsx) at startup, selects the editor by
`kind`, then loads the document from `./document/<filename>`. Editor modules
are lazy loaded by file type; the release still includes all four.

| Format | Browser component                                                                     | Current capability                                                                                 |
| ------ | ------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------- |
| DOCX   | [`docx-editor.tsx`](../../components/extend/docx-editor.tsx), `@extend-ai/react-docx` | Edit and export DOCX                                                                               |
| XLSX   | [`xlsx-editor.tsx`](../../components/extend/xlsx-editor.tsx), `@extend-ai/react-xlsx` | Edit and export XLSX; files above 5 MiB open read-only                                             |
| PPTX   | [`pptx-viewer.tsx`](../../components/extend/pptx-viewer.tsx), `@extend-ai/react-pptx` | View and download the original; no slide editing                                                   |
| PDF    | [`pdf-editor.tsx`](../../components/extend/pdf-editor.tsx), EmbedPDF and PDFium       | Annotate, fill forms, sign, redact, change pages, and download; no arbitrary existing-text editing |

The XLSX and PDF WASM URLs are set in
[`xlsx-artifact.tsx`](./src/xlsx-artifact.tsx) and
[`pdf-engine.ts`](./src/pdf-engine.ts). Their computation happens in the
visitor's browser. A browser Web Worker is a client-side thread; it is not a
Cool Computer VM or a Cloudflare Worker. Some optional PDF fonts and stamps
may fetch external resources, so a fully self-contained/offline release is not
yet guaranteed.

## Make a release

From the repository root, with Node and pnpm installed:

```sh
pnpm install
pnpm --filter v4 artifact-template:build
pnpm --filter v4 artifact-template:stage /absolute/path/to/report.xlsx /absolute/path/to/new-release-directory
cool publish /absolute/path/to/new-release-directory --name example
```

`artifact-template:build` runs Vite and writes the **bundle**: everything a
release needs except the document. The [`stage.mjs`](./stage.mjs) script copies
the bundle into a **new** directory, adds one input file, and writes the two
per-document files. Use a fresh output directory per artifact. Run the build
again after changing editor source.

```text
new-release-directory/
  index.html
  assets/...               # Vite JS, CSS, WASM, and related assets
  template.json            # bundle: {"template":"office","version":"<git describe>"}
  document/report.xlsx     # document: original uploaded bytes
  artifact.json            # document: {"kind":"xlsx","title":"report.xlsx","file":"document/report.xlsx"}
  cool.json                # document: {"spa":false}
```

`template.json` is written by the build and never by staging, so every release
staged from one build names the same bundle version. The version is
`git describe --always --dirty` of this repository at build time; a `-dirty`
version is a local build and is not meant for publishing.

The supported input extensions are `.docx`, `.xlsx`, `.pptx`, and `.pdf`.
Staging requires a regular file with an ASCII basename that uses letters,
digits, dots, underscores, or dashes. It checks a 25 MiB maximum per file,
20,000 files, and 1 GiB total, RuntimeVM's authenticated release bounds in
[`files.go`](https://github.com/The-Money-Company-Limited/runtimevm/blob/ff898d124b08dcfb47a0ca06ddebc63fe1ee5d8b/backend/internal/artifacts/files.go).
Publishing without an account has lower limits; use an authenticated account
for this workflow. The stage command removes a new output directory if it
fails. The generated `cool.json` has no `run`, `functions`, or `secrets`.

Every file type the bundle emits (`.html`, `.js`, `.css`, `.wasm`, `.json`) and
the four document types are in RuntimeVM's pinned content-type table
([`siteasset.go`](https://github.com/The-Money-Company-Limited/runtimevm/blob/ff898d124b08dcfb47a0ca06ddebc63fe1ee5d8b/infra/runtime-image/frontdoor/siteasset/siteasset.go)),
and no bundle path is hidden from serving (no leading `.`, no `node_modules`).
The bundle is about 19 MiB in 30 files, mostly the XLSX and PDF engines.

## How a request reaches the editor, with no computer VM

```text
Agent's local DOCX/XLSX/PPTX/PDF + this Vite build
  -> stage a complete static release
  -> Cool publish API: validate, store content-addressed bytes in R2, write the release index
  -> *.cool.computer gateway: read the slug's edge route, check access
  -> gateway Release entrypoint: resolve the path in the release index, stream the file from R2
  -> visitor's browser: load artifact.json, then the document; edit, export/download
```

A static release makes **no Cloudflare deploy**. Publishing stores each file
once in R2 under the computer, keyed by its sha256, and writes one immutable
release index (path to R2 key, ETag, content type). The gateway Worker serves
every file from that index, through Workers Cache, after its access check. Only
a release that declares `functions` deploys a Workers for Platforms script.
`cool.json` is read as release configuration and never served. `"spa":false`
answers missing paths with a 404 instead of the editor shell.

For an HTML page, the gateway also injects the site comment layer's script
while the computer's comments switch is on (it is on by default). The layer
anchors a note to a clicked DOM element; on the XLSX and PDF canvases that
anchor is coarse. See [Open questions](#open-questions).

There is **no per-document Linux VM**. The publish service and the gateway run
server-side; the browser does the document work. A computer gains a VM only
through an explicit Linux use or a release that declares `run`, and this
template does neither. If the computer later gains a VM, its site moves there
with the same files and URL.

## How an agent changes a published document

Releases are immutable; every change is a new release with its own URL
`r<n>--<slug>` while it is among the newest 20, so a rollback is instant.

- **Republish**: stage again with the new document and
  `cool publish <dir> --computer <slug>`. Unchanged bundle files are not uploaded
  again, because content is addressed by hash.
- **Replace one file**: on a computer without a VM, the files API
  (`cool files`) and the RuntimeVM operator's `write` tool publish a release
  with one file changed, under the computer's publish claim, so concurrent
  writes line up. An agent replaces `document/<file>` and leaves the bundle and
  `artifact.json` alone. The operator reads site files up to 6 MiB.

The app fetches `artifact.json` with `no-store`; the document's URL and ETag
change with the release, so a reload shows the new release.

## Contract for a later artifact template feature

A RuntimeVM "templated artifact" composes two parts into one ordinary release:

1. **Bundle**: this build's output, identified by `template.json`. The
   platform keeps each version it offers, unchanged.
2. **Document**: `document/<file>`, `artifact.json`, and `cool.json`, as
   [`stage.mjs`](./stage.mjs) writes them. Its checks (kind from the extension,
   the file-name grammar, size limits) are the composition rules to port.

No RuntimeVM change is needed to publish a composed release; the release is
exactly what `stage.mjs` produces. What a platform-side feature adds is
choosing a bundle version and composing on the server, so an upload of
`report.xlsx` alone becomes an artifact. Existing releases keep the bundle they
were published with; moving an artifact to a newer bundle is a new release.

Browser edits are session state until the user exports a file; there is no
autosave or save-back endpoint. Saving edited bytes needs a deliberate save
flow, authorization, and a new release (the files API path above already
does this for the owner). Public releases expose the original document to
visitors; use the computer's private access setting when the document should
be restricted. `artifact.json` is fetched by every visitor, so it holds no
secrets.

## Open questions

- **Two comment systems.** Comments stored in the document (DOCX
  `comments.xml`, XLSX threaded comments) round-trip through Office. The
  RuntimeVM site comment layer stores notes in Postgres and routes them to the
  owner's agent. A template page either opts out of the layer or bridges the
  two; RuntimeVM has no per-page opt-out today.
- **Bundle storage per computer.** R2 content is addressed per computer, so
  every new document artifact uploads the 19 MiB bundle once. A platform-side
  template could share bundle bytes across computers.
- **External resources.** Optional PDF fonts and stamps can still load from
  external URLs, so a private release is not fully self-contained.
