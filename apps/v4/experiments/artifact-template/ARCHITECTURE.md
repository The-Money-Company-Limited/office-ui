# Office artifact template: build and edge delivery

This is the implementation reference for the DOCX, XLSX, PPTX, and PDF artifact
experiment. It describes the source as reviewed on 2026-09-28. RuntimeVM is a
separate repository; this experiment does not change it. Its artifact service
was reviewed at
[`cb59f25`](https://github.com/The-Money-Company-Limited/runtimevm/tree/cb59f25fa4a823ef53b4619e9187db8b35275d83).

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

`artifact-template:build` runs Vite. The
[`stage.mjs`](./stage.mjs) script copies that build into a **new** directory,
adds one input file, and writes the two metadata files. Use a fresh output
directory per artifact. Run the build again after changing editor source.

```text
new-release-directory/
  index.html
  assets/...               # Vite JS, CSS, WASM, and related assets
  document/report.xlsx     # original uploaded bytes
  artifact.json            # {"kind":"xlsx","title":"report.xlsx","file":"document/report.xlsx"}
  cool.json                # {"spa":false}
```

The supported input extensions are `.docx`, `.xlsx`, `.pptx`, and `.pdf`.
Staging requires a regular file with an ASCII basename that uses letters,
digits, dots, underscores, or dashes. It checks a 25 MiB maximum per file,
20,000 files, and 1 GiB total. These match the authenticated release bounds
in RuntimeVM's
[`files.go`](https://github.com/The-Money-Company-Limited/runtimevm/blob/cb59f25fa4a823ef53b4619e9187db8b35275d83/backend/internal/artifacts/files.go).
Publishing without an account has lower limits; use an authenticated account
for this workflow. The stage command removes a new output directory if it
fails. The generated `cool.json` has no `run` command or `functions` module.

The `cool` CLI publishes the entire directory as one release. Without
`--computer`, it creates an artifact computer; `--computer <id-or-slug>`
publishes a new release to an existing one. A ZIP of the directory contents
can also be submitted through the publish UI. The output directory itself is
the release root: `index.html`, `artifact.json`, and `cool.json` must be at its
top level. The [CLI implementation](https://github.com/The-Money-Company-Limited/runtimevm/blob/cb59f25fa4a823ef53b4619e9187db8b35275d83/cli/internal/command/publish.go)
and [publish service](https://github.com/The-Money-Company-Limited/runtimevm/blob/cb59f25fa4a823ef53b4619e9187db8b35275d83/backend/internal/artifacts/publish.go)
are the source of truth for current publish behavior.

## How a request reaches the editor, with no computer VM

```text
Agent's local DOCX/XLSX/PPTX/PDF + this Vite build
  -> stage a complete static release
  -> Cool publish API: validate, store release manifest and bytes
  -> Cloudflare Workers for Platforms: deploy release Worker + Static Assets
  -> *.cool.computer gateway: resolve slug, release, and access
  -> release Worker: check computer ID/access, serve assets
  -> visitor's browser: load document, edit, export/download
```

RuntimeVM stores uploaded content in R2 and release metadata in Postgres. Its
[`deploy`](https://github.com/The-Money-Company-Limited/runtimevm/blob/cb59f25fa4a823ef53b4619e9187db8b35275d83/backend/internal/artifacts/service.go)
path sends the release's static assets to a Cloudflare Workers for Platforms
dispatch namespace. Cloudflare's asset upload session requests only missing
hashes; then RuntimeVM deploys a small **RuntimeVM-owned** Worker module with an
`ASSETS` binding. That module checks the routed computer ID and access before
calling `env.ASSETS.fetch(request)`. `run_worker_first: true` keeps these
checks in front of asset requests. See the
[`workersplatform` client](https://github.com/The-Money-Company-Limited/runtimevm/blob/cb59f25fa4a823ef53b4619e9187db8b35275d83/backend/internal/workersplatform/client.go)
and [template module](https://github.com/The-Money-Company-Limited/runtimevm/blob/cb59f25fa4a823ef53b4619e9187db8b35275d83/backend/internal/workersplatform/template.js).

The gateway uses an edge route to select the live release, a retained release
URL, or a VM route. For an artifact computer, it selects the edge release
Worker. `cool.json` is parsed as release configuration and is never served as
an asset. `"spa":false` gives missing paths a 404 instead of returning the
editor shell. See RuntimeVM's
[`routes.go`](https://github.com/The-Money-Company-Limited/runtimevm/blob/cb59f25fa4a823ef53b4619e9187db8b35275d83/backend/internal/artifacts/routes.go)
and [artifact computer reference](https://github.com/The-Money-Company-Limited/runtimevm/blob/cb59f25fa4a823ef53b4619e9187db8b35275d83/backend/docs/artifact-computers.md).

There is **no per-document Linux VM** for this static release. The publish
service and Cloudflare Workers still run server-side, and the browser performs
the document editing. A computer can gain a VM through an explicit Linux use
or a release that declares `run`; this template does neither. This experiment
was published and opened on Cloudflare in September 2026; its demo computers
were subsequently deleted, so there are no live example links to rely on.
Cloudflare also documents the underlying
[Workers for Platforms static asset flow](https://developers.cloudflare.com/cloudflare-for-platforms/workers-for-platforms/configuration/static-assets/)
and [`run_worker_first` behavior](https://developers.cloudflare.com/workers/static-assets/binding/).

## Contract for a later artifact upload feature

Today, RuntimeVM accepts a **complete release file set**, not a template ID
plus document bytes. An agent can already build/stage locally and publish the
result unchanged. To make upload automatic, the first integration can package
this pinned build with each document on the agent side and call the existing
publish API. Record the bundle version alongside the artifact so a later
editor build does not silently change old releases. Existing releases are
immutable; publishing an update creates another release.

A platform-side template registry would be a separate RuntimeVM feature: it
would need to select a versioned bundle, combine it with the document and
`artifact.json`, validate the resulting file set, then create a release. No
such registry or composition API exists in the reviewed code. The UI fetches
`artifact.json` in every visitor's browser, so do not put secrets in it.

Browser edits are session state until the user exports/downloads a file.
There is no autosave or upload-back endpoint in this template. Persisting
edited bytes would require a deliberate save flow, authorization, and a new
release or another storage contract. Public releases expose the original
document to visitors; use the computer's private access setting when the
document should be restricted.
