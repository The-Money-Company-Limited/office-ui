# Extend UI

[Extend UI](https://www.extend.ai/ui) is a set of open-source components for
document agents, user-facing document flows, and internal tools.

![Extend UI document components](apps/v4/public/extend-ui-readme.png)

The library includes 14 components and examples for PDF, DOCX, and XLSX
viewers, plus bounding box citations, file upload, file thumbnails,
e-signature, document splitting, layout blocks, and Finder-style file browsing.
Components are fully customizable, MIT licensed, and available through the
shadcn component registry.

Extend UI started as the internal component layer for
[Extend](https://www.extend.ai). We built it after evaluating existing file
viewer and document component libraries and needing more complete, polished
building blocks for document-heavy products. The components are maintained in
production at Extend and are used across workflows that process millions of
pages, so fixes and improvements continue to flow back into the project.

Use Extend UI when you need document viewers, review surfaces, file-management
interfaces, or agent-facing document tools without starting from a generic
viewer library. Because each component is installed as source, it also works
well with design-system agents and prototyping tools that need editable React
and Tailwind code.

## Links

- Documentation: [https://www.extend.ai/ui](https://www.extend.ai/ui)
- GitHub: [extend-hq/ui](https://github.com/extend-hq/ui)
- Registry namespace: `@extend/*`

## Getting Started

Configure the style-aware registry in your `components.json`:

```json
{
  "registries": {
    "@extend": "https://www.extend.ai/ui/r/styles/{style}/{name}.json"
  }
}
```

Then install a component with the shadcn CLI:

```bash
npx shadcn@latest add @extend/pdf-viewer
```

Then render the installed component from your app:

```tsx
import { PDFViewer } from "@/components/extend/pdf-viewer"

export default function Page() {
  return <PDFViewer src="/sample.pdf" className="h-[720px]" />
}
```

The `react-pdf` version of the PDF viewer is available at
[pdf-viewer.tsx](https://github.com/extend-hq/ui/blob/2dd84954269e817a2c17a44c672ae801df1bfc3d/apps/v4/components/ui/pdf-viewer.tsx).

Extend UI components are copied into your project as source, so you can adapt
them to your app. Shared primitives such as `Button`, `Select`, `Dialog`,
and `Tooltip` are imported directly from your existing UI components. The
registry selects Base or Radix source using your configured style and transforms
icons for your selected icon library. Keep existing primitives when the CLI
asks about overwriting them. If your project uses a different alias or design-system path, update the
generated imports to match, for example changing
`@/components/ui/button` or `@/components/ui/select` to your local primitive
paths. You can also set those aliases in `components.json` before installing so
new components are generated closer to your app structure.

Use current shadcn primitive APIs. The historical `default` style is
[deprecated upstream](https://ui.shadcn.com/docs/components-json#style) and must
be migrated before installing.

## Examples

### Layout Blocks

Inspect OCR and layout output with selectable blocks, confidence, markdown text,
and PDF overlays that stay connected to the rendered page.

![Layout Blocks component](apps/v4/public/readme/layout-blocks.png)

```bash
npx shadcn@latest add @extend/layout-blocks-block
```

```tsx
import type { ParsedOcrOutput } from "@/components/extend/layout-blocks"
import { OcrBlocksBlock } from "@/components/blocks/layout-blocks-block"

const output: ParsedOcrOutput = {
  chunks: [
    {
      blocks: [
        {
          id: "title",
          type: "heading",
          content: "Statement of Work",
          metadata: {
            page: { number: 1, width: 612, height: 792 },
            avgOcrConfidence: 0.99,
          },
          boundingBox: { left: 72, top: 96, right: 360, bottom: 124 },
        },
      ],
    },
  ],
}

export default function Page() {
  return <OcrBlocksBlock file="/documents/statement.pdf" output={output} />
}
```

### Bounding Box Citations

Review extracted values against the source PDF with field-level citations,
editable form controls, JSON diffing, and page overlays.

![Bounding Box Citations component](apps/v4/public/readme/bounding-box-citations.png)

```bash
npx shadcn@latest add @extend/bounding-box-citations-block
```

```tsx
import type { ReviewField } from "@/components/extend/bounding-box-citations"
import { HumanReviewBlock } from "@/components/blocks/bounding-box-citations-block"

const fields: ReviewField[] = [
  {
    key: "invoice_total",
    schema: {
      type: "number",
      title: "Invoice total",
      description: "Total amount due on the invoice.",
    },
    actual: 1280.5,
    expected: 1280.5,
    location: {
      page: 1,
      area: { left: 62, top: 82, width: 22, height: 4 },
    },
  },
]

export default function Page() {
  return (
    <HumanReviewBlock
      file="/documents/invoice.pdf"
      fields={fields}
      className="h-[720px]"
    />
  )
}
```

### File System

A Finder-style file browser for object-store manifests. It supports icon, list,
column, and gallery views, and opens PDF, DOCX, XLSX, and image files in viewer
dialogs.

![File System Finder component](apps/v4/public/readme/file-system-finder.png)

```bash
npx shadcn@latest add @extend/file-system-block
```

```tsx
import type { FileSystemItem } from "@/components/extend/file-system"
import { FileSystemBlock } from "@/components/blocks/file-system-block"

const items: FileSystemItem[] = [
  {
    kind: "folder",
    path: "reports/",
    hasChildren: true,
  },
  {
    kind: "file",
    path: "bank-statement.pdf",
    contentType: "application/pdf",
    url: "/documents/bank-statement.pdf",
    previewImageUrl: "/documents/bank-statement-preview.png",
  },
]

export default function Page() {
  return (
    <FileSystemBlock
      items={items}
      title="Documents"
      defaultView="icons"
      className="h-[720px]"
      getFileUrl={(file) =>
        `/api/files/sign?path=${encodeURIComponent(file.path)}`
      }
      loadChildren={async ({ path }) => {
        const response = await fetch(
          `/api/files?prefix=${encodeURIComponent(path)}`
        )
        return response.json()
      }}
    />
  )
}
```

## Popular Components

```bash
npx shadcn@latest add @extend/pdf-viewer
npx shadcn@latest add @extend/docx-viewer
npx shadcn@latest add @extend/xlsx-viewer
npx shadcn@latest add @extend/file-upload
npx shadcn@latest add @extend/file-thumbnail
npx shadcn@latest add @extend/e-signature
```

## Development

The app and the three engine forks share one pnpm 9.15.0 workspace and one
root lockfile. The forks live in `engines/react-docx`, `engines/react-xlsx`,
and `engines/react-pptx`, imported with full history using Git subtrees.
`apps/v4` consumes their packages with `workspace:*` dependencies.
`engines/duke-sheets` is the Rust spreadsheet engine behind XLSX, also a
subtree with full history; it is not a pnpm package, and react-xlsx vendors
its wasm build.

From the repository root, on Modal:

```bash
pnpm install
pnpm build
```

The root build uses Turbo's workspace dependency graph to build the engines
before v4. Engine root packages provide build/test tooling; their playground
apps are excluded from the workspace. After building, `pnpm v4:dev` serves
the site at `http://localhost:4000`. The existing v4 scripts also work,
including `pnpm --filter v4 build` after the engines have been built.

Builds, tests, and Rust/wasm rebuilds run on Modal, not on the laptop, through
`tools/modal/r`. From the root of the checkout (main or worktree), `.` names it.
In a worktree `.git` is a file and is not uploaded, so builds that stamp
`git describe` need a `git init` in the sandbox copy first:

```bash
r=tools/modal/r
repo=.
"$r" up
"$r" sync "$repo"
"$r" run "$repo" 'pnpm install --frozen-lockfile && pnpm build'
"$r" run "$repo" 'pnpm --filter v4 build'
"$r" run "$repo" 'cd engines/react-docx && pnpm exec vitest run'
"$r" run "$repo" 'cd engines/react-xlsx && pnpm test'
"$r" run "$repo" 'cd engines/react-pptx && pnpm exec vitest run'
"$r" down
```

When intentionally updating dependencies, use `pnpm install --no-frozen-lockfile`
on Modal and `"$r" pull "$repo" pnpm-lock.yaml` before shutting it down.
Normal CI installs use the checked-in root lockfile.

DOCX and PPTX ordinary builds use their checked-in wasm and generated JavaScript.
PPTX keeps its generated runtime in `packages/wasm/src/generated/`; `pkg/` is
temporary wasm-pack output and stays ignored. XLSX copies its vendored Duke Sheets
wasm via `scripts/copy-duke-wasm.mjs`. A fresh clone installs and builds without
Rust. To update these binaries, run the explicit rebuild commands on Modal:

```bash
pnpm --filter @extend-ai/react-docx-wasm build:wasm
pnpm --filter @extend-ai/react-pptx-wasm build:wasm
(cd engines/react-xlsx && bash scripts/rebuild-duke-wasm.sh)
```

Rebuilds require Rust, the `wasm32-unknown-unknown` target, and `wasm-pack`
(the remote helper supplies them). Pull the generated files back before shutting
Modal down. For PPTX, from the laptop using the helper variables above:

```bash
"$r" run "$repo" 'pnpm --filter @extend-ai/react-pptx-wasm build:wasm'
"$r" pull "$repo" engines/react-pptx/packages/wasm/src/generated
```

Include the generated JavaScript, declarations, and `.wasm` binary in the same
change as the Rust sources. Ordinary installs and builds never invoke wasm-pack.

### Updating engine subtrees

From the office-ui root, pull upstream changes (replace `<name>` with
`react-docx`, `react-xlsx`, or `react-pptx`):

```bash
git subtree pull --prefix=engines/<name> https://github.com/extend-hq/<name>.git main
git subtree pull --prefix=engines/duke-sheets https://github.com/guseggert/duke-sheets.git main
```

Keep full history: do not add `--squash`. Resolve workspace/configuration
conflicts by retaining the single root workspace and lockfile, merging new
pnpm settings into the root, and removing any reintroduced engine lockfiles
or workspace files. Install/build/test on Modal and bring the root lockfile
back before committing the update.

Engine changes are made here; the old The-Money-Company-Limited forks are
archived. To offer a fix upstream, push its subtree to a branch of a fork of
the upstream repo and open a PR there:

```bash
git subtree push --prefix=engines/<name> <your-fork-url> <fix-branch>
```

### Shared pnpm settings

- pnpm is pinned to 9.15.0 (office-ui previously used 9.0.6).
- Engine packages declare their own `tsup` and TypeScript build tools so Turbo
  can build them directly without relying on an engine root’s executable path.
- The app's React type pins stay at 19.2.2; they satisfy the engines' `^19.1`
  ranges. Engine dependencies use workspace references instead of sibling links.
- DOCX's `lodash` 4.18.1 and scoped `picomatch` 4.0.4 overrides are retained.
  Its global Vite 6.4.2 override is scoped to `react-docx-monorepo>vite`,
  with an explicit DOCX tooling dependency on 6.4.2. This preserves v4's
  Vite 7.3.2. PPTX now declares its Vite 8 tooling directly because its tests
  use Vite's native tsconfig-path support; it previously came from the playground.
- PPTX's `esbuild` 0.28.1 override is retained repo-wide, including packages
  whose original dependency range selected an older esbuild. Its
  `onlyBuiltDependencies: [esbuild]` allowlist is also retained at the root.
- The `@chenglou/pretext@0.0.8` patch stays in the DOCX subtree; the root
  `patchedDependencies` path includes `engines/react-docx/`.
- DOCX's install-time `prepare` hook was removed: setting `core.hooksPath`
  from a subtree would change office-ui's Git hook configuration.

## Created By

Extend UI is built and maintained by [Extend](https://www.extend.ai) for teams
building modern document processing products.

## License

Licensed under the [MIT license](./LICENSE.md).
