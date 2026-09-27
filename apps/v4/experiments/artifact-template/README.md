# Office artifact template experiment

This fork builds a static release for a Cool Computer artifact. Each release
contains the viewer/editor bundle, one document, `artifact.json`, and
`cool.json`. It runs in the browser without a computer VM or a RuntimeVM
code change.

The document fills the viewport edge to edge. Its editor controls use a copy
of RuntimeVM's current light and dark appearance tokens, with the spreadsheet
toolbar styled in place. The workbook's own appearance button persists the
choice under RuntimeVM's `cool-computers-theme` key. There is no artifact
header, logo, wallpaper, or outer padding.

| File | UI in this fork                                                               | Save path              |
| ---- | ----------------------------------------------------------------------------- | ---------------------- |
| XLSX | Experimental spreadsheet editor, including chart selection and movement       | Export XLSX            |
| DOCX | Experimental word processor                                                   | Export DOCX            |
| PPTX | Presentation viewer only; this fork has no PPTX editor                        | Download original PPTX |
| PDF  | PDF editor for annotations, forms, signatures, redaction, and page operations | Download PDF           |

The PDF editor does not rewrite arbitrary existing page text. Edits remain in
the browser session until exported; this template has no server-side save API.
The XLSX editor switches to read-only above 5 MiB. The bundled PDFium WASM is
served from the artifact itself; optional fonts and stamp resources may still
use external URLs.

Casual Slides is the selected candidate for a future editable PPTX experiment;
its integration is deferred while the ready DOCX, XLSX, and PDF editors are
prioritized.

## Build and stage

From the repository root:

```sh
pnpm install
pnpm --filter v4 artifact-template:build
pnpm --filter v4 artifact-template:stage /absolute/path/to/report.xlsx /absolute/path/to/new-release-directory
```

The stage command accepts `.xlsx`, `.docx`, `.pptx`, and `.pdf`. An omitted
output directory stages directly into
`apps/v4/experiments/artifact-template/dist`; a specified output directory
must not already exist. Rebuild after changing source. The release is a complete
static directory; Vite emits relative asset URLs so it works at an artifact
root or a local preview subpath. The stage command checks the 25 MiB per-asset,
20,000-file, and 1 GiB total limits. A failed stage removes a newly created
output directory.

To preview several releases from a common parent:

```sh
python3 -m http.server 4173 --bind 127.0.0.1 --directory /absolute/path/to/releases
```

To publish one release when ready:

```sh
cool publish /absolute/path/to/new-release-directory --name example
```

The corresponding ZIP may also be uploaded through the Cool Computers publish
UI. This experiment has not published to a Worker.

## RuntimeVM fit

The current RuntimeVM artifact API accepts complete static file sets rather
than a template identifier. A built-in upload feature would need RuntimeVM to
keep a versioned bundle and combine it with the uploaded document before
creating a release. This experiment leaves RuntimeVM untouched.
