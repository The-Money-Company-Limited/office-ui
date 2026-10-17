# HTML slides, phase 1

An artifact is `{ "kind": "slides", "title": "…", "file": "deck/deck.json" }`.
The deck is self-contained:

```
deck/deck.json
deck/slides/<id>.html
```

`deck.json` contains `title`, `order` (1–500 unique slide IDs), and optional
`sections: [{ title, slides: [ids] }]`. IDs start with an ASCII letter, then use
letters, numbers, underscores or dashes, at most 80 characters. Section members
must occur in `order`. Paths cannot escape `deck/slides`.

Each HTML file has exactly one `<section id="<id>">`. Its canvas is always
1920×1080 CSS pixels, relative positioned, border-box and clipped to its bounds.
A single optional final `<aside>` contains speaker notes. Notes are hidden in
both editing and presenting and included as PowerPoint speaker notes.

## Closed subset

HTML: `h1 h2 h3 p ul ol li b i u a span div img table hr`. Tables additionally
allow `thead tbody tfoot tr th td`. SVG additionally allows
`g path rect circle ellipse line polyline polygon text tspan title desc`.
No scripts, stylesheets, custom elements, frames, embedded HTML or SVG animation.
Unsupported elements are dropped with their subtree. Extra top-level content is
dropped; a missing section is an error. `validate.js` is the single validator
used by loading, editing and staging. It has no React dependency; a future CLI
can supply a standards-compatible `Document` to `sanitizeSlide`.

Attributes: `id`, `title`, `style`; `a[href]` permits HTTPS, HTTP and mailto;
`img[src,alt]` permits embedded base64 PNG/JPEG/GIF/WebP only. Embed images to
make ZIPs portable and avoid cross-origin export dependencies. Table cells
allow numeric `colspan`/`rowspan` (1–99). SVG attributes are `viewBox`, geometry
(`x y x1 y1 x2 y2 cx cy r rx ry width height d points`), `fill`, `stroke`,
`stroke-width`, `stroke-linecap`, `stroke-linejoin`, `opacity`, `fill-opacity`,
`stroke-opacity`, `transform`, `text-anchor`, `dominant-baseline`, `font-size`,
`font-family`, `font-weight`. SVG URLs and resource references are forbidden.

Inline CSS properties:

- `width height min-width min-height max-width max-height`, `margin` and
  `padding` (including four sides), `gap row-gap column-gap`.
- `color background background-color font-family font-size font-weight
  font-style line-height letter-spacing text-align text-decoration
  text-transform white-space word-break overflow-wrap`.
- `display flex flex-direction flex-wrap flex-grow flex-shrink flex-basis
  align-items align-self align-content justify-content justify-items
  justify-self order`.
- `grid-template-columns grid-template-rows grid-column grid-row grid-auto-flow
  grid-auto-columns grid-auto-rows`.
- `position top right bottom left z-index`, `border` (including four sides),
  `border-width border-style border-color border-radius box-shadow transform
  transform-origin opacity overflow object-fit object-position border-collapse
  border-spacing table-layout list-style-type box-sizing`.

CSSOM-expanded longhands for border sides/corners, background, text decoration,
grid row/column start/end and overflow x/y are also accepted.

Use px lengths, CSS colors, unitless numbers and grid `fr` tracks. Percent and
font/viewport/physical length units, CSS variables, URLs, escapes and `!important`
are dropped. Allowed functions are RGB/HSL colors, 2D translate/scale/rotate/skew
and matrix, grid repeat/minmax. Positions are static, relative or absolute;
display is block/inline/inline-block/flex/inline-flex/grid/none or table/list
layout. Browser-invalid declarations are dropped too. Every rejected element,
attribute or declaration logs a warning naming the slide and element.

The sanitizer gives elements stable `<slide-id>-…` IDs; supplied IDs in that
namespace survive edits and ZIP downloads. The rendered section retains the
slide ID and its container exposes `data-slide-id`. Thumbnails contain no IDs.
`#<slide-id>` links scroll to the real slide. The document scrolls the window;
only the fixed slide strip scrolls independently.

## Editing and exports

Click selects; drag pins an element at its measured box; eight handles resize.
Shift preserves the ratio. Arrow keys nudge 1px (Shift: 10px). Double-click or
Enter edits text. Escape ends text editing or clears selection. Delete removes;
Cmd/Ctrl+D duplicates; Cmd/Ctrl+Z and Shift+Cmd/Ctrl+Z undo/redo. Drag snapping
uses slide and other element edges/centres. History retains 100 edits in memory.
Text entry uses plaintext paste; inline formatting already in files is retained
until that text is replaced. There is no save path (COD-1462).

Download produces a ZIP of the current sanitized `deck/` files. Export uses
PptxGenJS 4.0.1 with real text boxes, simple shapes, embedded images and tables.
SVG, lists, merged tables and effects requiring rasterization become PNGs.
PowerPoint typography/layout is an approximation; complex effects are rendered
at slide resolution. Present uses fullscreen when available, otherwise a
viewport presentation; arrows, space and PageDown/PageUp navigate, Escape exits.

Dependencies are pinned in this directory's `package.json`/`package-lock.json`
to keep all new dependency files within the permitted slides path. Install with
`npm ci --prefix apps/v4/components/extend/slides --ignore-scripts` before the
app build. No external CDN or network request is needed to export.

See [tests/README.md](./tests/README.md) for the exact Modal build and browser
verification commands, including the negative control.

Phase 2: agent self-check CLI, save integration (COD-1462), live data and a PPTX
to HTML import pipeline. The existing read-only PPTX artifact viewer remains
available for legacy files.
