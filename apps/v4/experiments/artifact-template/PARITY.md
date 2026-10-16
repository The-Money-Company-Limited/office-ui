# Everyday parity with Google Docs, Sheets and Slides

What a person expects without thinking about it, checked by hand in Chrome
against the published demos. Status: ✅ works, 🔧 fixed (commit), ❌ missing,
⛔ needs the engine (no API to build it on).

## Sheets (XLSX)

| Behaviour | Status |
| --- | --- |
| Click a cell selects it; arrows, Shift+arrows, Home/End, PageUp/Down move | ✅ |
| Type to replace, F2 or double-click to edit in place, Escape cancels | ✅ |
| Enter saves and moves down, Tab right (Shift: up, left) | 🔧 react-xlsx `Enter and Tab in a cell save and move` |
| Backspace/Delete clear, Cmd+Z/Shift+Cmd+Z undo/redo | ✅ |
| Cmd+C/X/V copy, cut, paste | ✅ |
| Cmd+B/I/U | 🔧 office-ui `Cmd+B/I/U format cells` |
| Typed `=SUM(...)` calculates | 🔧 react-xlsx `Typing a formula into a cell calculates it` |
| Function suggestions and argument help | 🔧 react-xlsx `Formula suggestions and argument help` |
| Fill handle drag, column resize | ✅ |
| + adds a sheet and its tab comes into view | 🔧 office-ui |
| Right-click menu (cut, copy, paste, clear, comment); cut moves on paste | 🔧 react-xlsx + office-ui |
| Sum / Average / Count of the selection | 🔧 react-xlsx + office-ui |
| Cmd+F find in the sheet | 🔧 react-xlsx + office-ui |
| Name box shows the selection; typing a reference jumps there | 🔧 office-ui |
| A sheet opens with its saved active cell (or A1) selected | 🔧 react-xlsx |
| Rename, reorder sheets | ⛔ duke-sheets has no rename/move |
| Insert / delete rows and columns | ⛔ duke-sheets has no insert/delete |

## Docs (DOCX)

| Behaviour | Status |
| --- | --- |
| Typing, Backspace, selection, double-click word | ✅ |
| Enter at the end of a page keeps the caret (no page jump) | 🔧 react-docx `A blank paragraph before a page break stays on its page` |
| Cmd+B/I/U | 🔧 react-docx `Cmd/Ctrl+B, I and U format the selection` |
| Cmd+K link | 🔧 office-ui |
| Cmd+Z undo, Cmd+A select all | ✅ |
| `- `, `* `, `1. ` at a line start begin a list | 🔧 react-docx (COD-1457) |
| Tab / Shift+Tab change list level; Enter on an empty item outdents / ends | 🔧 react-docx (COD-1457) |
| Cmd+F find, Cmd+Shift+H replace (Editing and Suggesting) | 🔧 react-docx + office-ui (COD-1458) |
| One toolbar row, page fits narrow windows, Viewing hides formatting | 🔧 office-ui |

## Slides (PPTX)

| Behaviour | Status |
| --- | --- |
| Slide strip open beside the slides on wide windows, grey canvas | 🔧 office-ui |
| Click a thumbnail jumps to the slide | ✅ |
| Strip keeps the active thumbnail in view | 🔧 office-ui |
| Arrows / PageDown / Space / Home / End change slide | 🔧 office-ui |
| Present full screen | 🔧 office-ui |

## PDF

| Behaviour | Status |
| --- | --- |
| One toolbar row, mode dropdown, tools only outside View | 🔧 office-ui |
| PageDown / scrolling, page counter | ✅ |
| Cmd+F search | ✅ |

## Engine decision (Sheets)

Insert/delete rows and columns, rename and reorder sheets need structural
edits inside duke-sheets (Rust): shifting cells, styles, merges, comments,
charts and rewriting formula references across sheets. Codex Astra's
recommendation: add them to a fork in Rust and offer them upstream, about one
to two weeks for a prototype, six to twelve for a hardened release; emulating
them in JavaScript would amount to a second engine.
