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
| Right-click menu (cut, copy, paste, clear, comment) | ❌ |
| Sum / Average / Count of the selection | ❌ |
| Cmd+F find in the sheet | ❌ |
| Name box shows the anchor cell; typing a reference jumps there | ❌ |
| A sheet opens with A1 selected | ❌ (host-side attempt broke the first click; belongs in the grid) |
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
| `- `, `* `, `1. ` at a line start begin a list | ❌ |
| Tab / Shift+Tab change list level | unchecked |
| Cmd+F find | unchecked |
| One toolbar row, page fits narrow windows, Viewing hides formatting | 🔧 office-ui |

## Slides (PPTX) and PDF

Not yet checked item by item: arrow keys between slides, thumbnails,
present full screen; PDF page keys, Cmd+F, text selection and copy.
