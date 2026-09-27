import { setWasmSource } from "@extend-ai/react-xlsx"
import sheetsWasmUrl from "@extend-ai/react-xlsx/duke_sheets_wasm_bg.wasm?url"

import { XlsxEditorPreview } from "@/components/extend/xlsx-editor"

setWasmSource(sheetsWasmUrl)

export default XlsxEditorPreview
