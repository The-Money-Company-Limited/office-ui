#!/bin/sh
# Run from the laptop; all compilation happens in the shared Modal sandbox.
set -eu
root=$(CDPATH= cd -- "$(dirname -- "$0")/.." && pwd)
remote="$root/../office-remote/r"
"$remote" up
"$remote" sync duke-sheets
"$remote" run duke-sheets 'wasm-pack build bindings/wasm --target web --release --out-dir pkg'
"$remote" pull duke-sheets bindings/wasm/pkg
cp "$root/../duke-sheets/bindings/wasm/pkg/duke_sheets_wasm.js" "$root/packages/react-xlsx/vendor/sheets-wasm/"
cp "$root/../duke-sheets/bindings/wasm/pkg/duke_sheets_wasm.d.ts" "$root/packages/react-xlsx/vendor/sheets-wasm/"
cp "$root/../duke-sheets/bindings/wasm/pkg/duke_sheets_wasm_bg.wasm" "$root/packages/react-xlsx/vendor/sheets-wasm/"
cp "$root/../duke-sheets/bindings/wasm/pkg/duke_sheets_wasm_bg.wasm.d.ts" "$root/packages/react-xlsx/vendor/sheets-wasm/"
cp "$root/../duke-sheets/LICENSE" "$root/packages/react-xlsx/vendor/sheets-wasm/LICENSE"
