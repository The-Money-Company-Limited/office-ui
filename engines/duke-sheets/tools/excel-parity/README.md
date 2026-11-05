# Excel parity

Checks that the engine computes what Excel computes, on an engine build from `wasm-pack` (the
same build the office editor ships). Run it on Modal, never on a laptop:

```sh
r=~/Desktop/Code/office-ui/tools/modal/r
$r up
$r sync ~/Desktop/Code/office-ui/engines/duke-sheets
$r run ~/Desktop/Code/office-ui/engines/duke-sheets \
  'wasm-pack build bindings/wasm --target web --release --no-opt --out-dir pkg -- --locked'
```

**Edge cases** (`cases.mjs`): formulas with the result Excel gives, covering coercion, blanks and
empty text, operators, rounding, dates, text, lookups, arrays, criteria, references, names,
financial functions and number formats. Each runs in a fresh workbook. Cases listed in `known`
are reported but don't fail the run.

```sh
$r run ~/Desktop/Code/office-ui/engines/duke-sheets \
  'node tools/excel-parity/run-cases.mjs bindings/wasm/pkg /tmp/cases.json'
```

**Real workbooks** (`parity.mjs`): recalculates every workbook in a folder and compares each
formula with the value Excel saved in the file, leaving out NOW, TODAY, RAND, CELL and INFO. The
corpus is SpreadsheetBench Verified (800 files made in Excel, CC BY-SA 4.0), downloaded at run
time and never committed. `--max` fails the run when there are more mismatches than the baseline.

```sh
$r run ~/Desktop/Code/office-ui/engines/duke-sheets 'mkdir -p /work/ssb && cd /work/ssb &&
  curl -sSL https://huggingface.co/datasets/KAKA22/SpreadsheetBench/resolve/main/spreadsheetbench_verified_400.tar.gz | tar xz'
$r run ~/Desktop/Code/office-ui/engines/duke-sheets \
  'node tools/excel-parity/parity.mjs bindings/wasm/pkg /work/ssb /tmp/parity.json --max 609'
```

`parity.json` lists each mismatch with its file, cell, formula, Excel's value and the engine's.

Baseline (2026-11-05, COD-1769): 51,932 formulas compared, 609 mismatches in 39 of 800 files,
down from 7,245 in 191. What is left, and why, is listed on COD-1769.
