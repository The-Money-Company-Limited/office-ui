# Excel parity

Checks that the engine computes what Excel computes, on an engine build from `wasm-pack` (the
same build the office editor ships). Run it on the Cloudflare build box, never on a laptop:

```sh
r=~/Desktop/Code/office-ui/tools/cloudflare/r
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
  'node tools/excel-parity/parity.mjs bindings/wasm/pkg /work/ssb /tmp/parity.json --max 42'
```

`parity.json` lists each mismatch with its file, cell, formula, Excel's value and the engine's;
read it with `$r run … 'cat /tmp/parity.json'`. The corpus and results go when the box stops, so
download the corpus again on a fresh box.

Excel saves a formula's text result without trailing spaces, so an engine result that has extra
trailing spaces counts as the same. One missing a trailing space still counts as a mismatch.

The run fails when there are more mismatches than `--max`, or any workbook fails to open or
calculate (none does now).

Excel takes regional settings (day-first dates, month names, decimal commas) from the computer, and
a file doesn't record them. Cases set them with a trailing `locale` element (`inLocale` in `cases.mjs`); `parity.mjs` lists the four
workbooks that weren't computed in en-US, which their saved results show.

Formulas saved before dynamic arrays (no `cm="1"`, and every formula in a .xls file)
open with Excel 365's `@` where the old rules read one cell of a range (`=SUM(@A1:A3*@B1:B3)`),
and a Ctrl+Shift+Enter array formula (`t="array"` without `cm`) fills its block instead of
spilling.

Baseline (2026-11-06, COD-1790 and COD-1795): 68,264 formulas compared, 42 mismatches in 5 of 800
files, down from 7,245 in 191 before COD-1769; 398 edge cases. 38 are in 49196's init file, whose
dates are built with `TODAY()`: Excel's saved week numbers are from the day it was saved. 51354's
B6 (in both files) reads "Feb 22" as February 22 of the year it is computed in, so Excel's saved
"Jan 24" is from 2024. 48982's B323 includes itself in `MAX(B$7:B377)`: Excel saved 0 for that
circular reference, while the corpus's other circular cells keep the values Excel computed, as the
engine does. 55049's B23 multiplies whole columns, which can't spill from row 23; the engine reads
a whole column only as far as the used cells, so it doesn't see that.
