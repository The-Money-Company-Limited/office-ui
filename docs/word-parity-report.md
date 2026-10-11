# Word parity evaluation — 2026-10-11

The final **1,000-document rerun is complete** on the current source and rebuilt Rust/WASM parser. **791/995 qualified Word page counts match (79.50%)**. This is an improvement in page-count coverage; full Word page-break, geometry, and pixel parity remains incomplete.

| Check | Completed result |
| --- | --- |
| Documents attempted | 1,000 |
| Native Rust parsing | 1000/1,000 |
| Browser captures | 1000/1,000 |
| Qualified native count pairs | 995 |
| Matching counts | 791 |
| More pages than Word | 122 |
| Fewer pages than Word | 82 |
| Summed absolute page-count error | 1012 pages |
| Documents with qualified native page-text membership | 59 |
| Exact normalized page-text memberships | 20 |
| Matching native interior boundaries | 61/220 (197 uniquely anchored) |

On the same 994 qualified pairs as the earlier completed rerun, matches improve from **783 to 791**, with 8 new matches and 0 lost matches. The 5 documents without a qualified count comparison are excluded from the match-rate denominator. Their reasons are preserved in the [open mismatches](../output/playwright/corpus-1000-final-20261011/open-count-mismatches.json).

The automated regression gate passed against all three saved checkpoints. It checks source identity, capture availability, count error, retained matching boundaries and exact page memberships, and raw source-text credit. These checks cover 2559 source-paired comparisons; they do not assert unchanged pixels or universal geometry.

| Checkpoint | Common count pairs | Matching counts before → after | Absolute page error before → after | Count regressions |
| --- | --- | --- | --- | --- |
| Earlier completed 1,000-document run | 994 | 783 → 791 | 1036 → 1011 | 0 |
| Rejected 999-document checkpoint | 994 | 790 → 791 | 1003 → 1002 | 0 |
| Stopped 560-document checkpoint | 558 | 460 → 460 | 299 → 297 | 0 |

The [completed comparison](../output/playwright/corpus-1000-final-20261011/completed-comparison.json), [all 1,000 durable results](../output/playwright/corpus-1000-final-20261011/retained-1000-final.json), and [regression gate](../output/playwright/corpus-1000-final-20261011/regression-gate.json) contain the exact denominators and cases. The earlier rejected 999-document checkpoint is not presented as a completed rerun. The catalog contains 736,706 unique documents; this run covers the retained 1,000-document sample, not the entire catalog.

## Changes and native evidence

The final build retains the earlier body-clip measurement, natural table line-height, character scaling/tracking, wrapped-image frame and fractional-anchor, empty merged-row measurement, and narrowly qualified paragraph auto-spacing repairs. A broad script-weighted paragraph metric change was rejected after it made a document worse. That document is restored to 34 pages versus Word's 35; its remaining font-related discrepancy stays open.

Collapsed shared cell borders now contribute their actual minimum-row and fragment-edge extents. Fresh generic native controls distinguish minimum-height from exact-height rows and cell borders from table borders. This restores a long table document to **35 pages against Word's 35**, after the preceding candidate produced 34. Its final-table page cuts remain different. Broader fractional-row and footer-reserve changes were withheld because they caused another document to regress.

The border correction exposed a separate automatic table-width defect. Word's native print measures wider edge columns to fit unbroken tokens while preserving protected preferred widths. The parser now retains and round-trips cell `noWrap`, including explicit false and edits/clears. The viewer measures minimum token widths across equivalent formatting runs, preserves nonbreaking spaces, permits ordinary spaces to wrap, and refreshes cached widths when fonts change. It applies the new route to qualified simple horizontal tables whose column preferences are protected. Fixed layout, complex cells and unsupported inline geometry keep their previous paths.

The repaired width case now has **8 pages against Word's 8**. All 935 audited source paragraphs are complete. Generic short-text, phrase and token controls retain every one of their 96 markers; the supported protected-width cases agree with native column widths within 0.14 px. Unprotected automatic-table redistribution remains a separate gap. An unconditional saved-grid override was disproved by the short-text native control and was not promoted.

The fresh row audit for this eight-page document matches **2/7** interior row cuts. The older narrow eight-page capture nominally matched 3/7, but also placed rows beyond the body and, on some pages, beyond the physical page. The corrected widths gain the second cut and move two formerly exact nominal cuts; this fresh local audit is separate from the retained automated boundary gate. Equal counts therefore do not establish that every break is improved. See the [native row and width comparison](../output/playwright/followup-regression-20261011/protected-autofit-qualified/native-width-comparison.json) and [native measurements](../output/playwright/minimum-border-regression-20261011/NATIVE-REPORT.md).

REA inspected the installed Word layout framework and preserved original binary identity. The inspected minimum/maximum-width function constructs a formatted subpage and collects width constraints; a related function registers a minimum-width constraint. These observations support content-derived sizing. They do not recover Word's complete table algorithm or prove the viewer's tokenization policy. Native controlled documents qualify the implemented subset. See the [REA and candidate review](../output/playwright/minimum-border-regression-20261011/CANDIDATE-REVIEW.md). The preferred-width semantics are also documented in Microsoft's [noWrap reference](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.nowrap?view=openxml-3.0.1).

## Supplied document, blocks and words

The supplied wrapped-image document remains **one page**, matching Word. All **10/10 blocks**, **three figures**, and **92/92 words** resolve to browser rectangles. All four checked internal body line breaks match the native reference. Title-width error is about +0.250 px; figure horizontal differences remain about −1.3 px, with vertical differences below 1.04 px. This is improved geometry, not pixel identity. The actual browser faces were read back as Arial Black and Calibri.

The Playground retains formatted, Markdown and optional Words views, full text, and all-highlight mode. Rust extraction supplies document-bound source ranges; browser geometry is resolved after layout. The backend does not emit authoritative Word rectangles. Across this run, 16 documents have raw source-range credit gaps (419 UTF-16 units), 689 have unresolved mappings, and 127 have raw DOM rectangle outliers. These flags need individual qualification and are not interchangeable with lost visible text or painted overflow. Every word box in all 1,000 documents has not been compared with Word.

The [supplied-document comparison](../output/playwright/followup-regression-20261011/protected-autofit-qualified/priority-native-comparison.json) and [current rendered page](../output/playwright/followup-regression-20261011/protected-autofit-qualified/wrapped-figure-user-page.png) preserve the measurements.

## Parser coverage and validation

The final source passes **760 TypeScript tests** (one skipped), **129 Rust tests**, **46 Word-oracle harness tests**, workspace and Playground typechecks, and viewer/Playground builds. Two earlier runs under concurrent browser/build load exceeded unchanged timing budgets; the complete final reduced-concurrency run passed. Rebuilt WASM matches Mammoth raw body text in **17/17** retained upstream fixtures. This does not establish Mammoth converter/API compatibility or universal DOC/DOCX content support.

The historical 1,000-document Mammoth audit parsed all 1,000 with docx-core; Mammoth read 997. Of those, 934 matched after whitespace normalization and 967 retained Mammoth's characters in order. Thirty structural reading-order cases retained all exact paragraph occurrences. That historical comparison was not rerun as a new Mammoth-wide audit here. The current fidelity run re-parsed every document with the frozen native binary. See [Mammoth coverage](mammoth-coverage.md), [structured blocks](structured-blocks.md), and the [current validation receipt](../output/playwright/minimum-border-regression-20261011/qualified-code-validation.json).

## Reproducibility and disk use

The run used fresh 2-, 4- and 7-page native Word controls plus qualified source-bound saved native references. It did not export 1,000 fresh native PDFs. Browser imports were fresh, with profile/document font ownership checked. Source hashes, script hashes, parser output hashes and per-document parser wall times are recorded. A 60-second parser ceiling avoids a previously reproduced 25-second wall-time failure on a valid large document. The former failure and byte-identical qualified retry remain preserved.

- Viewer bundle: `a0ae52174aacdf75514df1693a7fe11464458eda8581ca19d15ffaac8cc74edf`.
- WASM: `ef88440993c08aa3bd55f5f1e3c7f910c3d853facd3867fd468daef740841677`.
- Native parser: `4439ad08638a7e0d54af1630a5b16a08a24e59ea8da7921478175c76a567990f`.
- Run contract: `0fc9d1c75f146f010f869f73d77407e0e8208095eaf57757ec0216cafd54881a`.
- Frozen source hashes verified against all 183 recorded files.

The run maintained an 8 GiB free-space reserve and a 512 MiB failure-evidence cap. It deleted 323,459,695 logical bytes of disposable batch products after durable result writes and retained 186,760,727 bytes of failure evidence. An additional 1,223 inactive Rust incremental-cache files occupied 290,676,736 allocated bytes and were removed after builds/tests completed. Earlier cleanup in this task removed about 8.2 GB of Docker build/dependency caches and 1.86 GB of closed task-browser cache blocks. User documents, native references, source and active outputs were preserved. Logical deletion totals and changes in system free space are different measurements.

The remaining work is listed in [open count mismatches](../output/playwright/corpus-1000-final-20261011/open-count-mismatches.json). Table row distribution, fonts and fallback, automatic width redistribution, complex stories and drawings still require native qualification. No claim of complete Word fidelity, complete Mammoth replacement, or all-catalog coverage is made.
