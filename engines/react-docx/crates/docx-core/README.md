# docx-core

A native Rust parser shared with the React viewer through WebAssembly. It accepts OOXML DOCX and Word 97–2003 binary DOC through the existing format adapter. No browser, Word installation, or network service is required for parsing.

```rust
let package = docx_core::parse_document_bytes(&bytes)?;
let model = docx_core::build_doc_model(&package);
let blocks = docx_core::extract_document_blocks(&model);
let json = serde_json::to_string(&blocks)?;
```

`parse_document_blocks(&bytes)` combines these steps. `cargo run --example parse_blocks -- input.docx` writes the block JSON to standard output.

Word spans are optional:

```rust
let options = docx_core::ExtractDocumentBlocksOptions { include_words: true };
let blocks = docx_core::parse_document_blocks_with_options(&bytes, options)?;
```

Use `extract_document_blocks_with_options(&model, options)` for an existing model, or add `--include-words` to the CLI. The TypeScript equivalent is `extractDocumentBlocks(model, { includeWords: true })`. Existing calls omit word data.

Paragraph-derived blocks, including those in tables and auxiliary stories, then include `words`. Each word has an ID, its parent `blockId`, original `content`, and document-bound source references with UTF-16 `start`/`end` offsets into the paragraph block's content. Segmentation uses Unicode word boundaries and includes punctuation and symbols; whitespace-only segments are omitted. It is not dictionary-based linguistic segmentation. Styled run boundaries do not split a word, and capitalization formatting does not rewrite its source text. Word rectangles are resolved by the viewer after layout, just like block rectangles.

The versioned output uses Extend's 15 block type names. Each block retains document-bound source references; paragraphs retain styled inline content, tables retain cells and spans, and text boxes retain nested content. Header/footer stories, notes, comments, form controls, and equations are represented where supported by the parser. Image bytes remain available in the parsed package.

`document_source_graph(&package)` provides an optional namespace-aware tree of XML elements, attributes, text, and an asset manifest. Original XML and binary assets remain in `package`. XML byte ranges refer to the decoded UTF-8 strings. For legacy DOC, the graph describes the adapter's OOXML output; keep the original DOC bytes to preserve unsupported binary structures.

Page numbers and bounding boxes require layout. The native parser deliberately returns `geometryStatus: "requires_layout"`. Use the matching React viewer to resolve source references against rendered pages. Source identity is a SHA-256 digest of the sorted, decoded package parts and assets, independent of ZIP compression and entry order. It is not the digest of the original ZIP bytes. Keep the original upload digest separately. Imported source IDs are stable for the same package and parser version; reimport edited bytes as a new document version.

The source graph preserves structure beyond the semantic block model. This is not a promise of complete semantic extraction or Word-identical rendering. Imported HTML, subdocuments, embedded programs/objects, diagrams, and equation layout have explicit diagnostics where recognized. Unknown extension markup can be inspected in the source graph. Dynamic fields use imported values; they are not a general field calculation engine. Encrypted files and older binary formats are not covered by the Word 97–2003 adapter.

See the repository's structured-block design document for the viewer contract, integration steps, validation evidence, and remaining coverage gaps.

## Third-party notices

The symbol-font Unicode mapping is derived from `dingbat-to-unicode` 1.0.1 by Michael Williamson, distributed under the BSD 2-Clause license. The complete notice is in `third-party/dingbat-to-unicode-LICENSE` and is included in the WASM and viewer distributions.
