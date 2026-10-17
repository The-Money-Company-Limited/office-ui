# Vendored Duke Sheets WASM

MIT-licensed build of `guseggert/duke-sheets` (upstream base `fec1cd04dbefa5eb65a41d29399ca4a7edc220f2`), WASM 0.1.24, with COD-1459 structural edits in the sibling `duke-sheets` working tree. `LICENSE` is copied from that repository.

Rebuild from `react-xlsx` with `sh scripts/rebuild-duke-wasm.sh`. The script uploads the source, compiles with wasm-pack **on Modal only**, pulls the generated bindings and binary, and copies them here. Review and commit the Rust source and this artifact together.

All source imports use this directory; the build bundles the generated JS and copies the matching `.wasm` into `dist`. The playground and tests use the same binary. There is no registry WASM dependency or alternate runtime path.
