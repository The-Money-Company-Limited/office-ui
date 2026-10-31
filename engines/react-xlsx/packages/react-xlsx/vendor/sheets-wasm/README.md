# Vendored Duke Sheets WASM

MIT-licensed build of `guseggert/duke-sheets` (upstream base `fec1cd04dbefa5eb65a41d29399ca4a7edc220f2`), WASM 0.1.24, with COD-1459 structural edits, kept in `engines/duke-sheets`. `LICENSE` is copied from that repository.

Rebuild from `react-xlsx` with `sh scripts/rebuild-duke-wasm.sh`. The script uploads the source, compiles with wasm-pack **on Modal only**, pulls the generated bindings and binary, and copies them here. The default command requires a clean, committed Rust tree and uses Cargo.lock. `BUILD.json` records a SHA-256 over sorted source paths and file digests (older manifests also name the duke-sheets commit) and the binary SHA-256. Commit the reviewed Rust fixes first, rebuild, then commit the generated bindings, binary, and manifest together. For development verification before review, append `--allow-dirty`; the manifest explicitly marks that source as dirty and must be regenerated after the Rust commit.

All source imports use this directory; the build bundles the generated JS and copies the matching `.wasm` into `dist`. The playground and tests use the same binary. There is no registry WASM dependency or alternate runtime path.
