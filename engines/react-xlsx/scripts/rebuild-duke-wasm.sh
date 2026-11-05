#!/bin/sh
# Run from the laptop; all compilation happens on the shared Cloudflare build box.
set -eu
exec node "$(dirname "$0")/rebuild-duke-wasm.mjs" "$@"
