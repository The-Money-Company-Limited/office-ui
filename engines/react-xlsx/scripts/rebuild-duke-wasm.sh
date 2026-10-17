#!/bin/sh
# Run from the laptop; all compilation happens in the shared Modal sandbox.
set -eu
exec node "$(dirname "$0")/rebuild-duke-wasm.mjs" "$@"
