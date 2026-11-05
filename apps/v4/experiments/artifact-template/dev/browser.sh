#!/bin/sh
set -eu
export TMPDIR=/work/office-ui/.cod1460/tmp
export FONTCONFIG_FILE=/work/office-ui/.cod1460/fonts.conf
export PLAYWRIGHT_BROWSERS_PATH=/work/office-ui/.cod1460/browsers
node apps/v4/experiments/artifact-template/dev/browser.mjs "$@"
