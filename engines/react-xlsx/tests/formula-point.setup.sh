#!/bin/sh
# Run through tools/modal/r; all browser dependencies stay inside this repo.
set -eu
case "$PWD" in /work/*/engines/react-xlsx) ;; *) echo 'Run on Modal in /work/<repo>/engines/react-xlsx'; exit 1 ;; esac
npm install --prefix tests/.browser-tools --no-save --package-lock=false playwright@1.58.2
PLAYWRIGHT_BROWSERS_PATH="$PWD/tests/.browser-cache" tests/.browser-tools/node_modules/.bin/playwright install chromium
mkdir -p tests/.browser-libs
cd tests/.browser-libs
apt-get download libglib2.0-0 libnspr4 libnss3 libatk1.0-0 libatk-bridge2.0-0 libdbus-1-3 libatspi2.0-0 libxcomposite1 libxdamage1 libxfixes3 libxrandr2 libgbm1 libxkbcommon0 libasound2 libxi6 libxrender1 libdrm2 libwayland-server0
for deb in ./*.deb; do dpkg-deb -x "$deb" root; done
