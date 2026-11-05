# The build box for office-ui: Rust with wasm-pack for the engines, Node with pnpm for the apps,
# and Playwright's Chromium for the browser tests. GitHub Actions builds it (wrangler deploy).
FROM docker.io/library/debian:bookworm-slim

RUN apt-get update && apt-get install -y --no-install-recommends \
      curl git build-essential pkg-config ca-certificates xz-utils unzip python3 \
    && curl -fsSL https://deb.nodesource.com/setup_22.x | bash - \
    && apt-get install -y --no-install-recommends nodejs \
    && npm install -g pnpm@9.15.0 \
    && rm -rf /var/lib/apt/lists/*

ENV PATH=/root/.cargo/bin:$PATH CI=1
RUN curl https://sh.rustup.rs -sSf | sh -s -- -y --profile minimal --target wasm32-unknown-unknown \
    && curl -fsSL https://github.com/rustwasm/wasm-pack/releases/download/v0.13.1/wasm-pack-v0.13.1-x86_64-unknown-linux-musl.tar.gz \
       | tar xz -C /usr/local/bin --strip-components=1 wasm-pack-v0.13.1-x86_64-unknown-linux-musl/wasm-pack

# The browser tests load Playwright and its browsers from <repo>/.cod1460; `r up` links it here.
ENV PLAYWRIGHT_BROWSERS_PATH=/opt/cod1460/browsers
RUN mkdir -p /opt/cod1460/tmp && cd /opt/cod1460 \
    && echo '{"name":"cod1460","private":true}' > package.json \
    && npm install --no-audit --no-fund playwright@1.49.1 \
    && npx playwright install --with-deps chromium \
    && printf '<?xml version="1.0"?><fontconfig><dir>/usr/share/fonts</dir><cachedir>/tmp/fc</cachedir></fontconfig>' > fonts.conf \
    && rm -rf /var/lib/apt/lists/*

RUN git config --global --add safe.directory '*'
WORKDIR /work
CMD ["sleep", "infinity"]
