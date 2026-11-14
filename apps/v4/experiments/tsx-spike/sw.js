// Service worker loader: serves src/**/*.ts(x) as browser ESM, transpiled on
// first use and kept in Cache Storage with the source file's ETag. Every file
// under src/ and vendor/ is revalidated with If-None-Match on each page load,
// so an edited file shows on the next reload and an unchanged one costs a 304.
// The site may answer no-store (a private computer), so this cache, not the
// HTTP cache, is what makes a warm load warm.
import { compile } from "./compile.js"

const VERSION = "tsx-spike-2"
const scope = new URL(self.registration.scope)
const base = scope.pathname
let generation = null // per page load: manifest + in-flight responses

self.addEventListener("install", () => self.skipWaiting())
self.addEventListener("activate", (event) => event.waitUntil(self.clients.claim()))

function newGeneration() {
  const manifest = fetch(new URL("loader-manifest.json", scope), { cache: "no-store" })
    .then((response) => response.json())
    .then((value) => ({ ...value, fileSet: new Set(value.files) }))
  return { manifest, responses: new Map(), stats: { compiled: 0, cached: 0, compileMs: 0, doneAt: 0 }, started: performance.now() }
}

const cache = caches.open(VERSION)

// One cached entry per served path: the response we serve (compiled JS for
// source) and, in x-etag, the ETag of the file it came from.
async function serve(g, path, kind) {
  const store = await cache
  const key = new Request(new URL(`__${kind}/${path}`, scope))
  const cached = await store.match(key)
  const etag = cached?.headers.get("x-etag")
  const response = await fetch(new URL(path, scope), { cache: "no-store", headers: etag ? { "if-none-match": etag } : {} })
  if (response.status === 304 && cached) {
    if (kind === "js") g.stats.cached++
    g.stats.doneAt = Math.round(performance.now() - g.started)
    return cached
  }
  if (!response.ok) return response
  const headers = { "x-etag": response.headers.get("etag") || "" }
  let body
  if (kind === "js") {
    const manifest = await g.manifest
    const source = await response.text()
    const started = performance.now()
    body = compile(path, source, manifest, (p) => manifest.fileSet.has(p))
    g.stats.compileMs += performance.now() - started
    g.stats.compiled++
    headers["content-type"] = "text/javascript; charset=utf-8"
  } else {
    body = await response.arrayBuffer()
    headers["content-type"] = response.headers.get("content-type") || "application/octet-stream"
  }
  const fresh = new Response(body, { headers })
  if (headers["x-etag"]) await store.put(key, fresh.clone())
  g.stats.doneAt = Math.round(performance.now() - g.started)
  return fresh
}

function served(g, path, kind) {
  if (!g.responses.has(kind + path)) g.responses.set(kind + path, serve(g, path, kind))
  return g.responses.get(kind + path).then((response) => response.clone())
}

const js = (body) => new Response(body, { headers: { "content-type": "text/javascript; charset=utf-8" } })

async function handle(request) {
  const url = new URL(request.url)
  let path = decodeURIComponent(url.pathname.slice(base.length))
  if (!generation) generation = newGeneration()
  const g = generation
  const manifest = await g.manifest
  if (url.search === "?raw") return served(g, path, "raw")
  if (url.search === "?url") return js(`export default ${JSON.stringify(new URL(path, scope).href)}`)
  if (url.search === "?css")
    return js(
      `if (!document.querySelector('link[data-app-css]')) { const l = document.createElement("link"); l.rel = "stylesheet"; l.dataset.appCss = ""; l.href = ${JSON.stringify(new URL(manifest.css, scope).href)}; document.head.append(l) }`
    )
  if (path.startsWith("src/")) {
    // TypeScript convention: "./x.js" names x.ts (the xlsx worker URL does).
    if (!manifest.fileSet.has(path) && path.endsWith(".js")) {
      const ts = [".ts", ".tsx"].map((e) => path.slice(0, -3) + e).find((p) => manifest.fileSet.has(p))
      if (ts) path = ts
    }
    if (/\.(tsx?|jsx)$/.test(path)) return served(g, path, "js")
  }
  return served(g, path, "raw")
}

self.addEventListener("message", (event) => {
  if (event.data === "stats") event.source.postMessage({ stats: generation?.stats })
})

self.addEventListener("fetch", (event) => {
  const url = new URL(event.request.url)
  if (url.origin !== scope.origin || !url.pathname.startsWith(base) || event.request.method !== "GET") return
  if (event.request.mode === "navigate") {
    generation = newGeneration()
    return
  }
  const path = url.pathname.slice(base.length)
  if (!path.startsWith("src/") && !path.startsWith("vendor/")) return
  event.respondWith(handle(event.request))
})
