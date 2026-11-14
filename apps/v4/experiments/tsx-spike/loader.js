// The loader: turns the site's src/**/*.tsx into one cached script and runs it.
//
// Every load revalidates each source file with If-None-Match against the
// ETags the cached script was built from (304s when nothing changed), so an
// edit on the site shows on the next reload. Changed files are transpiled
// with sucrase in a few workers. The result lives in Cache Storage, so a
// warm load works even when the site answers no-store (a private computer).
// vendor/vendor-*.js (npm dependencies) and the compiler are immutable,
// content-named files kept the same way without revalidation.
performance.mark("loader")
const NativeWorker = window.Worker
const root = new URL("./", location.href)
const store = await caches.open("tsx-registry-1")
const manifest = await (await fetch("./loader-manifest.json", { cache: "no-store" })).json()
const fileSet = new Set(manifest.files)
const modules = manifest.files.filter((f) => /\.(tsx?|jsx|js)$/.test(f))
const stats = { compiled: 0, compileMs: 0, changed: [], bundle: "hit", css: "shipped" }
window.__loaderStats = stats

async function immutable(path) {
  const key = new Request(new URL("__immutable/" + path, root))
  let response = await store.match(key)
  if (!response) {
    const fresh = await fetch(new URL(path, root), { cache: "no-store" })
    if (!fresh.ok) throw new Error(`${path}: HTTP ${fresh.status}`)
    response = new Response(await fresh.blob(), { headers: { "content-type": "text/javascript" } })
    await store.put(key, response.clone())
  }
  return URL.createObjectURL(await response.blob())
}
const vendorUrl = immutable(manifest.vendorFile)
const workerVendorUrl = immutable(manifest.workerVendorFile)

// 1. Revalidate every source file against the cached script's ETags.
const bundleKey = new Request(new URL("__bundle", root))
const bundle = await store.match(bundleKey)
const known = bundle ? JSON.parse(bundle.headers.get("x-etags")) : {}
const checks = await Promise.all(
  modules.map(async (path) => {
    const etag = known[path]
    const response = await fetch(new URL(path, root), { cache: "no-store", headers: etag ? { "if-none-match": etag } : {} })
    if (response.status === 304) return { path, etag }
    if (!response.ok) throw new Error(`${path}: HTTP ${response.status}`)
    return { path, etag: response.headers.get("etag"), source: await response.text() }
  })
)
performance.mark("revalidated")
const changed = checks.filter((c) => c.source !== undefined)
stats.changed = changed.map((c) => c.path)

// 2. Transpile what changed, in up to four workers, largest files first.
async function compile(jobs) {
  const started = performance.now()
  const compilerUrl = await immutable(manifest.compilerFile)
  const config = { aliases: manifest.aliases, prefixes: manifest.prefixes, vendor: manifest.vendor, files: manifest.files }
  const count = Math.min(jobs.length, 4, Math.max(1, (navigator.hardwareConcurrency || 2) - 1))
  const lanes = Array.from({ length: count }, () => ({ jobs: [], size: 0 }))
  for (const job of [...jobs].sort((a, b) => b.source.length - a.source.length)) {
    const lane = lanes.reduce((a, b) => (a.size <= b.size ? a : b))
    lane.jobs.push(job)
    lane.size += job.source.length
  }
  const boot = `importScripts(${JSON.stringify(compilerUrl)});
onmessage = ({ data: { jobs, config } }) => {
  const files = new Set(config.files)
  postMessage(jobs.map((job) => self.tsxCompiler.toRegistry(job.path, job.source, config, (p) => files.has(p))))
}`
  const bootUrl = URL.createObjectURL(new Blob([boot], { type: "text/javascript" }))
  const results = await Promise.all(
    lanes.map(
      (lane) =>
        new Promise((done, fail) => {
          const worker = new NativeWorker(bootUrl)
          worker.onmessage = ({ data }) => {
            worker.terminate()
            done(lane.jobs.map((job, i) => [job.path, data[i]]))
          }
          worker.onerror = (event) => fail(new Error(event.message))
          worker.postMessage({ jobs: lane.jobs.map(({ path, source }) => ({ path, source })), config })
        })
    )
  )
  stats.compileMs = Math.round(performance.now() - started)
  stats.compiled = jobs.length
  return new Map(results.flat())
}

const runtime = `(() => {
  const defs = Object.create(null), cache = Object.create(null)
  self.__define = (id, factory) => { defs[id] = factory }
  self.__require = function require(id) {
    if (id in cache) return cache[id].exports
    if (id[0] === "~") {
      const ns = self.__vendor[id.slice(1)]
      if (!ns) throw new Error("Vendor module not found: " + id.slice(1))
      return (cache[id] = { exports: Object.create(ns, { __esModule: { value: true } }) }).exports
    }
    const q = id.indexOf("?"), path = q < 0 ? id : id.slice(0, q), query = q < 0 ? "" : id.slice(q + 1)
    if (query === "url") return { __esModule: true, default: new URL(path, self.__loaderBase).href }
    if (query === "css") return {}
    const factory = defs[path]
    if (!factory) throw new Error("Module not found: " + id)
    const module = (cache[id] = { exports: {} })
    factory(require, module, module.exports, new URL(path, self.__loaderBase).href)
    return module.exports
  }
})();
`

let code
let workerCode = null
let cssKey = bundle?.headers.get("x-css") ?? null
if (!changed.length && bundle) {
  code = await bundle.blob()
} else {
  stats.bundle = "rebuilt"
  const compiled = await compile(changed)
  const parts = await Promise.all(
    modules.map(async (path) => {
      const key = new Request(new URL("__module/" + path, root))
      if (compiled.has(path)) {
        await store.put(key, new Response(compiled.get(path)))
        return compiled.get(path)
      }
      return (await store.match(key)).text()
    })
  )
  code = new Blob([runtime, ...parts], { type: "text/javascript" })
  workerCode = Object.fromEntries(
    Object.entries(manifest.workerModules).map(([entry, list]) => [
      entry,
      new Blob([runtime, ...list.map((p) => parts[modules.indexOf(p)])], { type: "text/javascript" }),
    ])
  )
  for (const [entry, blob] of Object.entries(workerCode)) await store.put(new Request(new URL("__bundle-worker/" + entry, root)), new Response(blob))
  const etags = Object.fromEntries(checks.map((c) => [c.path, c.etag]))
  // Tailwind: the shipped CSS matches the shipped source; a first load can
  // check that from the source it just fetched. After an edit, rebuild.
  if (!bundle && changed.length === modules.length) {
    const sources = modules.filter((p) => /\.(tsx?|jsx)$/.test(p)).map((p) => [p, checks.find((c) => c.path === p).source])
    cssKey = (await sha1(sources.map(([p, s]) => p + "\0" + s).join("\0"))) === manifest.tailwind.sourceHash ? "shipped" : null
  } else cssKey = null
  const save = (css) =>
    store.put(bundleKey, new Response(code, { headers: { "x-etags": JSON.stringify(etags), ...(css ? { "x-css": css } : {}) } }))
  await save(cssKey)
  if (!cssKey) setTimeout(() => rebuildCss(parts, etags).then(save), 0)
}
performance.mark("bundled")

async function sha1(text) {
  const digest = await crypto.subtle.digest("SHA-1", new TextEncoder().encode(text))
  return [...new Uint8Array(digest)].map((b) => b.toString(16).padStart(2, "0")).join("")
}

function applyCss(css) {
  document.querySelector("link[data-app-css]").href = URL.createObjectURL(new Blob([css], { type: "text/css" }))
}

async function rebuildCss(parts, etags) {
  const started = performance.now()
  const key = "__css/" + (await sha1(JSON.stringify(etags)))
  const text = async (p) => (await fetch(new URL(p, root), { cache: "no-store" })).text()
  const { buildCss } = await import(new URL(manifest.tailwind.compiler, root).href)
  const { entry, packages, plain } = manifest.tailwind
  const css = (await buildCss(entry, parts, packages, text)) + "\n" + (await Promise.all(plain.map(text))).join("\n")
  await store.put(new Request(new URL(key, root)), new Response(css))
  applyCss(css)
  stats.css = "rebuilt"
  stats.cssMs = Math.round(performance.now() - started)
  return key
}

if (cssKey && cssKey !== "shipped") {
  applyCss(await (await store.match(new Request(new URL(cssKey, root)))).text())
  stats.css = "cached"
}

// 3. Run: vendor first, then the modules; workers made from a source file
// boot from the same two scripts.
self.__loaderBase = root.href
const bundleUrl = URL.createObjectURL(code)
const vendor = await vendorUrl
const workerVendor = await workerVendorUrl
const workerUrls = {}
for (const entry of Object.keys(manifest.workerModules)) {
  const blob = workerCode?.[entry] ?? (await (await store.match(new Request(new URL("__bundle-worker/" + entry, root)))).blob())
  workerUrls[entry] = URL.createObjectURL(blob)
}
window.Worker = class extends NativeWorker {
  constructor(url, options) {
    // Each worker made from a source file runs its own graph's script.
    const target = new URL(url, location.href)
    let path = target.pathname.slice(root.pathname.length)
    if (path.endsWith(".js") && !fileSet.has(path))
      path = [".ts", ".tsx"].map((e) => path.slice(0, -3) + e).find((p) => fileSet.has(p)) ?? path
    if (target.origin !== root.origin || !workerUrls[path]) {
      super(url, options)
      return
    }
    const boot = `self.__loaderBase = ${JSON.stringify(root.href)}; importScripts(${JSON.stringify(workerVendor)}, ${JSON.stringify(workerUrls[path])}); __require(${JSON.stringify(path)})`
    super(URL.createObjectURL(new Blob([boot], { type: "text/javascript" })), { name: options?.name })
  }
}
for (const src of [vendor, bundleUrl])
  await new Promise((done, fail) => {
    const script = document.createElement("script")
    script.src = src
    script.onload = done
    script.onerror = fail
    document.head.append(script)
  })
performance.mark("scripts-ready")
__require(manifest.entry)
performance.mark("main-evaluated")
