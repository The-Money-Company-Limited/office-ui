// The loader runs the site's TSX source with no build. Every load it
// revalidates each source file, looks up or compiles it, links the imports
// afresh against the files the site has now, and runs the result.
//
// The only cache is immutable: a file's compiled code and import specifiers,
// keyed by its path and ETag (compiling is a pure function of the source),
// the built CSS keyed by a hash of its inputs, and the content-named files in
// vendor/. Tabs cannot corrupt each other, and a reverted file finds its old
// entry. Each path's last ETag is kept only as an If-None-Match hint: a stale
// hint costs a download, never wrong code. Cache Storage works even when a
// private site answers no-store, so a warm load transfers only 304s.
import { candidates, codeFile, link } from "./resolve.js"

performance.mark("loader")
const NativeWorker = window.Worker
const root = new URL("./", location.href)
const at = (path) => new URL(path, root)
const sourceFile = (path) => codeFile.test(path) || path.endsWith(".css")
const stats = { compiled: 0, css: "shipped" }
window.__loader = stats

async function run() {
  const manifest = await (
    await fetch(at("loader-manifest.json"), { cache: "no-store" })
  ).json()
  // One cache per template build: code from an older compiler or vendor
  // never runs with a newer one.
  const cacheName = `template-${manifest.version}`
  for (const name of await caches.keys())
    if (name.startsWith("template-") && name !== cacheName) caches.delete(name)
  const store = await caches.open(cacheName)
  async function immutable(path) {
    let response = await store.match(at(path))
    if (!response) {
      const fresh = await fetch(at(path), { cache: "no-store" })
      if (!fresh.ok) throw new Error(`${path}: HTTP ${fresh.status}`)
      response = new Response(await fresh.blob())
      await store.put(at(path), response.clone())
    }
    return response.blob()
  }
  const vendor = immutable(manifest.vendor)
  const shippedCss = immutable(manifest.css.shipped)

  // 1. Revalidate each source file and look up its entry; what has none is
  // compiled below.
  const entryKey = (path, etag) =>
    at(`__loader/entry/${path}?etag=${encodeURIComponent(etag)}`)
  const hintKey = (path) => at("__loader/etag/" + path)
  const files = new Set(manifest.files)
  const entries = new Map() // path -> {code, specs, hash}
  const jobs = []
  // Cache writes never hold up the page; a failed one only costs a recompile.
  const write = (promise) => promise.catch(() => {})
  async function load(path) {
    const hint = await (await store.match(hintKey(path)))?.text()
    let response = await fetch(at(path), {
      cache: "no-store",
      headers: hint ? { "if-none-match": hint } : {},
    })
    if (response.status === 304) {
      const cached = await store.match(entryKey(path, hint))
      if (cached) {
        entries.set(path, await cached.json())
        return true
      }
      response = await fetch(at(path), { cache: "no-store" })
    }
    if (response.status === 404) {
      files.delete(path)
      return false
    }
    if (!response.ok) throw new Error(`${path}: HTTP ${response.status}`)
    const source = await response.text()
    const hash = await sha1(source)
    const etag = response.headers.get("etag") || hash
    const cached = await store.match(entryKey(path, etag))
    if (!cached) jobs.push({ path, etag, source, hash })
    else {
      entries.set(path, await cached.json())
      if (etag !== hint)
        write(store.put(hintKey(path), new Response(etag)))
    }
    return true
  }
  await Promise.all(manifest.files.filter(sourceFile).map(load))
  performance.mark("revalidated")

  // 2. Compile what has no entry and link. A local import that names no file
  // is looked for on the site, on every load; nothing about a failed lookup
  // is kept.
  const absent = new Set(manifest.absent)
  const notFound = new Set()
  const workerEntries = Object.keys(manifest.workers)
  let graphs
  for (;;) {
    if (jobs.length) await compile(jobs.splice(0))
    graphs = [manifest.entry, ...workerEntries].map((entry) =>
      link(
        entry,
        (path) => entries.get(path)?.specs,
        manifest.config,
        (path) => files.has(path)
      )
    )
    const missing = [
      ...new Set(
        graphs
          .flatMap((graph) => graph.missing)
          .filter((base) => !absent.has(base) && !notFound.has(base))
      ),
    ]
    if (!missing.length && !jobs.length) break
    await Promise.all(
      missing.map(async (base) => {
        for (const path of candidates(base).filter(sourceFile))
          if (await load(path)) return files.add(path)
        notFound.add(base)
      })
    )
  }
  const [page, ...workerGraphs] = graphs
  const script = (graph) =>
    new Blob(
      [
        runtime,
        ...graph.modules.map(
          (path) =>
            `__define(${JSON.stringify(path)}, ${JSON.stringify(graph.maps[path])}, ${entries.get(path).code});\n`
        ),
      ],
      { type: "text/javascript" }
    )
  performance.mark("linked")

  // 3. CSS, keyed by the content of every module (its classes) and
  // stylesheet: the shipped file, a rebuilt one from the cache, or the shipped
  // one while Tailwind rebuilds it in a worker.
  const modules = [...new Set(graphs.flatMap((graph) => graph.modules))]
  const cssKey = await sha1(
    JSON.stringify(
      [...modules, ...page.css]
        .sort()
        .map((path) => [path, entries.get(path).hash])
    )
  )
  const style = document.createElement("style")
  document.head.append(style)
  const rebuilt =
    cssKey === manifest.css.key
      ? null
      : await store.match(at("__loader/css/" + cssKey))
  if (rebuilt) {
    style.textContent = await rebuilt.text()
    stats.css = "cached"
  } else {
    style.textContent = await (await shippedCss).text()
    if (cssKey !== manifest.css.key) {
      stats.css = "rebuilding"
      const codes = modules.map((path) => entries.get(path).code)
      const texts = Object.fromEntries(
        page.css.map((path) => [path, entries.get(path).code])
      )
      setTimeout(async () => {
        const css = await rebuildCss(
          manifest.css,
          page.sheets,
          texts,
          codes,
          immutable
        )
        await store.put(at("__loader/css/" + cssKey), new Response(css))
        style.textContent = css
        stats.css = "rebuilt"
      }, 0)
    }
  }

  // 4. Run: vendor, then the page's modules. A worker made from a source file
  // boots from its own vendor and script.
  self.__loaderBase = root.href
  const workers = {}
  for (const [i, entry] of workerEntries.entries())
    workers[entry] = {
      vendor: URL.createObjectURL(await immutable(manifest.workers[entry])),
      script: URL.createObjectURL(script(workerGraphs[i])),
    }
  window.Worker = class extends NativeWorker {
    constructor(url, options) {
      const target = new URL(url, location.href)
      const path = target.pathname.slice(root.pathname.length)
      const entry = [
        path,
        path.replace(/\.js$/, ".ts"),
        path.replace(/\.js$/, ".tsx"),
      ].find((p) => workers[p])
      if (target.origin !== root.origin || !entry) {
        super(url, options)
        return
      }
      const boot = `self.__loaderBase = ${JSON.stringify(root.href)}; importScripts(${JSON.stringify(workers[entry].vendor)}, ${JSON.stringify(workers[entry].script)}); __require(${JSON.stringify(entry)})`
      super(
        URL.createObjectURL(new Blob([boot], { type: "text/javascript" })),
        { name: options?.name }
      )
    }
  }
  for (const blob of [await vendor, script(page)])
    await runScript(URL.createObjectURL(blob))
  performance.mark("scripts-ready")
  __require(manifest.entry)
  performance.mark("main-evaluated")

  // Compiles jobs in up to four workers, largest files first, and keeps each
  // result under its path and ETag.
  async function compile(jobs) {
    const compilerUrl = URL.createObjectURL(await immutable(manifest.compiler))
    const boot = `importScripts(${JSON.stringify(compilerUrl)});
onmessage = ({ data }) => postMessage(data.map(({ path, source }) => self.compiler.compileFile(path, source)))`
    const bootUrl = URL.createObjectURL(
      new Blob([boot], { type: "text/javascript" })
    )
    const count = Math.min(
      jobs.length,
      4,
      Math.max(1, (navigator.hardwareConcurrency || 2) - 1)
    )
    const lanes = Array.from({ length: count }, () => ({ jobs: [], size: 0 }))
    for (const job of [...jobs].sort(
      (a, b) => b.source.length - a.source.length
    )) {
      const lane = lanes.reduce((a, b) => (a.size <= b.size ? a : b))
      lane.jobs.push(job)
      lane.size += job.source.length
    }
    const results = await Promise.all(
      lanes.map(
        (lane) =>
          new Promise((done, fail) => {
            const worker = new NativeWorker(bootUrl)
            worker.onmessage = ({ data }) => {
              worker.terminate()
              done(lane.jobs.map((job, i) => [job, data[i]]))
            }
            worker.onerror = (event) => {
              worker.terminate()
              fail(new Error(event.message))
            }
            worker.postMessage(
              lane.jobs.map(({ path, source }) => ({ path, source }))
            )
          })
      )
    )
    for (const [job, result] of results.flat()) {
      const entry = { ...result, hash: job.hash }
      entries.set(job.path, entry)
      write(
        store
          .put(entryKey(job.path, job.etag), Response.json(entry))
          .then(() => store.put(hintKey(job.path), new Response(job.etag)))
      )
    }
    stats.compiled += jobs.length
  }

  function runScript(src) {
    return new Promise((done, fail) => {
      const element = document.createElement("script")
      element.src = src
      element.onload = done
      element.onerror = () => fail(new Error(`Could not run ${src}`))
      document.head.append(element)
    })
  }
}

async function sha1(text) {
  const digest = await crypto.subtle.digest(
    "SHA-1",
    new TextEncoder().encode(text)
  )
  return [...new Uint8Array(digest)]
    .map((byte) => byte.toString(16).padStart(2, "0"))
    .join("")
}

// Runs Tailwind in a worker, so the page starts while it works.
// It reads the stylesheets the loader fetched (texts), and fetches only the
// CSS packages' files in vendor/.
async function rebuildCss(css, sheets, texts, codes, immutable) {
  const compiler = URL.createObjectURL(await immutable(css.compiler))
  const boot = `import { buildCss } from ${JSON.stringify(compiler)}
onmessage = async ({ data: { sheets, texts, packages, codes, root } }) => {
  const text = async (path) => {
    if (path in texts) return texts[path]
    const response = await fetch(new URL(path, root), { cache: "no-store" })
    if (!response.ok) throw new Error(path + ": HTTP " + response.status)
    return response.text()
  }
  postMessage(await buildCss(sheets, packages, codes, text))
}`
  const worker = new NativeWorker(
    URL.createObjectURL(new Blob([boot], { type: "text/javascript" })),
    {
      type: "module",
    }
  )
  try {
    return await new Promise((done, fail) => {
      worker.onmessage = ({ data }) => done(data)
      worker.onerror = (event) => fail(new Error(event.message))
      worker.postMessage({
        sheets,
        texts,
        packages: css.packages,
        codes,
        root: root.href,
      })
    })
  } finally {
    worker.terminate()
  }
}

// The module runtime, at the top of every script. Each module carries its
// specifier-to-id map; an id is a served path, "path?url" for a file's URL,
// "path?css" for a stylesheet (applied by the loader), or "~" and an npm
// package from the vendor script.
const runtime = `(() => {
  const defs = Object.create(null), cache = Object.create(null)
  self.__define = (path, map, factory) => { defs[path] = { map, factory } }
  const load = (self.__require = (id) => {
    if (id in cache) return cache[id].exports
    if (id[0] === "~") {
      const namespace = self.__vendor[id.slice(1)]
      if (!namespace) throw new Error("Not in vendor/: " + id.slice(1))
      return (cache[id] = { exports: Object.create(namespace, { __esModule: { value: true } }) }).exports
    }
    const q = id.indexOf("?"), path = q < 0 ? id : id.slice(0, q), query = q < 0 ? "" : id.slice(q + 1)
    if (query === "url") return { __esModule: true, default: new URL(path, self.__loaderBase).href }
    if (query === "css") return {}
    const def = defs[path]
    if (!def) throw new Error("Module not found: " + id)
    const module = (cache[id] = { exports: {} })
    const require = (spec) => {
      if (!(spec in def.map)) throw new Error("Module not found: " + spec + " (imported by " + path + ")")
      return load(def.map[spec])
    }
    def.factory(require, module, module.exports, new URL(path, self.__loaderBase).href)
    return module.exports
  })
})();
`

try {
  await run()
} catch (error) {
  const message = document.createElement("pre")
  message.setAttribute("role", "alert")
  message.style.cssText =
    "margin:24px;white-space:pre-wrap;font:13px ui-monospace,monospace"
  message.textContent = `The page's source could not run.\n\n${error?.stack ?? error}`
  document.getElementById("root").replaceChildren(message)
  throw error
}
