// The loader runs the site's TSX source with no build: it fetches src/**,
// transpiles what changed (sucrase, in workers), keeps the result in Cache
// Storage and runs it. Cache Storage works even when a private site answers
// no-store, so a warm load downloads nothing but 304s.
//
// Every load revalidates each module with If-None-Match against the ETag its
// cached code was compiled from, so a saved edit shows on the next reload. A
// local import that names no known file is looked for on the site, so a new
// source file needs no manifest entry. Content-named files (vendor/, the
// compilers, the shipped CSS) are cached once and never revalidated.
performance.mark("loader")
const NativeWorker = window.Worker
const root = new URL("./", location.href)
const at = (path) => new URL(path, root)
const codeFile = /\.(tsx?|jsx?)$/
const stats = { compiled: 0, changed: [], css: "shipped" }
window.__loader = stats

async function run() {
  const manifest = await (
    await fetch(at("loader-manifest.json"), { cache: "no-store" })
  ).json()
  // One cache per template build: compiled code from an older compiler or
  // vendor never runs with a newer one.
  const cacheName = `template-${manifest.version}`
  for (const name of await caches.keys())
    if (name.startsWith("template-") && name !== cacheName) caches.delete(name)
  const store = await caches.open(cacheName)
  const stateKey = at("__loader/state")
  const moduleKey = (path) => at("__loader/module/" + path)
  const bundleKey = (entry) => at("__loader/bundle/" + entry)
  const state = (await (await store.match(stateKey))?.json()) ?? {
    etags: {},
    deps: {},
    found: [],
    css: null,
  }
  const cold = Object.keys(state.etags).length === 0

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
  const shippedCss =
    state.css === "shipped" || state.css === null
      ? immutable(manifest.css.shipped)
      : null

  // 1. Revalidate every module against the ETag its cached code came from.
  const known = new Set([...manifest.files, ...state.found])
  const absent = new Set(manifest.absent)
  const checks = await Promise.all(
    [...known]
      .filter((path) => codeFile.test(path))
      .map(async (path) => {
        const etag = state.etags[path]
        const response = await fetch(at(path), {
          cache: "no-store",
          headers: etag ? { "if-none-match": etag } : {},
        })
        if (response.status === 304) return { path }
        if (response.status === 404) return { path, gone: true }
        if (!response.ok) throw new Error(`${path}: HTTP ${response.status}`)
        return {
          path,
          etag: response.headers.get("etag"),
          source: await response.text(),
        }
      })
  )
  performance.mark("revalidated")
  for (const { path } of checks.filter((check) => check.gone)) {
    known.delete(path)
    delete state.etags[path]
    delete state.deps[path]
  }
  state.found = state.found.filter((path) => known.has(path))

  // 2. Compile what changed. An import of an unknown local file is looked for
  // on the site; when it is there, its importer is compiled again.
  const sources = new Map(
    checks
      .filter((check) => check.source !== undefined)
      .map((check) => [check.path, check])
  )
  const compiled = new Map()
  const notFound = new Set()
  let jobs = [...sources.values()]
  while (jobs.length) {
    const results = await compile(manifest, jobs, [...known])
    const retry = new Set()
    const looks = []
    for (const [path, result] of results) {
      compiled.set(path, result)
      for (const { base, candidates } of result.missing) {
        if (absent.has(base) || notFound.has(base)) continue
        looks.push(
          find(candidates).then((file) => {
            if (!file) return notFound.add(base)
            retry.add(path)
            if (known.has(file.path)) return
            known.add(file.path)
            state.found.push(file.path)
            if (codeFile.test(file.path)) {
              sources.set(file.path, file)
              retry.add(file.path)
            }
          })
        )
      }
    }
    await Promise.all(looks)
    jobs = [...retry].map((path) => sources.get(path))
  }
  stats.compiled = compiled.size
  stats.changed = [...sources.keys()]

  // 3. One script per entry (the page, each worker): the runtime, then every
  // module the entry reaches.
  const entries = [manifest.entry, ...Object.keys(manifest.workers)]
  const changed = compiled.size > 0 || checks.some((check) => check.gone)
  const bundles = {}
  if (!changed)
    for (const entry of entries)
      bundles[entry] = await (await store.match(bundleKey(entry)))?.blob()
  const codeOf = async (path) =>
    compiled.get(path)?.code ?? (await store.match(moduleKey(path))).text()
  if (changed || entries.some((entry) => !bundles[entry])) {
    for (const [path, result] of compiled) {
      await store.put(moduleKey(path), new Response(result.code))
      state.etags[path] = sources.get(path).etag
      state.deps[path] = result.files.filter((file) => codeFile.test(file))
    }
    for (const entry of entries) {
      const list = reach(entry, state.deps, known)
      bundles[entry] = new Blob(
        [runtime, ...(await Promise.all(list.map(codeOf)))],
        { type: "text/javascript" }
      )
      await store.put(bundleKey(entry), new Response(bundles[entry]))
    }
    // The shipped CSS was built from the shipped source. A first load can
    // tell from the source it just fetched; any later change rebuilds it.
    if (state.css && state.css !== "shipped") await store.delete(at(state.css))
    state.css =
      cold && (await sourceHash(manifest, sources)) === manifest.css.sourceHash
        ? "shipped"
        : null
    await store.put(stateKey, Response.json(state))
  }
  performance.mark("bundled")

  // 4. CSS: the shipped file, or the one rebuilt after the last change.
  const style = document.createElement("style")
  document.head.append(style)
  if (state.css === "shipped")
    style.textContent = await (
      await (shippedCss ?? immutable(manifest.css.shipped))
    ).text()
  else if (state.css) {
    style.textContent = await (await store.match(at(state.css))).text()
    stats.css = "cached"
  } else {
    // Until Tailwind has run, the shipped CSS styles everything that has not
    // changed.
    style.textContent = await (
      await (shippedCss ?? immutable(manifest.css.shipped))
    ).text()
    stats.css = "rebuilding"
    const modules = [
      ...new Set(entries.flatMap((entry) => reach(entry, state.deps, known))),
    ]
    setTimeout(
      async () =>
        rebuildCss(
          manifest,
          await Promise.all(modules.map(codeOf)),
          immutable,
          style,
          store,
          state,
          stateKey
        ),
      0
    )
  }

  // 5. Run: vendor, then the page's modules. A worker made from a source file
  // boots from its own vendor and bundle.
  self.__loaderBase = root.href
  const workers = {}
  for (const [entry, vendorPath] of Object.entries(manifest.workers))
    workers[entry] = {
      vendor: URL.createObjectURL(await immutable(vendorPath)),
      bundle: URL.createObjectURL(bundles[entry]),
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
      const boot = `self.__loaderBase = ${JSON.stringify(root.href)}; importScripts(${JSON.stringify(workers[entry].vendor)}, ${JSON.stringify(workers[entry].bundle)}); __require(${JSON.stringify(entry)})`
      super(
        URL.createObjectURL(new Blob([boot], { type: "text/javascript" })),
        { name: options?.name }
      )
    }
  }
  for (const blob of [await vendor, bundles[manifest.entry]])
    await script(URL.createObjectURL(blob))
  performance.mark("scripts-ready")
  __require(manifest.entry)
  performance.mark("main-evaluated")

  // Transpiles jobs in up to four workers, largest files first.
  async function compile(manifest, jobs, known) {
    const compilerUrl = URL.createObjectURL(await immutable(manifest.compiler))
    const boot = `importScripts(${JSON.stringify(compilerUrl)});
onmessage = ({ data: { jobs, config, known } }) => {
  const files = new Set(known)
  postMessage(jobs.map(({ path, source }) => {
    const result = self.compiler.compileModule(path, source, config, (p) => files.has(p))
    return { ...result, missing: result.missing.map((base) => ({ base, candidates: self.compiler.candidates(base) })) }
  }))
}`
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
              done(lane.jobs.map((job, i) => [job.path, data[i]]))
            }
            worker.onerror = (event) => {
              worker.terminate()
              fail(new Error(event.message))
            }
            worker.postMessage({
              jobs: lane.jobs.map(({ path, source }) => ({ path, source })),
              config: manifest.config,
              known,
            })
          })
      )
    )
    return results.flat()
  }
}

// The first candidate path the site serves, with its source.
async function find(candidates) {
  for (const path of candidates) {
    const response = await fetch(at(path), { cache: "no-store" })
    if (response.ok)
      return {
        path,
        etag: response.headers.get("etag"),
        source: await response.text(),
      }
  }
  return null
}

// The modules an entry reaches through its imports.
function reach(entry, deps, known) {
  const seen = new Set()
  const stack = [entry]
  while (stack.length) {
    const path = stack.pop()
    if (seen.has(path) || !known.has(path)) continue
    seen.add(path)
    stack.push(...(deps[path] ?? []))
  }
  return [...seen]
}

async function sourceHash(manifest, sources) {
  const text = manifest.files
    .filter((path) => codeFile.test(path))
    .map((path) => path + "\0" + (sources.get(path)?.source ?? ""))
    .join("\0")
  const digest = await crypto.subtle.digest(
    "SHA-1",
    new TextEncoder().encode(text)
  )
  return [...new Uint8Array(digest)]
    .map((byte) => byte.toString(16).padStart(2, "0"))
    .join("")
}

// Runs Tailwind in a worker, so the page starts while it works, then swaps the
// rebuilt CSS in and keeps it for later loads.
async function rebuildCss(
  manifest,
  codes,
  immutable,
  style,
  store,
  state,
  stateKey
) {
  const started = performance.now()
  const compiler = URL.createObjectURL(await immutable(manifest.css.compiler))
  const boot = `import { buildCss } from ${JSON.stringify(compiler)}
onmessage = async ({ data: { css, codes, root } }) => {
  const text = async (path) => {
    const response = await fetch(new URL(path, root), { cache: "no-store" })
    if (!response.ok) throw new Error(path + ": HTTP " + response.status)
    return response.text()
  }
  postMessage(await buildCss(css, codes, text))
}`
  const worker = new NativeWorker(
    URL.createObjectURL(new Blob([boot], { type: "text/javascript" })),
    { type: "module" }
  )
  const css = await new Promise((done, fail) => {
    worker.onmessage = ({ data }) => done(data)
    worker.onerror = (event) => fail(new Error(event.message))
    worker.postMessage({ css: manifest.css, codes, root: root.href })
  })
  worker.terminate()
  const key = "__loader/css/" + Date.now()
  await store.put(at(key), new Response(css))
  state.css = key
  await store.put(stateKey, Response.json(state))
  style.textContent = css
  stats.css = "rebuilt"
  stats.cssMs = Math.round(performance.now() - started)
}

function script(src) {
  return new Promise((done, fail) => {
    const element = document.createElement("script")
    element.src = src
    element.onload = done
    element.onerror = () => fail(new Error(`Could not run ${src}`))
    document.head.append(element)
  })
}

// The module runtime, at the top of every bundle. An id is a served path,
// with "?url" for a file's URL and "?css" for a stylesheet (already applied),
// or "~" and an npm package name from the vendor script.
const runtime = `(() => {
  const defs = Object.create(null), cache = Object.create(null)
  self.__define = (id, factory) => { defs[id] = factory }
  self.__require = function require(id) {
    if (id in cache) return cache[id].exports
    if (id[0] === "~") {
      const namespace = self.__vendor[id.slice(1)]
      if (!namespace) throw new Error("Not in vendor/: " + id.slice(1))
      return (cache[id] = { exports: Object.create(namespace, { __esModule: { value: true } }) }).exports
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
