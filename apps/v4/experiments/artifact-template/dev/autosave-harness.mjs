// node dev/autosave-harness.mjs <staged release folder> [port]
// POST /__doc/control sets mode, retryAfter, pollFailed or receipt (normal/drop/late).
import { createHash } from "node:crypto"
import { mkdir, readFile, rename, writeFile } from "node:fs/promises"
import { createServer } from "node:http"
import { createRequire } from "node:module"
import path from "node:path"
import { pathToFileURL } from "node:url"

const require = createRequire(import.meta.url)
const raw = createRequire(require.resolve("rehype-raw"))
const { parse } = createRequire(raw.resolve("hast-util-raw"))("parse5")
const script = `
(() => {
  const send = data => window.postMessage({ cool: "doc", ...data }, location.origin)
  let version, polling = false, retryAt = 0
  async function access() {
    if (polling) return
    polling = true
    try {
      const response = await fetch('/__doc/access')
      if (!response.ok) return
      const data = await response.json()
      send({ type: 'access', ...data })
      if (version && version !== data.version) send({ type: 'changed', version: data.version })
      version = data.version
    } catch {} finally { polling = false }
  }
  addEventListener('message', async event => {
    if (event.source !== window || event.origin !== location.origin || event.data?.cool !== 'doc') return
    const data = event.data
    if (data.type === 'hello') await access()
    if (data.type !== 'save') return
    const fail = reason => send({ type: 'save_failed', id: data.id, reason })
    if (typeof data.id !== 'string' || !data.id || data.id.length > 128 ||
        typeof data.base !== 'string' || !data.base || data.base.length > 128 ||
        !Array.isArray(data.files) || !data.files.length ||
        data.files.some(file => typeof file?.path !== 'string' || !(file.bytes instanceof ArrayBuffer)) ||
        new Set(data.files.map(file => file.path)).size !== data.files.length) return fail('error')
    if (Date.now() < retryAt) return send({ type: 'save_failed', id: data.id, reason: 'offline', retryAfter: (retryAt - Date.now()) / 1000 })
    const timeout = setTimeout(() => fail('offline'), 19 * 60_000)
    try {
      const files = data.files.map(file => {
        const bytes = new Uint8Array(file.bytes)
        let binary = ''
        for (let offset = 0; offset < bytes.length; offset += 32768)
          binary += String.fromCharCode(...bytes.subarray(offset, offset + 32768))
        return { path: file.path, bytes: btoa(binary) }
      })
      const result = await (await fetch('/__doc/save', { method: 'POST', body: JSON.stringify({ id: data.id, base: data.base, files }) })).json()
      clearTimeout(timeout)
      if (result.type === 'saved') version = result.version
      if (result.reason === 'offline' && Number.isFinite(result.retryAfter)) retryAt = Date.now() + result.retryAfter * 1000
      send(result)
    } catch { clearTimeout(timeout); fail('offline') }
  })
  // Poll frequently in this stand-in so browser checks need no real 20-second wait.
  setInterval(access, 250)
  addEventListener('focus', access)
  document.addEventListener('visibilitychange', () => { if (!document.hidden) access() })
})()
`
const types = {
  ".html": "text/html",
  ".js": "text/javascript",
  ".css": "text/css",
  ".json": "application/json",
  ".wasm": "application/wasm",
}
const slidePath = /^deck\/slides\/[A-Za-z0-9_-]{1,64}\.html$/
const relativePath = (value) =>
  typeof value === "string" &&
  !/[\\\x00-\x1f\x7f]/.test(value) &&
  value
    .split("/")
    .every((part) => part && part !== "." && part !== ".." && part !== ".git")
const imagePath = (value) =>
  relativePath(value) &&
  value.startsWith("deck/") &&
  /\.(png|jpg|jpeg|gif|webp|svg)$/i.test(value)

async function allowedPaths(artifact, files, folder) {
  if (files.some((file) => !relativePath(file.path))) return false
  if (artifact.kind !== "slides")
    return files.every(
      (file) => file.path === artifact.file && file.path.startsWith("document/")
    )
  if (
    files.some(
      (file) =>
        file.path !== "deck/deck.json" &&
        !slidePath.test(file.path) &&
        !imagePath(file.path)
    )
  )
    return false
  const images = files.filter((file) => imagePath(file.path))
  if (!images.length) return true
  const read = async (name) =>
    files.find((file) => file.path === name)?.bytes ??
    (await readFile(path.join(folder, name)))
  const referenced = new Set()
  try {
    const deck = JSON.parse((await read("deck/deck.json")).toString())
    for (const id of deck.order) {
      const name = `deck/slides/${id}.html`
      if (!slidePath.test(name)) return false
      const add = (value) => {
        if (/[\\?#]/.test(value) || /^(?:[a-z][\w+.-]*:|\/\/)/i.test(value))
          return
        try {
          const decoded = decodeURIComponent(value)
          const resolved = decoded.startsWith("/")
            ? decoded.slice(1)
            : path.posix.join(path.posix.dirname(name), decoded)
          if (imagePath(resolved)) referenced.add(resolved)
        } catch {
          /* Invalid URL escapes cannot reference a file. */
        }
      }
      const css = (value) => {
        for (const match of value.matchAll(
          /url\(\s*["']?([^\s"')]+)["']?\s*\)/gi
        ))
          add(match[1])
      }
      const walk = (node) => {
        for (const attr of node.attrs ?? []) {
          if (attr.name === "style") css(attr.value)
          if (
            ["img", "image", "source"].includes(node.tagName) &&
            ["src", "href"].includes(attr.name)
          )
            add(attr.value)
        }
        if (node.content) walk(node.content)
        for (const child of node.childNodes ?? []) {
          if (node.tagName === "style" && child.nodeName === "#text")
            css(child.value)
          walk(child)
        }
      }
      walk(parse((await read(name)).toString()))
    }
    return images.every((file) => referenced.has(file.path))
  } catch {
    return false
  }
}

export async function startHarness(folder, port = 0) {
  folder = path.resolve(folder)
  const artifact = JSON.parse(
    await readFile(path.join(folder, "artifact.json"), "utf8")
  )
  let version = 1
  let mode = "online"
  let retryAfter = 0
  let pollFailed = false
  let receipt = "normal"
  let now = 0
  const saves = [],
    attempts = [],
    receipts = new Map(),
    late = []
  let queue = Promise.resolve()
  const body = async (request) => {
    let text = ""
    for await (const chunk of request) text += chunk
    return JSON.parse(text)
  }
  const server = createServer(async (request, response) => {
    const json = (data) => {
      response.setHeader("content-type", "application/json")
      response.end(JSON.stringify(data))
    }
    response.setHeader("cache-control", "no-store")
    try {
      const url = new URL(request.url, "http://localhost")
      if (url.pathname === "/__doc/bridge.js") {
        response.setHeader("content-type", "text/javascript")
        return response.end(script)
      }
      if (url.pathname === "/__doc/access") {
        if (pollFailed) {
          response.statusCode = 503
          return json({})
        }
        const denied = ["read_only", "signed_out", "not_document"].includes(
          mode
        )
        return json({
          canSave: !denied,
          ...(denied ? { reason: mode } : {}),
          version: String(version),
        })
      }
      if (url.pathname === "/__doc/control") {
        if (request.method === "POST") {
          const update = await body(request)
          if (update.mode) mode = update.mode
          if (update.changed) version++
          if (update.retryAfter !== undefined) retryAfter = update.retryAfter
          if (update.pollFailed !== undefined) pollFailed = update.pollFailed
          if (update.receipt) receipt = update.receipt
          if (update.advance) now += update.advance
          if (update.deliver) for (const deliver of late.splice(0)) deliver()
        }
        return json({
          mode,
          version: String(version),
          saves,
          attempts,
          receipts: receipts.size,
        })
      }
      if (url.pathname === "/__doc/save" && request.method === "POST") {
        const data = await body(request)
        const fail = (reason) =>
          json({
            type: "save_failed",
            id: data.id,
            reason,
            ...(reason === "conflict" ? { version: String(version) } : {}),
            ...(reason === "offline" && retryAfter ? { retryAfter } : {}),
          })
        const save = async () => {
          if (
            typeof data.id !== "string" ||
            !data.id ||
            data.id.length > 128 ||
            typeof data.base !== "string" ||
            !data.base ||
            data.base.length > 128 ||
            !Array.isArray(data.files) ||
            !data.files.length ||
            data.files.some(
              (file) =>
                typeof file?.path !== "string" || typeof file.bytes !== "string"
            ) ||
            new Set(data.files.map((file) => file.path)).size !==
              data.files.length
          )
            return fail("error")
          attempts.push({
            id: data.id,
            base: data.base,
            files: data.files.map((file) => ({
              path: file.path,
              hash: createHash("sha256").update(file.bytes).digest("hex"),
            })),
            time: Date.now(),
          })
          if (["read_only", "signed_out"].includes(mode))
            return fail("forbidden")
          if (
            [
              "offline",
              "not_document",
              "path_not_allowed",
              "too_large",
              "error",
            ].includes(mode)
          )
            return fail(mode)
          const time = Date.now() + now
          for (const [id, value] of receipts)
            if (time - value.time >= 24 * 60 * 60_000) receipts.delete(id)
          if (receipts.has(data.id)) return json(receipts.get(data.id).result)
          if (mode === "conflict") {
            version++
            mode = "online"
          }
          if (data.base !== String(version)) return fail("conflict")
          const files = data.files.map((file) => ({
            path: file.path,
            bytes: Buffer.from(file.bytes, "base64"),
          }))
          if (
            files.some((file) => file.bytes.length > 25 * 1024 * 1024) ||
            files.reduce((size, file) => size + file.bytes.length, 0) >
              64 * 1024 * 1024
          )
            return fail("too_large")
          if (!(await allowedPaths(artifact, files, folder)))
            return fail("path_not_allowed")
          for (const file of files) {
            const target = path.join(folder, file.path)
            await mkdir(path.dirname(target), { recursive: true })
            await writeFile(target + ".pending", file.bytes)
          }
          for (const file of files)
            await rename(
              path.join(folder, file.path + ".pending"),
              path.join(folder, file.path)
            )
          saves.push({
            id: data.id,
            base: data.base,
            paths: files.map((file) => file.path),
          })
          version++
          const result = {
            type: "saved",
            id: data.id,
            version: String(version),
          }
          receipts.set(data.id, { time, result })
          while (receipts.size > 32)
            receipts.delete(receipts.keys().next().value)
          if (receipt === "late") late.push(() => json(result))
          else if (receipt === "drop") response.destroy()
          else json(result)
          receipt = "normal"
        }
        // A rejected write must not poison later requests.
        queue = queue.then(save).catch(() => fail("error"))
        await queue
        return
      }
      const relative = decodeURIComponent(url.pathname).slice(1) || "index.html"
      const file = path.resolve(folder, relative)
      if (!file.startsWith(folder + path.sep)) throw new Error("Invalid path")
      let bytes = await readFile(file)
      if (relative === "index.html")
        bytes = bytes
          .toString()
          .replace("<head>", '<head><script src="/__doc/bridge.js"></script>')
      response.setHeader(
        "content-type",
        types[path.extname(file)] ?? "application/octet-stream"
      )
      response.end(bytes)
    } catch (error) {
      response.statusCode = 500
      response.end(String(error))
    }
  })
  await new Promise((resolve) => server.listen(port, "127.0.0.1", resolve))
  return server
}

if (
  process.argv[1] &&
  import.meta.url === pathToFileURL(process.argv[1]).href
) {
  if (!process.argv[2])
    throw new Error("Usage: node autosave-harness.mjs <release folder> [port]")
  const server = await startHarness(
    process.argv[2],
    Number(process.argv[3] ?? 4312)
  )
  console.log(`Autosave bridge: http://127.0.0.1:${server.address().port}`)
}
