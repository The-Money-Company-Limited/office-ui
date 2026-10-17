// node dev/autosave-harness.mjs <staged release folder> [port]
// POST /__doc/control {mode: "online"|"offline"|"read_only"|"signed_out"|"conflict"}
// POST /__doc/control {changed: true} simulates another release going live.
import { mkdir, readFile, rename, writeFile } from "node:fs/promises"
import { createServer } from "node:http"
import path from "node:path"
import { pathToFileURL } from "node:url"

const script = `
(() => {
  const send = data => window.postMessage({ cool: "doc", ...data }, location.origin)
  let version
  async function access() {
    const data = await (await fetch('/__doc/access')).json()
    version = data.version
    send({ type: 'access', ...data })
  }
  addEventListener('message', async event => {
    if (event.source !== window || event.origin !== location.origin || event.data?.cool !== 'doc') return
    const data = event.data
    if (data.type === 'hello') await access()
    if (data.type !== 'save') return
    try {
      const files = data.files.map(file => ({ path: file.path, bytes: Array.from(new Uint8Array(file.bytes)) }))
      const result = await (await fetch('/__doc/save', { method: 'POST', body: JSON.stringify({ id: data.id, base: data.base, files }) })).json()
      if (result.type === 'saved') version = result.version
      send(result)
    } catch { send({ type: 'save_failed', id: data.id, reason: 'offline' }) }
  })
  setInterval(async () => {
    try {
      const data = await (await fetch('/__doc/access')).json()
      if (version && data.version !== version) { version = data.version; send({ type: 'changed', version }) }
    } catch {}
  }, 250)
})()
`
const types = {
  ".html": "text/html",
  ".js": "text/javascript",
  ".css": "text/css",
  ".json": "application/json",
  ".wasm": "application/wasm",
}

export async function startHarness(folder, port = 0) {
  folder = path.resolve(folder)
  const artifact = JSON.parse(
    await readFile(path.join(folder, "artifact.json"), "utf8")
  )
  const prefix = artifact.kind === "slides" ? "deck/" : "document/"
  let version = 1
  let mode = "online"
  const saves = []
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
      if (url.pathname === "/__doc/access")
        return json({
          canSave: !["read_only", "signed_out"].includes(mode),
          reason: mode,
          version: String(version),
        })
      if (url.pathname === "/__doc/control") {
        if (request.method === "POST") {
          const update = await body(request)
          if (update.mode) mode = update.mode
          if (update.changed) version++
        }
        return json({ mode, version: String(version), saves })
      }
      if (url.pathname === "/__doc/save" && request.method === "POST") {
        const data = await body(request)
        queue = queue.then(async () => {
          let reason =
            mode === "offline"
              ? "offline"
              : ["read_only", "signed_out"].includes(mode)
                ? "forbidden"
                : undefined
          if (mode === "conflict") {
            version++
            mode = "online"
          }
          if (!reason && data.base !== String(version)) reason = "conflict"
          if (reason)
            return json({
              type: "save_failed",
              id: data.id,
              reason,
              version: String(version),
            })
          if (
            !data.files.length ||
            data.files.some(
              (file) =>
                !file.path.startsWith(prefix) ||
                file.path
                  .split("/")
                  .some((part) => !part || part === "." || part === "..") ||
                file.path.includes("\\")
            )
          )
            throw new Error("Invalid document path")
          const files = data.files.map((file) => ({
            path: file.path,
            bytes: Buffer.from(file.bytes),
          }))
          if (
            files.reduce((size, file) => size + file.bytes.length, 0) >
            64 * 1024 * 1024
          )
            return json({
              type: "save_failed",
              id: data.id,
              reason: "too_large",
            })
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
          saves.push({ base: data.base, paths: files.map((file) => file.path) })
          version++
          json({ type: "saved", id: data.id, version: String(version) })
        })
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
