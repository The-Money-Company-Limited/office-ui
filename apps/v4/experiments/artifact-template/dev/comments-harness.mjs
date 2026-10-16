// Serves a staged release with RuntimeVM's comment script injected, the way
// the cool.computer gateway does, and a stand-in for the comment layer that
// saves every pick at once. It checks that Comments and Ink anchor on the
// office pages without publishing anything.
//
//   node dev/comments-harness.mjs <runtimevm checkout> <release dir> [port]
//
// The site is http://127.0.0.1:<port>, the layer http://127.0.0.1:<port+1>.
// GET <layer>/notes lists every saved anchor as the owner's agent reads it;
// DELETE <layer>/notes clears them.
import { readFile, stat } from "node:fs/promises"
import { createServer } from "node:http"
import { createRequire } from "node:module"
import path from "node:path"

// esbuild comes with Vite, which this app already depends on.
const fromVite = createRequire(createRequire(import.meta.url).resolve("vite"))
const { build } = fromVite("esbuild")

const [runtimevm, release, portArgument = "4310"] = process.argv.slice(2)
if (!runtimevm || !release) {
  throw new Error(
    "Usage: node dev/comments-harness.mjs <runtimevm checkout> <release dir> [port]"
  )
}
const port = Number(portArgument)
const layerOrigin = `http://127.0.0.1:${port + 1}`

// The gateway wraps the same body (workers/computer-gateway/src/picker.js).
const {
  outputFiles: [bundle],
} = await build({
  entryPoints: [
    path.join(runtimevm, "workers/computer-gateway/src/comments/main.js"),
  ],
  bundle: true,
  format: "iife",
  legalComments: "none",
  write: false,
})
const script = `;(function (window, layerOrigin) {${bundle.text}\n})(window, ${JSON.stringify(layerOrigin)})\n`

const types = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json",
  ".wasm": "application/wasm",
  ".docx":
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  ".pptx":
    "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  ".pdf": "application/pdf",
}

createServer(async (request, response) => {
  const url = new URL(request.url, `http://127.0.0.1:${port}`)
  if (url.pathname === "/__cool/comments.js") {
    response.writeHead(200, { "content-type": types[".js"] })
    return response.end(script)
  }
  let file = path.join(release, decodeURIComponent(url.pathname))
  if (!file.startsWith(path.resolve(release))) {
    response.writeHead(403)
    return response.end()
  }
  try {
    if ((await stat(file)).isDirectory()) file = path.join(file, "index.html")
    let body = await readFile(file)
    const type = types[path.extname(file)] ?? "application/octet-stream"
    if (type.startsWith("text/html")) {
      body = body
        .toString("utf8")
        .replace(/<head>/i, '<head><script src="/__cool/comments.js"></script>')
    }
    response.writeHead(200, {
      "content-type": type,
      "cache-control": "no-cache",
    })
    response.end(body)
  } catch {
    response.writeHead(404)
    response.end("not found")
  }
}).listen(port, "127.0.0.1")

// The layer: says ready, takes the port, says the person may comment, and
// saves each pick as an open note. The real layer asks for words first.
const notes = []
const layerPage = `<!doctype html><script>
let port
let saved = []
const send = (message) => port && port.postMessage(message)
const show = () => send({ type: "notes", notes: saved.map((note) => note.shown) })
addEventListener("message", (event) => {
  if (event.data?.cool !== "site" || event.data.type !== "hello") return
  port = event.ports[0]
  port.onmessage = async ({ data }) => {
    if (data.type !== "pick") return
    const response = await fetch("/notes", { method: "POST", body: JSON.stringify(data.anchor) })
    saved = await response.json()
    show()
  }
  fetch("/notes").then((response) => response.json()).then((list) => {
    saved = list
    send({ type: "status", state: "ready" })
    show()
  })
})
parent.postMessage({ cool: "layer", type: "ready", version: 1 }, "*")
</script>`

// What the script draws: an open note's ink is its screen and mark paths.
const shown = (note) => ({
  id: note.id,
  n: note.n,
  state: "open",
  anchor: {
    ...note.anchor,
    ink: note.anchor.ink && {
      screen: note.anchor.ink.screen,
      paths: note.anchor.ink.marks.map((mark) => mark.path),
    },
  },
})

createServer(async (request, response) => {
  const url = new URL(request.url, layerOrigin)
  if (url.pathname === "/comment-layer") {
    response.writeHead(200, { "content-type": types[".html"] })
    return response.end(layerPage)
  }
  if (url.pathname === "/notes") {
    if (request.method === "POST") {
      let body = ""
      for await (const chunk of request) body += chunk
      const anchor = JSON.parse(body)
      const n = notes.length + 1
      notes.push({ id: `note_${n}`, n, anchor })
    } else if (request.method === "DELETE") {
      notes.length = 0
    }
    response.writeHead(200, { "content-type": types[".json"] })
    return response.end(
      JSON.stringify(
        url.searchParams.has("agent")
          ? notes
          : notes.map((note) => ({ ...note, shown: shown(note) }))
      )
    )
  }
  response.writeHead(404)
  response.end()
}).listen(port + 1, "127.0.0.1")

console.log(`Site  http://127.0.0.1:${port}/`)
console.log(`Notes ${layerOrigin}/notes?agent`)
