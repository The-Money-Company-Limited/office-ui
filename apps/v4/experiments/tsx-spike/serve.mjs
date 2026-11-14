// Static server that answers like the Cool Computers gateway: ETag on every
// file, 304 for a matching If-None-Match, and either the public policy
// (max-age=0, must-revalidate) or the private one (private, no-store).
// Text types are gzipped as the edge would; .ts/.tsx have no content type in
// RuntimeVM's table, so they go uncompressed and untyped.
// It counts the bytes it sends: GET /__stats returns and resets the counters.
import { createHash } from "node:crypto"
import { readFile, stat } from "node:fs/promises"
import { execFileSync } from "node:child_process"
import { existsSync, readFileSync } from "node:fs"
import http from "node:http"
import http2 from "node:http2"
import path from "node:path"
import { gzipSync } from "node:zlib"

const types = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".mjs": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json",
  ".wasm": "application/wasm",
  ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  ".woff2": "font/woff2",
  ".svg": "image/svg+xml",
}
// Stand-in for the injected `cool: "doc"` page bridge: grants saving and
// acknowledges every save, so the editor opens in editing mode.
const bridge = `<script>(() => { let v = 1; addEventListener("message", (e) => { const d = e.data; if (e.source !== window || d?.cool !== "doc") return; const send = (x) => postMessage({ cool: "doc", ...x }, location.origin); if (d.type === "hello") send({ type: "access", canSave: true, version: String(v) }); if (d.type === "save") send({ type: "saved", id: d.id, version: String(++v) }) }) })()</script>`
const compressible = new Set([".html", ".js", ".mjs", ".css", ".json", ".wasm", ".svg"])

// latency: milliseconds added before each answer, a stand-in for the round
// trip to the edge. tls: serve HTTP/2 over a self-signed certificate, as the
// edge does, so many parallel requests share one connection.
export function serve(root, { port = 0, policy = "public", latency = 0, tls = false } = {}) {
  const memo = new Map()
  let stats = { requests: 0, bytes: 0, notModified: 0 }
  async function file(p) {
    const info = await stat(p)
    const key = p + ":" + info.mtimeMs + ":" + info.size
    if (memo.has(key)) return memo.get(key)
    let body = await readFile(p)
    if (path.basename(p) === "index.html") body = Buffer.from(body.toString().replace("<head>", "<head>" + bridge))
    const ext = path.extname(p)
    const entry = {
      body,
      gzip: compressible.has(ext) ? gzipSync(body, { level: 6 }) : null,
      etag: `"${createHash("sha1").update(body).digest("hex")}"`,
      type: types[ext],
    }
    memo.set(key, entry)
    return entry
  }
  const handler = async (req, res) => {
    const url = new URL(req.url, "http://x")
    if (latency && url.pathname !== "/__stats") await new Promise((done) => setTimeout(done, latency))
    if (url.pathname === "/__stats") {
      res.setHeader("content-type", "application/json")
      res.end(JSON.stringify(stats))
      stats = { requests: 0, bytes: 0, notModified: 0 }
      return
    }
    stats.requests++
    let p = path.join(root, decodeURIComponent(url.pathname))
    if (!p.startsWith(root)) return res.writeHead(403).end()
    try {
      if ((await stat(p)).isDirectory()) p = path.join(p, "index.html")
      const entry = await file(p)
      const headers = {
        etag: entry.etag,
        "cache-control": policy === "private" ? "private, no-store" : "public, max-age=0, must-revalidate",
        "x-content-type-options": "nosniff",
      }
      if (entry.type) headers["content-type"] = entry.type
      if (req.headers["if-none-match"] === entry.etag) {
        stats.notModified++
        res.writeHead(304, headers).end()
        return
      }
      let body = entry.body
      if (entry.gzip && /\bgzip\b/.test(req.headers["accept-encoding"] || "")) {
        body = entry.gzip
        headers["content-encoding"] = "gzip"
      }
      headers["content-length"] = body.length
      stats.bytes += body.length
      res.writeHead(200, headers).end(body)
    } catch {
      res.writeHead(404, { "cache-control": "no-store" }).end()
    }
  }
  let server
  if (tls) {
    const dir = "/tmp/tsxspike"
    if (!existsSync(`${dir}/key.pem`))
      execFileSync("openssl", ["req", "-x509", "-newkey", "rsa:2048", "-nodes", "-keyout", `${dir}/key.pem`, "-out", `${dir}/cert.pem`, "-days", "2", "-subj", "/CN=127.0.0.1"], { stdio: "ignore" })
    server = http2.createSecureServer({ key: readFileSync(`${dir}/key.pem`), cert: readFileSync(`${dir}/cert.pem`), allowHTTP1: true }, handler)
  } else server = http.createServer(handler)
  return new Promise((done) => server.listen(port, "127.0.0.1", () => done({ server, port: server.address().port })))
}

if (import.meta.url === `file://${process.argv[1]}`) {
  const { port } = await serve(path.resolve(process.argv[2]), { port: Number(process.argv[3] || 4173), policy: process.argv[4] })
  console.log(`serving ${process.argv[2]} on ${port}`)
}
