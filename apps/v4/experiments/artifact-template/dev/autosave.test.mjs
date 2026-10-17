import assert from "node:assert/strict"
import { mkdir, readFile } from "node:fs/promises"
import { createRequire } from "node:module"
import path from "node:path"
import { test } from "node:test"
import { pathToFileURL } from "node:url"

const require = createRequire(import.meta.url)
const { build } = createRequire(require.resolve("vite"))("esbuild")
const output = path.resolve("apps/v4/experiments/artifact-template/dev/results")
await mkdir(output, { recursive: true })
let source = await readFile(
  new URL("../src/autosave.ts", import.meta.url),
  "utf8"
)
if (process.env.AUTOSAVE_NEGATIVE)
  source = source.replace(
    'this.page.addEventListener("message", this.message)',
    ""
  )
await build({
  stdin: { contents: source, loader: "ts" },
  bundle: true,
  format: "esm",
  platform: "node",
  outfile: `${output}/autosave.mjs`,
})
const { Autosave } = await import(pathToFileURL(`${output}/autosave.mjs`))

function setup(t) {
  t.mock.timers.enable({ apis: ["setTimeout"] })
  const page = new EventTarget()
  page.document = new EventTarget()
  page.location = { origin: "https://artifact.example" }
  page.scrollX = 0
  page.scrollY = 100
  page.getSelection = () => null
  page.postMessage = (data, origin, transfer) =>
    messages.push({ data, origin, transfer })
  const messages = []
  const reloads = []
  const save = new Autosave(page, (version) => reloads.push(version))
  let contents = "first"
  save.register({
    serialize: () => ({
      files: [
        {
          path: "document/test.docx",
          bytes: new TextEncoder().encode(contents).buffer,
        },
      ],
    }),
    download: () => {},
  })
  const stop = save.start()
  t.after(stop)
  const reply = (data, origin = page.location.origin, source = page) => {
    const event = new Event("message")
    Object.assign(event, { data: { cool: "doc", ...data }, origin, source })
    page.dispatchEvent(event)
  }
  const access = () => reply({ type: "access", canSave: true, version: "v1" })
  const tick = async (ms) => {
    t.mock.timers.tick(ms)
    await Promise.resolve()
    await Promise.resolve()
  }
  const saves = () => messages.filter((message) => message.data.type === "save")
  return {
    page,
    messages,
    reloads,
    save,
    reply,
    access,
    tick,
    saves,
    set: (value) => {
      contents = value
      save.dirty()
    },
  }
}

test("same-window handshake retries and read-only never serializes", async (t) => {
  const h = setup(t)
  assert.equal(h.messages[0].data.type, "hello")
  await h.tick(3000)
  assert.equal(h.messages.length, 2)
  h.reply(
    { type: "access", canSave: true, version: "v1" },
    "https://other.example"
  )
  h.reply({ type: "access", canSave: true, version: "v1" }, undefined, {})
  h.set("forbidden")
  await h.tick(2000)
  assert.equal(h.save.getSnapshot().canSave, false)
  assert.equal(h.saves().length, 0)
  h.access()
  assert.equal(h.save.getSnapshot().canSave, true)
})

test("debounce, transfer, single flight, edits during save, and hide flush", async (t) => {
  const h = setup(t)
  h.access()
  h.set("one")
  await h.tick(1900)
  h.set("two")
  await h.tick(1900)
  assert.equal(h.saves().length, 0)
  await h.tick(100)
  const first = h.saves()[0]
  assert.equal(first.data.base, "v1")
  assert.equal(new TextDecoder().decode(first.data.files[0].bytes), "two")
  assert.equal(first.transfer[0], first.data.files[0].bytes)
  h.set("three")
  h.page.dispatchEvent(new Event("pagehide"))
  await h.tick(4000)
  assert.equal(h.saves().length, 1)
  h.reply({ type: "saved", id: first.data.id, version: "v2" })
  await h.tick(0)
  const second = h.saves()[1].data
  assert.equal(second.base, "v2")
  assert.equal(new TextDecoder().decode(second.files[0].bytes), "three")
  h.reply({ type: "saved", id: second.id, version: "v3" })
  assert.equal(h.save.getSnapshot().status, "Saved")
  h.set("four")
  h.page.document.visibilityState = "hidden"
  h.page.document.dispatchEvent(new Event("visibilitychange"))
  await h.tick(0)
  assert.equal(h.saves().length, 3)
  h.reply({ type: "changed", version: "v5" })
  h.reply({ type: "saved", id: h.saves()[2].data.id, version: "v4" })
  assert.deepEqual(h.reloads, ["v5"])
})

test("offline retries back off, preserve base, and ignore unrelated acknowledgements", async (t) => {
  const h = setup(t)
  h.access()
  h.set("one")
  await h.tick(2000)
  h.reply({ type: "saved", id: "unrelated", version: "wrong" })
  h.reply({ type: "save_failed", id: h.saves()[0].data.id, reason: "offline" })
  assert.equal(h.save.getSnapshot().status, "Offline, will retry")
  await h.tick(1999)
  assert.equal(h.saves().length, 1)
  await h.tick(1)
  assert.equal(h.saves()[1].data.base, "v1")
  h.reply({ type: "save_failed", id: h.saves()[1].data.id, reason: "offline" })
  h.set("two")
  await h.tick(3999)
  assert.equal(h.saves().length, 2)
  await h.tick(1)
  assert.equal(h.saves().length, 3)
  h.set("during the unacknowledged save")
  await h.tick(30_000)
  assert.equal(h.save.getSnapshot().status, "Offline, will retry")
})

test("changed reloads only clean documents; conflicts never overwrite", async (t) => {
  const h = setup(t)
  h.access()
  h.reply({ type: "changed", version: "v2" })
  assert.deepEqual(h.reloads, ["v2"])
  h.set("mine")
  h.reply({ type: "changed", version: "v3" })
  await h.tick(2000)
  assert.equal(h.saves()[0].data.base, "v2")
  h.reply({
    type: "save_failed",
    id: h.saves()[0].data.id,
    reason: "conflict",
    version: "v3",
  })
  assert.equal(h.save.getSnapshot().conflict, true)
  h.set("still mine")
  await h.tick(60_000)
  assert.equal(h.saves().length, 1)
  h.save.reloadTheirs()
  assert.deepEqual(h.reloads, ["v2", "v3"])
})

test("permanent failures retain local edits and do not retry", async (t) => {
  const h = setup(t)
  h.access()
  h.set("mine")
  await h.tick(2000)
  h.reply({
    type: "save_failed",
    id: h.saves()[0].data.id,
    reason: "forbidden",
  })
  assert.equal(h.save.getSnapshot().canSave, false)
  assert.match(h.save.getSnapshot().status, /permission/)
  await h.tick(60_000)
  assert.equal(h.saves().length, 1)
})
