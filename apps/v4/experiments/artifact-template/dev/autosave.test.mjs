import assert from "node:assert/strict"
import { execFileSync } from "node:child_process"
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
if (process.env.AUTOSAVE_BASELINE)
  source = execFileSync(
    "git",
    ["show", "59895b6:apps/v4/experiments/artifact-template/src/autosave.ts"],
    { encoding: "utf8" }
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
  t.mock.timers.enable({ apis: ["setTimeout", "Date"] })
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
  await h.tick(20 * 60_000)
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
  assert.match(h.save.getSnapshot().status, /can no longer edit/)
  await h.tick(60_000)
  assert.equal(h.saves().length, 1)
})

test("an expired sign-in resumes saving when access returns", async (t) => {
  const h = setup(t)
  h.access()
  h.set("mine")
  await h.tick(2000)
  h.reply({ type: "save_failed", id: h.saves()[0].data.id, reason: "signed_out" })
  assert.equal(h.save.getSnapshot().canSave, false)
  assert.equal(h.save.getSnapshot().reason, "signed_out")
  assert.match(h.save.getSnapshot().status, /Sign in again/)
  await h.tick(60_000)
  assert.equal(h.saves().length, 1)
  h.access()
  assert.equal(h.save.getSnapshot().canSave, true)
  await h.tick(0)
  assert.equal(h.save.getSnapshot().status, "Saving…")
  assert.equal(h.saves().length, 2)
})

test("the access label waits for the first access answer", (t) => {
  const h = setup(t)
  assert.equal(h.save.getSnapshot().known, false)
  h.access()
  assert.equal(h.save.getSnapshot().known, true)
})

test("timeout retry retains id, base and transferred bytes despite newer edits", async (t) => {
  const h = setup(t)
  h.access()
  h.set("original write")
  await h.tick(2000)
  const first = structuredClone(h.saves()[0].data)
  structuredClone(h.saves()[0].data, { transfer: h.saves()[0].transfer })
  h.set("newer edit")
  await h.tick(19 * 60_000)
  assert.equal(h.save.getSnapshot().status, "Saving…")
  await h.tick(60_000)
  assert.equal(h.save.getSnapshot().status, "Offline, will retry")
  await h.tick(2000)
  assert.deepEqual(h.saves()[1].data, first)
  h.reply({ type: "saved", id: first.id, version: "v2" })
  await h.tick(0)
  const next = h.saves()[2].data
  assert.notEqual(next.id, first.id)
  assert.equal(next.base, "v2")
  assert.equal(new TextDecoder().decode(next.files[0].bytes), "newer edit")
})

test("late saved preserves newer edits and duplicate receipts cannot regress the base", async (t) => {
  const h = setup(t)
  h.access()
  h.set("first")
  await h.tick(2000)
  const first = h.saves()[0].data
  await h.tick(20 * 60_000)
  h.set("second")
  h.reply({ type: "saved", id: first.id, version: "v2" })
  await h.tick(0)
  assert.equal(h.saves().length, 2)
  const second = h.saves()[1].data
  assert.equal(second.base, "v2")
  assert.equal(new TextDecoder().decode(second.files[0].bytes), "second")
  h.reply({ type: "saved", id: second.id, version: "v3" })
  h.reply({ type: "saved", id: first.id, version: "v2" })
  h.set("third")
  await h.tick(2000)
  assert.equal(h.saves()[2].data.base, "v3")
})

test("beforeunload prompts only while dirty, including edits after a late receipt", async (t) => {
  const h = setup(t)
  const prompted = () =>
    !h.page.dispatchEvent(new Event("beforeunload", { cancelable: true }))
  h.access()
  assert.equal(prompted(), false)
  h.set("one")
  assert.equal(prompted(), true)
  await h.tick(2000)
  const first = h.saves()[0].data
  await h.tick(20 * 60_000)
  h.set("two")
  h.reply({ type: "saved", id: first.id, version: "v2" })
  assert.equal(prompted(), true)
  await h.tick(0)
  h.reply({ type: "saved", id: h.saves()[1].data.id, version: "v3" })
  assert.equal(prompted(), false)
  h.set("three")
  h.save.reloadTheirs()
  assert.equal(prompted(), false)
})

test("retryAfter is a minimum delay even when editing or hiding", async (t) => {
  const h = setup(t)
  h.access()
  h.set("one")
  await h.tick(2000)
  const first = h.saves()[0].data
  h.reply({
    type: "save_failed",
    id: first.id,
    reason: "offline",
    retryAfter: 60,
  })
  h.set("two")
  h.page.dispatchEvent(new Event("pagehide"))
  await h.tick(59_999)
  assert.equal(h.saves().length, 1)
  h.page.document.visibilityState = "hidden"
  h.page.document.dispatchEvent(new Event("visibilitychange"))
  await h.tick(0)
  assert.equal(h.saves().length, 1)
  await h.tick(1)
  assert.deepEqual(h.saves()[1].data, first)
})

for (const reason of ["path_not_allowed", "not_document"]) {
  test(`${reason} has its own message and retains unsaved edits`, async (t) => {
    const h = setup(t)
    h.access()
    h.set("one")
    await h.tick(2000)
    h.reply({ type: "save_failed", id: h.saves()[0].data.id, reason })
    assert.equal(h.save.getSnapshot().canSave, false)
    assert.match(
      h.save.getSnapshot().status,
      reason === "not_document"
        ? /This page is no longer a document/
        : /can't be saved/
    )
    assert.equal(
      h.page.dispatchEvent(new Event("beforeunload", { cancelable: true })),
      false
    )
    await h.tick(60_000)
    assert.equal(h.saves().length, 1)
  })
}

test("not_document access displays the document removal message", (t) => {
  const h = setup(t)
  h.reply({
    type: "access",
    canSave: false,
    version: "v1",
    reason: "not_document",
  })
  assert.equal(h.save.getSnapshot().canSave, false)
  assert.match(h.save.getSnapshot().status, /This page is no longer a document/)
})

test("a late receipt does not bypass retryAfter for the next edit", async (t) => {
  const h = setup(t)
  h.access()
  h.set("first")
  await h.tick(2000)
  const first = h.saves()[0].data
  h.reply({
    type: "save_failed",
    id: first.id,
    reason: "offline",
    retryAfter: 60,
  })
  h.set("second")
  h.reply({ type: "saved", id: first.id, version: "v2" })
  assert.notEqual(h.save.getSnapshot().status, "Saved")
  await h.tick(0)
  h.page.dispatchEvent(new Event("pagehide"))
  await h.tick(59_999)
  assert.equal(h.saves().length, 1)
  await h.tick(1)
  assert.equal(h.saves().length, 2)
  assert.equal(h.saves()[1].data.base, "v2")
})
