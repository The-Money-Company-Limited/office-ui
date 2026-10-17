import assert from "node:assert/strict"
import { mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises"
import path from "node:path"
import { test } from "node:test"

import { validateManifest } from "../../../components/extend/slides/validate.js"
import { startHarness } from "./autosave-harness.mjs"

async function setup(t, kind = "docx") {
  const root = path.resolve("apps/v4/experiments/artifact-template/dev/results")
  await mkdir(root, { recursive: true })
  const folder = await mkdtemp(`${root}/contract-`)
  const file = kind === "slides" ? "deck/deck.json" : `document/sample.${kind}`
  await mkdir(path.join(folder, path.dirname(file)), { recursive: true })
  await writeFile(
    path.join(folder, "artifact.json"),
    JSON.stringify({ kind, file })
  )
  await writeFile(
    path.join(folder, file),
    kind === "slides" ? '{"order":["one"]}' : "original"
  )
  if (kind === "slides") {
    await mkdir(`${folder}/deck/slides`)
    await writeFile(
      `${folder}/deck/slides/one.html`,
      '<img src="../existing.png">'
    )
  }
  const server = await startHarness(folder)
  t.after(async () => {
    server.closeAllConnections()
    await new Promise((resolve) => server.close(resolve))
    await rm(folder, { recursive: true })
  })
  const url = `http://127.0.0.1:${server.address().port}`
  const post = async (route, data) =>
    (
      await fetch(`${url}/__doc/${route}`, {
        method: "POST",
        body: JSON.stringify(data),
      })
    ).json()
  const save = (files, base = "1", id = crypto.randomUUID()) =>
    post("save", {
      id,
      base,
      files: Object.entries(files).map(([path, bytes]) => ({
        path,
        bytes: Buffer.from(bytes).toString("base64"),
      })),
    })
  return { folder, file, save, post, url }
}

for (const kind of ["docx", "xlsx", "pptx", "pdf"]) {
  test(`${kind} accepts only artifact.json.file and rejects invalid destinations`, async (t) => {
    const h = await setup(t, kind)
    for (const name of [
      `document/other.${kind}`,
      "artifact.json",
      "index.html",
      "document/../outside",
      "/document/sample",
      "document\\sample",
      "document//sample",
      "document/./sample",
    ]) {
      assert.equal(
        (await h.save({ [name]: "bad" })).reason,
        "path_not_allowed",
        name
      )
    }
    assert.equal(
      await readFile(path.join(h.folder, h.file), "utf8"),
      "original"
    )
    assert.equal((await h.save({ [h.file]: "saved" })).type, "saved")
  })
}

test("malformed requests answer error immediately and do not poison the next save", async (t) => {
  const h = await setup(t)
  const file = { path: h.file, bytes: "YQ==" }
  for (const patch of [
    { id: "" },
    { id: "a".repeat(129) },
    { base: "" },
    { base: "a".repeat(129) },
    { files: [] },
    { files: [file, file] },
    { files: [{ path: h.file }] },
  ]) {
    assert.equal(
      (await h.post("save", { id: "id", base: "1", files: [file], ...patch }))
        .reason,
      "error"
    )
  }
  assert.equal(
    (await h.save({ [h.file]: "ok" }, "1", "a".repeat(128))).type,
    "saved"
  )
})

test("receipts precede conflict checks and survive newer writes", async (t) => {
  const h = await setup(t)
  const first = await h.save({ [h.file]: "first" }, "1", "first")
  await h.save({ [h.file]: "second" }, first.version, "second")
  assert.deepEqual(await h.save({ [h.file]: "first" }, "1", "first"), first)
  assert.equal(await readFile(path.join(h.folder, h.file), "utf8"), "second")
  assert.equal((await h.post("control", {})).saves.length, 2)
})

test("receipt retention is bounded to 32 and 24 hours", async (t) => {
  const h = await setup(t)
  for (let index = 1; index <= 33; index++)
    assert.equal(
      (await h.save({ [h.file]: String(index) }, String(index), String(index)))
        .type,
      "saved"
    )
  assert.equal((await h.post("control", {})).receipts, 32)
  assert.equal((await h.save({ [h.file]: "1" }, "1", "1")).reason, "conflict")
  assert.equal((await h.save({ [h.file]: "2" }, "2", "2")).version, "3")
  await h.post("control", { advance: 24 * 60 * 60_000 })
  assert.equal((await h.save({ [h.file]: "2" }, "2", "2")).reason, "conflict")
})

test("offline includes retryAfter and not_document is reported on access and save", async (t) => {
  const h = await setup(t)
  await h.post("control", { mode: "offline", retryAfter: 12 })
  const failure = await h.save({ [h.file]: "one" })
  assert.equal(failure.reason, "offline")
  assert.equal(failure.retryAfter, 12)
  await h.post("control", { mode: "not_document" })
  assert.equal((await h.save({ [h.file]: "one" })).reason, "not_document")
  const access = await (await fetch(`${h.url}/__doc/access`)).json()
  assert.equal(access.canSave, false)
  assert.equal(access.reason, "not_document")
})

test("slides enforce id grammar and images referenced by the resulting deck", async (t) => {
  const h = await setup(t, "slides")
  for (const name of [
    "deck/evil.js",
    "deck/other.json",
    "deck/slides/a.b.html",
    `deck/slides/${"a".repeat(65)}.html`,
    "deck/unreferenced.png",
  ]) {
    assert.equal(
      (await h.save({ [name]: "bad" })).reason,
      "path_not_allowed",
      name
    )
  }
  assert.equal((await h.save({ "deck/existing.png": "image" })).type, "saved")
  const files = {
    "deck/deck.json": '{"order":["new_-1"]}',
    "deck/slides/new_-1.html":
      '<img src="../new.JPG"><style>.x{background:url(/deck/bg.svg)}</style><div style="background:url(../inline.webp)"></div>',
    "deck/new.JPG": "new",
    "deck/bg.svg": "bg",
    "deck/inline.webp": "inline",
  }
  assert.equal((await h.save(files, "2")).type, "saved")
  assert.equal(
    (await h.save({ "deck/existing.png": "no longer referenced" }, "3")).reason,
    "path_not_allowed"
  )
  assert.equal(
    (
      await h.save(
        {
          "deck/slides/new_-1.html":
            '<!-- <img src="../comment.png"> --><img src="https://external.test/deck/external.png">',
          "deck/comment.png": "bad",
        },
        "3"
      )
    ).reason,
    "path_not_allowed"
  )
  assert.equal(
    (await h.save({ [`deck/slides/${"a".repeat(64)}.html`]: "ok" }, "3")).type,
    "saved"
  )
})

test("25 MiB per file and 64 MiB combined are inclusive limits", async (t) => {
  const h = await setup(t, "slides")
  const mib = 1024 * 1024
  assert.equal(
    (await h.save({ "deck/slides/big.html": Buffer.alloc(25 * mib + 1) }))
      .reason,
    "too_large"
  )
  const files = {
    "deck/slides/a.html": Buffer.alloc(25 * mib),
    "deck/slides/b.html": Buffer.alloc(25 * mib),
    "deck/slides/c.html": Buffer.alloc(14 * mib + 1),
  }
  assert.equal((await h.save(files)).reason, "too_large")
  files["deck/slides/c.html"] = Buffer.alloc(14 * mib)
  assert.equal((await h.save(files)).type, "saved")
})

test("slide manifests use the same 1–64 character IDs as save paths", () => {
  for (const id of ["0", "_", "-", "a".repeat(64)])
    assert.equal(validateManifest({ title: "Deck", order: [id] }).order[0], id)
  for (const id of ["", "a".repeat(65), "a.b", "../slide"])
    assert.throws(() => validateManifest({ title: "Deck", order: [id] }))
})
