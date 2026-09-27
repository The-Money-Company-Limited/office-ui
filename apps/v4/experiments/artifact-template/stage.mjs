import {
  copyFile,
  cp,
  mkdir,
  readdir,
  rm,
  stat,
  writeFile,
} from "node:fs/promises"
import path from "node:path"
import { fileURLToPath } from "node:url"

const root = path.dirname(fileURLToPath(import.meta.url))
const dist = path.join(root, "dist")
const source = process.argv[2]
const output = process.argv[3] ? path.resolve(process.argv[3]) : dist

if (!source || process.argv.length > 4) {
  throw new Error(
    "Usage: pnpm --filter v4 artifact-template:stage /path/to/file.docx|xlsx|pptx|pdf [new-output-directory]"
  )
}

const kind = path.extname(source).slice(1).toLowerCase()
if (!["docx", "xlsx", "pptx", "pdf"].includes(kind)) {
  throw new Error("Only DOCX, XLSX, PPTX and PDF files are supported")
}

const name = path.basename(source)
if (
  !/^[a-zA-Z0-9][a-zA-Z0-9._-]*\.(docx|xlsx|pptx|pdf)$/.test(name) ||
  !name.endsWith(`.${kind}`)
) {
  throw new Error(
    "Use an ASCII file name containing only letters, numbers, dots, underscores and dashes"
  )
}

const info = await stat(source)
if (!info.isFile() || info.size > 25 * 1024 * 1024) {
  throw new Error("The source must be a regular file no larger than 25 MiB")
}

await stat(path.join(dist, "index.html"))
let createdOutput = false
try {
  if (output !== dist) {
    await mkdir(path.dirname(output), { recursive: true })
    try {
      await mkdir(output)
    } catch (error) {
      if (error.code === "EEXIST") {
        throw new Error(`Output directory already exists: ${output}`)
      }
      throw error
    }
    createdOutput = true
    await cp(dist, output, { recursive: true })
  }

  const documentDir = path.join(output, "document")
  await rm(documentDir, { recursive: true, force: true })
  await mkdir(documentDir)
  await copyFile(source, path.join(documentDir, name))
  await writeFile(
    path.join(output, "artifact.json"),
    JSON.stringify({ kind, title: name, file: `document/${name}` }) + "\n"
  )
  await writeFile(path.join(output, "cool.json"), '{"spa":false}\n')

  let count = 0
  let bytes = 0
  async function checkDirectory(directory) {
    for (const entry of await readdir(directory, { withFileTypes: true })) {
      const item = path.join(directory, entry.name)
      if (entry.isDirectory()) {
        await checkDirectory(item)
      } else if (entry.isFile()) {
        const { size } = await stat(item)
        if (size > 25 * 1024 * 1024) {
          throw new Error(
            `${path.relative(output, item)} exceeds the 25 MiB asset limit`
          )
        }
        count++
        bytes += size
      } else {
        throw new Error(`${path.relative(output, item)} is not a regular file`)
      }
    }
  }

  await checkDirectory(output)
  if (count > 20_000 || bytes > 1024 * 1024 * 1024) {
    throw new Error(
      "Release exceeds the artifact file count or total size limit"
    )
  }
  console.log(
    `Staged ${kind.toUpperCase()} artifact: ${count} files, ${bytes} bytes in ${output}`
  )
} catch (error) {
  if (createdOutput) await rm(output, { recursive: true, force: true })
  throw error
}
