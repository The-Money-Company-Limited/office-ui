// template.json names the template a release was staged from, so a platform
// that keeps several can tell which one serves an artifact: this repository's
// commit and the commit of each engine it ships.
import { execFileSync } from "node:child_process"
import { readFileSync, realpathSync } from "node:fs"
import path from "node:path"
import { fileURLToPath } from "node:url"

const root = path.dirname(fileURLToPath(import.meta.url))
const appRoot = path.resolve(root, "../..")

const describe = (cwd) =>
  execFileSync("git", ["describe", "--always", "--dirty"], {
    cwd,
    encoding: "utf8",
  }).trim()

// An engine linked from a fork checkout is named by its commit; one installed
// from npm by its version.
function engineVersion(name) {
  const directory = realpathSync(path.join(appRoot, "node_modules", name))
  if (!directory.includes(`${path.sep}node_modules${path.sep}`)) {
    return describe(directory)
  }
  return JSON.parse(readFileSync(path.join(directory, "package.json"), "utf8"))
    .version
}

export function templateJson(engines) {
  return (
    JSON.stringify({
      template: "office",
      version: describe(root),
      engines: Object.fromEntries(
        engines.map((name) => [name, engineVersion(name)])
      ),
    }) + "\n"
  )
}
