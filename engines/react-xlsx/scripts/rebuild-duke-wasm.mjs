import { execFileSync } from "node:child_process";
import { createHash } from "node:crypto";
import { copyFileSync, readFileSync, writeFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

const source = fileURLToPath(new URL("../../duke-sheets/", import.meta.url));
const remote = fileURLToPath(new URL("../../../tools/modal/r", import.meta.url));
const vendor = new URL("../packages/react-xlsx/vendor/sheets-wasm/", import.meta.url);
const git = (...args) => execFileSync("git", ["-C", source, ...args], { encoding: "utf8" });
const run = (...args) => execFileSync(remote, args, { stdio: "inherit" });
const sha256 = (bytes) => createHash("sha256").update(bytes).digest("hex");
const allowDirty = process.argv.slice(2).join(" ") === "--allow-dirty";
if (process.argv.length > 2 && !allowDirty) throw new Error("Usage: sh scripts/rebuild-duke-wasm.sh [--allow-dirty]");

function sourceState() {
  const files = [...new Set(git("ls-files", "-co", "--exclude-standard", "-z").split("\0").filter(Boolean))].sort();
  const hash = createHash("sha256");
  for (const file of files) {
    hash.update(file).update("\0").update(sha256(readFileSync(`${source}/${file}`))).update("\n");
  }
  return {
    dirty: git("status", "--porcelain", "--untracked-files=all", "--", ".").length > 0,
    sourceSha256: hash.digest("hex"),
  };
}

const before = sourceState();
if (before.dirty && !allowDirty) {
  throw new Error("Commit the reviewed duke-sheets changes before the final rebuild. For development verification only, pass --allow-dirty.");
}
const command = "wasm-pack build bindings/wasm --target web --release --out-dir pkg -- --locked";
run("up");
run("sync", source);
run("run", source, command);
run("pull", source, "bindings/wasm/pkg");
if (JSON.stringify(before) !== JSON.stringify(sourceState())) {
  throw new Error("duke-sheets source changed during the build; rerun before installing the artifact.");
}
for (const file of ["duke_sheets_wasm.js", "duke_sheets_wasm.d.ts", "duke_sheets_wasm_bg.wasm", "duke_sheets_wasm_bg.wasm.d.ts"]) {
  copyFileSync(`${source}/bindings/wasm/pkg/${file}`, new URL(file, vendor));
}
copyFileSync(`${source}/LICENSE`, new URL("LICENSE", vendor));
writeFileSync(new URL("BUILD.json", vendor), JSON.stringify({
  dukeSheets: before,
  build: { command, platform: "Modal Linux" },
  artifact: { file: "duke_sheets_wasm_bg.wasm", sha256: sha256(readFileSync(new URL("duke_sheets_wasm_bg.wasm", vendor))) },
}, null, 2) + "\n");
