// Run through tools/modal/r. Disable only the new transaction to check the boundary tests.
import { readFileSync, writeFileSync } from "node:fs";
import { spawnSync } from "node:child_process";
const path = new URL("../packages/react-xlsx/src/structural-edits.ts", import.meta.url);
const original = readFileSync(path, "utf8");
try {
  writeFileSync(path, original.replace("  const workbook = Workbook.fromBytes(bytes);", "  return bytes;\n  const workbook = Workbook.fromBytes(bytes);"));
  const result = spawnSync(process.execPath, ["--experimental-strip-types", "--test", "packages/react-xlsx/src/structural-edits.test.ts"], { encoding: "utf8" });
  const log = `${result.stdout}${result.stderr}`;
  writeFileSync(new URL("../.tmp-structural-red.log", import.meta.url), log);
  const failures = [...result.stdout.matchAll(/^not ok /gm)].length;
  if (result.status === 0 || failures !== 3) throw new Error(log);
  console.log("PASS negative control: all 3 structural tests fail with the transaction disabled.");
} finally {
  writeFileSync(path, original);
}
