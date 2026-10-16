#!/usr/bin/env node
// Applies agent ops to a DOCX and prints one JSON result per op:
//   react-docx-agent <in.docx> <out.docx> '[{"op":"suggest","paragraphId":"p-n1","find":"a","replace":"b","author":"Agent"}]'
// Ops: paragraphs, find, locate, suggest, comment, reply, resolve, comments, revisions, accept, reject.
import { readFile, writeFile } from "node:fs/promises";
import { openDocx } from "./dist/agent.js";

const [input, output, opsJson] = process.argv.slice(2);
if (!input || !output || !opsJson) {
  console.error("usage: react-docx-agent <in.docx> <out.docx> '<json ops array>'");
  process.exit(2);
}

const doc = await openDocx(await readFile(input));
const run = {
  paragraphs: () => doc.paragraphs(),
  find: ({ text, ...options }) => doc.find(text, options),
  locate: ({ anchor }) => doc.locate(anchor),
  suggest: (args) => doc.suggest(args),
  comment: (args) => doc.comment(args),
  reply: (args) => doc.reply(args),
  resolve: ({ commentId, resolved }) => doc.resolve(commentId, resolved),
  comments: () => doc.comments(),
  revisions: () => doc.revisions(),
  accept: ({ id }) => doc.accept(id),
  reject: ({ id }) => doc.reject(id),
};

const results = [];
for (const { op, ...args } of JSON.parse(opsJson)) {
  if (!run[op]) {
    console.error(`unknown op "${op}"; use one of: ${Object.keys(run).join(", ")}`);
    process.exit(2);
  }
  try {
    results.push({ op, result: run[op](args) ?? null });
  } catch (error) {
    console.log(JSON.stringify([...results, { op, error: error.message }], null, 2));
    process.exit(1);
  }
}
await writeFile(output, await doc.save());
console.log(JSON.stringify(results, null, 2));
