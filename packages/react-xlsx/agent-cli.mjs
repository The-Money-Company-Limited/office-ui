#!/usr/bin/env node
// Applies a JSON list of operations to a workbook with the headless agent API.
//
//   agent-cli.mjs in.xlsx out.xlsx '[{"op":"comment","sheet":"Sheet1","ref":"B2","text":"Why?","author":"Agent"}]'
//
// Operations: sheets, cells {sheet, range?}, comments, suggestions,
// comment {sheet, ref, text, author}, reply {id, text, author}, resolve {id, resolved?},
// delete {id}, suggest {sheet, ref, value, author, note?}, accept {id|"all", author?},
// reject {id|"all", author?}. An id of "$N" means the thread created or changed by
// operation N. Prints one JSON result per operation and writes out.xlsx.
import { readFileSync, writeFileSync } from "node:fs";
import { openXlsx } from "./dist/agent.js";

const [input, output, operationsJson] = process.argv.slice(2);
if (!input || !output || !operationsJson) {
  console.error("Usage: agent-cli.mjs in.xlsx out.xlsx '<json operations>'");
  process.exit(2);
}

const workbook = openXlsx(new Uint8Array(readFileSync(input)));
const operations = JSON.parse(operationsJson);
const results = [];

const threadIdOf = (result) => (typeof result === "string" ? result : result?.id ?? result?.[0]);
const resolveId = (id) => {
  const reference = typeof id === "string" ? /^\$(\d+)$/.exec(id) : null;
  return reference ? threadIdOf(results[Number(reference[1])]) : id;
};

for (const [index, { op, ...args }] of (Array.isArray(operations) ? operations : [operations]).entries()) {
  if ("id" in args) {
    args.id = resolveId(args.id);
  }
  try {
    switch (op) {
      case "sheets": results.push(workbook.sheets()); break;
      case "cells": results.push(workbook.cells(args.sheet, args.range)); break;
      case "comments": results.push(workbook.comments()); break;
      case "suggestions": results.push(workbook.suggestions()); break;
      case "comment": results.push(workbook.comment(args)); break;
      case "reply": results.push(workbook.reply(args)); break;
      case "resolve": results.push(workbook.resolve(args)); break;
      case "delete": results.push(workbook.delete(args) ?? args.id); break;
      case "suggest": results.push(workbook.suggest(args)); break;
      case "accept": results.push(workbook.accept(args.id, args.author)); break;
      case "reject": results.push(workbook.reject(args.id, args.author)); break;
      default: throw new Error(`Unknown operation "${op}".`);
    }
  } catch (error) {
    console.error(`Operation ${index} (${op}) failed: ${error.message}`);
    process.exit(1);
  }
}

writeFileSync(output, workbook.save());
console.log(JSON.stringify(results, null, 2));
