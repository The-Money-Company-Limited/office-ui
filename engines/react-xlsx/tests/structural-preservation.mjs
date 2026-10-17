import assert from "node:assert/strict";
import { createRequire } from "node:module";
const require = createRequire(new URL("../packages/react-xlsx/package.json", import.meta.url));
const { strFromU8, strToU8, unzipSync, zipSync } = require("fflate");

export function addPreservationChart(workbook) {
  workbook.getSheet(0).addDrawing({
    kind: "chart",
    anchor: { type: "twoCell", from: { col: 5, row: 1 }, to: { col: 10, row: 8 } },
    chart: {
      chartType: "BarClustered",
      series: [{ name: "Sheet1!$B$3", values: { refType: "formula", formula: "Sheet1!$B$3:$B$4" } }]
    }
  });
  const archive = unzipSync(workbook.saveXlsxBytes());
  const path = Object.keys(archive).find((name) => /^xl\/charts\/chart\d+\.xml$/.test(name));
  assert.ok(path, "fixture chart exists");
  const errorBars = '<c:errBars><c:errDir val="y"/><c:errBarType val="both"/><c:errValType val="cust"/><c:plus><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>7</c:v></c:pt><c:pt idx="1"><c:v>11</c:v></c:pt></c:numLit></c:plus><c:minus><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>5</c:v></c:pt></c:numLit></c:minus><c:spPr><a:ln w="28575"><a:solidFill><a:srgbClr val="0055AA"/></a:solidFill></a:ln></c:spPr></c:errBars>';
  archive[path] = strToU8(strFromU8(archive[path]).replace('</c:ser>', `${errorBars}</c:ser>`));
  return zipSync(archive);
}

export function assertPreservedErrorBars(bytes) {
  const archive = unzipSync(bytes);
  const charts = Object.entries(archive).filter(([path]) => /^xl\/charts\/chart\d+\.xml$/.test(path));
  assert.ok(charts.length > 0, "export retains chart");
  const xml = charts.map(([, bytes]) => strFromU8(bytes)).join("\n");
  assert.match(xml, /<c:plus>.*?<c:v>7<\/c:v>.*?<c:v>11<\/c:v>.*?<\/c:plus>/, "preserve positive custom error values");
  assert.match(xml, /<c:errBars>.*?<a:ln w="28575">.*?<a:srgbClr val="0055AA".*?<\/c:errBars>/, "preserve original error-bar formatting");
  assert.match(xml, /<c:minus>.*?<c:v>3<\/c:v>.*?<c:v>5<\/c:v>.*?<\/c:minus>/, "preserve negative custom error values");
}

export function assertChartReference(bytes, expected) {
  const archive = unzipSync(bytes);
  const charts = Object.entries(archive).filter(([path]) => /^xl\/charts\/chart\d+\.xml$/.test(path));
  const xml = charts.map(([, bytes]) => strFromU8(bytes)).join("\n");
  const formulas = [...xml.matchAll(/<(?:\w+:)?f>(.*?)<\/(?:\w+:)?f>/g)].map((match) => match[1]
    .replace(/&apos;/g, "'").replace(/&quot;/g, '"').replace(/&amp;/g, "&"));
  assert.ok(formulas.includes(expected), `chart reference ${expected}; actual ${JSON.stringify(formulas)}`);
  const name = expected.split(":")[0];
  const titles = [...xml.matchAll(/<(?:\w+:)?tx>(.*?)<\/(?:\w+:)?tx>/g)].map((match) => match[1]
    .replace(/&apos;/g, "'").replace(/&quot;/g, '"').replace(/&amp;/g, "&"));
  assert.ok(titles.some((title) => title.includes(`>${name}<`)), `chart series name reference ${name}`);
}
