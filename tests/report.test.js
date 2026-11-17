"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");
const XLSX = require("../dist/assets/xlsx-js-style.bundle.js");
const JSZip = require("../dist/assets/jszip.min.js");
const CorrectionEngine = require("../dist/corrections.js");
const DiffEngine = require("../dist/diff.js");

function fakeNode() {
  return {
    addEventListener() {},
    append() {},
    click() {},
    remove() {},
    replaceChildren() {},
    querySelectorAll() { return []; },
    classList: { add() {}, remove() {}, toggle() {} },
    dataset: {},
    files: [],
    textContent: "",
    value: "",
  };
}

const nodes = new Map();
const context = {
  console,
  JSZip,
  XLSX,
  CorrectionEngine,
  DiffEngine,
  TextEncoder,
  TextDecoder,
  Uint8Array,
  Blob,
  File,
  Map,
  Set,
  Date,
  URL,
  crypto,
  setTimeout,
  clearTimeout,
  fetch: async () => ({ ok: true, json: async () => ({ headers: ["values", "values-sk"], rows: [] }) }),
  document: {
    querySelector(selector) {
      if (!nodes.has(selector)) nodes.set(selector, fakeNode());
      return nodes.get(selector);
    },
    createElement() { return fakeNode(); },
    body: fakeNode(),
  },
};
vm.createContext(context);
const source = fs.readFileSync(require.resolve("../dist/app.js"), "utf8");
vm.runInContext(`${source}\nglobalThis.__test = { archiveRecordKey, diffWorkbook, resourcePathTemplate, resourceIdentityKey };`, context);
const { archiveRecordKey, diffWorkbook, resourcePathTemplate, resourceIdentityKey } = context.__test;

assert.equal(
  resourcePathTemplate("APK/TunerInput/res/atv_res/values-sk/strings.xml", "values-sk"),
  "APK/TunerInput/res/atv_res/{values}/strings.xml",
);
assert.notEqual(
  resourceIdentityKey("APK/App/res/{values}/arrays.xml", "string-array", "same", 0),
  resourceIdentityKey("APK/App/res/{values}/strings.xml", "string-array", "same", 0),
);
assert.notEqual(
  resourceIdentityKey("APK/App/res/{values}/strings.xml", "string", "same", null),
  resourceIdentityKey("APK/App/res/{values}/strings.xml", "string-array", "same", null),
);

const xmlPath = "APK/App/res/values-sk/strings.xml";
const key = archiveRecordKey(xmlPath, "string", "greeting", null);
const sourceRecord = {
  app: "App", language: "values-sk", xmlPath, filename: "strings.xml",
  resourceType: "string", id: "greeting", item: null, value: "staré",
};
const outputRecord = { ...sourceRecord, value: "nové" };
const sourceArchive = { records: new Map([[key, sourceRecord]]), xmlFiles: 1 };
const outputArchive = { records: new Map([[key, outputRecord]]), xmlFiles: 1 };
const result = DiffEngine.compare(sourceArchive.records, outputArchive.records);
const modifications = new Map([[key, {
  key,
  app: "App",
  id: "greeting",
  item: null,
  language: "values-sk",
  xmlPath,
  filename: "strings.xml",
  resourceType: "string",
  reasons: ["Correction file: Slovak corrections.xlsx (values-sk)"],
  checks: [{
    file: "Slovak corrections.xlsx", language: "values-sk", sourceRow: 2,
    originalValue: "staré", expectedValue: "nové",
  }],
}], ["unmatched:Slovak corrections.xlsx\u001f9\u001fvalues-sk", {
  key: "unmatched:Slovak corrections.xlsx\u001f9\u001fvalues-sk",
  app: "MissingApp", id: "missing", item: null, language: "values-sk", xmlPath: "", filename: "", resourceType: "",
  reasons: ["Correction rule did not match any workbook row: Slovak corrections.xlsx (values-sk)"],
  checks: [{
    file: "Slovak corrections.xlsx", language: "values-sk", sourceRow: 9,
    originalValue: "old", expectedValue: "new", targetMissing: true,
  }],
}]]);

const workbook = diffWorkbook(
  { name: "source.zip" },
  { name: "output.zip" },
  sourceArchive,
  outputArchive,
  result,
  modifications,
);
const summary = XLSX.utils.sheet_to_json(workbook.Sheets.Summary, { header: 1, raw: true, defval: null });
const differences = XLSX.utils.sheet_to_json(workbook.Sheets.Differences, { header: 1, raw: true, defval: null });
const summaryHeader = summary.find((row) => row[0] === "Metric");
const slovakColumn = summaryHeader.indexOf("values-sk");
assert.equal(slovakColumn, 1, "corrected language must be the first language column");
assert.equal(summary.find((row) => row[0] === "Verified correction checks")[slovakColumn], 1);
assert.equal(summary.find((row) => row[0] === "Mismatched correction checks")[slovakColumn], 0);
assert.equal(summary.find((row) => row[0] === "Missing correction outputs")[slovakColumn], 0);
assert.equal(summary.find((row) => row[0] === "Unmatched correction rules")[slovakColumn], 1);
assert.ok(summary.some((row) => row[0] === "How the comparison works"));
assert.ok(summary.some((row) => row[0] === "1. Exact resource identity"));
assert.ok(summary.some((row) => row[0] === "Output ZIP vs correction XLSX files"));
const differenceHeader = differences.find((row) => row[0] === "Status");
assert.ok(differenceHeader.includes("values-sk correction source"));
assert.ok(differenceHeader.includes("values-sk correction original value"));
assert.ok(differenceHeader.includes("values-sk expected corrected value"));
const verificationColumn = differenceHeader.indexOf("values-sk verification");
const differenceHeaderIndex = differences.indexOf(differenceHeader);
assert.ok(differences.slice(differenceHeaderIndex + 1).some((row) => String(row[verificationColumn]).includes("MATCH")));
assert.ok(differences.slice(differenceHeaderIndex + 1).some((row) => String(row[verificationColumn]).includes("TARGET NOT FOUND")));
assert.match(workbook.Sheets.Differences["!autofilter"].ref, /^A7:O9$/);
assert.equal(workbook.Sheets.Differences.A2.s.font.bold, true);
assert.equal(workbook.Sheets.Differences.B5.s.fill.fgColor.rgb, "FFC7CE");
const roundTripWorkbook = XLSX.read(XLSX.write(workbook, { bookType: "xlsx", type: "buffer", cellStyles: true }), { type: "buffer", cellStyles: true });
assert.equal(roundTripWorkbook.Sheets.Differences.B5.s.fgColor.rgb, "FFC7CE");

const normalizedLineEndings = context.__test.diffWorkbook(
  { name: "source.zip" }, { name: "output.zip" }, sourceArchive,
  { records: new Map([[key, { ...outputRecord, value: "line 1\nline 2" }]]), xmlFiles: 1 },
  DiffEngine.compare(sourceArchive.records, new Map([[key, { ...outputRecord, value: "line 1\nline 2" }]])),
  new Map([[key, { ...modifications.get(key), checks: [{
    file: "Slovak corrections.xlsx", language: "values-sk", sourceRow: 2,
    originalValue: "staré", expectedValue: "line 1\r\nline 2",
  }] }]]),
);
const normalizedRows = XLSX.utils.sheet_to_json(normalizedLineEndings.Sheets.Differences, { header: 1, raw: true, defval: null });
const normalizedHeader = normalizedRows.find((row) => row[0] === "Status");
const normalizedVerification = normalizedHeader.indexOf("values-sk verification");
assert.ok(normalizedRows.slice(normalizedRows.indexOf(normalizedHeader) + 1).some((row) => String(row[normalizedVerification]).includes("MATCH (XML line endings normalized)")));
const normalizedSummary = XLSX.utils.sheet_to_json(normalizedLineEndings.Sheets.Summary, { header: 1, raw: true, defval: null });
assert.equal(normalizedSummary.find((row) => row[0] === "XML newline normalizations")[1], 1);

console.log("report and resource identity tests passed");
