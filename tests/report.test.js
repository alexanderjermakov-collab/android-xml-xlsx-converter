"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");
const XLSX = require("../dist/assets/xlsx.full.min.js");
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
const summaryHeader = summary.findIndex((row) => row[0] === "Language");
const slovakSummary = summary.slice(summaryHeader + 1).find((row) => row[0] === "values-sk");
assert.ok(slovakSummary);
assert.equal(slovakSummary[10], 1);
assert.equal(slovakSummary[11], 0);
assert.equal(slovakSummary[12], 0);
assert.equal(slovakSummary[13], 1);
assert.ok(differences[0].includes("values-sk correction source"));
assert.ok(differences[0].includes("values-sk correction original value"));
assert.ok(differences[0].includes("values-sk expected corrected value"));
const verificationColumn = differences[0].indexOf("values-sk verification");
assert.ok(differences.slice(1).some((row) => String(row[verificationColumn]).includes("MATCH")));
assert.ok(differences.slice(1).some((row) => String(row[verificationColumn]).includes("TARGET NOT FOUND")));
assert.match(workbook.Sheets.Differences["!autofilter"].ref, /^A1:O3$/);

console.log("report and resource identity tests passed");
