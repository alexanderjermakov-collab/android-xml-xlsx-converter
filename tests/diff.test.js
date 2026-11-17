"use strict";

const assert = require("node:assert/strict");
const { compare } = require("../dist/diff.js");

function record(value, id) {
  return { value, id, app: "app", language: "values-sk", xmlPath: "APK/app/res/values-sk/strings.xml", filename: "strings.xml", resourceType: "string", item: null };
}

const source = new Map([
  ["same", record("Same", "same")],
  ["changed", record("Old", "changed")],
  ["removed", record("Removed", "removed")],
]);
const output = new Map([
  ["same", record("Same", "same")],
  ["changed", record("New", "changed")],
  ["added", record("Added", "added")],
]);

const result = compare(source, output);
assert.deepEqual(
  { matched: result.matched, changed: result.changed, onlySource: result.onlySource, onlyOutput: result.onlyOutput, different: result.different, total: result.total },
  { matched: 1, changed: 1, onlySource: 1, onlyOutput: 1, different: 3, total: 4 },
);
assert.equal(result.differences.length, 3);
assert.deepEqual(new Set(result.differences.map((row) => row.status)), new Set(["Changed", "Only in source", "Only in output"]));

console.log("diff engine tests passed");
