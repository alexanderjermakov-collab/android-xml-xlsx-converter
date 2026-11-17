"use strict";

const assert = require("node:assert/strict");
const engine = require("../dist/corrections.js");

const corrections = [
  ["Language", "Old text", "New text"],
  ["values-sk; cs", "staré", "nové"],
  ["all", "shared", "updated"],
];
const matrix = [
  ["App", "string ID", "item ID", "values", "values-sk", "values-cs", "values-tr"],
  ["app", "a", null, "Shared", "staré", "staré", "eski"],
  ["app", "b", null, "shared", "shared", "shared", "shared"],
  ["app", "c", null, "Settings", "nastavenia", "nastavení", "istanbul"],
  ["app", "d", null, "lowercase", "malé", "malé", "küçük"],
];

const rules = engine.parseRules(corrections, matrix[0].slice(3));
const replacements = engine.applyRules(matrix, rules);
assert.equal(replacements.replacements, 6);
assert.equal(replacements.unmatchedRules, 0);
assert.equal(matrix[1][4], "nové");
assert.equal(matrix[1][5], "nové");
assert.equal(matrix[2][3], "updated");
assert.equal(matrix[2][6], "updated");

const capitalization = engine.applyCapitalization(matrix);
assert.equal(capitalization.capitalized, 6);
assert.equal(capitalization.changes[0].language.startsWith("values-"), true);
assert.equal(matrix[3][4], "Nastavenia");
assert.equal(matrix[3][5], "Nastavení");
assert.equal(matrix[3][6], "İstanbul");
assert.equal(matrix[4][4], "malé");
assert.equal(engine.isUppercaseStart("…Hello"), true);
assert.equal(engine.uppercaseFirstLetter("(hello)", "en"), "(Hello)");

const keyedMatrix = [
  ["App", "string ID", "item ID", "values", "values-sk", "values-cs"],
  ["ATVInput", "choice", 2, "Development only", "Development only", "Pouze pro vývoj"],
  ["ATVInput", "shared", null, "Shared", "rovnaké", "stejné"],
];
const slovakFile = [
  ["App", "string ID", "item ID", "values", "values-sk (original)", "values-sk (corrected)"],
  ["ATVInput", "choice", 2, "Development only", "Development only", "Iba na vývoj"],
];
const czechFile = [
  ["App", "string ID", "item ID", "values", "values-cs (original)", "values-cs (corrected)"],
  ["ATVInput", "shared", null, "Shared", "stejné", "shodné"],
];
const keyedRules = [
  ...engine.parseRules(slovakFile, keyedMatrix[0].slice(3)),
  ...engine.parseRules(czechFile, keyedMatrix[0].slice(3)),
];
const keyedResult = engine.applyRules(keyedMatrix, keyedRules);
assert.equal(keyedResult.replacements, 2);
assert.equal(keyedResult.unmatchedRules, 0);
assert.deepEqual(new Set(keyedResult.changes.map((change) => change.language)), new Set(["values-sk", "values-cs"]));
assert.equal(keyedMatrix[1][4], "Iba na vývoj");
assert.equal(keyedMatrix[2][5], "shodné");
const repeatedResult = engine.applyRules(keyedMatrix, keyedRules);
assert.equal(repeatedResult.alreadyApplied, 2);
assert.equal(repeatedResult.changes.every((change) => change.alreadyApplied), true);

const duplicateResourceMatrix = [
  ["App", "string ID", "item ID", "values", "values-sk"],
  ["TunerInput", "shared", null, "ATV text", "staré"],
  ["TunerInput", "shared", null, "DTV text", "staré"],
];
const duplicateRules = engine.parseRules([
  ["App", "string ID", "item ID", "values", "values-sk (original)", "values-sk (corrected)"],
  ["TunerInput", "shared", null, "", "staré", "nové"],
], duplicateResourceMatrix[0].slice(3));
const duplicateResult = engine.applyRules(duplicateResourceMatrix, duplicateRules);
assert.equal(duplicateResult.replacements, 2);
assert.equal(duplicateResourceMatrix[1][4], "nové");
assert.equal(duplicateResourceMatrix[2][4], "nové");
assert.equal(duplicateResult.changes.every((change) => change.oldText === "staré" && change.newText === "nové"), true);

const unmatchedResult = engine.applyRules(duplicateResourceMatrix, [{
  kind: "keyed", language: "values-sk", app: "Missing", id: "missing", item: null,
  oldText: "old", newText: "new", sourceRow: 9,
}]);
assert.equal(unmatchedResult.unmatchedRules, 1);
assert.equal(unmatchedResult.unmatchedRuleDetails[0].sourceRow, 9);

console.log("correction engine tests passed");
