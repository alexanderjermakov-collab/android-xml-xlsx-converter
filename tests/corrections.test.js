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
assert.equal(matrix[3][4], "Nastavenia");
assert.equal(matrix[3][5], "Nastavení");
assert.equal(matrix[3][6], "İstanbul");
assert.equal(matrix[4][4], "malé");
assert.equal(engine.isUppercaseStart("…Hello"), true);
assert.equal(engine.uppercaseFirstLetter("(hello)", "en"), "(Hello)");

console.log("correction engine tests passed");
