(function exposeDiffEngine(root, factory) {
  const api = factory();
  if (typeof module === "object" && module.exports) module.exports = api;
  if (root) root.DiffEngine = api;
}(typeof globalThis !== "undefined" ? globalThis : this, () => {
  "use strict";

  function text(value) {
    return value === null || value === undefined ? "" : String(value);
  }

  function compare(sourceRecords, outputRecords) {
    const keys = new Set([...sourceRecords.keys(), ...outputRecords.keys()]);
    const differences = [];
    let matched = 0;
    let changed = 0;
    let onlySource = 0;
    let onlyOutput = 0;

    for (const key of keys) {
      const source = sourceRecords.get(key);
      const output = outputRecords.get(key);
      if (source && output && text(source.value) === text(output.value)) {
        matched += 1;
        continue;
      }

      let status;
      if (source && output) {
        status = "Changed";
        changed += 1;
      } else if (source) {
        status = "Only in source";
        onlySource += 1;
      } else {
        status = "Only in output";
        onlyOutput += 1;
      }
      const reference = source || output;
      differences.push({
        status,
        app: reference.app,
        language: reference.language,
        xmlPath: reference.xmlPath,
        filename: reference.filename,
        resourceType: reference.resourceType,
        id: reference.id,
        item: reference.item,
        sourceValue: source ? text(source.value) : null,
        outputValue: output ? text(output.value) : null,
      });
    }

    differences.sort((left, right) => [
      left.app,
      left.language,
      left.xmlPath,
      left.id,
      text(left.item),
      left.status,
    ].join("\u001f").localeCompare([
      right.app,
      right.language,
      right.xmlPath,
      right.id,
      text(right.item),
      right.status,
    ].join("\u001f")));

    return {
      matched,
      changed,
      onlySource,
      onlyOutput,
      different: changed + onlySource + onlyOutput,
      total: keys.size,
      differences,
    };
  }

  return { compare };
}));
