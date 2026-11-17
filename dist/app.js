/* global JSZip, XLSX, CorrectionEngine, DiffEngine */
"use strict";

const ROUNDTRIP_SHEET = "_RoundTrip";
const ROUNDTRIP_MARKER = "ANDROID_XML_XLSX_CONVERTER_V1";
const RESOURCE_MAP_SHEET = "_ResourceMap";
const RESOURCE_MAP_MARKER = "ANDROID_XML_RESOURCE_MAP_V2";
const SHEET_NAME = "All translations";
const CHUNK_SIZE = 30000;
const textEncoder = new TextEncoder();
const textDecoder = new TextDecoder("utf-8");
const schemaPromise = fetch("reference-schema.json").then((response) => {
  if (!response.ok) throw new Error("Could not load the reference workbook schema.");
  return response.json();
});

const ui = {
  zipInput: document.querySelector("#zipInput"),
  xlsxInput: document.querySelector("#xlsxInput"),
  referenceInput: document.querySelector("#referenceInput"),
  zipButton: document.querySelector("#zipButton"),
  xlsxButton: document.querySelector("#xlsxButton"),
  zipStatus: document.querySelector("#zipStatus"),
  xlsxStatus: document.querySelector("#xlsxStatus"),
  correctionInput: document.querySelector("#correctionInput"),
  correctionSelection: document.querySelector("#correctionSelection"),
  correctionSelectionCount: document.querySelector("#correctionSelectionCount"),
  correctionFileList: document.querySelector("#correctionFileList"),
  clearCorrectionFiles: document.querySelector("#clearCorrectionFiles"),
  correctionName: document.querySelector("#correctionName"),
  correctionTemplateButton: document.querySelector("#correctionTemplateButton"),
  capitalizationRule: document.querySelector("#capitalizationRule"),
  correctedXlsxButton: document.querySelector("#correctedXlsxButton"),
  compareSourceInput: document.querySelector("#compareSourceInput"),
  compareOutputInput: document.querySelector("#compareOutputInput"),
  compareButton: document.querySelector("#compareButton"),
  compareStatus: document.querySelector("#compareStatus"),
  xlsxName: document.querySelector("#xlsxName"),
  compareSourceName: document.querySelector("#compareSourceName"),
  compareOutputName: document.querySelector("#compareOutputName"),
  languagePanel: document.querySelector("#languagePanel"),
  languageList: document.querySelector("#languageList"),
  languageCount: document.querySelector("#languageCount"),
  languageSearch: document.querySelector("#languageSearch"),
  selectAllLanguages: document.querySelector("#selectAllLanguages"),
  clearLanguages: document.querySelector("#clearLanguages"),
};

let languageState = { fileToken: "", columns: [], ready: false };
const pipeline = {
  activeWorkbook: null,
  activeModifications: new Map(),
  sourceZip: null,
  outputZip: null,
  outputModifications: new Map(),
};
const correctionFileState = { entries: [] };

function showStatus(node, message, kind = "working") {
  node.textContent = message;
  node.className = `status visible${kind === "working" ? "" : ` ${kind}`}`;
}

function setBusy(button, busy) {
  button.disabled = busy;
  button.dataset.label ||= button.textContent;
  button.textContent = busy ? "Working…" : button.dataset.label;
}

function wireFileInput(input, nameNodeId, dropNodeId) {
  const nameNode = document.querySelector(`#${nameNodeId}`);
  const dropNode = document.querySelector(`#${dropNodeId}`);
  input.addEventListener("change", () => {
    const files = Array.from(input.files || []);
    nameNode.textContent = files.length > 1 ? `${files.length} files: ${files.map((file) => file.name).join(", ")}` : (files[0]?.name || "");
  });
  ["dragenter", "dragover"].forEach((eventName) => dropNode.addEventListener(eventName, () => dropNode.classList.add("dragging")));
  ["dragleave", "drop"].forEach((eventName) => dropNode.addEventListener(eventName, () => dropNode.classList.remove("dragging")));
}

wireFileInput(ui.zipInput, "zipName", "zipDrop");
wireFileInput(ui.xlsxInput, "xlsxName", "xlsxDrop");
wireFileInput(ui.referenceInput, "referenceName", "referenceDrop");
wireFileInput(ui.correctionInput, "correctionName", "correctionDrop");
wireFileInput(ui.compareSourceInput, "compareSourceName", "compareSourceDrop");
wireFileInput(ui.compareOutputInput, "compareOutputName", "compareOutputDrop");

function correctionFileKey(file) {
  return `${file.name}\u001f${file.size}\u001f${file.lastModified}`;
}

async function inspectCorrectionFile(file) {
  const bytes = new Uint8Array(await file.arrayBuffer());
  const workbook = XLSX.read(bytes, { type: "array", cellFormula: false, cellStyles: false });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) throw new Error("No worksheet found");
  const matrix = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, raw: true, defval: null });
  const headers = (matrix[0] || []).map((value) => String(value || ""));
  const languages = headers.flatMap((header) => {
    const match = /^\s*(values(?:-[^\s(]+)?)\s*\(original\)\s*$/i.exec(header);
    return match ? [match[1]] : [];
  });
  const legacy = headers.some((header) => header.trim().toLocaleLowerCase() === "language");
  if (!languages.length && !legacy) throw new Error("Correction language columns not found");
  const rows = matrix.slice(1).filter((row) => (row || []).some((value) => value !== null && value !== undefined && value !== "")).length;
  return {
    label: languages.length ? languages.join(", ") : "languages specified per row",
    rows,
  };
}

function renderCorrectionFiles() {
  const entries = correctionFileState.entries;
  ui.correctionSelection.hidden = entries.length === 0;
  ui.correctionSelectionCount.textContent = `${entries.length.toLocaleString()} correction ${entries.length === 1 ? "file" : "files"} loaded`;
  ui.correctionName.textContent = entries.length ? `${entries.length} loaded; add more files at any time` : "";
  ui.correctionFileList.replaceChildren();

  for (const entry of entries) {
    const item = document.createElement("div");
    item.className = "correction-file-item";
    const details = document.createElement("div");
    const name = document.createElement("strong");
    name.textContent = entry.file.name;
    const meta = document.createElement("span");
    if (entry.error) {
      meta.className = "error";
      meta.textContent = entry.error;
    } else if (entry.info) {
      meta.textContent = `${entry.info.label} · ${entry.info.rows.toLocaleString()} correction rows`;
    } else {
      meta.textContent = "Reading workbook…";
    }
    details.append(name, meta);
    const remove = document.createElement("button");
    remove.type = "button";
    remove.className = "text-button correction-remove";
    remove.textContent = "Remove";
    remove.addEventListener("click", () => {
      correctionFileState.entries = correctionFileState.entries.filter((candidate) => candidate.key !== entry.key);
      renderCorrectionFiles();
    });
    item.append(details, remove);
    ui.correctionFileList.append(item);
  }
}

async function addCorrectionFiles(files) {
  const known = new Set(correctionFileState.entries.map((entry) => entry.key));
  const additions = [];
  for (const file of files) {
    const key = correctionFileKey(file);
    if (known.has(key)) continue;
    known.add(key);
    const entry = { key, file, info: null, error: "" };
    correctionFileState.entries.push(entry);
    additions.push(entry);
  }
  renderCorrectionFiles();
  await Promise.all(additions.map(async (entry) => {
    try {
      entry.info = await inspectCorrectionFile(entry.file);
    } catch (error) {
      entry.error = error.message || String(error);
    }
  }));
  renderCorrectionFiles();
}

function selectedCorrectionFiles() {
  return correctionFileState.entries.map((entry) => entry.file);
}

ui.correctionInput.addEventListener("change", async () => {
  const files = Array.from(ui.correctionInput.files || []);
  ui.correctionInput.value = "";
  await addCorrectionFiles(files);
});

ui.clearCorrectionFiles.addEventListener("click", () => {
  correctionFileState.entries = [];
  renderCorrectionFiles();
});

function browserFile(bytes, filename, mime) {
  return new File([bytes], filename, { type: mime, lastModified: Date.now() });
}

function showDefaultFile(node, input, file, description) {
  if (!input.files.length && file) node.textContent = `${description}: ${file.name}`;
}

function setActiveWorkbook(file, description, modifications = new Map()) {
  pipeline.activeWorkbook = file;
  pipeline.activeModifications = new Map(modifications);
  ui.xlsxName.textContent = `${description}: ${file.name}`;
}

function activeWorkbook() {
  return pipeline.activeWorkbook || ui.xlsxInput.files[0] || null;
}

function setSourceZip(file) {
  if (pipeline.sourceZip && pipeline.sourceZip !== file) {
    pipeline.outputZip = null;
    pipeline.outputModifications = new Map();
    if (!ui.compareOutputInput.files.length) ui.compareOutputName.textContent = "";
  }
  pipeline.sourceZip = file;
  showDefaultFile(ui.compareSourceName, ui.compareSourceInput, file, "Default source");
}

function setOutputZip(file, modifications) {
  pipeline.outputZip = file;
  pipeline.outputModifications = new Map(modifications);
  showDefaultFile(ui.compareOutputName, ui.compareOutputInput, file, "Default output");
}

ui.xlsxInput.addEventListener("change", () => {
  const file = ui.xlsxInput.files[0];
  if (file) setActiveWorkbook(file, "Selected workbook");
});

function downloadBytes(bytes, filename, mime) {
  const blob = new Blob([bytes], { type: mime });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

function outputName(inputName, extension) {
  const stem = inputName.replace(/\.[^.]+$/, "");
  return `${stem}_converted.${extension}`;
}

function namedOutput(inputName, suffix, extension) {
  const stem = inputName.replace(/\.[^.]+$/, "");
  return `${stem}_${suffix}.${extension}`;
}

async function sha256(bytes) {
  const digest = await crypto.subtle.digest("SHA-256", bytes);
  return Array.from(new Uint8Array(digest), (value) => value.toString(16).padStart(2, "0")).join("");
}

function bytesToBase64(bytes) {
  const blocks = [];
  for (let offset = 0; offset < bytes.length; offset += 0x8000) {
    blocks.push(String.fromCharCode(...bytes.subarray(offset, offset + 0x8000)));
  }
  return btoa(blocks.join(""));
}

function base64ToBytes(value) {
  const binary = atob(value);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
  return bytes;
}

function canonicalItem(value) {
  if (value === null || value === undefined || value === "") return "n:";
  if (typeof value === "number") return `d:${value}`;
  return `s:${String(value)}`;
}

function rowKey(app, id, item) {
  return `${app}\u001f${id}\u001f${canonicalItem(item)}`;
}

function resourcePathTemplate(path, valuesDir) {
  const needle = `/${valuesDir}/`;
  const position = path.lastIndexOf(needle);
  if (position < 0) return path;
  return `${path.slice(0, position)}/{values}/${path.slice(position + needle.length)}`;
}

function resourceIdentityKey(templatePath, resourceType, id, item) {
  return JSON.stringify([templatePath, resourceType, String(id), item === null || item === undefined ? null : String(item)]);
}

function resourcePathForColumn(templatePath, column, variant) {
  return templatePath.replace("/{values}/", `/${targetValuesDir(column, variant)}/`);
}

async function matrixSignature(matrix) {
  const parts = [];
  for (let row = 0; row < matrix.length; row += 1) {
    const values = matrix[row] || [];
    for (let column = 0; column < values.length; column += 1) {
      const value = values[column];
      if (value !== null && value !== undefined && value !== "") {
        parts.push(`${row}:${column}:${canonicalItem(value)}\u001e`);
      }
    }
  }
  return sha256(textEncoder.encode(parts.join("")));
}

function splitValuesDirectory(valuesDir) {
  const variant = valuesDir.endsWith("-v23") ? "-v23" : null;
  const column = variant ? valuesDir.slice(0, -4) : valuesDir;
  const isLocale = /^values(?:-(?:[A-Za-z]{2,3}(?:-r[A-Za-z]{2,3})?|b\+[A-Za-z0-9+]+))?$/.test(column);
  return isLocale ? { column, variant } : null;
}

function parseResourcePath(path) {
  const match = /^APK\/([^/]+)\/(?:.*\/)?(values(?:-[^/]+)?)\/([^/]+\.xml)$/i.exec(path);
  if (!match) return null;
  const directory = splitValuesDirectory(match[2]);
  if (!directory) return null;
  return { app: match[1], valuesDir: match[2], filename: match[3], ...directory };
}

function fileToken(file) {
  return file ? `${file.name}:${file.size}:${file.lastModified}` : "";
}

function updateLanguageCount() {
  const options = Array.from(ui.languageList.querySelectorAll('input[type="checkbox"]:not(:disabled)'));
  const selected = options.filter((option) => option.checked).length;
  ui.languageCount.textContent = `${selected} of ${options.length} languages selected`;
}

function renderLanguages(columns) {
  ui.languageList.replaceChildren();
  const baseLabel = document.createElement("label");
  baseLabel.className = "language-option base";
  const baseCheckbox = document.createElement("input");
  baseCheckbox.type = "checkbox";
  baseCheckbox.checked = true;
  baseCheckbox.disabled = true;
  const baseText = document.createElement("span");
  baseText.textContent = "values — base/source language (always included)";
  baseLabel.append(baseCheckbox, baseText);
  ui.languageList.append(baseLabel);

  for (const column of columns.filter((value) => value !== "values")) {
    const label = document.createElement("label");
    label.className = "language-option";
    label.dataset.search = column.toLocaleLowerCase();
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.value = column;
    checkbox.checked = true;
    checkbox.addEventListener("change", updateLanguageCount);
    const text = document.createElement("span");
    text.textContent = column;
    label.append(checkbox, text);
    ui.languageList.append(label);
  }
  ui.languagePanel.hidden = false;
  ui.languageSearch.value = "";
  updateLanguageCount();
}

async function scanArchiveLanguages(file) {
  const token = fileToken(file);
  languageState = { fileToken: token, columns: [], ready: false };
  ui.languagePanel.hidden = false;
  ui.languageList.replaceChildren();
  ui.languageCount.textContent = "Scanning archive…";
  const zip = await JSZip.loadAsync(new Uint8Array(await file.arrayBuffer()));
  const found = new Set();
  for (const entry of Object.values(zip.files)) {
    if (entry.dir) continue;
    const info = parseResourcePath(entry.name);
    if (info) found.add(info.column);
  }
  const schema = await schemaPromise;
  const ordered = schema.headers.filter((header) => found.has(header));
  for (const column of found) if (!ordered.includes(column)) ordered.push(column);
  languageState = { fileToken: token, columns: ordered, ready: true };
  renderLanguages(ordered);
}

function selectedLanguageColumns() {
  const selected = new Set(["values"]);
  for (const checkbox of ui.languageList.querySelectorAll('input[type="checkbox"]:checked:not(:disabled)')) {
    selected.add(checkbox.value);
  }
  return selected;
}

ui.zipInput.addEventListener("change", async () => {
  const file = ui.zipInput.files[0];
  if (!file) {
    ui.languagePanel.hidden = true;
    languageState = { fileToken: "", columns: [], ready: false };
    return;
  }
  setSourceZip(file);
  ui.zipButton.disabled = true;
  showStatus(ui.zipStatus, "Scanning available languages…");
  try {
    await scanArchiveLanguages(file);
    showStatus(ui.zipStatus, `Found ${Math.max(0, languageState.columns.length - 1)} selectable languages.`, "success");
  } catch (error) {
    languageState = { fileToken: fileToken(file), columns: [], ready: false };
    showStatus(ui.zipStatus, error.message || String(error), "error");
  } finally {
    ui.zipButton.disabled = false;
  }
});

ui.selectAllLanguages.addEventListener("click", () => {
  ui.languageList.querySelectorAll('input[type="checkbox"]:not(:disabled)').forEach((checkbox) => { checkbox.checked = true; });
  updateLanguageCount();
});

ui.clearLanguages.addEventListener("click", () => {
  ui.languageList.querySelectorAll('input[type="checkbox"]:not(:disabled)').forEach((checkbox) => { checkbox.checked = false; });
  updateLanguageCount();
});

ui.languageSearch.addEventListener("input", () => {
  const query = ui.languageSearch.value.trim().toLocaleLowerCase();
  ui.languageList.querySelectorAll(".language-option:not(.base)").forEach((option) => {
    option.classList.toggle("filtered", Boolean(query) && !option.dataset.search.includes(query));
  });
});

function xmlDocument(text, path) {
  const documentNode = new DOMParser().parseFromString(text, "application/xml");
  const error = documentNode.querySelector("parsererror");
  if (error) throw new Error(`Invalid XML in ${path}: ${error.textContent.split("\n")[0]}`);
  return documentNode;
}

function directChildren(element, localName) {
  return Array.from(element.children).filter((child) => child.localName === localName);
}

function groupKey(templatePath, resourceType, id, variant) {
  return resourceIdentityKey(templatePath, resourceType, id, variant || null);
}

async function extractArchive(zipBytes, selectedColumns, onProgress) {
  const zip = await JSZip.loadAsync(zipBytes);
  const entries = Object.values(zip.files).filter((entry) => {
    if (entry.dir) return false;
    const info = parseResourcePath(entry.name);
    return info && selectedColumns.has(info.column);
  });
  const groups = new Map();
  const discoveredColumns = [];
  const columnSet = new Set();

  for (let entryIndex = 0; entryIndex < entries.length; entryIndex += 1) {
    const entry = entries[entryIndex];
    const info = parseResourcePath(entry.name);
    if (!columnSet.has(info.column)) {
      columnSet.add(info.column);
      discoveredColumns.push(info.column);
    }
    const documentNode = xmlDocument(await entry.async("string"), entry.name);
    const resources = documentNode.documentElement;
    for (const element of Array.from(resources.children)) {
      const id = element.getAttribute("name");
      if (!id) continue;
      const isArrayFile = info.filename === "arrays.xml" || info.filename === "strings.xml";
      if (element.localName === "string-array" && isArrayFile) {
        const templatePath = resourcePathTemplate(entry.name, info.valuesDir);
        const key = groupKey(templatePath, "string-array", id, info.variant);
        let group = groups.get(key);
        if (!group) {
          group = { app: info.app, id, variant: info.variant, templatePath, resourceType: "string-array", isArray: true, items: [] };
          groups.set(key, group);
        }
        const items = directChildren(element, "item");
        items.forEach((item, index) => {
          group.items[index] ||= new Map();
          group.items[index].set(info.column, item.textContent || null);
        });
      } else if (
        element.localName === "string" &&
        (info.filename.startsWith("strings") || info.filename.endsWith("strings.xml"))
      ) {
        const templatePath = resourcePathTemplate(entry.name, info.valuesDir);
        const key = groupKey(templatePath, "string", id, info.variant);
        let group = groups.get(key);
        if (!group) {
          group = { app: info.app, id, variant: info.variant, templatePath, resourceType: "string", isArray: false, items: [] };
          groups.set(key, group);
        }
        group.items[0] ||= new Map();
        group.items[0].set(info.column, element.textContent || null);
      }
    }
    if (entryIndex % 25 === 0) {
      onProgress?.(`Reading XML resources… ${entryIndex + 1} / ${entries.length}`);
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
  }
  return { groups, discoveredColumns, processedFiles: entries.length };
}

function archiveRecordKey(xmlPath, resourceType, id, item) {
  return JSON.stringify([xmlPath, resourceType, id, item === null || item === undefined ? null : String(item)]);
}

async function archiveRecords(file, label, onProgress) {
  const zip = await JSZip.loadAsync(new Uint8Array(await file.arrayBuffer()));
  const entries = Object.values(zip.files).filter((entry) => !entry.dir && parseResourcePath(entry.name));
  const records = new Map();

  for (let entryIndex = 0; entryIndex < entries.length; entryIndex += 1) {
    const entry = entries[entryIndex];
    const info = parseResourcePath(entry.name);
    const documentNode = xmlDocument(await entry.async("string"), entry.name);
    const resources = documentNode.documentElement;
    for (const element of Array.from(resources.children)) {
      const id = element.getAttribute("name");
      if (!id) continue;
      const isArrayFile = info.filename === "arrays.xml" || info.filename === "strings.xml";
      if (element.localName === "string-array" && isArrayFile) {
        directChildren(element, "item").forEach((item, itemIndex) => {
          const itemValue = info.variant ? `${itemIndex}${info.variant}` : itemIndex;
          const key = archiveRecordKey(entry.name, "string-array", id, itemValue);
          records.set(key, {
            app: info.app,
            language: info.column,
            xmlPath: entry.name,
            filename: info.filename,
            resourceType: "string-array",
            id,
            item: itemValue,
            value: item.textContent || "",
          });
        });
      } else if (
        element.localName === "string" &&
        (info.filename.startsWith("strings") || info.filename.endsWith("strings.xml"))
      ) {
        const itemValue = info.variant || null;
        const key = archiveRecordKey(entry.name, "string", id, itemValue);
        records.set(key, {
          app: info.app,
          language: info.column,
          xmlPath: entry.name,
          filename: info.filename,
          resourceType: "string",
          id,
          item: itemValue,
          value: element.textContent || "",
        });
      }
    }
    if (entryIndex % 25 === 0) {
      onProgress?.(`Reading ${label} XML files… ${entryIndex + 1} / ${entries.length}`);
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
  }
  return { records, xmlFiles: entries.length };
}

function auditEntryForRow(row, modifications) {
  if (row.auditKey && modifications.has(row.auditKey)) return modifications.get(row.auditKey);
  const exactKey = archiveRecordKey(row.xmlPath, row.resourceType, row.id, row.item);
  return modifications.get(exactKey) || modifications.get(legacyModificationKey(row.app, row.id, row.item, row.language)) || null;
}

function differenceConclusion(row, modifications) {
  if (row.status === "Only in source") return "Removed from output ZIP";
  if (row.status === "Only in output") return "Added to output ZIP";
  const audit = auditEntryForRow(row, modifications);
  return audit?.reasons?.length ? audit.reasons.join("; ") : "Modified value (reason not recorded in this browser session)";
}

function correctionEvaluation(audit, outputRecord) {
  const checks = audit?.checks || [];
  const expectedValues = new Set(checks.map((check) => String(check.expectedValue ?? "")));
  const conflict = expectedValues.size > 1;
  const results = checks.map((check) => {
    let status;
    if (check.targetMissing) status = "TARGET NOT FOUND";
    else if (!outputRecord) status = "MISSING OUTPUT";
    else if (String(outputRecord.value ?? "") === String(check.expectedValue ?? "")) status = "MATCH";
    else if (String(outputRecord.value ?? "").replace(/\r\n?/g, "\n") === String(check.expectedValue ?? "").replace(/\r\n?/g, "\n")) {
      status = "MATCH (XML line endings normalized)";
    }
    else status = "MISMATCH";
    return { ...check, status };
  });
  let status = "NOT REQUESTED";
  if (checks.length) {
    if (conflict) status = "CONFLICT";
    else if (results.some((check) => check.status === "TARGET NOT FOUND")) status = "TARGET NOT FOUND";
    else if (results.every((check) => check.status.startsWith("MATCH"))) status = "MATCH";
    else if (results.some((check) => check.status === "MISSING OUTPUT")) status = "MISSING OUTPUT";
    else status = "MISMATCH";
  }
  return { checks, results, conflict, status };
}

function perLanguageStatistics(sourceArchive, outputArchive, modifications) {
  const stats = new Map();
  const ensure = (language) => {
    if (!stats.has(language)) stats.set(language, {
      language, total: 0, matched: 0, changed: 0, onlySource: 0, onlyOutput: 0,
      correctionTargets: 0, correctionChecks: 0, verified: 0, mismatched: 0,
      missing: 0, unmatched: 0, conflicts: 0, capitalization: 0,
    });
    return stats.get(language);
  };
  const keys = new Set([...sourceArchive.records.keys(), ...outputArchive.records.keys()]);
  for (const key of keys) {
    const source = sourceArchive.records.get(key);
    const output = outputArchive.records.get(key);
    const row = ensure((source || output).language);
    row.total += 1;
    if (source && output && String(source.value ?? "") === String(output.value ?? "")) row.matched += 1;
    else if (source && output) row.changed += 1;
    else if (source) row.onlySource += 1;
    else row.onlyOutput += 1;
  }
  for (const [key, audit] of modifications) {
    const row = ensure(audit.language || "unknown");
    const evaluation = correctionEvaluation(audit, outputArchive.records.get(key));
    if (evaluation.checks.length) {
      if (!evaluation.checks.every((check) => check.targetMissing)) row.correctionTargets += 1;
      row.correctionChecks += evaluation.checks.length;
      row.verified += evaluation.results.filter((check) => check.status.startsWith("MATCH")).length;
      row.mismatched += evaluation.results.filter((check) => check.status === "MISMATCH").length;
      row.missing += evaluation.results.filter((check) => check.status === "MISSING OUTPUT").length;
      row.unmatched += evaluation.results.filter((check) => check.status === "TARGET NOT FOUND").length;
      if (evaluation.conflict) row.conflicts += 1;
    }
    if (audit.reasons?.includes("Capitalization")) row.capitalization += 1;
  }
  return [...stats.values()].sort((left, right) => left.language.localeCompare(right.language));
}

function diffWorkbook(sourceFile, outputFile, sourceArchive, outputArchive, result, modifications) {
  const workbook = XLSX.utils.book_new();
  const matchRate = result.total ? result.matched / result.total : 1;
  const languageStats = perLanguageStatistics(sourceArchive, outputArchive, modifications);
  const summaryRows = [
    ["ZIP translation comparison", null],
    ["Source ZIP", sourceFile.name],
    ["Output ZIP", outputFile.name],
    ["Generated", new Date().toLocaleString()],
    [null, null],
    ["Metric", "Count"],
    ["Total translation rows", result.total],
    ["Matching rows", result.matched],
    ["Different rows", result.different],
    ["Changed values", result.changed],
    ["Only in source", result.onlySource],
    ["Only in output", result.onlyOutput],
    ["Match rate", matchRate],
    ["Source translation rows", sourceArchive.records.size],
    ["Output translation rows", outputArchive.records.size],
    ["Source translation XML files", sourceArchive.xmlFiles],
    ["Output translation XML files", outputArchive.xmlFiles],
    [null, null],
    ["Statistics by language"],
    [
      "Language", "Total rows", "Matching", "Different", "Changed", "Only in source", "Only in output", "Match rate",
      "Correction targets", "Correction checks", "Verified checks", "Mismatched checks", "Missing checks", "Unmatched correction rules", "Conflicts", "Capitalization changes",
    ],
    ...languageStats.map((row) => {
      const different = row.changed + row.onlySource + row.onlyOutput;
      return [
        row.language, row.total, row.matched, different, row.changed, row.onlySource, row.onlyOutput,
        row.total ? row.matched / row.total : 1,
        row.correctionTargets, row.correctionChecks, row.verified, row.mismatched, row.missing, row.unmatched, row.conflicts, row.capitalization,
      ];
    }),
  ];
  const summarySheet = XLSX.utils.aoa_to_sheet(summaryRows);
  summarySheet["!cols"] = [
    { wch: 34 }, { wch: 18 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 16 }, { wch: 16 }, { wch: 14 },
    { wch: 18 }, { wch: 18 }, { wch: 17 }, { wch: 19 }, { wch: 16 }, { wch: 26 }, { wch: 12 }, { wch: 22 },
  ];
  if (summarySheet.B13) summarySheet.B13.z = "0.0%";
  const languageHeaderIndex = 19;
  for (let row = languageHeaderIndex + 1; row < summaryRows.length; row += 1) {
    const address = XLSX.utils.encode_cell({ r: row, c: 7 });
    if (summarySheet[address]) summarySheet[address].z = "0.0%";
  }
  summarySheet["!autofilter"] = { ref: `A${languageHeaderIndex + 1}:P${Math.max(languageHeaderIndex + 1, summaryRows.length)}` };
  XLSX.utils.book_append_sheet(workbook, summarySheet, "Summary");

  const correctionLanguages = [...new Set([...modifications.values()].flatMap((entry) =>
    (entry.checks || []).map((check) => check.language)
  ))].sort((left, right) => left.localeCompare(right));
  const differenceRows = [[
    "Status",
    "App",
    "Language",
    "XML path",
    "XML file",
    "Resource type",
    "string ID",
    "item ID",
    "Source value",
    "Output value",
    "Conclusion",
    ...correctionLanguages.flatMap((language) => [
      `${language} correction source`,
      `${language} correction original value`,
      `${language} expected corrected value`,
      `${language} verification`,
    ]),
  ]];
  const reportRows = new Map();
  result.differences.forEach((row) => {
    reportRows.set(archiveRecordKey(row.xmlPath, row.resourceType, row.id, row.item), row);
  });
  for (const [key, audit] of modifications) {
    if (!(audit.checks || []).length || reportRows.has(key)) continue;
    const source = sourceArchive.records.get(key);
    const output = outputArchive.records.get(key);
    const evaluation = correctionEvaluation(audit, output);
    reportRows.set(key, {
      status: evaluation.status === "MATCH" ? "Correction verified" :
        (evaluation.status === "TARGET NOT FOUND" ? "Correction target not found" : "Correction check failed"),
      auditKey: key,
      app: audit.app,
      language: audit.language,
      xmlPath: audit.xmlPath,
      filename: audit.filename,
      resourceType: audit.resourceType,
      id: audit.id,
      item: audit.item,
      sourceValue: source ? String(source.value ?? "") : null,
      outputValue: output ? String(output.value ?? "") : null,
    });
  }
  [...reportRows.values()].sort((left, right) => [
    left.app, left.language, left.xmlPath, left.id, String(left.item ?? ""), left.status,
  ].join("\u001f").localeCompare([
    right.app, right.language, right.xmlPath, right.id, String(right.item ?? ""), right.status,
  ].join("\u001f"))).forEach((row) => {
    row.conclusion = differenceConclusion(row, modifications);
    const audit = auditEntryForRow(row, modifications);
    const outputRecord = row.auditKey ? outputArchive.records.get(row.auditKey) :
      outputArchive.records.get(archiveRecordKey(row.xmlPath, row.resourceType, row.id, row.item));
    const evaluation = correctionEvaluation(audit, outputRecord);
    differenceRows.push([
      row.status,
      row.app,
      row.language,
      row.xmlPath,
      row.filename,
      row.resourceType,
      row.id,
      row.item,
      row.sourceValue,
      row.outputValue,
      row.conclusion,
      ...correctionLanguages.flatMap((language) => {
        if (!audit || audit.language !== language || !evaluation.checks.length) return [null, null, null, null];
        return [
          evaluation.checks.map((check) => `${check.file} (row ${check.sourceRow})`).join("\n"),
          evaluation.checks.map((check) => `${check.file}: ${check.originalValue}`).join("\n"),
          evaluation.checks.map((check) => `${check.file}: ${check.expectedValue}`).join("\n"),
          evaluation.results.map((check) => `${check.file}: ${check.status}`).join("\n") + (evaluation.conflict ? "\nConflicting expected values" : ""),
        ];
      }),
    ]);
  });
  const differencesSheet = XLSX.utils.aoa_to_sheet(differenceRows);
  differencesSheet["!cols"] = [
    { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 64 }, { wch: 20 },
    { wch: 16 }, { wch: 40 }, { wch: 12 }, { wch: 64 }, { wch: 64 }, { wch: 56 },
    ...correctionLanguages.flatMap(() => [{ wch: 38 }, { wch: 54 }, { wch: 54 }, { wch: 32 }]),
  ];
  const lastColumn = XLSX.utils.encode_col(differenceRows[0].length - 1);
  differencesSheet["!autofilter"] = { ref: `A1:${lastColumn}${Math.max(1, differenceRows.length)}` };
  XLSX.utils.book_append_sheet(workbook, differencesSheet, "Differences");
  return workbook;
}

async function compareArchives(sourceFile, outputFile, modifications = new Map()) {
  const sourceArchive = await archiveRecords(sourceFile, "source", (message) => showStatus(ui.compareStatus, message));
  const outputArchive = await archiveRecords(outputFile, "output", (message) => showStatus(ui.compareStatus, message));
  const result = DiffEngine.compare(sourceArchive.records, outputArchive.records);
  const workbook = diffWorkbook(sourceFile, outputFile, sourceArchive, outputArchive, result, modifications);
  const workbookBytes = new Uint8Array(XLSX.write(workbook, { bookType: "xlsx", type: "array", compression: true }));
  const sourceStem = sourceFile.name.replace(/\.[^.]+$/, "");
  const outputStem = outputFile.name.replace(/\.[^.]+$/, "");
  downloadBytes(
    workbookBytes,
    `${sourceStem}_vs_${outputStem}_differences.xlsx`,
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  );
  return result;
}

function flattenGroups(groups) {
  const rows = [];
  for (const group of groups.values()) {
    group.items.forEach((values, index) => {
      if (!values) return;
      let item = null;
      if (group.variant) item = group.isArray ? `${index}${group.variant}` : group.variant;
      else if (group.isArray) item = index;
      rows.push({
        app: group.app,
        id: group.id,
        item,
        values,
        templatePath: group.templatePath,
        resourceType: group.resourceType,
        variant: group.variant,
      });
    });
  }
  return rows;
}

function orderRows(rows, schema) {
  const byKey = new Map();
  rows.forEach((row) => {
    const key = rowKey(row.app, row.id, row.item);
    if (!byKey.has(key)) byKey.set(key, []);
    byKey.get(key).push(row);
  });
  const ordered = [];
  for (const [app, id, item] of schema.rows) {
    const key = rowKey(app, id, item);
    const bucket = byKey.get(key);
    if (bucket?.length) {
      ordered.push(bucket.shift());
      if (!bucket.length) byKey.delete(key);
    }
  }
  for (const bucket of byKey.values()) ordered.push(...bucket);
  return ordered;
}

function buildMatrix(rows, headers) {
  const columnIndex = new Map(headers.map((header, index) => [header, index + 3]));
  const matrix = [["App", "string ID", "item ID", ...headers]];
  for (const row of rows) {
    const output = new Array(headers.length + 3).fill(null);
    output[0] = row.app;
    output[1] = row.id;
    output[2] = row.item;
    for (const [column, value] of row.values) {
      const index = columnIndex.get(column);
      if (index !== undefined && value !== null && value !== "") output[index] = value;
    }
    matrix.push(output);
  }
  return matrix;
}

function makeRoundTripSheet(filename, signature, zipBytes) {
  const payload = bytesToBase64(zipBytes);
  const chunks = [];
  for (let offset = 0; offset < payload.length; offset += CHUNK_SIZE) chunks.push(payload.slice(offset, offset + CHUNK_SIZE));
  return XLSX.utils.aoa_to_sheet([
    ["marker", ROUNDTRIP_MARKER],
    ["source_name", filename],
    ["sheet_signature", signature],
    ["chunk_count", chunks.length],
    ...chunks.map((chunk, index) => [`chunk_${String(index).padStart(6, "0")}`, chunk]),
  ]);
}

function makeResourceMapSheet(rows) {
  return XLSX.utils.aoa_to_sheet([
    ["marker", RESOURCE_MAP_MARKER],
    ["worksheet row", "App", "string ID", "item ID", "XML path template", "Resource type", "Variant"],
    ...rows.map((row, index) => [
      index + 2,
      row.app,
      row.id,
      row.item,
      row.templatePath,
      row.resourceType,
      row.variant || null,
    ]),
  ]);
}

function readResourceMap(workbook, matrix) {
  const sheet = workbook.Sheets[RESOURCE_MAP_SHEET];
  if (!sheet) return null;
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null });
  if (rows[0]?.[1] !== RESOURCE_MAP_MARKER) return null;
  const result = new Map();
  for (let index = 2; index < rows.length; index += 1) {
    const source = rows[index] || [];
    const matrixIndex = Number(source[0]) - 1;
    const visible = matrix[matrixIndex] || [];
    if (!Number.isInteger(matrixIndex) || matrixIndex < 1) continue;
    if (String(visible[0] ?? "") !== String(source[1] ?? "") ||
        String(visible[1] ?? "") !== String(source[2] ?? "") ||
        canonicalItem(visible[2]) !== canonicalItem(source[3])) {
      throw new Error(`The hidden resource map no longer matches worksheet row ${matrixIndex + 1}. Regenerate the XLSX instead of sorting its rows.`);
    }
    result.set(matrixIndex, {
      app: String(source[1]),
      id: String(source[2]),
      item: visible[2],
      templatePath: String(source[4]),
      resourceType: String(source[5]),
      variant: source[6] ? String(source[6]) : null,
    });
  }
  return result;
}

function translationSheetName(workbook) {
  return workbook.SheetNames.includes(SHEET_NAME) ? SHEET_NAME : workbook.SheetNames[0];
}

function workbookMatrix(workbook) {
  const sheetName = translationSheetName(workbook);
  if (!sheetName) throw new Error("The workbook contains no worksheets.");
  const matrix = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, raw: true, defval: null });
  const header = matrix[0] || [];
  if (header[0] !== "App" || header[1] !== "string ID" || header[2] !== "item ID") {
    throw new Error("Expected columns App, string ID, and item ID in the first row.");
  }
  return matrix;
}

async function correctionRules(file, translationHeaders) {
  const bytes = new Uint8Array(await file.arrayBuffer());
  const workbook = XLSX.read(bytes, { type: "array", cellFormula: false, cellStyles: false });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) throw new Error("The correction workbook contains no worksheets.");
  const matrix = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, raw: true, defval: null });
  return CorrectionEngine.parseRules(matrix, translationHeaders);
}

async function applyCorrectionOptions(matrix, correctionFiles, capitalizationEnabled) {
  const headers = (matrix[0] || []).slice(3).map(String);
  const changes = new Map();
  const addChange = (change, reason, correctionCheck = null) => {
    const key = `${change.row}:${change.column}`;
    const existing = changes.get(key);
    const reasons = new Set(existing?.reasons || []);
    reasons.add(reason);
    const correctionChecks = [...(existing?.correctionChecks || [])];
    if (correctionCheck && !correctionChecks.some((check) =>
      check.file === correctionCheck.file &&
      check.sourceRow === correctionCheck.sourceRow &&
      check.language === correctionCheck.language
    )) correctionChecks.push(correctionCheck);
    changes.set(key, { ...existing, ...change, reasons: [...reasons], correctionChecks });
  };
  let replacements = 0;
  let matchedRules = 0;
  let unmatchedRules = 0;
  let alreadyApplied = 0;
  let capitalized = 0;
  const unmatchedChecks = [];

  for (const correctionFile of correctionFiles) {
    let rules;
    try {
      rules = await correctionRules(correctionFile, headers);
    } catch (error) {
      throw new Error(`${correctionFile.name}: ${error.message || String(error)}`);
    }
    const result = CorrectionEngine.applyRules(matrix, rules);
    replacements += result.replacements;
    matchedRules += result.matchedRules;
    unmatchedRules += result.unmatchedRules;
    alreadyApplied += result.alreadyApplied || 0;
    result.changes.forEach((change) => {
      const language = change.language || headers[change.column - 3];
      addChange(
        change,
        `Correction file: ${correctionFile.name} (${language})`,
        {
          file: correctionFile.name,
          language,
          sourceRow: change.sourceRow,
          originalValue: change.oldText,
          expectedValue: change.newText ?? change.value,
        },
      );
    });
    for (const rule of result.unmatchedRuleDetails || []) {
      const languages = rule.kind === "keyed" ? [rule.language] : rule.languages;
      for (const language of languages) unmatchedChecks.push({
        file: correctionFile.name,
        language,
        sourceRow: rule.sourceRow,
        originalValue: rule.oldText,
        expectedValue: rule.newText,
        app: rule.app || "",
        id: rule.id || "",
        item: rule.item,
        targetMissing: true,
      });
    }
  }
  if (capitalizationEnabled) {
    const result = CorrectionEngine.applyCapitalization(matrix);
    capitalized = result.capitalized;
    result.changes.forEach((change) => addChange(change, "Capitalization"));
  }
  return {
    files: correctionFiles.length,
    replacements,
    matchedRules,
    unmatchedRules,
    alreadyApplied,
    capitalized,
    unmatchedChecks,
    changes: [...changes.values()],
  };
}

function legacyModificationKey(app, id, item, language) {
  const itemValue = item === null || item === undefined || item === "" ? "" : String(item);
  return `legacy:${String(app)}\u001f${String(id)}\u001f${itemValue}\u001f${String(language)}`;
}

function modificationAudit(matrix, changes, existing = new Map(), resourceMap = null, unmatchedChecks = []) {
  const audit = new Map(existing);
  const headers = (matrix[0] || []).slice(3).map(String);
  for (const change of changes) {
    const row = matrix[change.row] || [];
    const language = change.language || headers[change.column - 3];
    const identity = resourceMap?.get(change.row) || null;
    const xmlPath = identity ? resourcePathForColumn(identity.templatePath, language, identity.variant) : null;
    const key = identity
      ? archiveRecordKey(xmlPath, identity.resourceType, identity.id, identity.item)
      : legacyModificationKey(row[0], row[1], row[2], language);
    const previous = audit.get(key) || {};
    const reasons = new Set(previous.reasons || []);
    (change.reasons || []).forEach((reason) => reasons.add(reason));
    const checks = [...(previous.checks || [])];
    for (const check of change.correctionChecks || []) {
      if (!checks.some((current) => current.file === check.file && current.sourceRow === check.sourceRow && current.language === check.language)) {
        checks.push(check);
      }
    }
    audit.set(key, {
      key,
      app: String(row[0] ?? ""),
      id: String(row[1] ?? ""),
      item: row[2],
      language,
      xmlPath,
      filename: xmlPath ? xmlPath.split("/").pop() : "",
      resourceType: identity?.resourceType || "",
      reasons: [...reasons],
      checks,
    });
  }
  for (const check of unmatchedChecks) {
    const key = `unmatched:${check.file}\u001f${check.sourceRow}\u001f${check.language}`;
    audit.set(key, {
      key,
      app: check.app,
      id: check.id,
      item: check.item,
      language: check.language,
      xmlPath: "",
      filename: "",
      resourceType: "",
      reasons: [`Correction rule did not match any workbook row: ${check.file} (${check.language})`],
      checks: [check],
    });
  }
  return audit;
}

function writeWorkbookChanges(workbook, changes) {
  const sheet = workbook.Sheets[translationSheetName(workbook)];
  for (const change of changes) {
    const address = XLSX.utils.encode_cell({ r: change.row, c: change.column });
    const cell = sheet[address] || {};
    cell.t = "s";
    cell.v = String(change.value);
    delete cell.f;
    delete cell.w;
    delete cell.h;
    sheet[address] = cell;
  }
}

function correctionSummary(result) {
  const parts = [];
  if (result.files) parts.push(`${result.files.toLocaleString()} correction files`);
  if (result.replacements) parts.push(`${result.replacements.toLocaleString()} text replacements`);
  if (result.alreadyApplied) parts.push(`${result.alreadyApplied.toLocaleString()} already applied`);
  if (result.capitalized) parts.push(`${result.capitalized.toLocaleString()} capitalization fixes`);
  if (result.unmatchedRules) parts.push(`${result.unmatchedRules.toLocaleString()} unmatched correction rules`);
  return parts.length ? parts.join(", ") : "no cell changes";
}

function downloadCorrectionTemplate() {
  const workbook = XLSX.utils.book_new();
  const corrections = XLSX.utils.aoa_to_sheet([
    ["App", "string ID", "item ID", "values", "values-xx (original)", "values-xx (corrected)"],
  ]);
  corrections["!cols"] = [{ wch: 10 }, { wch: 36 }, { wch: 8 }, { wch: 56 }, { wch: 28 }, { wch: 28 }];
  XLSX.utils.book_append_sheet(workbook, corrections, "List1");
  const bytes = new Uint8Array(XLSX.write(workbook, { bookType: "xlsx", type: "array", compression: true }));
  downloadBytes(bytes, "language_correction_template.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
}

async function createCorrectedWorkbook(xlsxFile, correctionFiles, capitalizationEnabled) {
  if (!correctionFiles.length && !capitalizationEnabled) throw new Error("Choose one or more correction workbooks or enable the capitalization rule.");
  const xlsxBytes = new Uint8Array(await xlsxFile.arrayBuffer());
  const workbook = XLSX.read(xlsxBytes, { type: "array", cellFormula: false, cellStyles: true });
  const matrix = workbookMatrix(workbook);
  const resourceMap = readResourceMap(workbook, matrix);
  const result = await applyCorrectionOptions(matrix, correctionFiles, capitalizationEnabled);
  const audit = modificationAudit(matrix, result.changes, pipeline.activeModifications, resourceMap, result.unmatchedChecks);
  writeWorkbookChanges(workbook, result.changes);
  const output = new Uint8Array(XLSX.write(workbook, { bookType: "xlsx", type: "array", compression: true }));
  const filename = namedOutput(xlsxFile.name, "corrected", "xlsx");
  downloadBytes(output, filename, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  const generatedFile = browserFile(output, filename, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  setActiveWorkbook(generatedFile, "Default corrected workbook", audit);
  return { message: `Corrected XLSX created and selected for ZIP conversion: ${correctionSummary(result)}.`, result };
}

function readRoundTrip(workbook) {
  const sheet = workbook.Sheets[ROUNDTRIP_SHEET];
  if (!sheet) return null;
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null });
  const values = new Map(rows.map((row) => [String(row[0]), row[1]]));
  if (values.get("marker") !== ROUNDTRIP_MARKER) return null;
  const count = Number(values.get("chunk_count"));
  const chunks = [];
  for (let index = 0; index < count; index += 1) chunks.push(values.get(`chunk_${String(index).padStart(6, "0")}`) || "");
  return {
    sourceName: String(values.get("source_name") || "translations.zip"),
    signature: String(values.get("sheet_signature") || ""),
    zipBytes: base64ToBytes(chunks.join("")),
  };
}

async function createWorkbook(zipFile, selectedColumns) {
  const zipBytes = new Uint8Array(await zipFile.arrayBuffer());
  const schema = await schemaPromise;
  const headers = schema.headers.filter((header) => selectedColumns.has(header));
  for (const column of languageState.columns) {
    if (selectedColumns.has(column) && !headers.includes(column)) headers.push(column);
  }
  const extracted = await extractArchive(zipBytes, selectedColumns, (message) => showStatus(ui.zipStatus, message));
  const rows = orderRows(flattenGroups(extracted.groups), schema);
  const headerSet = new Set(headers);
  for (const column of extracted.discoveredColumns) {
    if (!headerSet.has(column)) {
      headerSet.add(column);
      headers.push(column);
    }
  }
  showStatus(ui.zipStatus, `Building ${rows.length.toLocaleString()} rows across ${headers.length.toLocaleString()} translation columns…`);
  const matrix = buildMatrix(rows, headers);
  const signature = await matrixSignature(matrix);
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet(matrix);
  sheet["!ref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: matrix.length - 1, c: headers.length + 2 } });
  sheet["!autofilter"] = { ref: sheet["!ref"] };
  sheet["!cols"] = [{ wch: 9 }, { wch: 38 }, { wch: 12 }, ...headers.map(() => ({ wch: 28 }))];
  XLSX.utils.book_append_sheet(workbook, sheet, SHEET_NAME);
  XLSX.utils.book_append_sheet(workbook, makeRoundTripSheet(zipFile.name, signature, zipBytes), ROUNDTRIP_SHEET);
  XLSX.utils.book_append_sheet(workbook, makeResourceMapSheet(rows), RESOURCE_MAP_SHEET);
  workbook.Workbook = { Sheets: [
    { name: SHEET_NAME, Hidden: 0 },
    { name: ROUNDTRIP_SHEET, Hidden: 2 },
    { name: RESOURCE_MAP_SHEET, Hidden: 2 },
  ] };
  const workbookBytes = new Uint8Array(XLSX.write(workbook, { bookType: "xlsx", type: "array", compression: true }));
  const filename = outputName(zipFile.name, "xlsx");
  downloadBytes(workbookBytes, filename, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  const generatedFile = browserFile(workbookBytes, filename, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  setActiveWorkbook(generatedFile, "Default generated workbook");
  const languageCount = headers.filter((header) => header !== "values").length;
  return {
    message: `XLSX created: ${rows.length.toLocaleString()} rows, ${languageCount.toLocaleString()} languages, ${extracted.processedFiles.toLocaleString()} XML files.`,
    exact: false,
  };
}

function workbookData(matrix, resourceMap) {
  const headers = (matrix[0] || []).slice(3).map(String);
  const records = new Map();
  for (let index = 1; index < matrix.length; index += 1) {
    const row = matrix[index] || [];
    if (row[0] === null || row[1] === null || row[0] === undefined || row[1] === undefined) continue;
    const identity = resourceMap?.get(index) || null;
    const key = identity
      ? resourceIdentityKey(identity.templatePath, identity.resourceType, identity.id, identity.item)
      : rowKey(String(row[0]), String(row[1]), row[2]);
    if (records.has(key)) throw new Error(`Duplicate workbook resource at worksheet row ${index + 1}.`);
    records.set(key, { row, index, identity });
  }
  return { headers, records, mapped: Boolean(resourceMap) };
}

function nonBlank(value) {
  return value !== null && value !== undefined && value !== "";
}

function cloneAttributes(element) {
  return Array.from(element.attributes).map((attribute) => [attribute.name, attribute.value]);
}

function applyAttributes(element, attributes) {
  attributes.forEach(([name, value]) => element.setAttribute(name, value));
}

function targetValuesDir(column, variant) {
  return variant ? `${column}-v23` : column;
}

function replaceValuesDir(path, current, target) {
  const needle = `/${current}/`;
  const position = path.lastIndexOf(needle);
  if (position < 0) return null;
  return `${path.slice(0, position)}/${target}/${path.slice(position + needle.length)}`;
}

async function rebuildArchive(templateBytes, matrix, resourceMap, onProgress) {
  const zip = await JSZip.loadAsync(templateBytes);
  const workbook = workbookData(matrix, resourceMap);
  const columnIndex = new Map(workbook.headers.map((header, index) => [header, index + 3]));
  const seen = new Set();
  const definitions = new Map();
  const documents = new Map();
  const entries = Object.values(zip.files).filter((entry) => !entry.dir && parseResourcePath(entry.name));
  let changes = 0;

  for (let entryIndex = 0; entryIndex < entries.length; entryIndex += 1) {
    const entry = entries[entryIndex];
    const info = parseResourcePath(entry.name);
    const originalText = await entry.async("string");
    const documentNode = xmlDocument(originalText, entry.name);
    const state = { entry, info, documentNode, originalText, changed: false };
    documents.set(entry.name, state);
    const resources = documentNode.documentElement;
    for (const element of Array.from(resources.children)) {
      const id = element.getAttribute("name");
      if (!id) continue;
      const isArrayFile = info.filename === "arrays.xml" || info.filename === "strings.xml";
      if (element.localName === "string-array" && isArrayFile) {
        const items = directChildren(element, "item");
        items.forEach((item, itemIndex) => {
          const itemValue = info.variant ? `${itemIndex}${info.variant}` : itemIndex;
          const templatePath = resourcePathTemplate(entry.name, info.valuesDir);
          const key = workbook.mapped
            ? resourceIdentityKey(templatePath, "string-array", id, itemValue)
            : rowKey(info.app, id, itemValue);
          if (!workbook.mapped && definitions.has(key) &&
              (definitions.get(key).templatePath !== templatePath || definitions.get(key).kind !== "array")) {
            throw new Error(`Legacy workbook mapping is ambiguous for ${info.app} / ${id} / item ${itemValue}. Regenerate the XLSX with this converter before rebuilding the ZIP.`);
          }
          definitions.set(key, {
            kind: "array",
            sourcePath: entry.name,
            templatePath,
            info,
            parentAttributes: cloneAttributes(element),
            itemAttributes: cloneAttributes(item),
            itemIndex,
          });
          seen.add(`${key}\u001f${info.column}`);
          const record = workbook.records.get(key);
          const column = columnIndex.get(info.column);
          const value = record && column !== undefined ? record.row[column] : null;
          if (nonBlank(value)) {
            const workbookValue = String(value);
            if (workbookValue !== item.textContent) {
              changes += 1;
              item.textContent = workbookValue;
              state.changed = true;
            }
          }
        });
      } else if (
        element.localName === "string" &&
        (info.filename.startsWith("strings") || info.filename.endsWith("strings.xml"))
      ) {
        const itemValue = info.variant || null;
        const templatePath = resourcePathTemplate(entry.name, info.valuesDir);
        let key = workbook.mapped
          ? resourceIdentityKey(templatePath, "string", id, itemValue)
          : rowKey(info.app, id, itemValue);
        if (!workbook.mapped && !workbook.records.has(key) && !info.variant) {
          const numericKey = rowKey(info.app, id, 0);
          if (workbook.records.has(numericKey)) key = numericKey;
        }
        if (!workbook.mapped && definitions.has(key) &&
            (definitions.get(key).templatePath !== templatePath || definitions.get(key).kind !== "string")) {
          throw new Error(`Legacy workbook mapping is ambiguous for ${info.app} / ${id}. Regenerate the XLSX with this converter before rebuilding the ZIP.`);
        }
        definitions.set(key, { kind: "string", sourcePath: entry.name, templatePath, info, attributes: cloneAttributes(element) });
        seen.add(`${key}\u001f${info.column}`);
        const record = workbook.records.get(key);
        const column = columnIndex.get(info.column);
        const value = record && column !== undefined ? record.row[column] : null;
        if (nonBlank(value)) {
          const workbookValue = String(value);
          if (workbookValue !== element.textContent) {
            changes += 1;
            element.textContent = workbookValue;
            state.changed = true;
          }
        }
      }
    }
    if (entryIndex % 25 === 0) {
      onProgress?.(`Applying translations… ${entryIndex + 1} / ${entries.length}`);
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
  }

  for (const [key, record] of workbook.records) {
    const definition = definitions.get(key);
    if (!definition) continue;
    for (const [columnName, matrixColumn] of columnIndex) {
      const value = record.row[matrixColumn];
      if (!nonBlank(value) || seen.has(`${key}\u001f${columnName}`)) continue;
      const targetDir = targetValuesDir(columnName, definition.info.variant);
      const targetPath = definition.templatePath
        ? resourcePathForColumn(definition.templatePath, columnName, definition.info.variant)
        : replaceValuesDir(definition.sourcePath, definition.info.valuesDir, targetDir);
      if (!targetPath) continue;
      let state = documents.get(targetPath);
      if (!state) {
        const existing = zip.file(targetPath);
        const originalText = existing ? await existing.async("string") : '<?xml version="1.0" encoding="utf-8"?>\n<resources></resources>';
        const targetInfo = parseResourcePath(targetPath) || { ...definition.info, valuesDir: targetDir, column: columnName };
        state = { entry: existing, info: targetInfo, documentNode: xmlDocument(originalText, targetPath), originalText, changed: false };
        documents.set(targetPath, state);
      }
      const root = state.documentNode.documentElement;
      if (definition.kind === "string") {
        const element = state.documentNode.createElement("string");
        applyAttributes(element, definition.attributes);
        element.setAttribute("name", record.row[1]);
        element.textContent = String(value);
        root.appendChild(element);
      } else {
        let parent = Array.from(root.children).find((element) => element.localName === "string-array" && element.getAttribute("name") === record.row[1]);
        if (!parent) {
          parent = state.documentNode.createElement("string-array");
          applyAttributes(parent, definition.parentAttributes);
          parent.setAttribute("name", record.row[1]);
          root.appendChild(parent);
        }
        const item = state.documentNode.createElement("item");
        applyAttributes(item, definition.itemAttributes);
        item.textContent = String(value);
        parent.appendChild(item);
      }
      state.changed = true;
      changes += 1;
      seen.add(`${key}\u001f${columnName}`);
    }
  }

  const serializer = new XMLSerializer();
  let rewrittenDocuments = 0;
  for (const [path, state] of documents) {
    if (!state.changed) continue;
    let xml = serializer.serializeToString(state.documentNode);
    if (!xml.startsWith("<?xml")) xml = `<?xml version="1.0" encoding="utf-8"?>\n${xml}`;
    zip.file(path, xml);
    rewrittenDocuments += 1;
  }
  if (changes === 0) {
    return {
      output: templateBytes.slice(),
      changes,
      rewrittenDocuments: 0,
      validatedDocuments: entries.length,
      processedRows: workbook.records.size,
      languageCount: workbook.headers.filter((header) => header !== "values").length,
      exact: true,
    };
  }
  const output = await zip.generateAsync(
    { type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 6 } },
    (metadata) => onProgress?.(`Writing ZIP… ${Math.round(metadata.percent)}%`),
  );
  return {
    output,
    changes,
    rewrittenDocuments,
    validatedDocuments: entries.length,
    processedRows: workbook.records.size,
    languageCount: workbook.headers.filter((header) => header !== "values").length,
    exact: false,
  };
}

async function createArchive(xlsxFile, referenceFile, correctionFiles, capitalizationEnabled) {
  const xlsxBytes = new Uint8Array(await xlsxFile.arrayBuffer());
  const workbook = XLSX.read(xlsxBytes, { type: "array", cellFormula: false, cellStyles: false });
  const matrix = workbookMatrix(workbook);
  const resourceMap = readResourceMap(workbook, matrix);
  const metadata = readRoundTrip(workbook);
  const corrections = await applyCorrectionOptions(matrix, correctionFiles, capitalizationEnabled);
  const audit = modificationAudit(matrix, corrections.changes, pipeline.activeModifications, resourceMap, corrections.unmatchedChecks);

  let templateBytes;
  if (metadata) templateBytes = metadata.zipBytes;
  else if (referenceFile) templateBytes = new Uint8Array(await referenceFile.arrayBuffer());
  else throw new Error("This legacy workbook has no embedded archive. Select its source ZIP under “Reference ZIP”.");

  const rebuilt = await rebuildArchive(templateBytes, matrix, resourceMap, (message) => showStatus(ui.xlsxStatus, message));
  const filename = outputName(xlsxFile.name, "zip");
  downloadBytes(rebuilt.output, filename, "application/zip");
  const outputFile = browserFile(rebuilt.output, filename, "application/zip");
  setOutputZip(outputFile, audit);
  return {
    message: `ZIP created and selected for comparison: ${rebuilt.processedRows.toLocaleString()} rows, ${rebuilt.languageCount.toLocaleString()} languages, ${rebuilt.validatedDocuments.toLocaleString()} XML files validated, ${rebuilt.rewrittenDocuments.toLocaleString()} changed XML files rewritten${rebuilt.exact ? "; unchanged ZIP preserved byte-for-byte" : ""}. Corrections: ${correctionSummary(corrections)}.`,
    exact: rebuilt.exact,
    outputFile,
  };
}

ui.zipButton.addEventListener("click", async () => {
  const file = ui.zipInput.files[0];
  if (!file) {
    showStatus(ui.zipStatus, "Choose a ZIP archive first.", "error");
    return;
  }
  setBusy(ui.zipButton, true);
  showStatus(ui.zipStatus, "Opening archive…");
  try {
    setSourceZip(file);
    if (!languageState.ready || languageState.fileToken !== fileToken(file)) await scanArchiveLanguages(file);
    const result = await createWorkbook(file, selectedLanguageColumns());
    showStatus(ui.zipStatus, result.message, "success");
  } catch (error) {
    console.error(error);
    showStatus(ui.zipStatus, error.message || String(error), "error");
  } finally {
    setBusy(ui.zipButton, false);
  }
});

ui.xlsxButton.addEventListener("click", async () => {
  const file = activeWorkbook();
  if (!file) {
    showStatus(ui.xlsxStatus, "Create or choose an XLSX workbook first.", "error");
    return;
  }
  setBusy(ui.xlsxButton, true);
  showStatus(ui.xlsxStatus, "Opening workbook…");
  try {
    const result = await createArchive(
      file,
      ui.referenceInput.files[0] || null,
      selectedCorrectionFiles(),
      ui.capitalizationRule.checked,
    );
    showStatus(ui.xlsxStatus, result.message, "success");
  } catch (error) {
    console.error(error);
    showStatus(ui.xlsxStatus, error.message || String(error), "error");
  } finally {
    setBusy(ui.xlsxButton, false);
  }
});

ui.correctionTemplateButton.addEventListener("click", () => {
  downloadCorrectionTemplate();
  showStatus(ui.xlsxStatus, "Correction template downloaded.", "success");
});

ui.correctedXlsxButton.addEventListener("click", async () => {
  const file = activeWorkbook();
  if (!file) {
    showStatus(ui.xlsxStatus, "Create or choose an XLSX workbook first.", "error");
    return;
  }
  setBusy(ui.correctedXlsxButton, true);
  showStatus(ui.xlsxStatus, "Applying workbook corrections…");
  try {
    const result = await createCorrectedWorkbook(
      file,
      selectedCorrectionFiles(),
      ui.capitalizationRule.checked,
    );
    showStatus(ui.xlsxStatus, result.message, "success");
  } catch (error) {
    console.error(error);
    showStatus(ui.xlsxStatus, error.message || String(error), "error");
  } finally {
    setBusy(ui.correctedXlsxButton, false);
  }
});

ui.compareButton.addEventListener("click", async () => {
  const sourceFile = ui.compareSourceInput.files[0] || pipeline.sourceZip;
  const outputFile = ui.compareOutputInput.files[0] || pipeline.outputZip;
  if (!sourceFile || !outputFile) {
    showStatus(ui.compareStatus, "Choose both the source ZIP and output ZIP.", "error");
    return;
  }
  setBusy(ui.compareButton, true);
  showStatus(ui.compareStatus, "Opening archives…");
  try {
    const modifications = outputFile === pipeline.outputZip ? pipeline.outputModifications : new Map();
    const result = await compareArchives(sourceFile, outputFile, modifications);
    showStatus(
      ui.compareStatus,
      `Compared ${result.total.toLocaleString()} translation rows: ${result.matched.toLocaleString()} match, ${result.different.toLocaleString()} differ. XLSX report downloaded.`,
      "success",
    );
  } catch (error) {
    console.error(error);
    showStatus(ui.compareStatus, error.message || String(error), "error");
  } finally {
    setBusy(ui.compareButton, false);
  }
});
