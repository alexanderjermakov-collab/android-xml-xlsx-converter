/* global JSZip, XLSX */
"use strict";

const ROUNDTRIP_SHEET = "_RoundTrip";
const ROUNDTRIP_MARKER = "ANDROID_XML_XLSX_CONVERTER_V1";
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
  languagePanel: document.querySelector("#languagePanel"),
  languageList: document.querySelector("#languageList"),
  languageCount: document.querySelector("#languageCount"),
  languageSearch: document.querySelector("#languageSearch"),
  selectAllLanguages: document.querySelector("#selectAllLanguages"),
  clearLanguages: document.querySelector("#clearLanguages"),
};

let languageState = { fileToken: "", columns: [], ready: false };

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
    nameNode.textContent = input.files[0]?.name || "";
  });
  ["dragenter", "dragover"].forEach((eventName) => dropNode.addEventListener(eventName, () => dropNode.classList.add("dragging")));
  ["dragleave", "drop"].forEach((eventName) => dropNode.addEventListener(eventName, () => dropNode.classList.remove("dragging")));
}

wireFileInput(ui.zipInput, "zipName", "zipDrop");
wireFileInput(ui.xlsxInput, "xlsxName", "xlsxDrop");
wireFileInput(ui.referenceInput, "referenceName", "referenceDrop");

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

function groupKey(app, id, variant) {
  return `${app}\u001f${id}\u001f${variant || ""}`;
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
      const key = groupKey(info.app, id, info.variant);
      let group = groups.get(key);
      if (!group) {
        group = { app: info.app, id, variant: info.variant, isArray: false, items: [] };
        groups.set(key, group);
      }
      const isArrayFile = info.filename === "arrays.xml" || info.filename === "strings.xml";
      if (element.localName === "string-array" && isArrayFile) {
        group.isArray = true;
        const items = directChildren(element, "item");
        items.forEach((item, index) => {
          group.items[index] ||= new Map();
          group.items[index].set(info.column, item.textContent || null);
        });
      } else if (
        element.localName === "string" &&
        (info.filename.startsWith("strings") || info.filename.endsWith("strings.xml"))
      ) {
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

function flattenGroups(groups) {
  const rows = [];
  for (const group of groups.values()) {
    group.items.forEach((values, index) => {
      if (!values) return;
      let item = null;
      if (group.variant) item = group.isArray ? `${index}${group.variant}` : group.variant;
      else if (group.isArray) item = index;
      rows.push({ app: group.app, id: group.id, item, values });
    });
  }
  return rows;
}

function orderRows(rows, schema) {
  const byKey = new Map(rows.map((row) => [rowKey(row.app, row.id, row.item), row]));
  const ordered = [];
  for (const [app, id, item] of schema.rows) {
    const key = rowKey(app, id, item);
    const row = byKey.get(key);
    if (row) {
      ordered.push(row);
      byKey.delete(key);
    }
  }
  ordered.push(...byKey.values());
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

function workbookMatrix(workbook) {
  const sheetName = workbook.SheetNames.includes(SHEET_NAME) ? SHEET_NAME : workbook.SheetNames[0];
  if (!sheetName) throw new Error("The workbook contains no worksheets.");
  const matrix = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, raw: true, defval: null });
  const header = matrix[0] || [];
  if (header[0] !== "App" || header[1] !== "string ID" || header[2] !== "item ID") {
    throw new Error("Expected columns App, string ID, and item ID in the first row.");
  }
  return matrix;
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
  workbook.Workbook = { Sheets: [{ name: SHEET_NAME, Hidden: 0 }, { name: ROUNDTRIP_SHEET, Hidden: 2 }] };
  const workbookBytes = new Uint8Array(XLSX.write(workbook, { bookType: "xlsx", type: "array", compression: true }));
  const filename = outputName(zipFile.name, "xlsx");
  downloadBytes(workbookBytes, filename, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  const languageCount = headers.filter((header) => header !== "values").length;
  return {
    message: `XLSX created: ${rows.length.toLocaleString()} rows, ${languageCount.toLocaleString()} languages, ${extracted.processedFiles.toLocaleString()} XML files.`,
    exact: false,
  };
}

function workbookData(matrix) {
  const headers = (matrix[0] || []).slice(3).map(String);
  const records = new Map();
  for (let index = 1; index < matrix.length; index += 1) {
    const row = matrix[index] || [];
    if (row[0] === null || row[1] === null || row[0] === undefined || row[1] === undefined) continue;
    records.set(rowKey(String(row[0]), String(row[1]), row[2]), { row, index });
  }
  return { headers, records };
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

async function rebuildArchive(templateBytes, matrix, onProgress) {
  const zip = await JSZip.loadAsync(templateBytes);
  const workbook = workbookData(matrix);
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
          const key = rowKey(info.app, id, itemValue);
          definitions.set(key, {
            kind: "array",
            sourcePath: entry.name,
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
            if (workbookValue !== item.textContent) changes += 1;
            item.textContent = workbookValue;
            state.changed = true;
          }
        });
      } else if (
        element.localName === "string" &&
        (info.filename.startsWith("strings") || info.filename.endsWith("strings.xml"))
      ) {
        const numericKey = rowKey(info.app, id, 0);
        const itemValue = info.variant ? info.variant : (workbook.records.has(numericKey) ? 0 : null);
        const key = rowKey(info.app, id, itemValue);
        definitions.set(key, { kind: "string", sourcePath: entry.name, info, attributes: cloneAttributes(element) });
        seen.add(`${key}\u001f${info.column}`);
        const record = workbook.records.get(key);
        const column = columnIndex.get(info.column);
        const value = record && column !== undefined ? record.row[column] : null;
        if (nonBlank(value)) {
          const workbookValue = String(value);
          if (workbookValue !== element.textContent) changes += 1;
          element.textContent = workbookValue;
          state.changed = true;
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
      const targetPath = replaceValuesDir(definition.sourcePath, definition.info.valuesDir, targetDir);
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
    let xml = serializer.serializeToString(state.documentNode);
    if (!xml.startsWith("<?xml")) xml = `<?xml version="1.0" encoding="utf-8"?>\n${xml}`;
    zip.file(path, xml);
    rewrittenDocuments += 1;
  }
  const output = await zip.generateAsync(
    { type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 6 } },
    (metadata) => onProgress?.(`Writing ZIP… ${Math.round(metadata.percent)}%`),
  );
  return {
    output,
    changes,
    rewrittenDocuments,
    processedRows: workbook.records.size,
    languageCount: workbook.headers.filter((header) => header !== "values").length,
  };
}

async function createArchive(xlsxFile, referenceFile) {
  const xlsxBytes = new Uint8Array(await xlsxFile.arrayBuffer());
  const workbook = XLSX.read(xlsxBytes, { type: "array", cellFormula: false, cellStyles: false });
  const matrix = workbookMatrix(workbook);
  const metadata = readRoundTrip(workbook);

  let templateBytes;
  if (metadata) templateBytes = metadata.zipBytes;
  else if (referenceFile) templateBytes = new Uint8Array(await referenceFile.arrayBuffer());
  else throw new Error("This legacy workbook has no embedded archive. Select its source ZIP under “Reference ZIP”.");

  const rebuilt = await rebuildArchive(templateBytes, matrix, (message) => showStatus(ui.xlsxStatus, message));
  const filename = outputName(xlsxFile.name, "zip");
  downloadBytes(rebuilt.output, filename, "application/zip");
  return {
    message: `ZIP created: ${rebuilt.processedRows.toLocaleString()} rows, ${rebuilt.languageCount.toLocaleString()} languages, ${rebuilt.rewrittenDocuments.toLocaleString()} XML files (${rebuilt.changes.toLocaleString()} changed values).`,
    exact: false,
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
  const file = ui.xlsxInput.files[0];
  if (!file) {
    showStatus(ui.xlsxStatus, "Choose an XLSX workbook first.", "error");
    return;
  }
  setBusy(ui.xlsxButton, true);
  showStatus(ui.xlsxStatus, "Opening workbook…");
  try {
    const result = await createArchive(file, ui.referenceInput.files[0] || null);
    showStatus(ui.xlsxStatus, result.message, "success");
  } catch (error) {
    console.error(error);
    showStatus(ui.xlsxStatus, error.message || String(error), "error");
  } finally {
    setBusy(ui.xlsxButton, false);
  }
});
