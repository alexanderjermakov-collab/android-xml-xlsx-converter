(function exposeCorrectionEngine(root, factory) {
  const api = factory();
  if (typeof module === "object" && module.exports) module.exports = api;
  if (root) root.CorrectionEngine = api;
}(typeof globalThis !== "undefined" ? globalThis : this, () => {
  "use strict";

  function text(value) {
    return value === null || value === undefined ? "" : String(value);
  }

  function normalized(value) {
    return text(value).trim().toLocaleLowerCase().replace(/[\s_-]+/g, " ");
  }

  function headerIndex(headers, aliases) {
    const accepted = new Set(aliases.map(normalized));
    return headers.findIndex((header) => accepted.has(normalized(header)));
  }

  function resolveLanguageToken(token, headers) {
    const value = text(token).trim();
    if (!value) return [];
    const lower = value.toLocaleLowerCase().replace(/_/g, "-");
    if (["*", "all", "all languages"].includes(lower)) return [...headers];

    const exact = headers.find((header) => header.toLocaleLowerCase() === lower);
    if (exact) return [exact];

    const candidate = lower.startsWith("values-") ? lower : `values-${lower}`;
    const matched = headers.find((header) => header.toLocaleLowerCase() === candidate);
    if (matched) return [matched];

    if (["en", "english"].includes(lower) && headers.includes("values")) return ["values"];
    return [];
  }

  function itemKey(value) {
    if (value === null || value === undefined || value === "") return "";
    return String(value);
  }

  function recordKey(app, id, item) {
    return `${text(app)}\u001f${text(id)}\u001f${itemKey(item)}`;
  }

  function parseKeyedRules(matrix, translationHeaders) {
    const headings = matrix[0] || [];
    const appColumn = headerIndex(headings, ["App"]);
    const idColumn = headerIndex(headings, ["string ID"]);
    const itemColumn = headerIndex(headings, ["item ID"]);
    if (appColumn < 0 || idColumn < 0 || itemColumn < 0) return null;

    const pairs = [];
    for (let column = 0; column < headings.length; column += 1) {
      const match = /^\s*(values(?:-[^\s(]+)?)\s*\(original\)\s*$/i.exec(text(headings[column]));
      if (!match) continue;
      const language = translationHeaders.find((header) => header.toLocaleLowerCase() === match[1].toLocaleLowerCase());
      if (!language) throw new Error(`Correction workbook targets missing language ${match[1]}.`);
      const correctedColumn = headings.findIndex((heading) => {
        const corrected = /^\s*(values(?:-[^\s(]+)?)\s*\(corrected\)\s*$/i.exec(text(heading));
        return corrected && corrected[1].toLocaleLowerCase() === match[1].toLocaleLowerCase();
      });
      if (correctedColumn < 0) throw new Error(`Correction workbook is missing ${match[1]} (corrected).`);
      pairs.push({ language, originalColumn: column, correctedColumn });
    }
    if (!pairs.length) return null;

    const rules = [];
    for (let rowIndex = 1; rowIndex < matrix.length; rowIndex += 1) {
      const row = matrix[rowIndex] || [];
      const app = text(row[appColumn]);
      const id = text(row[idColumn]);
      if (!app && !id && pairs.every((pair) => !text(row[pair.originalColumn]) && !text(row[pair.correctedColumn]))) continue;
      if (!app || !id) throw new Error(`Correction row ${rowIndex + 1}: App and string ID are required.`);
      for (const pair of pairs) {
        const oldText = text(row[pair.originalColumn]);
        const newText = text(row[pair.correctedColumn]);
        if (!oldText && !newText) continue;
        if (!oldText) throw new Error(`Correction row ${rowIndex + 1}: ${pair.language} (original) is empty.`);
        if (!newText) throw new Error(`Correction row ${rowIndex + 1}: ${pair.language} (corrected) is empty.`);
        rules.push({
          kind: "keyed",
          language: pair.language,
          app,
          id,
          item: row[itemColumn],
          oldText,
          newText,
          sourceRow: rowIndex + 1,
        });
      }
    }
    if (!rules.length) throw new Error("The correction workbook contains no correction rows.");
    return rules;
  }

  function parseRules(matrix, translationHeaders) {
    const keyedRules = parseKeyedRules(matrix, translationHeaders);
    if (keyedRules) return keyedRules;
    const headings = matrix[0] || [];
    const languageColumn = headerIndex(headings, ["Language", "Languages", "Language folder", "Values folder"]);
    const oldColumn = headerIndex(headings, ["Old text", "Current text", "Old", "Current"]);
    const newColumn = headerIndex(headings, ["New text", "Replacement text", "New", "Replacement"]);
    if (languageColumn < 0 || oldColumn < 0 || newColumn < 0) {
      throw new Error("Correction workbook must contain Language, Old text, and New text columns.");
    }

    const rules = [];
    for (let rowIndex = 1; rowIndex < matrix.length; rowIndex += 1) {
      const row = matrix[rowIndex] || [];
      const languageValue = text(row[languageColumn]).trim();
      const oldText = text(row[oldColumn]);
      const newText = text(row[newColumn]);
      if (!languageValue && !oldText && !newText) continue;
      if (!oldText) throw new Error(`Correction row ${rowIndex + 1}: Old text is empty.`);
      if (!newText) throw new Error(`Correction row ${rowIndex + 1}: New text is empty.`);

      const tokens = languageValue ? languageValue.split(/[;,\n]+/) : ["all"];
      const languages = [];
      const unknown = [];
      for (const token of tokens) {
        const matches = resolveLanguageToken(token, translationHeaders);
        if (matches.length) languages.push(...matches);
        else if (token.trim()) unknown.push(token.trim());
      }
      if (unknown.length) throw new Error(`Correction row ${rowIndex + 1}: unknown language ${unknown.join(", ")}.`);
      rules.push({ kind: "text", languages: [...new Set(languages)], oldText, newText, sourceRow: rowIndex + 1 });
    }
    if (!rules.length) throw new Error("The correction workbook contains no replacement rows.");
    return rules;
  }

  function applyRules(matrix, rules) {
    const headers = (matrix[0] || []).slice(3).map(text);
    const columnByLanguage = new Map(headers.map((header, index) => [header, index + 3]));
    const rowByKey = new Map();
    for (let row = 1; row < matrix.length; row += 1) {
      const values = matrix[row] || [];
      const key = recordKey(values[0], values[1], values[2]);
      if (!rowByKey.has(key)) rowByKey.set(key, []);
      rowByKey.get(key).push(row);
    }
    const changedCells = new Map();
    let replacements = 0;
    let matchedRules = 0;
    let alreadyApplied = 0;
    const ruleOutcomes = [];

    for (const rule of rules) {
      let ruleMatches = 0;
      if (rule.kind === "keyed") {
        const column = columnByLanguage.get(rule.language);
        const rows = rowByKey.get(recordKey(rule.app, rule.id, rule.item)) || [];
        if (column !== undefined) for (const row of rows) {
          const current = text(matrix[row]?.[column]);
          if (current === rule.oldText) {
            matrix[row][column] = rule.newText;
            changedCells.set(`${row}:${column}`, {
              row,
              column,
              value: rule.newText,
              language: rule.language,
              sourceRow: rule.sourceRow,
              oldText: rule.oldText,
              newText: rule.newText,
              alreadyApplied: false,
            });
            replacements += 1;
            ruleMatches = 1;
          } else if (current === rule.newText) {
            changedCells.set(`${row}:${column}`, {
              row,
              column,
              value: rule.newText,
              language: rule.language,
              sourceRow: rule.sourceRow,
              oldText: rule.oldText,
              newText: rule.newText,
              alreadyApplied: true,
            });
            alreadyApplied += 1;
            ruleMatches = 1;
          }
        }
      } else {
        for (const language of rule.languages) {
          const column = columnByLanguage.get(language);
          if (column === undefined) continue;
          for (let row = 1; row < matrix.length; row += 1) {
            if (text(matrix[row]?.[column]) !== rule.oldText) continue;
            matrix[row][column] = rule.newText;
            changedCells.set(`${row}:${column}`, {
              row,
              column,
              value: rule.newText,
              language,
              sourceRow: rule.sourceRow,
              oldText: rule.oldText,
              newText: rule.newText,
              alreadyApplied: false,
            });
            replacements += 1;
            ruleMatches += 1;
          }
        }
      }
      if (ruleMatches) matchedRules += 1;
      ruleOutcomes.push({ rule, matches: ruleMatches });
    }
    return {
      replacements,
      matchedRules,
      unmatchedRules: rules.length - matchedRules,
      unmatchedRuleDetails: ruleOutcomes.filter((outcome) => outcome.matches === 0).map((outcome) => outcome.rule),
      alreadyApplied,
      changes: [...changedCells.values()],
    };
  }

  function firstLetter(value) {
    const source = text(value);
    for (let index = 0; index < source.length;) {
      const character = String.fromCodePoint(source.codePointAt(index));
      if (/\p{L}/u.test(character)) return { character, index, length: character.length };
      index += character.length;
    }
    return null;
  }

  function isUppercaseStart(value) {
    const found = firstLetter(value);
    if (!found) return false;
    const upper = found.character.toLocaleUpperCase();
    const lower = found.character.toLocaleLowerCase();
    return upper === found.character && lower !== found.character;
  }

  function localeForColumn(column) {
    if (!column || column === "values") return undefined;
    let qualifier = column.replace(/^values-/, "");
    if (qualifier.startsWith("b+")) qualifier = qualifier.slice(2).replace(/\+/g, "-");
    else qualifier = qualifier.replace(/-r([A-Za-z]{2,3})$/, "-$1");
    const locale = qualifier.split("-").filter(Boolean).join("-");
    return locale || undefined;
  }

  function uppercaseFirstLetter(value, locale) {
    const source = text(value);
    const found = firstLetter(source);
    if (!found) return source;
    let replacement;
    try {
      replacement = locale ? found.character.toLocaleUpperCase(locale) : found.character.toLocaleUpperCase();
    } catch (_error) {
      replacement = found.character.toLocaleUpperCase();
    }
    if (replacement === found.character) return source;
    return `${source.slice(0, found.index)}${replacement}${source.slice(found.index + found.length)}`;
  }

  function englishColumn(headers) {
    const preferred = ["values", "values-en", "values-en-rUS", "values-en-rGB"];
    for (const name of preferred) {
      const index = headers.findIndex((header) => header.toLocaleLowerCase() === name.toLocaleLowerCase());
      if (index >= 0) return index + 3;
    }
    const index = headers.findIndex((header) => /^values-en(?:-|$)/i.test(header));
    return index >= 0 ? index + 3 : -1;
  }

  function applyCapitalization(matrix) {
    const headers = (matrix[0] || []).slice(3).map(text);
    const sourceColumn = englishColumn(headers);
    if (sourceColumn < 0) throw new Error("Capitalization rule needs an English column (values or values-en).");
    const changedCells = new Map();
    let capitalized = 0;

    for (let row = 1; row < matrix.length; row += 1) {
      if (!isUppercaseStart(matrix[row]?.[sourceColumn])) continue;
      for (let column = 3; column < (matrix[row] || []).length; column += 1) {
        if (column === sourceColumn || matrix[row][column] === null || matrix[row][column] === undefined || matrix[row][column] === "") continue;
        const header = headers[column - 3];
        const updated = uppercaseFirstLetter(matrix[row][column], localeForColumn(header));
        if (updated === text(matrix[row][column])) continue;
        matrix[row][column] = updated;
        changedCells.set(`${row}:${column}`, { row, column, value: updated, language: header });
        capitalized += 1;
      }
    }
    return { capitalized, changes: [...changedCells.values()] };
  }

  return { applyCapitalization, applyRules, parseRules, uppercaseFirstLetter, isUppercaseStart };
}));
