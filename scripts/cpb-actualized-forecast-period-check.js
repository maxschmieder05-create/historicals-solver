const path = require("node:path");
const ExcelJS = require("exceljs");
const { postWorkbook } = require("./fill-workbook-api");

const repoRoot = path.resolve(__dirname, "..");
const ticker = process.env.CPB_TICKER || "CPB";
const cik = process.env.CPB_CIK || "0000016732";
const inputWorkbook = process.env.CPB_INPUT_WORKBOOK || "/Users/maxschmieder/Downloads/CPB_historicals_filled.xlsx";
const outputWorkbook =
  process.env.CPB_OUTPUT_WORKBOOK || path.join(repoRoot, "tmp", "cpb-actualized-forecast-period-output.xlsx");
const apiUrl = process.env.FILL_API_URL || "http://localhost:3000/api/fill-model";

const secHeaders = {
  "User-Agent": process.env.SEC_USER_AGENT || "HistoricalsSolver CPB regression contact@example.com"
};

const actualizedFormulaRows = [
  "Cash & Cash Equivalents",
  "Other Non-Current Assets",
  "Other Current Liabilities",
  "LT Debt (Incl. Current Portion)",
  "Other Non-Current Liabilities",
  "Retained Earnings"
];

const balanceSheetChecks = [
  { name: "total assets", labels: ["Total Assets"], concepts: ["Assets"] },
  { name: "total liabilities", labels: ["Total Liabilities"], concepts: ["Liabilities"] },
  { name: "shareholders' equity", labels: ["Total Shareholder's Equity", "Total Shareholders' Equity"], concepts: ["StockholdersEquity"] },
  { name: "liabilities plus equity", labels: ["Total Liabilities & Shareholder's Equity", "Total Liabilities and Shareholder's Equity"], concepts: ["Assets"] },
  { name: "balance sheet check", labels: ["Balance Sheet Check"], expected: 0 }
];

function cellValue(cell) {
  const value = cell.value;
  if (typeof value === "number") return value;
  if (typeof value === "string") return value;
  if (value && typeof value === "object") return value.result ?? value.text ?? value.formula ?? value.sharedFormula ?? null;
  return value;
}

function cellFormula(cell) {
  const value = cell.value;
  if (value && typeof value === "object" && typeof value.formula === "string") return value.formula;
  if (value && typeof value === "object" && typeof value.sharedFormula === "string") return value.sharedFormula;
  return null;
}

function normalize(value) {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function normalizePeriodLabel(label) {
  const compact = String(label ?? "")
    .trim()
    .replace(/\s+/g, "")
    .replace(/[’']/g, "")
    .replace(/(?:E|EST|ESTIMATE)$/i, "");
  const direct = compact.match(/^([1-4])Q(\d{2}|\d{4})$/i);
  if (direct) return `${direct[1]}Q${direct[2].slice(-2)}`.toUpperCase();
  const fiscalYear = compact.match(/^(?:FY(\d{2}|\d{4})|(\d{4}))A?$/i);
  if (fiscalYear) return `FY${(fiscalYear[1] ?? fiscalYear[2]).slice(-2)}`.toUpperCase();
  return compact.toUpperCase();
}

function isEstimatePeriodLabel(label) {
  const compact = String(label ?? "")
    .trim()
    .replace(/\s+/g, "")
    .replace(/[’']/g, "");
  return /(?:^|[^a-z])(?:E|EST|ESTIMATE)$/i.test(compact) || /(?:\d{2}|\d{4})E$/i.test(compact);
}

function bestPeriodHeaderRow(sheet) {
  let best = null;
  for (let row = 1; row <= Math.min(sheet.rowCount, 120); row += 1) {
    let validCount = 0;
    let quarterCount = 0;
    for (let col = 1; col <= Math.min(sheet.columnCount, 180); col += 1) {
      const period = normalizePeriodLabel(cellValue(sheet.getCell(row, col)));
      if (/^[1-4]Q\d{2}$/.test(period)) {
        validCount += 1;
        quarterCount += 1;
      } else if (/^FY\d{2}$/.test(period)) {
        validCount += 1;
      }
    }
    const score = validCount + quarterCount * 3;
    if (validCount && (!best || score > best.score)) best = { row, score };
  }
  return best?.row ?? null;
}

function findPeriodColumn(sheet, period) {
  const headerRow = bestPeriodHeaderRow(sheet);
  if (!headerRow) return null;
  for (let col = 1; col <= Math.min(sheet.columnCount, 180); col += 1) {
    if (normalizePeriodLabel(cellValue(sheet.getCell(headerRow, col))) === period.toUpperCase()) return col;
  }
  return null;
}

function rowLabel(sheet, row) {
  const candidates = [];
  for (const col of [1, 2, 3, 4, 5]) {
    const text = String(cellValue(sheet.getCell(row, col)) ?? "").trim();
    if (text && !/^x$/i.test(text)) candidates.push(text);
  }
  return candidates.sort((a, b) => b.length - a.length)[0] ?? "";
}

function findRowInBalanceSheet(sheet, labels) {
  const wanted = new Set(labels.map(normalize));
  for (const start of findRows(sheet, ["Balance Sheet"])) {
    for (let row = start + 1; row <= sheet.rowCount; row += 1) {
      const label = rowLabel(sheet, row);
      if (!label) continue;
      if (/working capital|cash flow statement|cashflow statement|income statement|schedule|analysis|drivers/i.test(label)) break;
      if (wanted.has(normalize(label))) return row;
    }
  }
  return null;
}

function findRow(sheet, labels) {
  return findRows(sheet, labels)[0] ?? null;
}

function findRows(sheet, labels) {
  const wanted = new Set(labels.map(normalize));
  const rows = [];
  for (let row = 1; row <= sheet.rowCount; row += 1) {
    if (wanted.has(normalize(rowLabel(sheet, row)))) rows.push(row);
  }
  return rows;
}

function valuesMatch(actual, expected, tolerance = 0.5) {
  return typeof actual === "number" && Number.isFinite(actual) && Math.abs(actual - expected) <= tolerance;
}

async function readWorkbook(file) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(file);
  return workbook;
}

async function fetchCompanyFacts() {
  const response = await fetch(`https://data.sec.gov/api/xbrl/companyfacts/CIK${cik}.json`, { headers: secHeaders });
  if (!response.ok) throw new Error(`Could not load SEC companyfacts for CIK ${cik}: ${response.status} ${response.statusText}`);
  return response.json();
}

function expectedConceptValue(payload, period, concepts) {
  for (const concept of concepts) {
    const fact = expectedFact(payload, period, concept);
    if (fact) return fact.val / 1_000_000;
  }
  return null;
}

function expectedFact(payload, period, concept) {
  const unitFacts = payload.facts?.["us-gaap"]?.[concept]?.units?.USD;
  if (!Array.isArray(unitFacts)) return null;
  const candidates = unitFacts.filter((fact) => factMatchesPeriod(fact, period) && fact.end && !fact.start && (fact.form === "10-Q" || fact.form === "10-K"));
  return candidates.sort(compareExpectedFactPreference)[0] ?? null;
}

function factMatchesPeriod(fact, period) {
  const quarter = Number(period[0]);
  const fiscalYear = 2000 + Number(period.slice(2));
  if (fact.fy !== fiscalYear) return false;
  if (quarter === 4) return fact.fp === "FY" || fact.fp === "Q4";
  return fact.fp === `Q${quarter}`;
}

function compareExpectedFactPreference(a, b) {
  const endCompare = String(b.end).localeCompare(String(a.end));
  if (endCompare !== 0) return endCompare;
  const formCompare = formScore(b.form) - formScore(a.form);
  if (formCompare !== 0) return formCompare;
  return String(b.filed).localeCompare(String(a.filed));
}

function formScore(form) {
  if (form === "10-K") return 2;
  if (form === "10-Q") return 1;
  return 0;
}

function auditRowsFor(audit, cell, period) {
  const rows = [];
  if (!audit) return rows;
  const headers = auditHeaderColumns(audit);
  const cellCol = headers.get("cell/range") ?? 2;
  const periodCol = headers.get("period") ?? 5;
  const formulaStatusCol = headers.get("formula status") ?? 19;
  const notesCol = headers.get("notes") ?? 24;
  for (let row = 2; row <= audit.rowCount; row += 1) {
    if (String(cellValue(audit.getCell(row, cellCol)) ?? "") !== cell) continue;
    if (String(cellValue(audit.getCell(row, periodCol)) ?? "").toUpperCase() !== period.toUpperCase()) continue;
    rows.push({
      formulaStatus: String(cellValue(audit.getCell(row, formulaStatusCol)) ?? ""),
      notes: String(cellValue(audit.getCell(row, notesCol)) ?? "")
    });
  }
  return rows;
}

function auditHeaderColumns(audit) {
  const headers = new Map();
  for (let col = 1; col <= audit.columnCount; col += 1) {
    const header = String(cellValue(audit.getCell(1, col)) ?? "")
      .trim()
      .toLowerCase();
    if (header) headers.set(header, col);
  }
  return headers;
}

function filingMapRowsFor(filingMap, period) {
  const rows = [];
  if (!filingMap) return rows;
  for (let row = 2; row <= filingMap.rowCount; row += 1) {
    if (String(cellValue(filingMap.getCell(row, 1)) ?? "").toUpperCase() !== period.toUpperCase()) continue;
    rows.push({
      form: String(cellValue(filingMap.getCell(row, 3)) ?? ""),
      periodEndDate: String(cellValue(filingMap.getCell(row, 4)) ?? "")
    });
  }
  return rows;
}

async function main() {
  await postWorkbook({ apiUrl, ticker, inputWorkbook, outputWorkbook });

  const output = await readWorkbook(outputWorkbook);
  const model = output.getWorksheet("Model");
  const audit = output.getWorksheet("Mapping Audit");
  const filingMap = output.getWorksheet("Filing Period Map");
  if (!model) throw new Error("Output workbook does not contain a Model sheet.");
  if (!audit) throw new Error("Output workbook does not contain a Mapping Audit sheet.");
  if (!filingMap) throw new Error("Output workbook does not contain a Filing Period Map sheet.");

  const errors = [];
  const headerRow = bestPeriodHeaderRow(model);
  const actualCol = findPeriodColumn(model, "3Q26");
  const forecastCol = findPeriodColumn(model, "4Q26");
  if (!headerRow) errors.push("Could not find a period header row.");
  if (!actualCol) errors.push("Could not find 3Q26 column in Model sheet.");
  if (!forecastCol) errors.push("Could not find 4Q26e column in Model sheet.");
  if (!headerRow || !actualCol || !forecastCol) throw new Error(errors.join("\n"));

  const actualHeader = String(cellValue(model.getCell(headerRow, actualCol)) ?? "");
  const forecastHeader = String(cellValue(model.getCell(headerRow, forecastCol)) ?? "");
  if (isEstimatePeriodLabel(actualHeader)) errors.push(`3Q26 header should be actualized, got "${actualHeader}".`);
  if (!isEstimatePeriodLabel(forecastHeader)) errors.push(`4Q26 header should remain forecast/protected, got "${forecastHeader}".`);

  const filingRows = filingMapRowsFor(filingMap, "3Q26");
  if (!filingRows.some((row) => row.form === "10-Q" && row.periodEndDate === "2026-05-03")) {
    errors.push("SEC Filing Period Map should tie 3Q26 to Campbell's 10-Q period ended 2026-05-03.");
  }

  for (const label of actualizedFormulaRows) {
    const row = findRowInBalanceSheet(model, [label]);
    if (!row) {
      errors.push(`Could not find balance-sheet row "${label}".`);
      continue;
    }
    const actualCell = model.getCell(row, actualCol);
    const forecastCell = model.getCell(row, forecastCol);
    if (cellFormula(actualCell)) errors.push(`3Q26 ${label} ${actualCell.address}: actual SEC period should not preserve forecast formula.`);
    if (!cellFormula(forecastCell)) errors.push(`4Q26e ${label} ${forecastCell.address}: future forecast formula should be preserved.`);
    const auditRows = auditRowsFor(audit, actualCell.address, "3Q26");
    if (!auditRows.some((row) => /actualized forecast formula replaced/i.test(row.formulaStatus) || /Actualized forecast column/i.test(row.notes))) {
      errors.push(`Mapping Audit should mark ${actualCell.address} as an actualized forecast formula replacement.`);
    }
  }

  const facts = await fetchCompanyFacts();
  for (const check of balanceSheetChecks) {
    const row = findRowInBalanceSheet(model, check.labels);
    if (!row) {
      errors.push(`Could not find ${check.name} row.`);
      continue;
    }
    const expected = check.expected ?? expectedConceptValue(facts, "3Q26", check.concepts);
    if (expected === null) {
      errors.push(`Could not resolve SEC expected value for ${check.name}.`);
      continue;
    }
    const actual = cellValue(model.getCell(row, actualCol));
    if (!valuesMatch(actual, expected)) {
      errors.push(`3Q26 ${check.name} ${model.getCell(row, actualCol).address}: expected ${expected}, got ${actual ?? "[blank]"}.`);
    }
  }

  if (errors.length) {
    console.error(errors.join("\n"));
    throw new Error(`CPB actualized forecast period regression failed with ${errors.length} issue(s).`);
  }

  console.log(`CPB actualized forecast period regression passed: ${outputWorkbook}`);
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
