const path = require("node:path");
const ExcelJS = require("exceljs");
const { postWorkbook } = require("./fill-workbook-api");

const repoRoot = path.resolve(__dirname, "..");
const inputWorkbook = process.env.LLY_INPUT_WORKBOOK || "/Users/maxschmieder/Downloads/LLY_historicals_filled.xlsx";
const outputWorkbook = process.env.LLY_OUTPUT_WORKBOOK || path.join(repoRoot, "tmp", "lly-income-statement-classification-output.xlsx");
const apiUrl = process.env.FILL_API_URL || "http://localhost:3000/api/fill-model";
const ticker = process.env.LLY_TICKER || "LLY";

const expected = {
  "1Q24": {
    "Depreciation & Amortization": 0,
    "Other Operating Income (Expense)": -110.5,
    EBIT: 2509
  },
  "3Q25": {
    "Depreciation & Amortization": 0,
    "Other Operating Income (Expense)": -1020.6,
    EBIT: 7365.5
  },
  FY25: {
    "Depreciation & Amortization": 0,
    "Other Operating Income (Expense)": -3394,
    EBIT: 26302
  },
  "1Q26": {
    "Depreciation & Amortization": 0,
    "Other Operating Income (Expense)": -863,
    EBIT: 8915,
    "Interest (Expense)": -211,
    "Other Non-Operating Income (Expense)": -65
  }
};

function cellValue(cell) {
  const value = cell.value;
  if (typeof value === "number" || typeof value === "string") return value;
  if (value && typeof value === "object") return value.result ?? value.text ?? value.formula ?? null;
  return value;
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

function bestPeriodHeaderRow(sheet) {
  let best = null;
  for (let row = 1; row <= Math.min(sheet.rowCount, 120); row += 1) {
    let validCount = 0;
    let quarterCount = 0;
    for (let col = 1; col <= Math.min(sheet.columnCount, 160); col += 1) {
      const label = String(cellValue(sheet.getCell(row, col)) ?? "").trim();
      if (/^[1-4]Q\d{2}$/i.test(label)) {
        validCount += 1;
        quarterCount += 1;
      } else if (/^FY\d{2}$/i.test(label) || /^20\d{2}$/.test(label)) {
        validCount += 1;
      }
    }
    if (!validCount) continue;
    const score = validCount + quarterCount * 3;
    if (!best || score > best.score) best = { row, score };
  }
  return best?.row ?? null;
}

function findPeriodColumn(sheet, period) {
  const headerRow = bestPeriodHeaderRow(sheet);
  if (!headerRow) return null;
  for (let col = 1; col <= Math.min(sheet.columnCount, 160); col += 1) {
    const label = String(cellValue(sheet.getCell(headerRow, col)) ?? "").trim();
    if (normalizePeriodLabel(label) === period.toUpperCase()) return col;
  }
  return null;
}

function findRow(sheet, label) {
  const wanted = normalize(label);
  for (let row = 1; row <= sheet.rowCount; row += 1) {
    for (const col of [1, 2, 3, 4, 5]) {
      if (normalize(cellValue(sheet.getCell(row, col))) === wanted) return row;
    }
  }
  return null;
}

function valuesMatch(actual, expectedValue) {
  return typeof actual === "number" && Math.abs(actual - expectedValue) <= 0.5;
}

function auditRowsFor(audit, cell, period) {
  const rows = [];
  audit.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    if (String(cellValue(row.getCell(2)) ?? "") !== cell) return;
    if (String(cellValue(row.getCell(5)) ?? "").toUpperCase() !== period.toUpperCase()) return;
    rows.push({
      concepts: String(cellValue(row.getCell(8)) ?? ""),
      labels: String(cellValue(row.getCell(9)) ?? ""),
      validationStatus: String(cellValue(row.getCell(23)) ?? "")
    });
  });
  return rows;
}

async function main() {
  await postWorkbook({ apiUrl, ticker, inputWorkbook, outputWorkbook });

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(outputWorkbook);
  const model = workbook.getWorksheet("Model");
  const audit = workbook.getWorksheet("Mapping Audit");
  if (!model) throw new Error("Output workbook does not contain a Model sheet.");
  if (!audit) throw new Error("Output workbook does not contain a Mapping Audit sheet.");

  const errors = [];
  for (const [period, rows] of Object.entries(expected)) {
    const col = findPeriodColumn(model, period);
    if (!col) {
      errors.push(`Could not find ${period} column in Model sheet.`);
      continue;
    }

    for (const [label, expectedValue] of Object.entries(rows)) {
      const row = findRow(model, label);
      if (!row) {
        errors.push(`Could not find row "${label}" in Model sheet.`);
        continue;
      }
      const actual = cellValue(model.getCell(row, col));
      if (!valuesMatch(actual, expectedValue)) {
        errors.push(`${period} ${label}: expected ${expectedValue}, got ${actual ?? "[blank]"}.`);
      }
    }
  }

  const q126DaRows = auditRowsFor(audit, "U34", "1Q26");
  if (!q126DaRows.some((row) => /IncomeStatementDepreciationAmortizationNotReported/.test(row.concepts))) {
    errors.push("Mapping Audit U34 should document the no-standalone-income-statement-D&A zero policy for LLY 1Q26.");
  }
  if (q126DaRows.some((row) => /DepreciationDepletionAndAmortization/.test(row.concepts))) {
    errors.push("Mapping Audit U34 should not use cash-flow DepreciationDepletionAndAmortization for LLY income-statement D&A.");
  }

  const q126OtherOperatingRows = auditRowsFor(audit, "U35", "1Q26");
  if (!q126OtherOperatingRows.some((row) => /ResearchAndDevelopmentAssetAcquiredOtherThanThroughBusinessCombinationWrittenOff/.test(row.concepts))) {
    errors.push("Mapping Audit U35 should include acquired IPR&D in other operating income/expense for LLY 1Q26.");
  }
  if (!q126OtherOperatingRows.some((row) => /RestructuringSettlementAndImpairmentProvisions/.test(row.concepts))) {
    errors.push("Mapping Audit U35 should include restructuring / impairment provisions in other operating income/expense for LLY 1Q26.");
  }

  const q126OtherNonOperatingRows = auditRowsFor(audit, "U41", "1Q26");
  if (q126OtherNonOperatingRows.some((row) => /ResearchAndDevelopmentAssetAcquired|RestructuringSettlementAndImpairment/.test(row.concepts))) {
    errors.push("Mapping Audit U41 should not absorb LLY operating IPR&D or restructuring charges into other non-operating income/expense.");
  }

  if (errors.length) {
    console.error(errors.join("\n"));
    throw new Error(`LLY income statement classification regression failed with ${errors.length} issue(s).`);
  }

  console.log(`LLY income statement classification regression passed: ${outputWorkbook}`);
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
