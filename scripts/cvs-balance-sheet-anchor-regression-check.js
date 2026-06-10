const path = require("node:path");
const ExcelJS = require("exceljs");
const { postWorkbook } = require("./fill-workbook-api");

const repoRoot = path.resolve(__dirname, "..");
const inputWorkbook = process.env.CVS_INPUT_WORKBOOK || "/Users/maxschmieder/Downloads/CVS_historicals_filled.xlsx";
const outputWorkbook = process.env.CVS_BALANCE_SHEET_OUTPUT_WORKBOOK || path.join(repoRoot, "tmp", "cvs-balance-sheet-anchor-output.xlsx");
const apiUrl = process.env.FILL_API_URL || "http://localhost:3000/api/fill-model";

const expected1Q26 = {
  "Cash & Cash Equivalents": 11802,
  "Total Current Assets": 74817,
  "Other Non-Current Assets": 54572,
  "Total Assets": 252974,
  "Total Current Liabilities (Excl. Debt)": 83826,
  "LT Debt (Incl. Current Portion)": 63111,
  "Other Non-Current Liabilities": 24629,
  "Total Non-Current Liabilities": 91511,
  "Total Liabilities": 175337,
  "Total Shareholder's Equity": 77456,
  "Noncontrolling Interests": 181,
  "Total Equity": 77637,
  "Total Liabilities & Shareholder's Equity": 252974,
  "Balance Sheet Check": 0
};

const failedRegressionValues = {
  "Total Current Assets": 70161,
  "Total Assets": 199071,
  "Total Liabilities": 121172,
  "Total Liabilities & Shareholder's Equity": 198597
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
    const score = validCount + quarterCount * 3;
    if (score && (!best || score > best.score)) best = { row, score };
  }
  return best?.row ?? null;
}

function findPeriodColumn(sheet, period) {
  const headerRow = bestPeriodHeaderRow(sheet);
  if (!headerRow) return null;
  for (let col = 1; col <= Math.min(sheet.columnCount, 160); col += 1) {
    if (normalizePeriodLabel(cellValue(sheet.getCell(headerRow, col))) === period.toUpperCase()) return col;
  }
  return null;
}

function rowLabel(sheet, row) {
  for (const col of [1, 2, 3, 4, 5]) {
    const text = String(cellValue(sheet.getCell(row, col)) ?? "").trim();
    if (text && !/^x$/i.test(text)) return text;
  }
  return "";
}

function findPrimaryBalanceSheetRow(sheet, label) {
  const wanted = normalize(label);
  for (let row = 80; row <= Math.min(sheet.rowCount, 180); row += 1) {
    if (normalize(rowLabel(sheet, row)) === wanted) return row;
  }
  return null;
}

function findRow(sheet, label) {
  const wanted = normalize(label);
  for (let row = 1; row <= sheet.rowCount; row += 1) {
    if (normalize(rowLabel(sheet, row)) === wanted) return row;
  }
  return null;
}

function findRowAfter(sheet, startRow, label) {
  if (!startRow) return null;
  const wanted = normalize(label);
  for (let row = startRow + 1; row <= sheet.rowCount; row += 1) {
    if (normalize(rowLabel(sheet, row)) === wanted) return row;
  }
  return null;
}

function valuesMatch(actual, expected) {
  return typeof actual === "number" && Math.abs(actual - expected) <= 0.5;
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
      notes: String(cellValue(row.getCell(24)) ?? ""),
      status: String(cellValue(row.getCell(23)) ?? "")
    });
  });
  return rows;
}

async function main() {
  await postWorkbook({ apiUrl, ticker: "CVS", inputWorkbook, outputWorkbook });

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(outputWorkbook);
  const model = workbook.getWorksheet("Model");
  const audit = workbook.getWorksheet("Mapping Audit");
  if (!model) throw new Error("Output workbook does not contain a Model sheet.");
  if (!audit) throw new Error("Output workbook does not contain a Mapping Audit sheet.");

  const col = findPeriodColumn(model, "1Q26");
  if (!col) throw new Error("Could not find 1Q26 column in Model sheet.");

  const errors = [];
  for (const [label, expected] of Object.entries(expected1Q26)) {
    const row = findPrimaryBalanceSheetRow(model, label);
    if (!row) {
      errors.push(`Could not find primary balance-sheet row "${label}".`);
      continue;
    }
    const actual = cellValue(model.getCell(row, col));
    if (!valuesMatch(actual === 1e-12 ? 0 : actual, expected)) {
      errors.push(`1Q26 ${label}: expected ${expected}, got ${actual ?? "[blank]"}.`);
    }
  }

  for (const [label, failedValue] of Object.entries(failedRegressionValues)) {
    const row = findPrimaryBalanceSheetRow(model, label);
    if (!row) continue;
    const actual = cellValue(model.getCell(row, col));
    if (valuesMatch(actual, failedValue)) errors.push(`1Q26 ${label} still equals failed regression value ${failedValue}.`);
  }

  const cashRow = findPrimaryBalanceSheetRow(model, "Cash & Cash Equivalents");
  if (!cashRow) errors.push("Could not find primary Cash & Cash Equivalents row for audit validation.");
  const cashAudit = cashRow ? auditRowsFor(audit, model.getCell(cashRow, col).address, "1Q26") : [];
  if (!cashAudit.some((row) => /CashAndCashEquivalentsAtCarryingValue=9542mm/.test(row.concepts) && /ShortTermInvestments=2260mm/.test(row.concepts))) {
    errors.push("1Q26 Cash & Cash Equivalents should include EDGAR cash and short-term investments.");
  }

  const otherAssetRow = findPrimaryBalanceSheetRow(model, "Other Non-Current Assets");
  if (!otherAssetRow) errors.push("Could not find primary Other Non-Current Assets row for audit validation.");
  const otherAssetAudit = otherAssetRow ? auditRowsFor(audit, model.getCell(otherAssetRow, col).address, "1Q26") : [];
  if (!otherAssetAudit.some((row) => /LongTermInvestments=32407mm/.test(row.concepts) && /OperatingLeaseRightOfUseAsset=14741mm/.test(row.concepts) && /OtherAssetsNoncurrent=7424mm/.test(row.concepts))) {
    errors.push("1Q26 Other Non-Current Assets should include long-term investments, operating lease ROU assets, and other assets.");
  }

  const debtRow = findPrimaryBalanceSheetRow(model, "LT Debt (Incl. Current Portion)");
  if (!debtRow) errors.push("Could not find primary LT Debt (Incl. Current Portion) row for audit validation.");
  const debtAudit = debtRow ? auditRowsFor(audit, model.getCell(debtRow, col).address, "1Q26") : [];
  if (!debtAudit.some((row) => /LongTermDebtAndCapitalLeaseObligations=60531mm/.test(row.concepts) && /LongTermDebtAndCapitalLeaseObligationsCurrent=2580mm/.test(row.concepts))) {
    errors.push("1Q26 LT Debt (Incl. Current Portion) should include long-term debt and current portion of long-term debt.");
  }

  if (errors.length) {
    console.error(errors.join("\n"));
    throw new Error(`CVS balance sheet anchor regression failed with ${errors.length} issue(s).`);
  }

  console.log(`CVS balance sheet anchor regression passed: ${outputWorkbook}`);
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
