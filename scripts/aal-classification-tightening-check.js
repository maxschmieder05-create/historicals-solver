const path = require("node:path");
const ExcelJS = require("exceljs");
const { postWorkbook } = require("./fill-workbook-api");

const repoRoot = path.resolve(__dirname, "..");
const inputWorkbook = process.env.AAL_INPUT_WORKBOOK || "/Users/maxschmieder/Downloads/AAL_historicals_filled.xlsx";
const outputWorkbook = process.env.AAL_CLASSIFICATION_OUTPUT_WORKBOOK || path.join(repoRoot, "tmp", "aal-classification-tightening-output.xlsx");
const apiUrl = process.env.FILL_API_URL || "http://localhost:3000/api/fill-model";
const ticker = process.env.AAL_TICKER || "AAL";

const expected1Q26 = {
  "Cost of Goods Sold": -16292,
  "Selling, General & Administration (SG&A)": -507,
  "Depreciation & Amortization": -475,
  "Other Operating Income (Expense)": -14,
  EBIT: -41,
  "Interest Income": 55,
  "Interest (Expense)": -397,
  "Other Non-Operating Income (Expense)": -93,
  "Pre-Tax Income (Loss)": -476,
  "Cash & Cash Equivalents": 7393,
  Inventory: 3131,
  "Prepaid & Other Current Assets": 1452,
  "Total Current Assets": 13984,
  Revolver: 0,
  "LT Debt (Incl. Current Portion)": 27305,
  "Other Non-Current Liabilities": 16187,
  "Total Non-Current Liabilities": 43492,
  "Total Liabilities": 67815
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
    for (let col = 1; col <= Math.min(sheet.columnCount, 180); col += 1) {
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
  for (let col = 1; col <= Math.min(sheet.columnCount, 180); col += 1) {
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
      value: cellValue(row.getCell(6)),
      mappingType: String(cellValue(row.getCell(7)) ?? ""),
      concepts: String(cellValue(row.getCell(8)) ?? ""),
      labels: String(cellValue(row.getCell(9)) ?? ""),
      notes: String(cellValue(row.getCell(24)) ?? "")
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
  const col = findPeriodColumn(model, "1Q26");
  if (!col) throw new Error("Could not find 1Q26 column in Model sheet.");

  for (const [label, expected] of Object.entries(expected1Q26)) {
    const row = findRow(model, label);
    if (!row) {
      errors.push(`Could not find row "${label}" in Model sheet.`);
      continue;
    }
    const actual = cellValue(model.getCell(row, col));
    if (!valuesMatch(actual, expected)) {
      errors.push(`1Q26 ${label}: expected ${expected}, got ${actual ?? "[blank]"}.`);
    }
  }

  const cogsRows = auditRowsFor(audit, "U29", "1Q26");
  if (!cogsRows.some((row) => row.mappingType === "formula preserved" && Number(row.value) === -16292)) {
    errors.push("Mapping Audit U29 should preserve the existing COGS formula cell instead of overwriting it.");
  }
  if (cogsRows.some((row) => /CostsAndExpenses=13953mm/.test(row.concepts))) {
    errors.push("Mapping Audit U29 should not use total operating expenses as COGS when expense details are mapped separately.");
  }

  const otherNonOpRows = auditRowsFor(audit, "U41", "1Q26");
  if (!otherNonOpRows.some((row) => /OtherNonoperatingIncomeExpense=-93mm/.test(row.concepts))) {
    errors.push("Mapping Audit U41 should map AAL Other expense, net directly to Other Non-Operating Income (Expense).");
  }
  if (otherNonOpRows.some((row) => /OtherNonOperatingIncomeExpenseFromReportedLine|NonoperatingIncomeExpense=-435mm|InterestExpenseNonoperating=397mm/.test(row.concepts))) {
    errors.push("Mapping Audit U41 should not use a non-operating residual, subtotal, or interest line for AAL other non-operating.");
  }

  const inventoryRows = auditRowsFor(audit, "U122", "1Q26");
  if (!inventoryRows.some((row) => /AirlineRelatedInventoryNet=3131mm/.test(row.concepts))) {
    errors.push("Mapping Audit U122 should classify Airline Related Inventory, Net into Inventory.");
  }

  const revolverRows = auditRowsFor(audit, "U139", "1Q26");
  if (!revolverRows.some((row) => /ShortTermBorrowings=0mm/.test(row.concepts))) {
    errors.push("Mapping Audit U139 should set Revolver to zero when no true short-term borrowing/revolver is reported.");
  }
  if (revolverRows.some((row) => /LongTermDebtCurrent=4083mm/.test(row.concepts))) {
    errors.push("Mapping Audit U139 should not classify current maturities of long-term debt as Revolver.");
  }

  const longTermDebtRows = auditRowsFor(audit, "U140", "1Q26");
  if (!longTermDebtRows.some((row) => /DebtLongtermAndShorttermCombinedAmount=27305mm/.test(row.concepts))) {
    errors.push("Mapping Audit U140 should use EDGAR long-term and short-term debt combined amount for LT Debt (Incl. Current Portion).");
  }

  if (errors.length) {
    console.error(errors.join("\n"));
    throw new Error(`AAL classification tightening regression failed with ${errors.length} issue(s).`);
  }

  console.log(`AAL classification tightening regression passed: ${outputWorkbook}`);
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
