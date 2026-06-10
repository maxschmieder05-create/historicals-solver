const path = require("node:path");
const ExcelJS = require("exceljs");
const { postWorkbook } = require("./fill-workbook-api");

const repoRoot = path.resolve(__dirname, "..");
const inputWorkbook = process.env.INTC_BALANCE_INPUT_WORKBOOK || "/Users/maxschmieder/Downloads/INTC_historicals_filled (1).xlsx";
const outputWorkbook =
  process.env.INTC_BALANCE_OUTPUT_WORKBOOK || path.join(repoRoot, "tmp", "intc-balance-sheet-formula-protection-output.xlsx");
const apiUrl = process.env.FILL_API_URL || "http://localhost:3000/api/fill-model";
const ticker = process.env.INTC_TICKER || "INTC";

const expected1Q26 = {
  "Cash & Cash Equivalents": 32789,
  "Prepaid & Other Current Assets": 12876,
  "Total Current Assets": 62157,
  "Total Current Liabilities (Excl. Debt)": 24881,
  Revolver: 0,
  "LT Debt (Incl. Current Portion)": 45031
};

function cellValue(cell) {
  const value = cell.value;
  if (typeof value === "number" || typeof value === "string") return value;
  if (value && typeof value === "object") return value.result ?? value.text ?? value.formula ?? value.sharedFormula ?? null;
  return value;
}

function cellFormula(cell) {
  return cell.formula ?? null;
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
    .replace(/[']/g, "")
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
      const period = normalizePeriodLabel(cellValue(sheet.getCell(row, col)));
      if (/^[1-4]Q\d{2}$/i.test(period)) {
        validCount += 1;
        quarterCount += 1;
      } else if (/^FY\d{2}$/i.test(period)) {
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
    if (normalizePeriodLabel(cellValue(sheet.getCell(headerRow, col))) === period.toUpperCase()) return col;
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

function snapshotFormulas(workbook) {
  const formulas = new Map();
  for (const sheetName of ["Model", "Segment Analysis"]) {
    const sheet = workbook.getWorksheet(sheetName);
    if (!sheet) continue;
    for (let row = 1; row <= sheet.rowCount; row += 1) {
      for (let col = 1; col <= sheet.columnCount; col += 1) {
        const formula = cellFormula(sheet.getCell(row, col));
        if (formula) formulas.set(`${sheetName}!${sheet.getCell(row, col).address}`, formula);
      }
    }
  }
  return formulas;
}

function compareFormulaSnapshots(before, afterWorkbook, errors) {
  for (const [address, expected] of before.entries()) {
    const [sheetName, cellAddress] = address.split("!");
    const sheet = afterWorkbook.getWorksheet(sheetName);
    const actual = sheet ? cellFormula(sheet.getCell(cellAddress)) : null;
    if (actual !== expected) errors.push(`${address}: formula changed from "${expected}" to "${actual ?? "[hardcoded/blank]"}".`);
  }
}

function compareProtectedCheckRow(inputModel, outputModel, errors) {
  const checkRow = findRow(inputModel, "Balance Sheet Check");
  if (!checkRow) {
    errors.push("Input workbook did not contain a Balance Sheet Check row.");
    return;
  }
  for (let col = 6; col <= Math.min(inputModel.columnCount, 40); col += 1) {
    const inputCell = inputModel.getCell(checkRow, col);
    const outputCell = outputModel.getCell(checkRow, col);
    const inputFormula = cellFormula(inputCell);
    const outputFormula = cellFormula(outputCell);
    if (inputFormula || outputFormula) {
      if (inputFormula !== outputFormula) {
        errors.push(`Balance Sheet Check ${inputCell.address}: formula changed from "${inputFormula ?? ""}" to "${outputFormula ?? ""}".`);
      }
      continue;
    }
    const before = cellValue(inputCell);
    const after = cellValue(outputCell);
    if (String(before ?? "") !== String(after ?? "")) {
      errors.push(`Balance Sheet Check ${inputCell.address}: value changed from "${before ?? ""}" to "${after ?? ""}".`);
    }
  }
}

function auditRowsFor(audit, cell, period) {
  const rows = [];
  for (let row = 2; row <= audit.rowCount; row += 1) {
    if (String(cellValue(audit.getCell(row, 2)) ?? "") !== cell) continue;
    if (String(cellValue(audit.getCell(row, 5)) ?? "").toUpperCase() !== period.toUpperCase()) continue;
    rows.push({
      concepts: String(cellValue(audit.getCell(row, 8)) ?? ""),
      labels: String(cellValue(audit.getCell(row, 9)) ?? ""),
      notes: String(cellValue(audit.getCell(row, 24)) ?? "")
    });
  }
  return rows;
}

async function readWorkbook(file) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(file);
  return workbook;
}

async function main() {
  const input = await readWorkbook(inputWorkbook);
  const inputModel = input.getWorksheet("Model");
  if (!inputModel) throw new Error("Input workbook does not contain a Model sheet.");
  const inputFormulas = snapshotFormulas(input);

  await postWorkbook({ apiUrl, ticker, inputWorkbook, outputWorkbook });

  const output = await readWorkbook(outputWorkbook);
  const model = output.getWorksheet("Model");
  const audit = output.getWorksheet("Mapping Audit");
  if (!model) throw new Error("Output workbook does not contain a Model sheet.");
  if (!audit) throw new Error("Output workbook does not contain a Mapping Audit sheet.");

  const errors = [];
  compareFormulaSnapshots(inputFormulas, output, errors);
  compareProtectedCheckRow(inputModel, model, errors);

  const col = findPeriodColumn(model, "1Q26");
  if (!col) {
    errors.push("Could not find 1Q26 column in Model sheet.");
  } else {
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
  }

  const cashRows = auditRowsFor(audit, "U120", "1Q26");
  if (!cashRows.some((row) => /CashAndCashEquivalentsAtCarryingValue=17247mm/.test(row.concepts) && /AvailableForSaleSecuritiesDebtSecuritiesCurrent=15542mm/.test(row.concepts))) {
    errors.push("Mapping Audit U120 should show Cash & Cash Equivalents includes cash and current short-term investments.");
  }

  const prepaidRows = auditRowsFor(audit, "U123", "1Q26");
  if (prepaidRows.some((row) => /AvailableForSaleSecuritiesDebtSecuritiesCurrent|ShortTermInvestments|MarketableSecuritiesCurrent/.test(row.concepts))) {
    errors.push("Mapping Audit U123 should not bury current investments inside Prepaid & Other Current Assets.");
  }

  const revolverRows = auditRowsFor(audit, "U139", "1Q26");
  if (!revolverRows.some((row) => /ShortTermBorrowings=0mm/.test(row.concepts))) {
    errors.push("Mapping Audit U139 should set Revolver to zero when INTC reports no standalone short-term borrowing/revolver balance.");
  }
  if (revolverRows.some((row) => /DebtCurrent=2004mm|LongTermDebtCurrent=/.test(row.concepts))) {
    errors.push("Mapping Audit U139 should not classify current maturities/current debt as Revolver.");
  }

  const debtRows = auditRowsFor(audit, "U140", "1Q26");
  if (!debtRows.some((row) => /LongTermDebtNoncurrent=43027mm/.test(row.concepts) && /DebtCurrent=2004mm/.test(row.concepts))) {
    errors.push("Mapping Audit U140 should include current debt in LT Debt (Incl. Current Portion).");
  }

  if (errors.length) {
    console.error(errors.join("\n"));
    throw new Error(`INTC balance-sheet/formula-protection regression failed with ${errors.length} issue(s).`);
  }

  console.log(`INTC balance-sheet/formula-protection regression passed: ${outputWorkbook}`);
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
