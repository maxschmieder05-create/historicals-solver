const fs = require("node:fs/promises");
const path = require("node:path");
const ExcelJS = require("exceljs");
const { postWorkbook } = require("./fill-workbook-api");

const repoRoot = path.resolve(__dirname, "..");
const inputWorkbook =
  process.env.FINANCIAL_COMPANY_INPUT_WORKBOOK ||
  "/Users/maxschmieder/Downloads/American Express Company (AXP)_Valuation Workbook (05-May-2026).xlsx";
const outputWorkbook = process.env.FINANCIAL_COMPANY_OUTPUT_WORKBOOK || path.join(repoRoot, "tmp", "financial-company-regression-output.xlsx");
const apiUrl = process.env.FILL_API_URL || "http://localhost:3000/api/fill-model";
const ticker = process.env.FINANCIAL_COMPANY_TICKER || "AXP";

const expectedCells = [
  ["Model", "F28", 8846, "1Q23 total revenue"],
  ["Model", "F29", -1433, "1Q23 interest expense"],
  ["Model", "F30", 7413, "1Q23 net revenue"],
  ["Model", "F31", -1055, "1Q23 provision for credit losses"],
  ["Model", "F39", 2167, "1Q23 operating/pre-tax income"],
  ["Model", "F42", 2167, "1Q23 pre-tax income"],
  ["Model", "F44", -351, "1Q23 income tax expense"],
  ["Model", "F45", 1816, "1Q23 net income"],
  ["Model", "F132", 236000, "1Q23 total assets"],
  ["Model", "F145", 210008, "1Q23 total liabilities"],
  ["Model", "F158", 0, "1Q23 balance sheet check"]
];

function cellValue(cell) {
  const value = cell.value;
  if (typeof value === "number") return value;
  if (value && typeof value === "object") {
    if (typeof value.result === "number") return value.result;
    if (typeof value.result === "string") return value.result;
  }
  return value;
}

function valuesMatch(actual, expected) {
  return typeof actual === "number" && Math.abs(actual - expected) <= 0.5;
}

async function fillWorkbook() {
  await postWorkbook({ apiUrl, ticker, inputWorkbook, outputWorkbook });
}

async function main() {
  await fillWorkbook();

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(outputWorkbook);
  const errors = [];

  for (const [sheetName, address, expected, label] of expectedCells) {
    const sheet = workbook.getWorksheet(sheetName);
    const actual = sheet ? cellValue(sheet.getCell(address)) : null;
    if (!valuesMatch(actual, expected)) {
      errors.push(`${label} ${sheetName}!${address}: expected ${expected}, got ${actual ?? "[blank]"}.`);
    }
  }

  const audit = workbook.getWorksheet("Mapping Audit");
  if (audit) {
    let hasOperatingProfitResolution = false;
    for (let row = 1; row <= audit.rowCount; row += 1) {
      const cell = String(cellValue(audit.getCell(row, 2)) ?? "");
      const concepts = String(cellValue(audit.getCell(row, 8)) ?? "");
      if (cell === "F38" && /ModeledOperatingProfitResolved|IncomeLossFromContinuingOperationsBeforeIncomeTaxes/.test(concepts)) {
        hasOperatingProfitResolution = true;
      }
      if (cell === "F44" && /NetIncomeLoss|ProfitLoss/.test(concepts)) {
        errors.push("Model!F44 should map to reported income tax expense, not become a net-income residual plug.");
      }
    }
    if (!hasOperatingProfitResolution) {
      errors.push("Model!F38 should carry the operating residual that reconciles financial-company operating/pre-tax income to EDGAR pre-tax income when no operating income concept is reported.");
    }
  }

  if (errors.length) {
    console.error(errors.join("\n"));
    throw new Error(`Financial company regression failed with ${errors.length} issue(s).`);
  }

  console.log(`Financial company regression passed: ${outputWorkbook}`);
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
