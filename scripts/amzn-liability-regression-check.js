const fs = require("node:fs/promises");
const path = require("node:path");
const ExcelJS = require("exceljs");
const { postWorkbook } = require("./fill-workbook-api");

const repoRoot = path.resolve(__dirname, "..");
const inputWorkbook = process.env.AMZN_INPUT_WORKBOOK || "/Users/maxschmieder/Downloads/AMZN_historicals_filled (4).xlsx";
const outputWorkbook = process.env.AMZN_LIABILITY_OUTPUT_WORKBOOK || path.join(repoRoot, "tmp", "amzn-liability-regression-output.xlsx");
const apiUrl = process.env.FILL_API_URL || "http://localhost:3000/api/fill-model";

const expectedCells = [
  ["Model", "F135", 59606, "1Q23 accrued liabilities excluding current debt"],
  ["Model", "F139", 1100, "1Q23 short-term borrowings / revolver"],
  ["Model", "F140", 72760, "1Q23 LT debt including current maturities"],
  ["Model", "F142", 95198, "1Q23 other non-current liabilities"],
  ["Model", "F145", 309852, "1Q23 total liabilities"],
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
  await postWorkbook({ apiUrl, ticker: "AMZN", inputWorkbook, outputWorkbook });
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
    for (let row = 1; row <= audit.rowCount; row += 1) {
      const cell = String(cellValue(audit.getCell(row, 2)) ?? "");
      const concepts = String(cellValue(audit.getCell(row, 8)) ?? "");
      if (cell === "F135" && !/(?:LongTermDebtCurrent|FinanceLeaseLiabilityCurrent|ShortTermBorrowings)=/.test(concepts)) {
        errors.push("Model!F135 should exclude separately disclosed current debt from current liabilities excluding debt.");
      }
      if (cell === "F139" && !/ShortTermBorrowings=1100mm/.test(concepts)) {
        errors.push("Model!F139 should map short-term borrowings to the revolver / short-term debt row.");
      }
      if (cell === "F140" && !/LongTermDebtCurrent=2000mm/.test(concepts)) {
        errors.push("Model!F140 should include current maturities of long-term debt.");
      }
      if (cell === "F140" && /ShortTermBorrowings=1100mm/.test(concepts)) {
        errors.push("Model!F140 should not absorb short-term borrowings when a revolver / short-term debt row is available.");
      }
    }
  }

  if (errors.length) {
    console.error(errors.join("\n"));
    throw new Error(`AMZN liability regression failed with ${errors.length} issue(s).`);
  }

  console.log(`AMZN liability regression passed: ${outputWorkbook}`);
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
