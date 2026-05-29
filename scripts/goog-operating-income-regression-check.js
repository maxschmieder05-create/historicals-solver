const path = require("node:path");
const ExcelJS = require("exceljs");
const { postWorkbook } = require("./fill-workbook-api");

const repoRoot = path.resolve(__dirname, "..");
const inputWorkbook =
  process.env.GOOG_OPERATING_INPUT_WORKBOOK || "/Users/maxschmieder/Downloads/GOOG_historicals_filled (1).xlsx";
const outputWorkbook = process.env.GOOG_OPERATING_OUTPUT_WORKBOOK || path.join(repoRoot, "tmp", "goog-operating-income-regression-output.xlsx");
const apiUrl = process.env.FILL_API_URL || "http://localhost:3000/api/fill-model";
const ticker = process.env.GOOG_TICKER || "GOOG";

const expectedCells = [
  ["Model", "F35", 0, "1Q23 other operating income/expense should not double-count embedded restructuring charges"],
  ["Model", "F36", 17415, "1Q23 EBIT should equal reported operating income"],
  ["Model", "F41", 73, "1Q23 other non-operating income/expense should split the reported non-operating line after separately classified interest"],
  ["Model", "F42", 18205, "1Q23 pre-tax income"],
  ["Model", "F59", 17415, "1Q23 analysis EBIT should equal reported operating income"]
];

function cellValue(cell) {
  const value = cell.value;
  if (typeof value === "number") return value;
  if (value && typeof value === "object") {
    if (typeof value.result === "number") return value.result;
    if (typeof value.result === "string") return Number(value.result);
  }
  return value;
}

function valuesMatch(actual, expected) {
  return typeof actual === "number" && Math.abs(actual - expected) <= 0.5;
}

async function main() {
  await postWorkbook({ apiUrl, ticker, inputWorkbook, outputWorkbook });

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
      const period = String(cellValue(audit.getCell(row, 5)) ?? "");
      const value = cellValue(audit.getCell(row, 6));
      const concepts = String(cellValue(audit.getCell(row, 8)) ?? "");
      if (cell === "F35" && period === "1Q23" && value !== 0 && /RestructuringCharges/.test(concepts)) {
        errors.push("Model!F35 should not map GOOG 1Q23 RestructuringCharges into Other Operating Income (Expense) when operating income already reconciles.");
      }
      if (cell === "F41" && period === "1Q23" && /OtherNonOperatingIncomeExpenseResidual/.test(concepts)) {
        errors.push("Model!F41 should not use a pure pre-tax residual for GOOG 1Q23 other non-operating income (expense).");
      }
    }
  }

  if (errors.length) {
    console.error(errors.join("\n"));
    throw new Error(`GOOG operating-income regression failed with ${errors.length} issue(s).`);
  }

  console.log(`GOOG operating-income regression passed: ${outputWorkbook}`);
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
