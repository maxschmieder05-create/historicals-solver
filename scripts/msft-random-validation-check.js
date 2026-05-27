const path = require("node:path");
const ExcelJS = require("exceljs");
const { postWorkbook } = require("./fill-workbook-api");

const repoRoot = path.resolve(__dirname, "..");
const inputWorkbook =
  process.env.MSFT_INPUT_WORKBOOK || "/Users/maxschmieder/Downloads/Owl Fund Integrated Model Template (03-Sep-2025)_v25 (3).xlsx";
const outputWorkbook = process.env.MSFT_OUTPUT_WORKBOOK || path.join(repoRoot, "tmp", "random-msft-validation-output.xlsx");
const apiUrl = process.env.FILL_API_URL || "http://localhost:3000/api/fill-model";
const ticker = process.env.MSFT_TICKER || "MSFT";

const expectedCells = [
  ["Model", "I158", 0, "4Q23 balance sheet check"],
  ["Model", "J158", 0, "FY23 balance sheet check"]
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

  if (errors.length) {
    console.error(errors.join("\n"));
    throw new Error(`MSFT random-company validation failed with ${errors.length} issue(s).`);
  }

  console.log(`MSFT random-company validation passed: ${outputWorkbook}`);
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
