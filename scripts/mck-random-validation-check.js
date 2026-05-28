const fs = require("node:fs/promises");
const path = require("node:path");
const ExcelJS = require("exceljs");
const { postWorkbook } = require("./fill-workbook-api");

const repoRoot = path.resolve(__dirname, "..");
const inputWorkbook =
  process.env.MCK_INPUT_WORKBOOK || "/Users/maxschmieder/Downloads/McKesson Corporation (MCK)_Valuation Workbook (12-May-2026).xlsx";
const outputWorkbook = process.env.MCK_OUTPUT_WORKBOOK || path.join(repoRoot, "tmp", "random-mck-validation-output.xlsx");
const apiUrl = process.env.FILL_API_URL || "http://localhost:3000/api/fill-model";
const ticker = process.env.MCK_TICKER || "MCK";

const expectedCells = [
  ["Model", "F28", 74483, "1Q24 revenue"],
  ["Model", "F29", -71461, "1Q24 COGS"],
  ["Model", "F36", 1100, "1Q24 EBIT"],
  ["Model", "F41", 38, "1Q24 other non-operating income / expense residual"],
  ["Model", "F42", 1091, "1Q24 pre-tax income"],
  ["Model", "F44", -133, "1Q24 tax expense"],
  ["Model", "F45", 958, "1Q24 GAAP net income"],
  ["Model", "F51", 920.1, "1Q24 adjusted net income after NCI row"],
  ["Model", "F61", 1062.1, "1Q24 EBITDA"],
  ["Model", "F132", 64096, "1Q24 total assets"],
  ["Model", "F145", 65336, "1Q24 total liabilities"],
  ["Model", "F154", -1240, "1Q24 shareholders' equity"],
  ["Model", "F158", 0, "1Q24 balance sheet check"],
  ["Model", "G45", 664, "2Q24 GAAP net income"],
  ["Model", "H45", 589, "3Q24 GAAP net income"],
  ["Model", "I45", 791, "4Q24 GAAP net income"],
  ["Model", "I158", 0, "4Q24 balance sheet check"]
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

  if (errors.length) {
    console.error(errors.join("\n"));
    throw new Error(`MCK random-company validation failed with ${errors.length} issue(s).`);
  }

  console.log(`MCK random-company validation passed: ${outputWorkbook}`);
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
