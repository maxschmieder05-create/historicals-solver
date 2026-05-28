const path = require("node:path");
const ExcelJS = require("exceljs");
const { postWorkbook } = require("./fill-workbook-api");

const repoRoot = path.resolve(__dirname, "..");
const inputWorkbook = process.env.COST_INPUT_WORKBOOK || "/Users/maxschmieder/Downloads/COST_historicals_filled.xlsx";
const outputWorkbook = process.env.COST_OUTPUT_WORKBOOK || path.join(repoRoot, "tmp", "cost-balance-sheet-regression-output.xlsx");
const apiUrl = process.env.FILL_API_URL || "http://localhost:3000/api/fill-model";
const ticker = process.env.COST_TICKER || "COST";

const expectedCells = [
  ["Model", "G120", 12970, "2Q23 cash and cash equivalents"],
  ["Model", "G121", 2714, "2Q23 receivables"],
  ["Model", "G122", 16081, "2Q23 inventory"],
  ["Model", "G132", 66848, "2Q23 total assets"],
  ["Model", "G148", 17341, "2Q23 retained earnings"],
  ["Model", "G150", -1672, "2Q23 AOCI"],
  ["Model", "G158", 0, "2Q23 balance sheet check"],
  ["Model", "L120", 9095, "2Q24 cash and cash equivalents"],
  ["Model", "L121", 2779, "2Q24 receivables"],
  ["Model", "L122", 17075, "2Q24 inventory"],
  ["Model", "L132", 66323, "2Q24 total assets"],
  ["Model", "L148", 14980, "2Q24 retained earnings"],
  ["Model", "L150", -1842, "2Q24 AOCI"],
  ["Model", "L158", 0, "2Q24 balance sheet check"],
  ["Model", "Q120", 12356, "2Q25 cash and cash equivalents"],
  ["Model", "Q121", 3060, "2Q25 receivables"],
  ["Model", "Q122", 18754, "2Q25 inventory"],
  ["Model", "Q132", 73224, "2Q25 total assets"],
  ["Model", "Q148", 19770, "2Q25 retained earnings"],
  ["Model", "Q150", -2242, "2Q25 AOCI"],
  ["Model", "Q158", 0, "2Q25 balance sheet check"]
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

  if (errors.length) {
    console.error(errors.join("\n"));
    throw new Error(`COST balance-sheet regression failed with ${errors.length} issue(s).`);
  }

  console.log(`COST balance-sheet regression passed: ${outputWorkbook}`);
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
