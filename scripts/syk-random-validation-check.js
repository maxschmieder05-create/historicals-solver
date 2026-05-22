const fs = require("node:fs/promises");
const path = require("node:path");
const ExcelJS = require("exceljs");

const repoRoot = path.resolve(__dirname, "..");
const inputWorkbook =
  process.env.SYK_INPUT_WORKBOOK || "/Users/maxschmieder/Downloads/Stryker Corp (SYK)_Valuation Workbook (13-Mar-2026).xlsx";
const outputWorkbook = process.env.SYK_OUTPUT_WORKBOOK || path.join(repoRoot, "tmp", "random-syk-validation-output.xlsx");
const apiUrl = process.env.FILL_API_URL || "http://localhost:3000/api/fill-model";
const ticker = process.env.SYK_TICKER || "SYK";

const expectedCells = [
  ["Model", "F28", 4778, "1Q23 revenue"],
  ["Model", "F29", -1762, "1Q23 COGS / cost of revenue"],
  ["Model", "F36", 735, "1Q23 EBIT / operating income"],
  ["Model", "F42", 679, "1Q23 pre-tax income"],
  ["Model", "F44", -87, "1Q23 tax expense"],
  ["Model", "F45", 592, "1Q23 net income"],
  ["Model", "F59", 735, "1Q23 EBITDA"],
  ["Model", "F123", 850, "1Q23 prepaid and other current assets"],
  ["Model", "F124", 10155, "1Q23 total current assets"],
  ["Model", "F129", 4010, "1Q23 other non-current assets"],
  ["Model", "F132", 36830, "1Q23 total assets"],
  ["Model", "F137", 4662, "1Q23 modeled current liabilities excluding current debt carried in total debt"],
  ["Model", "F140", 13061, "1Q23 debt including current portion"],
  ["Model", "F142", 2212, "1Q23 other non-current liabilities"],
  ["Model", "F143", 15273, "1Q23 modeled non-current liabilities including total debt"],
  ["Model", "F145", 19935, "1Q23 total liabilities"],
  ["Model", "F154", 16895, "1Q23 shareholders' equity"],
  ["Model", "F156", 36830, "1Q23 liabilities and equity"],
  ["Model", "F158", 0, "1Q23 balance sheet check"],
  ["Model", "M123", 961, "3Q24 prepaid and other current assets excluding other short-term investments"],
  ["Model", "M158", 0, "3Q24 balance sheet check"],
  ["Model", "N123", 1593, "4Q24 prepaid and other current assets excluding other short-term investments"],
  ["Model", "N158", 0, "4Q24 balance sheet check"]
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
  await fs.mkdir(path.dirname(outputWorkbook), { recursive: true });
  const bytes = await fs.readFile(inputWorkbook);
  const formData = new FormData();
  formData.append("ticker", ticker);
  formData.append("file", new Blob([bytes]), path.basename(inputWorkbook));

  const response = await fetch(apiUrl, { method: "POST", body: formData });
  const body = Buffer.from(await response.arrayBuffer());
  if (!response.ok) throw new Error(`Fill API failed (${response.status}): ${body.toString("utf8")}`);
  await fs.writeFile(outputWorkbook, body);
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
    throw new Error(`SYK random-company validation failed with ${errors.length} issue(s).`);
  }

  console.log(`SYK random-company validation passed: ${outputWorkbook}`);
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
