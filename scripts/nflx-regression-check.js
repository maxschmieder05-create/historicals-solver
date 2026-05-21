const fs = require("node:fs/promises");
const path = require("node:path");
const ExcelJS = require("exceljs");

const repoRoot = path.resolve(__dirname, "..");
const inputWorkbook =
  process.env.NFLX_INPUT_WORKBOOK || "/Users/maxschmieder/Downloads/Netflix Inc. (NFLX)_Valuation Workbook (15-Apr-2026) (2).xlsx";
const outputWorkbook = process.env.NFLX_OUTPUT_WORKBOOK || path.join(repoRoot, "tmp", "nflx-regression-output.xlsx");
const apiUrl = process.env.FILL_API_URL || "http://localhost:3000/api/fill-model";
const ticker = process.env.NFLX_TICKER || "NFLX";

const expectedCells = [
  ["Model", "F28", 8161.5, "1Q23 revenue"],
  ["Model", "F29", -4803.625, "1Q23 cost of revenue"],
  ["Model", "F32", -956.3, "1Q23 SG&A from marketing plus G&A"],
  ["Model", "F33", -687.275, "1Q23 technology/R&D expense"],
  ["Model", "F44", -163.754, "1Q23 income tax expense"],
  ["Model", "F120", 6714.594, "1Q23 cash and cash equivalents"],
  ["Model", "F129", 5245.444, "1Q23 other non-current assets"],
  ["Model", "F132", 49490.3, "1Q23 total assets"],
  ["Model", "F134", 591.987, "1Q23 accounts payable"],
  ["Model", "F135", 3379.503, "1Q23 accrued/current-liability residual"],
  ["Model", "F136", 4344.58, "1Q23 current content liabilities"],
  ["Model", "F137", 8316.1, "1Q23 total current liabilities"],
  ["Model", "F145", 27662.1, "1Q23 total liabilities"],
  ["Model", "F156", 49490.3, "1Q23 liabilities and equity"]
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
  if (!response.ok) {
    throw new Error(`Fill API failed (${response.status}): ${body.toString("utf8")}`);
  }
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

  const audit = workbook.getWorksheet("Mapping Audit");
  if (audit) {
    for (let row = 1; row <= audit.rowCount; row += 1) {
      const cell = String(cellValue(audit.getCell(row, 2)) ?? "");
      const mappingType = String(cellValue(audit.getCell(row, 7)) ?? "");
      const concepts = String(cellValue(audit.getCell(row, 8)) ?? "");
      if (cell === "F129" && mappingType === "calculated") {
        errors.push("Model!F129 was overwritten by a balance-sheet residual instead of retaining EDGAR OtherAssetsNoncurrent.");
      }
      if (cell === "F136" && !concepts.includes("ContentLiabilitiesCurrent")) {
        errors.push("Model!F136 did not map to EDGAR ContentLiabilitiesCurrent for Netflix current content liabilities.");
      }
    }
  }

  if (errors.length) {
    console.error(errors.join("\n"));
    throw new Error(`NFLX regression failed with ${errors.length} issue(s).`);
  }

  console.log(`NFLX regression passed: ${outputWorkbook}`);
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
