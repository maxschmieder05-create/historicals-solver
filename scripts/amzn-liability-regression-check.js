const fs = require("node:fs/promises");
const path = require("node:path");
const ExcelJS = require("exceljs");

const repoRoot = path.resolve(__dirname, "..");
const inputWorkbook = process.env.AMZN_INPUT_WORKBOOK || "/Users/maxschmieder/Downloads/AMZN_historicals_filled (4).xlsx";
const outputWorkbook = process.env.AMZN_LIABILITY_OUTPUT_WORKBOOK || path.join(repoRoot, "tmp", "amzn-liability-regression-output.xlsx");
const apiUrl = process.env.FILL_API_URL || "http://localhost:3000/api/fill-model";

const expectedCells = [
  ["Model", "F135", 66382, "1Q23 accrued liabilities"],
  ["Model", "F140", 67084, "1Q23 LT debt"],
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
  await fs.mkdir(path.dirname(outputWorkbook), { recursive: true });
  const bytes = await fs.readFile(inputWorkbook);
  const formData = new FormData();
  formData.append("ticker", "AMZN");
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

  const audit = workbook.getWorksheet("Mapping Audit");
  if (audit) {
    for (let row = 1; row <= audit.rowCount; row += 1) {
      const cell = String(cellValue(audit.getCell(row, 2)) ?? "");
      const concepts = String(cellValue(audit.getCell(row, 8)) ?? "");
      if (cell === "F135" && /(?:LongTermDebtCurrent|ShortTermBorrowings)=(?!0(?:\.0+)?mm)/.test(concepts)) {
        errors.push("Model!F135 should not be reduced by debt concepts when current debt is embedded in accrued expenses and other.");
      }
      if (cell === "F140" && /LongTermDebtCurrent|ShortTermBorrowings/.test(concepts)) {
        errors.push("Model!F140 should map to non-current long-term debt only.");
      }
      if (cell === "F142" && /LongTermDebtCurrent|ShortTermBorrowings/.test(concepts)) {
        errors.push("Model!F142 should not use current debt concepts in the other non-current liability residual.");
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
