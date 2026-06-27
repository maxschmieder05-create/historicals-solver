const path = require("node:path");
const fs = require("node:fs/promises");
const ExcelJS = require("exceljs");
const { postWorkbook } = require("./fill-workbook-api");

const repoRoot = path.resolve(__dirname, "..");
const inputWorkbook = process.env.MSFT_LIABILITY_INPUT_WORKBOOK || "/Users/maxschmieder/Downloads/MSFT_historicals_filled (1).xlsx";
const preparedInputWorkbook = path.join(repoRoot, "tmp", "msft-liability-classification-input.xlsx");
const outputWorkbook = process.env.MSFT_LIABILITY_OUTPUT_WORKBOOK || path.join(repoRoot, "tmp", "msft-liability-classification-output.xlsx");
const apiUrl = process.env.FILL_API_URL || "http://localhost:3000/api/fill-model";

const expectedCells = [
  ["Model", "T135", 20920, "FY25 accrued liabilities excluding long-term income taxes"],
  ["Model", "T136", 89575, "FY25 other current liabilities including short-term unearned revenue and other current liabilities"],
  ["Model", "T137", 138219, "FY25 current liabilities excluding debt"],
  ["Model", "U135", 12856, "1Q26 accrued liabilities excluding long-term income taxes"],
  ["Model", "U136", 81728, "1Q26 other current liabilities including short-term unearned revenue and other current liabilities"],
  ["Model", "U137", 127164, "1Q26 current liabilities excluding debt"],
  ["Model", "T140", 43151, "FY25 debt including current portion"],
  ["Model", "T141", 2835, "FY25 deferred income taxes"],
  ["Model", "T142", 91319, "FY25 other non-current liabilities excluding deferred taxes"],
  ["Model", "U142", 100051, "1Q26 other non-current liabilities including long-term income taxes and long-term unearned revenue"],
  ["Model", "T143", 137305, "FY25 total non-current liabilities including debt current portion"],
  ["Model", "T145", 275524, "FY25 total liabilities"],
  ["Model", "T158", 0, "FY25 balance sheet check"]
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
  await fs.mkdir(path.dirname(preparedInputWorkbook), { recursive: true });
  const input = new ExcelJS.Workbook();
  await input.xlsx.readFile(inputWorkbook);
  const segmentSheet = input.getWorksheet("Segment Analysis");
  if (segmentSheet) input.removeWorksheet(segmentSheet.id);
  await input.xlsx.writeFile(preparedInputWorkbook);

  await postWorkbook({ apiUrl, ticker: "MSFT", inputWorkbook: preparedInputWorkbook, outputWorkbook });

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
  const auditRows = [];
  if (audit) {
    for (let row = 1; row <= audit.rowCount; row += 1) {
      const cell = String(cellValue(audit.getCell(row, 2)) ?? "");
      if (cell === "T135" || cell === "T136" || cell === "T141" || cell === "T142" || cell === "U135" || cell === "U136" || cell === "U142") {
        auditRows.push({
          cell,
          label: String(cellValue(audit.getCell(row, 3)) ?? ""),
          concepts: String(cellValue(audit.getCell(row, 8)) ?? "")
        });
      }
    }
  }

  if (!auditRows.some((row) => row.cell === "T141" && /DeferredIncomeTaxLiabilitiesNet=2835mm/.test(row.concepts))) {
    errors.push("Model!T141 should map FY25 deferred income taxes to DeferredIncomeTaxLiabilitiesNet=2835mm.");
  }

  if (auditRows.some((row) => row.cell === "T142" && /DeferredIncomeTaxLiabilitiesNet=2835mm/.test(row.concepts) && !/Liabilities=275524mm/.test(row.concepts))) {
    errors.push("Model!T142 should not directly bury the deferred tax liability in other non-current liabilities.");
  }

  if (auditRows.some((row) => row.cell === "U135" && /AccruedIncomeTaxesNoncurrent=26569mm/.test(row.concepts))) {
    errors.push("Model!U135 should not include long-term income taxes in current accrued liabilities.");
  }

  if (!auditRows.some((row) => row.cell === "U136" && /ContractWithCustomerLiabilityCurrent=58987mm/.test(row.concepts) && /OtherLiabilitiesCurrent=22741mm/.test(row.concepts))) {
    errors.push("Model!U136 should group MSFT 1Q26 short-term unearned revenue and other current liabilities into Other Current Liabilities.");
  }

  if (!auditRows.some((row) => row.cell === "U142" && /AccruedIncomeTaxesNoncurrent=26569mm/.test(row.concepts) && /ContractWithCustomerLiabilityNoncurrent=2546mm/.test(row.concepts) && /OtherLiabilitiesNoncurrent=53588mm/.test(row.concepts))) {
    errors.push("Model!U142 should include MSFT 1Q26 long-term income taxes, long-term unearned revenue, and other long-term liabilities.");
  }

  if (errors.length) {
    console.error(errors.join("\n"));
    throw new Error(`MSFT liability classification regression failed with ${errors.length} issue(s).`);
  }

  console.log(`MSFT liability classification regression passed: ${outputWorkbook}`);
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
