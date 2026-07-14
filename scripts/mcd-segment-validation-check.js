const path = require("node:path");
const ExcelJS = require("exceljs");
const { postWorkbook } = require("./fill-workbook-api");

const repoRoot = path.resolve(__dirname, "..");
const inputWorkbook =
  process.env.MCD_INPUT_WORKBOOK || "/Users/maxschmieder/Downloads/Owl Fund Integrated Model Template (03-Sep-2025)_v25 (3).xlsx";
const outputWorkbook = process.env.MCD_SEGMENT_OUTPUT_WORKBOOK || path.join(repoRoot, "tmp", "mcd-segment-validation-output.xlsx");
const apiUrl = process.env.FILL_API_URL || "http://localhost:3000/api/fill-model";

const expectedLabels = [
  ["C8", "Reported Revenue"],
  ["C16", "Reported"]
];

const revenueColumns = ["F", "G", "H", "I", "K", "L", "M", "N", "P", "Q", "R", "S", "U"];

function cellValue(cell) {
  const value = cell.value;
  if (typeof value === "number" || typeof value === "string" || value === null) return value;
  if (value && typeof value === "object") {
    if (typeof value.result === "number" || typeof value.result === "string") return value.result;
    if (typeof value.formula === "string") return value.result ?? `=${value.formula}`;
  }
  return value;
}

function numericCell(cell) {
  const value = cellValue(cell);
  return typeof value === "number" ? value : null;
}

function valuesMatch(actual, expected, tolerance = 1) {
  return typeof actual === "number" && Math.abs(actual - expected) <= tolerance;
}

async function fillWorkbook() {
  await postWorkbook({ apiUrl, ticker: "MCD", inputWorkbook, outputWorkbook });
}

async function main() {
  await fillWorkbook();

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(outputWorkbook);
  const segmentSheet = workbook.getWorksheet("Segment Analysis");
  const modelSheet = workbook.getWorksheet("Model");
  const errors = [];

  if (!segmentSheet || !modelSheet) {
    throw new Error("Workbook is missing Model or Segment Analysis sheet.");
  }

  for (const [address, expected] of expectedLabels) {
    const actual = String(cellValue(segmentSheet.getCell(address)) ?? "");
    if (actual !== expected) {
      errors.push(`Segment Analysis!${address}: expected label "${expected}", got "${actual}".`);
    }
  }

  for (let row = 8; row <= 13; row += 1) {
    const label = String(cellValue(segmentSheet.getCell(`C${row}`)) ?? "");
    if (/U\.S\. Market|International Operated Markets|International Developmental Licensed Markets/i.test(label)) {
      errors.push(`Segment Analysis!C${row} should not include geographic revenue label "${label}".`);
    }
  }

  for (const col of revenueColumns) {
    const segmentRevenue = [8, 9, 10, 11, 12, 13].reduce((sum, row) => sum + (numericCell(segmentSheet.getCell(`${col}${row}`)) ?? 0), 0);
    const segmentTotal = numericCell(segmentSheet.getCell(`${col}7`));
    const modelRevenue = numericCell(modelSheet.getCell(`${col}28`));
    if (!valuesMatch(segmentRevenue, segmentTotal ?? NaN)) {
      errors.push(`Segment Analysis!${col}7: segment rows sum to ${segmentRevenue}, but total row is ${segmentTotal ?? "[blank]"}.`);
    }
    if (!valuesMatch(segmentRevenue, modelRevenue ?? NaN)) {
      errors.push(`Segment Analysis ${col}: segment rows sum to ${segmentRevenue}, but Model!${col}28 revenue is ${modelRevenue ?? "[blank]"}.`);
    }
    const reportedRevenue = numericCell(segmentSheet.getCell(`${col}8`));
    if (!valuesMatch(reportedRevenue, modelRevenue ?? NaN)) {
      errors.push(`Segment Analysis!${col}8 Reported Revenue is ${reportedRevenue ?? "[blank]"}, but Model!${col}28 revenue is ${modelRevenue ?? "[blank]"}.`);
    }
  }

  if (errors.length) {
    console.error(errors.join("\n"));
    throw new Error(`MCD segment validation failed with ${errors.length} issue(s).`);
  }

  console.log(`MCD segment validation passed: ${outputWorkbook}`);
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
