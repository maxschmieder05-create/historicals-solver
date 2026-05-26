const fs = require("node:fs/promises");
const path = require("node:path");
const ExcelJS = require("exceljs");

const repoRoot = path.resolve(__dirname, "..");
const inputWorkbook =
  process.env.AAPL_INPUT_WORKBOOK || "/Users/maxschmieder/Downloads/Owl Fund Integrated Model Template (03-Sep-2025)_v25 (3).xlsx";
const outputWorkbook = process.env.AAPL_SEGMENT_OUTPUT_WORKBOOK || path.join(repoRoot, "tmp", "aapl-segment-validation-output.xlsx");
const apiUrl = process.env.FILL_API_URL || "http://localhost:3000/api/fill-model";

const expectedLabels = [
  ["C8", "Services Revenue"],
  ["C9", "iPhone Revenue"],
  ["C10", "Mac Revenue"],
  ["C11", "iPad Revenue"],
  ["C12", "Wearables, Home and Accessories Revenue"]
];

const revenueColumns = ["G", "H", "I", "K", "L", "N", "P", "Q", "S", "U"];

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

function valuesMatch(actual, expected) {
  return typeof actual === "number" && Math.abs(actual - expected) <= 0.5;
}

async function fillWorkbook() {
  await fs.mkdir(path.dirname(outputWorkbook), { recursive: true });
  const bytes = await fs.readFile(inputWorkbook);
  const formData = new FormData();
  formData.append("ticker", "AAPL");
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
  const segmentSheet = workbook.getWorksheet("Segment Analysis");
  const modelSheet = workbook.getWorksheet("Model");
  const errors = [];

  if (!segmentSheet || !modelSheet) throw new Error("Workbook is missing Model or Segment Analysis sheet.");

  for (const [address, expected] of expectedLabels) {
    const actual = String(cellValue(segmentSheet.getCell(address)) ?? "");
    if (actual !== expected) errors.push(`Segment Analysis!${address}: expected label "${expected}", got "${actual}".`);
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
  }

  for (const col of ["G", "H", "I", "K", "L", "M"]) {
    const goodwill = numericCell(modelSheet.getCell(`${col}128`));
    const check = numericCell(modelSheet.getCell(`${col}158`));
    if (!valuesMatch(goodwill ?? NaN, 0)) errors.push(`Model!${col}128 Goodwill should be cleared to 0, got ${goodwill ?? "[blank]"}.`);
    if (!valuesMatch(check ?? NaN, 0)) errors.push(`Model!${col}158 Balance Sheet Check should be 0, got ${check ?? "[blank]"}.`);
  }

  if (errors.length) {
    console.error(errors.join("\n"));
    throw new Error(`AAPL segment validation failed with ${errors.length} issue(s).`);
  }

  console.log(`AAPL segment validation passed: ${outputWorkbook}`);
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
