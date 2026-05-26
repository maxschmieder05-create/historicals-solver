const fs = require("node:fs/promises");
const path = require("node:path");
const ExcelJS = require("exceljs");

const repoRoot = path.resolve(__dirname, "..");
const inputWorkbook =
  process.env.MCD_INPUT_WORKBOOK || "/Users/maxschmieder/Downloads/Owl Fund Integrated Model Template (03-Sep-2025)_v25 (3).xlsx";
const outputWorkbook = process.env.MCD_SEGMENT_OUTPUT_WORKBOOK || path.join(repoRoot, "tmp", "mcd-segment-validation-output.xlsx");
const apiUrl = process.env.FILL_API_URL || "http://localhost:3000/api/fill-model";

const expectedLabels = [
  ["C8", "U.S. Market Revenue"],
  ["C9", "International Operated Markets Revenue"],
  ["C10", "International Developmental Licensed Markets and Corporate Revenue"],
  ["C16", "U.S. Market"],
  ["C17", "International Operated Markets"],
  ["C18", "International Developmental Licensed Markets and Corporate"]
];

const expectedRevenueCells = [
  ["F8", 2487.6, "1Q23 U.S. Market revenue"],
  ["F9", 2794.8, "1Q23 International Operated Markets revenue"],
  ["F10", 615.6, "1Q23 IDL Markets and Corporate revenue"],
  ["I8", 2675.6, "4Q23 U.S. Market revenue"],
  ["I9", 3131.1, "4Q23 International Operated Markets revenue"],
  ["I10", 599.3, "4Q23 IDL Markets and Corporate revenue"],
  ["S8", 2778, "4Q25 U.S. Market revenue"],
  ["S9", 3597, "4Q25 International Operated Markets revenue"],
  ["S10", 634, "4Q25 IDL Markets and Corporate revenue"]
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

function valuesMatch(actual, expected) {
  return typeof actual === "number" && Math.abs(actual - expected) <= 0.5;
}

async function fillWorkbook() {
  await fs.mkdir(path.dirname(outputWorkbook), { recursive: true });
  const bytes = await fs.readFile(inputWorkbook);
  const formData = new FormData();
  formData.append("ticker", "MCD");
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

  if (!segmentSheet || !modelSheet) {
    throw new Error("Workbook is missing Model or Segment Analysis sheet.");
  }

  for (const [address, expected] of expectedLabels) {
    const actual = String(cellValue(segmentSheet.getCell(address)) ?? "");
    if (actual !== expected) {
      errors.push(`Segment Analysis!${address}: expected label "${expected}", got "${actual}".`);
    }
  }

  for (const [address, expected, label] of expectedRevenueCells) {
    const actual = numericCell(segmentSheet.getCell(address));
    if (!valuesMatch(actual, expected)) {
      errors.push(`${label} Segment Analysis!${address}: expected ${expected}, got ${actual ?? "[blank]"}.`);
    }
  }

  for (const col of revenueColumns) {
    const segmentRevenue =
      (numericCell(segmentSheet.getCell(`${col}8`)) ?? 0) +
      (numericCell(segmentSheet.getCell(`${col}9`)) ?? 0) +
      (numericCell(segmentSheet.getCell(`${col}10`)) ?? 0);
    const segmentTotal = numericCell(segmentSheet.getCell(`${col}7`));
    const modelRevenue = numericCell(modelSheet.getCell(`${col}28`));
    if (!valuesMatch(segmentRevenue, segmentTotal ?? NaN)) {
      errors.push(`Segment Analysis!${col}7: segment rows sum to ${segmentRevenue}, but total row is ${segmentTotal ?? "[blank]"}.`);
    }
    if (!valuesMatch(segmentRevenue, modelRevenue ?? NaN)) {
      errors.push(`Segment Analysis ${col}: segment rows sum to ${segmentRevenue}, but Model!${col}28 revenue is ${modelRevenue ?? "[blank]"}.`);
    }
  }

  const fourthQuarterUs = segmentSheet.getCell("I8").value;
  if (!fourthQuarterUs || typeof fourthQuarterUs !== "object" || !String(fourthQuarterUs.formula ?? "").startsWith("10568.4-SUM(F8:H8)")) {
    errors.push("Segment Analysis!I8 should preserve a 4Q segment bridge formula based on U.S. Market annual revenue, not consolidated company revenue.");
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
