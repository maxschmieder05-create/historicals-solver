const path = require("node:path");
const ExcelJS = require("exceljs");
const { postWorkbook } = require("./fill-workbook-api");

const repoRoot = path.resolve(__dirname, "..");
const inputWorkbook = process.env.RSG_INPUT_WORKBOOK || "/Users/maxschmieder/Downloads/Blank RSG.xlsx";
const outputWorkbook = process.env.RSG_SEGMENT_OUTPUT_WORKBOOK || path.join(repoRoot, "tmp", "rsg-segment-operating-income-output.xlsx");
const apiUrl = process.env.FILL_API_URL || "http://localhost:3000/api/fill-model";

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

function normalizeLabel(value) {
  return String(value ?? "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "");
}

function rowLabel(sheet, rowNumber) {
  for (const col of [1, 2, 3, 4, 5]) {
    const label = cellValue(sheet.getCell(rowNumber, col));
    if (label && !/^x$/i.test(String(label))) return String(label);
  }
  return "";
}

function findRow(sheet, label) {
  for (let row = 1; row <= sheet.rowCount; row += 1) {
    if (normalizeLabel(rowLabel(sheet, row)) === normalizeLabel(label)) return row;
  }
  return null;
}

function valuesMatch(actual, expected) {
  return typeof actual === "number" && typeof expected === "number" && Math.abs(actual - expected) <= 0.5;
}

function columnLetter(col) {
  let value = col;
  let letters = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    letters = String.fromCharCode(65 + remainder) + letters;
    value = Math.floor((value - 1) / 26);
  }
  return letters;
}

async function main() {
  await postWorkbook({ apiUrl, ticker: "RSG", inputWorkbook, outputWorkbook });

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(outputWorkbook);
  const segmentSheet = workbook.getWorksheet("Segment Analysis");
  const modelSheet = workbook.getWorksheet("Model");
  const errors = [];

  if (!segmentSheet || !modelSheet) {
    throw new Error("Workbook is missing Model or Segment Analysis sheet.");
  }

  const totalRow = findRow(segmentSheet, "Total Company Operating Income");
  const checkRow = findRow(segmentSheet, "Operating Income Check");
  const modelEbitRow = findRow(modelSheet, "EBIT") ?? findRow(modelSheet, "Operating Income");
  const modelDaRow = findRow(modelSheet, "Depreciation & Amortization");
  const otherOperatingRow = findRow(modelSheet, "Other Operating Income (Expense)");

  if (!totalRow || !checkRow || !modelEbitRow || !modelDaRow || !otherOperatingRow) {
    throw new Error("Workbook is missing Segment Analysis operating income rows or Model EBIT/D&A/other operating rows.");
  }

  const historicalColumns = Array.from({ length: 15 }, (_, index) => index + 6);
  const nonzeroDetailLabels = new Set();
  const expectedDaByPeriod = new Map([
    ["1Q23", -358.7],
    ["2023", -1501.3]
  ]);
  const expectedOtherOperatingByPeriod = new Map([
    ["1Q23", -29.6],
    ["2023", -132.1]
  ]);

  for (const col of historicalColumns) {
    const period = cellValue(segmentSheet.getCell(5, col));
    const segmentTotal = numericCell(segmentSheet.getCell(totalRow, col));
    const modelEbit = numericCell(modelSheet.getCell(modelEbitRow, col));
    const check = numericCell(segmentSheet.getCell(checkRow, col)) ?? 0;

    if (!valuesMatch(segmentTotal, modelEbit)) {
      errors.push(`Segment Analysis!${columnLetter(col)}${totalRow} ${period}: expected Model EBIT ${modelEbit ?? "[blank]"}, got ${segmentTotal ?? "[blank]"}.`);
    }
    if (Math.abs(check) > 0.05) {
      errors.push(`Segment Analysis!${columnLetter(col)}${checkRow} ${period}: expected operating income check near zero, got ${check}.`);
    }

    const expectedDa = expectedDaByPeriod.get(String(period));
    if (expectedDa !== undefined && !valuesMatch(numericCell(modelSheet.getCell(modelDaRow, col)), expectedDa)) {
      errors.push(`Model!${columnLetter(col)}${modelDaRow} ${period}: expected primary-statement D&A ${expectedDa}, got ${numericCell(modelSheet.getCell(modelDaRow, col)) ?? "[blank]"}.`);
    }
    const expectedOtherOperating = expectedOtherOperatingByPeriod.get(String(period));
    if (expectedOtherOperating !== undefined && !valuesMatch(numericCell(modelSheet.getCell(otherOperatingRow, col)), expectedOtherOperating)) {
      errors.push(
        `Model!${columnLetter(col)}${otherOperatingRow} ${period}: expected explicit accretion/restructuring operating items ${expectedOtherOperating}, got ${numericCell(modelSheet.getCell(otherOperatingRow, col)) ?? "[blank]"}.`
      );
    }

    for (let row = totalRow + 1; row < checkRow; row += 1) {
      const value = numericCell(segmentSheet.getCell(row, col)) ?? 0;
      if (Math.abs(value) > 0.05) nonzeroDetailLabels.add(rowLabel(segmentSheet, row));
    }
  }

  if (![...nonzeroDetailLabels].some((label) => /other|reconciliation/i.test(label))) {
    errors.push("Segment Analysis operating income should use an Other/Reconciliation residual when EDGAR does not provide valid segment operating income.");
  }

  for (const label of nonzeroDetailLabels) {
    if (/reclassification|accumulated other comprehensive income|cash flow hedge|aoci/i.test(label)) {
      errors.push(`Segment Analysis should not promote OCI/reclassification member "${label}" to an operating income segment.`);
    }
  }

  if (errors.length) {
    console.error(errors.join("\n"));
    throw new Error(`RSG segment operating income regression failed with ${errors.length} issue(s).`);
  }

  console.log(`RSG segment operating income regression passed: ${outputWorkbook}`);
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
