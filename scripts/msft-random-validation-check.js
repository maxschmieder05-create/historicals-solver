const path = require("node:path");
const ExcelJS = require("exceljs");
const { postWorkbook } = require("./fill-workbook-api");

const repoRoot = path.resolve(__dirname, "..");
const inputWorkbook =
  process.env.MSFT_INPUT_WORKBOOK || "/Users/maxschmieder/Downloads/MSFT_historicals_filled (1).xlsx";
const outputWorkbook = process.env.MSFT_OUTPUT_WORKBOOK || path.join(repoRoot, "tmp", "random-msft-validation-output.xlsx");
const apiUrl = process.env.FILL_API_URL || "http://localhost:3000/api/fill-model";
const ticker = process.env.MSFT_TICKER || "MSFT";

const expectedCells = [
  ["Model", "I158", 0, "4Q23 balance sheet check"],
  ["Model", "J158", 0, "FY23 balance sheet check"],
  ["Segment Analysis", "I7", 56189, "4Q23 total company segment revenue"]
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

function normalizeLabel(value) {
  return String(value ?? "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "");
}

function findRow(sheet, label) {
  const wanted = normalizeLabel(label);
  for (let row = 1; row <= sheet.rowCount; row += 1) {
    const actual = cellValue(sheet.getCell(row, 3)) ?? cellValue(sheet.getCell(row, 2)) ?? cellValue(sheet.getCell(row, 1));
    if (normalizeLabel(actual) === wanted) return row;
  }
  return null;
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

function historicalPeriodColumns(sheet) {
  const columns = [];
  for (let col = 1; col <= sheet.columnCount; col += 1) {
    const label = String(cellValue(sheet.getCell(5, col)) ?? "").trim();
    if (!label || /e$/i.test(label)) continue;
    if (/^(?:[1-4]Q\d{2}|20\d{2})$/i.test(label)) columns.push(col);
  }
  return columns;
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

  const segmentSheet = workbook.getWorksheet("Segment Analysis");
  const modelSheet = workbook.getWorksheet("Model");
  if (!segmentSheet || !modelSheet) {
    errors.push("Workbook is missing Model or Segment Analysis sheet.");
  } else {
    const expectedLabels = [
      ["C8", "Productivity and Business Processes Revenue"],
      ["C9", "Intelligent Cloud Revenue"],
      ["C10", "More Personal Computing Revenue"]
    ];
    for (const [address, expected] of expectedLabels) {
      const actual = String(cellValue(segmentSheet.getCell(address)) ?? "");
      if (actual !== expected) errors.push(`Segment Analysis!${address}: expected "${expected}", got "${actual}".`);
    }

    const revenueMixRow = findRow(segmentSheet, "Revenue Mix");
    const totalRevenueRow = findRow(segmentSheet, "Total Company Revenue");
    const segmentRows = [8, 9, 10];
    const detailRows = [];
    if (totalRevenueRow && revenueMixRow) {
      for (let row = totalRevenueRow + 1; row < revenueMixRow; row += 1) detailRows.push(row);
    }

    for (const row of detailRows) {
      const label = String(cellValue(segmentSheet.getCell(row, 3)) ?? "");
      if (/\$|(?:^|\s)\d[\d,]*(?:\.\d+)?(?:\s|$)/.test(label)) {
        errors.push(`Segment Analysis!C${row}: invalid numeric/currency-derived segment label "${label}".`);
      }
    }

    for (const col of historicalPeriodColumns(segmentSheet)) {
      const period = cellValue(segmentSheet.getCell(5, col));
      const letter = columnLetter(col);
      const disclosedRevenue = segmentRows.reduce((sum, row) => sum + (cellValue(segmentSheet.getCell(row, col)) ?? 0), 0);
      const detailRevenue = detailRows.reduce((sum, row) => sum + (cellValue(segmentSheet.getCell(row, col)) ?? 0), 0);
      const segmentTotal = cellValue(segmentSheet.getCell(7, col));
      const modelRevenue = cellValue(modelSheet.getCell(28, col));
      for (const row of segmentRows) {
        const value = cellValue(segmentSheet.getCell(row, col));
        if (typeof value !== "number" || Math.abs(value) <= 0.0001) {
          errors.push(`Segment Analysis!${letter}${row} ${period}: expected disclosed MSFT segment revenue, got ${value ?? "[blank]"}.`);
        }
      }
      if (!valuesMatch(disclosedRevenue, modelRevenue)) {
        errors.push(`${period} disclosed segment revenue rows sum to ${disclosedRevenue}, but Model!${letter}28 revenue is ${modelRevenue ?? "[blank]"}.`);
      }
      if (!valuesMatch(detailRevenue, modelRevenue)) {
        errors.push(`${period} segment detail rows sum to ${detailRevenue}, but Model!${letter}28 revenue is ${modelRevenue ?? "[blank]"}.`);
      }
      if (!valuesMatch(segmentTotal, modelRevenue)) {
        errors.push(`${period} Segment Analysis!${letter}7 total is ${segmentTotal ?? "[blank]"}, but Model!${letter}28 revenue is ${modelRevenue ?? "[blank]"}.`);
      }
    }

    const fourthQuarterSegmentRevenue = segmentRows.reduce((sum, row) => sum + (cellValue(segmentSheet.getCell(`I${row}`)) ?? 0), 0);
    if (valuesMatch(fourthQuarterSegmentRevenue, 211915)) {
      errors.push("4Q23 segment revenue rows still contain the FY23 segment revenue total instead of standalone 4Q revenue.");
    }
  }

  if (errors.length) {
    console.error(errors.join("\n"));
    throw new Error(`MSFT random-company validation failed with ${errors.length} issue(s).`);
  }

  console.log(`MSFT random-company validation passed: ${outputWorkbook}`);
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
