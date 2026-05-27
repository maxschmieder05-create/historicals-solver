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

const expectedRevenueCells = [
  ["F7", 117154, "1Q23 total revenue"],
  ["I7", 89498, "4Q23 total revenue"],
  ["K7", 119575, "1Q24 total revenue"],
  ["I8", 22314, "4Q23 Services revenue"],
  ["I9", 43805, "4Q23 iPhone revenue"],
  ["I10", 7614, "4Q23 Mac revenue"],
  ["I11", 6443, "4Q23 iPad revenue"],
  ["I12", 9322, "4Q23 Wearables revenue"],
  ["N8", 24972, "4Q24 Services revenue"],
  ["N9", 46222, "4Q24 iPhone revenue"],
  ["N10", 7744, "4Q24 Mac revenue"],
  ["N11", 6950, "4Q24 iPad revenue"],
  ["N12", 9042, "4Q24 Wearables revenue"],
  ["S8", 28750, "4Q25 Services revenue"],
  ["S9", 49025, "4Q25 iPhone revenue"],
  ["S10", 8726, "4Q25 Mac revenue"],
  ["S11", 6952, "4Q25 iPad revenue"],
  ["S12", 9013, "4Q25 Wearables revenue"]
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

function normalizePeriodLabel(label) {
  const compact = String(label ?? "")
    .trim()
    .replace(/\s+/g, "")
    .replace(/[’']/g, "")
    .replace(/(?:E|EST|ESTIMATE)$/i, "");
  if (/^[1-4]Q(\d{2}|\d{4})$/i.test(compact)) return compact.toUpperCase();
  if (/^(\d{4}|\d{2})A?$/i.test(compact)) return `FY${compact.slice(-2)}`;
  return compact;
}

function isEstimatePeriodLabel(label) {
  const compact = String(label ?? "")
    .trim()
    .replace(/\s+/g, "")
    .replace(/[’']/g, "");
  return /(?:^|[^a-z])(?:E|EST|ESTIMATE)$/i.test(compact) || /(?:\d{2}|\d{4})E$/i.test(compact);
}

function projectedColumns(sheet) {
  let best = [];
  for (let row = 1; row <= Math.min(sheet.rowCount, 120); row += 1) {
    const infos = [];
    for (let col = 4; col <= Math.min(sheet.columnCount, 160); col += 1) {
      const label = cellValue(sheet.getCell(row, col));
      const period = normalizePeriodLabel(label);
      if (!/^(?:[1-4]Q\d{2}|FY\d{2})$/i.test(period)) continue;
      infos.push({ col, isEstimate: isEstimatePeriodLabel(label) });
    }
    if (infos.length > best.length) best = infos;
  }
  return best.filter((info) => info.isEstimate).map((info) => info.col);
}

function balanceSheetRows(sheet) {
  const rows = [];
  let inBalanceSheet = false;
  for (let row = 1; row <= sheet.rowCount; row += 1) {
    const label = rowLabel(sheet, row);
    const normalized = normalizeLabel(label);
    if (!inBalanceSheet) {
      if (normalized === normalizeLabel("Balance Sheet")) inBalanceSheet = true;
      continue;
    }
    if (/workingcapital|cashflowstatement|incomestatement|schedule|analysis|drivers/i.test(normalized)) break;
    if (label) rows.push(row);
  }
  return rows;
}

function cellFingerprint(cell) {
  return JSON.stringify({ value: cell.value ?? null, note: cell.note ?? null });
}

function assertProjectedBalanceSheetPreserved(sourceSheet, outputSheet, errors) {
  const cols = projectedColumns(sourceSheet);
  const rows = balanceSheetRows(sourceSheet);
  for (const row of rows) {
    for (const col of cols) {
      const source = sourceSheet.getCell(row, col);
      const output = outputSheet.getCell(row, col);
      if (cellFingerprint(source) !== cellFingerprint(output)) {
        errors.push(`Model!${output.address}: projected Balance Sheet cell changed but should be preserved exactly.`);
      }
    }
  }
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

  const input = new ExcelJS.Workbook();
  await input.xlsx.readFile(inputWorkbook);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(outputWorkbook);
  const inputModelSheet = input.getWorksheet("Model");
  const segmentSheet = workbook.getWorksheet("Segment Analysis");
  const modelSheet = workbook.getWorksheet("Model");
  const errors = [];

  if (!segmentSheet || !modelSheet) throw new Error("Workbook is missing Model or Segment Analysis sheet.");
  if (!inputModelSheet) throw new Error("Input workbook is missing Model sheet.");
  assertProjectedBalanceSheetPreserved(inputModelSheet, modelSheet, errors);

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

  for (const [address, expected, label] of expectedRevenueCells) {
    const actual = numericCell(segmentSheet.getCell(address));
    if (!valuesMatch(actual, expected)) errors.push(`${label} Segment Analysis!${address}: expected ${expected}, got ${actual ?? "[blank]"}.`);
  }

  const operatingIncomeReconciliation = numericCell(segmentSheet.getCell("F35")) ?? 0;
  const operatingIncomeLabel = String(cellValue(segmentSheet.getCell("C35")) ?? "");
  if (operatingIncomeLabel.includes("Other / Reconciliation") || !valuesMatch(operatingIncomeReconciliation, 0)) {
    errors.push("Segment operating income should not invent an Other / Reconciliation row when Apple does not report product-level operating income.");
  }

  for (const col of ["I", "N", "S"]) {
    const otherRevenue = numericCell(segmentSheet.getCell(`${col}13`)) ?? 0;
    const otherLabel = String(cellValue(segmentSheet.getCell("C13")) ?? "");
    if (otherLabel.includes("Other / Reconciliation") && !valuesMatch(otherRevenue, 0)) {
      errors.push(`Segment Analysis!${col}13 should not absorb all 4Q revenue when annual product/service detail is available.`);
    }
  }

  for (const col of ["F", "G", "H", "I", "K", "L", "M"]) {
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
