const path = require("node:path");
const ExcelJS = require("exceljs");
const { postWorkbook } = require("./fill-workbook-api");

const repoRoot = path.resolve(__dirname, "..");
const inputWorkbook =
  process.env.SEGMENT_RELIABILITY_INPUT_WORKBOOK || "/Users/maxschmieder/Downloads/Owl Fund Integrated Model Template (03-Sep-2025)_v25 (3).xlsx";
const outputDir = process.env.SEGMENT_RELIABILITY_OUTPUT_DIR || path.join(repoRoot, "tmp");
const apiUrl = process.env.FILL_API_URL || "http://localhost:3000/api/fill-model";
const tickers = (process.env.SEGMENT_RELIABILITY_TICKERS || "GOOG,NVDA")
  .split(",")
  .map((ticker) => ticker.trim().toUpperCase())
  .filter(Boolean);

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
  return typeof actual === "number" && typeof expected === "number" && Math.abs(actual - expected) <= 0.5;
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
  const wanted = normalizeLabel(label);
  for (let row = 1; row <= sheet.rowCount; row += 1) {
    if (normalizeLabel(rowLabel(sheet, row)) === wanted) return row;
  }
  return null;
}

function normalizePeriodLabel(label) {
  const compact = String(label ?? "")
    .trim()
    .replace(/\s+/g, "")
    .replace(/[’']/g, "")
    .replace(/(?:E|EST|ESTIMATE)$/i, "");
  if (/^[1-4]Q(\d{2}|\d{4})$/i.test(compact)) return compact.toUpperCase().replace(/(\d{4})$/, (year) => year.slice(-2));
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

function historicalPeriodColumns(sheet) {
  let best = [];
  for (let row = 1; row <= Math.min(sheet.rowCount, 120); row += 1) {
    const infos = [];
    for (let col = 4; col <= Math.min(sheet.columnCount, 160); col += 1) {
      const label = cellValue(sheet.getCell(row, col));
      const period = normalizePeriodLabel(label);
      if (!/^(?:[1-4]Q\d{2}|FY\d{2})$/i.test(period)) continue;
      infos.push({ col, period, isEstimate: isEstimatePeriodLabel(label) });
    }
    const score = infos.length + infos.filter((info) => /^[1-4]Q/i.test(info.period)).length * 3;
    const bestScore = best.length + best.filter((info) => /^[1-4]Q/i.test(info.period)).length * 3;
    if (score > bestScore) best = infos;
  }
  return best.filter((info) => !info.isEstimate && /^[1-4]Q/i.test(info.period));
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

function columnIndex(letters) {
  return String(letters)
    .toUpperCase()
    .split("")
    .reduce((total, char) => total * 26 + char.charCodeAt(0) - 64, 0);
}

function segmentBaseLabel(label, suffix) {
  return String(label ?? "")
    .replace(/^"+|"+$/g, "")
    .replace(/^=/, "")
    .replace(new RegExp(`\\s*${suffix.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}\\s*$`, "i"), "")
    .replace(/\s*Revenue\s*$/i, "")
    .replace(/\s*Operating Income\s*$/i, "")
    .replace(/\s*D&A\s*$/i, "")
    .trim();
}

function isInvalidSegmentLabel(label) {
  if (!label) return false;
  const letters = label.match(/[A-Za-z]/g)?.length ?? 0;
  const numericCurrency = label.match(/[\d$,.()\-]/g)?.length ?? 0;
  return /^segment\s*\d+$/i.test(label) || /^\(?\$?\d[\d,]*(?:\.\d+)?\)?$/.test(label) || (/\d/.test(label) && numericCurrency > letters);
}

function detailRowsBetween(sheet, startLabel, endLabel) {
  const startRow = findRow(sheet, startLabel);
  const endRow = findRow(sheet, endLabel);
  if (!startRow || !endRow || endRow <= startRow) return [];
  const rows = [];
  for (let row = startRow + 1; row < endRow; row += 1) rows.push(row);
  return rows;
}

function auditedHistoricalColumns(workbook, modelRevenueRow, segmentRevenueTotalRow) {
  const audit = workbook.getWorksheet("Mapping Audit");
  if (!audit) return null;
  const columns = new Set();
  for (let row = 2; row <= audit.rowCount; row += 1) {
    const sheetName = String(cellValue(audit.getCell(row, 1)) ?? "");
    const cellAddress = String(cellValue(audit.getCell(row, 2)) ?? "");
    const modelLabel = String(cellValue(audit.getCell(row, 3)) ?? "");
    const match = cellAddress.match(/^([A-Z]+)(\d+)$/i);
    if (!match) continue;
    const rowNumber = Number(match[2]);
    if (sheetName === "Model" && rowNumber === modelRevenueRow && normalizeLabel(modelLabel) === normalizeLabel("Revenue")) {
      columns.add(columnIndex(match[1]));
    }
    if (sheetName === "Segment Analysis" && rowNumber === segmentRevenueTotalRow && normalizeLabel(modelLabel) === normalizeLabel("Total Company Revenue")) {
      columns.add(columnIndex(match[1]));
    }
  }
  return columns.size ? columns : null;
}

async function assertWorkbook(ticker, outputWorkbook) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(outputWorkbook);
  const segmentSheet = workbook.getWorksheet("Segment Analysis");
  const modelSheet = workbook.getWorksheet("Model");
  const errors = [];
  if (!segmentSheet || !modelSheet) throw new Error(`${ticker}: workbook is missing Model or Segment Analysis sheet.`);

  let revenueColumns = historicalPeriodColumns(segmentSheet);
  const modelRevenueRow = findRow(modelSheet, "Revenue");
  const modelEbitRow = findRow(modelSheet, "EBIT");
  const segmentRevenueTotalRow = findRow(segmentSheet, "Total Company Revenue");
  const segmentOperatingTotalRow = findRow(segmentSheet, "Total Company Operating Income");
  if (!modelRevenueRow || !modelEbitRow || !segmentRevenueTotalRow || !segmentOperatingTotalRow) {
    throw new Error(`${ticker}: missing required revenue or operating income rows.`);
  }
  const auditedColumns = auditedHistoricalColumns(workbook, modelRevenueRow, segmentRevenueTotalRow);
  if (auditedColumns) revenueColumns = revenueColumns.filter(({ col }) => auditedColumns.has(col));

  const revenueRows = detailRowsBetween(segmentSheet, "Total Company Revenue", "Revenue Mix");
  const activeRevenueRows = revenueRows.filter((row) => rowLabel(segmentSheet, row).trim());
  const fallbackRow = activeRevenueRows.find((row) => /^Reported Revenue$/i.test(rowLabel(segmentSheet, row).trim()));

  for (const row of revenueRows) {
    const label = rowLabel(segmentSheet, row).trim();
    const baseLabel = segmentBaseLabel(label, "Revenue");
    const absoluteRevenue = revenueColumns.reduce((sum, { col }) => sum + Math.abs(numericCell(segmentSheet.getCell(row, col)) ?? 0), 0);
    if (label && absoluteRevenue <= 0.0001) errors.push(`${ticker}: Segment Analysis row ${row} has stale zero-value label "${label}".`);
    if (isInvalidSegmentLabel(baseLabel)) errors.push(`${ticker}: Segment Analysis row ${row} has invalid segment label "${label}".`);
  }

  for (const { col, period } of revenueColumns) {
    const letter = columnLetter(col);
    const modelRevenue = numericCell(modelSheet.getCell(modelRevenueRow, col));
    const segmentTotal = numericCell(segmentSheet.getCell(segmentRevenueTotalRow, col));
    const detailTotal = activeRevenueRows.reduce((sum, row) => sum + (numericCell(segmentSheet.getCell(row, col)) ?? 0), 0);
    const modelEbit = numericCell(modelSheet.getCell(modelEbitRow, col));
    const segmentOperatingTotal = numericCell(segmentSheet.getCell(segmentOperatingTotalRow, col));

    if (!valuesMatch(segmentTotal, modelRevenue)) {
      errors.push(`${ticker} ${period}: Segment Analysis!${letter}${segmentRevenueTotalRow} revenue total ${segmentTotal ?? "[blank]"} does not tie to Model revenue ${modelRevenue ?? "[blank]"}.`);
    }
    if (!valuesMatch(detailTotal, modelRevenue)) {
      errors.push(`${ticker} ${period}: segment revenue rows sum to ${detailTotal}, but Model revenue is ${modelRevenue ?? "[blank]"}.`);
    }
    if (!valuesMatch(segmentOperatingTotal, modelEbit)) {
      errors.push(`${ticker} ${period}: Segment Analysis operating income total ${segmentOperatingTotal ?? "[blank]"} does not tie to Model EBIT ${modelEbit ?? "[blank]"}.`);
    }
    if (fallbackRow) {
      const fallbackRevenue = numericCell(segmentSheet.getCell(fallbackRow, col));
      if (!valuesMatch(fallbackRevenue, modelRevenue)) {
        errors.push(`${ticker} ${period}: Reported Revenue fallback row ${fallbackRevenue ?? "[blank]"} does not tie to Model revenue ${modelRevenue ?? "[blank]"}.`);
      }
    }
  }

  if (errors.length) {
    console.error(errors.join("\n"));
    throw new Error(`${ticker} segment first-run reliability failed with ${errors.length} issue(s).`);
  }
}

async function main() {
  for (const ticker of tickers) {
    const outputWorkbook = path.join(outputDir, `${ticker.toLowerCase()}-segment-first-run-output.xlsx`);
    await postWorkbook({ apiUrl, ticker, inputWorkbook, outputWorkbook });
    await assertWorkbook(ticker, outputWorkbook);
    console.log(`${ticker} segment first-run reliability passed: ${outputWorkbook}`);
  }
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
