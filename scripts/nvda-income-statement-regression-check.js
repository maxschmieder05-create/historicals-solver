const path = require("node:path");
const ExcelJS = require("exceljs");
const { postWorkbook } = require("./fill-workbook-api");

const repoRoot = path.resolve(__dirname, "..");
const inputWorkbook = process.env.NVDA_INPUT_WORKBOOK || "/Users/maxschmieder/Downloads/NVDA_historicals_filled.xlsx";
const outputWorkbook = process.env.NVDA_OUTPUT_WORKBOOK || path.join(repoRoot, "tmp", "nvda-income-statement-regression-output.xlsx");
const apiUrl = process.env.FILL_API_URL || "http://localhost:3000/api/fill-model";
const ticker = process.env.NVDA_TICKER || "NVDA";

const expectedAnchors = {
  "4Q25": {
    Revenue: 39331,
    "Gross Profit": 28723,
    EBIT: 24034,
    "Pre-Tax Income (Loss)": 25217,
    "Income Tax Benefit (Expense)": -3126,
    "Net Income (Loss)": 22091
  },
  FY25: {
    Revenue: 130497,
    "Gross Profit": 97858,
    EBIT: 81453,
    "Pre-Tax Income (Loss)": 84026,
    "Income Tax Benefit (Expense)": -11146,
    "Net Income (Loss)": 72880
  }
};

function cellValue(cell) {
  const value = cell.value;
  if (typeof value === "number" || typeof value === "string") return value;
  if (value && typeof value === "object") return value.result ?? value.text ?? value.formula ?? null;
  return value;
}

function normalize(value) {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function valuesMatch(actual, expected) {
  return typeof actual === "number" && Math.abs(actual - expected) <= 0.5;
}

function findPeriodColumn(sheet, period) {
  const headerRow = bestPeriodHeaderRow(sheet);
  if (!headerRow) return null;
  for (let col = 1; col <= Math.min(sheet.columnCount, 160); col += 1) {
    const label = String(cellValue(sheet.getCell(headerRow, col)) ?? "").trim();
    if (period === "FY25" && label === "2025") return col;
    if (label.toUpperCase() === period) return col;
  }
  return null;
}

function bestPeriodHeaderRow(sheet) {
  let best = null;
  for (let row = 1; row <= Math.min(sheet.rowCount, 120); row += 1) {
    let validCount = 0;
    let quarterCount = 0;
    for (let col = 1; col <= Math.min(sheet.columnCount, 160); col += 1) {
      const label = String(cellValue(sheet.getCell(row, col)) ?? "").trim();
      if (/^[1-4]Q\d{2}$/i.test(label)) {
        validCount += 1;
        quarterCount += 1;
      } else if (/^20\d{2}$/.test(label)) {
        validCount += 1;
      }
    }
    if (!validCount) continue;
    const score = validCount + quarterCount * 3;
    if (!best || score > best.score) best = { row, score };
  }
  return best?.row ?? null;
}

function findRow(sheet, label) {
  const wanted = normalize(label);
  for (let row = 1; row <= sheet.rowCount; row += 1) {
    for (const col of [1, 2, 3, 4, 5]) {
      if (normalize(cellValue(sheet.getCell(row, col))) === wanted) return row;
    }
  }
  return null;
}

async function main() {
  await postWorkbook({ apiUrl, ticker, inputWorkbook, outputWorkbook });

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(outputWorkbook);
  const model = workbook.getWorksheet("Model");
  if (!model) throw new Error("Output workbook does not contain a Model sheet.");

  const errors = [];
  for (const [period, anchors] of Object.entries(expectedAnchors)) {
    const col = findPeriodColumn(model, period);
    if (!col) {
      errors.push(`Could not find ${period} column in Model sheet.`);
      continue;
    }

    for (const [label, expected] of Object.entries(anchors)) {
      const row = findRow(model, label);
      if (!row) {
        errors.push(`Could not find row "${label}" in Model sheet.`);
        continue;
      }
      const actual = cellValue(model.getCell(row, col));
      if (!valuesMatch(actual, expected)) {
        errors.push(`${period} ${label}: expected ${expected}, got ${actual ?? "[blank]"}.`);
      }
    }
  }

  if (errors.length) {
    console.error(errors.join("\n"));
    throw new Error(`NVDA income statement regression failed with ${errors.length} issue(s).`);
  }

  console.log(`NVDA income statement regression passed: ${outputWorkbook}`);
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
