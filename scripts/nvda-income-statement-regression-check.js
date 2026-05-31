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

const expectedFiscalPeriodAnchors = {
  "1Q26": {
    Revenue: 44062,
    "Total Assets": 125254,
    "Total Liabilities & Shareholder's Equity": 125254
  },
  "3Q26": {
    Revenue: 57006,
    "Total Assets": 161148,
    "Total Liabilities & Shareholder's Equity": 161148
  },
  "4Q26": {
    Revenue: 68127,
    "Total Assets": 206803,
    "Total Liabilities & Shareholder's Equity": 206803
  },
  FY26: {
    Revenue: 215938,
    "Total Assets": 206803,
    "Total Liabilities & Shareholder's Equity": 206803
  },
  "1Q27": {
    Revenue: 81615,
    "Total Assets": 259474,
    "Total Liabilities & Shareholder's Equity": 259474
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
    if (normalizePeriodLabel(label) === period.toUpperCase()) return col;
  }
  return null;
}

function normalizePeriodLabel(label) {
  const compact = String(label ?? "")
    .trim()
    .replace(/\s+/g, "")
    .replace(/[’']/g, "")
    .replace(/(?:E|EST|ESTIMATE)$/i, "");
  const direct = compact.match(/^([1-4])Q(\d{2}|\d{4})$/i);
  if (direct) return `${direct[1]}Q${direct[2].slice(-2)}`.toUpperCase();
  const fiscalYear = compact.match(/^(?:FY(\d{2}|\d{4})|(\d{4}))A?$/i);
  if (fiscalYear) return `FY${(fiscalYear[1] ?? fiscalYear[2]).slice(-2)}`.toUpperCase();
  return compact.toUpperCase();
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
  for (const [period, anchors] of Object.entries({ ...expectedAnchors, ...expectedFiscalPeriodAnchors })) {
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

  const segment = workbook.getWorksheet("Segment Analysis");
  if (!segment) {
    errors.push("Output workbook does not contain a Segment Analysis sheet.");
  } else {
    const segmentRevenueRow = findRow(segment, "Total Company Revenue");
    const modelRevenueRow = findRow(model, "Revenue");
    if (!segmentRevenueRow || !modelRevenueRow) {
      errors.push("Could not find Segment Analysis Total Company Revenue or Model Revenue row.");
    } else {
      for (const period of Object.keys(expectedFiscalPeriodAnchors)) {
        const col = findPeriodColumn(model, period);
        if (!col) continue;
        const modelRevenue = cellValue(model.getCell(modelRevenueRow, col));
        const segmentRevenue = cellValue(segment.getCell(segmentRevenueRow, col));
        if (!valuesMatch(segmentRevenue, modelRevenue)) {
          errors.push(`${period} Segment Analysis Total Company Revenue should tie to Model Revenue: expected ${modelRevenue}, got ${segmentRevenue ?? "[blank]"}.`);
        }
      }
    }
  }

  const fy26AssetsCol = findPeriodColumn(model, "FY26");
  const q426AssetsCol = findPeriodColumn(model, "4Q26");
  const assetsRow = findRow(model, "Total Assets");
  if (fy26AssetsCol && q426AssetsCol && assetsRow) {
    const fy26Assets = cellValue(model.getCell(assetsRow, fy26AssetsCol));
    const q426Assets = cellValue(model.getCell(assetsRow, q426AssetsCol));
    if (!valuesMatch(fy26Assets, q426Assets)) errors.push(`FY26 balance sheet should equal 4Q26 year-end balance sheet: expected ${q426Assets}, got ${fy26Assets}.`);
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
