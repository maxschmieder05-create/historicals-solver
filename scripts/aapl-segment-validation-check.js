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
  return best.filter((info) => !info.isEstimate);
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

function findBalanceSheetRow(sheet, labels) {
  const wanted = labels.map(normalizeLabel);
  return balanceSheetRows(sheet).find((row) => wanted.includes(normalizeLabel(rowLabel(sheet, row)))) ?? null;
}

function assertTie(errors, label, actual, expected, tolerance = 0.5) {
  if (typeof actual !== "number" || typeof expected !== "number" || Math.abs(actual - expected) > tolerance) {
    errors.push(`${label}: expected ${expected ?? "[blank]"}, got ${actual ?? "[blank]"}.`);
  }
}

function assertAaplIncomeStatementConsistency(modelSheet, errors) {
  const periodColumns = new Map(historicalPeriodColumns(modelSheet).map((info) => [info.period.toUpperCase(), info.col]));
  const rowByLabel = new Map();
  const incomeStatementStart = Array.from({ length: modelSheet.rowCount }, (_, index) => index + 1)
    .filter((row) => normalizeLabel(rowLabel(modelSheet, row)) === normalizeLabel("Income Statement"))
    .find((row) => {
      for (let offset = 1; offset <= 8; offset += 1) {
        if (normalizeLabel(rowLabel(modelSheet, row + offset)) === normalizeLabel("Revenue")) return true;
      }
      return false;
    });
  if (!incomeStatementStart) {
    errors.push("Model Income Statement section was not found.");
    return;
  }

  for (let row = incomeStatementStart + 1; row <= modelSheet.rowCount; row += 1) {
    const sectionLabel = normalizeLabel(rowLabel(modelSheet, row));
    if (/incomestatementanalysis|cashflowstatement|balancesheet|workingcapital|schedule|drivers/i.test(sectionLabel)) break;
    const label = normalizeLabel(rowLabel(modelSheet, row));
    if (label) rowByLabel.set(label, row);
  }

  const row = (label) => rowByLabel.get(normalizeLabel(label));
  const cell = (label, period) => {
    const rowNumber = row(label);
    const col = periodColumns.get(period);
    return rowNumber && col ? numericCell(modelSheet.getCell(rowNumber, col)) : null;
  };

  for (const period of ["1Q25", "2Q25", "3Q25", "4Q25", "1Q26", "2Q26"]) {
    assertTie(errors, `Model Depreciation & Amortization ${period} should use income-statement-only zero policy`, cell("Depreciation & Amortization", period), 0);
  }

  const expected2Q26 = [
    ["Revenue", 111184],
    ["Cost of Goods Sold", -56403],
    ["Gross Profit", 54781],
    ["Selling, General & Administration (SG&A)", -7477],
    ["Research & Development (R&D)", -11419],
    ["Depreciation & Amortization", 0],
    ["Other Operating Income (Expense)", 0],
    ["EBIT", 35885],
    ["Interest Income", 0],
    ["Interest (Expense)", -216],
    ["Other Non-Operating Income (Expense)", 164],
    ["Pre-Tax Income (Loss)", 35833],
    ["Income Tax Benefit (Expense)", -6255],
    ["Net Income (Loss)", 29578],
    ["Pre-Tax Adjustments", 0],
    ["Post-Tax Adjustments", 0],
    ["Discontinued Operations", 0],
    ["Income (Loss) due to Non-Controlling Interest", 0],
    ["Adj. Net Income (Loss)", 29578]
  ];

  for (const [label, expected] of expected2Q26) {
    assertTie(errors, `Model ${label} 2Q26 should follow reported income-statement methodology`, cell(label, "2Q26"), expected);
  }
}

function assertBalanceSheetIntegrity(sheet, errors) {
  const periodColumns = historicalPeriodColumns(sheet);
  const periodsByName = new Map(periodColumns.map((info) => [info.period.toUpperCase(), info.col]));
  const rows = balanceSheetRows(sheet);

  for (const year of ["23", "24", "25"]) {
    const fourthQuarterCol = periodsByName.get(`4Q${year}`);
    const annualCol = periodsByName.get(`FY${year}`);
    if (!fourthQuarterCol || !annualCol) continue;
    for (const row of rows) {
      const label = rowLabel(sheet, row);
      if (!label || /balance sheet/i.test(label)) continue;
      const fourthQuarter = numericCell(sheet.getCell(row, fourthQuarterCol));
      const annual = numericCell(sheet.getCell(row, annualCol));
      if (fourthQuarter === null || annual === null) continue;
      assertTie(errors, `Model ${label} FY20${year} annual should equal 4Q`, annual, fourthQuarter);
    }
  }

  const totalAssetsRow = findBalanceSheetRow(sheet, ["Total Assets"]);
  const totalLiabilitiesAndEquityRow = findBalanceSheetRow(sheet, [
    "Total Liabilities & Shareholder's Equity",
    "Total Liabilities and Shareholder's Equity"
  ]);
  const balanceCheckRow = findBalanceSheetRow(sheet, ["Balance Sheet Check"]);
  const totalCurrentAssetsRow = findBalanceSheetRow(sheet, ["Total Current Assets"]);
  const currentAssetRows = [
    findBalanceSheetRow(sheet, ["Cash & Cash Equivalents", "Cash and Cash Equivalents"]),
    findBalanceSheetRow(sheet, ["Accounts Receivable", "Accounts Receivable, Net"]),
    findBalanceSheetRow(sheet, ["Inventory"]),
    findBalanceSheetRow(sheet, ["Prepaid & Other Current Assets", "Prepaid and Other Current Assets", "Other Current Assets"])
  ].filter(Boolean);
  const totalCurrentLiabilitiesRow = findBalanceSheetRow(sheet, [
    "Total Current Liabilities",
    "Total Current Liabilities (Excl. Debt)",
    "Total Current Liabilities Excl. Debt"
  ]);
  const currentLiabilityRows = [
    findBalanceSheetRow(sheet, ["Accounts Payable"]),
    findBalanceSheetRow(sheet, ["Accrued Liabilities", "Accrued Expenses"]),
    findBalanceSheetRow(sheet, ["Other Current Liabilities", "Other Current Liabs"])
  ].filter(Boolean);
  const totalEquityRow = findBalanceSheetRow(sheet, [
    "Total Shareholder's Equity",
    "Total Shareholders' Equity",
    "Total Stockholders' Equity"
  ]);
  const equityRows = [
    findBalanceSheetRow(sheet, ["Common Stock & APIC", "Common Stock and APIC"]),
    findBalanceSheetRow(sheet, ["Retained Earnings"]),
    findBalanceSheetRow(sheet, ["Treasury Stock"]),
    findBalanceSheetRow(sheet, ["Accumulated Other Comprehensive Income (AOCI)", "Accumulated Other Comprehensive Income", "AOCI"])
  ].filter(Boolean);

  for (const { col, period } of periodColumns) {
    if (totalAssetsRow && totalLiabilitiesAndEquityRow) {
      assertTie(
        errors,
        `Model Balance Sheet ${period} total assets should equal liabilities and equity`,
        numericCell(sheet.getCell(totalAssetsRow, col)),
        numericCell(sheet.getCell(totalLiabilitiesAndEquityRow, col))
      );
    }
    if (balanceCheckRow) {
      assertTie(errors, `Model Balance Sheet ${period} check should be zero`, numericCell(sheet.getCell(balanceCheckRow, col)), 0);
    }
    if (totalCurrentAssetsRow && currentAssetRows.length >= 2) {
      const componentSum = currentAssetRows.reduce((sum, row) => sum + (numericCell(sheet.getCell(row, col)) ?? 0), 0);
      assertTie(errors, `Model Balance Sheet ${period} current assets should sum`, numericCell(sheet.getCell(totalCurrentAssetsRow, col)), componentSum);
    }
    if (totalCurrentLiabilitiesRow && currentLiabilityRows.length >= 2) {
      const componentSum = currentLiabilityRows.reduce((sum, row) => sum + (numericCell(sheet.getCell(row, col)) ?? 0), 0);
      assertTie(errors, `Model Balance Sheet ${period} current liabilities should sum`, numericCell(sheet.getCell(totalCurrentLiabilitiesRow, col)), componentSum);
    }
    if (totalEquityRow && equityRows.length >= 2) {
      const componentSum = equityRows.reduce((sum, row) => sum + (numericCell(sheet.getCell(row, col)) ?? 0), 0);
      assertTie(errors, `Model Balance Sheet ${period} shareholders' equity should sum`, numericCell(sheet.getCell(totalEquityRow, col)), componentSum);
    }
  }
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
  assertBalanceSheetIntegrity(modelSheet, errors);
  assertAaplIncomeStatementConsistency(modelSheet, errors);

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

  const operatingIncomeBreakout = [34, 35, 36, 37, 38, 39].reduce(
    (sum, row) => sum + Math.abs(numericCell(segmentSheet.getCell(`S${row}`)) ?? 0),
    0
  );
  if (operatingIncomeBreakout <= 0.5) {
    errors.push("Segment operating income should populate when Apple discloses segment operating income detail.");
  }

  for (const col of ["I", "N", "S"]) {
    const otherRevenue = numericCell(segmentSheet.getCell(`${col}13`)) ?? 0;
    const otherLabel = String(cellValue(segmentSheet.getCell("C13")) ?? "");
    if (otherLabel.includes("Other / Reconciliation") && !valuesMatch(otherRevenue, 0)) {
      errors.push(`Segment Analysis!${col}13 should not absorb all 4Q revenue when annual product/service detail is available.`);
    }
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
