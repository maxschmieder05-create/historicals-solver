const path = require("node:path");
const ExcelJS = require("exceljs");
const { postWorkbook } = require("./fill-workbook-api");

const repoRoot = path.resolve(__dirname, "..");
const inputWorkbook = process.env.IBM_INPUT_WORKBOOK || "/Users/maxschmieder/Downloads/IBM_historicals_filled.xlsx";
const outputWorkbook = process.env.IBM_OUTPUT_WORKBOOK || path.join(repoRoot, "tmp", "ibm-income-statement-classification-output.xlsx");
const apiUrl = process.env.FILL_API_URL || "http://localhost:3000/api/fill-model";
const ticker = process.env.IBM_TICKER || "IBM";

function cellValue(cell) {
  const value = cell.value;
  if (typeof value === "number" || typeof value === "string") return value;
  if (value && typeof value === "object") return value.result ?? value.text ?? value.formula ?? null;
  return value;
}

function numericCell(cell) {
  const value = cellValue(cell);
  return typeof value === "number" ? value : null;
}

function normalize(value) {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function normalizePeriodLabel(label) {
  const compact = String(label ?? "")
    .trim()
    .replace(/\s+/g, "")
    .replace(/[’']/g, "")
    .replace(/(?:E|EST|ESTIMATE)$/i, "");
  const quarter = compact.match(/^([1-4])Q(\d{2}|\d{4})$/i);
  if (quarter) return `${quarter[1]}Q${quarter[2].slice(-2)}`.toUpperCase();
  const fiscalYear = compact.match(/^(?:FY(\d{2}|\d{4})|(\d{4}))A?$/i);
  if (fiscalYear) return `FY${(fiscalYear[1] ?? fiscalYear[2]).slice(-2)}`.toUpperCase();
  return compact.toUpperCase();
}

function rowLabel(sheet, rowNumber) {
  for (const col of [1, 2, 3, 4, 5, 6, 7, 8]) {
    const label = cellValue(sheet.getCell(rowNumber, col));
    if (label && !/^x$/i.test(String(label))) return String(label);
  }
  return "";
}

function bestPeriodHeaderRow(sheet) {
  let best = null;
  for (let row = 1; row <= Math.min(sheet.rowCount, 120); row += 1) {
    let validCount = 0;
    let quarterCount = 0;
    for (let col = 1; col <= Math.min(sheet.columnCount, 160); col += 1) {
      const period = normalizePeriodLabel(cellValue(sheet.getCell(row, col)));
      if (/^[1-4]Q\d{2}$/.test(period)) {
        validCount += 1;
        quarterCount += 1;
      } else if (/^FY\d{2}$/.test(period)) {
        validCount += 1;
      }
    }
    const score = validCount + quarterCount * 3;
    if (score && (!best || score > best.score)) best = { row, score };
  }
  return best?.row ?? null;
}

function findPeriodColumn(sheet, period) {
  const headerRow = bestPeriodHeaderRow(sheet);
  if (!headerRow) return null;
  for (let col = 1; col <= Math.min(sheet.columnCount, 160); col += 1) {
    if (normalizePeriodLabel(cellValue(sheet.getCell(headerRow, col))) === period.toUpperCase()) return col;
  }
  return null;
}

function incomeStatementRows(sheet) {
  let start = null;
  for (let row = 1; row <= sheet.rowCount; row += 1) {
    if (normalize(rowLabel(sheet, row)) !== "incomestatement") continue;
    for (let offset = 1; offset <= 8; offset += 1) {
      if (/revenue/.test(normalize(rowLabel(sheet, row + offset)))) {
        start = row;
        break;
      }
    }
    if (start) break;
  }
  if (!start) return new Map();

  const rows = new Map();
  for (let row = start + 1; row <= sheet.rowCount; row += 1) {
    const label = rowLabel(sheet, row);
    const normalized = normalize(label);
    if (/incomestatementanalysis|cashflowstatement|balancesheet|workingcapital|schedule|drivers/.test(normalized)) break;
    if (label) rows.set(normalized, row);
  }
  return rows;
}

function valuesMatch(actual, expected) {
  return typeof actual === "number" && Math.abs(actual - expected) <= 0.5;
}

function assertTie(errors, label, actual, expected) {
  if (!valuesMatch(actual, expected)) {
    errors.push(`${label}: expected ${expected}, got ${actual ?? "[blank]"}.`);
  }
}

function auditRowsFor(audit, cell, period) {
  const rows = [];
  for (let row = 2; row <= audit.rowCount; row += 1) {
    if (String(cellValue(audit.getCell(row, 2)) ?? "") !== cell) continue;
    if (String(cellValue(audit.getCell(row, 5)) ?? "") !== period) continue;
    rows.push({
      concepts: String(cellValue(audit.getCell(row, 8)) ?? ""),
      labels: String(cellValue(audit.getCell(row, 9)) ?? ""),
      value: cellValue(audit.getCell(row, 6)),
      mappingType: String(cellValue(audit.getCell(row, 7)) ?? "")
    });
  }
  return rows;
}

async function main() {
  await postWorkbook({ apiUrl, ticker, inputWorkbook, outputWorkbook });

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(outputWorkbook);
  const model = workbook.getWorksheet("Model");
  const audit = workbook.getWorksheet("Mapping Audit");
  const errors = [];

  if (!model) errors.push("Model sheet was not found.");
  if (!audit) errors.push("Mapping Audit sheet was not found.");
  if (errors.length) throw new Error(errors.join("\n"));

  const period = "1Q26";
  const col = findPeriodColumn(model, period);
  const rows = incomeStatementRows(model);
  if (!col) errors.push(`${period} column was not found.`);

  const cell = (label) => {
    const row = rows.get(normalize(label));
    return row && col ? numericCell(model.getCell(row, col)) : null;
  };

  assertTie(errors, "IBM 1Q26 SG&A should map to reported SG&A", cell("Selling, General & Administration (SG&A)"), -5089);
  assertTie(errors, "IBM 1Q26 R&D should map to reported R&D", cell("Research & Development (R&D)"), -2173);
  assertTie(errors, "IBM 1Q26 D&A should stay zero when no standalone income-statement D&A line is reported", cell("Depreciation & Amortization"), 0);
  assertTie(errors, "IBM 1Q26 other operating should map reported intellectual property/custom development income", cell("Other Operating Income (Expense)"), 172);
  assertTie(errors, "IBM 1Q26 EBIT should derive from the reported operating bridge rows", cell("EBIT"), 1860);
  assertTie(errors, "IBM 1Q26 Interest Income should stay zero without a standalone line", cell("Interest Income"), 0);
  assertTie(errors, "IBM 1Q26 Interest Expense should map direct reported interest expense", cell("Interest (Expense)"), -473);
  assertTie(errors, "IBM 1Q26 Other Non-Operating should map reported other income/expense", cell("Other Non-Operating Income (Expense)"), 1);
  assertTie(errors, "IBM 1Q26 Pre-Tax Income should remain reported pre-tax income", cell("Pre-Tax Income (Loss)"), 1387);
  assertTie(errors, "IBM 1Q26 Net Income should remain reported net income", cell("Net Income (Loss)"), 1216);

  const otherOperatingRows = auditRowsFor(audit, "U35", period);
  if (!otherOperatingRows.some((row) => /IntellectualPropertyAndCustomDevelopmentIncome/.test(row.concepts))) {
    errors.push("Mapping Audit U35 should cite IntellectualPropertyAndCustomDevelopmentIncome for IBM 1Q26 other operating income/expense.");
  }
  if (otherOperatingRows.some((row) => /RestructuringCharges/.test(row.concepts))) {
    errors.push("Mapping Audit U35 should not classify note-level restructuring charges as IBM 1Q26 other operating income/expense.");
  }

  const interestIncomeRows = auditRowsFor(audit, "U38", period);
  if (!interestIncomeRows.some((row) => row.value === 0 && /InterestIncomeNotReported/.test(row.concepts))) {
    errors.push("Mapping Audit U38 should document the no-standalone-interest-income zero policy.");
  }
  if (interestIncomeRows.some((row) => /InterestIncomeExpenseOperatingAndNonoperatingAdjustedToExcludeFinancingSegment/.test(row.concepts))) {
    errors.push("Mapping Audit U38 should not use IBM's adjusted combined interest income/expense disclosure as standalone interest income.");
  }

  const otherNonOperatingRows = auditRowsFor(audit, "U41", period);
  if (!otherNonOperatingRows.some((row) => /OtherExpenseAndIncome/.test(row.concepts))) {
    errors.push("Mapping Audit U41 should cite IBM's reported OtherExpenseAndIncome line.");
  }
  if (otherNonOperatingRows.some((row) => /OtherNonOperatingIncomeExpenseFromReportedLine|InterestIncomeExpenseOperatingAndNonoperatingAdjustedToExcludeFinancingSegment/.test(row.concepts))) {
    errors.push("Mapping Audit U41 should not use a derived residual split or adjusted interest disclosure for IBM 1Q26 other non-operating income/expense.");
  }

  if (errors.length) {
    console.error(errors.join("\n"));
    throw new Error(`IBM income statement classification regression failed with ${errors.length} issue(s).`);
  }

  console.log(`IBM income statement classification regression passed: ${outputWorkbook}`);
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
