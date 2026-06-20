const path = require("node:path");
const ExcelJS = require("exceljs");
const { postWorkbook } = require("./fill-workbook-api");

const repoRoot = path.resolve(__dirname, "..");
const inputWorkbook = process.env.KO_INPUT_WORKBOOK || "/Users/maxschmieder/Downloads/KO_historicals_filled.xlsx";
const outputWorkbook = process.env.KO_OUTPUT_WORKBOOK || path.join(repoRoot, "tmp", "ko-income-statement-classification-output.xlsx");
const apiUrl = process.env.FILL_API_URL || "http://localhost:3000/api/fill-model";

const expected1Q26 = {
  Revenue: 12472,
  "Cost of Goods Sold": -4620,
  "Gross Profit": 7852,
  "Selling, General & Administration (SG&A)": -3472,
  "Other Operating Income (Expense)": -21,
  EBIT: 4359,
  "Interest Income": 222,
  "Interest (Expense)": -375,
  "Other Non-Operating Income (Expense)": 405,
  "Pre-Tax Income (Loss)": 4611,
  "Income Tax Benefit (Expense)": -645,
  "Net Income (Loss)": 3966,
  "Post-Tax Adjustments": 0,
  "Income (Loss) due to Non-Controlling Interest": -42,
  "Adj. Net Income (Loss)": 3924
};

function cellValue(cell) {
  const value = cell.value;
  if (typeof value === "number" || typeof value === "string") return value;
  if (value && typeof value === "object") return value.result ?? value.text ?? value.formula ?? null;
  return value;
}

function formula(cell) {
  const value = cell.value;
  return value && typeof value === "object" ? value.formula || value.sharedFormula || "" : "";
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
    if (label && !rows.has(normalized)) rows.set(normalized, row);
  }
  return rows;
}

function valuesMatch(actual, expected) {
  return typeof actual === "number" && Math.abs(actual - expected) <= 0.5;
}

function auditRowsFor(audit, cell, period) {
  const rows = [];
  for (let rowNumber = 2; rowNumber <= audit.rowCount; rowNumber += 1) {
    if (String(cellValue(audit.getCell(rowNumber, 2)) ?? "") !== cell) continue;
    if (String(cellValue(audit.getCell(rowNumber, 5)) ?? "") !== period) continue;
    rows.push({
      value: cellValue(audit.getCell(rowNumber, 6)),
      mappingType: String(cellValue(audit.getCell(rowNumber, 7)) ?? ""),
      concepts: String(cellValue(audit.getCell(rowNumber, 8)) ?? ""),
      labels: String(cellValue(audit.getCell(rowNumber, 9)) ?? ""),
      formulaStatus: String(cellValue(audit.getCell(rowNumber, 19)) ?? "")
    });
  }
  return rows;
}

async function main() {
  await postWorkbook({ apiUrl, ticker: "KO", inputWorkbook, outputWorkbook });

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
  if (!col) throw new Error("Could not find 1Q26 column.");
  const colLetter = model.getColumn(col).letter;
  const rows = incomeStatementRows(model);
  const row = (label) => rows.get(normalize(label));

  for (const [label, expected] of Object.entries(expected1Q26)) {
    const rowNumber = row(label);
    if (!rowNumber) {
      errors.push(`Could not find income-statement row "${label}".`);
      continue;
    }
    const actual = cellValue(model.getCell(rowNumber, col));
    if (!valuesMatch(actual, expected)) errors.push(`${period} ${label}: expected ${expected}, got ${actual ?? "[blank]"}.`);
  }

  const ebitFormula = formula(model.getCell(row("EBIT"), col));
  if (!ebitFormula || !new RegExp(`${colLetter}${row("Other Operating Income \\(Expense\\)")}`).test(ebitFormula)) {
    errors.push(`${period} EBIT formula should include the Other Operating Income (Expense) row; got ${ebitFormula || "[none]"}.`);
  }

  const pretaxFormula = formula(model.getCell(row("Pre-Tax Income (Loss)"), col));
  if (!pretaxFormula || !new RegExp(`${colLetter}${row("Other Non-Operating Income \\(Expense\\)")}`).test(pretaxFormula)) {
    errors.push(`${period} pre-tax formula should include the Other Non-Operating Income (Expense) row; got ${pretaxFormula || "[none]"}.`);
  }

  const otherOperatingRows = auditRowsFor(audit, `${colLetter}${row("Other Operating Income (Expense)")}`, period);
  if (!otherOperatingRows.some((entry) => /(OtherOperatingIncomeExpenseDerivedFromOperatingIncomeBridge|OtherSellingGeneralAndAdministrativeExpense|OtherCostAndExpenseOperating)/i.test(entry.concepts) && valuesMatch(entry.value, -21))) {
    errors.push(`Mapping Audit ${colLetter}${row("Other Operating Income (Expense)")} should support KO's other operating charges as -21 from EDGAR explicit other-operating charges or the operating-income bridge. Entries: ${JSON.stringify(otherOperatingRows)}`);
  }
  if (otherOperatingRows.some((entry) => /OtherComprehensiveIncomeLoss|AOCI/i.test(`${entry.concepts} ${entry.labels}`))) {
    errors.push(`Mapping Audit ${colLetter}${row("Other Operating Income (Expense)")} should not include OCI/AOCI reclassification rows. Entries: ${JSON.stringify(otherOperatingRows)}`);
  }

  const otherNonOperatingRows = auditRowsFor(audit, `${colLetter}${row("Other Non-Operating Income (Expense)")}`, period);
  const otherNonOperatingText = otherNonOperatingRows.map((entry) => `${entry.concepts} ${entry.labels}`).join(" | ");
  if (!/IncomeLossFromEquityMethodInvestments=384mm/.test(otherNonOperatingText)) {
    errors.push(`Mapping Audit ${colLetter}${row("Other Non-Operating Income (Expense)")} should include equity-method income in other non-operating income. Entries: ${JSON.stringify(otherNonOperatingRows)}`);
  }
  if (!/OtherNonoperatingIncomeExpense=21mm/.test(otherNonOperatingText)) {
    errors.push(`Mapping Audit ${colLetter}${row("Other Non-Operating Income (Expense)")} should include other income (loss), net. Entries: ${JSON.stringify(otherNonOperatingRows)}`);
  }

  const postTaxRows = auditRowsFor(audit, `${colLetter}${row("Post-Tax Adjustments")}`, period);
  if (postTaxRows.some((entry) => /IncomeLossFromEquityMethodInvestments/.test(entry.concepts))) {
    errors.push(`Mapping Audit ${colLetter}${row("Post-Tax Adjustments")} should not map equity-method income when it is reported above pre-tax income.`);
  }

  const netIncomeRows = auditRowsFor(audit, `${colLetter}${row("Net Income (Loss)")}`, period);
  if (netIncomeRows.some((entry) => /post-tax equity-method/i.test(entry.formulaStatus))) {
    errors.push(`${period} net-income formula should not be rewritten for a post-tax equity-method bridge.`);
  }

  if (errors.length) throw new Error(`KO income statement classification regression failed with ${errors.length} issue(s):\n${errors.join("\n")}`);
  console.log(`KO income statement classification regression passed: ${outputWorkbook}`);
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
