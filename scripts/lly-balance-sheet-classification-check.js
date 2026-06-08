const path = require("node:path");
const ExcelJS = require("exceljs");
const { postWorkbook } = require("./fill-workbook-api");

const repoRoot = path.resolve(__dirname, "..");
const inputWorkbook = process.env.LLY_INPUT_WORKBOOK || "/Users/maxschmieder/Downloads/LLY_historicals_filled.xlsx";
const outputWorkbook = process.env.LLY_BALANCE_SHEET_OUTPUT_WORKBOOK || path.join(repoRoot, "tmp", "lly-balance-sheet-classification-output.xlsx");
const apiUrl = process.env.FILL_API_URL || "http://localhost:3000/api/fill-model";
const ticker = process.env.LLY_TICKER || "LLY";

const expected1Q26 = {
  "Total Current Liabilities (Excl. Debt)": 32634,
  Revolver: 4000,
  "LT Debt (Incl. Current Portion)": 39370,
  "Other Non-Current Liabilities": 9374,
  "Total Non-Current Liabilities": 48744,
  "Total Liabilities": 85378,
  "Common Stock & APIC": 7530,
  "Retained Earnings": 29514,
  "Treasury Stock": -3013,
  "Accumulated Other Comprehensive Income (AOCI)": -2833,
  "Total Shareholder's Equity": 31198,
  "Total Liabilities & Shareholder's Equity": 116576,
  "Balance Sheet Check": 0
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
      } else if (/^FY\d{2}$/i.test(label) || /^20\d{2}$/.test(label)) {
        validCount += 1;
      }
    }
    if (!validCount) continue;
    const score = validCount + quarterCount * 3;
    if (!best || score > best.score) best = { row, score };
  }
  return best?.row ?? null;
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

function findRow(sheet, label) {
  const wanted = normalize(label);
  for (let row = 1; row <= sheet.rowCount; row += 1) {
    for (const col of [1, 2, 3, 4, 5]) {
      if (normalize(cellValue(sheet.getCell(row, col))) === wanted) return row;
    }
  }
  return null;
}

function valuesMatch(actual, expected) {
  return typeof actual === "number" && Math.abs(actual - expected) <= 0.5;
}

function auditRowsFor(audit, cell, period) {
  const rows = [];
  audit.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    if (String(cellValue(row.getCell(2)) ?? "") !== cell) return;
    if (String(cellValue(row.getCell(5)) ?? "").toUpperCase() !== period.toUpperCase()) return;
    rows.push({
      concepts: String(cellValue(row.getCell(8)) ?? ""),
      labels: String(cellValue(row.getCell(9)) ?? ""),
      notes: String(cellValue(row.getCell(24)) ?? "")
    });
  });
  return rows;
}

async function main() {
  await postWorkbook({ apiUrl, ticker, inputWorkbook, outputWorkbook });

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(outputWorkbook);
  const model = workbook.getWorksheet("Model");
  const audit = workbook.getWorksheet("Mapping Audit");
  if (!model) throw new Error("Output workbook does not contain a Model sheet.");
  if (!audit) throw new Error("Output workbook does not contain a Mapping Audit sheet.");

  const errors = [];
  const col = findPeriodColumn(model, "1Q26");
  if (!col) throw new Error("Could not find 1Q26 column in Model sheet.");

  for (const [label, expected] of Object.entries(expected1Q26)) {
    const row = findRow(model, label);
    if (!row) {
      errors.push(`Could not find row "${label}" in Model sheet.`);
      continue;
    }
    const actual = cellValue(model.getCell(row, col));
    if (!valuesMatch(actual, expected)) {
      errors.push(`1Q26 ${label}: expected ${expected}, got ${actual ?? "[blank]"}.`);
    }
  }

  const currentLiabilitiesExDebt = cellValue(model.getCell(findRow(model, "Total Current Liabilities (Excl. Debt)"), col));
  const revolver = cellValue(model.getCell(findRow(model, "Revolver"), col));
  if (!valuesMatch(currentLiabilitiesExDebt + revolver, 36634)) {
    errors.push(`1Q26 current liabilities excluding debt plus revolver should equal EDGAR current liabilities 36634, got ${currentLiabilitiesExDebt + revolver}.`);
  }

  const revolverAudit = auditRowsFor(audit, "U139", "1Q26");
  if (!revolverAudit.some((row) => /DebtCurrent=4000mm/.test(row.concepts))) {
    errors.push("Mapping Audit U139 should map 1Q26 Revolver to EDGAR DebtCurrent=4000mm.");
  }

  const otherNonCurrentAudit = auditRowsFor(audit, "U142", "1Q26");
  if (!otherNonCurrentAudit.some((row) => /AccruedIncomeTaxesNoncurrent=5289mm/.test(row.concepts) && /OtherLiabilitiesNoncurrent=4085mm/.test(row.concepts))) {
    errors.push("Mapping Audit U142 should include only LLY 1Q26 non-current income taxes payable and other non-current liabilities.");
  }
  if (otherNonCurrentAudit.some((row) => /DebtCurrent=4000mm/.test(row.concepts))) {
    errors.push("Mapping Audit U142 should not bury current debt inside other non-current liabilities.");
  }

  const commonApicAudit = auditRowsFor(audit, "U147", "1Q26");
  if (!commonApicAudit.some((row) => /CommonStockValue=590mm/.test(row.concepts) && /AdditionalPaidInCapitalCommonStock=6921mm/.test(row.concepts) && /OtherAdditionalCapital=19mm/.test(row.concepts))) {
    errors.push("Mapping Audit U147 should classify common stock, APIC, and other additional capital into Common Stock & APIC.");
  }

  const treasuryAudit = auditRowsFor(audit, "U149", "1Q26");
  if (!treasuryAudit.some((row) => /CommonStockSharesHeldInEmployeeTrust=3013mm/.test(row.concepts))) {
    errors.push("Mapping Audit U149 should classify employee trust shares as contra-equity in Treasury Stock.");
  }

  if (errors.length) {
    console.error(errors.join("\n"));
    throw new Error(`LLY balance sheet classification regression failed with ${errors.length} issue(s).`);
  }

  console.log(`LLY balance sheet classification regression passed: ${outputWorkbook}`);
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
