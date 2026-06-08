const path = require("node:path");
const ExcelJS = require("exceljs");
const { postWorkbook } = require("./fill-workbook-api");

const repoRoot = path.resolve(__dirname, "..");
const inputWorkbook = process.env.INTC_INPUT_WORKBOOK || "/Users/maxschmieder/Downloads/INTC_historicals_filled.xlsx";
const outputWorkbook = process.env.INTC_OUTPUT_WORKBOOK || path.join(repoRoot, "tmp", "intc-income-statement-classification-output.xlsx");
const apiUrl = process.env.FILL_API_URL || "http://localhost:3000/api/fill-model";
const ticker = process.env.INTC_TICKER || "INTC";

const expectedCells = [
  ["Model", "U36", -3136, "1Q26 EBIT should equal EDGAR operating income"],
  ["Model", "U38", 0, "1Q26 interest income should stay zero without a standalone primary-statement line"],
  ["Model", "U39", 0, "1Q26 interest expense should stay zero because the primary statement reports a combined interest-and-other line"],
  ["Model", "U40", 0, "1Q26 goodwill impairment should stay zero without a standalone primary-statement line"],
  ["Model", "U41", -810, "1Q26 other non-operating should sum reported below-EBIT primary-statement rows"],
  ["Model", "U42", -3946, "1Q26 pre-tax income"],
  ["Model", "U44", -335, "1Q26 income tax expense"],
  ["Model", "U45", -4281, "1Q26 net income should equal EDGAR primary-statement net income"]
];

function cellValue(cell) {
  const value = cell.value;
  if (typeof value === "number") return value;
  if (value && typeof value === "object") {
    if (typeof value.result === "number") return value.result;
    if (typeof value.result === "string") return Number(value.result);
  }
  return value;
}

function valuesMatch(actual, expected) {
  return typeof actual === "number" && Math.abs(actual - expected) <= 0.5;
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
  const errors = [];

  for (const [sheetName, address, expected, label] of expectedCells) {
    const sheet = workbook.getWorksheet(sheetName);
    const actual = sheet ? cellValue(sheet.getCell(address)) : null;
    if (!valuesMatch(actual, expected)) {
      errors.push(`${label} ${sheetName}!${address}: expected ${expected}, got ${actual ?? "[blank]"}.`);
    }
  }

  const audit = workbook.getWorksheet("Mapping Audit");
  if (!audit) {
    errors.push("Mapping Audit sheet was not found.");
  } else {
    const period = "1Q26";
    const interestExpenseRows = auditRowsFor(audit, "U39", period);
    if (!interestExpenseRows.some((row) => row.value === 0 && /InterestExpenseNotReported/.test(row.concepts))) {
      errors.push("Mapping Audit U39 should document that no standalone primary-statement interest expense line was reported.");
    }
    if (interestExpenseRows.some((row) => /InterestExpenseNonoperating/.test(row.concepts))) {
      errors.push("Mapping Audit U39 should not map INTC note-level InterestExpenseNonoperating into standalone interest expense.");
    }

    const goodwillRows = auditRowsFor(audit, "U40", period);
    if (!goodwillRows.some((row) => row.value === 0 && /GoodwillImpairmentNotReported/.test(row.concepts))) {
      errors.push("Mapping Audit U40 should document the no-standalone-goodwill-impairment zero policy.");
    }
    if (goodwillRows.some((row) => /GoodwillImpairmentLoss/.test(row.concepts))) {
      errors.push("Mapping Audit U40 should not map INTC note-level GoodwillImpairmentLoss into Goodwill Impairment.");
    }

    const otherNonOperatingRows = auditRowsFor(audit, "U41", period);
    if (!otherNonOperatingRows.some((row) => /NonoperatingIncomeExpense/.test(row.concepts) && /EquitySecuritiesFvNiGainLoss/.test(row.concepts))) {
      errors.push("Mapping Audit U41 should cite INTC's reported NonoperatingIncomeExpense and EquitySecuritiesFvNiGainLoss below-EBIT rows.");
    }
    if (otherNonOperatingRows.some((row) => /OtherNonOperatingIncomeExpenseFromPreTaxBridge|GoodwillImpairmentLoss|InterestExpenseNonoperating/.test(row.concepts))) {
      errors.push("Mapping Audit U41 should not use goodwill, standalone interest expense, or a pre-tax residual bridge as other non-operating support.");
    }

    const netIncomeRows = auditRowsFor(audit, "U45", period);
    if (!netIncomeRows.some((row) => /ProfitLoss/.test(row.concepts))) {
      errors.push("Mapping Audit U45 should cite ProfitLoss for INTC primary-statement net income.");
    }
    if (netIncomeRows.some((row) => /NetIncomeLoss=/.test(row.concepts))) {
      errors.push("Mapping Audit U45 should not use parent-attributable NetIncomeLoss for the broad Net Income row when ProfitLoss is reported.");
    }
  }

  if (errors.length) {
    console.error(errors.join("\n"));
    throw new Error(`INTC income statement classification regression failed with ${errors.length} issue(s).`);
  }

  console.log(`INTC income statement classification regression passed: ${outputWorkbook}`);
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
