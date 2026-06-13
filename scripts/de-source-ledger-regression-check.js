const path = require("node:path");
const ExcelJS = require("exceljs");
const { postWorkbook } = require("./fill-workbook-api");

const repoRoot = path.resolve(__dirname, "..");
const inputWorkbook = process.env.DE_INPUT_WORKBOOK || "/Users/maxschmieder/Downloads/DE_historicals_filled.xlsx";
const outputWorkbook = process.env.DE_OUTPUT_WORKBOOK || path.join(repoRoot, "tmp", "de-source-ledger-regression-output.xlsx");
const apiUrl = process.env.FILL_API_URL || "http://localhost:3000/api/fill-model";

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

function assertTie(errors, label, actual, expected) {
  if (!valuesMatch(actual, expected)) errors.push(`${label}: expected ${expected}, got ${actual ?? "[blank]"}.`);
}

function ledgerRowsByCell(ledger, cell) {
  const header = ledger.getRow(1).values;
  const cols = {};
  header.forEach((value, index) => {
    if (value) cols[String(value)] = index;
  });
  const rows = [];
  for (let rowNumber = 2; rowNumber <= ledger.rowCount; rowNumber += 1) {
    const row = ledger.getRow(rowNumber);
    if (String(cellValue(row.getCell(cols.cell)) ?? "") !== cell) continue;
    rows.push({
      value: cellValue(row.getCell(cols.value)),
      status: String(cellValue(row.getCell(cols["mapping status"])) ?? ""),
      ticker: String(cellValue(row.getCell(cols.ticker)) ?? ""),
      cik: String(cellValue(row.getCell(cols.CIK)) ?? ""),
      accession: String(cellValue(row.getCell(cols["accession number"])) ?? ""),
      tag: String(cellValue(row.getCell(cols["source XBRL tag"])) ?? ""),
      label: String(cellValue(row.getCell(cols["source line item label"])) ?? "")
    });
  }
  return rows;
}

function requireLedger(errors, ledger, cell, predicate, description) {
  const rows = ledgerRowsByCell(ledger, cell);
  if (!rows.some(predicate)) errors.push(`${cell}: missing Source Ledger entry for ${description}. Entries: ${JSON.stringify(rows)}`);
}

async function main() {
  await postWorkbook({ apiUrl, ticker: "DE", inputWorkbook, outputWorkbook });

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(outputWorkbook);
  const model = workbook.getWorksheet("Model");
  const ledger = workbook.getWorksheet("Source Ledger");
  const errors = [];

  if (!model) errors.push("Model sheet was not found.");
  if (!ledger) errors.push("Source Ledger sheet was not found.");
  if (errors.length) throw new Error(errors.join("\n"));

  const col = findPeriodColumn(model, "1Q26");
  if (!col) throw new Error("Could not find 1Q26 column.");
  const rows = incomeStatementRows(model);
  const row = (label) => rows.get(normalize(label));
  const valueFor = (label) => {
    const rowNumber = row(label);
    if (!rowNumber) {
      errors.push(`Could not find income-statement row ${label}.`);
      return null;
    }
    return cellValue(model.getCell(rowNumber, col));
  };

  assertTie(errors, "1Q26 Revenue", valueFor("Revenue"), 9611);
  assertTie(errors, "1Q26 COGS", valueFor("Cost of Goods Sold"), -6280);
  assertTie(errors, "1Q26 SG&A", valueFor("Selling, General & Administration (SG&A)"), -972);
  assertTie(errors, "1Q26 R&D", valueFor("Research & Development (R&D)"), -554);
  assertTie(errors, "1Q26 Other Operating", valueFor("Other Operating Income (Expense)"), -250);
  assertTie(errors, "1Q26 EBIT", valueFor("EBIT"), 1555);
  assertTie(errors, "1Q26 Interest Expense", valueFor("Interest (Expense)"), -719);
  assertTie(errors, "1Q26 Other Non-Operating", valueFor("Other Non-Operating Income (Expense)"), 0);
  assertTie(errors, "1Q26 Pre-Tax Income", valueFor("Pre-Tax Income (Loss)"), 836);
  assertTie(errors, "1Q26 Tax", valueFor("Income Tax Benefit (Expense)"), -196);
  assertTie(errors, "1Q26 Net Income", valueFor("Net Income (Loss)"), 655);
  assertTie(errors, "1Q26 Post-Tax Adjustments", valueFor("Post-Tax Adjustments"), 15);
  assertTie(errors, "1Q26 NCI", valueFor("Income (Loss) due to Non-Controlling Interest"), 1);
  assertTie(errors, "1Q26 Adj. Net Income", valueFor("Adj. Net Income (Loss)"), 656);

  const cogsRow = row("Cost of Goods Sold");
  if (cogsRow && valuesMatch(cellValue(model.getCell(cogsRow, col)), 16292)) {
    errors.push("1Q26 COGS still equals stale prior-run value 16,292.");
  }

  const netIncomeRow = row("Net Income (Loss)");
  const adjustedNetIncomeRow = row("Adj. Net Income (Loss)");
  const colLetter = model.getColumn(col).letter;
  if (
    netIncomeRow &&
    formula(model.getCell(netIncomeRow, col)) !==
      `${colLetter}${row("Pre-Tax Income (Loss)")}+${colLetter}${row("Income Tax Benefit (Expense)")}+${colLetter}${row("Post-Tax Adjustments")}`
  ) {
    errors.push(`1Q26 net-income formula did not include post-tax equity-method bridge: ${formula(model.getCell(netIncomeRow, col))}`);
  }
  if (adjustedNetIncomeRow && !/^SUM\(.+,.+\)$/i.test(formula(model.getCell(adjustedNetIncomeRow, col)))) {
    errors.push(`1Q26 adjusted-net-income formula was not rewritten to avoid the post-tax double count: ${formula(model.getCell(adjustedNetIncomeRow, col))}`);
  }

  requireLedger(
    errors,
    ledger,
    `${model.getColumn(col).letter}${cogsRow}`,
    (entry) =>
      entry.status === "validated_current_company_derived_value" &&
      entry.ticker === "DE" &&
      entry.cik === "0000315189" &&
      entry.accession === "0001104659-26-020158" &&
      entry.tag.includes("CostOfRevenueDerivedFromCostsAndExpenses") &&
      entry.tag.includes("OtherCostAndExpenseOperating"),
    "current-company derived COGS from SEC costs and expenses"
  );
  requireLedger(
    errors,
    ledger,
    `${model.getColumn(col).letter}${row("Other Non-Operating Income (Expense)")}`,
    (entry) => entry.status === "explicit_zero_no_source_disclosed" && entry.tag.includes("OtherNonOperatingIncomeExpenseNotReported"),
    "explicit zero other non-operating row"
  );
  requireLedger(
    errors,
    ledger,
    `${model.getColumn(col).letter}${row("Post-Tax Adjustments")}`,
    (entry) => entry.status === "explicit_current_sec_source" && entry.tag.includes("IncomeLossFromEquityMethodInvestments"),
    "post-tax equity-method income"
  );
  requireLedger(
    errors,
    ledger,
    `${model.getColumn(col).letter}${row("Income (Loss) due to Non-Controlling Interest")}`,
    (entry) => entry.status === "explicit_current_sec_source" && entry.tag.includes("NetIncomeLossAttributableToNoncontrollingInterest"),
    "noncontrolling-interest bridge"
  );

  if (errors.length) throw new Error(`DE source-ledger regression failed with ${errors.length} issue(s):\n${errors.join("\n")}`);
  console.log(`DE source-ledger regression passed: ${outputWorkbook}`);
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
