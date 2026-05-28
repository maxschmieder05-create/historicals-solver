const path = require("node:path");
const ExcelJS = require("exceljs");
const { postWorkbook } = require("./fill-workbook-api");

const repoRoot = path.resolve(__dirname, "..");
const ticker = process.env.BALANCE_SHEET_TICKER || process.env.COST_TICKER || "COST";
const cik = process.env.BALANCE_SHEET_CIK || process.env.COST_CIK || "0000909832";
const inputWorkbook =
  process.env.BALANCE_SHEET_INPUT_WORKBOOK || process.env.COST_INPUT_WORKBOOK || "/Users/maxschmieder/Downloads/COST_historicals_filled.xlsx";
const outputWorkbook =
  process.env.BALANCE_SHEET_OUTPUT_WORKBOOK ||
  process.env.COST_OUTPUT_WORKBOOK ||
  path.join(repoRoot, `tmp/${ticker.toLowerCase()}-balance-sheet-instant-regression-output.xlsx`);
const apiUrl = process.env.FILL_API_URL || "http://localhost:3000/api/fill-model";

const secHeaders = {
  "User-Agent": process.env.SEC_USER_AGENT || "HistoricalsSolver regression contact@example.com"
};

const directBalanceSheetChecks = [
  {
    row: 120,
    label: "cash and cash equivalents",
    concepts: [
      "CashAndCashEquivalentsAtCarryingValue",
      "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents",
      "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalentsIncludingDisposalGroupAndDiscontinuedOperations"
    ]
  },
  {
    row: 121,
    label: "receivables",
    concepts: ["AccountsReceivableNetCurrent", "AccountsReceivableNet", "TradeAccountsReceivableNetCurrent", "ReceivablesNetCurrent"]
  },
  { row: 122, label: "inventory", concepts: ["InventoryNet"] },
  { row: 124, label: "total current assets", concepts: ["AssetsCurrent"], formulaResult: true },
  {
    row: 126,
    label: "PP&E, net",
    concepts: [
      "PropertyPlantAndEquipmentAndOperatingLeaseRightofUseAssetAfterAccumulatedDepreciationAndAmortization",
      "PropertyPlantAndEquipmentAndFinanceLeaseRightOfUseAssetAfterAccumulatedDepreciationAndAmortization",
      "PropertyPlantAndEquipmentNet",
      "PropertyAndEquipmentNet"
    ]
  },
  { row: 132, label: "total assets", concepts: ["Assets"], formulaResult: true },
  { row: 134, label: "accounts payable", concepts: ["AccountsPayableCurrent", "AccountsPayableTradeCurrent"] },
  { row: 145, label: "total liabilities", concepts: ["Liabilities"], formulaResult: true },
  { row: 148, label: "retained earnings", concepts: ["RetainedEarningsAccumulatedDeficit"] },
  { row: 150, label: "AOCI", concepts: ["AccumulatedOtherComprehensiveIncomeLossNetOfTax"] },
  { row: 151, label: "stockholders' equity", concepts: ["StockholdersEquity"], formulaResult: true }
];

const costcoNetOtherIncomeBridgeChecks = [
  { period: "1Q25", row: 39, label: "interest expense separated from net other income", expected: 0 },
  { period: "2Q25", row: 39, label: "interest expense separated from net other income", expected: 0 },
  { period: "3Q25", row: 39, label: "interest expense separated from net other income", expected: 0 },
  { period: "4Q25", row: 39, label: "interest expense separated from net other income", expected: 0 },
  { period: "2025", row: 39, label: "interest expense separated from net other income", expected: 0 },
  { period: "3Q25", row: 41, label: "other non-operating income bridge", expected: 50 },
  { period: "4Q25", row: 41, label: "other non-operating income bridge", expected: 169 },
  { period: "2025", row: 41, label: "other non-operating income bridge", expected: 435 },
  { period: "3Q25", row: 42, label: "pre-tax income", expected: 2580 },
  { period: "4Q25", row: 42, label: "pre-tax income", expected: 3510 },
  { period: "2025", row: 42, label: "pre-tax income", expected: 10818 },
  { period: "3Q25", row: 44, label: "income tax expense", expected: -677 },
  { period: "4Q25", row: 44, label: "income tax expense", expected: -900 },
  { period: "2025", row: 44, label: "income tax expense", expected: -2719 },
  { period: "3Q25", row: 45, label: "net income", expected: 1903 },
  { period: "4Q25", row: 45, label: "net income", expected: 2610 },
  { period: "2025", row: 45, label: "net income", expected: 8099 }
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

async function main() {
  await postWorkbook({ apiUrl, ticker, inputWorkbook, outputWorkbook });

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(outputWorkbook);
  const facts = await fetchCompanyFacts();
  const errors = [];
  const model = workbook.getWorksheet("Model");
  if (!model) throw new Error("Output workbook does not contain a Model sheet.");
  const periods = modelQuarterPeriods(model);

  for (const { period, col } of periods) {
    for (const check of directBalanceSheetChecks) {
      const expected = expectedConceptValue(facts, period, check.concepts);
      if (expected === null) continue;
      const address = `${columnLetter(col)}${check.row}`;
      const actual = cellValue(model.getCell(check.row, col));
      if (!valuesMatch(actual, expected)) {
        errors.push(`${period} ${check.label} Model!${address}: expected ${expected}, got ${actual ?? "[blank]"}.`);
      }
    }

    const checkAddress = `${columnLetter(col)}158`;
    const balanceSheetCheck = cellValue(model.getCell(158, col));
    if (!valuesMatch(balanceSheetCheck, 0)) {
      errors.push(`${period} balance sheet check Model!${checkAddress}: expected 0, got ${balanceSheetCheck ?? "[blank]"}.`);
    }
  }

  const noFrameInstantPeriods = periods.filter(({ period }) => {
    const fact = expectedFact(facts, period, "InventoryNet");
    return fact && !fact.frame && fact.end && !fact.start;
  });
  if (!noFrameInstantPeriods.length) {
    errors.push(`${ticker}: no no-frame InventoryNet instant facts were available to exercise the SEC instant classification regression.`);
  }

  if (ticker.toUpperCase() === "COST") {
    errors.push(...validateCostcoNetOtherIncomeBridge(model));
  }

  if (errors.length) {
    console.error(errors.join("\n"));
    throw new Error(`${ticker} balance-sheet instant regression failed with ${errors.length} issue(s).`);
  }

  console.log(`${ticker} balance-sheet instant regression passed across ${periods.length} quarter column(s): ${outputWorkbook}`);
}

function validateCostcoNetOtherIncomeBridge(sheet) {
  const errors = [];
  for (const check of costcoNetOtherIncomeBridgeChecks) {
    const col = findPeriodColumn(sheet, check.period);
    if (!col) {
      errors.push(`COST ${check.label}: could not find ${check.period} column.`);
      continue;
    }
    const actual = cellValue(sheet.getCell(check.row, col));
    if (!valuesMatch(actual, check.expected)) {
      errors.push(`COST ${check.period} ${check.label} Model!${columnLetter(col)}${check.row}: expected ${check.expected}, got ${actual ?? "[blank]"}.`);
    }
  }
  return errors;
}

function findPeriodColumn(sheet, period) {
  for (let col = 1; col <= sheet.columnCount; col += 1) {
    if (String(sheet.getCell(25, col).text || "").replace(/[’']/g, "").trim().toUpperCase() === period.toUpperCase()) return col;
  }
  return null;
}

async function fetchCompanyFacts() {
  const response = await fetch(`https://data.sec.gov/api/xbrl/companyfacts/CIK${cik}.json`, { headers: secHeaders });
  if (!response.ok) throw new Error(`Could not load ${ticker} SEC companyfacts: ${response.status} ${response.statusText}`);
  return response.json();
}

function modelQuarterPeriods(sheet) {
  const periods = [];
  for (let col = 1; col <= sheet.columnCount; col += 1) {
    const label = String(sheet.getCell(25, col).text || "").replace(/[’']/g, "").trim();
    if (!/^[1-4]Q\d{2}$/i.test(label)) continue;
    if (/e$/i.test(String(sheet.getCell(25, col).text || ""))) continue;
    periods.push({ period: label.toUpperCase(), col });
  }
  return periods;
}

function expectedConceptValue(payload, period, concepts) {
  for (const concept of concepts) {
    const fact = expectedFact(payload, period, concept);
    if (fact) return fact.val / 1_000_000;
  }
  return null;
}

function expectedFact(payload, period, concept) {
  const unitFacts = payload.facts?.["us-gaap"]?.[concept]?.units?.USD;
  if (!Array.isArray(unitFacts)) return null;
  const candidates = unitFacts.filter((fact) => factMatchesPeriod(fact, period) && fact.end && !fact.start && (fact.form === "10-Q" || fact.form === "10-K"));
  return candidates.sort(compareExpectedFactPreference)[0] ?? null;
}

function factMatchesPeriod(fact, period) {
  const quarter = Number(period[0]);
  const fiscalYear = 2000 + Number(period.slice(2));
  if (fact.fy !== fiscalYear) return false;
  if (quarter === 4) return fact.fp === "FY" || fact.fp === "Q4";
  return fact.fp === `Q${quarter}`;
}

function compareExpectedFactPreference(a, b) {
  const endCompare = String(b.end).localeCompare(String(a.end));
  if (endCompare !== 0) return endCompare;
  const formCompare = formScore(b.form) - formScore(a.form);
  if (formCompare !== 0) return formCompare;
  return String(b.filed).localeCompare(String(a.filed));
}

function formScore(form) {
  if (form === "10-K") return 2;
  if (form === "10-Q") return 1;
  return 0;
}

function columnLetter(col) {
  let letter = "";
  while (col > 0) {
    const modulo = (col - 1) % 26;
    letter = String.fromCharCode(65 + modulo) + letter;
    col = Math.floor((col - modulo) / 26);
  }
  return letter;
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
