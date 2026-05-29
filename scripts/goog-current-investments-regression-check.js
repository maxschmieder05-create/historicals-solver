const path = require("node:path");
const ExcelJS = require("exceljs");
const { postWorkbook } = require("./fill-workbook-api");

const repoRoot = path.resolve(__dirname, "..");
const inputWorkbook =
  process.env.GOOG_INPUT_WORKBOOK || "/Users/maxschmieder/Downloads/Owl Fund Integrated Model Template (03-Sep-2025)_v25 (3).xlsx";
const outputWorkbook = process.env.GOOG_OUTPUT_WORKBOOK || path.join(repoRoot, "tmp", "goog-current-investments-regression-output.xlsx");
const apiUrl = process.env.FILL_API_URL || "http://localhost:3000/api/fill-model";
const ticker = process.env.GOOG_TICKER || "GOOG";
const cik = process.env.GOOG_CIK || "0001652044";

const secHeaders = {
  "User-Agent": process.env.SEC_USER_AGENT || "HistoricalsSolver regression contact@example.com"
};

const checks = [
  { row: 120, label: "cash plus current marketable securities", resolver: expectedCashAndCurrentInvestments },
  { row: 123, label: "prepaid and other current assets excluding investments already in cash", resolver: expectedOtherCurrentAssets },
  { row: 124, label: "total current assets", concepts: ["AssetsCurrent"] },
  { row: 132, label: "total assets", concepts: ["Assets"] },
  { row: 156, label: "total liabilities plus shareholder's equity", concepts: ["Assets"] },
  { row: 158, label: "balance sheet check", resolver: () => 0 }
];

const flowChecks = [
  {
    row: 28,
    label: "revenue",
    concepts: ["Revenues", "RevenueFromContractWithCustomerExcludingAssessedTax", "SalesRevenueNet"]
  }
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
  const model = workbook.getWorksheet("Model");
  if (!model) throw new Error("Output workbook does not contain a Model sheet.");

  const errors = [];
  for (const { period, col } of modelQuarterPeriods(model)) {
    for (const check of flowChecks) {
      const expected = expectedDurationConceptValue(facts, period, check.concepts);
      if (expected === null) continue;
      const actual = cellValue(model.getCell(check.row, col));
      if (!valuesMatch(actual, expected)) {
        errors.push(`${period} ${check.label} Model!${columnLetter(col)}${check.row}: expected ${expected}, got ${actual ?? "[blank]"}.`);
      }
    }

    for (const check of checks) {
      const expected =
        check.resolver?.(facts, period) ??
        (check.concepts ? expectedConceptValue(facts, period, check.concepts) : null);
      if (expected === null) continue;
      const actual = cellValue(model.getCell(check.row, col));
      if (!valuesMatch(actual, expected)) {
        errors.push(`${period} ${check.label} Model!${columnLetter(col)}${check.row}: expected ${expected}, got ${actual ?? "[blank]"}.`);
      }
    }
  }

  if (errors.length) {
    console.error(errors.join("\n"));
    throw new Error(`GOOG current-investments regression failed with ${errors.length} issue(s).`);
  }

  console.log(`GOOG current-investments regression passed: ${outputWorkbook}`);
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

function expectedCashAndCurrentInvestments(payload, period) {
  const cash = expectedConceptValue(payload, period, [
    "CashAndCashEquivalentsAtCarryingValue",
    "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents",
    "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalentsIncludingDisposalGroupAndDiscontinuedOperations"
  ]);
  const investments = expectedConceptValue(payload, period, [
    "ShortTermInvestments",
    "MarketableSecuritiesCurrent",
    "AvailableForSaleSecuritiesCurrent"
  ]);
  if (cash === null) return null;
  return cash + (investments ?? 0);
}

function expectedOtherCurrentAssets(payload, period) {
  const currentAssets = expectedConceptValue(payload, period, ["AssetsCurrent"]);
  const cashAndInvestments = expectedCashAndCurrentInvestments(payload, period);
  const receivables = expectedConceptValue(payload, period, [
    "AccountsReceivableNetCurrent",
    "AccountsReceivableNet",
    "TradeAccountsReceivableNetCurrent",
    "ReceivablesNetCurrent"
  ]) ?? 0;
  const inventory = expectedConceptValue(payload, period, ["InventoryNet"]) ?? 0;
  if (currentAssets === null || cashAndInvestments === null) return null;
  return currentAssets - cashAndInvestments - receivables - inventory;
}

function expectedConceptValue(payload, period, concepts) {
  for (const concept of concepts) {
    const fact = expectedFact(payload, period, concept);
    if (fact) return fact.val / 1_000_000;
  }
  return null;
}

function expectedDurationConceptValue(payload, period, concepts) {
  for (const concept of concepts) {
    const fact = expectedDurationFact(payload, period, concept);
    if (fact) return fact.val / 1_000_000;
    if (/^4Q\d{2}$/i.test(period)) {
      const fourthQuarter = expectedFourthQuarterDurationValue(payload, period, concept);
      if (fourthQuarter !== null) return fourthQuarter;
    }
  }
  return null;
}

function expectedFourthQuarterDurationValue(payload, period, concept) {
  const year = period.slice(2);
  const annual = expectedDurationFact(payload, `FY${year}`, concept);
  const q1 = expectedDurationFact(payload, `1Q${year}`, concept);
  const q2 = expectedDurationFact(payload, `2Q${year}`, concept);
  const q3 = expectedDurationFact(payload, `3Q${year}`, concept);
  if (!annual || !q1 || !q2 || !q3) return null;
  return (annual.val - q1.val - q2.val - q3.val) / 1_000_000;
}

function expectedDurationFact(payload, period, concept) {
  const unitFacts = payload.facts?.["us-gaap"]?.[concept]?.units?.USD;
  if (!Array.isArray(unitFacts)) return null;
  const candidates = unitFacts.filter((fact) => durationFactMatchesPeriod(fact, period) && fact.start && fact.end && (fact.form === "10-Q" || fact.form === "10-K"));
  return candidates.sort(compareExpectedFactPreference)[0] ?? null;
}

function durationFactMatchesPeriod(fact, period) {
  const fiscalYear = 2000 + Number(period.slice(2));
  if (fact.fy !== fiscalYear) return false;
  if (/^FY\d{2}$/i.test(period)) return fact.fp === "FY";
  const quarter = Number(period[0]);
  if (quarter === 4) return fact.fp === "Q4";
  return fact.fp === `Q${quarter}`;
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
