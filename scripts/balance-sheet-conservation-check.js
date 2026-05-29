const path = require("node:path");
const ExcelJS = require("exceljs");
const { postWorkbook } = require("./fill-workbook-api");

const repoRoot = path.resolve(__dirname, "..");
const apiUrl = process.env.FILL_API_URL || "http://localhost:3000/api/fill-model";

const cases = [
  {
    key: "AMZN",
    ticker: "AMZN",
    cik: "0001018724",
    inputWorkbook: process.env.AMZN_INPUT_WORKBOOK || "/Users/maxschmieder/Downloads/AMZN_historicals_filled (4).xlsx"
  },
  {
    key: "GOOG",
    ticker: "GOOG",
    cik: "0001652044",
    inputWorkbook: process.env.GOOG_INPUT_WORKBOOK || "/Users/maxschmieder/Downloads/Owl Fund Integrated Model Template (03-Sep-2025)_v25 (3).xlsx"
  },
  {
    key: "COST",
    ticker: "COST",
    cik: "0000909832",
    inputWorkbook: process.env.COST_INPUT_WORKBOOK || "/Users/maxschmieder/Downloads/COST_historicals_filled.xlsx"
  },
  {
    key: "MCK",
    ticker: "MCK",
    cik: "0000927653",
    inputWorkbook: process.env.MCK_INPUT_WORKBOOK || "/Users/maxschmieder/Downloads/McKesson Corporation (MCK)_Valuation Workbook (12-May-2026).xlsx"
  },
  {
    key: "SYK",
    ticker: "SYK",
    cik: "0000310764",
    inputWorkbook: process.env.SYK_INPUT_WORKBOOK || "/Users/maxschmieder/Downloads/Stryker Corp (SYK)_Valuation Workbook (13-Mar-2026).xlsx"
  }
];

const enabledCases = new Set(
  (process.env.BALANCE_SHEET_CONSERVATION_CASES || "AMZN,GOOG")
    .split(",")
    .map((item) => item.trim().toUpperCase())
    .filter(Boolean)
);

const secHeaders = {
  "User-Agent": process.env.SEC_USER_AGENT || "HistoricalsSolver conservation regression contact@example.com"
};

const checks = [
  { name: "total assets", labels: ["Total Assets"], concepts: ["Assets"] },
  { name: "total liabilities", labels: ["Total Liabilities"], concepts: ["Liabilities"] },
  {
    name: "shareholders' equity",
    labels: ["Total Shareholder's Equity", "Total Shareholders' Equity", "Total Shareholders Equity", "Total Stockholders' Equity", "Total Stockholders Equity", "Total Equity"],
    concepts: ["StockholdersEquity", "StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest"]
  },
  { name: "liabilities plus equity", labels: ["Total Liabilities & Shareholder's Equity", "Total Liabilities and Shareholder's Equity", "Total Liabilities & Stockholders' Equity"], concepts: ["Assets"] },
  { name: "balance sheet check", labels: ["Balance Sheet Check"], expected: 0 }
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

function valuesMatch(actual, expected, tolerance = 0.5) {
  return typeof actual === "number" && Number.isFinite(actual) && Math.abs(actual - expected) <= tolerance;
}

async function main() {
  const selected = cases.filter((item) => enabledCases.has(item.key));
  if (!selected.length) throw new Error("No balance sheet conservation cases were selected.");

  const allErrors = [];
  for (const testCase of selected) {
    const outputWorkbook = path.join(repoRoot, "tmp", `${testCase.key.toLowerCase()}-balance-sheet-conservation-output.xlsx`);
    await postWorkbook({ apiUrl, ticker: testCase.ticker, inputWorkbook: testCase.inputWorkbook, outputWorkbook });
    const errors = await validateCase(testCase, outputWorkbook);
    allErrors.push(...errors);
  }

  if (allErrors.length) {
    console.error(allErrors.join("\n"));
    throw new Error(`Balance sheet conservation failed with ${allErrors.length} issue(s).`);
  }

  console.log(`Balance sheet conservation passed for ${selected.map((item) => item.key).join(", ")}.`);
}

async function validateCase(testCase, outputWorkbook) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(outputWorkbook);
  const model = workbook.getWorksheet("Model");
  if (!model) return [`${testCase.key}: output workbook is missing Model sheet.`];

  const facts = await fetchCompanyFacts(testCase.cik);
  const periods = modelQuarterPeriods(model);
  const errors = [];
  for (const { period, col } of periods) {
    for (const check of checks) {
      const row = findRowInBalanceSheet(model, check.labels);
      if (!row) {
        errors.push(`${testCase.key} ${period}: could not find ${check.name} row.`);
        continue;
      }
      const expected = check.expected ?? expectedConceptValue(facts, period, check.concepts);
      if (expected === null) continue;
      const actual = cellValue(model.getCell(row, col));
      if (!valuesMatch(actual, expected)) {
        errors.push(`${testCase.key} ${period} ${check.name} ${model.getCell(row, col).address}: expected ${expected}, got ${actual ?? "[blank]"}.`);
      }
    }
  }
  return errors;
}

async function fetchCompanyFacts(cik) {
  const response = await fetch(`https://data.sec.gov/api/xbrl/companyfacts/CIK${cik}.json`, { headers: secHeaders });
  if (!response.ok) throw new Error(`Could not load SEC companyfacts for CIK ${cik}: ${response.status} ${response.statusText}`);
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

function findRowInBalanceSheet(sheet, labels) {
  const wanted = new Set(labels.map(normalize));
  for (let row = 1; row <= sheet.rowCount; row += 1) {
    const label = rowLabel(sheet, row);
    if (row > 80 && wanted.has(normalize(label))) return row;
  }
  return null;
}

function rowLabel(sheet, row) {
  for (let col = 1; col <= Math.min(sheet.columnCount, 8); col += 1) {
    const text = String(sheet.getCell(row, col).text || "").trim();
    if (/^x$/i.test(text)) continue;
    if (text) return text;
  }
  return "";
}

function normalize(value) {
  return String(value).toLowerCase().replace(/[^a-z0-9]/g, "");
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

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
