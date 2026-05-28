const path = require("node:path");
const fs = require("node:fs/promises");
const ExcelJS = require("exceljs");
const JSZip = require("jszip");
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
  { period: "1Q23", row: 38, label: "combined interest income left out of dedicated interest income", expected: 0 },
  { period: "1Q23", row: 39, label: "standalone interest expense", expected: -34 },
  { period: "1Q23", row: 41, label: "interest income and other, net", expected: 53 },
  { period: "1Q23", row: 42, label: "pre-tax income", expected: 1770 },
  { period: "1Q25", row: 39, label: "standalone interest expense", expected: -37 },
  { period: "2Q25", row: 39, label: "standalone interest expense", expected: -36 },
  { period: "3Q25", row: 39, label: "standalone interest expense", expected: -35 },
  { period: "4Q25", row: 39, label: "standalone interest expense", expected: -46 },
  { period: "2025", row: 39, label: "standalone interest expense", expected: -154 },
  { period: "3Q25", row: 41, label: "interest income and other, net", expected: 85 },
  { period: "4Q25", row: 41, label: "interest income and other, net", expected: 215 },
  { period: "2025", row: 41, label: "interest income and other, net", expected: 589 },
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
  await validateWorkbookRecalculationMetadata(outputWorkbook);
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
    errors.push(...validateCostcoPreTaxBridgeConsistency(model));
  }

  if (errors.length) {
    console.error(errors.join("\n"));
    throw new Error(`${ticker} balance-sheet instant regression failed with ${errors.length} issue(s).`);
  }

  console.log(`${ticker} balance-sheet instant regression passed across ${periods.length} quarter column(s): ${outputWorkbook}`);
}

async function validateWorkbookRecalculationMetadata(file) {
  const zip = await JSZip.loadAsync(await fs.readFile(file));
  const workbookXml = await zip.file("xl/workbook.xml")?.async("string");
  const errors = [];
  if (!workbookXml || !/<calcPr\b[^>]*calcMode="auto"/.test(workbookXml)) {
    errors.push("workbook.xml is missing automatic calculation mode.");
  }
  if (!workbookXml || !/<calcPr\b[^>]*fullCalcOnLoad="1"/.test(workbookXml) || !/<calcPr\b[^>]*forceFullCalc="1"/.test(workbookXml)) {
    errors.push("workbook.xml is not forcing a full formula recalculation on open.");
  }
  if (zip.file("xl/calcChain.xml")) {
    errors.push("xl/calcChain.xml should be removed so Excel rebuilds formula dependencies.");
  }

  const sharedStrings = await sharedStringValues(zip);
  let formulaCount = 0;
  let recalculationMarkedCount = 0;
  for (const path of Object.keys(zip.files).filter((item) => /^xl\/worksheets\/sheet\d+\.xml$/i.test(item))) {
    const xml = await zip.file(path)?.async("string");
    if (!xml) continue;
    for (const cellXml of xml.match(/<c\b[^>]*>[\s\S]*?<\/c>/g) ?? []) {
      const address = cellXml.match(/\br="([^"]+)"/)?.[1] ?? "?";
      if (/<f\b/.test(cellXml)) {
        formulaCount += 1;
        if (/<f\b[^>]*\bca="1"/.test(cellXml)) recalculationMarkedCount += 1;
        else errors.push(`${path}!${address}: formula is not marked for recalculation.`);
        const formula = unescapeXml(cellXml.match(/<f\b[^>]*>([\s\S]*?)<\/f>/)?.[1] ?? "");
        const cachedValue = formulaStringValue(cellXml, sharedStrings);
        if (!/#REF!?/i.test(formula) && cachedValue && /^#REF!?$/i.test(cachedValue)) {
          errors.push(`${path}!${address}: formula has stale cached #REF result.`);
        }
        continue;
      }
      const textValue = formulaStringValue(cellXml, sharedStrings);
      if (textValue && isPlainTextFormula(textValue)) {
        errors.push(`${path}!${address}: formula is stored as text.`);
      }
    }
  }
  if (!formulaCount) errors.push("No formula cells were found in the generated workbook.");
  if (!recalculationMarkedCount) errors.push("No formula cells were marked for recalculation.");
  if (errors.length) throw new Error(`Generated workbook recalculation metadata failed: ${errors.slice(0, 12).join(" | ")}`);
}

async function sharedStringValues(zip) {
  const values = new Map();
  const sharedStringsFile = zip.file("xl/sharedStrings.xml");
  if (!sharedStringsFile) return values;

  const xml = await sharedStringsFile.async("string");
  let index = 0;
  for (const match of xml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)) {
    const text = Array.from(match[1].matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g))
      .map((item) => unescapeXml(item[1]))
      .join("");
    values.set(index, text);
    index += 1;
  }
  return values;
}

function formulaStringValue(cellXml, sharedStrings) {
  if (/\bt=(["'])s\1/.test(cellXml)) {
    const sharedStringIndex = Number(cellXml.match(/<v>(\d+)<\/v>/)?.[1]);
    return Number.isFinite(sharedStringIndex) ? sharedStrings.get(sharedStringIndex) ?? null : null;
  }
  const value = cellXml.match(/<v>([\s\S]*?)<\/v>/)?.[1];
  if (value !== undefined) return unescapeXml(value);
  const inlineText = cellXml.match(/<is>\s*<t[^>]*>([\s\S]*?)<\/t>\s*<\/is>/)?.[1];
  return inlineText === undefined ? null : unescapeXml(inlineText);
}

function isPlainTextFormula(value) {
  return /^=\s*(?:[A-Z_@]|\d|[+\-.]|\()/i.test(value.trim());
}

function unescapeXml(value) {
  return value.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");
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

function validateCostcoPreTaxBridgeConsistency(sheet) {
  const errors = [];
  for (const { period, col } of modelQuarterPeriods(sheet)) {
    const ebit = cellValue(sheet.getCell(36, col));
    const interestIncome = cellValue(sheet.getCell(38, col));
    const interestExpense = cellValue(sheet.getCell(39, col));
    const goodwillImpairment = cellValue(sheet.getCell(40, col));
    const otherNonOperating = cellValue(sheet.getCell(41, col));
    const pretax = cellValue(sheet.getCell(42, col));
    const bridgeValues = [ebit, interestIncome, interestExpense, goodwillImpairment, otherNonOperating, pretax];
    if (!bridgeValues.every((value) => typeof value === "number")) {
      errors.push(`COST ${period} pre-tax bridge: missing numeric EBIT, interest, goodwill, other non-operating, or pre-tax value.`);
      continue;
    }
    if (interestExpense > 0) {
      errors.push(`COST ${period} standalone interest expense should be stored as a model reduction, got ${interestExpense}.`);
    }
    const expectedPreTax = ebit + interestIncome + interestExpense + goodwillImpairment + otherNonOperating;
    if (!valuesMatch(expectedPreTax, pretax)) {
      errors.push(`COST ${period} pre-tax bridge: expected ${expectedPreTax}, got ${pretax}.`);
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
