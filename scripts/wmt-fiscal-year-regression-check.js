const path = require("node:path");
const ExcelJS = require("exceljs");
const { postWorkbook } = require("./fill-workbook-api");

const repoRoot = path.resolve(__dirname, "..");
const ticker = process.env.WMT_TICKER || "WMT";
const cik = process.env.WMT_CIK || "0000104169";
const inputWorkbook = process.env.WMT_INPUT_WORKBOOK || "/Users/maxschmieder/Downloads/WMT_historicals_filled.xlsx";
const outputWorkbook = process.env.WMT_OUTPUT_WORKBOOK || path.join(repoRoot, "tmp", "wmt-fiscal-year-regression-output.xlsx");
const apiUrl = process.env.FILL_API_URL || "http://localhost:3000/api/fill-model";

const secHeaders = {
  "User-Agent": process.env.SEC_USER_AGENT || "HistoricalsSolver WMT regression contact@example.com"
};

const anchorRows = [
  { label: "Revenue", concepts: ["Revenues", "RevenueFromContractWithCustomerExcludingAssessedTax"], kind: "duration" },
  { label: "Cash & Cash Equivalents", concepts: ["CashAndCashEquivalentsAtCarryingValue"], kind: "instant" },
  { label: "Total Assets", concepts: ["Assets"], kind: "instant" },
  { label: "Total Liabilities & Shareholder's Equity", concepts: ["LiabilitiesAndStockholdersEquity", "Assets"], kind: "instant" }
];

const expectedIncomeStatementBridge = {
  "1Q25": {
    EBIT: 6841,
    "Interest Income": 114,
    "Interest (Expense)": -714,
    "Goodwill Impairment": 0,
    "Other Non-Operating Income (Expense)": 794,
    "Pre-Tax Income (Loss)": 7035
  },
  "2Q25": {
    EBIT: 7940,
    "Interest Income": 114,
    "Interest (Expense)": -679,
    "Goodwill Impairment": 0,
    "Other Non-Operating Income (Expense)": -1162,
    "Pre-Tax Income (Loss)": 6213
  },
  "1Q26": {
    EBIT: 7135,
    "Interest Income": 93,
    "Interest (Expense)": -637,
    "Goodwill Impairment": 0,
    "Other Non-Operating Income (Expense)": -597,
    "Pre-Tax Income (Loss)": 5994
  },
  "2Q26": {
    EBIT: 7286,
    "Interest Income": 94,
    "Interest (Expense)": -769,
    "Goodwill Impairment": 0,
    "Other Non-Operating Income (Expense)": 2708,
    "Pre-Tax Income (Loss)": 9319
  }
};

function cellValue(cell) {
  const value = cell.value;
  if (typeof value === "number" || typeof value === "string") return value;
  if (value && typeof value === "object") return value.result ?? value.text ?? value.formula ?? value.sharedFormula ?? null;
  return value;
}

function cellFormula(cell) {
  const value = cell.value;
  if (value && typeof value === "object" && typeof value.formula === "string") return value.formula;
  if (value && typeof value === "object" && typeof value.sharedFormula === "string") return value.sharedFormula;
  return null;
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
    .replace(/['\u2019]/g, "")
    .replace(/(?:E|EST|ESTIMATE)$/i, "");
  const direct = compact.match(/^([1-4])Q(\d{2}|\d{4})$/i);
  if (direct) return `${direct[1]}Q${direct[2].slice(-2)}`.toUpperCase();
  const qFirst = compact.match(/^Q([1-4])(\d{2}|\d{4})$/i);
  if (qFirst) return `${qFirst[1]}Q${qFirst[2].slice(-2)}`.toUpperCase();
  const fiscalYear = compact.match(/^(?:FY(\d{2}|\d{4})|(\d{4}))A?$/i);
  if (fiscalYear) return `FY${(fiscalYear[1] ?? fiscalYear[2]).slice(-2)}`.toUpperCase();
  return compact.toUpperCase();
}

function bestPeriodHeaderRow(sheet) {
  let best = null;
  for (let row = 1; row <= Math.min(sheet.rowCount, 120); row += 1) {
    const infos = [];
    for (let col = 1; col <= Math.min(sheet.columnCount, 180); col += 1) {
      const period = normalizePeriodLabel(cellValue(sheet.getCell(row, col)));
      if (!/^(?:[1-4]Q\d{2}|FY\d{2})$/.test(period)) continue;
      infos.push({ col, period });
    }
    const quarterCount = infos.filter((info) => /^[1-4]Q/.test(info.period)).length;
    const score = infos.length + quarterCount * 3;
    if (infos.length && (!best || score > best.score)) best = { row, infos, score };
  }
  return best;
}

function findPeriodColumn(sheet, period) {
  const header = bestPeriodHeaderRow(sheet);
  return header?.infos.find((info) => info.period === period.toUpperCase())?.col ?? null;
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

function valuesMatch(actual, expected, tolerance = 0.5) {
  return typeof actual === "number" && Number.isFinite(actual) && Math.abs(actual - expected) <= tolerance;
}

async function readWorkbook(file) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(file);
  return workbook;
}

async function fetchJson(url) {
  const response = await fetch(url, { headers: secHeaders });
  if (!response.ok) throw new Error(`Could not load ${url}: ${response.status} ${response.statusText}`);
  return response.json();
}

async function fetchCompanyFacts() {
  return fetchJson(`https://data.sec.gov/api/xbrl/companyfacts/CIK${cik}.json`);
}

function fiscalFocusFilingFromFacts(payload) {
  const candidates = [];
  for (const concept of ["Revenues", "RevenueFromContractWithCustomerExcludingAssessedTax"]) {
    const unitFacts = payload.facts?.["us-gaap"]?.[concept]?.units?.USD;
    if (!Array.isArray(unitFacts)) continue;
    for (const fact of unitFacts) {
      if (fact.form !== "10-Q") continue;
      if (Number(fact.fy) !== 2027 || fact.fp !== "Q1") continue;
      if (!fact.start || !fact.end || factDurationDays(fact) > 115) continue;
      candidates.push({
        accessionNumber: fact.accn,
        filingDate: fact.filed,
        reportDate: fact.end,
        fiscalYear: Number(fact.fy),
        fiscalPeriod: fact.fp
      });
    }
  }
  candidates.sort((a, b) => String(b.reportDate).localeCompare(String(a.reportDate)) || String(b.filingDate).localeCompare(String(a.filingDate)));
  const filing = candidates[0];
  if (!filing?.reportDate || !filing.accessionNumber) throw new Error("Could not find Walmart FY2027 Q1 10-Q facts in SEC companyfacts.");
  return filing;
}

function expectedConceptValue(payload, targetEndDate, concepts, kind) {
  for (const concept of concepts) {
    const fact = expectedFact(payload, targetEndDate, concept, kind);
    if (fact) return { concept, fact, value: fact.val / 1_000_000 };
  }
  return null;
}

function expectedFact(payload, targetEndDate, concept, kind) {
  const unitFacts = payload.facts?.["us-gaap"]?.[concept]?.units?.USD;
  if (!Array.isArray(unitFacts)) return null;
  const candidates = unitFacts.filter((fact) => {
    if (fact.form !== "10-Q") return false;
    if (fact.end !== targetEndDate) return false;
    if (Number(fact.fy) !== 2027 || fact.fp !== "Q1") return false;
    if (kind === "instant") return !fact.start;
    return Boolean(fact.start) && factDurationDays(fact) <= 115;
  });
  return candidates.sort(compareExpectedFactPreference)[0] ?? null;
}

function compareExpectedFactPreference(a, b) {
  const filedCompare = String(b.filed).localeCompare(String(a.filed));
  if (filedCompare !== 0) return filedCompare;
  return String(b.accn).localeCompare(String(a.accn));
}

function factDurationDays(fact) {
  return (new Date(`${fact.end}T00:00:00Z`).getTime() - new Date(`${fact.start}T00:00:00Z`).getTime()) / 86_400_000;
}

function filingMapRowsFor(sheet, period) {
  const rows = [];
  if (!sheet) return rows;
  for (let row = 2; row <= sheet.rowCount; row += 1) {
    if (String(cellValue(sheet.getCell(row, 1)) ?? "").toUpperCase() !== period.toUpperCase()) continue;
    rows.push({
      form: String(cellValue(sheet.getCell(row, 3)) ?? ""),
      periodEndDate: String(cellValue(sheet.getCell(row, 4)) ?? ""),
      fiscalYearLabel: String(cellValue(sheet.getCell(row, 5)) ?? ""),
      fiscalQuarterLabel: String(cellValue(sheet.getCell(row, 6)) ?? ""),
      accessionNumber: String(cellValue(sheet.getCell(row, 7)) ?? "")
    });
  }
  return rows;
}

function assignmentRowsFor(sheet, period, labelPattern) {
  const rows = [];
  if (!sheet) return rows;
  for (let row = 2; row <= sheet.rowCount; row += 1) {
    const fiscalPeriod = String(cellValue(sheet.getCell(row, 1)) ?? "");
    const label = String(cellValue(sheet.getCell(row, 4)) ?? "");
    if (fiscalPeriod !== period || !labelPattern.test(label)) continue;
    rows.push({
      sourceLineItemLabel: label,
      sourceAmount: cellValue(sheet.getCell(row, 5)),
      sourceXbrlTag: String(cellValue(sheet.getCell(row, 7)) ?? ""),
      assignedModelRow: String(cellValue(sheet.getCell(row, 8)) ?? ""),
      assignmentStatus: String(cellValue(sheet.getCell(row, 9)) ?? ""),
      classificationReason: String(cellValue(sheet.getCell(row, 10)) ?? ""),
      sourceSection: String(cellValue(sheet.getCell(row, 13)) ?? "")
    });
  }
  return rows;
}

async function main() {
  const facts = await fetchCompanyFacts();
  const filing = fiscalFocusFilingFromFacts(facts);
  await postWorkbook({ apiUrl, ticker, inputWorkbook, outputWorkbook });

  const workbook = await readWorkbook(outputWorkbook);
  const model = workbook.getWorksheet("Model");
  const filingMap = workbook.getWorksheet("Filing Period Map");
  const incomeStatementLedger = workbook.getWorksheet("Income Stmt Assignment Ledger");
  if (!model) throw new Error("Output workbook does not contain a Model sheet.");
  if (!filingMap) throw new Error("Output workbook does not contain a Filing Period Map sheet.");
  if (!incomeStatementLedger) throw new Error("Output workbook does not contain an Income Stmt Assignment Ledger sheet.");

  const errors = [];
  const col = findPeriodColumn(model, "1Q27");
  if (!col) errors.push("Could not find 1Q27 column in Model sheet.");

  const filingRows = filingMapRowsFor(filingMap, "1Q27");
  if (
    !filingRows.some(
      (row) =>
        row.form === "10-Q" &&
        row.periodEndDate === filing.reportDate &&
        row.fiscalYearLabel === "FY27" &&
        row.fiscalQuarterLabel === "Q1" &&
        row.accessionNumber === filing.accessionNumber
    )
  ) {
    errors.push(
      `Filing Period Map should tie 1Q27 to Walmart's FY2027 Q1 10-Q period ended ${filing.reportDate}, not a prior-year comparative period.`
    );
  }

  if (col) {
    for (const check of anchorRows) {
      const row = findRow(model, check.label);
      if (!row) {
        errors.push(`Could not find row "${check.label}" in Model sheet.`);
        continue;
      }
      const expected = expectedConceptValue(facts, filing.reportDate, check.concepts, check.kind);
      if (!expected) {
        errors.push(`Could not resolve SEC expected value for ${check.label} at ${filing.reportDate}.`);
        continue;
      }
      const actual = cellValue(model.getCell(row, col));
      if (!valuesMatch(actual, expected.value)) {
        errors.push(
          `1Q27 ${check.label}: expected ${expected.value} from EDGAR ${expected.concept} at ${filing.reportDate}, got ${actual ?? "[blank]"}.`
        );
      }
      const formula = cellFormula(model.getCell(row, col));
      if (check.label === "Revenue" && formula) {
        errors.push(`1Q27 Revenue should be a durable SEC actual after recalculation, but ${model.getCell(row, col).address} still contains formula "${formula}".`);
      }
    }
  }

  for (const [period, expectedRows] of Object.entries(expectedIncomeStatementBridge)) {
    const periodCol = findPeriodColumn(model, period);
    if (!periodCol) {
      errors.push(`Could not find ${period} column in Model sheet.`);
      continue;
    }

    const values = {};
    for (const [label, expected] of Object.entries(expectedRows)) {
      const row = findRow(model, label);
      if (!row) {
        errors.push(`Could not find income statement row "${label}" for ${period}.`);
        continue;
      }
      const actual = cellValue(model.getCell(row, periodCol));
      values[label] = actual;
      if (!valuesMatch(actual, expected)) {
        errors.push(`${period} ${label}: expected ${expected}, got ${actual ?? "[blank]"}.`);
      }
    }

    const bridge = ["EBIT", "Interest Income", "Interest (Expense)", "Goodwill Impairment", "Other Non-Operating Income (Expense)"]
      .map((label) => values[label])
      .reduce((total, value) => (typeof value === "number" ? total + value : NaN), 0);
    if (!valuesMatch(bridge, expectedRows["Pre-Tax Income (Loss)"])) {
      errors.push(`${period} pre-tax bridge should equal EBIT plus below-operating rows, got ${bridge}.`);
    }
  }

  for (const period of Object.keys(expectedIncomeStatementBridge)) {
    const membershipRows = assignmentRowsFor(incomeStatementLedger, period, /membership and other income/i);
    if (!membershipRows.length) {
      errors.push(`${period} Income Stmt Assignment Ledger should include WMT Membership and other income.`);
    }
    if (membershipRows.some((row) => /Other Non-Operating Income/i.test(row.assignedModelRow))) {
      errors.push(`${period} Membership and other income should be treated as revenue-section detail, not other non-operating: ${JSON.stringify(membershipRows)}`);
    }
    if (!membershipRows.some((row) => row.assignmentStatus === "subtotal_or_total_excluded" && row.sourceSection === "revenue")) {
      errors.push(`${period} Membership and other income should be excluded as revenue detail covered by total revenues: ${JSON.stringify(membershipRows)}`);
    }

    const otherGainLossRows = assignmentRowsFor(incomeStatementLedger, period, /other .*gains.*losses|other .*losses.*gains/i);
    if (!otherGainLossRows.some((row) => /Other Non-Operating Income/i.test(row.assignedModelRow))) {
      errors.push(`${period} Other gains/losses should map to Other Non-Operating Income (Expense): ${JSON.stringify(otherGainLossRows)}`);
    }
  }

  if (errors.length) {
    console.error(errors.join("\n"));
    throw new Error(`WMT fiscal-year regression failed with ${errors.length} issue(s).`);
  }

  console.log(`WMT fiscal-year regression passed: ${outputWorkbook}`);
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
