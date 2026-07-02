const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const Module = require("node:module");
const ts = require("typescript");
const ExcelJS = require("exceljs");

const repoRoot = path.resolve(__dirname, "..");
const sourcePath = path.join(repoRoot, "server", "fill-model", "fill-model-service.ts");

function compileTypeScript(source) {
  return ts.transpileModule(source, {
    compilerOptions: {
      module: ts.ModuleKind.CommonJS,
      target: ts.ScriptTarget.ES2020,
      esModuleInterop: true
    }
  }).outputText;
}

function registerTypeScriptRequire() {
  if (require.extensions[".ts"]) return;
  require.extensions[".ts"] = (mod, file) => {
    mod._compile(compileTypeScript(fs.readFileSync(file, "utf8")), file);
  };
}

function loadTypeScriptModule(file) {
  registerTypeScriptRequire();
  const source = fs.readFileSync(file, "utf8");
  const compiled = compileTypeScript(source);
  const mod = new Module(file, module);
  mod.filename = file;
  mod.paths = Module._nodeModulePaths(path.dirname(file));
  mod._compile(compiled, file);
  return mod.exports;
}

const { __fillModelServiceTestHooks: hooks } = loadTypeScriptModule(sourcePath);

const accession = "000000248826000001";
const reportDate = "2026-03-28";
const filingEntry = {
  accessionNumber: accession,
  accessionKey: accession,
  form: "10-Q",
  filingDate: "2026-05-06",
  reportDate,
  fiscalYear: 2026,
  fiscalQuarter: 1,
  quarterPeriod: "1Q26"
};

function balanceRow(rowOrder, rowLabel, xbrlConcept, value, section = "current") {
  return {
    statementName: "Condensed Consolidated Balance Sheets",
    sourceTableType: "primary_statement",
    rowLabel,
    xbrlConcept,
    taxonomy: "us-gaap",
    value,
    unit: "USD",
    period: {
      instant: reportDate,
      periodType: "instant"
    },
    consolidated: true,
    dimensions: [],
    rowOrder,
    accession,
    reportingPeriod: reportDate,
    currentNonCurrentSection: section,
    parentSubtotal: { label: "Current liabilities", concept: "LiabilitiesCurrent" }
  };
}

const rows = [
  balanceRow(1, "Accounts payable", "AccountsPayableCurrent", 2_997_000_000),
  balanceRow(2, "Accrued liabilities", "AccruedLiabilitiesCurrentAndNoncurrent", 5_785_000_000),
  balanceRow(3, "Current portion of long-term debt, net", "LongTermDebtCurrent", 874_000_000),
  balanceRow(4, "Other current liabilities", "OtherLiabilitiesCurrent", 850_000_000),
  balanceRow(5, "Total current liabilities", "LiabilitiesCurrent", 10_506_000_000)
];

const ctx = {
  duration: new Map(),
  instant: new Map([
    [
      "1Q26",
      new Map(
        rows.map((row) => [
          row.xbrlConcept,
          {
            concept: row.xbrlConcept,
            label: row.rowLabel,
            value: row.value,
            unit: "USD",
            taxonomy: row.taxonomy,
            sourceLayer: "sec_filing_package",
            accn: accession,
            end: reportDate,
            periodKey: "1Q26",
            periodType: "instant",
            reportDate
          }
        ])
      )
    ]
  ]),
  filingPackageStatements: [
    {
      statementName: "Condensed Consolidated Balance Sheets",
      sourceTableType: "primary_statement",
      accession,
      reportingPeriod: reportDate,
      form: "10-Q",
      filingDate: "2026-05-06",
      rows
    }
  ],
  fiscalPeriods: {
    entries: [filingEntry],
    byAccession: new Map([[accession, filingEntry]]),
    byReportDate: new Map([[reportDate, filingEntry]]),
    reportedPeriods: new Set(["1Q26"]),
    fiscalYearEndMonth: 12,
    fiscalYearEndDay: 31
  }
};

const fillRows = [
  "Accounts Payable",
  "Accrued Liabilities",
  "LT Debt (Incl. Current Portion)",
  "Other Current Liabilities"
].map((label, index) => ({
  row: index + 1,
  label,
  classification: "direct",
  statement: "balance",
  kind: "instant",
  scale: 1_000_000
}));

const ledger = hooks.buildPrimaryBalanceSheetAssignmentLedgerRows(["1Q26"], ctx, fillRows);
const byLabel = new Map(ledger.map((row) => [row.sourceLineItemLabel, row]));

assert.equal(byLabel.get("Accounts payable").assignedModelRow, "Accounts Payable");
assert.equal(byLabel.get("Accounts payable").amount, 2_997_000_000);
assert.equal(byLabel.get("Accrued liabilities").assignedModelRow, "Accrued Liabilities");
assert.equal(byLabel.get("Accrued liabilities").amount, 5_785_000_000);
assert.equal(byLabel.get("Current portion of long-term debt, net").assignedModelRow, "LT Debt (Incl. Current Portion)");
assert.equal(byLabel.get("Current portion of long-term debt, net").amount, 874_000_000);
assert.equal(byLabel.get("Other current liabilities").assignedModelRow, "Other Current Liabilities");
assert.equal(byLabel.get("Other current liabilities").amount, 850_000_000);
assert.equal(byLabel.has("Total current liabilities"), false);

const accrued = hooks.resolveAccruedLiabilities("1Q26", ctx);
assert.equal(accrued.value, 5_785_000_000);
assert.equal(accrued.sources.some((source) => source.concept === "LongTermDebtCurrent"), false);

const otherCurrent = hooks.resolveOtherCurrentLiabilities("1Q26", ctx);
assert.equal(otherCurrent.value, 850_000_000);
assert.equal(otherCurrent.sources.some((source) => source.concept === "AccruedLiabilitiesCurrentAndNoncurrent"), false);

const tslaAccession = "000162828026000001";
const tslaReportDate = "2026-03-31";
const tslaFilingEntry = {
  accessionNumber: tslaAccession,
  accessionKey: tslaAccession,
  form: "10-Q",
  filingDate: "2026-04-23",
  reportDate: tslaReportDate,
  fiscalYear: 2026,
  fiscalQuarter: 1,
  quarterPeriod: "1Q26"
};

function tslaBalanceRow(rowOrder, rowLabel, xbrlConcept, value, section, parentLabel, parentConcept) {
  return {
    statementName: "Condensed Consolidated Balance Sheets",
    sourceTableType: "primary_statement",
    rowLabel,
    xbrlConcept,
    taxonomy: "us-gaap",
    value,
    unit: "USD",
    period: {
      instant: tslaReportDate,
      periodType: "instant"
    },
    consolidated: true,
    dimensions: [],
    rowOrder,
    accession: tslaAccession,
    reportingPeriod: tslaReportDate,
    currentNonCurrentSection: section,
    parentSubtotal: { label: parentLabel, concept: parentConcept }
  };
}

const tslaRows = [
  tslaBalanceRow(1, "Cash and cash equivalents", "CashAndCashEquivalentsAtCarryingValue", 16_603_000_000, "current", "Current assets", "AssetsCurrent"),
  tslaBalanceRow(2, "Accounts receivable, net", "AccountsReceivableNetCurrent", 3_959_000_000, "current", "Current assets", "AssetsCurrent"),
  tslaBalanceRow(3, "Total assets", "Assets", 143_724_000_000, "total", "Assets", "Assets"),
  tslaBalanceRow(4, "Accounts payable", "AccountsPayableCurrent", 14_696_000_000, "current", "Current liabilities", "LiabilitiesCurrent"),
  tslaBalanceRow(5, "Accrued liabilities and other", "AccruedAndOtherCurrentLiabilities", 14_554_000_000, "current", "Current liabilities", "LiabilitiesCurrent"),
  tslaBalanceRow(6, "Current portion of debt and finance leases", "DebtCurrent", 1_374_000_000, "current", "Current liabilities", "LiabilitiesCurrent"),
  tslaBalanceRow(7, "Operating lease liability, current", "OperatingLeaseLiabilityCurrent", 988_000_000, "current", "Current liabilities", "LiabilitiesCurrent"),
  tslaBalanceRow(8, "Total current liabilities", "LiabilitiesCurrent", 34_138_000_000, "current", "Current liabilities", "LiabilitiesCurrent"),
  tslaBalanceRow(9, "Digital assets", "CryptoAssetFairValueNoncurrent", 786_000_000, "non_current", "Assets", "Assets"),
  tslaBalanceRow(10, "Deferred revenue, net of current portion", "ContractWithCustomerLiabilityNoncurrent", 3_847_000_000, "non_current", "Liabilities", "Liabilities"),
  tslaBalanceRow(11, "Other long-term liabilities", "OtherLiabilitiesNoncurrent", 13_155_000_000, "non_current", "Liabilities", "Liabilities"),
  tslaBalanceRow(12, "Total liabilities", "Liabilities", 58_922_000_000, "total", "Liabilities", "Liabilities"),
  tslaBalanceRow(13, "Redeemable noncontrolling interests in subsidiaries", "RedeemableNoncontrollingInterestEquityCarryingAmount", 57_000_000, "equity", "Equity", "StockholdersEquity"),
  tslaBalanceRow(14, "Noncontrolling interests in subsidiaries", "MinorityInterest", 629_000_000, "equity", "Equity", "StockholdersEquity")
];

const tslaInstantFacts = new Map(
  tslaRows.map((row) => [
    row.xbrlConcept,
    {
      concept: row.xbrlConcept,
      label: row.rowLabel,
      value: row.value,
      unit: "USD",
      taxonomy: row.taxonomy,
      sourceLayer: "sec_filing_package",
      accn: tslaAccession,
      end: tslaReportDate,
      periodKey: "1Q26",
      periodType: "instant",
      reportDate: tslaReportDate
    }
  ])
);

const templateWithoutMezzanine = {
  hasCashRow: true,
  hasCashAndCurrentInvestmentRow: false,
  hasCurrentInvestmentRow: true,
  hasCurrentDebtRow: true,
  hasCurrentDebtMaturitiesRow: true,
  hasShortTermBorrowingsRow: false,
  hasCurrentLiabilitiesExcludingDebtRow: true,
  hasOtherCurrentLiabilityRow: true,
  hasDebtInclCurrentPortionRow: true,
  hasDeferredTaxLiabilityRow: true,
  hasNonCurrentLeaseLiabilityRow: false,
  hasPensionLiabilityRow: false,
  hasMezzanineEquityRow: false
};

const tslaCtx = {
  duration: new Map(),
  instant: new Map([
    ["FY25", new Map([["Goodwill", { concept: "Goodwill", label: "Goodwill", value: 257_000_000, sourceLayer: "sec_filing_package", periodKey: "FY25", periodType: "instant" }]])],
    ["1Q26", tslaInstantFacts]
  ]),
  filingPackageStatements: [
    {
      statementName: "Condensed Consolidated Balance Sheets",
      sourceTableType: "primary_statement",
      accession: tslaAccession,
      reportingPeriod: tslaReportDate,
      form: "10-Q",
      filingDate: "2026-04-23",
      rows: tslaRows
    }
  ],
  fiscalPeriods: {
    entries: [tslaFilingEntry],
    byAccession: new Map([[tslaAccession, tslaFilingEntry]]),
    byReportDate: new Map([[tslaReportDate, tslaFilingEntry]]),
    reportedPeriods: new Set(["1Q26"]),
    fiscalYearEndMonth: 12,
    fiscalYearEndDay: 31
  },
  template: templateWithoutMezzanine
};

const tslaFillRows = [
  "Accounts Payable",
  "Accrued Liabilities",
  "Other Current Liabilities",
  "Other Non-Current Assets",
  "Other Non-Current Liabilities",
  "Noncontrolling Interests"
].map((label, index) => ({
  row: index + 20,
  label,
  classification: "direct",
  statement: "balance",
  kind: "instant",
  scale: 1_000_000
}));

const tslaLedger = hooks.buildPrimaryBalanceSheetAssignmentLedgerRows(["1Q26"], tslaCtx, tslaFillRows);
const tslaByLabel = new Map(tslaLedger.map((row) => [row.sourceLineItemLabel, row]));

assert.equal(hooks.resolveAccruedLiabilities("1Q26", tslaCtx).value, 14_554_000_000);
assert.equal(hooks.resolveGoodwill("1Q26", tslaCtx).value, 0);

const tslaOtherNonCurrentLiabilities = hooks.resolveOtherNonCurrentLiabilities("1Q26", tslaCtx);
assert.equal(tslaOtherNonCurrentLiabilities.value, 17_059_000_000);
assert.equal(tslaOtherNonCurrentLiabilities.sources.some((source) => source.concept === "CryptoAssetFairValueNoncurrent"), false);
assert.equal(tslaOtherNonCurrentLiabilities.sources.some((source) => source.concept === "RedeemableNoncontrollingInterestEquityCarryingAmount"), true);
assert.equal(hooks.resolveTotalLiabilities("1Q26", tslaCtx).value, 58_979_000_000);
assert.equal(tslaByLabel.get("Digital assets").assignedModelRow, "Other Non-Current Assets");
assert.equal(tslaByLabel.get("Redeemable noncontrolling interests in subsidiaries").assignedModelRow, "Other Non-Current Liabilities");
assert.notEqual(tslaByLabel.get("Redeemable noncontrolling interests in subsidiaries").assignmentStatus, "explicitly_excluded_with_reason");

const workbook = new ExcelJS.Workbook();
const formulaSheet = workbook.addWorksheet("Model");
formulaSheet.getCell("A1").value = 100;
formulaSheet.getCell("A2").value = 43;
formulaSheet.getCell("A3").value = { formula: "A1-A2", result: 0 };
const formulaEvaluator = new hooks.FormulaEvaluator(formulaSheet, { useCachedFormulaResults: true });
assert.equal(hooks.statementMetricCellValue(formulaSheet.getCell("A3"), formulaEvaluator, 0), 57);

console.log("AMD balance sheet liability classification regression passed.");
