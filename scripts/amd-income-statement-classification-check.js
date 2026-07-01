const assert = require("node:assert/strict");
const ExcelJS = require("exceljs");
const fs = require("node:fs");
const path = require("node:path");
const Module = require("node:module");
const ts = require("typescript");

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

function statementRow(rowOrder, rowLabel, xbrlConcept, value) {
  return {
    statementName: "Condensed Consolidated Statements of Operations",
    sourceTableType: "primary_statement",
    rowLabel,
    xbrlConcept,
    taxonomy: xbrlConcept.startsWith("AmortizationOfAcquisition") ? "amd" : "us-gaap",
    value,
    unit: "USD",
    period: {
      start: "2025-12-29",
      end: reportDate,
      periodType: "duration"
    },
    consolidated: true,
    dimensions: [],
    rowOrder,
    accession,
    reportingPeriod: reportDate
  };
}

const rows = [
  statementRow(1, "Net revenue", "RevenueFromContractWithCustomerExcludingAssessedTax", 10_253_000_000),
  statementRow(2, "Cost of sales", "CostOfGoodsAndServiceExcludingDepreciationDepletionAndAmortization", 4_576_000_000),
  statementRow(3, "Amortization of acquisition-related intangibles", "AmortizationOfAcquisitionRelatedIntangiblesCOGS", 261_000_000),
  statementRow(4, "Total cost of sales", "CostOfGoodsAndServicesSold", 4_837_000_000),
  statementRow(5, "Gross profit", "GrossProfit", 5_416_000_000),
  statementRow(6, "Research and development", "ResearchAndDevelopmentExpense", 2_397_000_000),
  statementRow(7, "Marketing, general and administrative", "SellingGeneralAndAdministrativeExpense", 1_253_000_000),
  statementRow(8, "Amortization of acquisition-related intangibles", "AmortizationOfAcquisitionRelatedIntangiblesOpex", 290_000_000),
  statementRow(9, "Total operating expenses", "OperatingExpenses", 3_940_000_000),
  statementRow(10, "Operating income", "OperatingIncomeLoss", 1_476_000_000),
  statementRow(11, "Interest expense", "InterestExpense", 37_000_000),
  statementRow(12, "Other income (expense), net", "OtherNonoperatingIncomeExpense", 165_000_000),
  statementRow(
    13,
    "Income from continuing operations before income taxes and equity income",
    "IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments",
    1_604_000_000
  ),
  statementRow(14, "Income tax provision", "IncomeTaxExpenseBenefit", 238_000_000),
  statementRow(15, "Equity income in investee", "IncomeLossFromEquityMethodInvestments", 6_000_000),
  statementRow(16, "Income from continuing operations, net of tax", "IncomeLossFromContinuingOperations", 1_372_000_000),
  statementRow(
    17,
    "Income from discontinued operations, net of tax",
    "IncomeLossFromDiscontinuedOperationsNetOfTaxAttributableToReportingEntity",
    11_000_000
  ),
  statementRow(18, "Net income", "NetIncomeLoss", 1_383_000_000)
];

const ctx = {
  duration: new Map([
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
            start: row.period.start,
            end: reportDate,
            periodKey: "1Q26",
            periodType: "quarterly",
            reportDate
          }
        ])
      )
    ]
  ]),
  instant: new Map(),
  filingPackageStatements: [
    {
      statementName: "Condensed Consolidated Statements of Operations",
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
  "Revenue",
  "Cost of Goods Sold",
  "Selling, General & Administration (SG&A)",
  "Research & Development (R&D)",
  "Depreciation & Amortization",
  "Other Operating Income (Expense)",
  "Interest (Expense)",
  "Other Non-Operating Income (Expense)",
  "Pre-Tax Income (Loss)",
  "Income Tax Benefit (Expense)",
  "Net Income (Loss)",
  "Post-Tax Adjustments",
  "Discontinued Operations"
].map((label, index) => ({
  row: index + 1,
  label,
  classification: "direct",
  statement: "income",
  kind: "duration",
  scale: 1_000_000
}));

const ledger = hooks.buildPrimaryIncomeStatementAssignmentLedgerRows(["1Q26"], ctx, fillRows);
const byLabel = new Map(ledger.map((row) => [row.sourceLineItemLabel, row]));
const byLabelRows = (label) => ledger.filter((row) => row.sourceLineItemLabel === label);

assert.equal(byLabel.get("Cost of sales").assignedModelRow, "Cost of Goods Sold");
assert.equal(byLabel.get("Cost of sales").modelAmount, -4_576_000_000);
assert.equal(byLabel.get("Total cost of sales").assignmentStatus, "subtotal_or_total_excluded");

const amortizationRows = byLabelRows("Amortization of acquisition-related intangibles");
assert.equal(amortizationRows.length, 2);
assert.equal(
  amortizationRows.find((row) => row.sourceXbrlTag === "AmortizationOfAcquisitionRelatedIntangiblesCOGS").assignedModelRow,
  "Cost of Goods Sold"
);
assert.equal(
  amortizationRows.find((row) => row.sourceXbrlTag === "AmortizationOfAcquisitionRelatedIntangiblesCOGS").modelAmount,
  -261_000_000
);
assert.equal(
  amortizationRows.find((row) => row.sourceXbrlTag === "AmortizationOfAcquisitionRelatedIntangiblesOpex").assignedModelRow,
  "Depreciation & Amortization"
);
assert.equal(
  amortizationRows.find((row) => row.sourceXbrlTag === "AmortizationOfAcquisitionRelatedIntangiblesOpex").modelAmount,
  -290_000_000
);

assert.equal(byLabel.get("Other income (expense), net").assignedModelRow, "Other Non-Operating Income (Expense)");
assert.equal(byLabel.get("Other income (expense), net").modelAmount, 165_000_000);
assert.equal(byLabel.get("Equity income in investee").assignedModelRow, "Post-Tax Adjustments");
assert.equal(byLabel.get("Equity income in investee").modelAmount, 6_000_000);
assert.equal(byLabel.get("Income from discontinued operations, net of tax").assignedModelRow, "Discontinued Operations");
assert.equal(byLabel.get("Income from discontinued operations, net of tax").modelAmount, 11_000_000);

const primarySources = hooks.primaryIncomeStatementSourcesForPeriod("1Q26", ctx);
assert.equal(primarySources.some((source) => source.label === "Cost of sales"), true);

const primaryCost = hooks.resolvePrimaryStatementCostOfRevenue("1Q26", ctx);
assert.equal(primaryCost.value, -4_576_000_000);
assert.equal(hooks.primaryStatementHasCostOfRevenueSplitWithSeparateDa("1Q26", ctx), true);

const cogs = hooks.resolveCostOfRevenue("1Q26", ctx);
assert.equal(cogs.value, -4_837_000_000);
assert.match(cogs.note, /consolidated SEC cost of revenue/i);

const da = hooks.resolveIncomeStatementDepreciationAmortization("1Q26", ctx);
assert.equal(da.value, -290_000_000);
assert.match(da.note, /primary income statement/i);

const otherOperating = hooks.resolveOtherOperatingIncomeExpense("1Q26", ctx);
assert.equal(otherOperating.value, 0);

const workbook = new ExcelJS.Workbook();
const sheet = workbook.addWorksheet("Model");
sheet.getCell("A40").value = "x";
sheet.getCell("C40").value = "Income Statement";
sheet.getCell("C42").value = "Pre-Tax Income (Loss)";
sheet.getCell("C44").value = "Income Tax Benefit (Expense)";
sheet.getCell("C45").value = "Net Income (Loss)";
sheet.getCell("C48").value = "Post-Tax Adjustments";
sheet.getCell("C49").value = "Discontinued Operations";
sheet.getCell("C50").value = "Adj. Net Income (Loss)";
assert.equal(hooks.isPostTaxEquityMethodBridgeFormulaUpdate(sheet.getCell("Q45"), "Q42+Q44+Q48+Q49"), true);
assert.equal(hooks.isPostTaxEquityMethodBridgeFormulaUpdate(sheet.getCell("T45"), "T42+T44+T49"), true);
assert.equal(hooks.isPostTaxEquityMethodBridgeFormulaUpdate(sheet.getCell("Q50"), "SUM(Q45:Q47)"), true);
assert.equal(hooks.isPostTaxEquityMethodBridgeFormulaUpdate(sheet.getCell("T50"), "SUM(T45:T48)"), true);

console.log("AMD income statement classification regression passed.");
