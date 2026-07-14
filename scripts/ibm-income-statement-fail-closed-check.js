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

const { __fillModelServiceTestHooks: hooks, validateHistoricalSourceLedger } = loadTypeScriptModule(sourcePath);

const accession = "000005114326000001";
const reportDate = "2026-03-31";

function statementRow(rowOrder, rowLabel, xbrlConcept, value) {
  return {
    statementName: "Consolidated Statement of Earnings",
    sourceTableType: "primary_statement",
    rowLabel,
    xbrlConcept,
    taxonomy: "us-gaap",
    value,
    unit: "USD",
    period: {
      start: "2026-01-01",
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

const primaryRows = [
  statementRow(1, "Total revenue", "Revenues", 15_917_000_000),
  statementRow(2, "Income from continuing operations before income taxes", "IncomeLossFromContinuingOperationsBeforeIncomeTaxes", 1_387_000_000),
  statementRow(3, "Net income", "NetIncomeLoss", 1_216_000_000)
];

const filingEntry = {
  accessionNumber: accession,
  accessionKey: accession,
  form: "10-Q",
  filingDate: "2026-04-30",
  reportDate,
  fiscalYear: 2026,
  fiscalQuarter: 1,
  quarterPeriod: "1Q26"
};

const ctx = {
  duration: new Map([
    [
      "1Q26",
      new Map(
        primaryRows.map((row) => [
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
      statementName: "Consolidated Statement of Earnings",
      sourceTableType: "primary_statement",
      accession,
      reportingPeriod: reportDate,
      form: "10-Q",
      filingDate: "2026-04-30",
      rows: primaryRows
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

assert.equal(
  hooks.otherNonOperatingValue({
    concept: "OtherIncomeAndExpense",
    label: "Other (income) and expense",
    value: -233_000_000,
    unit: "USD",
    sourceLayer: "sec_filing_package"
  }),
  233_000_000,
  "a broad OtherIncomeAndExpense filing concept uses the filing's expense-sign convention and must be inverted for the model"
);

const priorDiscontinuedSource = {
  concept: "IncomeLossFromDiscontinuedOperationsNetOfTax",
  label: "Income (loss) from discontinued operations, net of tax",
  value: -7_000_000,
  unit: "USD",
  taxonomy: "us-gaap",
  sourceLayer: "sec_filing_package",
  accn: "000005114323000001",
  periodKey: "1Q23",
  periodType: "quarterly",
  start: "2023-01-01",
  end: "2023-03-31"
};
const priorPeriodCtx = {
  ...ctx,
  duration: new Map([...ctx.duration, ["1Q23", new Map([[priorDiscontinuedSource.concept, priorDiscontinuedSource]])]])
};
const priorDiscontinued = hooks.resolveDiscontinuedOperationsBridge("1Q23", priorPeriodCtx);
assert.equal(priorDiscontinued.value, -7_000_000, "direct prior-year discontinued operations must not be replaced by a model-only zero");
assert.equal(priorDiscontinued.sources[0], priorDiscontinuedSource);

function ibm2023Row(rowOrder, rowLabel, xbrlConcept, value) {
  return {
    ...statementRow(rowOrder, rowLabel, xbrlConcept, value),
    accession: "000155837023006656",
    reportingPeriod: "2023-03-31",
    period: { start: "2023-01-01", end: "2023-03-31", periodType: "duration" }
  };
}

const ibm2023PrimaryRows = [
  ibm2023Row(1, "Revenue", "Revenues", 14_252_000_000),
  ibm2023Row(2, "Cost", "CostOfRevenue", 6_743_000_000),
  ibm2023Row(3, "Gross profit", "GrossProfit", 7_509_000_000),
  ibm2023Row(4, "SG&A expense", "SellingGeneralAndAdministrativeExpense", 4_853_000_000),
  // The SEC statement parser can omit R&D even though the same-period filing
  // fact is available. The full pre-tax equation must still anchor EBIT.
  ibm2023Row(5, "Intellectual property and custom development income", "IntellectualPropertyAndCustomDevelopmentIncome", 180_000_000),
  ibm2023Row(6, "Total other (income) and expense", "OtherIncomeAndExpense", -245_000_000),
  ibm2023Row(7, "Interest expense", "InterestExpense", 367_000_000),
  ibm2023Row(8, "Income from continuing operations before income taxes", "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest", 1_058_000_000),
  ibm2023Row(9, "Provision for income taxes", "IncomeTaxExpenseBenefit", 124_000_000),
  ibm2023Row(10, "Income from continuing operations", "IncomeLossFromContinuingOperations", 934_000_000),
  ibm2023Row(11, "Income from discontinued operations, net of tax", "IncomeLossFromDiscontinuedOperationsNetOfTaxAttributableToReportingEntity", -7_000_000),
  ibm2023Row(12, "Net income", "NetIncomeLoss", 927_000_000)
];
const ibm2023FactRows = [
  ...ibm2023PrimaryRows,
  ibm2023Row(5, "Research and development", "ResearchAndDevelopmentExpense", 1_655_000_000)
];
const ibm2023Facts = new Map(
  ibm2023FactRows.map((row) => [
    row.xbrlConcept,
    {
      concept: row.xbrlConcept,
      label: row.rowLabel,
      value: row.value,
      unit: "USD",
      taxonomy: "us-gaap",
      sourceLayer: "sec_filing_package",
      accn: row.accession,
      start: row.period.start,
      end: row.period.end,
      periodKey: "1Q23",
      periodType: "quarterly",
      reportDate: row.reportingPeriod
    }
  ])
);
const ibm2023Entry = {
  accessionNumber: "000155837023006656",
  accessionKey: "000155837023006656",
  form: "10-Q",
  filingDate: "2023-05-10",
  reportDate: "2023-03-31",
  fiscalYear: 2023,
  fiscalQuarter: 1,
  quarterPeriod: "1Q23"
};
const ibm2023Ctx = {
  ...ctx,
  duration: new Map([["1Q23", ibm2023Facts]]),
  filingPackageStatements: [
    {
      statementName: "Statement Consolidated Income Statement",
      sourceTableType: "primary_statement",
      accession: ibm2023Entry.accessionNumber,
      reportingPeriod: ibm2023Entry.reportDate,
      form: ibm2023Entry.form,
      filingDate: ibm2023Entry.filingDate,
      rows: ibm2023PrimaryRows
    }
  ],
  fiscalPeriods: {
    ...ctx.fiscalPeriods,
    entries: [ibm2023Entry],
    byAccession: new Map([[ibm2023Entry.accessionKey, ibm2023Entry]]),
    byReportDate: new Map([[ibm2023Entry.reportDate, ibm2023Entry]]),
    reportedPeriods: new Set(["1Q23"])
  }
};
const operatingResolverStartedAt = Date.now();
const ibm2023OperatingIncome = hooks.resolveOperatingIncome("1Q23", ibm2023Ctx);
assert.ok(Date.now() - operatingResolverStartedAt < 2_000, "the missing-subtotal operating-income equation must resolve without runaway recursion");
assert.equal(ibm2023OperatingIncome.value, 1_181_000_000, "IBM-shaped operating components must reconcile through the SEC pre-tax equation");
assert.ok(
  ibm2023OperatingIncome.sources[0].derivationCalculation?.terms.length,
  "the source-backed operating-income equation must retain replayable calculation provenance"
);
assert.ok(
  !ibm2023OperatingIncome.sources[0].derivationCalculation.terms.some((term) => /BeforeIncomeTaxes/.test(term.concept)),
  "operating-income amount provenance must replay from operating components while pre-tax remains a reconciliation check"
);
assert.equal(hooks.resolveOtherNonOperatingIncomeExpense("1Q23", ibm2023Ctx).value, 245_000_000);

function scaledIbmPeriod(period, start, end, accession, multiplier, periodType, form) {
  const rows = ibm2023PrimaryRows.map((row) => ({
    ...row,
    value: row.value * multiplier,
    accession,
    reportingPeriod: end,
    period: { start, end, periodType: "duration" }
  }));
  const rdRow = {
    ...ibm2023FactRows.find((row) => row.xbrlConcept === "ResearchAndDevelopmentExpense"),
    value: 1_655_000_000 * multiplier,
    accession,
    reportingPeriod: end,
    period: { start, end, periodType: "duration" }
  };
  const facts = new Map(
    [...rows, rdRow].map((row) => [
      row.xbrlConcept,
      {
        concept: row.xbrlConcept,
        label: row.rowLabel,
        value: row.value,
        unit: "USD",
        taxonomy: "us-gaap",
        sourceLayer: "sec_filing_package",
        accn: accession,
        start,
        end,
        periodKey: period,
        periodType,
        reportDate: end
      }
    ])
  );
  return {
    period,
    rows,
    facts,
    statement: {
      statementName: "Statement Consolidated Income Statement",
      sourceTableType: "primary_statement",
      accession,
      reportingPeriod: end,
      form,
      filingDate: end,
      rows
    },
    entry: {
      accessionNumber: accession,
      accessionKey: accession,
      form,
      filingDate: end,
      reportDate: end,
      fiscalYear: 2023,
      fiscalQuarter: period.startsWith("FY") ? 4 : Number(period[0]),
      quarterPeriod: period.startsWith("FY") ? "4Q23" : period,
      annualPeriod: "FY23"
    }
  };
}

const ibmMultiPeriods = [
  scaledIbmPeriod("1Q23", "2023-01-01", "2023-03-31", "000000000023000001", 1, "quarterly", "10-Q"),
  scaledIbmPeriod("2Q23", "2023-04-01", "2023-06-30", "000000000023000002", 1, "quarterly", "10-Q"),
  scaledIbmPeriod("3Q23", "2023-07-01", "2023-09-30", "000000000023000003", 1, "quarterly", "10-Q"),
  scaledIbmPeriod("FY23", "2023-01-01", "2023-12-31", "000000000023000004", 4, "annual", "10-K")
];
const ibmMultiPeriodCtx = {
  ...ctx,
  duration: new Map(ibmMultiPeriods.map((item) => [item.period, item.facts])),
  filingPackageStatements: ibmMultiPeriods.map((item) => item.statement),
  fiscalPeriods: {
    ...ctx.fiscalPeriods,
    entries: ibmMultiPeriods.map((item) => item.entry),
    byAccession: new Map(ibmMultiPeriods.map((item) => [item.entry.accessionKey, item.entry])),
    byReportDate: new Map(ibmMultiPeriods.map((item) => [item.entry.reportDate, item.entry])),
    reportedPeriods: new Set(["1Q23", "2Q23", "3Q23", "4Q23"])
  }
};
const fourthQuarterResolverStartedAt = Date.now();
const ibm2023FourthQuarterOperatingIncome = hooks.resolveOperatingIncome("4Q23", ibmMultiPeriodCtx);
assert.ok(Date.now() - fourthQuarterResolverStartedAt < 2_000, "annual-to-fourth-quarter operating bridges must use bounded resolver work");
assert.ok(
  Math.abs((ibm2023FourthQuarterOperatingIncome.value ?? 0) - 1_181_000_000) <= 5_000_000,
  "the synthetic fourth-quarter bridge must stay within accumulated whole-million SEC rounding"
);

const aggregateAnnualFacts = new Map(ibmMultiPeriods.find((item) => item.period === "FY23").facts);
aggregateAnnualFacts.delete("OtherIncomeAndExpense");
aggregateAnnualFacts.set("OtherNonoperatingIncomeExpense", {
  concept: "OtherNonoperatingIncomeExpense",
  label: "Other Nonoperating Income (Expense)",
  value: 266_000_000,
  unit: "USD",
  taxonomy: "us-gaap",
  sourceLayer: "sec_live_companyfacts",
  accn: "000000000023000004",
  start: "2023-01-01",
  end: "2023-12-31",
  periodKey: "FY23",
  periodType: "annual",
  reportDate: "2023-12-31"
});
aggregateAnnualFacts.set("InterestExpense", {
  ...aggregateAnnualFacts.get("InterestExpense"),
  value: 1_607_000_000
});
const aggregatePreTaxConcept = "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest";
aggregateAnnualFacts.set(aggregatePreTaxConcept, {
  ...aggregateAnnualFacts.get(aggregatePreTaxConcept),
  value: 8_690_000_000
});
aggregateAnnualFacts.set("OperatingIncomeLoss", {
  ...aggregateAnnualFacts.get(aggregatePreTaxConcept),
  concept: "OperatingIncomeLoss",
  label: "Operating income",
  value: 8_424_000_000
});
const aggregateAnnualCtx = {
  ...ibmMultiPeriodCtx,
  duration: new Map([...ibmMultiPeriodCtx.duration, ["FY23", aggregateAnnualFacts]]),
  filingPackageStatements: ibmMultiPeriodCtx.filingPackageStatements.map((statement) =>
    statement.reportingPeriod === "2023-12-31"
      ? {
          ...statement,
          rows: [
            ...statement.rows
              .filter((row) => !["OtherIncomeAndExpense", "InterestExpense"].includes(row.xbrlConcept))
              .map((row) => row.xbrlConcept === aggregatePreTaxConcept ? { ...row, value: 8_690_000_000 } : row),
            {
              ...statement.rows.find((row) => row.xbrlConcept === aggregatePreTaxConcept),
              rowOrder: 7,
              rowLabel: "Operating income",
              xbrlConcept: "OperatingIncomeLoss",
              value: 8_424_000_000
            }
          ]
        }
      : statement
  )
};
assert.equal(
  hooks.resolveOperatingIncome("FY23", aggregateAnnualCtx).value,
  8_424_000_000,
  "the reported annual operating-income subtotal must remain authoritative in the aggregate split"
);
assert.equal(
  hooks.resolveInterestExpense("FY23", aggregateAnnualCtx).value,
  -1_607_000_000,
  "a same-period SEC interest fact must be usable when the reported nonoperating aggregate bridges operating income to pre-tax income"
);
assert.equal(
  hooks.resolveOtherNonOperatingIncomeExpense("FY23", aggregateAnnualCtx).value,
  1_873_000_000,
  "a non-primary nonoperating aggregate must exclude separately modeled annual interest expense"
);

const derivedOtherOperatingValidation = hooks.validateResolvedValueForWrite(
  { cik: "0000051143", ticker: "IBM", title: "International Business Machines Corporation" },
  {
    row: 35,
    label: "Other Operating Income (Expense)",
    classification: "grouped",
    statement: "income",
    kind: "duration",
    concepts: []
  },
  "4Q23",
  {
    value: -716_000_000,
    classification: "grouped",
    sources: [
      {
        concept: "OtherOperatingIncomeExpenseFromPreTaxEquation",
        label: "Other operating income/expense derived from the complete SEC pre-tax equation",
        value: -716_000_000,
        sourceLayer: "derived",
        periodKey: "4Q23",
        periodType: "quarterly"
      },
      {
        concept: "Revenues",
        label: "Fourth-quarter revenue bridge",
        value: 17_381_000_000,
        sourceLayer: "derived",
        periodKey: "4Q23",
        periodType: "quarterly"
      },
      {
        concept: "Revenues",
        label: "Annual revenue",
        value: 61_860_000_000,
        sourceLayer: "sec_live_companyfacts",
        periodKey: "FY23",
        periodType: "annual"
      }
    ]
  }
);
assert.notEqual(
  derivedOtherOperatingValidation.status,
  "blocked",
  "a compatible final derived output must treat nested component bridges as traced inputs rather than competing row classifications"
);

const revenueFillRowWithBadStatementInference = {
  row: 28,
  label: "Revenue",
  classification: "direct",
  statement: "support",
  kind: "duration",
  scale: 1_000_000,
  modelContext: {
    sheetName: "Model",
    row: 28,
    label: "Revenue",
    sectionHeader: "Income Statement",
    previousLabel: "Income Statement",
    nextLabel: "Cost of Goods Sold",
    indentation: 0,
    hasHistoricalFormula: false,
    hasHardcodedInput: true,
    hasNetRevenueInterestExpenseAbove: false,
    projectedColumns: 0,
    signConvention: 1
  }
};

const candidateRows = hooks.incomeStatementAssignmentCandidateRows([revenueFillRowWithBadStatementInference]);
assert.equal(candidateRows.length, 1);
assert.equal(candidateRows[0].label, "Revenue");

const ledger = hooks.buildPrimaryIncomeStatementAssignmentLedgerRows(["1Q26"], ctx, [revenueFillRowWithBadStatementInference]);
const revenueAssignment = ledger.find((row) => row.sourceLineItemLabel === "Total revenue");
assert.ok(revenueAssignment);
assert.equal(revenueAssignment.assignedModelRow, "Revenue");
assert.equal(revenueAssignment.modelAmount, 15_917_000_000);

const workbook = new ExcelJS.Workbook();
const sheet = workbook.addWorksheet("Model");
sheet.getCell("C10").value = "Income Statement";
sheet.getCell("C11").value = "EBIT";
sheet.getCell("U11").value = 1666;
const evaluator = new hooks.FormulaEvaluator(sheet);
const warnings = [];
const errors = hooks.validateIncomeStatementMetricAgainstEdgar(
  sheet,
  ["1Q26"],
  [21],
  ctx,
  evaluator,
  warnings,
  "EBIT",
  ["EBIT", "Operating Income"],
  ["OperatingIncomeLoss"],
  () => ({
    value: 1_859_000_000,
    sources: [
      {
        concept: "OperatingIncomeLoss",
        label: "Operating income",
        value: 1_859_000_000,
        unit: "USD",
        sourceLayer: "sec_filing_package",
        periodKey: "1Q26",
        periodType: "quarterly"
      }
    ],
    classification: "direct"
  }),
  { hard: true }
);

assert.equal(warnings.length, 0);
assert.equal(errors.length, 1);
assert.match(errors[0], /does not match EDGAR 1859/);
assert.match(errors[0], /must be refreshed or remapped before output/);

sheet.getCell("U12").value = 100;
sheet.getCell("U11").value = { formula: "U12+5", result: 1859 };
const poisonedCacheErrors = hooks.validateIncomeStatementMetricAgainstEdgar(
  sheet,
  ["1Q26"],
  [21],
  ctx,
  new hooks.FormulaEvaluator(sheet, { useCachedFormulaResults: true }),
  [],
  "EBIT",
  ["EBIT", "Operating Income"],
  ["OperatingIncomeLoss"],
  () => ({
    value: 1_859_000_000,
    sources: [{ concept: "OperatingIncomeLoss", label: "Operating income", value: 1_859_000_000, unit: "USD", sourceLayer: "sec_filing_package" }],
    classification: "direct"
  }),
  { hard: true }
);
assert.equal(poisonedCacheErrors.length, 1);
assert.match(poisonedCacheErrors[0], /EBIT 105 does not match EDGAR 1859/);

sheet.getCell("U13").value = { formula: "UNSUPPORTED(1)", result: 1859 };
assert.equal(
  new hooks.FormulaEvaluator(sheet).evaluateCell(sheet.getCell("U13")),
  null,
  "cached formula results must be opt-in rather than the evaluator default"
);
assert.equal(
  hooks.statementMetricCellValue(sheet.getCell("U13"), new hooks.FormulaEvaluator(sheet), 1859),
  null,
  "statement validation must not accept a poisoned cache when the formula cannot be evaluated"
);
assert.equal(
  hooks.incomeStatementClassificationCellTiesResolvedValue(1859, null, 1859),
  false,
  "classification reconciliation must not treat a matching cache as proof of an unsupported formula"
);
sheet.getCell("T14").value = { formula: 'UNSUPPORTED("Y")', result: "Y" };
sheet.getCell("U14").value = { formula: 'IF(T14="Y",1859,0)', result: 1859 };
assert.equal(
  new hooks.FormulaEvaluator(sheet).evaluateCell(sheet.getCell("U14")),
  null,
  "a poisoned text-formula cache must not choose a validating IF branch"
);

const revenueWorkbook = new ExcelJS.Workbook();
const revenueSheet = revenueWorkbook.addWorksheet("Model");
revenueSheet.getCell("C10").value = "Income Statement";
revenueSheet.getCell("C11").value = "Revenue";
revenueSheet.getCell("U12").value = 100;
revenueSheet.getCell("U11").value = { formula: "U12+5", result: 15_917 };
hooks.markReportedPeriodColumns(revenueSheet, [{ period: "1Q26", col: 21 }]);
const revenueAuditRows = [];
const revenueReconciliation = hooks.reconcileIncomeStatementFormulaMetricToEdgar(
  revenueSheet,
  ["1Q26"],
  [21],
  ctx,
  revenueAuditRows,
  ["Revenue"],
  () => ({
    value: 15_917_000_000,
    sources: [{ concept: "Revenues", label: "Total revenue", value: 15_917_000_000, unit: "USD", sourceLayer: "sec_filing_package" }],
    classification: "direct"
  }),
  "revenue"
);
assert.equal(revenueReconciliation.filledCells, 1);
assert.equal(revenueSheet.getCell("U11").value, 15_917);
assert.equal(revenueAuditRows[0].formulaPreserved, false);
assert.match(revenueAuditRows[0].formulaStatus, /formula replaced/i);

revenueSheet.getCell("U12").value = 15_912;
revenueSheet.getCell("U11").value = { formula: "U12+5", result: 15_917 };
const tiedRevenueAuditRows = [];
const tiedRevenueReconciliation = hooks.reconcileIncomeStatementFormulaMetricToEdgar(
  revenueSheet,
  ["1Q26"],
  [21],
  ctx,
  tiedRevenueAuditRows,
  ["Revenue"],
  () => ({
    value: 15_917_000_000,
    sources: [{ concept: "Revenues", label: "Total revenue", value: 15_917_000_000, unit: "USD", sourceLayer: "sec_filing_package" }],
    classification: "direct"
  }),
  "revenue"
);
assert.equal(tiedRevenueReconciliation.filledCells, 0, "an already-current formula cache does not require a workbook write");
assert.equal(tiedRevenueAuditRows.length, 1, "an already-correct reported formula still needs current SEC provenance in the source ledger");
assert.equal(tiedRevenueAuditRows[0].formulaPreserved, true);

const balanceWorkbook = new ExcelJS.Workbook();
const balanceSheet = balanceWorkbook.addWorksheet("Model");
balanceSheet.getCell("C10").value = "Balance Sheet";
balanceSheet.getCell("C11").value = "Total Assets";
balanceSheet.getCell("U11").value = 100;
balanceSheet.getCell("V11").value = { formula: "U11+5", result: 100 };
hooks.markReportedPeriodColumns(balanceSheet, [
  { period: "4Q26", col: 21 },
  { period: "FY26", col: 22 }
]);
const balanceAuditRows = [];
const annualCopy = hooks.copyBalanceSheetFourthQuarterToAnnualColumns(
  balanceSheet,
  ["4Q26", "FY26"],
  [21, 22],
  { duration: new Map(), instant: new Map() },
  balanceAuditRows
);
assert.equal(annualCopy.filledCells, 1);
assert.equal(balanceSheet.getCell("V11").value, 100);
assert.equal(balanceAuditRows[0].formulaPreserved, false);
assert.match(balanceAuditRows[0].formulaStatus, /formula replaced/i);

const exactAnnualBalanceWorkbook = new ExcelJS.Workbook();
const exactAnnualBalanceSheet = exactAnnualBalanceWorkbook.addWorksheet("Model");
exactAnnualBalanceSheet.getCell("C10").value = "Balance Sheet";
exactAnnualBalanceSheet.getCell("C11").value = "Total Assets";
exactAnnualBalanceSheet.getCell("U11").value = 100;
exactAnnualBalanceSheet.getCell("V11").value = 100;
hooks.markReportedPeriodColumns(exactAnnualBalanceSheet, [
  { period: "4Q26", col: 21 },
  { period: "FY26", col: 22 }
]);
const exactAnnualSource = {
  concept: "Assets",
  label: "Total assets",
  value: 100_000_000,
  unit: "USD",
  sourceLayer: "sec_filing_package",
  accn: "000000000026000001",
  end: "2026-12-31",
  periodKey: "4Q26",
  periodType: "instant"
};
const exactAnnualAuditRows = [];
const exactAnnualCopy = hooks.copyBalanceSheetFourthQuarterToAnnualColumns(
  exactAnnualBalanceSheet,
  ["4Q26", "FY26"],
  [21, 22],
  { duration: new Map(), instant: new Map([["4Q26", new Map([["Assets", exactAnnualSource]])]]) },
  exactAnnualAuditRows
);
assert.equal(exactAnnualCopy.filledCells, 0, "an already-correct annual balance-sheet hardcode does not require a workbook write");
assert.equal(exactAnnualAuditRows.length, 1, "an already-correct annual balance-sheet hardcode still needs refreshed SEC provenance");
assert.equal(exactAnnualAuditRows[0].sourceProvenance[0].role, "sec_source");
assert.equal(exactAnnualAuditRows[0].sourceProvenance[0].value, 100, "annual SEC audit provenance must use the workbook's millions scale");

const annualFormulaWorkbook = new ExcelJS.Workbook();
const annualFormulaSheet = annualFormulaWorkbook.addWorksheet("Model");
annualFormulaSheet.getCell("C10").value = "Balance Sheet";
annualFormulaSheet.getCell("C11").value = "Total Assets";
annualFormulaSheet.getCell("U11").value = 100;
annualFormulaSheet.getCell("V11").value = { formula: "U11", result: 95 };
hooks.markReportedPeriodColumns(annualFormulaSheet, [
  { period: "4Q26", col: 21 },
  { period: "FY26", col: 22 }
]);
const annualFormulaAuditRows = [];
const annualFormulaCopy = hooks.copyBalanceSheetFourthQuarterToAnnualColumns(
  annualFormulaSheet,
  ["4Q26", "FY26"],
  [21, 22],
  { duration: new Map(), instant: new Map() },
  annualFormulaAuditRows
);
assert.equal(annualFormulaCopy.filledCells, 1);
assert.deepEqual(annualFormulaSheet.getCell("V11").value, { formula: "U11", result: 100 });
assert.equal(annualFormulaAuditRows.length, 1);
assert.equal(annualFormulaAuditRows[0].mappingType, "formula preserved");
assert.equal(annualFormulaAuditRows[0].formulaPreserved, true);
assert.equal(annualFormulaAuditRows[0].conceptsUsed, "");
assert.equal(annualFormulaAuditRows[0].accession, "");
assert.deepEqual(
  annualFormulaAuditRows[0].sourceProvenance,
  [],
  "a preserved annual balance-sheet formula must not manufacture a bare derived-output source"
);

const annualAccession = "000000000026000001";
const q4AuditRow = hooks.statementTotalAuditRow(
  annualFormulaSheet,
  annualFormulaSheet.getCell("U11"),
  "Total Assets",
  "4Q26",
  100,
  {
    concept: "Assets",
    label: "Total assets",
    value: 100_000_000,
    unit: "USD",
    sourceLayer: "sec_filing_package",
    accn: annualAccession,
    form: "10-K",
    filed: "2027-02-15",
    end: "2026-12-31",
    periodKey: "4Q26",
    periodType: "instant"
  },
  "balance",
  "Total assets maps to the current SEC balance sheet."
);
const annualLedgerPeriodEntries = [
  {
    period: "4Q26",
    column: 21,
    modelColumn: "U",
    accessionKey: annualAccession,
    accessionNumber: annualAccession,
    form: "10-K",
    filingDate: "2027-02-15",
    periodEndDate: "2026-12-31",
    fiscalYearLabel: "FY26",
    fiscalQuarterLabel: "Q4"
  },
  {
    period: "FY26",
    column: 22,
    modelColumn: "V",
    accessionKey: annualAccession,
    accessionNumber: annualAccession,
    form: "10-K",
    filingDate: "2027-02-15",
    periodEndDate: "2026-12-31",
    fiscalYearLabel: "FY26",
    fiscalQuarterLabel: "FY"
  }
];
const annualLedgerRows = hooks.buildHistoricalSourceLedgerRows(
  { cik: "0000000000", ticker: "EXM", title: "Example Corp." },
  annualLedgerPeriodEntries,
  annualFormulaSheet,
  [{ row: 11, label: "Total Assets", classification: "formula", statement: "balance", kind: "instant" }],
  [
    { period: "4Q26", col: 21 },
    { period: "FY26", col: 22 }
  ],
  [q4AuditRow, ...annualFormulaAuditRows]
);
const annualLedgerFormulaRow = annualLedgerRows.find((row) => row.cell === "V11");
assert.ok(annualLedgerFormulaRow);
assert.equal(annualLedgerFormulaRow.mappingStatus, "formula_preserved");
assert.equal(annualLedgerFormulaRow.workbookFormulaPrecedents, "Model!U11");
assert.deepEqual(
  annualLedgerFormulaRow.sourceProvenance,
  [],
  "Source Ledger must recurse to the Q4 precedent instead of validating a copied synthetic amount"
);

assert.deepEqual(
  hooks.llmMappingReviewFailureBlockingErrors("LLM mapping review attempted_failed (Key limit exceeded)."),
  [],
  "LLM provider outages are advisory; deterministic SEC validation is the release gate"
);

validateHistoricalSourceLedger(
  annualLedgerRows,
  annualLedgerPeriodEntries,
  { cik: "0000000000", ticker: "EXM", title: "Example Corp." },
  new Map([[annualAccession, {}]])
)
  .then((annualLedgerErrors) => {
    assert.deepEqual(
      annualLedgerErrors,
      [],
      "an FY balance-sheet formula must close recursively through its current-company Q4 SEC-backed precedent"
    );
    console.log("IBM income-statement fail-closed regression passed.");
  })
  .catch((error) => {
    console.error(error);
    process.exitCode = 1;
  });
