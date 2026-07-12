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
