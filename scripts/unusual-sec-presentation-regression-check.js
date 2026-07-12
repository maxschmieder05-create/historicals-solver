const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const Module = require("node:module");
const ts = require("typescript");
const ExcelJS = require("exceljs");

function compileTypeScript(source) {
  return ts.transpileModule(source, {
    compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020, esModuleInterop: true }
  }).outputText;
}

function loadTypeScriptModule(file) {
  if (!require.extensions[".ts"]) {
    require.extensions[".ts"] = (mod, filename) => mod._compile(compileTypeScript(fs.readFileSync(filename, "utf8")), filename);
  }
  const mod = new Module(file, module);
  mod.filename = file;
  mod.paths = Module._nodeModulePaths(path.dirname(file));
  mod._compile(compileTypeScript(fs.readFileSync(file, "utf8")), file);
  return mod.exports;
}

const sourcePath = path.join(__dirname, "..", "server", "fill-model", "fill-model-service.ts");
const hooks = loadTypeScriptModule(sourcePath).__fillModelServiceTestHooks;

function statementStructure(statementName, sourceTableType = "primary_statement") {
  return {
    statementName,
    sourceTableType,
    accession: "0000000000-26-000001",
    rows: []
  };
}

assert.equal(
  hooks.isPrimaryIncomeStatementStructure(
    statementStructure("CONSOLIDATED STATEMENTS OF INCOME AND OTHER COMPREHENSIVE INCOME")
  ),
  true,
  "a combined income and OCI primary statement must be eligible as the primary income statement"
);
assert.equal(
  hooks.isPrimaryIncomeStatementStructure(
    statementStructure("CONSOLIDATED STATEMENTS OF OPERATIONS AND COMPREHENSIVE INCOME")
  ),
  true,
  "a combined operations and comprehensive-income primary statement must be eligible"
);
assert.equal(
  hooks.isPrimaryIncomeStatementStructure(
    statementStructure("CONSOLIDATED STATEMENTS OF INCOME AND OTHER COMPREHENSIVE LOSS")
  ),
  true,
  "a combined income and comprehensive-loss primary statement must be eligible"
);
assert.equal(
  hooks.isPrimaryIncomeStatementStructure(statementStructure("CONSOLIDATED STATEMENTS OF OTHER COMPREHENSIVE INCOME")),
  false,
  "a pure OCI statement must not be selected as the primary income statement"
);
assert.equal(
  hooks.isPrimaryIncomeStatementStructure(statementStructure("CONSOLIDATED STATEMENTS OF COMPREHENSIVE INCOME")),
  false,
  "a comprehensive-income-only title must not pass without an income/operations component"
);
assert.equal(
  hooks.isPrimaryIncomeStatementStructure(statementStructure("CONSOLIDATED STATEMENTS OF OTHER COMPREHENSIVE LOSS")),
  false,
  "a pure OCI-loss statement must not be selected as the primary income statement"
);
assert.equal(
  hooks.isPrimaryIncomeStatementStructure(statementStructure("CONSOLIDATED STATEMENTS OF COMPREHENSIVE LOSS")),
  false,
  "a comprehensive-loss-only title must not pass without an income/operations component"
);

function statementRow(rowOrder, rowLabel, xbrlConcept, value) {
  return {
    statementName: "Consolidated Balance Sheet",
    sourceTableType: "primary_statement",
    rowLabel,
    xbrlConcept,
    taxonomy: "us-gaap",
    value,
    unit: "USD",
    period: { instant: "2026-03-31", periodType: "instant" },
    consolidated: true,
    dimensions: [],
    currentNonCurrentSection: "current",
    rowOrder,
    accession: "0000000000-26-000001",
    reportingPeriod: "2026-03-31"
  };
}

const aggregateSource = { concept: "ExpenseAndOtherIncome", label: "Expense and other income", value: 900, unit: "USD", sourceLayer: "sec_filing_package" };
const aggregateRow = {
  statementName: "Consolidated Statement of Earnings",
  sourceTableType: "primary_statement",
  rowLabel: aggregateSource.label,
  xbrlConcept: aggregateSource.concept,
  value: aggregateSource.value,
  unit: "USD",
  period: { start: "2026-01-01", end: "2026-03-31", periodType: "duration" },
  consolidated: true,
  dimensions: [],
  rowOrder: 8,
  accession: "0000000000-26-000001"
};
assert.equal(hooks.classifyIncomeStatementPresentationRole(aggregateSource, aggregateRow, "operating expenses"), "subtotal");
assert.equal(hooks.incomeStatementSourceIsSubtotalOrTotalToExclude(aggregateSource, "operating expenses"), true);

const workbook = new ExcelJS.Workbook();
const segmentSheet = workbook.addWorksheet("Segment Analysis");
segmentSheet.getCell("C8").value = "Software Revenue";
segmentSheet.getCell("F8").value = { formula: "0", result: 0 };
segmentSheet.getCell("C9").value = "Consulting Revenue";
segmentSheet.getCell("F9").value = { formula: "0", result: 0 };
assert.equal(hooks.findSegmentResidualRow(segmentSheet, [8, 9], 6, "Revenue"), null);
const segment = {
  label: "Infrastructure",
  family: "reportable",
  values: new Map([["1Q26", 100]]),
  operatingIncome: new Map(),
  depreciationAmortization: new Map()
};
assert.match(hooks.unmatchedSegmentCoverageWarnings(segmentSheet, [8, 9], [segment], "Revenue", "values", ["1Q26"])[0], /non-blocking disclosure-coverage warning/);
segmentSheet.getCell("C9").value = "Other / Reconciliation Revenue";
assert.equal(hooks.findSegmentResidualRow(segmentSheet, [8, 9], 6, "Revenue"), 9);

const labelCoverageWorkbook = new ExcelJS.Workbook();
const labelCoverageSheet = labelCoverageWorkbook.addWorksheet("Segment Analysis");
labelCoverageSheet.getCell("C7").value = "Total Company Revenue";
labelCoverageSheet.getCell("C8").value = "Cloud Revenue";
labelCoverageSheet.getCell("C9").value = "Services Revenue";
labelCoverageSheet.getCell("C14").value = "Revenue Mix";
for (let row = 8; row <= 13; row += 1) labelCoverageSheet.getCell(row, 6).value = 0;
const disclosedSegment = (label) => ({
  label,
  family: "reportable",
  values: new Map([["1Q26", 100]]),
  operatingIncome: new Map(),
  depreciationAmortization: new Map()
});
assert.equal(
  hooks.shouldPreserveExistingSegmentLabels(
    labelCoverageSheet,
    ["1Q26"],
    [6],
    [disclosedSegment("Cloud"), disclosedSegment("Services")]
  ),
  true
);
assert.equal(
  hooks.shouldPreserveExistingSegmentLabels(
    labelCoverageSheet,
    ["1Q26"],
    [6],
    [disclosedSegment("Cloud"), disclosedSegment("Services"), disclosedSegment("Security")]
  ),
  false,
  "label preservation requires 75% coverage of disclosed labels, not merely the smaller label set"
);
labelCoverageSheet.getCell("C10").value = "Security Revenue";
assert.equal(
  hooks.shouldPreserveExistingSegmentLabels(
    labelCoverageSheet,
    ["1Q26"],
    [6],
    [disclosedSegment("Cloud"), disclosedSegment("Services")]
  ),
  false,
  "label preservation also requires 75% coverage of existing non-placeholder labels"
);
const partialCoverageSegments = ["Cloud", "Services", "Security", "Data"].map((label) => ({
  ...disclosedSegment(label),
  values: new Map([["1Q26", 100_000_000]])
}));
assert.equal(
  hooks.shouldPreserveExistingSegmentLabels(labelCoverageSheet, ["1Q26"], [6], partialCoverageSegments),
  true,
  "three matched labels out of four disclosures meet the two-sided 75% preservation threshold"
);
const partialCoverageAuditRows = [];
const partialCoverageFill = hooks.fillSegmentMetricRows(
  labelCoverageSheet,
  ["1Q26"],
  [6],
  [8, 9, 10, 11, 12, 13],
  partialCoverageSegments,
  "values",
  "Revenue",
  partialCoverageAuditRows,
  { forceOrderedAssignment: false, clearUnmatchedLabels: true, matchExistingLabelsOnly: false }
);
assert.equal(partialCoverageFill.filledCells, 4);
assert.equal(labelCoverageSheet.getCell("C8").text, "Cloud Revenue");
assert.equal(labelCoverageSheet.getCell("C9").text, "Services Revenue");
assert.equal(labelCoverageSheet.getCell("C10").text, "Security Revenue");
assert.equal(labelCoverageSheet.getCell("C11").text, "Data Revenue");
assert.equal(labelCoverageSheet.getCell("F11").value, 100);
assert.equal(partialCoverageAuditRows.length, 4, "the formerly unmatched disclosure must receive full source-audit coverage");
labelCoverageSheet.getCell("C20").value = "Cloud Operating Income";
labelCoverageSheet.getCell("C21").value = "Services Operating Income";
labelCoverageSheet.getCell("C22").value = "Security Operating Income";
for (let row = 20; row <= 25; row += 1) labelCoverageSheet.getCell(row, 6).value = 0;
const partialOperatingSegments = partialCoverageSegments.map((segment) => ({
  ...segment,
  operatingIncome: new Map([["1Q26", 25_000_000]])
}));
const partialOperatingAuditRows = [];
const partialOperatingFill = hooks.fillSegmentMetricRows(
  labelCoverageSheet,
  ["1Q26"],
  [6],
  [20, 21, 22, 23, 24, 25],
  partialOperatingSegments,
  "operatingIncome",
  "Operating Income",
  partialOperatingAuditRows,
  { clearUnmatchedLabels: true, matchExistingLabelsOnly: false }
);
assert.equal(partialOperatingFill.filledCells, 4);
assert.equal(labelCoverageSheet.getCell("C23").text, "Data Operating Income");
assert.equal(labelCoverageSheet.getCell("F23").value, 25);
assert.equal(partialOperatingAuditRows.length, 4);

const structureWorkbook = new ExcelJS.Workbook();
const model = structureWorkbook.addWorksheet("Model");
model.getCell("C2").value = "Revenue";
model.getCell("F2").value = 10;
model.getCell("G2").value = { formula: "F2", result: 10 };
model.getColumn(3).width = 24;
model.getRow(2).height = 18;
model.mergeCells("A1:B1");
model.getCell("F2").dataValidation = { type: "decimal", operator: "greaterThanOrEqual", formulae: [0] };
model.getCell("F3").style = { numFmt: "0.0x" };
structureWorkbook.definedNames.add("Model!$F$2", "HistoricalInput");
const snapshot = hooks.snapshotWorkbook(structureWorkbook, ["Model"], 6);
assert.deepEqual(hooks.validateWorkbookPreservation(structureWorkbook, snapshot), []);
const returnedParserSnapshot = hooks.workbookSnapshotForReturnedValidation(
  { ...snapshot, definedNamesFingerprint: "pre-sanitization ExcelJS parser view" },
  structureWorkbook
);
assert.deepEqual(
  hooks.validateWorkbookPreservation(structureWorkbook, returnedParserSnapshot),
  [],
  "post-serialization validation relies on the exact OOXML defined-name invariant and must not compare a lossy pre-sanitization ExcelJS parser view"
);
model.getCell("F3").value = 42;
model.getCell("H3").value = 7;
assert.deepEqual(
  hooks.validateWorkbookPreservation(structureWorkbook, snapshot),
  [],
  "ordinary writes into blank historical cells must not be mistaken for structural mutations"
);
model.getColumn(3).width = 30;
assert.ok(hooks.validateWorkbookPreservation(structureWorkbook, snapshot).some((error) => /protected template structure/.test(error)));
model.getColumn(3).width = 24;
model.getCell("G2").value = { formula: "F2+1", result: 11 };
assert.ok(hooks.validateWorkbookPreservation(structureWorkbook, snapshot).some((error) => /formula changed/.test(error)));

const segmentStructureWorkbook = new ExcelJS.Workbook();
const segmentStructureSheet = segmentStructureWorkbook.addWorksheet("Segment Analysis");
segmentStructureSheet.getCell("C7").value = "Total Company Revenue";
segmentStructureSheet.getCell("C8").value = "Segment 1 Revenue";
segmentStructureSheet.getCell("C14").value = "Revenue Mix";
let segmentStructureSnapshot = hooks.snapshotWorkbook(segmentStructureWorkbook, ["Segment Analysis"], 6);
hooks.setSegmentMetricRowLabel(segmentStructureSheet, 8, "Revenue", "Infrastructure");
assert.deepEqual(
  hooks.validateWorkbookPreservation(segmentStructureWorkbook, segmentStructureSnapshot),
  [],
  "the generated segment-label rewrite must be a permitted preservation mutation"
);
segmentStructureSheet.getCell("C8").value = "Unregistered Label Mutation";
assert.ok(
  hooks.validateWorkbookPreservation(segmentStructureWorkbook, segmentStructureSnapshot).some((error) => /row label changed/.test(error)),
  "only the exact generated segment label is permitted; a subsequent unrelated mutation remains blocked"
);
segmentStructureSheet.getCell("C8").value = "Infrastructure Revenue";
segmentStructureSnapshot = hooks.snapshotWorkbook(segmentStructureWorkbook, ["Segment Analysis"], 6);
hooks.clearSegmentMetricRowLabel(segmentStructureSheet, 8);
assert.deepEqual(
  hooks.validateWorkbookPreservation(segmentStructureWorkbook, segmentStructureSnapshot),
  [],
  "the generated clearing of an unused segment label must be a permitted preservation mutation"
);
segmentStructureSheet.getCell("C7").value = "Mutated Total Label";
assert.ok(
  hooks.validateWorkbookPreservation(segmentStructureWorkbook, segmentStructureSnapshot).some((error) => /row label changed/.test(error)),
  "segment-label exceptions must not permit mutation of a neighboring total label"
);

const intentionalFormulaWorkbook = new ExcelJS.Workbook();
const intentionalSegmentSheet = intentionalFormulaWorkbook.addWorksheet("Segment Analysis");
intentionalSegmentSheet.getCell("C7").value = "Total Company Revenue";
intentionalSegmentSheet.getCell("C8").value = "Cloud Revenue";
intentionalSegmentSheet.getCell("C14").value = "Revenue Mix";
intentionalSegmentSheet.getCell("F7").value = { formula: "SUM(F8:F13)", result: 100 };
intentionalSegmentSheet.getCell("F8").value = 100;
let intentionalFormulaSnapshot = hooks.snapshotWorkbook(intentionalFormulaWorkbook, ["Segment Analysis"], 6);
intentionalSegmentSheet.getCell("F7").value = { formula: "SUM(F8:F11)", result: 100 };
assert.deepEqual(
  hooks.validateWorkbookPreservation(intentionalFormulaWorkbook, intentionalFormulaSnapshot),
  [],
  "an intentional Segment Analysis total-formula rewrite must use the existing narrow preservation allowance"
);

const postTaxWorkbook = new ExcelJS.Workbook();
const postTaxSheet = postTaxWorkbook.addWorksheet("Model");
postTaxSheet.getCell("C5").value = "Income Statement";
postTaxSheet.getCell("C6").value = "Pre-Tax Income (Loss)";
postTaxSheet.getCell("C7").value = "Income Tax Benefit (Expense)";
postTaxSheet.getCell("C8").value = "Net Income (Loss)";
postTaxSheet.getCell("C9").value = "Post-Tax Adjustments";
postTaxSheet.getCell("C10").value = "Adj. Net Income (Loss)";
postTaxSheet.getCell("F8").value = { formula: "F6+F7", result: 90 };
intentionalFormulaSnapshot = hooks.snapshotWorkbook(postTaxWorkbook, ["Model"], 6);
postTaxSheet.getCell("F8").value = { formula: "F6+F7+F9", result: 95 };
assert.deepEqual(
  hooks.validateWorkbookPreservation(postTaxWorkbook, intentionalFormulaSnapshot),
  [],
  "an intentional post-tax bridge formula update must use the existing narrowly validated bridge allowance"
);

const reportedBalanceSheetWorkbook = new ExcelJS.Workbook();
const reportedBalanceSheet = reportedBalanceSheetWorkbook.addWorksheet("Model");
reportedBalanceSheet.getCell("C5").value = "Balance Sheet";
reportedBalanceSheet.getCell("C6").value = "Cash & Cash Equivalents";
reportedBalanceSheet.getCell("C7").value = "Working Capital";
reportedBalanceSheet.getCell("F6").value = { formula: "0", result: 0 };
hooks.markReportedPeriodColumns(reportedBalanceSheet, [{ period: "1Q26", col: 6 }]);
intentionalFormulaSnapshot = hooks.snapshotWorkbook(reportedBalanceSheetWorkbook, ["Model"], 6);
reportedBalanceSheet.getCell("F6").value = 125;
assert.deepEqual(
  hooks.validateWorkbookPreservation(reportedBalanceSheetWorkbook, intentionalFormulaSnapshot),
  [],
  "an EDGAR-backed reported Balance Sheet formula replacement must use its existing narrow preservation allowance"
);

const reportedWeightedSharesWorkbook = new ExcelJS.Workbook();
const reportedWeightedSharesSheet = reportedWeightedSharesWorkbook.addWorksheet("Model");
reportedWeightedSharesSheet.getCell("C5").value = "Shares Outstanding Schedule";
reportedWeightedSharesSheet.getCell("C6").value = "Weighted Average Basic Shares";
reportedWeightedSharesSheet.getCell("F6").value = { formula: "E6", result: 100 };
hooks.markReportedPeriodColumns(reportedWeightedSharesSheet, [{ period: "FY26", col: 6 }]);
intentionalFormulaSnapshot = hooks.snapshotWorkbook(reportedWeightedSharesWorkbook, ["Model"], 6);
reportedWeightedSharesSheet.getCell("F6").value = 123;
assert.ok(
  hooks
    .validateWorkbookPreservation(reportedWeightedSharesWorkbook, intentionalFormulaSnapshot)
    .some((error) => /formula changed/.test(error)),
  "a weighted-average share formula replacement without a current SEC audit row must remain blocked"
);
assert.deepEqual(
  hooks.validateWorkbookPreservation(reportedWeightedSharesWorkbook, intentionalFormulaSnapshot, [
    {
      sheetName: "Model",
      cell: "F6",
      modelRowLabel: "Weighted Average Basic Shares",
      period: "FY26",
      valueWritten: 123,
      mappingType: "direct",
      conceptsUsed: "WeightedAverageNumberOfSharesOutstandingBasic=123mm",
      sourceStatement: "support",
      accession: "0000000000-26-000001",
      sourceUrl: "https://www.sec.gov/Archives/edgar/data/0/example.htm",
      cellWritable: true,
      formulaPreserved: false,
      formulaStatus: "reported-period formula replaced with direct annual SEC fact",
      writeBlockedReason: "",
      signConvention: "reported",
      confidence: "high",
      validationStatus: "OK!",
      notes: ""
    }
  ]),
  [],
  "a reported weighted-average share formula may be replaced only when its exact cell has current SEC provenance"
);

reportedWeightedSharesSheet.getCell("C7").value = "Weighted Average Dilutive Shares";
reportedWeightedSharesSheet.getCell("F7").value = { formula: "E7", result: 101 };
intentionalFormulaSnapshot = hooks.snapshotWorkbook(reportedWeightedSharesWorkbook, ["Model"], 6);
reportedWeightedSharesSheet.getCell("F7").value = 124;
assert.deepEqual(
  hooks.validateWorkbookPreservation(reportedWeightedSharesWorkbook, intentionalFormulaSnapshot, [
    {
      sheetName: "Model",
      cell: "F7",
      modelRowLabel: "Weighted Average Dilutive Shares",
      period: "FY26",
      valueWritten: 124,
      mappingType: "direct",
      conceptsUsed: "WeightedAverageNumberOfDilutedSharesOutstanding=124mm",
      sourceStatement: "support",
      accession: "0000000000-26-000001",
      sourceUrl: "https://www.sec.gov/Archives/edgar/data/0/example.htm",
      cellWritable: true,
      formulaPreserved: false,
      formulaStatus: "reported-period formula replaced with direct annual SEC fact",
      writeBlockedReason: "",
      signConvention: "reported",
      confidence: "high",
      validationStatus: "OK!",
      notes: ""
    }
  ]),
  [],
  "diluted weighted-average shares use the same audited non-additive replacement rule"
);

const receivableRows = [
  statementRow(1, "Accounts receivable, net", "AccountsReceivableNetCurrent", 100_000_000),
  statementRow(2, "Receivables classified as held for sale", "FinancingReceivablesHeldForSaleCurrent", 25_000_000)
];
const filingEntry = {
  accessionNumber: "0000000000-26-000001",
  accessionKey: "000000000026000001",
  form: "10-Q",
  filingDate: "2026-04-30",
  reportDate: "2026-03-31",
  fiscalYear: 2026,
  fiscalQuarter: 1,
  quarterPeriod: "1Q26"
};
const ctx = {
  duration: new Map(),
  instant: new Map(),
  filingPackageStatements: [{
    statementName: "Consolidated Balance Sheet",
    sourceTableType: "primary_statement",
    accession: filingEntry.accessionNumber,
    reportingPeriod: filingEntry.reportDate,
    form: filingEntry.form,
    filingDate: filingEntry.filingDate,
    rows: receivableRows
  }],
  fiscalPeriods: {
    entries: [filingEntry],
    byAccession: new Map([[filingEntry.accessionNumber, filingEntry]]),
    byReportDate: new Map([[filingEntry.reportDate, filingEntry]]),
    reportedPeriods: new Set(["1Q26"]),
    fiscalYearEndMonth: 12,
    fiscalYearEndDay: 31
  }
};
const receivables = hooks.resolveAccountsReceivable("1Q26", ctx);
assert.equal(receivables.value, 125_000_000);
assert.equal(receivables.sources.length, 2);
assert.equal(new Set(receivables.sources.map((source) => `${source.concept}|${source.value}`)).size, 2);
assert.equal(hooks.sourceLooksLikeReceivableHeldForSale(receivables.sources[1]), true);

const duplicateRows = [{ sourceRowKey: "same" }, { sourceRowKey: "same" }];
assert.deepEqual(hooks.duplicateBalanceSheetAssignmentKeys(duplicateRows), ["same"]);
assert.deepEqual(hooks.duplicateIncomeStatementAssignmentKeys(duplicateRows), ["same"]);
assert.equal(hooks.validationRepairStrategy([], model, "same", null), "targeted");
assert.equal(hooks.shouldStopRepeatedUnrepairableValidationFailure("same", "same", 1), true);
assert.equal(hooks.shouldStopRepeatedUnrepairableValidationFailure("same", "different", 0), true);

assert.equal(
  hooks.llmApiKeyForEndpoint("https://openrouter.ai/api/v1/chat/completions", {
    OPENAI_API_KEY: "openai-secret",
    OPENROUTER_API_KEY: ""
  }),
  "",
  "an OpenAI credential must never be sent to OpenRouter"
);
assert.equal(
  hooks.llmApiKeyForEndpoint("https://api.openai.com/v1/chat/completions", {
    OPENAI_API_KEY: "openai-secret",
    OPENROUTER_API_KEY: "openrouter-secret"
  }),
  "openai-secret"
);
assert.equal(
  hooks.llmApiKeyForEndpoint("https://llm.internal.example/v1/chat/completions", {
    OPENAI_API_KEY: "openai-secret",
    OPENROUTER_API_KEY: "openrouter-secret",
    ACCOUNTING_LLM_API_KEY: "internal-secret"
  }),
  "internal-secret"
);

console.log("Unusual SEC presentation, structural immutability, coverage-warning, receivable aggregation, duplicate-source, and retry regressions passed.");
