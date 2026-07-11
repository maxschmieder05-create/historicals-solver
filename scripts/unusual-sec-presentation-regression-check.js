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

const structureWorkbook = new ExcelJS.Workbook();
const model = structureWorkbook.addWorksheet("Model");
model.getCell("C2").value = "Revenue";
model.getCell("F2").value = 10;
model.getCell("G2").value = { formula: "F2", result: 10 };
model.getColumn(3).width = 24;
model.getRow(2).height = 18;
model.mergeCells("A1:B1");
model.getCell("F2").dataValidation = { type: "decimal", operator: "greaterThanOrEqual", formulae: [0] };
structureWorkbook.definedNames.add("Model!$F$2", "HistoricalInput");
const snapshot = hooks.snapshotWorkbook(structureWorkbook, ["Model"], 6);
assert.deepEqual(hooks.validateWorkbookPreservation(structureWorkbook, snapshot), []);
model.getColumn(3).width = 30;
assert.ok(hooks.validateWorkbookPreservation(structureWorkbook, snapshot).some((error) => /protected template structure/.test(error)));
model.getColumn(3).width = 24;
model.getCell("G2").value = { formula: "F2+1", result: 11 };
assert.ok(hooks.validateWorkbookPreservation(structureWorkbook, snapshot).some((error) => /formula changed/.test(error)));

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

console.log("Unusual SEC presentation, structural immutability, coverage-warning, receivable aggregation, duplicate-source, and retry regressions passed.");
