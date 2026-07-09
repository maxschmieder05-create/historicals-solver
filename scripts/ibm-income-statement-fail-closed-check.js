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

assert.deepEqual(
  hooks.llmMappingReviewFailureBlockingErrors("LLM mapping review attempted_failed (Key limit exceeded)."),
  ["LLM mapping review attempted_failed (Key limit exceeded)."]
);

console.log("IBM income-statement fail-closed regression passed.");
