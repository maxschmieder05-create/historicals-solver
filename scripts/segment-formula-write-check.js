#!/usr/bin/env node

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const Module = require("node:module");
const ts = require("typescript");
const ExcelJS = require("exceljs");

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
  const compiled = compileTypeScript(fs.readFileSync(file, "utf8"));
  const mod = new Module(file, module);
  mod.filename = file;
  mod.paths = Module._nodeModulePaths(path.dirname(file));
  mod._compile(compiled, file);
  return mod.exports;
}

const { __fillModelServiceTestHooks } = loadTypeScriptModule(
  path.join(__dirname, "..", "server", "fill-model", "fill-model-service.ts")
);
const {
  writeSegmentFourthQuarterBridgeFormula,
  writeSegmentMetricPreservingFormula,
  fillSegmentMetricRows,
  reconcileSegmentMetricRowsToModelRow,
  writeHistoricalEbitdaDaAddback,
  findSegmentResidualRow,
  parseInlineSegmentRevenue,
  segmentFamilyFromMembers,
  segmentFamilyCandidates,
  segmentRevenueTemplateCapacityIssue,
  selectReconciledRevenueSegmentFamilyForTemplate,
  setSegmentMetricRowLabel,
  refreshSegmentLinkedLabelFormulaResults,
  refreshCompanyMetadataFormulaResults,
  reportedRevenueFallbackSegment,
  reconcileSegmentMetricFamilyToStatement,
  segmentAnalysisAssignmentStatus,
  segmentAnalysisAssignmentReason,
  segmentAnalysisAssignmentSourceStatement,
  validateFinancialSegmentAssignmentCoverage,
  validateWorkbookBeforeReturn,
  snapshotWorkbook,
  validateWorkbookPreservation,
  FormulaEvaluator
} = __fillModelServiceTestHooks;

const linkedLabelWorkbook = new ExcelJS.Workbook();
const linkedLabelSheet = linkedLabelWorkbook.addWorksheet("Segment Analysis");
linkedLabelSheet.getCell("C8").value = {
  formula: 'C16&" Revenue"',
  result: "Prior Company Revenue",
  ref: "C8:C9",
  shareType: "shared"
};
linkedLabelSheet.getCell("C9").value = { sharedFormula: "C8", result: "Segment 2 Revenue" };
linkedLabelSheet.getCell("C16").value = "Prior Company";
linkedLabelSheet.getCell("C17").value = "Segment 2";

const linkedAssignment = setSegmentMetricRowLabel(linkedLabelSheet, 8, "Revenue", "Reported");
assert.equal(linkedAssignment, "Reported");
assert.equal(linkedLabelSheet.getCell("C16").text, "Reported");
assert.equal(linkedLabelSheet.getCell("C8").text, "Reported Revenue");
assert.equal(linkedLabelSheet.getCell("C8").formula, 'C16&" Revenue"');

const sharedLinkedAssignment = setSegmentMetricRowLabel(linkedLabelSheet, 9, "Revenue", "Cloud");
assert.equal(sharedLinkedAssignment, "Cloud");
assert.equal(linkedLabelSheet.getCell("C17").text, "Cloud");
assert.equal(linkedLabelSheet.getCell("C9").text, "Cloud Revenue");
assert.equal(linkedLabelSheet.getCell("C9").formula, 'C17&" Revenue"');

linkedLabelSheet.getCell("F8").value = null;
const linkedFillAuditRows = [];
const linkedFill = fillSegmentMetricRows(
  linkedLabelSheet,
  ["1Q26"],
  [6],
  [8],
  [
    {
      label: "Reported",
      family: "reported_revenue_fallback",
      values: new Map([["1Q26", 2_709_000_000]]),
      operatingIncome: new Map(),
      depreciationAmortization: new Map()
    }
  ],
  "values",
  "Revenue",
  linkedFillAuditRows,
  { forceOrderedAssignment: true }
);
assert.equal(linkedFill.filledCells, 1, "a formula-linked fallback row must remain assigned after its base label changes");
assert.equal(linkedLabelSheet.getCell("F8").value, 2709);
assert.equal(linkedFillAuditRows.length, 1);

const missingPeriodWorkbook = new ExcelJS.Workbook();
const missingPeriodSheet = missingPeriodWorkbook.addWorksheet("Segment Analysis");
missingPeriodSheet.getCell("C8").value = "Americas Revenue";
missingPeriodSheet.getCell("F8").value = 111;
missingPeriodSheet.getCell("G8").value = 999;
const missingPeriodAuditRows = [];
fillSegmentMetricRows(
  missingPeriodSheet,
  ["1Q26", "2Q26"],
  [6, 7],
  [8],
  [
    {
      label: "Americas",
      family: "geographic",
      values: new Map([["1Q26", 111_000_000]]),
      revenueSources: new Map([
        [
          "1Q26",
          [
            {
              concept: "RevenueFromContractWithCustomerExcludingAssessedTax",
              label: "Americas revenue",
              value: 111_000_000,
              sourceLayer: "sec_inline_xbrl",
              sourceUrl: "https://www.sec.gov/Archives/edgar/data/1/000000000126000001/example.htm",
              form: "10-Q",
              filed: "2026-05-01",
              accn: "0000000001-26-000001",
              start: "2026-01-01",
              end: "2026-03-31",
              periodKey: "1Q26",
              periodType: "quarterly"
            }
          ]
        ]
      ]),
      operatingIncome: new Map(),
      depreciationAmortization: new Map()
    }
  ],
  "values",
  "Revenue",
  missingPeriodAuditRows,
  { forceOrderedAssignment: true }
);
assert.equal(missingPeriodSheet.getCell("F8").value, 111);
assert.equal(
  missingPeriodSheet.getCell("G8").value,
  null,
  "a missing same-period segment disclosure must clear a stale template value instead of inventing zero"
);
assert.equal(missingPeriodAuditRows.find((row) => row.cell === "G8")?.mappingType, "cleared");
assert.equal(missingPeriodAuditRows.find((row) => row.cell === "G8")?.accession, "");

const metadataCacheWorkbook = new ExcelJS.Workbook();
const metadataCacheSheet = metadataCacheWorkbook.addWorksheet("Summary");
metadataCacheSheet.getCell("C1").value = {
  formula: 'Company_Name&" ("&Ticker&")"',
  result: "Lowe's Companies, Inc. (LOW)"
};
metadataCacheSheet.getCell("C2").value = { formula: "Ticker", result: "LOW" };
metadataCacheSheet.getCell("C3").value = { formula: '"Low"', result: "Low" };
metadataCacheSheet.getCell("C4").value = { formula: '"Unrelated"', result: "Unrelated" };
assert.equal(
  refreshCompanyMetadataFormulaResults(metadataCacheWorkbook, "Lowe's Companies, Inc.", "LOW", "Abbott Laboratories", "ABT"),
  2
);
assert.equal(metadataCacheSheet.getCell("C1").text, "Abbott Laboratories (ABT)");
assert.equal(metadataCacheSheet.getCell("C2").text, "ABT");
assert.equal(metadataCacheSheet.getCell("C3").text, "Low", "ordinary title-case words must not be mistaken for the ticker");
assert.equal(metadataCacheSheet.getCell("C4").text, "Unrelated");

const hardcodedMetadataWorkbook = new ExcelJS.Workbook();
const anchoredMetadataSheet = hardcodedMetadataWorkbook.addWorksheet("Anchored Metadata");
anchoredMetadataSheet.getCell("A1").value = { formula: "Company_Name", result: "Lowe's Companies, Inc." };
anchoredMetadataSheet.getCell("B1").value = "Lowe's Companies, Inc.";
anchoredMetadataSheet.getCell("C1").value = "Lowe's Companies, Inc. peer commentary";
anchoredMetadataSheet.getCell("A2").value = { formula: "Ticker", result: "LOW" };
anchoredMetadataSheet.getCell("B2").value = "LOW";
anchoredMetadataSheet.getCell("C2").value = "LOWE";
anchoredMetadataSheet.getCell("D2").value = "Low";
const transitiveMetadataSheet = hardcodedMetadataWorkbook.addWorksheet("Transitive Metadata");
transitiveMetadataSheet.getCell("A1").value = { formula: "Company_Name", result: "Lowe's Companies, Inc." };
transitiveMetadataSheet.getCell("D10").value = { formula: "$A$1", result: "Lowe's Companies, Inc." };
transitiveMetadataSheet.getCell("F10").value = "LOW";
const unrelatedPeerSheet = hardcodedMetadataWorkbook.addWorksheet("Unrelated Peer Data");
unrelatedPeerSheet.getCell("A1").value = "Lowe's Companies, Inc.";
unrelatedPeerSheet.getCell("A2").value = "LOW";
refreshCompanyMetadataFormulaResults(
  hardcodedMetadataWorkbook,
  "Lowe's Companies, Inc.",
  "LOW",
  "Abbott Laboratories",
  "ABT"
);
assert.equal(anchoredMetadataSheet.getCell("B1").value, "Abbott Laboratories");
assert.equal(anchoredMetadataSheet.getCell("B2").value, "ABT");
assert.equal(anchoredMetadataSheet.getCell("C1").value, "Lowe's Companies, Inc. peer commentary");
assert.equal(anchoredMetadataSheet.getCell("C2").value, "LOWE");
assert.equal(anchoredMetadataSheet.getCell("D2").value, "Low");
assert.equal(transitiveMetadataSheet.getCell("F10").value, "ABT");
assert.equal(unrelatedPeerSheet.getCell("A1").value, "Lowe's Companies, Inc.");
assert.equal(unrelatedPeerSheet.getCell("A2").value, "LOW");

const staleSegmentWorkbook = new ExcelJS.Workbook();
const staleSegmentSheet = staleSegmentWorkbook.addWorksheet("Segment Analysis");
staleSegmentSheet.getCell("C8").value = "Current Segment Revenue";
staleSegmentSheet.getCell("C9").value = "Prior Company Segment Revenue";
staleSegmentSheet.getCell("F8").value = 0;
staleSegmentSheet.getCell("F9").value = { formula: "999", result: 999 };
const staleSegmentAuditRows = [];
fillSegmentMetricRows(
  staleSegmentSheet,
  ["1Q26"],
  [6],
  [8, 9],
  [
    {
      label: "Current Segment",
      family: "reportable",
      values: new Map([["1Q26", 2_000_000_000]]),
      operatingIncome: new Map(),
      depreciationAmortization: new Map()
    }
  ],
  "values",
  "Revenue",
  staleSegmentAuditRows,
  { clearUnmatchedLabels: true, clearUnmatchedFormulas: true }
);
assert.equal(staleSegmentSheet.getCell("F8").value, 2000);
assert.equal(staleSegmentSheet.getCell("F9").value, 0, "an unmatched prior-company historical segment formula must be cleared");

const metricOnlyWorkbook = new ExcelJS.Workbook();
const metricOnlySheet = metricOnlyWorkbook.addWorksheet("Segment Analysis");
metricOnlySheet.getCell("C34").value = { formula: 'C16&" Operating Income"', result: "Diagnostics Operating Income" };
metricOnlySheet.getCell("C35").value = { formula: 'C17&" Operating Income"', result: "" };
metricOnlySheet.getCell("C16").value = "Diagnostics";
metricOnlySheet.getCell("C17").value = "";
metricOnlySheet.getCell("F34").value = 0;
metricOnlySheet.getCell("F35").value = 0;
const metricOnlyAuditRows = [];
const metricOnlyFill = fillSegmentMetricRows(
  metricOnlySheet,
  ["FY25"],
  [6],
  [34, 35],
  [
    {
      label: "Diagnostics",
      family: "reportable_segments",
      values: new Map(),
      operatingIncome: new Map([["FY25", 4_000_000_000]]),
      depreciationAmortization: new Map()
    },
    {
      label: "Corporate / Reconciliation",
      family: "reportable_segments:corporate_reconciliation",
      values: new Map(),
      operatingIncome: new Map([["FY25", -1_000_000_000]]),
      depreciationAmortization: new Map()
    }
  ],
  "operatingIncome",
  "Operating Income",
  metricOnlyAuditRows,
  { preserveLinkedBaseLabels: true }
);
assert.equal(metricOnlyFill.filledCells, 2, "a metric-only reconciliation should use an available blank linked row");
assert.equal(metricOnlySheet.getCell("C16").text, "Diagnostics", "a meaningful revenue-linked base label must remain unchanged");
assert.equal(metricOnlySheet.getCell("C17").text, "Corporate / Reconciliation");
assert.equal(metricOnlySheet.getCell("C35").text, "Corporate / Reconciliation Operating Income");
assert.equal(metricOnlySheet.getCell("F35").value, -1000);

const independentMetricWorkbook = new ExcelJS.Workbook();
const independentMetricSheet = independentMetricWorkbook.addWorksheet("Segment Analysis");
independentMetricSheet.getCell("C34").value = { formula: 'C44&" Operating Income"', result: "Home Improvement Retail Operating Income" };
independentMetricSheet.getCell("C44").value = "Home Improvement Retail";
independentMetricSheet.getCell("F34").value = 0;
fillSegmentMetricRows(
  independentMetricSheet,
  ["FY25"],
  [6],
  [34],
  [
    {
      label: "Established Pharmaceutical Products",
      family: "reportable_segments",
      values: new Map(),
      operatingIncome: new Map([["FY25", 1_290_000_000]]),
      depreciationAmortization: new Map()
    }
  ],
  "operatingIncome",
  "Operating Income",
  [],
  { forceOrderedAssignment: true, preserveLinkedBaseLabels: false }
);
assert.equal(
  independentMetricSheet.getCell("C44").text,
  "Established Pharmaceutical Products",
  "an independently linked operating-income label block must replace prior-company labels instead of preserving them as if shared with revenue"
);
assert.equal(independentMetricSheet.getCell("C34").text, "Established Pharmaceutical Products Operating Income");

const propagatedLabelWorkbook = new ExcelJS.Workbook();
const propagatedLabelSheet = propagatedLabelWorkbook.addWorksheet("Segment Analysis");
propagatedLabelSheet.getCell("C34").value = { formula: 'C44&" Operating Income"', result: "Home Improvement Retail Operating Income" };
propagatedLabelSheet.getCell("C44").value = "Home Improvement Retail";
propagatedLabelSheet.getCell("C52").value = { formula: 'C44&" D&A"', result: "Home Improvement Retail D&A" };
propagatedLabelSheet.getCell("C70").value = { formula: 'C44&" EBITDA"', result: "Home Improvement Retail EBITDA" };
const propagatedDownstreamSheet = propagatedLabelWorkbook.addWorksheet("Revenue Build");
propagatedDownstreamSheet.getCell("C1").value = {
  formula: "'Segment Analysis'!C52",
  result: "Home Improvement Retail D&A"
};
propagatedDownstreamSheet.getCell("C3").value = {
  formula: '"Multiple Used for "&C1',
  result: "Multiple Used for Home Improvement Retail D&A"
};
propagatedDownstreamSheet.getCell("A2").value = "Actual unrelated label";
propagatedDownstreamSheet.getCell("C2").value = { formula: "A2", result: "Do not touch unrelated cache" };
const propagatedLabelSnapshot = snapshotWorkbook(propagatedLabelWorkbook, ["Segment Analysis", "Revenue Build"], 6);
setSegmentMetricRowLabel(propagatedLabelSheet, 34, "Operating Income", "Established Pharmaceutical Products");
assert.equal(refreshSegmentLinkedLabelFormulaResults(propagatedLabelSheet, [34, 52]), 4);
assert.equal(propagatedLabelSheet.getCell("C34").text, "Established Pharmaceutical Products Operating Income");
assert.equal(propagatedLabelSheet.getCell("C52").text, "Established Pharmaceutical Products D&A");
assert.equal(propagatedLabelSheet.getCell("C70").text, "Established Pharmaceutical Products EBITDA");
assert.equal(propagatedDownstreamSheet.getCell("C1").text, "Established Pharmaceutical Products D&A");
assert.equal(propagatedDownstreamSheet.getCell("C3").text, "Multiple Used for Established Pharmaceutical Products D&A");
assert.equal(propagatedDownstreamSheet.getCell("C2").text, "Do not touch unrelated cache");
assert.deepEqual(
  validateWorkbookPreservation(propagatedLabelWorkbook, propagatedLabelSnapshot),
  [],
  "linked segment-label formula caches must refresh without appearing as unauthorized template-label edits"
);

function buildReturnedLabelWorkbook(baseLabel) {
  const result = new ExcelJS.Workbook();
  const resultSheet = result.addWorksheet("Segment Analysis");
  resultSheet.getCell("C7").value = "Total Company Revenue";
  resultSheet.getCell("C8").value = { formula: 'C16&" Revenue"', result: "Segment 1 Revenue" };
  resultSheet.getCell("C9").value = { formula: 'C17&" Revenue"', result: `${baseLabel} Revenue` };
  resultSheet.getCell("C10").value = "Revenue Mix";
  resultSheet.getCell("C16").value = "Segment 1";
  resultSheet.getCell("C17").value = baseLabel;
  resultSheet.getCell("H17");
  return result;
}

const preSerializationLabelWorkbook = buildReturnedLabelWorkbook("Segment 2");
const preSerializationLabelSheet = preSerializationLabelWorkbook.getWorksheet("Segment Analysis");
const returnedLabelSnapshot = snapshotWorkbook(preSerializationLabelWorkbook, ["Segment Analysis"], 6);
setSegmentMetricRowLabel(preSerializationLabelSheet, 9, "Revenue", "Reported");

const reloadedLabelWorkbook = buildReturnedLabelWorkbook("Reported");
assert.deepEqual(
  validateWorkbookPreservation(reloadedLabelWorkbook, returnedLabelSnapshot),
  [],
  "an exact generated segment-label mutation must remain authorized after workbook serialization creates a new workbook instance"
);
reloadedLabelWorkbook.getWorksheet("Segment Analysis").getCell("C17").value = "Tampered";
assert.match(
  validateWorkbookPreservation(reloadedLabelWorkbook, returnedLabelSnapshot).join("\n"),
  /C17: row label changed from "Segment 2" to "Tampered"/,
  "returned-workbook preservation must still reject an unrecorded label value at an authorized address"
);

const workbook = new ExcelJS.Workbook();
const sheet = workbook.addWorksheet("Segment Analysis");
sheet.getCell("C8").value = "Foods and Sundries Revenue";
sheet.getCell("F8").value = 10;
sheet.getCell("G8").value = 20;
sheet.getCell("H8").value = 30;
sheet.getCell("I8").value = { formula: "70-SUM(F8:H8)", result: 10 };
sheet.getCell("J8").value = { formula: "SUM(F8:I8)", result: 70 };

const periods = ["1Q23", "2Q23", "3Q23", "4Q23", "FY23"];
const columns = [6, 7, 8, 9, 10];
const segment = {
  label: "Foods and Sundries",
  family: "merchandise_category",
  values: new Map([
    ["4Q23", 40_000_000],
    ["FY23", 100_000_000]
  ]),
  operatingIncome: new Map(),
  depreciationAmortization: new Map(),
  annualValues: new Map([["FY23", 70_000_000]])
};

assert.equal(writeSegmentFourthQuarterBridgeFormula(sheet, 8, columns, periods, 3, segment, "values", 40), false);
assert.deepEqual(sheet.getCell("I8").value, { formula: "70-SUM(F8:H8)", result: 10 });
assert.equal(
  writeSegmentMetricPreservingFormula(sheet.getCell("I8"), 40, {
    sheet,
    rowNumber: 8,
    columns,
    periods,
    periodIndex: 3,
    segment,
    metric: "values"
  }),
  true,
  "a stale 4Q historical bridge must be replaced by the exact EDGAR-sourced quarter"
);
assert.equal(sheet.getCell("I8").value, 40);
assert.equal(
  writeSegmentMetricPreservingFormula(sheet.getCell("J8"), 100, {
    sheet,
    rowNumber: 8,
    columns,
    periods,
    periodIndex: 4,
    segment,
    metric: "values"
  }),
  true
);
assert.deepEqual(
  sheet.getCell("J8").value,
  { formula: "SUM(F8:I8)", result: 100 },
  "the annual formula may be preserved only after the sourced quarters make it strictly tie"
);

sheet.getCell("C35").value = "Canada Operating Income";
for (const address of ["F35", "G35", "H35", "I35"]) sheet.getCell(address).value = 0;
sheet.getCell("J35").value = { formula: "SUM(F35:I35)", result: 0 };
assert.equal(
  writeSegmentMetricPreservingFormula(sheet.getCell("J35"), 50, {
    sheet,
    rowNumber: 35,
    columns,
    periods,
    periodIndex: 4,
    segment: { ...segment, operatingIncome: new Map([["FY23", 50_000_000]]) },
    metric: "operatingIncome"
  }),
  true
);
assert.equal(sheet.getCell("J35").value, 50, "a stale annual formula must be replaced by the EDGAR-sourced operating income total");

sheet.getCell("C80").value = "Legacy Segment Operating Income";
sheet.getCell("C81").value = "Legacy Other Operating Income";
sheet.getCell("F80").value = { formula: "4000+18", result: 4018 };
sheet.getCell("F81").value = { formula: "100+1", result: 101 };
const fallbackOperatingIncomeAudit = [];
const reportedOperatingIncomeSegment = {
  label: "Reported",
  family: "reported_operating_income_fallback",
  values: new Map(),
  operatingIncome: new Map([["1Q23", 484_600_000]]),
  depreciationAmortization: new Map(),
  reportedOperatingIncomeSources: new Map([
    [
      "1Q23",
      {
        concept: "OperatingIncomeLoss",
        label: "Operating income",
        value: 484_600_000,
        accn: "0000000000-23-000001",
        form: "10-Q"
      }
    ]
  ])
};
const fallbackOperatingIncomeFill = fillSegmentMetricRows(
  sheet,
  ["1Q23"],
  [6],
  [80, 81],
  [reportedOperatingIncomeSegment],
  "operatingIncome",
  "Operating Income",
  fallbackOperatingIncomeAudit,
  {
    forceOrderedAssignment: true,
    clearUnmatchedLabels: true,
    clearUnmatchedFormulas: true,
    consolidatedFallback: true
  }
);
assert.equal(fallbackOperatingIncomeFill.filledCells, 2);
assert.equal(sheet.getCell("C80").text, "Reported Operating Income");
assert.equal(sheet.getCell("F80").value, 484.6);
assert.equal(sheet.getCell("C81").text, "");
assert.equal(sheet.getCell("F81").value, 0);
assert.equal(fallbackOperatingIncomeAudit[0].formulaStatus, "stale historical formula replaced with the EDGAR-sourced value after strict evaluation failed");
assert.equal(fallbackOperatingIncomeAudit[0].accession, "0000000000-23-000001");
assert.equal(fallbackOperatingIncomeAudit[1].validationStatus, "cleared");

sheet.getCell("K8").value = { formula: "'Revenue Build'!K8", result: 1 };
assert.equal(
  writeSegmentMetricPreservingFormula(sheet.getCell("K8"), 25),
  false,
  "a missing cross-sheet precedent must not be masked by replacing the formula's cached result"
);
assert.deepEqual(sheet.getCell("K8").value, { formula: "'Revenue Build'!K8", result: 1 });
const revenueBuild = workbook.addWorksheet("Revenue Build");
revenueBuild.getCell("K8").value = 10;
assert.equal(
  writeSegmentMetricPreservingFormula(sheet.getCell("K8"), 25),
  false,
  "a cross-sheet precedent that does not tie to reconciled SEC detail must fail closed"
);
assert.deepEqual(sheet.getCell("K8").value, { formula: "'Revenue Build'!K8", result: 1 });
revenueBuild.getCell("K8").value = 25;
assert.equal(writeSegmentMetricPreservingFormula(sheet.getCell("K8"), 25), true);
assert.deepEqual(sheet.getCell("K8").value, { formula: "'Revenue Build'!K8", result: 25 });

revenueBuild.getCell("K9").value = { formula: "UNSUPPORTED(1)", result: 25 };
sheet.getCell("K9").value = { formula: "'Revenue Build'!K9", result: 25 };
assert.equal(
  writeSegmentMetricPreservingFormula(sheet.getCell("K9"), 25),
  false,
  "a cross-sheet precedent must be genuinely evaluable; its cached result is not sufficient evidence"
);

const model = workbook.addWorksheet("Model");
model.getCell("C55").value = "Income Statement";
model.getCell("C59").value = "EBIT";
model.getCell("C60").value = "Depreciation & Amortization";
model.getCell("C61").value = "EBITDA";
model.getCell("F60").value = { formula: "-F34", result: 0 };
const daAuditRows = [];
const daAddback = writeHistoricalEbitdaDaAddback(
  model,
  ["1Q23"],
  [6],
  {
    duration: new Map([
      [
        "1Q23",
        new Map([
          [
            "DepreciationDepletionAndAmortization",
            {
              concept: "DepreciationDepletionAndAmortization",
              label: "Depreciation, depletion and amortization",
              value: 25_000_000,
              periodType: "quarterly",
              form: "10-Q",
              accn: "0000000000-23-000001"
            }
          ]
        ])
      ]
    ])
  },
  daAuditRows
);
assert.equal(daAddback.filledCells, 1);
assert.equal(model.getCell("F60").value, 25);
assert.match(daAddback.warnings[0], /formula was replaced with the SEC actual/i);
assert.equal(daAuditRows.length, 1);
assert.equal(daAuditRows[0].formulaPreserved, false);
sheet.getCell("C51").value = "Total D&A";
sheet.getCell("C59").value = "D&A Check";
for (let row = 52; row <= 57; row += 1) sheet.getCell(row, 6).value = 0;
const auditRows = [];
const reconciliation = reconcileSegmentMetricRowsToModelRow(sheet, ["1Q23"], [6], [52, 53, 54, 55, 56, 57], "D&A", auditRows, [
  "Depreciation & Amortization"
]);
assert.equal(reconciliation.filledCells, 0);
assert.equal(
  [52, 53, 54, 55, 56, 57].reduce((total, row) => total + (Number(sheet.getCell(row, 6).value) || 0), 0),
  0
);
assert.equal(sheet.getCell("C57").text, "");
assert.equal(auditRows.length, 0);

sheet.getCell("C70").value = "Other International Operating Income";
sheet.getCell("F70").value = 0;
sheet.getCell("C71").value = "451";
sheet.getCell("F71").value = 0;
assert.equal(findSegmentResidualRow(sheet, [70, 71], 6, "Operating Income"), 70);
assert.equal(sheet.getCell("C71").text, "451");
sheet.getCell("C71").value = "Other / Reconciliation Operating Income";
assert.equal(findSegmentResidualRow(sheet, [70, 71], 6, "Operating Income"), 71);

const syntheticDuration = new Map();
const syntheticPeriods = ["1Q23", "2Q23", "3Q23", "4Q23", "FY23"];
const quarterlyRevenue = [10, 20, 30];
const quarterlyOperatingIncome = [1, 2, 3];
for (let index = 0; index < 3; index += 1) {
  const period = syntheticPeriods[index];
  syntheticDuration.set(
    period,
    new Map([
      [
        "RevenueFromContractWithCustomerExcludingAssessedTax",
        {
          concept: "RevenueFromContractWithCustomerExcludingAssessedTax",
          label: "Revenue",
          value: quarterlyRevenue[index] * 1_000_000,
          accn: `quarter-revenue-${period}`,
          periodType: "quarterly"
        }
      ],
      [
        "OperatingIncomeLoss",
        {
          concept: "OperatingIncomeLoss",
          label: "Operating income",
          value: quarterlyOperatingIncome[index] * 1_000_000,
          accn: `quarter-oi-${period}`,
          periodType: "quarterly"
        }
      ]
    ])
  );
}
syntheticDuration.set(
  "4Q23",
  new Map([
    [
      "RevenueFromContractWithCustomerExcludingAssessedTax",
      {
        concept: "RevenueFromContractWithCustomerExcludingAssessedTax",
        label: "Revenue",
        value: 40_000_000,
        accn: "annual-revenue-FY23",
        periodType: "quarterly"
      }
    ]
  ])
);
syntheticDuration.set(
  "FY23",
  new Map([
    [
      "RevenueFromContractWithCustomerExcludingAssessedTax",
      {
        concept: "RevenueFromContractWithCustomerExcludingAssessedTax",
        label: "Revenue",
        value: 100_000_000,
        accn: "annual-revenue-FY23",
        periodType: "annual"
      }
    ],
    [
      "OperatingIncomeLoss",
      {
        concept: "OperatingIncomeLoss",
        label: "Operating income",
        value: 10_000_000,
        accn: "annual-oi-FY23",
        periodType: "annual"
      }
    ]
  ])
);
const reportedFallback = reportedRevenueFallbackSegment(syntheticPeriods, { duration: syntheticDuration, instant: new Map() });
assert.ok(reportedFallback);
assert.equal(reportedFallback.values.get("4Q23"), 40_000_000);
assert.equal(reportedFallback.operatingIncome.get("4Q23"), 4_000_000);
assert.equal(reportedFallback.reportedRevenueSources.get("FY23").accn, "annual-revenue-FY23");
assert.equal(reportedFallback.reportedOperatingIncomeSources.get("4Q23").value, 4_000_000);
assert.match(reportedFallback.reportedOperatingIncomeSources.get("4Q23").accn, /annual-oi-FY23/);
assert.equal(segmentAnalysisAssignmentStatus(reportedFallback, "values"), "reported_total_fallback");
assert.equal(segmentAnalysisAssignmentStatus(reportedFallback, "operatingIncome"), "reported_total_fallback");
assert.match(segmentAnalysisAssignmentReason(reportedFallback, "operatingIncome", "reported_total_fallback"), /not presented as a disclosed segment or as a reconciliation plug/i);
assert.equal(segmentAnalysisAssignmentSourceStatement(reportedFallback, "operatingIncome"), "Consolidated EDGAR operating income fallback");

const capacityContext = {
  duration: new Map([
    [
      "1Q26",
      new Map([
        [
          "RevenueFromContractWithCustomerExcludingAssessedTax",
          {
            concept: "RevenueFromContractWithCustomerExcludingAssessedTax",
            label: "Revenue",
            value: 700_000_000,
            accn: "capacity-quarter-revenue-1Q26",
            periodType: "quarterly"
          }
        ]
      ])
    ]
  ]),
  instant: new Map()
};
const sevenDisclosedSegments = ["Alpha", "Bravo", "Charlie", "Delta", "Echo", "Foxtrot", "Golf"].map(
  (label, sourceOrder) => ({
    label,
    family: "segment",
    disclosureKind: "segment",
    disclosurePriority: 4,
    sourceOrder,
    values: new Map([["1Q26", 100_000_000]]),
    annualValues: new Map(),
    operatingIncome: new Map(),
    depreciationAmortization: new Map()
  })
);
function segmentCapacitySheet(detailRows) {
  const result = new ExcelJS.Workbook().addWorksheet("Segment Analysis");
  result.getCell("C7").value = "Total Company Revenue";
  for (let offset = 0; offset < detailRows; offset += 1) {
    result.getCell(8 + offset, 3).value = `Segment ${offset + 1} Revenue`;
    result.getCell(8 + offset, 6).value = 0;
  }
  result.getCell(8 + detailRows, 3).value = "Revenue Mix";
  return result;
}

const sixRowCapacityIssue = segmentRevenueTemplateCapacityIssue(
  segmentCapacitySheet(6),
  ["1Q26"],
  [6],
  sevenDisclosedSegments,
  capacityContext
);
assert.ok(sixRowCapacityIssue, "seven disclosed segments must not collapse into a one-row reported-revenue fallback");
assert.equal(sixRowCapacityIssue.availableRows, 6);
assert.equal(sixRowCapacityIssue.requiredRows, 7);
assert.deepEqual(sixRowCapacityIssue.disclosedLabels, ["Alpha", "Bravo", "Charlie", "Delta", "Echo", "Foxtrot", "Golf"]);
assert.match(sixRowCapacityIssue.message, /requires 7 detail rows/i);
assert.match(sixRowCapacityIssue.message, /only 6 structurally writable rows/i);
assert.match(sixRowCapacityIssue.message, /would discard SEC segment detail/i);

assert.equal(
  segmentRevenueTemplateCapacityIssue(
    segmentCapacitySheet(7),
    ["1Q26"],
    [6],
    sevenDisclosedSegments,
    capacityContext
  ),
  null,
  "a template that structurally exposes seven rows should support the full seven-segment family"
);
assert.deepEqual(
  selectReconciledRevenueSegmentFamilyForTemplate(
    sevenDisclosedSegments,
    ["1Q26"],
    capacityContext,
    7
  )?.segments.map((item) => item.label),
  ["Alpha", "Bravo", "Charlie", "Delta", "Echo", "Foxtrot", "Golf"]
);

const nestedHealthcareDisclosureHtml = `
  <ix:nonNumeric name="dei:DocumentFiscalYearFocus">2025</ix:nonNumeric>
  <ix:nonNumeric name="dei:DocumentFiscalPeriodFocus">FY</ix:nonNumeric>
  <ix:nonNumeric name="dei:DocumentPeriodEndDate">2025-12-31</ix:nonNumeric>
  <xbrli:context id="total"><xbrli:period><xbrli:startDate>2025-01-01</xbrli:startDate><xbrli:endDate>2025-12-31</xbrli:endDate></xbrli:period></xbrli:context>
  <xbrli:context id="a1"><xbrli:entity><xbrli:segment><xbrldi:explicitMember dimension="us-gaap:SubsegmentsAxis">test:ProductA1Member</xbrldi:explicitMember><xbrldi:explicitMember dimension="us-gaap:StatementBusinessSegmentsAxis">test:SegmentAMember</xbrldi:explicitMember><xbrldi:explicitMember dimension="srt:ConsolidationItemsAxis">us-gaap:OperatingSegmentsMember</xbrldi:explicitMember></xbrli:segment></xbrli:entity><xbrli:period><xbrli:startDate>2025-01-01</xbrli:startDate><xbrli:endDate>2025-12-31</xbrli:endDate></xbrli:period></xbrli:context>
  <xbrli:context id="a2"><xbrli:entity><xbrli:segment><xbrldi:explicitMember dimension="us-gaap:SubsegmentsAxis">test:ProductA2Member</xbrldi:explicitMember><xbrldi:explicitMember dimension="us-gaap:StatementBusinessSegmentsAxis">test:SegmentAMember</xbrldi:explicitMember><xbrldi:explicitMember dimension="srt:ConsolidationItemsAxis">us-gaap:OperatingSegmentsMember</xbrldi:explicitMember></xbrli:segment></xbrli:entity><xbrli:period><xbrli:startDate>2025-01-01</xbrli:startDate><xbrli:endDate>2025-12-31</xbrli:endDate></xbrli:period></xbrli:context>
  <xbrli:context id="a"><xbrli:entity><xbrli:segment><xbrldi:explicitMember dimension="us-gaap:StatementBusinessSegmentsAxis">test:SegmentAMember</xbrldi:explicitMember><xbrldi:explicitMember dimension="srt:ConsolidationItemsAxis">us-gaap:OperatingSegmentsMember</xbrldi:explicitMember></xbrli:segment></xbrli:entity><xbrli:period><xbrli:startDate>2025-01-01</xbrli:startDate><xbrli:endDate>2025-12-31</xbrli:endDate></xbrli:period></xbrli:context>
  <xbrli:context id="b"><xbrli:entity><xbrli:segment><xbrldi:explicitMember dimension="us-gaap:StatementBusinessSegmentsAxis">test:SegmentBMember</xbrldi:explicitMember><xbrldi:explicitMember dimension="srt:ConsolidationItemsAxis">us-gaap:OperatingSegmentsMember</xbrldi:explicitMember></xbrli:segment></xbrli:entity><xbrli:period><xbrli:startDate>2025-01-01</xbrli:startDate><xbrli:endDate>2025-12-31</xbrli:endDate></xbrli:period></xbrli:context>
  <xbrli:context id="narrative"><xbrli:entity><xbrli:segment><xbrldi:explicitMember dimension="us-gaap:StatementBusinessSegmentsAxis">test:SegmentAMember</xbrldi:explicitMember><xbrldi:explicitMember dimension="srt:ProductOrServiceAxis">test:OutbreakTestingMember</xbrldi:explicitMember></xbrli:segment></xbrli:entity><xbrli:period><xbrli:startDate>2025-01-01</xbrli:startDate><xbrli:endDate>2025-12-31</xbrli:endDate></xbrli:period></xbrli:context>
  <table>
    <tr><td>Total</td><td><ix:nonFraction name="us-gaap:Revenues" contextRef="total" scale="6">100</ix:nonFraction></td></tr>
    <tr><td>Product A1</td><td><ix:nonFraction name="us-gaap:Revenues" contextRef="a1" scale="6">10</ix:nonFraction></td></tr>
    <tr><td>Product A2</td><td><ix:nonFraction name="us-gaap:Revenues" contextRef="a2" scale="6">30</ix:nonFraction></td></tr>
    <tr><td>Segment A</td><td><ix:nonFraction name="us-gaap:Revenues" contextRef="a" scale="6">40</ix:nonFraction></td></tr>
    <tr><td>Segment B</td><td><ix:nonFraction name="us-gaap:Revenues" contextRef="b" scale="6">60</ix:nonFraction></td></tr>
    <tr><td>Outbreak testing sales</td><td><ix:nonFraction name="us-gaap:Revenues" contextRef="narrative" scale="6">5</ix:nonFraction></td></tr>
  </table>`;
const nestedHealthcareParsed = parseInlineSegmentRevenue(nestedHealthcareDisclosureHtml, "10-K");
const nestedHealthcareRows = nestedHealthcareParsed.annual.get("4Q25");
assert.ok(nestedHealthcareRows, "nested segment fixture should parse one annual disclosure family");
const nestedHealthcareMetrics = Array.from(nestedHealthcareRows.values());
assert.equal(nestedHealthcareMetrics.some((item) => item.label === "Outbreak Testing"), false, "a nested narrative metric that makes the family exceed consolidated revenue must be removed");
assert.equal(nestedHealthcareMetrics.find((item) => item.label === "Product A1")?.aggregateParent, "Segment A");
assert.equal(nestedHealthcareMetrics.find((item) => item.label === "Segment A")?.aggregate, true);
const nestedHealthcareSegments = nestedHealthcareMetrics.map((item) => ({
  label: item.label,
  family: item.family,
  aggregate: item.aggregate,
  aggregateParent: item.aggregateParent,
  sourceOrder: item.sourceOrder,
  values: new Map([["FY25", item.revenue || 0]]),
  annualValues: new Map([["FY25", item.revenue || 0]]),
  operatingIncome: new Map(),
  depreciationAmortization: new Map()
}));
const nestedHealthcareCandidate = segmentFamilyCandidates(nestedHealthcareSegments, 2).find((candidate) => candidate.family === "segment");
assert.deepEqual(nestedHealthcareCandidate?.segments.map((item) => item.label), ["Segment A", "Segment B"]);
assert.equal(nestedHealthcareCandidate.segments.reduce((sum, item) => sum + item.values.get("FY25"), 0), 100_000_000);
assert.equal(
  segmentFamilyFromMembers(
    [
      "srt:ConsolidationItemsAxis=us-gaap:OperatingSegmentsMember",
      "us-gaap:StatementBusinessSegmentsAxis=test:EstablishedPharmaceuticalProductsMember"
    ],
    "Established Pharmaceutical Products"
  ),
  "reportable",
  "a company name containing the letters 'mac' must not be misclassified as a Mac product family"
);

const disclosedOperatingSegments = [
  ["Established Pharmaceuticals", 3_000_000_000],
  ["Nutritionals", 2_500_000_000],
  ["Diagnostics", 4_000_000_000],
  ["Medical Devices", 2_300_000_000]
].map(([label, value], sourceOrder) => ({
  label,
  family: "reportable_segments",
  sourceOrder,
  values: new Map(),
  operatingIncome: new Map([["FY25", value]]),
  depreciationAmortization: new Map()
}));
const consolidatedOperatingIncomeResolver = () => ({
  value: 8_053_000_000,
  sources: [
    {
      concept: "OperatingIncomeLoss",
      label: "Operating income",
      value: 8_053_000_000,
      unit: "USD",
      sourceLayer: "sec_filing_package",
      accn: "000000180026000001",
      start: "2025-01-01",
      end: "2025-12-31",
      periodKey: "FY25",
      periodType: "annual"
    }
  ]
});
const operatingSegmentsWithCorporate = reconcileSegmentMetricFamilyToStatement(
  disclosedOperatingSegments,
  ["FY25"],
  { duration: new Map(), instant: new Map() },
  "operatingIncome",
  consolidatedOperatingIncomeResolver,
  6
);
assert.equal(operatingSegmentsWithCorporate.length, 5, "disclosed operating segments should be retained when one reconciliation row fits the template");
const corporateReconciliation = operatingSegmentsWithCorporate.at(-1);
assert.equal(corporateReconciliation.label, "Corporate / Reconciliation");
assert.equal(corporateReconciliation.operatingIncome.get("FY25"), -3_747_000_000);
assert.equal(
  operatingSegmentsWithCorporate.reduce((total, item) => total + (item.operatingIncome.get("FY25") || 0), 0),
  8_053_000_000,
  "the SEC-anchored corporate reconciliation must bridge disclosed segment earnings to consolidated operating income"
);
assert.equal(segmentAnalysisAssignmentStatus(corporateReconciliation, "operatingIncome"), "grouped_into_segment_reconciliation");
assert.match(
  corporateReconciliation.operatingIncomeSources.get("FY25")[0].concept,
  /SegmentOperatingIncomeCorporateReconciliation/
);
assert.deepEqual(
  reconcileSegmentMetricFamilyToStatement(
    disclosedOperatingSegments,
    ["FY25"],
    { duration: new Map(), instant: new Map() },
    "operatingIncome",
    consolidatedOperatingIncomeResolver,
    4
  ),
  [],
  "the system must fall back cleanly when no template row is available for the reconciliation"
);

const financialWorkbook = new ExcelJS.Workbook();
const financialSegmentSheet = financialWorkbook.addWorksheet("Segment Analysis");
financialSegmentSheet.getCell("C7").value = "Total Company Revenue";
financialSegmentSheet.getCell("C8").value = "Advisory Revenue";
financialSegmentSheet.getCell("C9").value = "Trading Revenue";
financialSegmentSheet.getCell("C10").value = "Revenue Mix";
financialSegmentSheet.getCell("F7").value = 500;
financialSegmentSheet.getCell("F8").value = 12.5;
financialSegmentSheet.getCell("F9").value = 7.25;

function financialSegmentAssignment(cell, modelRow, sourceAmount, overrides = {}) {
  return {
    fiscalPeriod: "1Q26",
    sourceFilingAccession: "0000000000-26-000001",
    sourceFilingForm: "10-Q",
    sourcePeriodEndDate: "2026-03-31",
    sourceUrl: "https://www.sec.gov/Archives/edgar/data/0/000000000026000001/example.htm",
    sourceStatement: "SEC segment/disaggregation table",
    sourceLineItemLabel: modelRow,
    sourceMetric: "Revenue",
    sourceAmount,
    modelAmount: sourceAmount,
    sourceXbrlTag: "segment:values",
    assignedSheet: "Segment Analysis",
    assignedModelRow: modelRow,
    assignedCell: cell,
    assignmentStatus: "mapped_to_segment_row",
    classificationReason: "Current SEC component maps one-to-one to the workbook row.",
    validationStatus: "OK!",
    sourceRowKey: `1Q26|segment:values|${modelRow.toLowerCase().replace(/[^a-z0-9]/g, "")}|${cell}`,
    ...overrides
  };
}

const advisoryAssignment = financialSegmentAssignment("F8", "Advisory Revenue", 12_500_000);
const tradingAssignment = financialSegmentAssignment("F9", "Trading Revenue", 7_250_000);
const financialEvaluator = new FormulaEvaluator(financialSegmentSheet, {
  useCachedFormulaResults: false,
  allowCachedFormulaResultFallback: false
});
assert.deepEqual(
  validateFinancialSegmentAssignmentCoverage(
    financialSegmentSheet,
    ["1Q26"],
    [6],
    [advisoryAssignment, tradingAssignment],
    financialEvaluator
  ),
  [],
  "financial-company component cells should pass exact provenance validation without being forced to add to the consolidated total"
);
const financialValidationWarnings = [];
assert.deepEqual(
  validateWorkbookBeforeReturn(
    financialWorkbook,
    ["1Q26"],
    [6],
    { duration: new Map(), instant: new Map() },
    financialValidationWarnings,
    "Model",
    {
      kind: "financial_company",
      confidence: "high",
      rationale: ["focused financial-segment provenance regression"],
      sheetName: "Model",
      hasSegmentAnalysis: true
    },
    ["1Q26"],
    [6],
    ["1Q26"],
    [6],
    [],
    [],
    [advisoryAssignment, tradingAssignment],
    ["1Q26"],
    [6]
  ),
  [],
  "the financial-company workbook return gate must invoke exact segment provenance validation"
);
assert.match(financialValidationWarnings.join("\n"), /additive tie-outs were skipped.*hard-gated/i);

financialSegmentSheet.getCell("F8").value = 999;
assert.match(
  validateFinancialSegmentAssignmentCoverage(
    financialSegmentSheet,
    ["1Q26"],
    [6],
    [advisoryAssignment, tradingAssignment],
    financialEvaluator
  ).join("\n"),
  /stale or poisoned component values cannot be returned/i
);
financialSegmentSheet.getCell("F8").value = 12.5;

assert.match(
  validateFinancialSegmentAssignmentCoverage(financialSegmentSheet, ["1Q26"], [6], [advisoryAssignment], financialEvaluator).join("\n"),
  /F9 1Q26:.*no same-period SEC Segment Assignment Ledger provenance/i
);
assert.match(
  validateFinancialSegmentAssignmentCoverage(
    financialSegmentSheet,
    ["1Q26"],
    [6],
    [advisoryAssignment, { ...advisoryAssignment }, tradingAssignment],
    financialEvaluator
  ).join("\n"),
  /F8 1Q26:.*exactly one same-period SEC assignment is required/i
);
assert.match(
  validateFinancialSegmentAssignmentCoverage(
    financialSegmentSheet,
    ["1Q26"],
    [6],
    [{ ...advisoryAssignment, sourceRowKey: "FY25|segment:values|advisoryrevenue|F8" }, tradingAssignment],
    financialEvaluator
  ).join("\n"),
  /F8 1Q26:.*complete current-SEC assignment provenance/i
);
assert.match(
  validateFinancialSegmentAssignmentCoverage(
    financialSegmentSheet,
    ["1Q26"],
    [6],
    [{ ...advisoryAssignment, sourceFilingAccession: "" }, tradingAssignment],
    financialEvaluator
  ).join("\n"),
  /F8 1Q26:.*complete current-SEC assignment provenance/i,
  "a segment assignment cannot claim current-SEC provenance without an actual filing accession"
);

async function validateSerializedSegmentLabelMutation() {
  const output = await preSerializationLabelWorkbook.xlsx.writeBuffer();
  const returnedWorkbook = new ExcelJS.Workbook();
  await returnedWorkbook.xlsx.load(output);
  const returnedLabelErrors = validateWorkbookPreservation(returnedWorkbook, returnedLabelSnapshot).filter((error) =>
    /row label changed/i.test(error)
  );
  assert.deepEqual(
    returnedLabelErrors,
    [],
    "the exact generated Reported label must remain authorized after a real ExcelJS buffer write/load round trip"
  );
  returnedWorkbook.getWorksheet("Segment Analysis").getCell("C17").value = "Tampered after reload";
  assert.match(
    validateWorkbookPreservation(returnedWorkbook, returnedLabelSnapshot).join("\n"),
    /C17: row label changed from "Segment 2" to "Tampered after reload"/,
    "the returned buffer must reject a label that differs from the exact authorized mutation"
  );

  const propagatedOutput = await propagatedLabelWorkbook.xlsx.writeBuffer();
  const returnedPropagatedWorkbook = new ExcelJS.Workbook();
  await returnedPropagatedWorkbook.xlsx.load(propagatedOutput);
  assert.equal(
    returnedPropagatedWorkbook.getWorksheet("Segment Analysis").getCell("C70").text,
    "Established Pharmaceutical Products EBITDA"
  );
  assert.equal(
    returnedPropagatedWorkbook.getWorksheet("Revenue Build").getCell("C1").text,
    "Established Pharmaceutical Products D&A"
  );
  assert.equal(
    returnedPropagatedWorkbook.getWorksheet("Revenue Build").getCell("C3").text,
    "Multiple Used for Established Pharmaceutical Products D&A"
  );
  assert.deepEqual(
    validateWorkbookPreservation(returnedPropagatedWorkbook, propagatedLabelSnapshot).filter((error) =>
      /row label changed|formula changed/i.test(error)
    ),
    [],
    "same-sheet and cross-sheet segment-label caches must survive the returned-workbook serialization boundary"
  );

  const metadataOutput = await metadataCacheWorkbook.xlsx.writeBuffer();
  const returnedMetadataWorkbook = new ExcelJS.Workbook();
  await returnedMetadataWorkbook.xlsx.load(metadataOutput);
  assert.equal(returnedMetadataWorkbook.getWorksheet("Summary").getCell("C1").text, "Abbott Laboratories (ABT)");
  assert.equal(returnedMetadataWorkbook.getWorksheet("Summary").getCell("C2").text, "ABT");

  const integratedTemplate = new ExcelJS.Workbook();
  await integratedTemplate.xlsx.readFile(
    path.join(__dirname, "..", "github", "templates", "general", "Owl Fund Integrated Model Template (03-Sep-2025)_v25 (3).xlsx")
  );
  refreshCompanyMetadataFormulaResults(
    integratedTemplate,
    "Lowe's Companies, Inc.",
    "LOW",
    "Abbott Laboratories",
    "ABT"
  );
  const comparables = integratedTemplate.getWorksheet("Comparables");
  const multiplesData = integratedTemplate.getWorksheet("Multiples Data");
  for (const address of ["F10", "AB10", "AX10", "F28", "AB28", "AX28"]) {
    assert.equal(comparables.getCell(address).value, "ABT", `${address} must refresh the hardcoded target ticker`);
  }
  for (const address of ["AX6", "BD6", "BR6"]) {
    assert.equal(multiplesData.getCell(address).value, "Abbott Laboratories", `${address} must refresh the hardcoded target company`);
  }
  for (const address of ["AX7", "BD7", "BR7"]) {
    assert.equal(multiplesData.getCell(address).value, "ABT", `${address} must refresh the hardcoded target ticker`);
  }
  assert.equal(comparables.getCell("F12").value, "AAPL", "unrelated primary peers must remain unchanged");
  assert.equal(comparables.getCell("AB12").value, "RTX", "unrelated defense peers must remain unchanged");
  assert.equal(comparables.getCell("AX12").value, "MMM", "unrelated capital-goods peers must remain unchanged");
  assert.equal(multiplesData.getCell("BM6").text, "3M Company", "unrelated peer metadata must remain unchanged");
}

validateSerializedSegmentLabelMutation()
  .then(() => console.log("Segment formula-write, label-preservation, and annual bridge checks passed."))
  .catch((error) => {
    console.error(error);
    process.exitCode = 1;
  });
