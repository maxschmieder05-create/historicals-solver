const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const ts = require("typescript");

const repoRoot = path.resolve(__dirname, "..");

require.extensions[".ts"] = function compileTypeScriptModule(module, filename) {
  const source = fs.readFileSync(filename, "utf8");
  const compiled = ts.transpileModule(source, {
    compilerOptions: {
      module: ts.ModuleKind.CommonJS,
      target: ts.ScriptTarget.ES2022,
      esModuleInterop: true
    }
  }).outputText;
  module._compile(compiled, filename);
};

const { __fillModelServiceTestHooks } = require(
  path.join(repoRoot, "server", "fill-model", "fill-model-service.ts")
);
const {
  segmentRevenueFromFilingPackageStatements,
  mergeSegmentRevenueFallback
} = __fillModelServiceTestHooks;

function statement(period, start, end, value, form = "10-Q") {
  return {
    statementName: "Reportable Segment Information",
    sourceTableType: "segment_table",
    accession: `0000000001-26-${period}`,
    reportingPeriod: end,
    form,
    filingDate: end,
    sourceUrl: "https://www.sec.gov/Archives/example",
    rows: [
      {
        statementName: "Reportable Segment Information",
        sourceTableType: "segment_table",
        rowLabel: "Revenue from external customers",
        xbrlConcept: "RevenueFromContractWithCustomerExcludingAssessedTax",
        taxonomy: "us-gaap",
        value,
        unit: "USD",
        period: { start, end, periodType: "duration" },
        consolidated: false,
        dimensions: [
          {
            dimension: "StatementBusinessSegmentsAxis",
            member: "CloudSegmentMember"
          }
        ],
        rowOrder: 1,
        accession: `0000000001-26-${period}`,
        reportingPeriod: end,
        sourceUrl: "https://www.sec.gov/Archives/example"
      }
    ]
  };
}

const fiscalEntries = [
  ["2026-03-31", "1Q26"],
  ["2026-06-30", "2Q26"],
  ["2026-09-30", "3Q26"],
  ["2026-12-31", "4Q26"]
].map(([reportDate, quarterPeriod], index) => ({
  accessionNumber: String(index + 1),
  accessionKey: String(index + 1),
  form: index === 3 ? "10-K" : "10-Q",
  filingDate: reportDate,
  reportDate,
  fiscalYear: 2026,
  fiscalQuarter: index + 1,
  quarterPeriod,
  annualPeriod: index === 3 ? "FY26" : undefined
}));
const ctx = {
  duration: new Map(),
  instant: new Map(),
  fiscalPeriods: {
    entries: fiscalEntries,
    byAccession: new Map(),
    byReportDate: new Map(fiscalEntries.map((entry) => [entry.reportDate, entry])),
    reportedPeriods: new Set(["1Q26", "2Q26", "3Q26", "4Q26", "FY26"]),
    fiscalYearEndMonth: 12,
    fiscalYearEndDay: 31
  }
};

const fallback = segmentRevenueFromFilingPackageStatements(
  [
    statement("001", "2026-01-01", "2026-03-31", 70),
    statement("002", "2026-04-01", "2026-06-30", 90),
    statement("003", "2026-07-01", "2026-09-30", 110),
    statement("004", "2026-01-01", "2026-12-31", 400, "10-K")
  ],
  ["1Q26", "2Q26", "3Q26", "4Q26"],
  ctx
);

assert.equal(fallback.length, 1);
assert.equal(fallback[0].label, "Cloud");
assert.equal(fallback[0].values.get("1Q26"), 70);
assert.equal(fallback[0].values.get("4Q26"), 130, "segment Q4 must equal FY less Q1-Q3 for flow metrics");
assert.equal(fallback[0].revenueSources.get("4Q26")[0].sourceLayer, "derived");

const primary = [
  {
    label: "Cloud",
    values: new Map([["1Q26", 75]]),
    annualValues: new Map(),
    operatingIncome: new Map(),
    depreciationAmortization: new Map(),
    revenueSources: new Map([["1Q26", [{ concept: "InlineRevenue", label: "Cloud revenue", value: 75 }]]]),
    operatingIncomeSources: new Map(),
    depreciationAmortizationSources: new Map()
  }
];
const merged = mergeSegmentRevenueFallback(primary, fallback, ["1Q26", "2Q26", "3Q26", "4Q26"]);
assert.equal(merged[0].values.get("1Q26"), 75, "existing inline segment facts must remain authoritative");
assert.equal(merged[0].values.get("2Q26"), 90, "filing-package facts should fill only missing segment periods");
assert.equal(merged[0].values.get("4Q26"), 130);

console.log("SEC filing-package segment fallback checks passed.");
