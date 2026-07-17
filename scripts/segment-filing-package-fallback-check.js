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

function statement(period, start, end, value, form = "10-Q", options = {}) {
  return {
    statementName: options.statementName || "Reportable Segment Information (Details)",
    roleDefinition: options.roleDefinition || "Disclosure - Segment Reporting - Reportable Segment Information (Details)",
    roleUri: options.roleUri || "https://issuer.example/role/SegmentDetails",
    reportCategory: options.reportCategory || "Details",
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
            dimension: options.dimension || "StatementBusinessSegmentsAxis",
            dimensionLabel: options.dimensionLabel,
            member: options.member || "OpaqueSegmentMember",
            memberLabel: options.memberLabel || "Cloud"
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

const detailQ1 = statement("001", "2026-01-01", "2026-03-31", 70);
detailQ1.rows.push({
  ...detailQ1.rows[0],
  rowLabel: "Gross Revenue",
  xbrlConcept: "Revenues",
  value: 999,
  rowOrder: 2
});

const fallback = segmentRevenueFromFilingPackageStatements(
  [
    statement("001-summary", "2026-01-01", "2026-03-31", 999, "10-Q", {
      statementName: "Segment Reporting (Tables)",
      roleDefinition: "Disclosure - Segment Reporting (Tables)",
      roleUri: "https://issuer.example/role/SegmentTables",
      reportCategory: "Tables"
    }),
    detailQ1,
    statement("001-geographic", "2026-01-01", "2026-03-31", 8, "10-Q", {
      statementName: "Segment Reporting - Narrative (Details)",
      roleDefinition: "Disclosure - Segment Reporting - Narrative (Details)",
      dimension: "StatementGeographicalAxis",
      dimensionLabel: "Geographical areas [Axis]",
      member: "CanadaMember",
      memberLabel: "Canada"
    }),
    statement("002", "2026-04-01", "2026-06-30", 90),
    statement("003", "2026-07-01", "2026-09-30", 110),
    statement("004", "2026-01-01", "2026-12-31", 400, "10-K")
  ],
  ["1Q26", "2Q26", "3Q26", "4Q26", "FY26"],
  ctx
);

assert.equal(fallback.length, 1);
assert.equal(fallback[0].label, "Cloud");
assert.equal(fallback[0].values.get("1Q26"), 70);
assert.notEqual(fallback[0].values.get("1Q26"), 999, "detail-role values must take precedence over high-level segment table duplicates");
assert.equal(
  fallback[0].revenueSources.get("1Q26")[0].concept,
  "RevenueFromContractWithCustomerExcludingAssessedTax",
  "net/external revenue must take precedence over a gross-revenue fact in the same segment-detail role"
);
assert.equal(fallback[0].values.get("4Q26"), 130, "segment Q4 must equal FY less Q1-Q3 for flow metrics");
assert.equal(fallback[0].values.get("FY26"), 400, "annual segment facts must remain available to annual workbook columns");
assert.equal(fallback[0].revenueSources.get("4Q26")[0].sourceLayer, "derived");

const numberedGroup = segmentRevenueFromFilingPackageStatements(
  [statement("group-1", "2026-01-01", "2026-03-31", 55, "10-Q", { member: "Group1SegmentMember", memberLabel: "Group 1" })],
  ["1Q26"],
  ctx
);
assert.equal(numberedGroup[0].label, "Group 1", "an official numbered member in an authoritative segment-detail role must remain usable");
assert.equal(numberedGroup[0].values.get("1Q26"), 55);

const renamedMember = segmentRevenueFromFilingPackageStatements(
  [
    statement("stable-member-q1", "2026-01-01", "2026-03-31", 55, "10-Q", {
      member: "CloudSegmentMember",
      memberLabel: "Cloud"
    }),
    statement("stable-member-q2", "2026-04-01", "2026-06-30", 65, "10-Q", {
      member: "CloudSegmentMember",
      memberLabel: "Cloud Platform"
    })
  ],
  ["1Q26", "2Q26"],
  ctx
);
assert.equal(renamedMember.length, 1, "the stable XBRL member QName must remain the segment identity when its display label changes");
assert.equal(renamedMember[0].values.get("1Q26"), 55);
assert.equal(renamedMember[0].values.get("2Q26"), 65);

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
