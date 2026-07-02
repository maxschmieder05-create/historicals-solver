const assert = require("node:assert/strict");
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

console.log("AMD balance sheet liability classification regression passed.");
