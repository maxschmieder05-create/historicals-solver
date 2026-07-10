const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const Module = require("node:module");
const ts = require("typescript");

const repoRoot = path.resolve(__dirname, "..");

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

const { isPrimaryBalanceSheetStructure, __fillModelServiceTestHooks } = loadTypeScriptModule(
  path.join(repoRoot, "server", "fill-model", "fill-model-service.ts")
);
const { primaryBalanceSheetRowsHaveStructuralAnchor, selectPrimaryBalanceSheetStatementCandidates } = __fillModelServiceTestHooks;
const { classifySourceTableType } = loadTypeScriptModule(path.join(repoRoot, "server", "fill-model", "sec-filing-package.ts"));

function statement(statementName) {
  return {
    statementName,
    sourceTableType: "primary_statement",
    accession: "0000000000-26-000001",
    rows: []
  };
}

assert.equal(
  isPrimaryBalanceSheetStructure(statement("CONSOLIDATED BALANCE SHEET - (CONTINUED) (UNAUDITED) LIABILITIES AND EQUITY")),
  true
);
assert.equal(isPrimaryBalanceSheetStructure(statement("CONSOLIDATED BALANCE SHEET ASSETS")), true);
assert.equal(isPrimaryBalanceSheetStructure(statement("CONSOLIDATED STATEMENT OF STOCKHOLDERS' EQUITY")), false);
assert.equal(isPrimaryBalanceSheetStructure(statement("CONSOLIDATED STATEMENTS OF CASH FLOWS")), false);
assert.equal(isPrimaryBalanceSheetStructure(statement("BALANCE SHEET PARENTHETICAL")), false);
assert.equal(
  isPrimaryBalanceSheetStructure(statement("Consolidated Balance Sheet at March 31, 2026 and December 31, 2025 were not material")),
  false
);
assert.equal(
  classifySourceTableType("Consolidated Balance Sheet at March 31, 2026 and December 31, 2025 were not material"),
  "support_table"
);
assert.equal(
  classifySourceTableType("Current-period accruals Accrual adjustments Charges incurred Balance at March 31 Product warranty accrual"),
  "roll_forward"
);
assert.equal(
  classifySourceTableType("Consolidated Balance Sheet is estimated to be the following Remainder of 2026 Thereafter finite-lived intangible assets amortization expense"),
  "support_table"
);

function balanceRow(xbrlConcept, rowLabel, value, rowOrder) {
  return {
    statementName: "CONSOLIDATED BALANCE SHEET",
    sourceTableType: "primary_statement",
    rowLabel,
    xbrlConcept,
    value,
    unit: "USD",
    period: { instant: "2026-03-31", periodType: "instant" },
    consolidated: true,
    dimensions: [],
    rowOrder,
    accession: "0000000000-26-000001"
  };
}

const assetRows = [
  balanceRow("CashAndCashEquivalentsAtCarryingValue", "Cash and cash equivalents", 100, 1),
  balanceRow("AssetsCurrent", "Total current assets", 300, 2),
  balanceRow("Assets", "Total assets", 900, 3)
];
const liabilityRows = [
  balanceRow("LiabilitiesCurrent", "Total current liabilities", 250, 1),
  balanceRow("StockholdersEquity", "Total stockholders' equity", 400, 2),
  balanceRow("LiabilitiesAndStockholdersEquity", "Total liabilities and stockholders' equity", 900, 3)
];
const orphanedDebtDetailRows = [
  balanceRow("DeferredFinanceCostsNet", "Less unamortized debt discounts and issuance costs", 14, 1)
];

assert.equal(primaryBalanceSheetRowsHaveStructuralAnchor(assetRows), true);
assert.equal(primaryBalanceSheetRowsHaveStructuralAnchor(liabilityRows), true);
assert.equal(primaryBalanceSheetRowsHaveStructuralAnchor(orphanedDebtDetailRows), false);

const selectedSplitStatements = selectPrimaryBalanceSheetStatementCandidates([
  { statement: { ...statement("CONSOLIDATED BALANCE SHEET ASSETS"), rows: assetRows }, rows: assetRows },
  { statement: { ...statement("CONSOLIDATED BALANCE SHEET LIABILITIES AND EQUITY"), rows: liabilityRows }, rows: liabilityRows },
  { statement: { ...statement("CONSOLIDATED BALANCE SHEET"), rows: orphanedDebtDetailRows }, rows: orphanedDebtDetailRows }
]);
assert.equal(selectedSplitStatements.length, 2);
assert.equal(selectedSplitStatements.some((candidate) => candidate.rows === orphanedDebtDetailRows), false);

const fallbackFragment = selectPrimaryBalanceSheetStatementCandidates([
  { statement: { ...statement("CONSOLIDATED BALANCE SHEET"), rows: orphanedDebtDetailRows }, rows: orphanedDebtDetailRows }
]);
assert.equal(fallbackFragment.length, 1);

console.log("Primary balance-sheet structure regression passed.");
