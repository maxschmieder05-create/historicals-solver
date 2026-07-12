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
const {
  primaryBalanceSheetRowsHaveStructuralAnchor,
  selectPrimaryBalanceSheetStatementCandidates,
  buildPrimaryBalanceSheetAssignmentLedgerRows
} = __fillModelServiceTestHooks;
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

const completeRows = [...assetRows, ...liabilityRows.map((row) => ({ ...row, rowOrder: row.rowOrder + 10 }))];
const selectedCompleteStatement = selectPrimaryBalanceSheetStatementCandidates([
  { statement: { ...statement("CONSOLIDATED BALANCE SHEET"), rows: completeRows }, rows: completeRows },
  { statement: { ...statement("CONSOLIDATED BALANCE SHEET ASSETS"), rows: assetRows }, rows: assetRows },
  { statement: { ...statement("CONSOLIDATED BALANCE SHEET LIABILITIES AND EQUITY"), rows: liabilityRows }, rows: liabilityRows }
]);
assert.equal(selectedCompleteStatement.length, 1, "one complete primary balance sheet must win over redundant anchored fragments");
assert.equal(selectedCompleteStatement[0].rows, completeRows);

const accession = "0000000000-26-000001";
const roleUri = "https://issuer.example/role/ConsolidatedBalanceSheet";
function ledgerRow({ rowOrder, label, concept, value, parentConcept, parentLabel, section }) {
  return {
    statementName: "CONSOLIDATED BALANCE SHEET",
    sourceTableType: "primary_statement",
    rowLabel: label,
    xbrlConcept: concept,
    value,
    unit: "USD",
    period: { instant: "2026-03-31", end: "2026-03-31", periodType: "instant" },
    consolidated: true,
    dimensions: [],
    currentNonCurrentSection: section,
    rowOrder,
    accession,
    reportingPeriod: "2026-03-31",
    parentSubtotal: {
      concept: parentConcept,
      label: parentLabel,
      relationship: "calculation",
      weight: 1,
      roleUri
    }
  };
}

const ledgerStatementRows = [
  ledgerRow({ rowOrder: 1, label: "Cash and cash equivalents", concept: "CashAndCashEquivalentsAtCarryingValue", value: 1_000, parentConcept: "AssetsCurrent", parentLabel: "Assets, Current", section: "current" }),
  ledgerRow({ rowOrder: 2, label: "Assets, Current", concept: "AssetsCurrent", value: 1_000, parentConcept: "Assets", parentLabel: "Assets", section: "current" }),
  ledgerRow({ rowOrder: 3, label: "Long-term investments", concept: "LongTermInvestments", value: 918, parentConcept: "Assets", parentLabel: "Assets", section: "non_current" }),
  ledgerRow({ rowOrder: 4, label: "Property, Plant and Equipment, Net", concept: "PropertyPlantAndEquipmentNet", value: 11_816, parentConcept: "Assets", parentLabel: "Assets" }),
  ledgerRow({ rowOrder: 5, label: "Intangible Assets, Net", concept: "IntangibleAssetsNetExcludingGoodwill", value: 5_526, parentConcept: "Assets", parentLabel: "Assets" }),
  ledgerRow({ rowOrder: 6, label: "Deferred Income Taxes and Other Assets, Noncurrent", concept: "DeferredIncomeTaxesAndOtherAssetsNoncurrent", value: 18_422, parentConcept: "Assets", parentLabel: "Assets", section: "non_current" }),
  ledgerRow({ rowOrder: 7, label: "Equity securities", concept: "EquitySecuritiesFvNi", value: 342, parentConcept: "AssetsFairValueDisclosure", parentLabel: "Assets, Fair Value Disclosure", section: "current" }),
  ledgerRow({ rowOrder: 8, label: "Assets", concept: "Assets", value: 37_682, parentConcept: "AssetsAbstract", parentLabel: "Assets [Abstract]" }),
  ledgerRow({ rowOrder: 9, label: "Accounts payable", concept: "AccountsPayableTradeCurrent", value: 2_000, parentConcept: "LiabilitiesCurrent", parentLabel: "Liabilities, Current", section: "current" }),
  ledgerRow({ rowOrder: 10, label: "Liabilities, Current", concept: "LiabilitiesCurrent", value: 2_000, parentConcept: "LiabilitiesAndStockholdersEquity", parentLabel: "Liabilities and Equity", section: "current" }),
  ledgerRow({ rowOrder: 11, label: "Amount of post-employment obligations, deferred tax liability, and other liabilities", concept: "PostEmploymentObligationsDeferredIncomeTaxesAndOtherLongTermLiabilities", value: 7_550, parentConcept: "LiabilitiesAndStockholdersEquity", parentLabel: "Liabilities and Equity" }),
  ledgerRow({ rowOrder: 12, label: "Post-employment obligations and other long-term liabilities", concept: "PostEmploymentObligationsDeferredIncomeTaxesAndOtherLongTermLiabilities", value: 7_550, parentConcept: "LiabilitiesAndStockholdersEquity", parentLabel: "Liabilities and Equity", section: "non_current" }),
  ledgerRow({ rowOrder: 13, label: "Common stock", concept: "CommonStockValue", value: 28_132, parentConcept: "StockholdersEquity", parentLabel: "Stockholders' Equity" }),
  ledgerRow({ rowOrder: 14, label: "Stockholders' Equity", concept: "StockholdersEquity", value: 28_132, parentConcept: "LiabilitiesAndStockholdersEquity", parentLabel: "Liabilities and Equity" }),
  ledgerRow({ rowOrder: 14.5, label: "Stockholders' Equity Attributable to Noncontrolling Interest", concept: "MinorityInterest", value: 222, parentConcept: "StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest", parentLabel: "Stockholders' Equity, Including Portion Attributable to Noncontrolling Interest" }),
  ledgerRow({ rowOrder: 15, label: "Liabilities and Equity", concept: "LiabilitiesAndStockholdersEquity", value: 37_682, parentConcept: "LiabilitiesAndStockholdersEquityAbstract", parentLabel: "Liabilities and Equity [Abstract]" })
];
const ledgerStatement = {
  statementName: "CONSOLIDATED BALANCE SHEET",
  sourceTableType: "primary_statement",
  roleUri,
  accession,
  reportingPeriod: "2026-03-31",
  form: "10-Q",
  filingDate: "2026-04-30",
  rows: ledgerStatementRows
};
const fiscalEntry = {
  accessionNumber: accession,
  accessionKey: accession.replace(/-/g, ""),
  form: "10-Q",
  filingDate: "2026-04-30",
  reportDate: "2026-03-31",
  fiscalYear: 2026,
  fiscalQuarter: 1,
  quarterPeriod: "1Q26"
};
const ledgerContext = {
  duration: new Map(),
  instant: new Map(),
  filingPackageStatements: [ledgerStatement],
  fiscalPeriods: {
    entries: [fiscalEntry],
    byAccession: new Map([[fiscalEntry.accessionKey, fiscalEntry]]),
    byReportDate: new Map([[fiscalEntry.reportDate, fiscalEntry]]),
    reportedPeriods: new Set(["1Q26"]),
    fiscalYearEndMonth: 12,
    fiscalYearEndDay: 31
  }
};
const ledgerFillRows = [
  "Cash & Cash Equivalents",
  "Other Non-Current Assets",
  "PP&E, Net",
  "Intangible Assets, Net",
  "Accounts Payable",
  "Other Non-Current Liabilities",
  "Common Stock & APIC",
  "Noncontrolling Interests"
].map((label, index) => ({ row: index + 1, label, classification: "direct", statement: "balance", kind: "instant", scale: 1_000_000 }));
const assignmentLedger = buildPrimaryBalanceSheetAssignmentLedgerRows(["1Q26"], ledgerContext, ledgerFillRows);
assert.equal(
  assignmentLedger.filter((row) => row.sourceXbrlTag === "PostEmploymentObligationsDeferredIncomeTaxesAndOtherLongTermLiabilities").length,
  1,
  "the same SEC fact identity must not be counted twice merely because presentation labels or inferred sections differ"
);
assert.equal(
  assignmentLedger.some((row) => row.sourceXbrlTag === "DeferredIncomeTaxesAndOtherAssetsNoncurrent"),
  true,
  "an 'and other' primary carrying amount must remain when unrelated sibling assets happen to approximate it"
);
assert.equal(
  assignmentLedger.some((row) => row.sourceXbrlTag === "EquitySecuritiesFvNi"),
  false,
  "fair-value disclosure composition facts must not be counted as primary balance-sheet carrying rows"
);
assert.equal(
  assignmentLedger.some((row) => row.sourceXbrlTag === "MinorityInterest" && row.assignedModelRow === "Noncontrolling Interests"),
  true,
  "equity explicitly attributable to noncontrolling interests is a component carrying amount, not an equity subtotal"
);

console.log("Primary balance-sheet structure regression passed.");
