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
  {
    statementName: "Condensed Consolidated Balance Sheets",
    sourceTableType: "primary_statement",
    rowLabel: "Accounts receivable (net of allowance of $12)",
    xbrlConcept: "AllowanceForDoubtfulAccountsReceivableCurrent",
    taxonomy: "us-gaap",
    value: 12_000_000,
    unit: "USD",
    period: {
      instant: reportDate,
      periodType: "instant"
    },
    consolidated: true,
    dimensions: [],
    rowOrder: 0,
    accession,
    reportingPeriod: reportDate,
    currentNonCurrentSection: "current",
    parentSubtotal: { label: "Current assets", concept: "AssetsCurrent" }
  },
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
assert.equal(byLabel.has("Accounts receivable (net of allowance of $12)"), false);

const accrued = hooks.resolveAccruedLiabilities("1Q26", ctx);
assert.equal(accrued.value, 5_785_000_000);

const completeNoDebtRows = [
  balanceRow(1, "Total assets", "Assets", 100_000_000, "non_current"),
  balanceRow(2, "Cash and cash equivalents", "CashAndCashEquivalentsAtCarryingValue", 20_000_000, "current"),
  balanceRow(3, "Accounts receivable", "AccountsReceivableNetCurrent", 10_000_000, "current"),
  balanceRow(4, "Property and equipment, net", "PropertyPlantAndEquipmentNet", 70_000_000, "non_current"),
  balanceRow(5, "Total liabilities", "Liabilities", 40_000_000, "non_current"),
  balanceRow(6, "Accounts payable", "AccountsPayableCurrent", 10_000_000, "current"),
  balanceRow(7, "Other liabilities", "OtherLiabilitiesCurrent", 30_000_000, "current"),
  balanceRow(8, "Stockholders' equity", "StockholdersEquity", 60_000_000, "equity"),
  balanceRow(9, "Total liabilities and stockholders' equity", "LiabilitiesAndStockholdersEquity", 100_000_000, "equity")
];
const completeNoDebtCtx = {
  ...ctx,
  instant: new Map([
    [
      "1Q26",
      new Map(
        completeNoDebtRows.map((row) => [
          row.xbrlConcept,
          {
            concept: row.xbrlConcept,
            label: row.rowLabel,
            value: row.value,
            unit: "USD",
            sourceLayer: "sec_filing_package",
            accn: accession,
            end: reportDate,
            periodKey: "1Q26",
            periodType: "instant"
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
      rows: completeNoDebtRows
    }
  ]
};
for (const [resolved, conceptPattern] of [
  [hooks.resolveCurrentDebt("1Q26", completeNoDebtCtx), /DebtDerivedZeroFromCompleteBalanceSheet/],
  [hooks.resolveRevolverCurrentDebt("1Q26", completeNoDebtCtx), /ShortTermBorrowingsDerivedZeroFromCompleteBalanceSheet/],
  [hooks.resolveLongTermDebtInclCurrentPortion("1Q26", completeNoDebtCtx), /DebtDerivedZeroFromCompleteBalanceSheet/],
  [hooks.resolveDeferredTaxLiability("1Q26", completeNoDebtCtx), /DeferredTaxLiabilityDerivedZeroFromCompleteBalanceSheet/]
]) {
  assert.equal(resolved.value, 0, "A complete, reconciled primary SEC balance sheet with no debt line should derive debt as zero.");
  assert.equal(resolved.sources[0].sourceLayer, "derived");
  assert.match(resolved.sources[0].concept, conceptPattern);
  assert.ok(resolved.sources.slice(1).some((source) => source.accn === accession), "The zero-debt proof must retain its SEC filing anchor.");
}
const incompleteNoDebtCtx = {
  ...completeNoDebtCtx,
  filingPackageStatements: [
    {
      ...completeNoDebtCtx.filingPackageStatements[0],
      rows: completeNoDebtRows.filter((row) => row.xbrlConcept !== "LiabilitiesAndStockholdersEquity" && row.xbrlConcept !== "StockholdersEquity")
    }
  ]
};
assert.equal(
  hooks.resolveCurrentDebt("1Q26", incompleteNoDebtCtx).value,
  null,
  "An incomplete balance sheet cannot certify debt absence as zero."
);
assert.equal(accrued.sources.some((source) => source.concept === "LongTermDebtCurrent"), false);

const otherCurrent = hooks.resolveOtherCurrentLiabilities("1Q26", ctx);
assert.equal(otherCurrent.value, 850_000_000);
assert.equal(otherCurrent.sources.some((source) => source.concept === "AccruedLiabilitiesCurrentAndNoncurrent"), false);

function ibmLikeRow(rowOrder, rowLabel, xbrlConcept, value, section = "current", statementName = "Consolidated Balance Sheet") {
  return {
    statementName,
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
    parentSubtotal: { label: section === "current" ? "Current assets" : "Assets", concept: section === "current" ? "AssetsCurrent" : "Assets" }
  };
}

const ibmLikeRows = [
  ibmLikeRow(1, "Finished goods", "InventoryFinishedGoodsNetOfReserves", 268_000_000),
  ibmLikeRow(2, "Work in process and raw materials", "InventoryWorkInProcessAndRawMaterialsNetOfReserves", 1_208_000_000),
  ibmLikeRow(3, "Total inventory", "InventoryNet", 1_476_000_000),
  ibmLikeRow(4, "Intangible assets — net", "FiniteLivedIntangibleAssetsNet", 14_624_000_000, "non_current"),
  ibmLikeRow(
    5,
    "Remainder of 2026",
    "FiniteLivedIntangibleAssetsAmortizationExpenseRemainderOfFiscalYear",
    1_500_000_000,
    "non_current",
    "Consolidated Balance Sheet is estimated to be the following"
  ),
  ibmLikeRow(
    6,
    "Thereafter",
    "FiniteLivedIntangibleAssetsAmortizationExpenseYearFiveAndAfterYearFive",
    13_124_000_000,
    "non_current",
    "Consolidated Balance Sheet is estimated to be the following"
  )
];
const ibmLikeCtx = {
  ...ctx,
  instant: new Map([
    [
      "1Q26",
      new Map(
        ibmLikeRows.map((row) => [
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
      statementName: "Consolidated Balance Sheet",
      sourceTableType: "primary_statement",
      accession,
      reportingPeriod: reportDate,
      form: "10-Q",
      filingDate: "2026-05-06",
      rows: ibmLikeRows
    }
  ]
};
const ibmLikeFillRows = ["Inventory", "Intangible Assets, Net"].map((label, index) => ({
  row: index + 10,
  label,
  classification: "direct",
  statement: "balance",
  kind: "instant",
  scale: 1_000_000
}));
const inventory = hooks.resolveInventory("1Q26", ibmLikeCtx);
assert.equal(inventory.value, 1_476_000_000);
assert.equal(inventory.sources[0].concept, "InventoryNet");
const ibmLikeLedger = hooks.buildPrimaryBalanceSheetAssignmentLedgerRows(["1Q26"], ibmLikeCtx, ibmLikeFillRows);
assert.equal(
  ibmLikeLedger.filter((row) => row.assignedModelRow === "Inventory").reduce((total, row) => total + row.amount, 0),
  1_476_000_000
);
assert.equal(
  ibmLikeLedger.filter((row) => row.assignedModelRow === "Intangible Assets, Net").reduce((total, row) => total + row.amount, 0),
  14_624_000_000
);
assert.equal(ibmLikeLedger.some((row) => /Remainder|Thereafter/.test(row.sourceLineItemLabel)), false);

const inventoryAndPpeHierarchyRows = [
  {
    ...ibmLikeRow(1, "Inventory, Finished Goods, Net of Reserves", "InventoryFinishedGoodsNetOfReserves", 500_000_000),
    parentSubtotal: { label: "Inventory, Net", concept: "InventoryNet", relationship: "calculation", weight: 1 }
  },
  {
    ...ibmLikeRow(2, "Inventory, Work in Process, Net of Reserves", "InventoryWorkInProcessNetOfReserves", 200_000_000),
    parentSubtotal: { label: "Inventory, Net", concept: "InventoryNet", relationship: "calculation", weight: 1 }
  },
  {
    ...ibmLikeRow(3, "Inventory, Raw Materials, Net of Reserves", "InventoryRawMaterialsNetOfReserves", 300_000_000),
    parentSubtotal: { label: "Inventory, Net", concept: "InventoryNet", relationship: "calculation", weight: 1 }
  },
  ibmLikeRow(4, "Inventory, Net", "InventoryNet", 1_000_000_000),
  ibmLikeRow(5, "Property, Plant and Equipment, Net", "PropertyPlantAndEquipmentNet", 2_000_000_000, "non_current"),
  {
    ...ibmLikeRow(6, "Land", "Land", 100_000_000, "non_current"),
    currentNonCurrentSection: undefined,
    parentSubtotal: undefined
  },
  {
    ...ibmLikeRow(
      7,
      "Amount of post-employment obligations and deferred tax liability, after deferred tax asset, and other liabilities expected to be paid after one year",
      "PostEmploymentObligationsDeferredIncomeTaxesAndOtherLongTermLiabilities",
      700_000_000,
      "non_current"
    ),
    parentSubtotal: { label: "Assets", concept: "Assets", relationship: "calculation" }
  },
  {
    ...ibmLikeRow(8, "Deferred Income Taxes and Other Assets, Noncurrent", "DeferredIncomeTaxesAndOtherAssetsNoncurrent", 400_000_000, "non_current"),
    parentSubtotal: { label: "Assets", concept: "Assets", relationship: "calculation" }
  }
];
const inventoryAndPpeHierarchyCtx = {
  ...ctx,
  instant: new Map([
    [
      "1Q26",
      new Map(
        inventoryAndPpeHierarchyRows.map((row) => [
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
      statementName: "Consolidated Balance Sheet",
      sourceTableType: "primary_statement",
      accession,
      reportingPeriod: reportDate,
      form: "10-Q",
      filingDate: "2026-05-06",
      rows: inventoryAndPpeHierarchyRows
    }
  ]
};
const inventoryAndPpeHierarchyFillRows = ["Inventory", "PP&E, Net", "Other Non-Current Assets", "Other Non-Current Liabilities"].map((label, index) => ({
  row: index + 10,
  label,
  classification: "direct",
  statement: "balance",
  kind: "instant",
  scale: 1_000_000
}));
const inventoryAndPpeHierarchyLedger = hooks.buildPrimaryBalanceSheetAssignmentLedgerRows(
  ["1Q26"],
  inventoryAndPpeHierarchyCtx,
  inventoryAndPpeHierarchyFillRows
);
assert.equal(
  inventoryAndPpeHierarchyLedger.filter((row) => row.assignedModelRow === "Inventory").reduce((total, row) => total + row.amount, 0),
  1_000_000_000,
  "A reported inventory carrying total and its raw/WIP/finished-goods detail must enter the assignment ledger exactly once."
);
assert.equal(
  inventoryAndPpeHierarchyLedger.some((row) => row.sourceXbrlTag === "InventoryNet"),
  false,
  "The inventory aggregate should be excluded when primary-statement component detail reconciles to it."
);
const standaloneLandAssignment = inventoryAndPpeHierarchyLedger.find((row) => row.sourceXbrlTag === "Land");
assert.equal(standaloneLandAssignment.assignmentStatus, "explicitly_excluded_with_reason");
assert.match(standaloneLandAssignment.classificationReason, /PP&E gross component detail/);
const compoundLongTermLiabilityAssignment = inventoryAndPpeHierarchyLedger.find(
  (row) => row.sourceXbrlTag === "PostEmploymentObligationsDeferredIncomeTaxesAndOtherLongTermLiabilities"
);
assert.equal(compoundLongTermLiabilityAssignment.assignedModelRow, "Other Non-Current Liabilities");
assert.equal(compoundLongTermLiabilityAssignment.side, "liabilities_and_equity");
const deferredTaxAndOtherAssetAssignment = inventoryAndPpeHierarchyLedger.find(
  (row) => row.sourceXbrlTag === "DeferredIncomeTaxesAndOtherAssetsNoncurrent"
);
assert.equal(deferredTaxAndOtherAssetAssignment.assignedModelRow, "Other Non-Current Assets");
assert.equal(deferredTaxAndOtherAssetAssignment.side, "assets");
assert.equal(
  inventoryAndPpeHierarchyLedger
    .filter((row) => row.side === "assets" && row.assignmentStatus !== "explicitly_excluded_with_reason")
    .reduce((total, row) => total + row.amount, 0),
  3_400_000_000,
  "Compound post-employment/deferred-tax liabilities and covered PP&E detail must not inflate the assignment-ledger asset total."
);

const treasuryPrimaryRow = {
  ...balanceRow(1, "Treasury Stock, Common, Value", "TreasuryStockCommonValue", 100_000_000, "equity"),
  parentSubtotal: { label: "Stockholders' equity", concept: "StockholdersEquity", relationship: "calculation" }
};
const treasuryCtx = {
  ...ctx,
  instant: new Map([
    [
      "1Q26",
      new Map([
        [
          "TreasuryStockCommonValue",
          {
            concept: "TreasuryStockCommonValue",
            label: "Treasury Stock, Common, Value",
            value: 100_000_000,
            unit: "USD",
            sourceLayer: "sec_filing_package",
            accn: accession,
            end: reportDate,
            periodKey: "1Q26",
            periodType: "instant"
          }
        ],
        [
          "CommonStockSharesHeldInEmployeeTrust",
          {
            concept: "CommonStockSharesHeldInEmployeeTrust",
            label: "Common Stock, Shares Held in Employee Trust",
            value: 10_000_000,
            unit: "USD",
            sourceLayer: "sec_live_companyfacts",
            accn: accession,
            end: reportDate,
            periodKey: "1Q26",
            periodType: "instant"
          }
        ]
      ])
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
      rows: [treasuryPrimaryRow]
    }
  ]
};
assert.equal(hooks.resolveTreasuryStockOnly("1Q26", treasuryCtx).value, -110_000_000);
const treasuryWorkbook = new ExcelJS.Workbook();
const treasurySheet = treasuryWorkbook.addWorksheet("Model");
treasurySheet.getCell("A1").value = "Balance Sheet";
treasurySheet.getCell("A2").value = "Treasury Stock";
treasurySheet.getCell("C2").value = -110;
treasurySheet.getCell("A3").value = "Cash Flow Statement";
const treasuryFillRows = [
  {
    row: 2,
    label: "Treasury Stock",
    classification: "direct",
    statement: "balance",
    kind: "instant",
    scale: 1_000_000,
    resolver: hooks.resolveTreasuryStockOnly
  }
];
const treasuryWarnings = [];
const treasuryErrors = hooks.validatePrimaryBalanceSheetAssignmentCoverage(
  treasurySheet,
  ["1Q26"],
  [3],
  treasuryCtx,
  new hooks.FormulaEvaluator(treasurySheet, { useCachedFormulaResults: false, allowCachedFormulaResultFallback: false }),
  treasuryWarnings,
  treasuryFillRows
);
assert.equal(treasuryErrors.length, 0, treasuryErrors.join("\n"));
assert.ok(
  treasuryWarnings.some((warning) => /employee-trust contra-equity support/.test(warning)),
  "A treasury row that ties current SEC treasury plus employee-trust contra-equity should pass with a narrow-primary-ledger advisory."
);

assert.equal(
  hooks.reportedLineItemCategory({
    concept: "AccountsReceivableNetCurrent",
    label: "Accounts Receivable, after Allowance for Credit Loss, Current",
    value: 6_493_000_000,
    sourceLayer: "sec_live_companyfacts",
    periodType: "instant"
  }),
  "current_assets"
);
assert.equal(
  hooks.reportedLineItemCategory({
    concept: "PropertyPlantAndEquipmentAndFinanceLeaseRightOfUseAssetAfterAccumulatedDepreciationAndAmortization",
    label: "Property, Plant, and Equipment and Finance Lease Right-of-Use Asset, after Accumulated Depreciation and Amortization",
    value: 5_781_000_000,
    sourceLayer: "sec_live_companyfacts",
    periodType: "instant"
  }),
  "non_current_assets"
);
assert.equal(
  hooks.reportedLineItemCategory({
    concept: "DebtSecuritiesAvailableForSaleExcludingAccruedInterestCurrent",
    label: "Marketable securities",
    value: 964_000_000,
    sourceLayer: "sec_filing_package",
    periodType: "instant"
  }),
  "current_assets"
);
assert.equal(
  hooks.reportedLineItemCategory({
    concept: "AllowanceForDoubtfulAccountsReceivableCurrent",
    label: "Allowance for doubtful accounts receivable, current",
    value: 12_000_000,
    sourceLayer: "sec_filing_package",
    periodType: "instant"
  }),
  "cash_flow_or_support"
);

const tslaAccession = "000162828026000001";
const tslaReportDate = "2026-03-31";
const tslaFilingEntry = {
  accessionNumber: tslaAccession,
  accessionKey: tslaAccession,
  form: "10-Q",
  filingDate: "2026-04-23",
  reportDate: tslaReportDate,
  fiscalYear: 2026,
  fiscalQuarter: 1,
  quarterPeriod: "1Q26"
};

function tslaBalanceRow(rowOrder, rowLabel, xbrlConcept, value, section, parentLabel, parentConcept) {
  return {
    statementName: "Condensed Consolidated Balance Sheets",
    sourceTableType: "primary_statement",
    rowLabel,
    xbrlConcept,
    taxonomy: "us-gaap",
    value,
    unit: "USD",
    period: {
      instant: tslaReportDate,
      periodType: "instant"
    },
    consolidated: true,
    dimensions: [],
    rowOrder,
    accession: tslaAccession,
    reportingPeriod: tslaReportDate,
    currentNonCurrentSection: section,
    parentSubtotal: { label: parentLabel, concept: parentConcept }
  };
}

const tslaRows = [
  tslaBalanceRow(1, "Cash and cash equivalents", "CashAndCashEquivalentsAtCarryingValue", 16_603_000_000, "current", "Current assets", "AssetsCurrent"),
  tslaBalanceRow(2, "Accounts receivable, net", "AccountsReceivableNetCurrent", 3_959_000_000, "current", "Current assets", "AssetsCurrent"),
  tslaBalanceRow(3, "Total assets", "Assets", 143_724_000_000, "total", "Assets", "Assets"),
  tslaBalanceRow(4, "Accounts payable", "AccountsPayableCurrent", 14_696_000_000, "current", "Current liabilities", "LiabilitiesCurrent"),
  tslaBalanceRow(5, "Accrued liabilities and other", "AccruedAndOtherCurrentLiabilities", 14_554_000_000, "current", "Current liabilities", "LiabilitiesCurrent"),
  tslaBalanceRow(6, "Current portion of debt and finance leases", "DebtCurrent", 1_374_000_000, "current", "Current liabilities", "LiabilitiesCurrent"),
  tslaBalanceRow(7, "Operating lease liability, current", "OperatingLeaseLiabilityCurrent", 988_000_000, "current", "Current liabilities", "LiabilitiesCurrent"),
  tslaBalanceRow(8, "Total current liabilities", "LiabilitiesCurrent", 34_138_000_000, "current", "Current liabilities", "LiabilitiesCurrent"),
  tslaBalanceRow(9, "Digital assets", "CryptoAssetFairValueNoncurrent", 786_000_000, "non_current", "Assets", "Assets"),
  tslaBalanceRow(10, "Deferred revenue, net of current portion", "ContractWithCustomerLiabilityNoncurrent", 3_847_000_000, "non_current", "Liabilities", "Liabilities"),
  tslaBalanceRow(11, "Other long-term liabilities", "OtherLiabilitiesNoncurrent", 13_155_000_000, "non_current", "Liabilities", "Liabilities"),
  tslaBalanceRow(12, "Total liabilities", "Liabilities", 58_922_000_000, "total", "Liabilities", "Liabilities"),
  tslaBalanceRow(13, "Redeemable noncontrolling interests in subsidiaries", "RedeemableNoncontrollingInterestEquityCarryingAmount", 57_000_000, "equity", "Equity", "StockholdersEquity"),
  tslaBalanceRow(14, "Noncontrolling interests in subsidiaries", "MinorityInterest", 629_000_000, "equity", "Equity", "StockholdersEquity")
];

const tslaInstantFacts = new Map(
  tslaRows.map((row) => [
    row.xbrlConcept,
    {
      concept: row.xbrlConcept,
      label: row.rowLabel,
      value: row.value,
      unit: "USD",
      taxonomy: row.taxonomy,
      sourceLayer: "sec_filing_package",
      accn: tslaAccession,
      end: tslaReportDate,
      periodKey: "1Q26",
      periodType: "instant",
      reportDate: tslaReportDate
    }
  ])
);

const templateWithoutMezzanine = {
  hasCashRow: true,
  hasCashAndCurrentInvestmentRow: false,
  hasCurrentInvestmentRow: true,
  hasCurrentDebtRow: true,
  hasCurrentDebtMaturitiesRow: true,
  hasShortTermBorrowingsRow: false,
  hasCurrentLiabilitiesExcludingDebtRow: true,
  hasOtherCurrentLiabilityRow: true,
  hasDebtInclCurrentPortionRow: true,
  hasDeferredTaxLiabilityRow: true,
  hasNonCurrentLeaseLiabilityRow: false,
  hasPensionLiabilityRow: false,
  hasMezzanineEquityRow: false
};

const tslaCtx = {
  duration: new Map(),
  instant: new Map([
    ["FY25", new Map([["Goodwill", { concept: "Goodwill", label: "Goodwill", value: 257_000_000, sourceLayer: "sec_filing_package", periodKey: "FY25", periodType: "instant" }]])],
    ["1Q26", tslaInstantFacts]
  ]),
  filingPackageStatements: [
    {
      statementName: "Condensed Consolidated Balance Sheets",
      sourceTableType: "primary_statement",
      accession: tslaAccession,
      reportingPeriod: tslaReportDate,
      form: "10-Q",
      filingDate: "2026-04-23",
      rows: tslaRows
    }
  ],
  fiscalPeriods: {
    entries: [tslaFilingEntry],
    byAccession: new Map([[tslaAccession, tslaFilingEntry]]),
    byReportDate: new Map([[tslaReportDate, tslaFilingEntry]]),
    reportedPeriods: new Set(["1Q26"]),
    fiscalYearEndMonth: 12,
    fiscalYearEndDay: 31
  },
  template: templateWithoutMezzanine
};

const tslaFillRows = [
  "Accounts Payable",
  "Accrued Liabilities",
  "Other Current Liabilities",
  "Other Non-Current Assets",
  "Other Non-Current Liabilities",
  "Noncontrolling Interests"
].map((label, index) => ({
  row: index + 20,
  label,
  classification: "direct",
  statement: "balance",
  kind: "instant",
  scale: 1_000_000
}));

const tslaLedger = hooks.buildPrimaryBalanceSheetAssignmentLedgerRows(["1Q26"], tslaCtx, tslaFillRows);
const tslaByLabel = new Map(tslaLedger.map((row) => [row.sourceLineItemLabel, row]));

assert.equal(hooks.resolveAccruedLiabilities("1Q26", tslaCtx).value, 14_554_000_000);
assert.equal(
  hooks.resolveGoodwill("1Q26", tslaCtx).value,
  null,
  "A missing goodwill disclosure must remain unresolved instead of being converted into an invented zero."
);

const tslaOtherNonCurrentLiabilities = hooks.resolveOtherNonCurrentLiabilities("1Q26", tslaCtx);
assert.equal(tslaOtherNonCurrentLiabilities.value, 17_059_000_000);
assert.equal(tslaOtherNonCurrentLiabilities.sources.some((source) => source.concept === "CryptoAssetFairValueNoncurrent"), false);
assert.equal(tslaOtherNonCurrentLiabilities.sources.some((source) => source.concept === "RedeemableNoncontrollingInterestEquityCarryingAmount"), true);
assert.equal(hooks.resolveTotalLiabilities("1Q26", tslaCtx).value, 58_979_000_000);
assert.equal(tslaByLabel.get("Digital assets").assignedModelRow, "Other Non-Current Assets");
assert.equal(tslaByLabel.get("Redeemable noncontrolling interests in subsidiaries").assignedModelRow, "Other Non-Current Liabilities");
assert.notEqual(tslaByLabel.get("Redeemable noncontrolling interests in subsidiaries").assignmentStatus, "explicitly_excluded_with_reason");

const priorAnnualAccession = "000000248825000010";
const priorAnnualEntry = {
  accessionNumber: priorAnnualAccession,
  accessionKey: priorAnnualAccession,
  form: "10-K",
  filingDate: "2026-02-01",
  reportDate: "2025-12-31",
  fiscalYear: 2025,
  fiscalQuarter: 4,
  quarterPeriod: "4Q25",
  annualPeriod: "FY25"
};
const coverageCtx = {
  ...ctx,
  fiscalPeriods: {
    entries: [priorAnnualEntry, filingEntry],
    byAccession: new Map([
      [priorAnnualAccession, priorAnnualEntry],
      [accession, filingEntry]
    ]),
    byReportDate: new Map([
      [priorAnnualEntry.reportDate, priorAnnualEntry],
      [reportDate, filingEntry]
    ]),
    reportedPeriods: new Set(["4Q25", "FY25", "1Q26"]),
    fiscalYearEndMonth: 12,
    fiscalYearEndDay: 31
  }
};
const validationWorkbook = new ExcelJS.Workbook();
const validationSheet = validationWorkbook.addWorksheet("Model");
validationSheet.getCell("A1").value = "Balance Sheet";
fillRows.forEach((row) => {
  validationSheet.getCell(row.row + 1, 1).value = row.label;
});
validationSheet.getCell(2, 3).value = 2_997;
validationSheet.getCell(3, 3).value = 5_785;
validationSheet.getCell(4, 3).value = 874;
validationSheet.getCell(5, 3).value = 850;
validationSheet.getCell("A6").value = "Cash Flow Statement";
const coverageWarnings = [];
const coverageErrors = hooks.validatePrimaryBalanceSheetAssignmentCoverage(
  validationSheet,
  ["4Q25", "1Q26"],
  [2, 3],
  coverageCtx,
  new hooks.FormulaEvaluator(validationSheet),
  coverageWarnings,
  fillRows.map((row) => ({ ...row, row: row.row + 1 }))
);
assert.ok(
  coverageErrors.some((error) => /Balance Sheet 4Q25: no primary balance sheet assignment ledger rows/.test(error)),
  "Every reported SEC balance-sheet period must fail closed when its assignment ledger is missing."
);
assert.equal(
  coverageWarnings.some((warning) => /Balance Sheet 4Q25: no primary balance sheet assignment ledger rows/.test(warning)),
  false,
  "Missing reported-period coverage is a blocking error, not an advisory warning."
);

const globallyEmptyBalanceLedgerCtx = {
  ...ctx,
  filingPackageStatements: []
};
const emptyBalanceLedgerErrors = hooks.validatePrimaryBalanceSheetAssignmentCoverage(
  validationSheet,
  ["1Q26"],
  [3],
  globallyEmptyBalanceLedgerCtx,
  new hooks.FormulaEvaluator(validationSheet, { useCachedFormulaResults: false, allowCachedFormulaResultFallback: false }),
  [],
  fillRows.map((row) => ({ ...row, row: row.row + 1 }))
);
assert.ok(
  emptyBalanceLedgerErrors.some((error) => /Balance Sheet 1Q26: no primary balance sheet assignment ledger rows/.test(error)),
  "A globally empty balance-sheet assignment ledger must fail closed when the requested period is SEC-reported."
);

const investmentAssetRow = {
  ...balanceRow(
    1,
    "Equity Securities without Readily Determinable Fair Value, Amount",
    "EquitySecuritiesWithoutReadilyDeterminableFairValueAmount",
    62_300_000,
    "non_current"
  ),
  parentSubtotal: { label: "Assets", concept: "Assets", relationship: "calculation" }
};
const investmentCtx = {
  ...ctx,
  instant: new Map([
    [
      "1Q26",
      new Map([
        [
          investmentAssetRow.xbrlConcept,
          {
            concept: investmentAssetRow.xbrlConcept,
            label: investmentAssetRow.rowLabel,
            value: investmentAssetRow.value,
            unit: "USD",
            sourceLayer: "sec_filing_package",
            accn: accession,
            end: reportDate,
            periodKey: "1Q26",
            periodType: "instant"
          }
        ]
      ])
    ]
  ]),
  filingPackageStatements: [
    {
      statementName: "Condensed Consolidated Balance Sheets",
      sourceTableType: "primary_statement",
      accession,
      reportingPeriod: reportDate,
      form: "10-Q",
      rows: [investmentAssetRow]
    }
  ]
};
const investmentLedger = hooks.buildPrimaryBalanceSheetAssignmentLedgerRows(
  ["1Q26"],
  investmentCtx,
  ["Other Non-Current Assets", "Common Stock & APIC"].map((label, index) => ({
    row: index + 1,
    label,
    classification: "direct",
    statement: "balance",
    kind: "instant",
    scale: 1_000_000
  }))
);
assert.equal(investmentLedger[0].assignedModelRow, "Other Non-Current Assets");
assert.equal(investmentLedger[0].side, "assets");

const duplicateAccruedRows = [
  balanceRow(1, "Accrued Liabilities, Current", "AccruedLiabilitiesCurrent", 263_100_000),
  balanceRow(2, "Accrued liabilities", "AccruedLiabilitiesCurrent", 263_100_000.00000003),
  balanceRow(3, "Accrued Income Taxes, Current", "AccruedIncomeTaxesCurrent", 249_100_000)
];
const duplicateAccruedCtx = {
  ...ctx,
  filingPackageStatements: [
    {
      statementName: "Condensed Consolidated Balance Sheets",
      sourceTableType: "primary_statement",
      accession,
      reportingPeriod: reportDate,
      form: "10-Q",
      rows: duplicateAccruedRows
    }
  ]
};
assert.equal(
  hooks.resolveAccruedLiabilities("1Q26", duplicateAccruedCtx).value,
  512_200_000,
  "Equivalent primary-statement facts from HTML and presentation structures must be counted once before adding separately reported tax accruals."
);

function combinedBalanceRow(rowOrder, rowLabel, xbrlConcept, value, section, parentConcept) {
  return {
    ...balanceRow(rowOrder, rowLabel, xbrlConcept, value, section),
    parentSubtotal: {
      label: parentConcept,
      concept: parentConcept,
      relationship: "calculation"
    }
  };
}

const combinedIntangibleRows = [
  combinedBalanceRow(1, "Cash and cash equivalents", "CashAndCashEquivalentsAtCarryingValue", 100_000_000, "current", "AssetsCurrent"),
  combinedBalanceRow(2, "Goodwill and acquisition-related intangible assets, net", "IntangibleAssetsNetIncludingGoodwill", 330_000_000, "non_current", "Assets"),
  combinedBalanceRow(3, "Other assets", "OtherAssetsNoncurrent", 670_000_000, "non_current", "Assets"),
  combinedBalanceRow(4, "Total assets", "Assets", 1_100_000_000, "non_current", "Assets"),
  combinedBalanceRow(5, "Accounts payable", "AccountsPayableCurrent", 100_000_000, "current", "LiabilitiesCurrent"),
  combinedBalanceRow(6, "Other long-term liabilities", "OtherLiabilitiesNoncurrent", 200_000_000, "non_current", "Liabilities"),
  combinedBalanceRow(7, "Additional paid-in capital", "AdditionalPaidInCapital", 800_000_000, "non_current", "StockholdersEquity"),
  combinedBalanceRow(8, "Total liabilities and stockholders' equity", "LiabilitiesAndStockholdersEquity", 1_100_000_000, "non_current", "LiabilitiesAndStockholdersEquity")
];
const separateIntangibleFact = {
  concept: "FiniteLivedIntangibleAssetsNet",
  label: "Finite-lived intangible assets, net",
  value: 62_000_000,
  unit: "USD",
  sourceLayer: "sec_filing_package",
  accn: accession,
  end: reportDate,
  periodKey: "1Q26",
  periodType: "instant"
};
const combinedIntangibleCtx = {
  ...ctx,
  instant: new Map([
    [
      "1Q26",
      new Map([
        ["FiniteLivedIntangibleAssetsNet", separateIntangibleFact],
        ...combinedIntangibleRows.map((row) => [
          row.xbrlConcept,
          {
            concept: row.xbrlConcept,
            label: row.rowLabel,
            value: row.value,
            unit: "USD",
            sourceLayer: "sec_filing_package",
            accn: accession,
            end: reportDate,
            periodKey: "1Q26",
            periodType: "instant"
          }
        ])
      ])
    ]
  ]),
  filingPackageStatements: [
    {
      statementName: "Condensed Consolidated Balance Sheets",
      sourceTableType: "primary_statement",
      accession,
      reportingPeriod: reportDate,
      form: "10-Q",
      rows: combinedIntangibleRows
    }
  ]
};
const combinedGoodwill = hooks.resolveGoodwill("1Q26", combinedIntangibleCtx);
assert.equal(
  combinedGoodwill.value,
  268_000_000,
  "Goodwill must be derived from an SEC combined intangibles-and-goodwill balance less same-period separately disclosed SEC intangibles."
);
assert.equal(combinedGoodwill.sources[0].concept, "GoodwillDerivedFromCombinedIntangibles");
assert.equal(combinedGoodwill.sources[0].value, 268_000_000, "The final goodwill derivation must be recorded separately from its SEC inputs.");

const combinedFillRows = [
  ["Cash & Cash Equivalents", 2],
  ["Intangible Assets, Net", 3],
  ["Goodwill", 4],
  ["Other Non-Current Assets", 5],
  ["Accounts Payable", 6],
  ["Other Non-Current Liabilities", 7],
  ["Common Stock & APIC", 8]
].map(([label, row]) => ({ row, label, classification: "direct", statement: "balance", kind: "instant", scale: 1_000_000 }));
const combinedWorkbook = new ExcelJS.Workbook();
const combinedSheet = combinedWorkbook.addWorksheet("Model");
combinedSheet.getCell("A1").value = "Balance Sheet";
for (const fillRow of combinedFillRows) combinedSheet.getCell(fillRow.row, 1).value = fillRow.label;
[100, 62, 268, 670, 100, 200, 800].forEach((value, index) => combinedSheet.getCell(index + 2, 2).value = value);
combinedSheet.getCell("A9").value = "Cash Flow Statement";
const combinedWarnings = [];
assert.deepEqual(
  hooks.validatePrimaryBalanceSheetAssignmentCoverage(
    combinedSheet,
    ["1Q26"],
    [2],
    combinedIntangibleCtx,
    new hooks.FormulaEvaluator(combinedSheet, { useCachedFormulaResults: false, allowCachedFormulaResultFallback: false }),
    combinedWarnings,
    combinedFillRows
  ),
  [],
  "A combined SEC carrying amount must reconcile across separate Intangible Assets and Goodwill model rows without double-counting."
);
assert.ok(combinedWarnings.some((warning) => /reports intangibles including goodwill/i.test(warning)));

const workbook = new ExcelJS.Workbook();
const formulaSheet = workbook.addWorksheet("Model");
formulaSheet.getCell("A1").value = 100;
formulaSheet.getCell("A2").value = 43;
formulaSheet.getCell("A3").value = { formula: "A1-A2", result: 0 };
const formulaEvaluator = new hooks.FormulaEvaluator(formulaSheet, { useCachedFormulaResults: true });
assert.equal(hooks.statementMetricCellValue(formulaSheet.getCell("A3"), formulaEvaluator, 0), 57);

console.log("AMD balance sheet liability classification regression passed.");
