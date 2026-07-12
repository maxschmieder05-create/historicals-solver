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

const servicePath = path.join(__dirname, "..", "server", "fill-model", "fill-model-service.ts");
const hooks = loadTypeScriptModule(servicePath).__fillModelServiceTestHooks;
const accession = "0000000000-26-000001";

function source(concept, value, period, periodType = "quarterly") {
  return {
    concept,
    label: concept.replace(/([a-z])([A-Z])/g, "$1 $2"),
    value,
    unit: "USD",
    sourceLayer: "sec_live_companyfacts",
    accn: accession,
    form: periodType === "annual" ? "10-K" : "10-Q",
    periodKey: period,
    periodType
  };
}

function factMap(entries) {
  const map = new Map();
  for (const [period, facts] of Object.entries(entries)) {
    map.set(period, new Map(facts.map((fact) => [fact.concept, fact])));
  }
  return map;
}

function context(duration = {}, instant = {}) {
  return { duration: factMap(duration), instant: factMap(instant) };
}

const combinedDaCtx = context({
  "1Q26": [
    source("DepreciationDepletionAndAmortization", 150, "1Q26"),
    source("Depreciation", 100, "1Q26"),
    source("AmortizationOfIntangibleAssets", 50, "1Q26")
  ]
});
const combinedDa = hooks.resolveCashFlowDepreciationAmortization("1Q26", combinedDaCtx);
assert.equal(combinedDa.value, 150);
assert.equal(combinedDa.sources[0].concept, "DepreciationDepletionAndAmortization", "the aggregate D&A fact must win without double counting components");

const separateDaCtx = context({
  "1Q26": [source("Depreciation", 100, "1Q26"), source("AmortizationOfIntangibleAssets", 50, "1Q26")]
});
const separateDa = hooks.resolveCashFlowDepreciationAmortization("1Q26", separateDaCtx);
assert.equal(separateDa.value, 150);
assert.equal(separateDa.sources[0].sourceLayer, "derived", "separate depreciation and amortization must expose one derived output first");

const derivedPpeDepreciationCtx = context({
  "1Q26": [source("DepreciationAndAmortization", 150, "1Q26"), source("AmortizationOfIntangibleAssets", 50, "1Q26")]
});
const ppeDepreciation = hooks.resolvePpeScheduleDepreciationExpense("1Q26", derivedPpeDepreciationCtx);
assert.equal(ppeDepreciation.value, -100);
assert.equal(ppeDepreciation.sources[0].concept, "PpeDepreciationDerivedFromDaLessIntangibleAmortization");

const coreCashFlowCtx = context({
  "1Q26": [
    source("PaymentsToAcquirePropertyPlantAndEquipment", 200, "1Q26"),
    source("PaymentsToAcquireIntangibleAssets", 30, "1Q26"),
    source("PaymentsForRepurchaseOfCommonStock", 40, "1Q26"),
    source("PaymentsToAcquireBusinessesNetOfCashAcquired", 100, "1Q26"),
    source("ProceedsFromDivestitureOfBusinessesNetOfCashDivested", 25, "1Q26"),
    source("ProceedsFromLinesOfCredit", 100, "1Q26"),
    source("RepaymentsOfLinesOfCredit", 70, "1Q26"),
    source("PaymentsToNoncontrollingInterests", 9, "1Q26"),
    source("ProceedsFromNoncontrollingInterests", 4, "1Q26")
  ]
});
assert.equal(
  hooks.resolveCashFlowDurationConcepts(
    "1Q26",
    coreCashFlowCtx,
    ["PaymentsToAcquirePropertyPlantAndEquipment"],
    "outflow",
    "CashFlowCapitalExpenditures",
    "capex"
  ).value,
  -200
);
assert.equal(hooks.resolvePpeScheduleCapitalExpenditures("1Q26", coreCashFlowCtx).value, 200, "cash-flow capex and PP&E additions must use opposite model signs");
assert.equal(
  hooks.resolveCashFlowDurationConcepts(
    "1Q26",
    coreCashFlowCtx,
    ["PaymentsToAcquireIntangibleAssets"],
    "outflow",
    "CashFlowPurchasesOfIntangibles",
    "intangibles"
  ).value,
  -30
);
assert.equal(hooks.resolveIntangibleSchedulePurchases("1Q26", coreCashFlowCtx).value, 30);
assert.equal(
  hooks.resolveCashFlowDurationConcepts(
    "1Q26",
    coreCashFlowCtx,
    ["PaymentsForRepurchaseOfCommonStock"],
    "outflow",
    "CashFlowShareRepurchases",
    "repurchases"
  ).value,
  -40
);
const acquisitions = hooks.resolveCashFlowBusinessAcquisitions("1Q26", coreCashFlowCtx);
assert.equal(acquisitions.value, -75);
assert.equal(acquisitions.sources[0].concept, "CashFlowBusinessAcquisitionsNet");
const revolver = hooks.resolveCashFlowRevolverIssuanceRepayment("1Q26", coreCashFlowCtx);
assert.equal(revolver.value, 30);
assert.equal(revolver.sources[0].concept, "RevolverIssuanceRepaymentNet");
const noncontrolling = hooks.resolveCashFlowNoncontrollingInterestChange("1Q26", coreCashFlowCtx);
assert.equal(noncontrolling.value, -5);
assert.equal(noncontrolling.sources[0].concept, "CashFlowNoncontrollingInterestsNet");

const q4CapexCtx = context({
  FY26: [source("PaymentsToAcquireProductiveAssets", 400, "FY26", "annual")],
  "1Q26": [source("PaymentsToAcquirePropertyPlantAndEquipment", 100, "1Q26")],
  "2Q26": [source("PaymentsToAcquirePropertyPlantAndEquipment", 90, "2Q26")],
  "3Q26": [source("PaymentsToAcquirePropertyPlantAndEquipment", 110, "3Q26")]
});
const q4Capex = hooks.resolveCashFlowDurationConcepts(
  "4Q26",
  q4CapexCtx,
  ["PaymentsToAcquirePropertyPlantAndEquipment", "PaymentsToAcquireProductiveAssets"],
  "outflow",
  "CashFlowCapitalExpenditures",
  "capex"
);
assert.equal(q4Capex.value, -100);
assert.equal(q4Capex.sources[0].sourceLayer, "derived");
assert.equal(q4Capex.sources[0].derivedTotalValue, -400, "an explicit annual-to-Q4 bridge must retain the sign-normalized SEC annual total, not recompute peer-input arithmetic as annual metadata");
assert.ok(
  q4Capex.sources.some((item) => item.concept === "PaymentsToAcquireProductiveAssets"),
  "an annual productive-assets concept may bridge quarterly PP&E purchases when the issuer changes equivalent SEC concepts by filing type"
);

const operatingDetailCtx = context({
  "1Q26": [
    source("NetCashProvidedByUsedInOperatingActivities", 100, "1Q26"),
    source("ProfitLoss", 20, "1Q26"),
    source("DepreciationDepletionAndAmortization", 10, "1Q26"),
    source("ShareBasedCompensation", 5, "1Q26"),
    source("IncreaseDecreaseInAccountsReceivable", -2, "1Q26"),
    source("IncreaseDecreaseInInventories", -3, "1Q26"),
    source("IncreaseDecreaseInOtherOperatingCapitalNet", 4, "1Q26")
  ]
});
const workingCapital = hooks.resolveCashFlowWorkingCapital("1Q26", operatingDetailCtx);
assert.equal(workingCapital.value, -1, "working capital should sum non-overlapping SEC cash-flow components when no aggregate is reported");
assert.equal(workingCapital.sources[0].concept, "CashFlowWorkingCapitalFromComponents");
const longTermItems = hooks.resolveCashFlowLongTermItems("1Q26", operatingDetailCtx);
assert.equal(longTermItems.value, 66, "the long-term-items catch-all should be the exact SEC operating-cash-flow residual");
assert.equal(longTermItems.sources[0].concept, "CashFlowLongTermItemsOperatingResidual");
const aggregateLongTermItems = hooks.resolveCashFlowLongTermItems(
  "1Q26",
  context({ "1Q26": [source("OtherOperatingActivitiesCashFlowStatement", 7, "1Q26")] })
);
assert.equal(aggregateLongTermItems.value, 7);
assert.equal(aggregateLongTermItems.sources[0].concept, "OtherOperatingActivitiesCashFlowStatement");
const singleLongTermComponent = hooks.resolveCashFlowLongTermItems(
  "1Q26",
  context({ "1Q26": [source("DeferredIncomeTaxExpenseBenefit", 3, "1Q26")] })
);
assert.equal(singleLongTermComponent.value, 3);
const ambiguousLongTermComponentsCtx = context({
  "1Q26": [source("OtherNoncashIncomeExpense", 4, "1Q26"), source("DeferredIncomeTaxExpenseBenefit", 3, "1Q26")]
});
assert.equal(
  hooks.resolveCashFlowLongTermItems("1Q26", ambiguousLongTermComponentsCtx).value,
  null,
  "multiple possible long-term components must fail closed without primary cash-flow statement evidence that they are separate rows"
);
ambiguousLongTermComponentsCtx.filingPackageStatements = [
  {
    statementName: "Consolidated Statements of Cash Flows",
    sourceTableType: "primary_statement",
    accession,
    rows: [
      { xbrlConcept: "OtherNoncashIncomeExpense" },
      { xbrlConcept: "DeferredIncomeTaxExpenseBenefit" }
    ]
  }
];
const provenLongTermComponents = hooks.resolveCashFlowLongTermItems("1Q26", ambiguousLongTermComponentsCtx);
assert.equal(provenLongTermComponents.value, 7);
assert.equal(provenLongTermComponents.sources[0].concept, "CashFlowLongTermItemsFromSeparateComponents");
assert.deepEqual(
  hooks.expectedUnitsForFillRow({ label: "Stock-Based Compensation", concepts: ["ShareBasedCompensation"] }),
  ["USD"],
  "monetary stock compensation must not be misclassified as a share-count fact merely because its concept contains Share"
);
assert.deepEqual(
  hooks.expectedUnitsForFillRow({ label: "Basic Shares", concepts: ["WeightedAverageNumberOfSharesOutstandingBasic"] }),
  ["shares"]
);
const incompleteWorkingCapitalCtx = context({
  "1Q26": [
    source("IncreaseDecreaseInAccountsReceivable", -2, "1Q26"),
    source("IncreaseDecreaseInInventories", -3, "1Q26")
  ]
});
assert.equal(
  hooks.resolveCashFlowWorkingCapital("1Q26", incompleteWorkingCapitalCtx).value,
  null,
  "two asset-only components without an SEC catch-all other-operating-capital line must not be presented as complete working capital"
);

const cashBalanceCtx = context(
  {},
  {
    "4Q25": [source("CashAndCashEquivalentsAtCarryingValue", 250, "4Q25", "instant")],
    "1Q26": [source("CashAndCashEquivalentsAtCarryingValue", 275, "1Q26", "instant")]
  }
);
const beginningCash = hooks.resolveBeginningCashBalance("1Q26", cashBalanceCtx);
assert.equal(beginningCash.value, 250);
assert.equal(beginningCash.sources[0].concept, "BeginningCashBalanceFromPriorPeriod");
assert.equal(beginningCash.sources[0].periodKey, "1Q26");
assert.equal(beginningCash.sources[1].periodKey, "4Q25");
assert.equal(hooks.resolveCashFlowEndingCashBalance("1Q26", cashBalanceCtx).value, 275);

const periodCalendarCtx = context();
periodCalendarCtx.fiscalPeriods = {
  entries: [
    { quarterPeriod: "4Q22", reportDate: "2022-12-31" },
    { quarterPeriod: "1Q23", reportDate: "2023-03-31" }
  ]
};
const periodCalendarWorkbook = new ExcelJS.Workbook();
const periodCalendarSheet = periodCalendarWorkbook.addWorksheet("Model");
periodCalendarSheet.getCell("C26").value = "Days In Period";
periodCalendarSheet.getCell("F26").value = 91;
const periodCalendarResult = hooks.refreshDaysInPeriodMetadata(
  periodCalendarSheet,
  [{ period: "1Q23", col: 6 }],
  periodCalendarCtx
);
assert.equal(periodCalendarSheet.getCell("F26").value, 90, "days-in-period metadata must refresh from the selected issuer's SEC report dates");
assert.equal(periodCalendarResult.filledCells, 1);

const workbook = new ExcelJS.Workbook();
const sheet = workbook.addWorksheet("Model");
sheet.getCell("C12").value = { formula: 'HYPERLINK("#Model!C73","Cash Flow Statement")', result: "Cash Flow Statement" };
sheet.getCell("C20").value = "Income Statement";
sheet.getCell("C45").value = "Net Income";
sheet.getCell("F45").value = 12;
sheet.getCell("G45").value = 13;
sheet.getCell("C73").value = "Cash Flow Statement";
const labels = new Map([
  [75, "Net Income"],
  [76, "Depreciation & Amortization"],
  [77, "Stock-Based Compensation"],
  [78, "(Increase)/Decrease in Working Capital"],
  [79, "(Increase)/Decrease in LT Items"],
  [80, "Net Cash From Operating Activities"],
  [82, "Capital Expenditures"],
  [83, "Purchases of Intangibles"],
  [84, "Proceeds From/(Acquisitions of) Businesses"],
  [85, "Net Cash From Investment Activities"],
  [87, "Cash Flow Available for Financing Activities"],
  [89, "Issuance/(Repayment) of Revolver"],
  [90, "Issuance of Debt"],
  [91, "(Repayment of Debt)"],
  [92, "Issuance of Equity"],
  [93, "(Repurchase) of Equity"],
  [94, "Dividends"],
  [95, "Change in Noncontrolling Interests"],
  [96, "Net Cash From Financing Activities"],
  [98, "Effect of FX Rate Changes on Cash"],
  [100, "Ending Cash Adjustments"],
  [103, "Net Change in Cash"],
  [106, "Free Cash Flow Analysis"]
]);
for (const [row, label] of labels) sheet.getCell(row, 3).value = label;
sheet.getCell("F80").value = { formula: "F75+F76+F77+F78+F79", result: 0 };
const formulaAudit = [];
const formulaResult = hooks.writeCashFlowHistoricalFormulas(sheet, ["1Q26", "2Q26"], [6, 7], formulaAudit);
assert.equal(hooks.cashFlowStatementRows(sheet)[0], 73, "the structurally complete statement must win over a table-of-contents link");
assert.equal(sheet.getCell("F80").value.formula, "F75+F76+F77+F78+F79", "existing cash-flow formulas must be preserved exactly");
assert.equal(sheet.getCell("G80").value.formula, "SUM(G75,G76,G77,G78,G79)");
assert.equal(sheet.getCell("F75").value.formula, "F45");
assert.equal(sheet.getCell("G85").value.formula, "SUM(G82,G83,G84)");
assert.equal(sheet.getCell("G87").value.formula, "SUM(G80,G85)");
assert.equal(sheet.getCell("G96").value.formula, "SUM(G89,G90,G91,G92,G93,G94,G95)");
assert.equal(sheet.getCell("G103").value.formula, "SUM(G87,G96,G98,G100)");
assert.equal(formulaResult.formulasPreserved, 1);
assert.equal(formulaResult.formulasCreated, 11);
assert.equal(formulaResult.formulasReplaced, 0);
assert.equal(formulaResult.formulasCleared, 0);
assert.equal(formulaAudit.length, 11);

const duplicateLabelCacheWorkbook = new ExcelJS.Workbook();
const duplicateLabelCacheSheet = duplicateLabelCacheWorkbook.addWorksheet("Model");
duplicateLabelCacheSheet.getCell("C20").value = "Income Statement";
duplicateLabelCacheSheet.getCell("C45").value = "Net Income";
duplicateLabelCacheSheet.getCell("F45").value = { formula: "40+60", result: 100 };
duplicateLabelCacheSheet.getCell("C73").value = "Cash Flow Statement";
duplicateLabelCacheSheet.getCell("C75").value = "Net Income";
duplicateLabelCacheSheet.getCell("F75").value = { formula: "F45", result: 999 };
duplicateLabelCacheSheet.getCell("C106").value = "Free Cash Flow Analysis";
duplicateLabelCacheSheet.getCell("C108").value = "Net Income";
duplicateLabelCacheSheet.getCell("F108").value = { formula: "F45", result: 999 };
hooks.markReportedPeriodColumns(duplicateLabelCacheSheet, [{ period: "1Q26", col: 6 }]);
hooks.refreshHistoricalFormulaCachedResults(duplicateLabelCacheWorkbook, [6], ["Model"]);
assert.equal(duplicateLabelCacheSheet.getCell("F75").value.result, 100, "cash-flow Net Income cache should refresh from current precedents");
assert.equal(duplicateLabelCacheSheet.getCell("F108").value.result, 100, "downstream duplicate Net Income labels must not be mistaken for the primary statement row");

const formulaDetailWorkbook = new ExcelJS.Workbook();
const formulaDetailSheet = formulaDetailWorkbook.addWorksheet("Model");
formulaDetailSheet.getCell("C73").value = "Cash Flow Statement";
formulaDetailSheet.getCell("C76").value = "Depreciation & Amortization";
formulaDetailSheet.getCell("F45").value = 20;
formulaDetailSheet.getCell("F76").value = { formula: "F45", result: 20 };
hooks.markReportedPeriodColumns(formulaDetailSheet, [{ period: "1Q26", col: 6 }]);
const formulaDetailRow = hooks.fillRowForContext({
  sheetName: "Model",
  row: 76,
  label: "Depreciation & Amortization",
  isCashFlowStatementRow: true,
  sectionHeader: "Cash Flow Statement",
  indentation: 0,
  hasHistoricalFormula: true,
  hasHardcodedInput: false,
  hasNetRevenueInterestExpenseAbove: false,
  projectedColumns: 0,
  signConvention: 1
});
assert.equal(formulaDetailRow.classification, "direct", "known cash-flow detail formulas must retain their SEC resolver instead of being downgraded to generic formulas");
assert.equal(typeof formulaDetailRow.resolver, "function");
assert.equal(
  hooks.validateCashFlowResolvableDetailCompleteness(
    formulaDetailSheet,
    [{ period: "1Q26", col: 6 }],
    context({ "1Q26": [source("DepreciationDepletionAndAmortization", 10_000_000, "1Q26")] }),
    [formulaDetailRow]
  ).length,
  1,
  "a formula-backed cash-flow detail row must fail when its strict value disagrees with SEC"
);
formulaDetailSheet.getCell("F45").value = 10;
assert.deepEqual(
  hooks.validateCashFlowResolvableDetailCompleteness(
    formulaDetailSheet,
    [{ period: "1Q26", col: 6 }],
    context({ "1Q26": [source("DepreciationDepletionAndAmortization", 10_000_000, "1Q26")] }),
    [formulaDetailRow]
  ),
  [],
  "a formula-backed cash-flow detail row may pass only when its strict value ties SEC"
);

const reportedTotalsWorkbook = new ExcelJS.Workbook();
const reportedTotalsSheet = reportedTotalsWorkbook.addWorksheet("Model");
reportedTotalsSheet.getCell("C12").value = { formula: 'HYPERLINK("#Model!C73","Cash Flow Statement")', result: "Cash Flow Statement" };
reportedTotalsSheet.getCell("C20").value = "Income Statement";
reportedTotalsSheet.getCell("C45").value = "Net Income";
reportedTotalsSheet.getCell("F45").value = 60;
reportedTotalsSheet.getCell("C73").value = "Cash Flow Statement";
for (const [row, label] of labels) reportedTotalsSheet.getCell(row, 3).value = label;
for (const [row, value] of [
  [76, 10],
  [77, 5],
  [78, 15],
  [79, 10],
  [82, -29],
  [83, 0],
  [84, 0],
  [89, 0],
  [90, 0],
  [91, 0],
  [92, 0],
  [93, 0],
  [94, -20],
  [95, 0],
  [98, 0],
  [100, 0]
]) {
  reportedTotalsSheet.getCell(row, 6).value = value;
}
const reportedTotalsCtx = context({
  "1Q26": [
    source("NetCashProvidedByUsedInOperatingActivities", 100_000_000, "1Q26"),
    source("NetCashProvidedByUsedInInvestingActivities", -30_000_000, "1Q26"),
    source("NetCashProvidedByUsedInFinancingActivities", -20_000_000, "1Q26"),
    source(
      "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalentsPeriodIncreaseDecreaseIncludingExchangeRateEffect",
      50_000_000,
      "1Q26"
    )
  ]
});
const reportedTotalsAudit = [];
hooks.writeCashFlowHistoricalFormulas(reportedTotalsSheet, ["1Q26"], [6], reportedTotalsAudit);
const reportedTotalsResult = hooks.reconcileCashFlowReportedTotalsForSheet(
  reportedTotalsSheet,
  ["1Q26"],
  [6],
  reportedTotalsCtx,
  reportedTotalsAudit
);
assert.equal(reportedTotalsSheet.getCell("F80").value.formula, "SUM(F75,F76,F77,F78,F79)");
assert.equal(reportedTotalsSheet.getCell("F80").value.result, 100);
assert.equal(
  reportedTotalsSheet.getCell("F85").value,
  -30,
  "a detail formula that does not tie must fall back to the direct reported SEC cash-flow total"
);
assert.equal(reportedTotalsSheet.getCell("F96").value.result, -20);
assert.equal(reportedTotalsSheet.getCell("F103").value.result, 50);
assert.equal(reportedTotalsResult.reportedTotalsWritten, 1);
assert.equal(reportedTotalsResult.reportedTotalsFormulaTied, 3);
assert.ok(
  reportedTotalsAudit.some((row) => row.cell === "F80" && row.formulaPreserved && row.accession === accession),
  "a formula-preserved cash-flow total must carry its own current SEC amount support"
);
assert.ok(
  reportedTotalsAudit.some((row) => row.cell === "F85" && row.mappingType === "direct" && row.accession === accession),
  "a direct cash-flow total fallback must retain current SEC provenance"
);
assert.equal(
  hooks.isRequiredCoreHistoricalInputCell(reportedTotalsSheet, {
    row: 80,
    label: "Net Cash From Operating Activities",
    classification: "formula",
    statement: "support",
    kind: "duration"
  }),
  true,
  "cash-flow financial formulas must undergo source-ledger amount or dependency validation"
);
assert.equal(
  hooks.isRequiredCoreHistoricalInputCell(reportedTotalsSheet, {
    row: 12,
    label: "Cash Flow Statement",
    classification: "formula",
    statement: "support",
    kind: "duration"
  }),
  false,
  "table-of-contents and header formulas must remain exempt from financial provenance checks"
);

const unsafeFormulaWorkbook = new ExcelJS.Workbook();
const unsafeFormulaSheet = unsafeFormulaWorkbook.addWorksheet("Model");
unsafeFormulaSheet.getCell("C73").value = "Cash Flow Statement";
for (const [row, label] of [
  [75, "Net Income"],
  [76, "Depreciation & Amortization"],
  [77, "Stock-Based Compensation"],
  [78, "(Increase)/Decrease in Working Capital"],
  [79, "(Increase)/Decrease in LT Items"],
  [80, "Net Cash From Operating Activities"],
  [90, "Free Cash Flow Analysis"]
]) {
  unsafeFormulaSheet.getCell(row, 3).value = label;
}
unsafeFormulaSheet.getCell("F75").value = { formula: "LegacyNetIncome", result: 12 };
unsafeFormulaSheet.getCell("F80").value = { formula: "SUM('[Other.xlsx]Model'!F75:F79)", result: 0 };
unsafeFormulaSheet.getCell("G80").value = { formula: "SUM(CashFlowInputs)", result: 0 };
unsafeFormulaSheet.getCell("H80").value = { formula: "SUM(H75:H80)", result: 0 };
unsafeFormulaSheet.getCell("I80").value = { formula: "I75+I76+I77+I78+I79", result: 0 };
unsafeFormulaSheet.getCell("J80").value = { formula: "SUM('Model'!$J$79,$J$77,$J$75,$J$78,$J$76)", result: 0 };
const unsafeFormulaAudit = [];
const unsafeFormulaResult = hooks.writeCashFlowHistoricalFormulas(
  unsafeFormulaSheet,
  ["1Q26", "2Q26", "3Q26", "4Q26", "FY26"],
  [6, 7, 8, 9, 10],
  unsafeFormulaAudit
);
assert.equal(unsafeFormulaSheet.getCell("F75").value, null, "named-range formula with no certifiable source row must fail closed");
assert.equal(unsafeFormulaSheet.getCell("F80").value.formula, "SUM(F75,F76,F77,F78,F79)", "external workbook formulas must be replaced");
assert.equal(unsafeFormulaSheet.getCell("G80").value.formula, "SUM(G75,G76,G77,G78,G79)", "named-range formulas must be replaced");
assert.equal(unsafeFormulaSheet.getCell("H80").value.formula, "SUM(H75,H76,H77,H78,H79)", "wrong-row formulas must be replaced");
assert.equal(unsafeFormulaSheet.getCell("I80").value.formula, "I75+I76+I77+I78+I79", "a valid additive equivalent formula must be preserved exactly");
assert.equal(
  unsafeFormulaSheet.getCell("J80").value.formula,
  "SUM('Model'!$J$79,$J$77,$J$75,$J$78,$J$76)",
  "a valid same-sheet formula with reordered absolute references must be preserved exactly"
);
assert.equal(unsafeFormulaResult.formulasCreated, 3);
assert.equal(unsafeFormulaResult.formulasReplaced, 3);
assert.equal(unsafeFormulaResult.formulasPreserved, 2);
assert.equal(unsafeFormulaResult.formulasCleared, 1);
assert.ok(
  unsafeFormulaAudit.some((row) => row.formulaStatus === "unsupported historical cash-flow formula cleared"),
  "fail-closed formula removal must be visible in the audit trail"
);

const reorderedWorkbook = new ExcelJS.Workbook();
const reorderedSheet = reorderedWorkbook.addWorksheet("Reordered Model");
reorderedSheet.getCell("C5").value = "Net Income";
reorderedSheet.getCell("C20").value = "Cash Flow Statement";
for (const [row, label] of [
  [23, "Stock-Based Compensation"],
  [25, "Net Income"],
  [26, "Inserted Unrelated Subtotal"],
  [27, "(Increase)/Decrease in Working Capital"],
  [29, "Depreciation & Amortization"],
  [31, "(Increase)/Decrease in LT Items"],
  [35, "Net Cash From Operating Activities"],
  [50, "Free Cash Flow Analysis"]
]) {
  reorderedSheet.getCell(row, 3).value = label;
}
reorderedSheet.getCell("F35").value = { formula: "SUM(F23:F31)", result: 0 };
reorderedSheet.getCell("G35").value = { formula: "SUM($G$23,$G$25,$G$27,$G$29,$G$31)", result: 0 };
const reorderedAudit = [];
const reorderedResult = hooks.writeCashFlowHistoricalFormulas(reorderedSheet, ["1Q26", "2Q26"], [6, 7], reorderedAudit);
assert.equal(
  reorderedSheet.getCell("F35").value.formula,
  "SUM(F25,F29,F23,F27,F31)",
  "a stale range that absorbs an inserted unrelated row must be replaced with explicit known-cell references"
);
assert.equal(
  reorderedSheet.getCell("G35").value.formula,
  "SUM($G$23,$G$25,$G$27,$G$29,$G$31)",
  "reordered templates must preserve formulas whose exact precedent multiset remains equivalent"
);
assert.equal(reorderedResult.formulasReplaced, 1);
assert.equal(reorderedResult.formulasPreserved, 1);

const discoveryWorkbook = new ExcelJS.Workbook();
const discoverySheet = discoveryWorkbook.addWorksheet("Model");
discoverySheet.getCell("C73").value = "Cash Flow Statement";
discoverySheet.getCell("C76").value = "Depreciation & Amortization";
discoverySheet.getCell("C80").value = "Net Cash From Operating Activities";
discoverySheet.getCell("F80").value = { formula: "SUM(F75:F79)", result: 0 };
discoverySheet.getCell("C82").value = "Capital Expenditures";
discoverySheet.getCell("F82").font = { color: { argb: "FF0000FF" } };
discoverySheet.getCell("C106").value = "Free Cash Flow Analysis";
const discovered = hooks.discoverFillRows(discoverySheet, [6], [{ period: "1Q26", isEstimate: false }]);
assert.ok(discovered.some((row) => row.row === 76 && row.resolver), "blank CFS D&A must be actively discoverable");
assert.ok(discovered.some((row) => row.row === 80 && row.classification === "formula"), "existing formula totals must remain protected formula rows");
assert.ok(discovered.some((row) => row.row === 82 && row.resolver), "blue blank CFS capex must be actively discoverable");

const finalizerWorkbook = new ExcelJS.Workbook();
const finalizerSheet = finalizerWorkbook.addWorksheet("Model");
finalizerSheet.getCell("C73").value = "Cash Flow Statement";
finalizerSheet.getCell("C75").value = "Depreciation & Amortization";
finalizerSheet.getCell("F75").value = 15;
finalizerSheet.getCell("C76").value = "Unsupported Model-Specific Cash Flow";
finalizerSheet.getCell("F76").value = 999;
finalizerSheet.getCell("C80").value = "Net Cash From Operating Activities";
finalizerSheet.getCell("F80").value = { formula: "SUM(F75:F79)", result: 15 };
finalizerSheet.getCell("C106").value = "Free Cash Flow Analysis";
const finalizerAudit = [
  {
    sheetName: "Model",
    cell: "F75",
    modelRowLabel: "Depreciation & Amortization",
    period: "1Q26",
    valueWritten: 15,
    mappingType: "direct",
    conceptsUsed: "DepreciationDepletionAndAmortization",
    sourceStatement: "support",
    accession,
    sourceUrl: "",
    cellWritable: true,
    formulaPreserved: false,
    writeBlockedReason: "",
    signConvention: "reported",
    confidence: "high",
    validationStatus: "OK!",
    notes: "SEC supported",
    sourceProvenance: [
      {
        role: "sec_source",
        concept: "DepreciationDepletionAndAmortization",
        label: "Depreciation, depletion and amortization",
        value: 15,
        sourceLayer: "sec_live_companyfacts",
        accession,
        form: "10-Q",
        filedDate: "",
        startDate: "",
        endDate: "",
        periodKey: "1Q26",
        periodType: "quarterly",
        periodEvidence: "source"
      }
    ]
  }
];
const finalizerResult = hooks.finalizeCashFlowStatementHistoricalInputs(finalizerSheet, ["1Q26"], [6], finalizerAudit);
assert.equal(finalizerSheet.getCell("F75").value, 15, "current-run SEC-backed cash-flow values must survive finalization");
assert.equal(finalizerSheet.getCell("F76").value, null, "unsupported stale cash-flow hardcodes must fail closed");
assert.equal(finalizerSheet.getCell("F80").value.formula, "SUM(F75:F79)", "formula-derived rows must survive finalization");
assert.equal(finalizerResult.preservedSecBackedCells, 1);
assert.equal(finalizerResult.clearedCells, 1);
assert.equal(finalizerResult.preservedFormulaCells, 1);

console.log("Cash-flow historical completion checks passed.");
