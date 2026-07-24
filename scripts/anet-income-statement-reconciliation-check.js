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

const { __fillModelServiceTestHooks: hooks } = loadTypeScriptModule(sourcePath);
assert.equal(hooks.statementMetricTies(2, 5), false, "a $3mm error on a small line item must never pass accounting validation");
assert.equal(hooks.statementMetricTies(2, 2.1), true, "one-decimal model display rounding remains acceptable");
assert.equal(hooks.incomeStatementOperatingBridgeTies("FY25", 100, 99), true, "a $1mm annual SEC display-rounding difference is acceptable");
assert.equal(hooks.incomeStatementOperatingBridgeTies("FY25", 100.1, 99), false, "an annual bridge difference above $1mm remains blocking");
assert.equal(hooks.incomeStatementOperatingBridgeTies("4Q25", 100, 99), true, "a $1mm quarterly SEC display-rounding difference is acceptable");
assert.equal(hooks.incomeStatementOperatingBridgeTies("4Q25", 100.1, 99), false, "a quarterly bridge difference above $1mm remains blocking");
assert.equal(hooks.segmentMetric("us-gaap:Depreciation"), "depreciationAmortization", "an exact SEC Depreciation segment tag must map to segment D&A");

function statementRow({ rowOrder, label, concept, value, accession, start, end }) {
  return {
    statementName: "Consolidated Statements of Operations",
    sourceTableType: "primary_statement",
    rowLabel: label,
    xbrlConcept: concept,
    taxonomy: "us-gaap",
    value,
    unit: "USD",
    period: { start, end, periodType: "duration" },
    consolidated: true,
    dimensions: [],
    rowOrder,
    accession,
    reportingPeriod: end
  };
}

function fact(row, period) {
  return {
    concept: row.xbrlConcept,
    label: row.rowLabel,
    value: row.value,
    unit: "USD",
    taxonomy: "us-gaap",
    sourceLayer: "sec_filing_package",
    accn: row.accession,
    start: row.period.start,
    end: row.period.end,
    periodKey: period,
    periodType: period.startsWith("FY") ? "annual" : "quarterly",
    reportDate: row.reportingPeriod
  };
}

function contextForStatements(definitions) {
  const duration = new Map();
  const statements = [];
  const entries = [];
  for (const definition of definitions) {
    duration.set(definition.period, new Map(definition.rows.map((row) => [row.xbrlConcept, fact(row, definition.period)])));
    statements.push({
      statementName: "Consolidated Statements of Operations",
      sourceTableType: "primary_statement",
      accession: definition.accession,
      reportingPeriod: definition.end,
      form: definition.form,
      filingDate: definition.filingDate,
      rows: definition.rows
    });
    entries.push({
      accessionNumber: definition.accession,
      accessionKey: definition.accession,
      form: definition.form,
      filingDate: definition.filingDate,
      reportDate: definition.end,
      fiscalYear: Number(`20${definition.period.slice(-2)}`),
      fiscalQuarter: definition.period.startsWith("FY") ? 4 : Number(definition.period[0]),
      quarterPeriod: definition.period.startsWith("FY") ? `4Q${definition.period.slice(-2)}` : definition.period,
      annualPeriod: definition.period.startsWith("FY") ? definition.period : `FY${definition.period.slice(-2)}`
    });
  }
  return {
    duration,
    instant: new Map(),
    filingPackageStatements: statements,
    fiscalPeriods: {
      entries,
      byAccession: new Map(entries.map((entry) => [entry.accessionKey, entry])),
      byReportDate: new Map(entries.map((entry) => [entry.reportDate, entry])),
      reportedPeriods: new Set(entries.map((entry) => entry.quarterPeriod)),
      fiscalYearEndMonth: 12,
      fiscalYearEndDay: 31
    }
  };
}

const accession = "000159653225000216";
const period = "2Q25";
const start = "2025-04-01";
const end = "2025-06-30";
const rows = [
  [1, "Revenue", "RevenueFromContractWithCustomerExcludingAssessedTax", 2_000_000_000],
  [2, "Cost of product and service sold", "CostOfGoodsAndServicesSold", 500_000_000],
  [3, "Research and development", "ResearchAndDevelopmentExpense", 200_000_000],
  [4, "Sales and marketing", "SellingAndMarketingExpense", 93_492_000],
  [5, "General and administrative", "GeneralAndAdministrativeExpense", 25_029_000],
  [6, "Operating income", "OperatingIncomeLoss", 986_200_000],
  [7, "Nonoperating income (expense)", "NonoperatingIncomeExpense", 94_000_000],
  [8, "Income before income taxes", "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest", 1_080_200_000],
  [9, "Income tax expense", "IncomeTaxExpenseBenefit", 191_400_000],
  [10, "Net income", "NetIncomeLoss", 888_800_000]
].map(([rowOrder, label, concept, value]) => statementRow({ rowOrder, label, concept, value, accession, start, end }));
const ctx = contextForStatements([{ period, accession, start, end, form: "10-Q", filingDate: "2025-08-07", rows }]);
const statementCategories = new Set(rows.map((row) => hooks.reportedLineItemCategory(fact(row, period))));
assert.ok(statementCategories.has("revenue") && statementCategories.has("pretax_income") && statementCategories.has("net_income"), JSON.stringify([...statementCategories]));

const nonoperatingSource = fact(rows[6], period);
assert.equal(
  hooks.incomeStatementSourceIsSubtotalOrTotalToExclude(nonoperatingSource, "below operating income"),
  false,
  "NonoperatingIncomeExpense must not be mistaken for an operating-income subtotal by substring matching"
);
const absentInterestExpense = hooks.resolveInterestExpense(period, ctx);
assert.equal(absentInterestExpense.value, 0);
assert.equal(absentInterestExpense.sources[0].sourceLayer, "derived");
assert.equal(absentInterestExpense.sources[0].accn, accession, "a presentation-backed zero must cite the same-period SEC filing");
assert.match(absentInterestExpense.sources[0].concept, /PresentationAbsence$/);
const absentInterestIncome = hooks.resolveInterestIncome(period, ctx);
assert.equal(absentInterestIncome.value, 0, absentInterestIncome.note);
assert.equal(hooks.resolveOtherNonOperatingIncomeExpense(period, ctx).value, 94_000_000);

const workbook = new ExcelJS.Workbook();
const sheet = workbook.addWorksheet("Model");
sheet.getCell(27, 3).value = "Income Statement";
[
  [28, "Revenue"],
  [29, "Cost of Goods Sold"],
  [32, "Selling, General & Administration (SG&A)"],
  [33, "Research & Development (R&D)"],
  [34, "Depreciation & Amortization"],
  [35, "Other Operating Income (Expense)"],
  [36, "EBIT"],
  [38, "Interest Income"],
  [39, "Interest (Expense)"],
  [40, "Goodwill Impairment"],
  [41, "Other Non-Operating Income (Expense)"],
  [42, "Pre-Tax Income (Loss)"]
].forEach(([row, label]) => sheet.getCell(row, 3).value = label);
sheet.getCell("Q28").value = 2_000;
sheet.getCell("Q29").value = -1;
sheet.getCell("Q32").value = -25.029;
sheet.getCell("Q33").value = -1;
sheet.getCell("Q34").value = { formula: "-1477", result: -1477 };
sheet.getCell("Q35").value = 0;
sheet.getCell("Q36").value = { formula: "Q28+SUM(Q29,Q32:Q35)", result: -999 };
sheet.getCell("Q38").value = { formula: "17", result: 17 };
sheet.getCell("Q39").value = { formula: "-219", result: -219 };
sheet.getCell("Q40").value = { formula: "3", result: 3 };
sheet.getCell("Q41").value = { formula: "0", result: 0 };
sheet.getCell("Q42").value = { formula: "Q36+SUM(Q38:Q41)", result: -999 };

const auditRows = [];
hooks.markReportedPeriodColumns(sheet, [{ period, col: 17 }]);
hooks.reconcileIncomeStatementClassificationRowsToEdgar(sheet, [period], [17], ctx, auditRows);
assert.equal(sheet.getCell("Q32").value, -118.521, "grouped SG&A must retain both primary-statement components");
assert.equal(sheet.getCell("Q34").value, 0, "stale D&A formulas must be cleared when no standalone primary-statement D&A exists");
assert.equal(sheet.getCell("Q39").value, 0, "stale interest expense must be cleared when the filing presents only a combined non-operating line");
assert.equal(sheet.getCell("Q41").value, 94, "the combined SEC non-operating line must populate other non-operating income/expense");
const evaluator = new hooks.FormulaEvaluator(sheet, { useCachedFormulaResults: false, allowCachedFormulaResultFallback: false });
assert.equal(evaluator.evaluateCell(sheet.getCell("Q36")), 986.2, "EBIT must reconcile through formula precedents without cache injection");
assert.equal(evaluator.evaluateCell(sheet.getCell("Q42")), 1080.2, "the below-operating bridge must reconcile to reported pre-tax income");

const formulaSheet = workbook.addWorksheet("Formula Bridge");
formulaSheet.getCell("F32").value = -20;
formulaSheet.getCell("G32").value = -30;
formulaSheet.getCell("H32").value = -25;
formulaSheet.getCell("I32").value = { formula: "-115-SUM(F32:H32)", result: -999 };
assert.equal(
  new hooks.FormulaEvaluator(formulaSheet, { useCachedFormulaResults: false, allowCachedFormulaResultFallback: false }).evaluateCell(formulaSheet.getCell("I32")),
  -40
);
formulaSheet.getCell("I32").value = { formula: "-999-SUM(F32:H32)", result: -924 };
const fourthQuarterRefresh = hooks.refreshReportedIncomeFormulaForTarget(formulaSheet.getCell("I32"), "4Q25", -40);
assert.deepEqual(fourthQuarterRefresh, { value: -40, formulaUpdated: true });
assert.equal(formulaSheet.getCell("I32").value.formula, "-115-SUM(F32:H32)");
formulaSheet.getCell("J32").value = { formula: "SUM(F32:I32)", result: -999 };
const annualRefresh = hooks.refreshReportedIncomeFormulaForTarget(formulaSheet.getCell("J32"), "FY25", -115);
assert.deepEqual(annualRefresh, { value: -115, formulaUpdated: false });
assert.equal(formulaSheet.getCell("J32").value.formula, "SUM(F32:I32)");

const daAddbackCtx = {
  duration: new Map([
    [
      period,
      new Map([
        [
          "Depreciation",
          { concept: "Depreciation", label: "Depreciation", value: 1_434_000_000, unit: "USD", accn: accession, start, end, periodKey: period, periodType: "quarterly", sourceLayer: "sec_live_companyfacts" }
        ],
        [
          "AmortizationOfIntangibleAssets",
          { concept: "AmortizationOfIntangibleAssets", label: "Amortization of intangible assets", value: 1_682_000_000, unit: "USD", accn: accession, start, end, periodKey: period, periodType: "quarterly", sourceLayer: "sec_live_companyfacts" }
        ]
      ])
    ]
  ]),
  instant: new Map(),
  filingPackageStatements: [],
  fiscalPeriods: ctx.fiscalPeriods
};
const separateDaAddback = hooks.resolveEbitdaDepreciationAmortizationAddback(period, daAddbackCtx);
assert.equal(separateDaAddback.value, 3_116_000_000, "separately reported depreciation and amortization must both be included in EBITDA D&A");
assert.equal(separateDaAddback.sources[0].concept, "EbitdaDepreciationAndAmortizationCombined");

const daSheet = workbook.addWorksheet("D&A Addback");
daSheet.getCell("C9").value = "Income Statement";
daSheet.getCell("C10").value = "EBIT";
daSheet.getCell("C11").value = "Depreciation & Amortization";
daSheet.getCell("C12").value = "EBITDA";
daSheet.getCell("F5").value = -1_682;
daSheet.getCell("F11").value = { formula: "-F5", result: 1_682 };
const daAuditRows = [];
const daWrite = hooks.writeHistoricalEbitdaDaAddback(daSheet, [period], [6], daAddbackCtx, daAuditRows);
assert.equal(daWrite.filledCells, 1);
assert.equal(daSheet.getCell("F11").value, 3_116, "an incomplete historical D&A link must be replaced by the complete SEC addback");
assert.equal(daAuditRows[0].formulaPreserved, false);

const ebitdaBridgeSheet = workbook.addWorksheet("EBITDA Bridge");
ebitdaBridgeSheet.getCell("C25").value = "Income Statement";
ebitdaBridgeSheet.getCell("C36").value = "EBIT";
ebitdaBridgeSheet.getCell("C59").value = "EBIT";
ebitdaBridgeSheet.getCell("C60").value = "Depreciation & Amortization";
ebitdaBridgeSheet.getCell("C61").value = "EBITDA";
ebitdaBridgeSheet.getCell("F36").value = 8_053;
ebitdaBridgeSheet.getCell("F59").value = { formula: "F51-F44-F41-F40-F39-F38", result: 11_714 };
ebitdaBridgeSheet.getCell("F60").value = -1;
ebitdaBridgeSheet.getCell("F61").value = { formula: "SUM(F59:F60)", result: 11_713 };
const ebitdaBridgeCtx = {
  ...daAddbackCtx,
  duration: new Map([
    [
      period,
      new Map([
        ...daAddbackCtx.duration.get(period).entries(),
        [
          "OperatingIncomeLoss",
          { concept: "OperatingIncomeLoss", label: "Operating income", value: 8_053_000_000, unit: "USD", accn: accession, start, end, periodKey: period, periodType: "quarterly", sourceLayer: "sec_live_companyfacts" }
        ]
      ])
    ]
  ])
};
const ebitdaBridgeWrite = hooks.writeHistoricalEbitdaDaAddback(ebitdaBridgeSheet, [period], [6], ebitdaBridgeCtx, []);
assert.equal(ebitdaBridgeWrite.filledCells, 3);
assert.deepEqual(ebitdaBridgeSheet.getCell("F59").value, { formula: "F36", result: 8_053 });
assert.deepEqual(ebitdaBridgeSheet.getCell("F61").value, { formula: "SUM(F59:F60)", result: 11_169 });

const preferred = hooks.preferredIncomeStatementResolverOverNarrowerAssignment(period, ctx, "Selling, General & Administration (SG&A)", [
  {
    fiscalPeriod: period,
    sourceFilingAccession: accession,
    sourceStatement: "Consolidated Statements of Operations",
    sourceLineItemLabel: "General and administrative",
    sourceAmount: 25_029_000,
    modelAmount: -25_029_000,
    sourceXbrlTag: "GeneralAndAdministrativeExpense",
    assignedModelRow: "Selling, General & Administration (SG&A)",
    assignmentStatus: "mapped_to_model_row",
    classificationReason: "component",
    llmUsed: false,
    validationStatus: "OK!",
    sourceSection: "operating expenses",
    sourceRowKey: "g-and-a"
  }
]);
assert.equal(preferred.value, -118_521_000, "a narrower assignment ledger must not overwrite a broader SEC-backed SG&A grouping");

const combinedNonOperatingRows = [
  [1, "Revenue", "RevenueFromContractWithCustomerExcludingAssessedTax", 10_000_000_000],
  [2, "Cost of revenue", "CostOfGoodsAndServicesSold", 6_000_000_000],
  [3, "Operating income", "OperatingIncomeLoss", 4_000_000_000],
  [4, "Other income/(expense), net", "NonoperatingIncomeExpense", -393_000_000],
  [5, "Income before taxes", "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest", 3_607_000_000],
  [6, "Income tax expense", "IncomeTaxExpenseBenefit", 600_000_000],
  [7, "Net income", "NetIncomeLoss", 3_007_000_000]
].map(([rowOrder, label, concept, value]) => statementRow({ rowOrder, label, concept, value, accession, start, end }));
const combinedNonOperatingCtx = contextForStatements([
  { period, accession, start, end, form: "10-Q", filingDate: end, rows: combinedNonOperatingRows }
]);
combinedNonOperatingCtx.duration.get(period).set("InterestExpense", {
  concept: "InterestExpense",
  label: "Interest expense disclosed outside the primary statement",
  value: 1_003_000_000,
  unit: "USD",
  accn: accession,
  start,
  end,
  periodKey: period,
  periodType: "quarterly",
  sourceLayer: "sec_live_companyfacts"
});
const splitReportedLinePreferred = hooks.preferredIncomeStatementResolverOverNarrowerAssignment(
  period,
  combinedNonOperatingCtx,
  "Other Non-Operating Income (Expense)",
  [
    {
      fiscalPeriod: period,
      sourceFilingAccession: accession,
      sourceStatement: "Consolidated Statements of Operations",
      sourceLineItemLabel: "Other income/(expense), net",
      sourceAmount: -393_000_000,
      modelAmount: -393_000_000,
      sourceXbrlTag: "NonoperatingIncomeExpense",
      assignedModelRow: "Other Non-Operating Income (Expense)",
      assignmentStatus: "mapped_to_model_row",
      classificationReason: "primary combined non-operating line",
      llmUsed: false,
      validationStatus: "OK!",
      sourceSection: "below operating income",
      sourceRowKey: "combined-other"
    }
  ]
);
assert.equal(
  splitReportedLinePreferred.value,
  610_000_000,
  "a combined primary non-operating line must be split after separately sourced interest expense instead of double-counting interest"
);

function annualRows(period, accession, start, end, values) {
  return [
    [1, "Revenue", "RevenueFromContractWithCustomerExcludingAssessedTax", values.revenue],
    [2, "Cost of revenue", "CostOfGoodsAndServicesSold", values.cogs],
    [3, "Research and development", "ResearchAndDevelopmentExpense", values.rd],
    [4, "Sales and marketing", "SellingAndMarketingExpense", values.sales],
    [5, "General and administrative", "GeneralAndAdministrativeExpense", values.ga],
    [6, "Operating income", "OperatingIncomeLoss", values.operating],
    [7, "Income before taxes", "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest", values.pretax],
    [8, "Income tax expense", "IncomeTaxExpenseBenefit", values.tax],
    [9, "Net income", "NetIncomeLoss", values.net]
  ].map(([rowOrder, label, concept, value]) => statementRow({ rowOrder, label, concept, value, accession, start, end }));
}

const quarterDefinitions = [
  ["1Q26", "0001", "2026-01-01", "2026-03-31", { revenue: 100, cogs: 50, rd: 5, sales: 10, ga: 5, operating: 30, pretax: 32, tax: 6, net: 26 }],
  ["2Q26", "0002", "2026-04-01", "2026-06-30", { revenue: 110, cogs: 55, rd: 6, sales: 10, ga: 5, operating: 34, pretax: 36, tax: 7, net: 29 }],
  ["3Q26", "0003", "2026-07-01", "2026-09-30", { revenue: 120, cogs: 60, rd: 6, sales: 10, ga: 5, operating: 39, pretax: 41, tax: 7, net: 34 }],
  ["FY26", "0004", "2026-01-01", "2026-12-31", { revenue: 460, cogs: 230, rd: 23, sales: 46, ga: 24, operating: 137, pretax: 147, tax: 27, net: 120 }]
].map(([fiscalPeriod, filingAccession, periodStart, periodEnd, values]) => ({
  period: fiscalPeriod,
  accession: filingAccession,
  start: periodStart,
  end: periodEnd,
  form: fiscalPeriod.startsWith("FY") ? "10-K" : "10-Q",
  filingDate: periodEnd,
  rows: annualRows(fiscalPeriod, filingAccession, periodStart, periodEnd, values)
}));
const annualCtx = contextForStatements(quarterDefinitions);
const impairmentDefinitions = quarterDefinitions.map((definition) => {
  if (definition.period !== "FY26") return definition;
  const impairmentRow = statementRow({
    rowOrder: 6.5,
    label: "Goodwill impairment",
    concept: "GoodwillImpairmentLoss",
    value: 40,
    accession: definition.accession,
    start: definition.start,
    end: definition.end
  });
  return { ...definition, rows: [...definition.rows, impairmentRow].sort((a, b) => a.rowOrder - b.rowOrder) };
});
const impairmentCtx = contextForStatements(impairmentDefinitions);
const fourthQuarterImpairment = hooks.resolveGoodwillImpairment("4Q26", impairmentCtx);
assert.equal(fourthQuarterImpairment.value, -40, "an annual-only goodwill impairment must be allocated to Q4, not erased by an absence zero");
assert.equal(fourthQuarterImpairment.sources[0].concept, "GoodwillImpairmentFourthQuarterBridge");

const incompleteImpairmentDefinitions = impairmentDefinitions.map((definition) =>
  definition.period === "2Q26"
    ? { ...definition, rows: definition.rows.filter((row) => row.xbrlConcept !== "NetIncomeLoss") }
    : definition
);
assert.equal(
  hooks.resolveGoodwillImpairment("4Q26", contextForStatements(incompleteImpairmentDefinitions)).value,
  null,
  "an annual-only goodwill impairment must remain unresolved when a Q1-Q3 presentation is incomplete"
);
const fillRows = [
  [28, "Revenue"],
  [29, "Cost of Goods Sold"],
  [32, "Selling, General & Administration (SG&A)"],
  [33, "Research & Development (R&D)"],
  [41, "Other Non-Operating Income (Expense)"],
  [44, "Income Tax Benefit (Expense)"]
].map(([row, label]) => ({ row, label, statement: "income", kind: "duration", classification: "grouped", scale: 1_000_000 }));
const ledger = hooks.buildPrimaryIncomeStatementAssignmentLedgerRows(["4Q26", "FY26"], annualCtx, fillRows);
const fourthQuarterSgaRows = ledger.filter((row) => row.fiscalPeriod === "4Q26" && /general.*administration|sg&a/i.test(row.assignedModelRow));
assert.equal(fourthQuarterSgaRows.length, 2, "4Q assignment coverage must derive both SG&A components from FY less Q1-Q3");
assert.equal(fourthQuarterSgaRows.reduce((sum, row) => sum + row.modelAmount, 0), -25, "derived 4Q SG&A must equal FY less Q1-Q3");
assert.ok(fourthQuarterSgaRows.every((row) => /derived fourth quarter/i.test(row.sourceLineItemLabel)));

console.log("ANET income-statement reconciliation regression passed.");
