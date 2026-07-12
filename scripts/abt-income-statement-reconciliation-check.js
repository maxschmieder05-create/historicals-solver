const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const Module = require("node:module");
const ts = require("typescript");

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

function statementRow({ rowOrder, label, concept, value, unit = "USD", accession, start, end }) {
  return {
    statementName: "Consolidated Statement of Earnings",
    sourceTableType: "primary_statement",
    rowLabel: label,
    xbrlConcept: concept,
    taxonomy: "us-gaap",
    value,
    unit,
    period: { start, end, periodType: "duration" },
    consolidated: true,
    dimensions: [],
    rowOrder,
    accession,
    reportingPeriod: end
  };
}

function statementDefinition({ period, accession, start, end, values, form }) {
  const rows = [
    [1, "Net sales", "RevenueFromContractWithCustomerExcludingAssessedTax", values.revenue],
    [2, "Cost of products sold, excluding amortization of intangible assets", "CostOfGoodsAndServicesSold", values.cogs],
    [3, "Amortization of intangible assets", "AmortizationOfIntangibleAssets", values.amortization],
    [4, "Research and development", "ResearchAndDevelopmentExpense", values.rd],
    [5, "Selling, general, and administrative", "SellingGeneralAndAdministrativeExpense", values.sga],
    [6, "Interest expense", "InterestExpenseNonoperating", values.interestExpense],
    [7, "Interest (income)", "InvestmentIncomeInterest", values.interestIncome],
    [8, "Net foreign exchange (gain) loss", "ForeignCurrencyTransactionGainLossBeforeTax", values.fx],
    [9, "Other (income) expense, net", "OtherNonoperatingIncomeExpense", values.other],
    [10, "Earnings before taxes", "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest", values.pretax],
    [11, "Taxes on earnings", "IncomeTaxExpenseBenefit", values.tax],
    [12, "Net earnings", "NetIncomeLoss", values.netIncome]
  ].map(([rowOrder, label, concept, value]) => statementRow({ rowOrder, label, concept, value, accession, start, end }));
  rows.push(
    statementRow({
      rowOrder: 13,
      label: "Weighted average common shares outstanding - basic",
      concept: "WeightedAverageNumberOfSharesOutstandingBasic",
      value: values.weightedAverageShares,
      unit: "shares",
      accession,
      start,
      end
    })
  );
  return { period, accession, start, end, form, rows };
}

function fact(row, period) {
  return {
    concept: row.xbrlConcept,
    label: row.rowLabel,
    value: row.value,
    unit: row.unit,
    taxonomy: row.taxonomy,
    sourceLayer: "sec_live_companyfacts",
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
    const facts = new Map(definition.rows.map((row) => [row.xbrlConcept, fact(row, definition.period)]));
    // Adversarial companyfacts candidate: same concept/accession/period, but a note/detail amount that
    // does not equal the primary-statement line. Primary-statement aggregation must ignore it.
    const reportedOther = facts.get("OtherNonoperatingIncomeExpense");
    if (reportedOther) facts.set("OtherNonoperatingIncomeExpense", { ...reportedOther, value: reportedOther.value + 1_011_000_000 });
    const revenue = facts.get("RevenueFromContractWithCustomerExcludingAssessedTax");
    if (revenue) {
      facts.set("OperatingIncomeLoss", {
        ...revenue,
        concept: "OperatingIncomeLoss",
        label: "Operating income from an unrelated note or segment presentation",
        value: -900_000_000
      });
    }
    duration.set(definition.period, facts);
    statements.push({
      statementName: "Consolidated Statement of Earnings",
      sourceTableType: "primary_statement",
      accession: definition.accession,
      reportingPeriod: definition.end,
      form: definition.form,
      filingDate: definition.end,
      rows: definition.rows
    });
    const year = Number(definition.end.slice(0, 4));
    entries.push({
      accessionNumber: definition.accession,
      accessionKey: definition.accession,
      form: definition.form,
      filingDate: definition.end,
      reportDate: definition.end,
      fiscalYear: year,
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
      reportedPeriods: new Set(entries.flatMap((entry) => [entry.quarterPeriod, entry.annualPeriod])),
      fiscalYearEndMonth: 12,
      fiscalYearEndDay: 31
    }
  };
}

const definitions = [
  statementDefinition({
    period: "1Q25",
    accession: "000000180025000001",
    start: "2025-01-01",
    end: "2025-03-31",
    form: "10-Q",
    values: { revenue: 10_358e6, cogs: 4_468e6, amortization: 420e6, rd: 716e6, sga: 3_061e6, interestExpense: 131e6, interestIncome: 82e6, fx: 7e6, other: 127e6, pretax: 1_778e6, tax: 453e6, netIncome: 1_325e6, weightedAverageShares: 1_740e6 }
  }),
  statementDefinition({
    period: "2Q25",
    accession: "000000180025000002",
    start: "2025-04-01",
    end: "2025-06-30",
    form: "10-Q",
    values: { revenue: 11_142e6, cogs: 4_854e6, amortization: 420e6, rd: 725e6, sga: 3_091e6, interestExpense: 121e6, interestIncome: 71e6, fx: 11e6, other: 137e6, pretax: 2_150e6, tax: 371e6, netIncome: 1_779e6, weightedAverageShares: 1_742e6 }
  }),
  statementDefinition({
    period: "3Q25",
    accession: "000000180025000003",
    start: "2025-07-01",
    end: "2025-09-30",
    form: "10-Q",
    values: { revenue: 11_369e6, cogs: 5_075e6, amortization: 420e6, rd: 766e6, sga: 3_051e6, interestExpense: 121e6, interestIncome: 77e6, fx: 17e6, other: 150e6, pretax: 2_180e6, tax: 536e6, netIncome: 1_644e6, weightedAverageShares: 1_745e6 }
  }),
  statementDefinition({
    period: "FY25",
    accession: "000000180026000004",
    start: "2025-01-01",
    end: "2025-12-31",
    form: "10-K",
    values: { revenue: 44_328e6, cogs: 19_319e6, amortization: 1_682e6, rd: 2_942e6, sga: 12_332e6, interestExpense: 493e6, interestIncome: 308e6, fx: 48e6, other: 550e6, pretax: 8_466e6, tax: 1_942e6, netIncome: 6_524e6, weightedAverageShares: 1_743e6 }
  }),
  statementDefinition({
    period: "1Q26",
    accession: "000000180026000005",
    start: "2026-01-01",
    end: "2026-03-31",
    form: "10-Q",
    values: { revenue: 11_164e6, cogs: 4_890e6, amortization: 422e6, rd: 767e6, sga: 3_740e6, interestExpense: 174e6, interestIncome: 106e6, fx: 13e6, other: 159e6, pretax: 1_449e6, tax: 372e6, netIncome: 1_077e6, weightedAverageShares: 1_750e6 }
  })
];

const ctx = contextForStatements(definitions);
const fyOperatingIncome = hooks.resolveOperatingIncome("FY25", ctx);
assert.equal(fyOperatingIncome.value, 8_053e6, "FY25 EBIT must derive from the complete primary operating components");
assert.equal(fyOperatingIncome.sources[0].concept, "OperatingIncomeDerivedFromPrimaryStatementComponents");
assert.doesNotMatch(fyOperatingIncome.note, /pre-tax income less/i);
assert.notEqual(
  fyOperatingIncome.value,
  -900e6,
  "a same-period Companyfacts OperatingIncomeLoss that is absent from the consolidated primary statement must not override the statement component bridge"
);
assert.equal(hooks.resolveOperatingIncome("1Q26", ctx).value, 1_345e6, "1Q26 EBIT must derive from the complete primary operating components");

assert.equal(
  hooks.resolveOtherNonOperatingIncomeExpense("FY25", ctx).value,
  598e6,
  "Other non-operating aggregation must use the primary-statement row amounts rather than a mismatched same-concept companyfact"
);
assert.equal(hooks.resolveOtherNonOperatingIncomeExpense("1Q26", ctx).value, 172e6);
const fourthQuarterOther = hooks.resolveOtherNonOperatingIncomeExpense("4Q25", ctx);
assert.equal(fourthQuarterOther.value, 149e6, "4Q other non-operating must retain every annual component through an annual-minus-Q1-Q3 bridge");
assert.equal(fourthQuarterOther.sources[0].concept, "OtherNonOperatingIncomeExpenseFourthQuarterBridge");

const modelRows = [
  "Revenue",
  "COGS / Cost of Goods Sold",
  "SG&A",
  "R&D",
  "D&A",
  "Other Operating Income / Expense",
  "Interest Income",
  "Interest Expense",
  "Goodwill Impairment",
  "Other Non-Operating Income / Expense",
  "Income Tax Benefit / Expense"
].map((label, index) => ({ row: index + 1, label, classification: "direct", statement: "income", kind: "duration", scale: 1_000_000 }));
const derivedAssignmentRows = hooks.deriveFourthQuarterPrimaryIncomeStatementAssignmentLedgerRows("4Q25", ctx, modelRows, []);
assert.ok(derivedAssignmentRows.length > 0, "additive primary-statement assignments should still derive fourth-quarter rows");
assert.equal(
  derivedAssignmentRows.some((row) => /weighted average|shares outstanding/i.test(`${row.sourceLineItemLabel} ${row.sourceXbrlTag}`)),
  false,
  "non-additive average-share disclosures must never be annual-minus-quarterly derived"
);
assert.equal(
  hooks.incomeStatementAssignmentRowIsQuarterDerivable({
    assignmentStatus: "explicitly_excluded_with_reason",
    assignedModelRow: "",
    sourceLineItemLabel: "Weighted average common shares outstanding - basic",
    sourceXbrlTag: "WeightedAverageNumberOfSharesOutstandingBasic",
    classificationReason: "Excluded non-USD per-share/share-count presentation row."
  }),
  false
);

console.log("ABT-style primary-component operating-income reconciliation checks passed.");
