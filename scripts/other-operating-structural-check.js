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

const accession = "000000000024000001";
const reportDate = "2023-12-31";

function statementRow(rowOrder, rowLabel, xbrlConcept, value) {
  return {
    statementName: "Consolidated Statements of Income",
    sourceTableType: "primary_statement",
    rowLabel,
    xbrlConcept,
    taxonomy: "us-gaap",
    value,
    unit: "USD",
    period: {
      start: "2023-01-01",
      end: reportDate,
      periodType: "duration"
    },
    consolidated: true,
    dimensions: [],
    rowOrder,
    accession,
    reportingPeriod: reportDate
  };
}

const rows = [
  statementRow(1, "Revenue", "RevenueFromContractWithCustomerExcludingAssessedTax", 14_964_500_000),
  statementRow(2, "Cost of operations", "CostOfGoodsAndServicesSold", 8_942_200_000),
  statementRow(3, "Depreciation, amortization and depletion", "DepreciationDepletionAndAmortization", 1_501_400_000),
  statementRow(4, "Selling, general and administrative", "SellingGeneralAndAdministrativeExpense", 1_608_700_000),
  statementRow(5, "Accretion", "AssetRetirementObligationAccretionExpense", 97_900_000),
  statementRow(6, "Adjustment to withdrawal liability for multiemployer pension funds", "PensionAndOtherPostretirementBenefitExpense", 4_500_000),
  statementRow(7, "Gain on business divestitures and impairments, net", "GainLossOnDispositionOfAssetsAndImpairmentsNet", 3_600_000),
  statementRow(8, "Restructuring charges", "RestructuringCharges", 33_200_000),
  statementRow(9, "Operating income", "OperatingIncomeLoss", 2_780_200_000),
  statementRow(10, "Interest expense", "InterestExpense", 508_200_000),
  statementRow(11, "Income before income taxes", "IncomeLossFromContinuingOperationsBeforeIncomeTaxes", 2_191_500_000)
];

const leaseRows = [
  statementRow(1, "Fixed lease cost", "FixedLeaseCost", 60_200_000),
  statementRow(2, "Short-term lease cost", "ShortTermLeaseCost", 89_500_000),
  statementRow(3, "Variable lease cost", "OperatingLeaseVariableLeaseCost", 26_800_000),
  statementRow(4, "Interest on lease liabilities", "FinanceLeaseInterestExpense", 8_800_000)
];

function quarterStatement(period, start, end, quarterAccession, accretion, restructuring, gain, operatingIncome, interestExpense) {
  const quarterRows = [
    [1, "Revenue", "RevenueFromContractWithCustomerExcludingAssessedTax", 1_000_000_000],
    [2, "Accretion", "AssetRetirementObligationAccretionExpense", accretion],
    ...(gain ? [[3, "Gain on business divestitures and impairments, net", "GainLossOnDispositionOfAssetsAndImpairmentsNet", gain]] : []),
    [4, "Restructuring charges", "RestructuringCharges", restructuring],
    [5, "Operating income", "OperatingIncomeLoss", operatingIncome],
    [6, "Interest expense", "InterestExpense", interestExpense]
  ].map(([rowOrder, rowLabel, xbrlConcept, value]) => ({
    statementName: "Consolidated Statements of Income",
    sourceTableType: "primary_statement",
    rowLabel,
    xbrlConcept,
    taxonomy: "us-gaap",
    value,
    unit: "USD",
    period: { start, end, periodType: "duration" },
    consolidated: true,
    dimensions: [],
    rowOrder,
    accession: quarterAccession,
    reportingPeriod: end
  }));
  const quarterNumber = Number(period[0]);
  const entry = {
    accessionNumber: quarterAccession,
    accessionKey: quarterAccession,
    form: "10-Q",
    filingDate: end,
    reportDate: end,
    fiscalYear: 2023,
    fiscalQuarter: quarterNumber,
    quarterPeriod: period
  };
  return { period, rows: quarterRows, entry };
}

const quarterStatements = [
  quarterStatement("1Q23", "2023-01-01", "2023-03-31", "000000000023000011", 24_100_000, 5_500_000, 0, 644_100_000, 126_700_000),
  quarterStatement("2Q23", "2023-04-01", "2023-06-30", "000000000023000012", 24_500_000, 15_500_000, 0, 707_200_000, 124_400_000),
  quarterStatement("3Q23", "2023-07-01", "2023-09-30", "000000000023000013", 24_600_000, 6_300_000, 1_500_000, 727_800_000, 127_600_000)
];

const filingEntry = {
  accessionNumber: accession,
  accessionKey: accession,
  form: "10-K",
  filingDate: "2024-02-29",
  reportDate,
  fiscalYear: 2023,
  fiscalQuarter: 4,
  quarterPeriod: "4Q23",
  annualPeriod: "FY23"
};

const ctx = {
  duration: new Map([
    ...quarterStatements.map((statement) => [
      statement.period,
      new Map(
        statement.rows.map((row) => [
          row.xbrlConcept,
          {
            concept: row.xbrlConcept,
            label: row.rowLabel,
            value: row.value,
            unit: "USD",
            taxonomy: row.taxonomy,
            sourceLayer: "sec_filing_package",
            accn: row.accession,
            start: row.period.start,
            end: row.period.end,
            periodKey: statement.period,
            periodType: "quarterly",
            reportDate: row.reportingPeriod
          }
        ])
      )
    ]),
    [
      "4Q23",
      new Map([
        [
          "InterestExpense",
          {
            concept: "InterestExpense",
            label: "Interest Expense (derived Q4)",
            value: 129_400_000,
            unit: "USD",
            taxonomy: "us-gaap",
            sourceLayer: "sec_live_companyfacts",
            accn: accession,
            periodKey: "4Q23",
            periodType: "quarterly",
            reportDate
          }
        ],
        [
          "InterestExpenseNonoperating",
          {
            concept: "InterestExpenseNonoperating",
            label: "Interest expense Nonoperating",
            value: 129_200_000,
            unit: "USD",
            taxonomy: "us-gaap",
            sourceLayer: "sec_live_companyfacts",
            accn: accession,
            periodKey: "4Q23",
            periodType: "quarterly",
            reportDate
          }
        ]
      ])
    ],
    [
      "FY23",
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
            start: row.period.start,
            end: reportDate,
            periodKey: "FY23",
            periodType: "annual",
            reportDate
          }
        ])
      )
    ]
  ]),
  instant: new Map(),
  filingPackageStatements: [
    ...quarterStatements.map((statement) => ({
      statementName: "Consolidated Statements of Income",
      sourceTableType: "primary_statement",
      accession: statement.entry.accessionNumber,
      reportingPeriod: statement.entry.reportDate,
      form: "10-Q",
      filingDate: statement.entry.filingDate,
      rows: statement.rows
    })),
    {
      statementName: "Consolidated Statements of Income",
      sourceTableType: "primary_statement",
      accession,
      reportingPeriod: reportDate,
      form: "10-K",
      filingDate: "2024-02-29",
      rows
    },
    {
      statementName: "Consolidated statements of income lease cost detail",
      sourceTableType: "primary_statement",
      accession,
      reportingPeriod: reportDate,
      form: "10-K",
      filingDate: "2024-02-29",
      rows: leaseRows
    }
  ],
  fiscalPeriods: {
    entries: [...quarterStatements.map((statement) => statement.entry), filingEntry],
    byAccession: new Map([...quarterStatements.map((statement) => [statement.entry.accessionKey, statement.entry]), [accession, filingEntry]]),
    byReportDate: new Map([...quarterStatements.map((statement) => [statement.entry.reportDate, statement.entry]), [reportDate, filingEntry]]),
    reportedPeriods: new Set(["1Q23", "2Q23", "3Q23", "4Q23", "FY23"]),
    fiscalYearEndMonth: 12,
    fiscalYearEndDay: 31
  }
};

const structuralSources = hooks.primaryIncomeStatementOtherOperatingLineSources("FY23", ctx);
assert.deepEqual(
  structuralSources.map((source) => source.concept).sort(),
  [
    "AssetRetirementObligationAccretionExpense",
    "GainLossOnDispositionOfAssetsAndImpairmentsNet",
    "PensionAndOtherPostretirementBenefitExpense",
    "RestructuringCharges"
  ].sort()
);

assert.equal(hooks.primaryIncomeStatementRowMatchesPeriod(rows[0], "FY23", ctx), true);
assert.equal(
  hooks.primaryIncomeStatementRowMatchesPeriod(
    {
      ...rows[0],
      value: 13_511_100_000,
      period: { start: "2022-01-01", end: "2022-12-31", periodType: "duration" }
    },
    "FY23",
    ctx
  ),
  false,
  "A comparative annual fact from the same 10-K accession must not be assigned to the filing's current fiscal year."
);

const otherOperating = hooks.resolveOtherOperatingIncomeExpense("FY23", ctx);
assert.equal(otherOperating.value, -132_000_000);
assert.equal(otherOperating.sources.some((source) => source.concept === "PensionAndOtherPostretirementBenefitExpense"), true);
assert.equal(otherOperating.sources.some((source) => source.concept === "GainLossOnDispositionOfAssetsAndImpairmentsNet"), true);
assert.equal(hooks.resolveOtherOperatingIncomeExpense("1Q23", ctx).value, -29_600_000);
assert.equal(hooks.resolveOtherOperatingIncomeExpense("2Q23", ctx).value, -40_000_000);
assert.equal(hooks.resolveOtherOperatingIncomeExpense("3Q23", ctx).value, -29_400_000);
assert.equal(hooks.resolveOtherOperatingIncomeExpense("4Q23", ctx).value, -33_000_000);
assert.equal(
  hooks.resolveOperatingIncome("4Q23", ctx).value,
  701_100_000,
  "Fourth-quarter operating income must use the annual-minus-Q1-Q3 bridge so the annual formula recalculates to the SEC total."
);
assert.equal(
  hooks.resolveInterestExpense("4Q23", ctx).value,
  -129_500_000,
  "Fourth-quarter interest expense must use the annual-minus-Q1-Q3 bridge so the annual formula recalculates to the SEC total."
);
assert.equal(
  hooks.deterministicModelRowCandidateForSource(
    {
      concept: "PensionAndOtherPostretirementBenefitExpense",
      label: "Adjustment to withdrawal liability for multiemployer pension funds",
      value: 4_500_000
    },
    "income_statement",
    "operating expenses"
  ),
  "Other Operating Income / Expense",
  "Primary-statement location must keep an operating pension-withdrawal charge above EBIT even when its concept name is ambiguous."
);
assert.equal(
  hooks.selectPrimaryStatementRevenueSource("FY23", ctx).value,
  14_964_500_000,
  "Annual revenue should come from the matching primary consolidated statement row."
);
assert.equal(
  hooks.selectPrimaryStatementRevenueSource("4Q23", ctx).value,
  11_964_500_000,
  "Fourth-quarter revenue should be derived from the matching annual primary statement less Q1-Q3."
);
assert.equal(
  hooks.completeQuarterlyAnnualValue(
    "FY23",
    new Map([
      ["1Q23", 3_581_100_000],
      ["2Q23", 3_725_900_000],
      ["3Q23", 3_825_900_000],
      ["4Q23", 3_831_600_000]
    ])
  ),
  14_964_500_000
);
assert.equal(hooks.completeQuarterlyAnnualValue("FY23", new Map([["1Q23", 3_581_100_000]])), null);

assert.equal(
  hooks.otherOperatingLineValue({ concept: "GainLossOnDispositionOfAssetsAndImpairmentsNet", label: "Gain on business divestitures and impairments, net", value: 3_600_000 }),
  3_600_000
);
assert.equal(
  hooks.otherOperatingLineValue({ concept: "GainLossOnDispositionOfAssetsAndImpairmentsNet", label: "(Gain) loss on business divestitures and impairments, net", value: 3_600_000 }),
  3_600_000
);
assert.equal(hooks.otherOperatingLineValue({ concept: "AssetDispositionLoss", label: "Loss on disposition of assets", value: 2_000_000 }), -2_000_000);
assert.equal(
  hooks.otherOperatingLineValue({ concept: "RestructuringCharges", label: "Restructuring and other charges", value: -62_000_000 }),
  62_000_000,
  "A negative reported charge is an operating credit/reversal and must not be forced back to an expense."
);

assert.equal(hooks.isNumericConstantFormula("-24.1-5.5"), true);
assert.equal(hooks.isNumericConstantFormula("=(-24.1)+(1.5)-6.3"), true);
assert.equal(hooks.isNumericConstantFormula("SUM(F35:I35)"), false);
assert.equal(hooks.isNumericConstantFormula("F30+F31"), false);
assert.equal(
  hooks.incomeStatementClassificationCellTiesResolvedValue(-90.2, -88.7, -90.2),
  false,
  "A cached SEC value must not hide a numeric formula whose recalculated result is wrong."
);
assert.equal(hooks.incomeStatementClassificationCellTiesResolvedValue(-90.2, -90.2, -90.2), true);

const noncurrentRestrictedCash = {
  concept: "RestrictedCashAndInvestmentsNoncurrent",
  label: "Restricted Cash and Investments, Noncurrent",
  value: 292_000_000
};
assert.equal(hooks.sourceLooksLikeCurrentInvestment(noncurrentRestrictedCash), false);
assert.equal(hooks.sourceLooksLikeCashLikeShortTermInvestment(noncurrentRestrictedCash), false);
assert.equal(hooks.sourceLooksLikeNonCurrentCashBalance(noncurrentRestrictedCash), true);
assert.equal(hooks.sourceLooksLikeCashBalance(noncurrentRestrictedCash), false);
assert.equal(
  hooks.sourceLooksLikeCashLikeShortTermInvestment({ concept: "MarketableSecuritiesCurrent", label: "Marketable securities, current", value: 50_000_000 }),
  true
);

const currentInsuranceReserve = {
  concept: "SelfInsuranceReserveCurrent",
  label: "Less: current portion",
  value: 216_600_000
};
const noncurrentInsuranceReserve = {
  concept: "SelfInsuranceReserveNoncurrent",
  label: "Insurance reserves, net of current portion",
  value: 348_800_000
};
assert.equal(hooks.sourceLooksLikeCurrentSelfInsuranceReserve(currentInsuranceReserve), true);
assert.equal(hooks.sourceLooksLikeCurrentDebtMaturity(currentInsuranceReserve), false);
assert.equal(hooks.sourceLooksLikeNonCurrentSelfInsuranceReserve(noncurrentInsuranceReserve), true);
assert.equal(hooks.sourceLooksLikeNonCurrentDebt(noncurrentInsuranceReserve), false);

console.log("Other-operating structural classification regression passed.");
