const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const ts = require("typescript");
const ExcelJS = require("exceljs");

const repoRoot = path.resolve(__dirname, "..");

require.extensions[".ts"] = function compileTypeScriptModule(module, filename) {
  const source = fs.readFileSync(filename, "utf8");
  const compiled = ts.transpileModule(source, {
    compilerOptions: {
      module: ts.ModuleKind.CommonJS,
      target: ts.ScriptTarget.ES2020,
      esModuleInterop: true
    }
  }).outputText;
  module._compile(compiled, filename);
};

const { sourceLedgerStatusForAuditRow, validateHistoricalSourceLedger, __fillModelServiceTestHooks } = require(
  path.join(repoRoot, "server", "fill-model", "fill-model-service.ts")
);

function auditRow(overrides = {}) {
  return {
    sheetName: "Model",
    cell: "F42",
    modelRowLabel: "Other Current Liabilities",
    period: "1Q23",
    valueWritten: 0,
    mappingType: "direct",
    conceptsUsed: "NoCurrentSecSource=0mm",
    sourceStatement: "balance",
    accession: "",
    sourceUrl: "",
    cellWritable: true,
    formulaPreserved: false,
    formulaStatus: "reported-period value explicitly sourced as zero",
    writeBlockedReason: "",
    signConvention: "explicit zero",
    confidence: "high",
    validationStatus: "OK!",
    notes: "Explicitly set to zero because the current SEC filing has no source for this row after prior filings reported a balance.",
    ...overrides
  };
}

function provenance(overrides = {}) {
  return {
    role: "sec_source",
    concept: "RevenueFromContractWithCustomerExcludingAssessedTax",
    label: "Revenue",
    value: 100,
    sourceLayer: "sec_filing_package",
    accession: "0000000000-23-000001",
    form: "10-Q",
    filedDate: "2023-05-01",
    startDate: "2023-01-01",
    endDate: "2023-03-31",
    periodKey: "1Q23",
    periodType: "quarterly",
    periodEvidence: "source",
    ...overrides
  };
}

function linearDerivationCalculation(terms) {
  return {
    operation: "signed_linear_combination",
    terms: terms.map(({ concept, label = concept, value, coefficient }) => ({ concept, label, value, coefficient }))
  };
}

assert.equal(
  sourceLedgerStatusForAuditRow(
    auditRow({
      formulaStatus: "unsupported reported-period value explicitly zeroed"
    })
  ),
  "stale_or_unsupported",
  "absence of a current SEC source must not be converted into a financial zero"
);

assert.equal(
  sourceLedgerStatusForAuditRow(
    auditRow({
      modelRowLabel: "Deferred Income Taxes",
      conceptsUsed: "NoCurrentSecSource:DeferredIncomeTaxLiabilitiesNet=0mm",
      secLabels: "No current SEC source disclosed",
      validationStatus: "warning: Value is a zero/derived model support value with no direct SEC fact.",
      notes: "No separate deferred tax liability was reported in the SEC balance sheet for this period."
    })
  ),
  "stale_or_unsupported",
  "model-layer no-source markers cannot certify a balance-sheet zero"
);

assert.equal(
  sourceLedgerStatusForAuditRow(
    auditRow({
      modelRowLabel: "Deferred Income Taxes",
      conceptsUsed: "DeferredIncomeTaxLiabilitiesNet=0mm",
      accession: "0000059478-26-000013",
      formulaStatus: "not a formula cell",
      validationStatus: "warning: Value is a zero/derived model support value with no direct SEC fact.",
      notes: ""
    })
  ),
  "explicit_current_sec_source",
  "a zero-valued SEC fact with an accession remains a direct SEC source"
);

assert.equal(
  sourceLedgerStatusForAuditRow(
    auditRow({
      modelRowLabel: "Noncontrolling Interests",
      conceptsUsed: "MinorityInterest=0mm",
      accession: "0000059478-26-000013",
      formulaStatus: "not a formula cell",
      validationStatus: "warning: Value is a zero/derived model support value with no direct SEC fact.",
      notes: ""
    })
  ),
  "explicit_current_sec_source",
  "a zero-valued noncontrolling-interest SEC fact remains a direct SEC source"
);

assert.equal(
  sourceLedgerStatusForAuditRow(
    auditRow({
      modelRowLabel: "Deferred Income Taxes",
      conceptsUsed: "DeferredIncomeTaxLiabilitiesNet=0mm",
      accession: "",
      formulaStatus: "not a formula cell",
      validationStatus: "",
      notes: ""
    })
  ),
  "stale_or_unsupported",
  "a zero-valued concept without an accession is not source support"
);

assert.equal(
  sourceLedgerStatusForAuditRow(
    auditRow({
      modelRowLabel: "Deferred Income Taxes",
      conceptsUsed: "DeferredIncomeTaxLiabilitiesNet=0mm",
      formulaStatus: "not a formula cell",
      validationStatus: "warning: Value is a zero/derived model support value with no direct SEC fact.",
      notes: "No separate deferred tax liability was reported in the SEC balance sheet for this period."
    })
  ),
  "stale_or_unsupported",
  "an explanatory absence note cannot substitute for an SEC-reported zero"
);

assert.equal(
  sourceLedgerStatusForAuditRow(
    auditRow({
      mappingType: "cleared",
      conceptsUsed: "",
      formulaStatus: "unsupported hardcoded historical value cleared to zero",
      validationStatus: "cleared",
      notes: "",
      writeBlockedReason: "Prior hardcoded value 5.9 had no explicit SEC source for this filing period."
    })
  ),
  "stale_or_unsupported",
  "generic stale hardcode clears still need source support and must not masquerade as explicit zeros"
);

assert.equal(
  sourceLedgerStatusForAuditRow(
    auditRow({
      valueWritten: 12.3,
      mappingType: "residual",
      conceptsUsed: "DerivedOutput:OtherCurrentLiabilitiesResidual=12.3mm; LiabilitiesCurrent=100mm; AccountsPayableCurrent=20mm; AccruedLiabilitiesCurrent=67.7mm",
      accession: "0000000000-23-000001",
      sourceProvenance: [
        provenance({ role: "derived_output", concept: "OtherCurrentLiabilitiesResidual", value: 12.3, sourceLayer: "derived", accession: "" }),
        provenance({ concept: "LiabilitiesCurrent", value: 100, periodType: "instant", startDate: "" }),
        provenance({ concept: "AccountsPayableCurrent", value: 20, periodType: "instant", startDate: "" }),
        provenance({ concept: "AccruedLiabilitiesCurrent", value: 67.7, periodType: "instant", startDate: "" })
      ],
      formulaStatus: "historical input refreshed from primary SEC balance-sheet assignment",
      notes: "Residual calculated from SEC current liabilities less SEC-sourced components."
    })
  ),
  "validated_current_company_derived_value"
);

assert.equal(
  sourceLedgerStatusForAuditRow(
    auditRow({
      valueWritten: 12.3,
      conceptsUsed: "OtherLiabilitiesCurrent=12.3mm",
      accession: "0000000000-23-000001",
      formulaStatus: "not a formula cell",
      notes: "Mapped directly to SEC-sourced balance-sheet line item support."
    })
  ),
  "explicit_current_sec_source"
);

const validPresentationAudit = auditRow({
  modelRowLabel: "Interest Income",
  sourceStatement: "income",
  mappingType: "grouped",
  conceptsUsed: "PresentationAbsence:InterestIncomePresentationAbsence=0mm",
  accession: "0000000000-23-000001",
  formulaStatus: "not a formula cell",
  notes: "Interest income was not separately presented on the primary income statement; explicit zero.",
  sourceProvenance: [
    provenance({
      role: "presentation_absence",
      concept: "InterestIncomePresentationAbsence",
      label: "Interest income not separately presented",
      value: 0,
      sourceLayer: "derived"
    })
  ]
});
assert.equal(sourceLedgerStatusForAuditRow(validPresentationAudit), "explicit_zero_no_source_disclosed");
assert.equal(
  sourceLedgerStatusForAuditRow({
    ...validPresentationAudit,
    notes: "",
    sourceProvenance: [
      {
        ...validPresentationAudit.sourceProvenance[0],
        label: "Interest income not separately presented on the primary income statement"
      }
    ]
  }),
  "explicit_zero_no_source_disclosed",
  "the structured presentation-absence label must retain the explanation when zero-valued derived sources are intentionally omitted from analyst-facing mapping comments"
);
assert.equal(
  sourceLedgerStatusForAuditRow({ ...validPresentationAudit, sourceStatement: "balance" }),
  "stale_or_unsupported",
  "an income-statement presentation marker cannot certify a balance-sheet zero"
);
assert.equal(
  sourceLedgerStatusForAuditRow({
    ...validPresentationAudit,
    sourceProvenance: [{ ...validPresentationAudit.sourceProvenance[0], value: 12 }]
  }),
  "stale_or_unsupported",
  "presentation absence must carry a zero-valued structured record"
);
assert.equal(
  sourceLedgerStatusForAuditRow({ ...validPresentationAudit, validationStatus: "blocked unsupported classification" }),
  "stale_or_unsupported",
  "a presentation marker cannot override blocked or unsupported audit status"
);

function ledgerRow(overrides = {}) {
  const row = {
    sheetName: "Model",
    modelRow: 10,
    modelRowLabel: "Revenue",
    modelColumn: "F",
    cell: "F10",
    fiscalPeriod: "1Q23",
    value: 100,
    company: "Example Corp.",
    ticker: "EXM",
    cik: "0000000000",
    accessionNumber: "0000000000-23-000001",
    accessionRaw: "0000000000-23-000001",
    accessionNormalized: "000000000023000001",
    filingFormType: "10-Q",
    reportingPeriodEndDate: "2023-03-31",
    filingPeriod: "2023-03-31",
    sourceStatement: "income",
    sourceTableType: "primary_statement",
    sourceLineItemLabel: "Revenue",
    sourceXbrlTag: "RevenueFromContractWithCustomerExcludingAssessedTax",
    mappingStatus: "explicit_current_sec_source",
    requiredCoreHistoricalInput: true,
    balanceSheetResolverStatus: "",
    workbookFormula: "",
    workbookFormulaPrecedents: "",
    residualFormula: "",
    rawSecValue: "RevenueFromContractWithCustomerExcludingAssessedTax=100mm",
    sourceMappingType: "direct",
    sourceSignConvention: "copied",
    llmUsed: false,
    classificationReason: "Mapped directly to the current SEC filing.",
    ...overrides
  };
  if (!Object.prototype.hasOwnProperty.call(overrides, "sourceProvenance")) {
    const rawAmount = Number(String(row.rawSecValue).match(/=([-+]?\d+(?:\.\d+)?)mm/i)?.[1]);
    row.sourceProvenance = [
      provenance({
        concept: String(row.sourceXbrlTag || "RevenueFromContractWithCustomerExcludingAssessedTax").split(";")[0].trim(),
        label: row.sourceLineItemLabel,
        value: Number.isFinite(rawAmount) ? rawAmount : null,
        accession: row.accessionRaw,
        form: row.filingFormType,
        startDate: row.sourceStatement === "balance" ? "" : "2023-01-01",
        endDate: row.reportingPeriodEndDate,
        periodKey: row.fiscalPeriod,
        periodType: row.sourceStatement === "balance" ? "instant" : row.fiscalPeriod.startsWith("FY") ? "annual" : "quarterly"
      })
    ];
  }
  const derivedOutputs = row.sourceProvenance.filter((source) => source.role === "derived_output");
  row.sourceProvenanceJson = JSON.stringify(row.sourceProvenance);
  row.derivedOutputConcept = derivedOutputs.length === 1 ? derivedOutputs[0].concept : "";
  row.derivedOutputValue = derivedOutputs.length === 1 ? derivedOutputs[0].value : null;
  row.derivedOutputCount = derivedOutputs.length;
  return row;
}

const positionalProvenance = __fillModelServiceTestHooks.sourceProvenanceForResolved(
  {
    value: 60_000_000,
    sources: [
      { concept: "FinalBridge", label: "Final bridge", value: 60_000_000, sourceLayer: "derived", periodKey: "1Q23" },
      { concept: "NestedBridge", label: "Nested bridge", value: 40_000_000, sourceLayer: "derived", periodKey: "1Q23" },
      { concept: "Revenue", label: "Revenue", value: 100_000_000, sourceLayer: "sec_filing_package", accn: "0000000000-23-000001" }
    ]
  },
  1_000_000
);
assert.deepEqual(
  positionalProvenance.map((source) => source.role),
  ["derived_output", "derived_input", "sec_source"],
  "only the first positional synthetic source is the final derived output; later synthetic sources remain nested inputs"
);
assert.equal(positionalProvenance[0].value, 60);

const nestedEquityBridge = {
  concept: "TotalEquityDerivedFromAssetsAndLiabilities",
  label: "Total equity derived from assets less liabilities",
  value: 60,
  sourceLayer: "derived",
  periodKey: "1Q23",
  periodType: "instant",
  derivationCalculation: linearDerivationCalculation([
    { concept: "Assets", value: 100, coefficient: 1 },
    { concept: "Liabilities", value: 40, coefficient: -1 }
  ])
};
const nestedEquityAuditSource = __fillModelServiceTestHooks.resolvedAuditSource(
  "1Q23",
  "StockholdersEquityResolved",
  "Resolved stockholders equity",
  {
    value: 60,
    sources: [
      nestedEquityBridge,
      {
        concept: "Assets",
        label: "Assets",
        value: 100,
        sourceLayer: "sec_filing_package",
        accn: "0000000000-23-000001",
        periodKey: "1Q23",
        periodType: "instant"
      },
      {
        concept: "Liabilities",
        label: "Liabilities",
        value: 40,
        sourceLayer: "sec_filing_package",
        accn: "0000000000-23-000001",
        periodKey: "1Q23",
        periodType: "instant"
      }
    ],
    classification: "residual"
  }
);
const nestedEquityAuditProvenance = __fillModelServiceTestHooks.sourceProvenanceForResolved(
  { value: 60, sources: [nestedEquityAuditSource] },
  1
);
assert.deepEqual(
  nestedEquityAuditProvenance.map((source) => [source.role, source.concept]),
  [
    ["derived_output", "Assets; Liabilities"],
    ["derived_input", "TotalEquityDerivedFromAssetsAndLiabilities"],
    ["sec_source", "Assets"],
    ["sec_source", "Liabilities"]
  ],
  "a summarized balance-sheet total audit source must retain the nested bridge and its SEC inputs instead of emitting a self-certifying synthetic output"
);

const nestedEquityTotalAuditWorkbook = new ExcelJS.Workbook();
const nestedEquityTotalAuditSheet = nestedEquityTotalAuditWorkbook.addWorksheet("Model");
const nestedEquityTotalAuditRow = __fillModelServiceTestHooks.statementTotalAuditRow(
  nestedEquityTotalAuditSheet,
  nestedEquityTotalAuditSheet.getCell("F154"),
  "Total Equity",
  "1Q23",
  60,
  nestedEquityAuditSource,
  "balance",
  "Total Equity maps to resolved SEC balance-sheet sources."
);
assert.deepEqual(
  nestedEquityTotalAuditRow.sourceProvenance.map((source) => [source.role, source.concept]),
  [
    ["derived_output", "Assets; Liabilities"],
    ["derived_input", "TotalEquityDerivedFromAssetsAndLiabilities"],
    ["sec_source", "Assets"],
    ["sec_source", "Liabilities"]
  ],
  "statementTotalAuditRow must expand nested derivation inputs so the Source Ledger preserves the final output, bridge, and direct SEC evidence"
);

const sameAmountYtdEvidenceProvenance = __fillModelServiceTestHooks.sourceProvenanceForResolved(
  {
    value: 221_000_000,
    sources: [
      {
        concept: "ProceedsFromIssuanceOfLongTermDebt",
        label: "Proceeds from issuance of long-term debt (derived Q2)",
        value: 221_000_000,
        sourceLayer: "derived",
        accn: "0000000000-23-000002",
        start: "2023-01-01",
        end: "2023-06-30",
        periodKey: "2Q23",
        periodType: "quarterly",
        derivationCalculation: linearDerivationCalculation([
          { concept: "ProceedsFromIssuanceOfLongTermDebt", value: 221_000_000, coefficient: 1 },
          { concept: "ProceedsFromIssuanceOfLongTermDebt", value: 0, coefficient: -1 }
        ]),
        derivationInputSources: [
          {
            concept: "ProceedsFromIssuanceOfLongTermDebt",
            label: "Proceeds from issuance of long-term debt, six months",
            value: 221_000_000,
            sourceLayer: "sec_filing_package",
            accn: "0000000000-23-000002",
            start: "2023-01-01",
            end: "2023-06-30",
            periodKey: "2Q23",
            periodType: "year_to_date"
          },
          {
            concept: "ProceedsFromIssuanceOfLongTermDebt",
            label: "Proceeds from issuance of long-term debt, first quarter",
            value: 0,
            sourceLayer: "sec_filing_package",
            accn: "0000000000-23-000001",
            start: "2023-01-01",
            end: "2023-03-31",
            periodKey: "1Q23",
            periodType: "quarterly"
          }
        ]
      }
    ]
  },
  1_000_000
);
assert.deepEqual(
  sameAmountYtdEvidenceProvenance.map((source) => source.role),
  ["derived_output", "sec_source", "sec_source"],
  "a derived quarter and same-valued YTD SEC input remain separate provenance records when the prior quarter is zero"
);

const groupedDerivedAggregate = __fillModelServiceTestHooks.resolvedWithFinalAuditDerivation(
  { label: "SG&A", kind: "duration" },
  "4Q23",
  {
    value: -147_000_000,
    classification: "grouped",
    sources: [
      { concept: "SellingAndMarketingExpense", label: "Selling and marketing", value: 105_000_000, sourceLayer: "derived", periodKey: "4Q23" },
      { concept: "GeneralAndAdministrativeExpense", label: "General and administrative", value: 42_000_000, sourceLayer: "derived", periodKey: "4Q23" }
    ]
  }
);
assert.equal(groupedDerivedAggregate.sources[0].value, -147_000_000);
assert.match(groupedDerivedAggregate.sources[0].concept, /DerivedAggregate$/);
assert.equal(groupedDerivedAggregate.sources[0].derivationCalculation.operation, "signed_linear_combination");
assert.equal(
  groupedDerivedAggregate.sources[0].derivationCalculation.terms.reduce(
    (total, term) => total + term.value * term.coefficient,
    0
  ),
  -147_000_000,
  "generated derived outputs must carry arithmetic that replays to the output value"
);
assert.deepEqual(
  __fillModelServiceTestHooks.sourceProvenanceForResolved(groupedDerivedAggregate, 1_000_000).map((source) => source.role),
  ["derived_output", "derived_input", "derived_input"],
  "multiple derived components receive one separate final aggregate output"
);

const modelZeroCalculation = __fillModelServiceTestHooks.derivationCalculationForBridge(
  60_000_000,
  [
    { concept: "Revenue", label: "Revenue", value: 100_000_000, sourceLayer: "sec_filing_package" },
    { concept: "CostOfRevenue", label: "Cost of revenue", value: 40_000_000, sourceLayer: "sec_filing_package" },
    {
      value: 0,
      sources: [{ concept: "PostTaxAdjustmentsNotReported", label: "Post-tax adjustments not reported", value: 0, sourceLayer: "model" }]
    }
  ],
  "signed_linear_combination"
);
assert.deepEqual(
  modelZeroCalculation.terms.map((term) => term.concept),
  ["Revenue", "CostOfRevenue"],
  "zero-valued model placeholders do not become unsupported arithmetic evidence terms"
);

const nestedRepresentativeCalculation = __fillModelServiceTestHooks.derivationCalculationForBridge(
  800_000_000,
  [
    {
      value: 800_000_000,
      sources: [
        {
          concept: "EbitdaDepreciationAndAmortizationCombined",
          label: "Combined D&A",
          value: 800_000_000,
          sourceLayer: "derived",
          derivationCalculation: linearDerivationCalculation([
            { concept: "Depreciation", value: 302_000_000, coefficient: 1 },
            { concept: "AmortizationOfIntangibleAssets", value: 498_000_000, coefficient: 1 }
          ])
        },
        {
          concept: "Depreciation",
          label: "Depreciation derived from YTD",
          value: 302_000_000,
          sourceLayer: "derived",
          derivedTotalValue: 620_000_000,
          derivedPriorPeriods: ["1Q23"]
        }
      ]
    }
  ],
  "signed_linear_combination"
);
assert.equal(
  nestedRepresentativeCalculation.terms[0].concept,
  "EbitdaDepreciationAndAmortizationCombined",
  "a grouped resolver's leading synthetic output represents the grouped value instead of an unrelated nested quarterly bridge"
);

assert.equal(
  __fillModelServiceTestHooks.derivedInputPeriodIsAllowed(
    { fiscalPeriod: "2Q23", sourceStatement: "cash_flow" },
    { concept: "BeginningCashBalanceFromPriorPeriod", label: "Beginning cash balance", periodKey: "2Q23", periodType: "instant" },
    { concept: "CashAndCashEquivalentsAtCarryingValue", label: "Prior ending cash", periodKey: "1Q23", periodType: "instant" }
  ),
  true,
  "beginning-balance roll-forwards allow the immediately prior reported instant before generic duration rules are applied"
);

const commonStockPeriod = "1Q23";
const commonStockFacts = new Map(
  [
    ["CommonStockValue", 24_488_000_000],
    ["StockholdersEquity", 37_010_000_000],
    ["RetainedEarningsAccumulatedDeficit", 35_868_000_000],
    ["TreasuryStockCommonValue", 15_307_000_000],
    ["PreferredStockValue", 0],
    ["AccumulatedOtherComprehensiveIncomeLossNetOfTax", -8_039_000_000]
  ].map(([concept, value]) => [
    concept,
    {
      concept,
      label: concept,
      value,
      sourceLayer: "sec_filing_package",
      accn: "0000000000-23-000001",
      periodKey: commonStockPeriod,
      periodType: "instant",
      end: "2023-03-31",
      unit: "USD"
    }
  ])
);
const commonStockResolution = __fillModelServiceTestHooks.resolveCommonStockAndApic(commonStockPeriod, {
  duration: new Map(),
  instant: new Map([[commonStockPeriod, commonStockFacts]])
});
assert.equal(commonStockResolution.value, 24_488_000_000);
assert.deepEqual(
  commonStockResolution.sources.map((source) => source.concept),
  ["CommonStockValue"],
  "a primary common-stock carrying line that exactly reconciles the equity bridge remains a direct source instead of inheriting residual sources"
);

const presentationProvenance = __fillModelServiceTestHooks.sourceProvenanceForResolved(
  {
    value: 0,
    sources: [
      {
        concept: "InterestIncomePresentationAbsence",
        label: "Interest income not separately presented",
        value: 0,
        sourceLayer: "derived",
        accn: "0000000000-23-000001",
        periodKey: "1Q23",
        periodType: "quarterly"
      }
    ]
  },
  1_000_000
);
assert.equal(presentationProvenance[0].role, "presentation_absence", "presentation absence is not an arithmetic derived output");
const nestedPresentationProvenance = __fillModelServiceTestHooks.sourceProvenanceForResolved(
  {
    value: 0,
    sources: [
      { concept: "FourthQuarterBridge", label: "Fourth-quarter bridge", value: 0, sourceLayer: "derived", periodKey: "4Q23", periodType: "quarterly" },
      {
        concept: "InterestIncomePresentationAbsence",
        label: "Annual interest income not separately presented",
        value: 0,
        sourceLayer: "derived",
        accn: "0000000000-23-000001",
        periodKey: "FY23",
        periodType: "annual"
      }
    ]
  },
  1_000_000
);
assert.deepEqual(nestedPresentationProvenance.map((source) => source.role), ["derived_output", "derived_input"]);

const q4DerivedWrite = {
  value: -25_000_000,
  classification: "grouped",
  sources: [
    {
      concept: "CashFlowCapitalExpendituresFourthQuarterBridge",
      label: "Fourth-quarter capital expenditures",
      value: -25_000_000,
      sourceLayer: "derived",
      periodKey: "4Q23",
      periodType: "quarterly"
    },
    {
      concept: "PaymentsToAcquirePropertyPlantAndEquipment",
      label: "Annual capital expenditures",
      value: 100_000_000,
      sourceLayer: "sec_filing_package",
      cik: "0000000000",
      unit: "USD",
      periodKey: "FY23",
      periodType: "annual"
    },
    ...["1Q23", "2Q23", "3Q23"].map((periodKey) => ({
      concept: "PaymentsToAcquirePropertyPlantAndEquipment",
      label: `${periodKey} capital expenditures`,
      value: 25_000_000,
      sourceLayer: "sec_filing_package",
      cik: "0000000000",
      unit: "USD",
      periodKey,
      periodType: "quarterly"
    }))
  ]
};
const q4WriteValidation = __fillModelServiceTestHooks.validateResolvedValueForWrite(
  { cik: "0000000000", ticker: "EXM", title: "Example Corp." },
  {
    row: 210,
    label: "Capital Expenditures",
    classification: "grouped",
    statement: "support",
    kind: "duration",
    concepts: ["PaymentsToAcquirePropertyPlantAndEquipment"]
  },
  "4Q23",
  q4DerivedWrite
);
assert.notEqual(
  q4WriteValidation.status,
  "blocked",
  "a valid Q4 derived output may trace annual and Q1-Q3 SEC inputs without treating those nested inputs as the workbook period"
);
const wrongFinalPeriodValidation = __fillModelServiceTestHooks.validateResolvedValueForWrite(
  { cik: "0000000000", ticker: "EXM", title: "Example Corp." },
  {
    row: 210,
    label: "Capital Expenditures",
    classification: "grouped",
    statement: "support",
    kind: "duration",
    concepts: ["PaymentsToAcquirePropertyPlantAndEquipment"]
  },
  "4Q23",
  {
    ...q4DerivedWrite,
    sources: [{ ...q4DerivedWrite.sources[0], periodKey: "FY23", periodType: "annual" }, ...q4DerivedWrite.sources.slice(1)]
  }
);
assert.equal(wrongFinalPeriodValidation.status, "blocked", "the final derived output itself cannot use an annual period to certify a Q4 workbook cell");

(async () => {
  const periodEntries = [
    {
      period: "1Q23",
      column: 6,
      modelColumn: "F",
      accessionKey: "000000000023000001",
      accessionNumber: "0000000000-23-000001",
      form: "10-Q",
      filingDate: "2023-05-01",
      periodEndDate: "2023-03-31",
      fiscalYearLabel: "FY23",
      fiscalQuarterLabel: "Q1"
    }
  ];
  const company = { cik: "0000000000", ticker: "EXM", title: "Example Corp." };

  const weightedAverageBasicSharesRow = __fillModelServiceTestHooks.fillRowForContext({
    sheetName: "Model",
    row: 282,
    label: "Weighted Average Basic Shares",
    sectionHeader: "Shares Outstanding Schedule",
    previousLabel: "Ending Balance - Basic",
    nextLabel: "Effects of Dilutive Securities",
    indentation: 0,
    hasHistoricalFormula: true,
    hasHardcodedInput: false,
    hasNetRevenueInterestExpenseAbove: false,
    projectedColumns: 0,
    signConvention: 1
  });
  assert.equal(
    weightedAverageBasicSharesRow.classification,
    "direct",
    "weighted-average shares remain a direct SEC duration fact even inside a shares-outstanding schedule"
  );
  assert.equal(weightedAverageBasicSharesRow.statement, "support");
  assert.equal(weightedAverageBasicSharesRow.kind, "duration");
  assert.ok(
    weightedAverageBasicSharesRow.concepts.some((concept) => /weightedaveragenumberofsharesoutstandingbasic/i.test(concept)),
    "the row must retain the SEC weighted-average basic share concepts"
  );

  const weightedSharesWorkbook = new ExcelJS.Workbook();
  const weightedSharesSheet = weightedSharesWorkbook.addWorksheet("Model");
  weightedSharesSheet.getCell("C282").value = "Weighted Average Basic Shares";
  weightedSharesSheet.getCell("J282").value = { formula: "I282", result: 100 };
  __fillModelServiceTestHooks.markReportedPeriodColumns(weightedSharesSheet, [{ period: "FY23", col: 10 }]);
  assert.equal(
    __fillModelServiceTestHooks.fillRowUsesAllReportedPeriods(weightedSharesSheet, weightedAverageBasicSharesRow),
    true,
    "non-additive weighted-average share facts must be populated for annual as well as quarterly reported periods"
  );
  assert.equal(
    __fillModelServiceTestHooks.isReportedNonAdditiveSupportFormulaInputCell(
      weightedAverageBasicSharesRow,
      weightedSharesSheet.getCell("J282"),
      "FY23",
      { fiscalPeriods: { reportedPeriods: new Set(["FY23"]) } }
    ),
    true,
    "an annual share formula linked to fourth-quarter shares must be recognized as a replaceable reported input"
  );
  assert.equal(
    __fillModelServiceTestHooks.historicalWriteDecision(
      weightedAverageBasicSharesRow,
      weightedSharesSheet.getCell("J282"),
      "FY23",
      { fiscalPeriods: { reportedPeriods: new Set(["FY23"]) } }
    ).writable,
    true,
    "the annual weighted-average share formula must be overwritten with the directly reported SEC annual fact"
  );

  const dilutedSharesContext = {
    sheetName: "Model",
    row: 284,
    label: "Weighted Average Dilutive Shares",
    sectionHeader: "Shares Repurchased",
    previousLabel: "Effects of Dilutive Securities",
    nextLabel: "Debt and Interest Schedule",
    indentation: 0,
    hasHistoricalFormula: false,
    hasHardcodedInput: true,
    hasNetRevenueInterestExpenseAbove: false,
    projectedColumns: 0,
    signConvention: 1
  };
  const dilutedSharesRow = __fillModelServiceTestHooks.fillRowForContext(dilutedSharesContext);
  dilutedSharesRow.modelContext = dilutedSharesContext;
  weightedSharesSheet.getCell("C284").value = "Weighted Average Dilutive Shares";
  weightedSharesSheet.getCell("F284").value = null;
  __fillModelServiceTestHooks.markReportedPeriodColumns(weightedSharesSheet, [
    { period: "1Q23", col: 6 },
    { period: "FY23", col: 10 }
  ]);
  assert.equal(
    __fillModelServiceTestHooks.historicalWriteDecision(
      dilutedSharesRow,
      weightedSharesSheet.getCell("F284"),
      "1Q23",
      { fiscalPeriods: { reportedPeriods: new Set(["1Q23", "FY23"]) } }
    ).writable,
    true,
    "a neighboring debt-schedule header must not turn a cleaned quarterly diluted-share input into an inactive debt helper"
  );

  const freeCashFlowDaContext = {
    sheetName: "Model",
    row: 109,
    label: "Depreciation & Amortization",
    sectionHeader: "Free Cash Flow Analysis",
    previousLabel: "Net Income",
    nextLabel: "Change in Operating Working Capital",
    indentation: 0,
    hasHistoricalFormula: true,
    hasHardcodedInput: false,
    hasNetRevenueInterestExpenseAbove: false,
    projectedColumns: 0,
    signConvention: 1
  };
  const freeCashFlowDaRow = __fillModelServiceTestHooks.fillRowForContext(freeCashFlowDaContext);
  freeCashFlowDaRow.modelContext = freeCashFlowDaContext;
  assert.equal(freeCashFlowDaRow.classification, "direct");
  assert.equal(typeof freeCashFlowDaRow.resolver, "function");
  assert.equal(
    __fillModelServiceTestHooks.fillRowUsesAllReportedPeriods(weightedSharesSheet, freeCashFlowDaRow),
    true,
    "SEC-resolvable free-cash-flow support inputs must populate annual and quarterly reported periods"
  );
  weightedSharesSheet.getCell("C109").value = "Depreciation & Amortization";
  weightedSharesSheet.getCell("J109").value = { formula: "-J194-J208", result: 25 };
  assert.equal(
    __fillModelServiceTestHooks.isReportedSecSupportFormulaInputCell(
      freeCashFlowDaRow,
      weightedSharesSheet.getCell("J109"),
      "FY23",
      { fiscalPeriods: { reportedPeriods: new Set(["FY23"]) } }
    ),
    true,
    "a reported free-cash-flow support formula is writable when its SEC resolver can supply the filing actual"
  );
  assert.equal(
    __fillModelServiceTestHooks.historicalWriteDecision(
      freeCashFlowDaRow,
      weightedSharesSheet.getCell("J109"),
      "FY23",
      { fiscalPeriods: { reportedPeriods: new Set(["FY23"]) } }
    ).writable,
    true,
    "a stale schedule link must not block a directly SEC-backed free-cash-flow historical input"
  );
  assert.equal(__fillModelServiceTestHooks.sourceBackedSupportFormulaTies(-30.8, -30.9), false);
  assert.equal(__fillModelServiceTestHooks.sourceBackedSupportFormulaTies(-30.85, -30.9), true);

  const repurchaseAmountRow = __fillModelServiceTestHooks.fillRowForContext({
    sheetName: "Model",
    row: 241,
    label: "Shares Repurchased ($ Amount)",
    sectionHeader: "Share Repurchase Assumptions",
    previousLabel: "Shares Repurchased (# shares in mm)",
    nextLabel: "Common Stock & APIC Assumptions",
    indentation: 0,
    hasHistoricalFormula: true,
    hasHardcodedInput: true,
    hasNetRevenueInterestExpenseAbove: false,
    projectedColumns: 0,
    signConvention: -1
  });
  assert.equal(repurchaseAmountRow.classification, "direct");
  assert.ok(repurchaseAmountRow.concepts.includes("PaymentsForRepurchaseOfCommonStock"));
  assert.equal(repurchaseAmountRow.sign, 1, "shareholder-schedule repurchase dollars use a positive amount while cash-flow rows use an outflow sign");

  const completenessWorkbook = new ExcelJS.Workbook();
  const completenessSheet = completenessWorkbook.addWorksheet("Model");
  completenessSheet.getCell("C5").value = "Income Statement";
  completenessSheet.getCell("C10").value = "Revenue";
  completenessSheet.getCell("C20").value = "Beginning Cash Balance";
  completenessSheet.getCell("C30").value = "Balance Sheet";
  completenessSheet.getCell("C35").value = "Cash & Cash Equivalents";
  const completenessRows = __fillModelServiceTestHooks.buildHistoricalSourceLedgerRows(
    company,
    periodEntries,
    completenessSheet,
    [
      {
        row: 10,
        label: "Revenue",
        classification: "direct",
        statement: "income",
        kind: "duration",
        resolver: () => ({ value: null, sources: [] }),
        scale: 1_000_000,
        modelContext: { sectionHeader: "Income Statement" }
      },
      {
        row: 20,
        label: "Beginning Cash Balance",
        classification: "direct",
        statement: "support",
        kind: "instant",
        concepts: ["CashAndCashEquivalentsAtCarryingValue"],
        scale: 1_000_000,
        modelContext: { sectionHeader: "Cash Flow Statement" }
      },
      {
        row: 35,
        label: "Cash & Cash Equivalents",
        classification: "direct",
        statement: "balance",
        kind: "instant",
        resolver: () => ({ value: null, sources: [] }),
        scale: 1_000_000,
        modelContext: { sectionHeader: "Balance Sheet" }
      }
    ],
    [{ period: "1Q23", col: 6 }],
    []
  );
  assert.deepEqual(
    completenessRows.map((row) => row.cell),
    ["F10", "F35"],
    "primary income-statement and strict balance-sheet blanks receive dispositions while a support-schedule blank is excluded"
  );
  assert.ok(completenessRows.every((row) => row.value === null && row.requiredCoreHistoricalInput));
  const silentCoreBlankErrors = await validateHistoricalSourceLedger(completenessRows, periodEntries, company, new Map());
  assert.ok(
    silentCoreBlankErrors.some((error) => /Model!F10 1Q23: required core historical input "Revenue" is blank/i.test(error)),
    "a silent blank on an active primary income-statement input must fail analyst-readiness validation"
  );
  assert.ok(
    silentCoreBlankErrors.some((error) => /Model!F35 1Q23: required core historical input "Cash & Cash Equivalents" is blank/i.test(error)),
    "a silent blank on a strict primary balance-sheet input must fail analyst-readiness validation"
  );

  const irrelevantSupportBlank = ledgerRow({
    modelRow: 20,
    modelRowLabel: "Beginning Cash Balance",
    cell: "F20",
    value: null,
    sourceStatement: "support",
    sourceTableType: "cash_flow_or_support",
    mappingStatus: "stale_or_unsupported",
    requiredCoreHistoricalInput: false
  });
  assert.deepEqual(
    await validateHistoricalSourceLedger([irrelevantSupportBlank], periodEntries, company, new Map()),
    [],
    "an irrelevant blank support-schedule cell does not block workbook completion"
  );

  assert.deepEqual(
    await validateHistoricalSourceLedger([ledgerRow({ value: 0, rawSecValue: "RevenueFromContractWithCustomerExcludingAssessedTax=0mm" })], periodEntries, company, new Map()),
    [],
    "an explicit zero backed by the current SEC filing satisfies the core-input completeness gate"
  );

  const supportedFormulaRows = [
    ledgerRow(),
    ledgerRow({ cell: "F11", modelRow: 11, modelRowLabel: "COGS", value: -40, sourceLineItemLabel: "Cost of revenue", sourceXbrlTag: "CostOfRevenue", rawSecValue: "CostOfRevenue=-40mm" }),
    ledgerRow({
      cell: "F12",
      modelRow: 12,
      modelRowLabel: "Gross Profit",
      value: 60,
      accessionNumber: "",
      accessionRaw: "",
      accessionNormalized: "",
      sourceLineItemLabel: "",
      sourceXbrlTag: "",
      mappingStatus: "formula_preserved",
      workbookFormula: "F10+F11",
      workbookFormulaPrecedents: "Model!F10; Model!F11",
      rawSecValue: "",
      sourceProvenance: [],
      classificationReason: "Template formula over SEC-backed precedents."
    })
  ];
  assert.deepEqual(
    await validateHistoricalSourceLedger(supportedFormulaRows, periodEntries, company, new Map()),
    [],
    "an internal formula is supported only when every precedent closes to the current-period SEC ledger"
  );

  const sourceTiedFormula = ledgerRow({
    value: 100,
    mappingStatus: "formula_preserved",
    workbookFormula: "'Segment Analysis'!F7",
    workbookFormulaPrecedents: "Segment Analysis!F7",
    sourceMappingType: "direct",
    sourceSignConvention: "copied",
    classificationReason: "Strict formula evaluation tied directly to current-period SEC revenue."
  });
  assert.deepEqual(
    await validateHistoricalSourceLedger([sourceTiedFormula], periodEntries, company, new Map()),
    [],
    "a required formula with its own same-period SEC amount tie does not need every cross-sheet helper precedent duplicated in this ledger"
  );
  const formulaDependingOnSourceTiedFormula = ledgerRow({
    cell: "F15",
    modelRow: 15,
    modelRowLabel: "Revenue-derived subtotal",
    value: 100,
    accessionNumber: "",
    accessionRaw: "",
    accessionNormalized: "",
    mappingStatus: "formula_preserved",
    workbookFormula: "F10",
    workbookFormulaPrecedents: "Model!F10",
    rawSecValue: "",
    sourceProvenance: [],
    sourceLineItemLabel: "",
    sourceXbrlTag: ""
  });
  assert.deepEqual(
    await validateHistoricalSourceLedger(
      [sourceTiedFormula, formulaDependingOnSourceTiedFormula],
      periodEntries,
      company,
      new Map()
    ),
    [],
    "a downstream formula must honor a precedent formula's independently validated same-period SEC amount support"
  );

  const formulaWithConceptOnlyMetadata = [
    ledgerRow(),
    ledgerRow({
      cell: "F13",
      modelRow: 13,
      modelRowLabel: "Income Tax Expense",
      value: 21,
      mappingStatus: "formula_preserved",
      workbookFormula: "F10*21%",
      workbookFormulaPrecedents: "Model!F10",
      rawSecValue: "IncomeTaxExpenseBenefit",
      sourceProvenance: [],
      sourceXbrlTag: "IncomeTaxExpenseBenefit",
      classificationReason: "Template formula over an SEC-backed precedent."
    })
  ];
  const conceptOnlyFormulaErrors = await validateHistoricalSourceLedger(
    formulaWithConceptOnlyMetadata,
    periodEntries,
    company,
    new Map()
  );
  assert.ok(
    conceptOnlyFormulaErrors.some((error) => /unsupported numeric literal\(s\) 21%/i.test(error)),
    "a concept name without structured amount provenance cannot certify a hardcoded historical tax-rate formula"
  );

  const additiveLiteralErrors = await validateHistoricalSourceLedger(
    [
      ledgerRow(),
      ledgerRow({
        cell: "F14",
        modelRow: 14,
        modelRowLabel: "Unsupported Historical Adjustment",
        value: 1099,
        accessionNumber: "",
        accessionRaw: "",
        accessionNormalized: "",
        mappingStatus: "formula_preserved",
        workbookFormula: "F10+999",
        workbookFormulaPrecedents: "Model!F10",
        rawSecValue: "",
        sourceProvenance: [],
        sourceLineItemLabel: "",
        sourceXbrlTag: ""
      })
    ],
    periodEntries,
    company,
    new Map()
  );
  assert.ok(
    additiveLiteralErrors.some((error) => /unsupported numeric literal\(s\) 999/i.test(error)),
    "an audited precedent cannot legitimize an additive hardcoded amount hidden inside a historical formula"
  );

  const oversizedRangeWorkbook = new ExcelJS.Workbook();
  const oversizedRangeSheet = oversizedRangeWorkbook.addWorksheet("Model");
  oversizedRangeSheet.getCell("C20").value = "Range-backed Historical Amount";
  oversizedRangeSheet.getCell("F20").value = { formula: "F30+SUM(A1:A501)", result: 1122 };
  oversizedRangeSheet.getCell("F30").value = 123;
  oversizedRangeSheet.getCell("A1").value = 999;
  const oversizedRangeRows = __fillModelServiceTestHooks.buildHistoricalSourceLedgerRows(
    company,
    periodEntries,
    oversizedRangeSheet,
    [{ row: 20, label: "Range-backed Historical Amount", classification: "formula", statement: "support", kind: "duration" }],
    [{ period: "1Q23", col: 6 }],
    [
      auditRow({
        sheetName: "Model",
        cell: "F30",
        modelRowLabel: "Audited Support",
        period: "1Q23",
        valueWritten: 123,
        mappingType: "direct",
        conceptsUsed: "OtherCurrentAssets=123mm",
        sourceStatement: "support",
        accession: "0000000000-23-000001",
        sourceProvenance: [provenance({ concept: "OtherCurrentAssets", label: "Other current assets", value: 123 })]
      })
    ]
  );
  assert.match(
    oversizedRangeRows.find((row) => row.cell === "F20").workbookFormulaUnsupportedReferences,
    /A1:A501 expands to 501 cells/i,
    "an oversized formula range must be recorded as unsupported instead of silently omitted"
  );
  const oversizedRangeErrors = await validateHistoricalSourceLedger(
    oversizedRangeRows,
    periodEntries,
    company,
    new Map()
  );
  assert.ok(
    oversizedRangeErrors.some((error) => /reference syntax that cannot be fully dependency-audited.*A1:A501/i.test(error)),
    "a formula whose range exceeds the dependency-expansion cap must fail closed"
  );

  const nonCoreHelperFormula = ledgerRow({
    modelRow: 24,
    modelRowLabel: "2023",
    cell: "F24",
    value: "2023",
    mappingStatus: "formula_preserved",
    requiredCoreHistoricalInput: false,
    workbookFormula: "G24",
    workbookFormulaPrecedents: "Model!G24",
    rawSecValue: "",
    sourceLineItemLabel: "",
    sourceXbrlTag: ""
  });
  assert.deepEqual(
    await validateHistoricalSourceLedger([nonCoreHelperFormula], periodEntries, company, new Map()),
    [],
    "calendar/header helper formulas outside financial historical values do not require SEC provenance"
  );

  const presentationWorkbook = new ExcelJS.Workbook();
  const presentationSheet = presentationWorkbook.addWorksheet("Model");
  presentationSheet.getCell("C25").value = "Income Statement";
  presentationSheet.getCell("F25").value = { formula: "F24", result: 2023 };
  presentationSheet.getCell("F24").value = 2023;
  presentationSheet.getCell("C26").value = "Days In Period";
  presentationSheet.getCell("F26").value = 90;
  const presentationLedgerRows = __fillModelServiceTestHooks.buildHistoricalSourceLedgerRows(
    company,
    periodEntries,
    presentationSheet,
    [
      { row: 25, label: "Income Statement", classification: "formula", statement: "income", kind: "duration" },
      { row: 26, label: "Days In Period", classification: "unused", statement: "support", kind: "duration" }
    ],
    [{ period: "1Q23", col: 6 }],
    []
  );
  assert.deepEqual(
    presentationLedgerRows,
    [],
    "statement headers and SEC-derived calendar metadata must stay outside the financial Source Ledger dependency audit"
  );

  presentationSheet.getCell("C27").value = "Working Capital Schedule";
  presentationSheet.getCell("F27").value = { formula: "F25", result: 2023 };
  assert.deepEqual(
    __fillModelServiceTestHooks.buildHistoricalSourceLedgerRows(
      company,
      periodEntries,
      presentationSheet,
      [{ row: 27, label: "Working Capital Schedule", classification: "formula", statement: "support", kind: "duration" }],
      [{ period: "1Q23", col: 6 }],
      []
    ),
    [],
    "schedule, driver, analysis, and assumption headings are presentation metadata rather than financial amounts"
  );

  const staleCachedFormulaWorkbook = new ExcelJS.Workbook();
  const staleCachedFormulaSheet = staleCachedFormulaWorkbook.addWorksheet("Model");
  staleCachedFormulaSheet.getCell("C20").value = "Optional Model Output";
  staleCachedFormulaSheet.getCell("F20").value = { formula: "F30", result: 999 };
  staleCachedFormulaSheet.getCell("F30").value = null;
  const staleCachedFormulaRows = __fillModelServiceTestHooks.buildHistoricalSourceLedgerRows(
    company,
    periodEntries,
    staleCachedFormulaSheet,
    [{ row: 20, label: "Optional Model Output", classification: "formula", statement: "support", kind: "duration" }],
    [{ period: "1Q23", col: 6 }],
    []
  );
  assert.equal(
    staleCachedFormulaRows.find((row) => row.cell === "F20").value,
    null,
    "a cached value from the uploaded template must not be treated as the current formula result when its precedent is blank"
  );
  assert.deepEqual(
    await validateHistoricalSourceLedger(staleCachedFormulaRows, periodEntries, company, new Map()),
    [],
    "a non-core optional formula with no strict current value stays outside financial provenance instead of surfacing stale cached data"
  );

  presentationSheet.getCell("C30").value = "Revenue";
  presentationSheet.getCell("F30").value = 100;
  presentationSheet.getCell("C36").value = "Calendar-adjusted support formula";
  presentationSheet.getCell("F36").value = { formula: "F30+F26+F25", result: 2193 };
  const calendarHelperDependencyRows = __fillModelServiceTestHooks.buildHistoricalSourceLedgerRows(
    company,
    periodEntries,
    presentationSheet,
    [{ row: 36, label: "Calendar-adjusted support formula", classification: "formula", statement: "support", kind: "duration" }],
    [{ period: "1Q23", col: 6 }],
    [
      auditRow({
        sheetName: "Model",
        cell: "F30",
        modelRowLabel: "Revenue",
        period: "1Q23",
        valueWritten: 100,
        conceptsUsed: "RevenueFromContractWithCustomerExcludingAssessedTax=100mm",
        sourceStatement: "income",
        accession: "0000000000-23-000001",
        sourceProvenance: [provenance({ value: 100 })]
      })
    ]
  );
  assert.equal(
    calendarHelperDependencyRows.find((row) => row.cell === "F36").workbookFormulaPrecedents,
    "Model!F30",
    "statement-header and Days-in-Period cells may drive calendar logic but must not be misclassified as financial amount precedents"
  );
  assert.deepEqual(
    calendarHelperDependencyRows.map((row) => row.cell).sort(),
    ["F30", "F36"],
    "the dependency closure must retain the audited financial precedent while excluding presentation metadata"
  );

  const textControlWorkbook = new ExcelJS.Workbook();
  const textControlSheet = textControlWorkbook.addWorksheet("Model");
  textControlSheet.getCell("C34").value = "Depreciation & Amortization";
  textControlSheet.getCell("F30").value = 100;
  textControlSheet.getCell("C36").value = "EBIT";
  textControlSheet.getCell("F36").value = {
    formula: 'IF($C$34="Depreciation & Amortization (incl. in SG&A)",F30,F30)',
    result: 100
  };
  const textControlLedgerRow = __fillModelServiceTestHooks.buildHistoricalSourceLedgerRows(
    company,
    periodEntries,
    textControlSheet,
    [{ row: 36, label: "EBIT", classification: "formula", statement: "income", kind: "duration" }],
    [{ period: "1Q23", col: 6 }],
    []
  )[0];
  assert.equal(
    textControlLedgerRow.workbookFormulaPrecedents,
    "Model!F30",
    "constant text controls may select a formula branch but are not financial amount precedents requiring SEC ledger rows"
  );

  const optionalAdjustmentWorkbook = new ExcelJS.Workbook();
  const optionalAdjustmentSheet = optionalAdjustmentWorkbook.addWorksheet("Model");
  optionalAdjustmentSheet.getCell("C48").value = "Post-Tax Adjustments";
  optionalAdjustmentSheet.getCell("H48").value = { formula: "0.26*H284", result: 209.82 };
  optionalAdjustmentSheet.getCell("J48").value = { formula: "SUM(F48:I48)", result: 209.82 };
  const optionalAdjustmentFillRow = __fillModelServiceTestHooks.fillRowForContext({
    sheetName: "Model",
    row: 48,
    label: "Post-Tax Adjustments",
    indentation: 0,
    hasHistoricalFormula: true,
    hasHardcodedInput: false,
    hasNetRevenueInterestExpenseAbove: false,
    projectedColumns: 0,
    signConvention: -1
  });
  const optionalAdjustmentAudit = [];
  const optionalAdjustmentResult = __fillModelServiceTestHooks.finalizeUnsupportedOptionalIncomeAdjustments(
    optionalAdjustmentSheet,
    [optionalAdjustmentFillRow],
    [
      { period: "3Q23", col: 8 },
      { period: "FY23", col: 10 }
    ],
    { duration: new Map(), instant: new Map() },
    optionalAdjustmentAudit
  );
  assert.equal(optionalAdjustmentResult.clearedCells, 2);
  assert.equal(optionalAdjustmentSheet.getCell("H48").value, null);
  assert.equal(optionalAdjustmentSheet.getCell("J48").value, null);
  assert.equal(optionalAdjustmentAudit.length, 2, "stale optional adjustment formulas must be explicitly audited as cleared");

  const namedRangeFormulaLedgerRow = (range) => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Model");
    sheet.getCell("F24").value = { formula: "LegacyNamedRange", result: 42 };
    sheet.getCell("G24").value = 42;
    workbook.definedNames.add(range, "LegacyNamedRange");
    return __fillModelServiceTestHooks.buildHistoricalSourceLedgerRows(
      company,
      periodEntries,
      sheet,
      [
        {
          row: 24,
          label: "Historical Support Amount",
          classification: "formula",
          statement: "support",
          kind: "duration",
          scale: 1_000_000,
          modelContext: { sectionHeader: "Historical Support Schedule", hasHistoricalFormula: true }
        }
      ],
      [{ period: "1Q23", col: 6 }],
      []
    )[0];
  };

  const externalNamedRangeFormula = namedRangeFormulaLedgerRow("'[Legacy.xlsx]Model'!$G$24");
  assert.match(
    externalNamedRangeFormula.workbookFormulaExternalReferences,
    /LegacyNamedRange.*Legacy\.xlsx/i,
    "a formula's external-workbook dependency is retained even when hidden behind a defined name"
  );
  const externalNamedRangeErrors = await validateHistoricalSourceLedger(
    [externalNamedRangeFormula],
    periodEntries,
    company,
    new Map()
  );
  assert.ok(
    externalNamedRangeErrors.some((error) => /external workbook.*defined name/i.test(error)),
    "=LegacyNamedRange must fail when the name resolves to a prior external workbook"
  );

  const internalNamedRangeFormula = namedRangeFormulaLedgerRow("Model!$G$24");
  assert.equal(
    internalNamedRangeFormula.workbookFormulaPrecedents,
    "Model!G24",
    "an internal defined name resolves to its concrete cell precedent"
  );
  const unledgeredNamedRangeErrors = await validateHistoricalSourceLedger(
    [internalNamedRangeFormula],
    periodEntries,
    company,
    new Map()
  );
  assert.ok(
    unledgeredNamedRangeErrors.some((error) => /formula precedent Model!G24 is outside the audited historical source ledger/i.test(error)),
    "an internally resolved name cannot certify an unledgered hardcoded historical support value"
  );

  const safeNamedRangeSource = ledgerRow({
    modelRow: 24,
    modelRowLabel: "Historical Support Source",
    cell: "G24",
    modelColumn: "G",
    value: 42,
    sourceStatement: "support",
    sourceTableType: "cash_flow_or_support",
    sourceLineItemLabel: "Other current assets",
    sourceXbrlTag: "OtherCurrentAssets",
    rawSecValue: "OtherCurrentAssets=42mm",
    requiredCoreHistoricalInput: false
  });
  assert.deepEqual(
    await validateHistoricalSourceLedger(
      [safeNamedRangeSource, internalNamedRangeFormula],
      periodEntries,
      company,
      new Map()
    ),
    [],
    "a defined name over a ledger-backed same-period internal dependency remains valid"
  );

  const poisonedValueErrors = await validateHistoricalSourceLedger([ledgerRow({ value: 999 })], periodEntries, company, new Map());
  assert.ok(
    poisonedValueErrors.some((error) => /workbook value 999 does not reconcile to structured SEC amount support/i.test(error)),
    "an accession-backed audit row cannot certify a workbook value that differs from its recorded SEC amount"
  );

  const signNormalizedRows = [
    ledgerRow({
      value: -100,
      sourceMappingType: "direct",
      sourceSignConvention: "inverted to match model sign convention"
    })
  ];
  assert.deepEqual(
    await validateHistoricalSourceLedger(signNormalizedRows, periodEntries, company, new Map()),
    [],
    "direct SEC amount integrity honors the mapping's recorded sign convention"
  );

  const groupedRows = [
    ledgerRow({
      value: 100,
      sourceMappingType: "grouped",
      rawSecValue: "ProductRevenue=60mm; ServiceRevenue=40mm",
      sourceXbrlTag: "ProductRevenue; ServiceRevenue",
      sourceProvenance: [
        provenance({ concept: "ProductRevenue", label: "Product revenue", value: 60 }),
        provenance({ concept: "ServiceRevenue", label: "Service revenue", value: 40 })
      ]
    })
  ];
  assert.deepEqual(
    await validateHistoricalSourceLedger(groupedRows, periodEntries, company, new Map()),
    [],
    "grouped direct mappings reconcile the workbook value to the sum of SEC-backed components"
  );

  const derivedRows = [
    ledgerRow({
      value: 60,
      mappingStatus: "validated_current_company_derived_value",
      sourceMappingType: "derived",
      rawSecValue: "DerivedOutput:GrossProfitDerivedFromRevenueAndCostOfRevenue=60mm; Revenue=100mm; CostOfRevenue=-40mm",
      sourceProvenance: [
        provenance({
          role: "derived_output",
          concept: "GrossProfitDerivedFromRevenueAndCostOfRevenue",
          value: 60,
          sourceLayer: "derived",
          accession: "",
          derivationCalculation: linearDerivationCalculation([
            { concept: "Revenue", value: 100, coefficient: 1 },
            { concept: "CostOfRevenue", value: -40, coefficient: 1 }
          ])
        }),
        provenance({ concept: "Revenue", value: 100 }),
        provenance({ concept: "CostOfRevenue", value: -40 })
      ],
      classificationReason: "Derived as SEC revenue less SEC cost of revenue."
    })
  ];
  assert.deepEqual(
    await validateHistoricalSourceLedger(derivedRows, periodEntries, company, new Map()),
    [],
    "derived arithmetic is validated through its provenance metadata without being mistaken for a direct source sum"
  );

  const nestedSignNormalizedExpenseRows = [
    ledgerRow({
      value: 60,
      mappingStatus: "validated_current_company_derived_value",
      sourceMappingType: "derived",
      rawSecValue: "DerivedOutput:OperatingIncomeBridge=60mm; DerivedInput:OperatingIncomeFromComponents=60mm; Revenue=100mm; CostOfRevenue=40mm",
      sourceProvenance: [
        provenance({
          role: "derived_output",
          concept: "OperatingIncomeBridge",
          value: 60,
          sourceLayer: "derived",
          accession: "",
          derivationCalculation: linearDerivationCalculation([
            { concept: "OperatingIncomeFromComponents", value: 60, coefficient: 1 }
          ])
        }),
        provenance({
          role: "derived_input",
          concept: "OperatingIncomeFromComponents",
          value: 60,
          sourceLayer: "derived",
          accession: "",
          derivationCalculation: linearDerivationCalculation([
            { concept: "Revenue", value: 100, coefficient: 1 },
            { concept: "CostOfRevenue", value: -40, coefficient: 1 }
          ])
        }),
        provenance({ concept: "Revenue", value: 100 }),
        provenance({ concept: "CostOfRevenue", value: 40 }),
        ...Array.from({ length: 15 }, (_unused, index) =>
          provenance({ concept: `UnrelatedOperatingFact${index + 1}`, value: index + 1 })
        )
      ],
      classificationReason: "Nested operating-income derivation normalizes positive SEC expense magnitudes to negative calculation terms."
    })
  ];
  assert.deepEqual(
    await validateHistoricalSourceLedger(nestedSignNormalizedExpenseRows, periodEntries, company, new Map()),
    [],
    "a nested derivation accepts same-concept, same-period positive SEC expense magnitude as support for its explicitly negative calculation term"
  );

  const largeGroupedNestedRows = [
    ledgerRow({
      value: 100,
      mappingStatus: "validated_current_company_derived_value",
      sourceMappingType: "derived",
      rawSecValue: "DerivedOutput:OuterGroupedBridge=100mm; DerivedInput:NestedGroupedInput=100mm; ProductRevenue=60mm; ServiceRevenue=40mm",
      sourceProvenance: [
        provenance({
          role: "derived_output",
          concept: "OuterGroupedBridge",
          value: 100,
          sourceLayer: "derived",
          accession: "",
          derivationCalculation: linearDerivationCalculation([
            { concept: "NestedGroupedInput", value: 100, coefficient: 1 }
          ])
        }),
        provenance({
          role: "derived_input",
          concept: "NestedGroupedInput",
          value: 100,
          sourceLayer: "derived",
          accession: "",
          derivationCalculation: linearDerivationCalculation([
            { concept: "ProductRevenue", value: 100, coefficient: 1 }
          ])
        }),
        provenance({ concept: "ProductRevenue", value: 60 }),
        provenance({ concept: "ServiceRevenue", value: 40 }),
        ...Array.from({ length: 17 }, (_unused, index) =>
          provenance({ concept: `UnrelatedGroupedFact${index + 1}`, value: index + 1 })
        )
      ],
      classificationReason: "A grouped nested derivation uses a representative concept for two same-period SEC facts."
    })
  ];
  assert.deepEqual(
    await validateHistoricalSourceLedger(largeGroupedNestedRows, periodEntries, company, new Map()),
    [],
    "a same-period two-source grouped derivation remains valid when the ledger row contains more than 16 provenance candidates"
  );

  const signNormalizedDerivedRows = [
    ledgerRow({
      value: -100,
      mappingStatus: "validated_current_company_derived_value",
      sourceMappingType: "derived",
      sourceSignConvention: "inverted to match model sign convention",
      rawSecValue: "DerivedOutput:CashOutflowBridge=-100mm; PaymentsToAcquireProductiveAssets=100mm",
      sourceProvenance: [
        provenance({
          role: "derived_output",
          concept: "CashOutflowBridge",
          value: -100,
          sourceLayer: "derived",
          accession: "",
          derivationCalculation: linearDerivationCalculation([
            { concept: "PaymentsToAcquireProductiveAssets", value: 100, coefficient: -1 }
          ])
        }),
        provenance({ concept: "PaymentsToAcquireProductiveAssets", value: 100 })
      ],
      classificationReason: "Derived cash outflow with the final model sign already normalized."
    })
  ];
  assert.deepEqual(
    await validateHistoricalSourceLedger(signNormalizedDerivedRows, periodEntries, company, new Map()),
    [],
    "a final derived output must not be inverted a second time after its replayable calculation already normalized the model sign"
  );

  const rawPositiveDerived = {
    value: 100_000_000,
    classification: "grouped",
    sources: [
      {
        concept: "RawPositiveBridge",
        label: "Raw positive derived amount",
        value: 100_000_000,
        sourceLayer: "derived",
        periodKey: "1Q23",
        periodType: "quarterly",
        derivationCalculation: linearDerivationCalculation([
          { concept: "PaymentsToAcquireProductiveAssets", value: 100_000_000, coefficient: 1 }
        ])
      },
      {
        ...provenance({ concept: "PaymentsToAcquireProductiveAssets", value: 100_000_000 }),
        sourceLayer: "sec_filing_package",
        accn: "0000000000-23-000001",
        periodKey: "1Q23",
        periodType: "quarterly"
      }
    ]
  };
  const normalizedDerived = __fillModelServiceTestHooks.resolvedWithFinalAuditDerivation(
    {
      row: 82,
      label: "Capital Expenditures",
      classification: "grouped",
      statement: "support",
      kind: "duration",
      scale: 1_000_000,
      sign: -1
    },
    "1Q23",
    rawPositiveDerived,
    -100
  );
  assert.equal(normalizedDerived.value, -100_000_000);
  assert.equal(normalizedDerived.sources[0].value, -100_000_000);
  assert.equal(
    normalizedDerived.sources[0].derivationCalculation.terms[0].coefficient,
    -1,
    "audit creation must wrap a pre-sign derived amount in one final model-normalized derivation"
  );

  const derivedCalculationRows = [
    ledgerRow({
      value: 60,
      mappingStatus: "validated_current_company_derived_value",
      sourceMappingType: "derived",
      rawSecValue: "DerivedOutput:GrossProfitDerivedFromRevenueAndCostOfRevenue=60mm; Revenue=100mm; CostOfRevenue=-40mm",
      sourceProvenance: [
        provenance({
          role: "derived_output",
          concept: "GrossProfitDerivedFromRevenueAndCostOfRevenue",
          value: 60,
          sourceLayer: "derived",
          accession: "",
          derivationCalculation: linearDerivationCalculation([
            { concept: "Revenue", value: 100, coefficient: 1 },
            { concept: "CostOfRevenue", value: -40, coefficient: 1 }
          ])
        }),
        provenance({ concept: "Revenue", value: 100 }),
        provenance({ concept: "CostOfRevenue", value: -40 })
      ],
      classificationReason: "Derived as SEC revenue less SEC cost of revenue."
    })
  ];
  assert.deepEqual(
    await validateHistoricalSourceLedger(derivedCalculationRows, periodEntries, company, new Map()),
    [],
    "a recorded synthetic derivation amount certifies the corresponding workbook value"
  );
  const poisonedDerivedErrors = await validateHistoricalSourceLedger(
    [{ ...derivedCalculationRows[0], value: 999 }],
    periodEntries,
    company,
    new Map()
  );
  assert.ok(
    poisonedDerivedErrors.some((error) => /workbook value 999 does not reconcile to final derived output/i.test(error)),
    "a synthetic derivation cannot certify a poisoned workbook amount"
  );
  const poisonedCalculationRow = ledgerRow({
    ...derivedCalculationRows[0],
    sourceProvenance: derivedCalculationRows[0].sourceProvenance.map((source) =>
      source.role === "derived_output"
        ? {
            ...source,
            derivationCalculation: linearDerivationCalculation([
              { concept: "Revenue", value: 1, coefficient: 1 },
              { concept: "CostOfRevenue", value: 2, coefficient: 1 }
            ])
          }
        : source
    )
  });
  const poisonedCalculationErrors = await validateHistoricalSourceLedger(
    [poisonedCalculationRow],
    periodEntries,
    company,
    new Map()
  );
  assert.ok(
    poisonedCalculationErrors.some((error) => /does not replay from its calculation terms/i.test(error)),
    "a synthetic output that agrees with the workbook cannot certify unrelated arithmetic inputs"
  );

  const modelSupportedCalculationRow = ledgerRow({
    ...derivedCalculationRows[0],
    sourceProvenance: [
      provenance({
        role: "derived_output",
        concept: "ModelSupportedDerivedOutput",
        value: 60,
        sourceLayer: "derived",
        accession: "",
        derivationCalculation: linearDerivationCalculation([{ concept: "ModelOnlyInput", value: 60, coefficient: 1 }])
      }),
      provenance({
        role: "model_source",
        concept: "ModelOnlyInput",
        label: "Existing model value",
        value: 60,
        sourceLayer: "model",
        accession: ""
      }),
      provenance({ concept: "Revenue", value: 100 })
    ]
  });
  const modelSupportedCalculationErrors = await validateHistoricalSourceLedger(
    [modelSupportedCalculationRow],
    periodEntries,
    company,
    new Map()
  );
  assert.ok(
    modelSupportedCalculationErrors.some((error) => /ModelOnlyInput=60mm.*not supported/i.test(error)),
    "a model-layer value cannot certify a derived SEC calculation term"
  );

  const unrelatedConceptCalculationRow = ledgerRow({
    ...derivedCalculationRows[0],
    sourceProvenance: [
      provenance({
        role: "derived_output",
        concept: "UnrelatedConceptDerivedOutput",
        value: 60,
        sourceLayer: "derived",
        accession: "",
        derivationCalculation: linearDerivationCalculation([{ concept: "UnrelatedConcept", value: 60, coefficient: 1 }])
      }),
      provenance({ concept: "Revenue", value: 100 }),
      provenance({ concept: "CostOfRevenue", value: -40 })
    ]
  });
  const unrelatedConceptCalculationErrors = await validateHistoricalSourceLedger(
    [unrelatedConceptCalculationRow],
    periodEntries,
    company,
    new Map()
  );
  assert.ok(
    unrelatedConceptCalculationErrors.some((error) => /UnrelatedConcept=60mm.*not supported/i.test(error)),
    "an arithmetic combination of unrelated SEC concepts cannot certify a named derivation term"
  );

  const unsupportedNestedCalculationRow = ledgerRow({
    ...derivedCalculationRows[0],
    sourceProvenance: [
      provenance({
        role: "derived_output",
        concept: "OuterDerivedOutput",
        value: 60,
        sourceLayer: "derived",
        accession: "",
        derivationCalculation: linearDerivationCalculation([{ concept: "NestedDerivedInput", value: 60, coefficient: 1 }])
      }),
      provenance({
        role: "derived_input",
        concept: "NestedDerivedInput",
        value: 60,
        sourceLayer: "derived",
        accession: "",
        derivationCalculation: undefined
      }),
      provenance({ concept: "Revenue", value: 100 }),
      provenance({ concept: "CostOfRevenue", value: -40 })
    ]
  });
  const unsupportedNestedCalculationErrors = await validateHistoricalSourceLedger(
    [unsupportedNestedCalculationRow],
    periodEntries,
    company,
    new Map()
  );
  assert.ok(
    unsupportedNestedCalculationErrors.some((error) => /NestedDerivedInput=60mm.*not supported/i.test(error)),
    "an uncalculated nested synthetic value cannot self-certify a final derived output"
  );

  const supportedNestedCalculationRow = ledgerRow({
    ...derivedCalculationRows[0],
    sourceProvenance: [
      provenance({
        role: "derived_output",
        concept: "OuterDerivedOutput",
        value: 60,
        sourceLayer: "derived",
        accession: "",
        derivationCalculation: linearDerivationCalculation([{ concept: "NestedDerivedInput", value: 60, coefficient: 1 }])
      }),
      provenance({
        role: "derived_input",
        concept: "NestedDerivedInput",
        value: 60,
        sourceLayer: "derived",
        accession: "",
        derivationCalculation: linearDerivationCalculation([
          { concept: "Revenue", value: 100, coefficient: 1 },
          { concept: "CostOfRevenue", value: -40, coefficient: 1 }
        ])
      }),
      provenance({ concept: "Revenue", value: 100 }),
      provenance({ concept: "CostOfRevenue", value: -40 })
    ]
  });
  assert.deepEqual(
    await validateHistoricalSourceLedger([supportedNestedCalculationRow], periodEntries, company, new Map()),
    [],
    "a nested synthetic input remains valid only when its own calculation replays from named SEC sources"
  );

  const missingCalculationRow = ledgerRow({
    ...derivedCalculationRows[0],
    sourceProvenance: derivedCalculationRows[0].sourceProvenance.map((source) =>
      source.role === "derived_output" ? { ...source, derivationCalculation: undefined } : source
    )
  });
  const missingCalculationErrors = await validateHistoricalSourceLedger(
    [missingCalculationRow],
    periodEntries,
    company,
    new Map()
  );
  assert.ok(
    missingCalculationErrors.some((error) => /missing replayable calculation terms/i.test(error)),
    "a derived output cannot self-certify without replayable arithmetic metadata"
  );

  const wrongPeriodDerivedRow = ledgerRow({
    ...derivedCalculationRows[0],
    sourceProvenance: derivedCalculationRows[0].sourceProvenance.map((source) =>
      source.role === "sec_source" && source.concept === "CostOfRevenue"
        ? { ...source, periodKey: "FY22", periodType: "annual", startDate: "2022-01-01", endDate: "2022-12-31" }
        : source
    )
  });
  const wrongPeriodDerivedErrors = await validateHistoricalSourceLedger(
    [wrongPeriodDerivedRow],
    periodEntries,
    company,
    new Map()
  );
  assert.ok(
    wrongPeriodDerivedErrors.some((error) => /derived input CostOfRevenue uses period FY22.*not allowed for 1Q23/i.test(error)),
    "a correct-looking derived amount cannot cite SEC inputs from an unrelated fiscal period"
  );

  const staleFormulaRows = supportedFormulaRows.map((row) => ({ ...row }));
  staleFormulaRows[1].mappingStatus = "stale_or_unsupported";
  const staleErrors = await validateHistoricalSourceLedger(staleFormulaRows, periodEntries, company, new Map());
  assert.ok(staleErrors.some((error) => /formula precedent Model!F11 lacks current-company source support/i.test(error)));

  const opaqueFormulaRows = supportedFormulaRows.map((row) => ({ ...row }));
  opaqueFormulaRows[2].workbookFormula = "NamedRevenue-NamedCosts";
  opaqueFormulaRows[2].workbookFormulaPrecedents = "";
  const opaqueErrors = await validateHistoricalSourceLedger(opaqueFormulaRows, periodEntries, company, new Map());
  assert.ok(opaqueErrors.some((error) => /no traceable audited cell precedents/i.test(error)));

  const wrongPeriodAccession = "0000000000-23-000001";
  const wrongPeriodSource = [
    ledgerRow({
      accessionNumber: wrongPeriodAccession,
      accessionRaw: wrongPeriodAccession,
      accessionNormalized: "000000000023000001",
      reportingPeriodEndDate: "2022-12-31",
      filingPeriod: "2022-12-31"
    })
  ];
  const wrongPeriodErrors = await validateHistoricalSourceLedger(
    wrongPeriodSource,
    periodEntries,
    company,
    new Map()
  );
  assert.ok(
    wrongPeriodErrors.some((error) => /source end date 2022-12-31 that does not match 1Q23/i.test(error)),
    "even an accession mapped to the row cannot excuse a wrong source reporting date"
  );

  const comparativeAccession = "0000000000-24-000099";
  const comparativeSource = [
    ledgerRow({
      modelRowLabel: "Cash & Cash Equivalents",
      sourceStatement: "balance",
      sourceLineItemLabel: "Cash and cash equivalents",
      sourceXbrlTag: "CashAndCashEquivalentsAtCarryingValue",
      accessionNumber: comparativeAccession,
      accessionRaw: comparativeAccession,
      accessionNormalized: "000000000024000099",
      reportingPeriodEndDate: "2023-03-31",
      filingPeriod: "2023-03-31"
    })
  ];
  assert.deepEqual(
    await validateHistoricalSourceLedger(comparativeSource, periodEntries, company, new Map([[comparativeAccession, {}]])),
    [],
    "a later current-company SEC filing may support the same comparative instant when its period end and period type exactly match the model period"
  );

  const annualBalancePeriodEntries = [
    {
      period: "FY23",
      column: 10,
      modelColumn: "J",
      accessionKey: "000000000023000010",
      accessionNumber: "0000000000-23-000010",
      form: "10-K",
      filingDate: "2024-02-01",
      periodEndDate: "2023-12-31",
      fiscalYearLabel: "FY23",
      fiscalQuarterLabel: "FY"
    }
  ];
  const annualBalanceInstant = ledgerRow({
    modelRowLabel: "Cash & Cash Equivalents",
    modelColumn: "J",
    cell: "J120",
    fiscalPeriod: "FY23",
    sourceStatement: "balance",
    sourceLineItemLabel: "Cash and cash equivalents",
    sourceXbrlTag: "CashAndCashEquivalentsAtCarryingValue",
    accessionNumber: "0000000000-23-000010",
    accessionRaw: "0000000000-23-000010",
    accessionNormalized: "000000000023000010",
    filingFormType: "10-K",
    reportingPeriodEndDate: "2023-12-31",
    filingPeriod: "2023-12-31",
    sourceProvenance: [
      provenance({
        concept: "CashAndCashEquivalentsAtCarryingValue",
        label: "Cash and cash equivalents",
        accession: "0000000000-23-000010",
        form: "10-K",
        startDate: "",
        endDate: "2023-12-31",
        periodKey: "4Q23",
        periodType: "instant"
      })
    ]
  });
  assert.deepEqual(
    await validateHistoricalSourceLedger([annualBalanceInstant], annualBalancePeriodEntries, company, new Map()),
    [],
    "an FY balance-sheet column may cite the equivalent 4Q instant key when the exact SEC period end and instant type match"
  );

  const submittingAgentAccession = "0001628280-23-000001";
  const submittingAgentPeriodEntries = [
    {
      ...periodEntries[0],
      accessionKey: "000162828023000001",
      accessionNumber: submittingAgentAccession
    }
  ];
  const submittingAgentSource = ledgerRow({
    accessionNumber: submittingAgentAccession,
    accessionRaw: submittingAgentAccession,
    accessionNormalized: "000162828023000001",
    sourceProvenance: [provenance({ accession: submittingAgentAccession })]
  });
  assert.deepEqual(
    await validateHistoricalSourceLedger([submittingAgentSource], submittingAgentPeriodEntries, company, new Map()),
    [],
    "an accession in the selected issuer's EDGAR filing metadata remains valid when its prefix identifies a submitting agent rather than the issuer CIK"
  );
  const mismatchedLedgerIdentityErrors = await validateHistoricalSourceLedger(
    [{ ...submittingAgentSource, cik: "0000000001", ticker: "WRONG" }],
    submittingAgentPeriodEntries,
    company,
    new Map()
  );
  assert.ok(
    mismatchedLedgerIdentityErrors.some((error) => /Source Ledger CIK 0000000001 does not match selected company CIK 0000000000/i.test(error)),
    "the ledger's explicit issuer CIK must match the selected company independently of accession-prefix semantics"
  );
  assert.ok(
    mismatchedLedgerIdentityErrors.some((error) => /Source Ledger ticker WRONG does not match selected company ticker EXM/i.test(error)),
    "the ledger's explicit issuer ticker must match the selected company"
  );

  const fabricatedAccession = "0000000000-24-999999";
  const fabricatedErrors = await validateHistoricalSourceLedger(
    [
      ledgerRow({
        accessionNumber: fabricatedAccession,
        accessionRaw: fabricatedAccession,
        accessionNormalized: "000000000024999999"
      })
    ],
    periodEntries,
    company,
    new Map()
  );
  assert.ok(
    fabricatedErrors.some((error) => /not present in the selected company's SEC filing metadata/i.test(error)),
    "a matching CIK prefix and reporting date cannot make a fabricated accession current"
  );

  const mixedAccessionRaw = `0000000000-23-000001; ${fabricatedAccession}`;
  const mixedAccessionErrors = await validateHistoricalSourceLedger(
    [
      ledgerRow({
        accessionNumber: mixedAccessionRaw,
        accessionRaw: mixedAccessionRaw,
        accessionNormalized: "000000000023000001; 000000000024999999",
        sourceProvenance: [
          provenance(),
          provenance({ concept: "ServiceRevenue", label: "Service revenue", value: 0, accession: fabricatedAccession })
        ]
      })
    ],
    periodEntries,
    company,
    new Map()
  );
  assert.ok(
    mixedAccessionErrors.some((error) => /000000000024999999.*not present in the selected company's SEC filing metadata/i.test(error)),
    "one mapped accession cannot hide an additional fabricated same-CIK accession"
  );

  const malformedDateErrors = await validateHistoricalSourceLedger(
    [
      ledgerRow({
        reportingPeriodEndDate: "2023-03-31-invalid",
        filingPeriod: "2023-03-31-invalid"
      })
    ],
    periodEntries,
    company,
    new Map()
  );
  assert.ok(
    malformedDateErrors.some((error) => /missing or malformed source end date 2023-03-31-invalid/i.test(error)),
    "source dates must be exact ISO dates rather than strings with a valid ten-character prefix"
  );

  const missingSourceDateErrors = await validateHistoricalSourceLedger(
    [ledgerRow({ sourceProvenance: [provenance({ endDate: "" })] })],
    periodEntries,
    company,
    new Map()
  );
  assert.ok(
    missingSourceDateErrors.some((error) => /missing or malformed source end date \[missing\]/i.test(error)),
    "the model filing's period end cannot be substituted for a missing structured source date"
  );

  const incompatiblePeriodTypeErrors = await validateHistoricalSourceLedger(
    [
      ledgerRow({
        sourceProvenance: [provenance({ periodType: "annual", startDate: "2023-01-01", endDate: "2023-03-31" })]
      })
    ],
    periodEntries,
    company,
    new Map()
  );
  assert.ok(
    incompatiblePeriodTypeErrors.some((error) => /source period type annual.*incompatible with 1Q23/i.test(error)),
    "an annual or YTD duration cannot certify a quarterly model cell merely because its end date matches"
  );

  const missingDerivedOutputErrors = await validateHistoricalSourceLedger(
    [
      ledgerRow({
        value: 60,
        mappingStatus: "validated_current_company_derived_value",
        sourceMappingType: "derived",
        sourceProvenance: [provenance({ concept: "Revenue", value: 100 }), provenance({ concept: "CostOfRevenue", value: -40 })]
      })
    ],
    periodEntries,
    company,
    new Map()
  );
  assert.ok(
    missingDerivedOutputErrors.some((error) => /missing its single explicit final derived output/i.test(error)),
    "derived status must fail closed when no final output is designated"
  );

  const multipleDerivedOutputErrors = await validateHistoricalSourceLedger(
    [
      ledgerRow({
        value: 60,
        mappingStatus: "validated_current_company_derived_value",
        sourceMappingType: "derived",
        sourceProvenance: [
          provenance({ role: "derived_output", concept: "GrossProfitBridge", value: 60, sourceLayer: "derived", accession: "" }),
          provenance({ role: "derived_output", concept: "NestedBridge", value: 40, sourceLayer: "derived", accession: "" }),
          provenance({ concept: "Revenue", value: 100 })
        ]
      })
    ],
    periodEntries,
    company,
    new Map()
  );
  assert.ok(
    multipleDerivedOutputErrors.some((error) => /has 2 final outputs; exactly one is required/i.test(error)),
    "multiple output markers cannot make the validator choose whichever amount happens to match"
  );

  const nestedSubstitutionErrors = await validateHistoricalSourceLedger(
    [
      ledgerRow({
        value: 40,
        mappingStatus: "validated_current_company_derived_value",
        sourceMappingType: "derived",
        sourceProvenance: [
          provenance({ role: "derived_output", concept: "GrossProfitBridge", value: 60, sourceLayer: "derived", accession: "" }),
          provenance({ role: "derived_input", concept: "NestedInterestBridge", value: 40, sourceLayer: "derived", accession: "" }),
          provenance({ concept: "Revenue", value: 100 })
        ]
      })
    ],
    periodEntries,
    company,
    new Map()
  );
  assert.ok(
    nestedSubstitutionErrors.some((error) => /workbook value 40 does not reconcile to final derived output GrossProfitBridge=60mm/i.test(error)),
    "a nested derived input cannot substitute for the one positional final output"
  );

  const otherCompanyPresentationZero = ledgerRow({
    value: 0,
    accessionNumber: "0000000001-23-000001",
    accessionRaw: "0000000001-23-000001",
    accessionNormalized: "000000000123000001",
    mappingStatus: "explicit_zero_no_source_disclosed",
    rawSecValue: "PresentationAbsence:InterestIncomePresentationAbsence=0mm",
    sourceXbrlTag: "InterestIncomePresentationAbsence",
    sourceProvenance: [
      provenance({
        role: "presentation_absence",
        concept: "InterestIncomePresentationAbsence",
        label: "Interest income not separately presented",
        value: 0,
        sourceLayer: "derived",
        accession: "0000000001-23-000001"
      })
    ],
    classificationReason: "Not separately presented."
  });
  const otherCompanyZeroErrors = await validateHistoricalSourceLedger(
    [otherCompanyPresentationZero],
    periodEntries,
    company,
    new Map()
  );
  assert.ok(
    otherCompanyZeroErrors.some((error) => /presentation-backed explicit zero cites accession 000000000123000001 that is not present in the selected company's SEC filing metadata/i.test(error)),
    "presentation absence cannot be certified with an accession outside the selected company's EDGAR filing metadata"
  );

  const savedLlmEnv = {
    OPENROUTER_API_KEY: process.env.OPENROUTER_API_KEY,
    LLM_MAPPING_ENABLED: process.env.LLM_MAPPING_ENABLED,
    LLM_MAPPING_REVIEW_ENABLED: process.env.LLM_MAPPING_REVIEW_ENABLED,
    ALLOW_LEGACY_LLM_WORKBOOK_REVIEW: process.env.ALLOW_LEGACY_LLM_WORKBOOK_REVIEW
  };
  process.env.OPENROUTER_API_KEY = "focused-test-key";
  process.env.LLM_MAPPING_ENABLED = "false";
  process.env.LLM_MAPPING_REVIEW_ENABLED = "true";
  process.env.ALLOW_LEGACY_LLM_WORKBOOK_REVIEW = "true";
  const reviewOnlyState = __fillModelServiceTestHooks.createLlmMappingState();
  const reviewOnlySummary = __fillModelServiceTestHooks.llmMappingStateSummary(reviewOnlyState);
  assert.equal(reviewOnlySummary.enabled, true);
  assert.equal(reviewOnlySummary.mappingEnabled, false);
  assert.equal(reviewOnlySummary.reviewEnabled, true);
  assert.equal(
    __fillModelServiceTestHooks.llmMappingCanUse(reviewOnlyState, 1),
    false,
    "review-only configuration must not accidentally enable line-item mapping calls"
  );
  for (const [key, value] of Object.entries(savedLlmEnv)) {
    if (value === undefined) delete process.env[key];
    else process.env[key] = value;
  }

  console.log("Source-ledger status regression passed.");
})().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
