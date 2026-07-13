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

const {
  authorizeCellAssignment,
  normalizeSourceFact,
  validateAuthorizedCellAssignments
} = require(path.join(repoRoot, "server", "fill-model", "normalized-layer.ts"));
const {
  findApprovedMapping
} = require(path.join(repoRoot, "server", "fill-model", "approved-mapping-cache.ts"));
const {
  quarterlyFlowDerivationAllowed,
  quarterlyFlowInputsCompatible
} = require(path.join(repoRoot, "server", "fill-model", "fiscal-period.ts"));
const {
  classifyFinancialLineItem,
  modelRowDefinitionsForRows,
  MODEL_ROW_DEFINITIONS
} = require(path.join(repoRoot, "server", "fill-model", "financial-line-item-classifier.ts"));

function normalizedFact(overrides = {}) {
  return normalizeSourceFact({
    concept: "RevenueFromContractWithCustomerExcludingAssessedTax",
    label: "Revenue",
    value: 100_000_000,
    taxonomy: "us-gaap",
    unit: "USD",
    cik: "0000000001",
    accession: "0000000001-26-000001",
    start: "2026-01-01",
    end: "2026-03-31",
    period: "1Q26",
    sourcePeriod: "1Q26",
    periodType: "quarterly",
    sourceLayer: "sec_filing_package",
    ...overrides
  });
}

function assignment(overrides = {}) {
  return authorizeCellAssignment({
    sheetName: "Model",
    cell: "F20",
    modelCategory: "revenue",
    modelRow: "Revenue",
    statement: "income",
    expectedPeriodType: "duration",
    expectedUnitFamilies: ["currency"],
    period: "1Q26",
    value: 100,
    facts: [normalizedFact()],
    mappingType: "direct",
    formulaPolicy: "hardcode",
    projected: false,
    reportedPeriod: true,
    ...overrides
  });
}

const authorized = assignment();
assert.equal(authorized.authorized, true);
assert.equal(authorized.facts[0].unitFamily, "currency");
assert.equal(authorized.directSourceFactIds.length, 1);

assert.equal(
  assignment({
    expectedPeriodType: "instant",
    modelCategory: "assets",
    modelRow: "Total Assets"
  }).authorized,
  false,
  "duration facts must never be authorized for balance-sheet cells"
);

assert.equal(
  assignment({
    projected: true,
    formulaPolicy: "protected"
  }).authorized,
  false,
  "projected/protected formula cells must fail closed"
);

const secondDirectUse = assignment({
  cell: "F21",
  modelCategory: "other_income",
  modelRow: "Other Income"
});
assert.match(
  validateAuthorizedCellAssignments([authorized, secondDirectUse]).join(" "),
  /directly assigned to both/i,
  "one source fact must not be directly allocated to two additive model lines"
);

assert.equal(
  quarterlyFlowDerivationAllowed("RevenueFromContractWithCustomerExcludingAssessedTax", {
    unit: "USD",
    periodType: "annual"
  }),
  true
);
assert.equal(
  quarterlyFlowDerivationAllowed("Assets", { unit: "USD", periodType: "instant" }),
  false,
  "Q4 must never be derived for balance-sheet instants"
);
assert.equal(
  quarterlyFlowDerivationAllowed("WeightedAverageNumberOfSharesOutstandingBasic", {
    unit: "shares",
    periodType: "annual"
  }),
  false,
  "non-additive share counts must not be quarter-subtracted"
);
assert.equal(
  quarterlyFlowInputsCompatible([
    { unit: "USD", cik: "1", periodType: "annual" },
    { unit: "shares", cik: "1", periodType: "quarterly" }
  ]),
  false,
  "quarter derivations must not mix units"
);

const approvedMappings = {
  version: 1,
  mappings: [
    {
      scope: "company_historical",
      companyTicker: "TEST",
      statement: "balance_sheet",
      section: "current assets",
      xbrlTag: "OtherCurrentAssetsExtension",
      modelRow: "Inventory",
      action: "map",
      explanation: "Company history",
      approvedBy: "analyst@example.com",
      approvedAt: "2026-06-01T00:00:00.000Z"
    },
    {
      scope: "approved_exact",
      statement: "balance_sheet",
      section: "current assets",
      xbrlTag: "OtherCurrentAssetsExtension",
      modelRow: "Prepaid & Other Current Assets",
      action: "merge_into_other",
      explanation: "Exact concept and statement-context approval",
      approvedBy: "controller@example.com",
      approvedAt: "2026-06-02T00:00:00.000Z"
    }
  ]
};

const lookup = findApprovedMapping(approvedMappings, {
  company: { name: "Test Company", ticker: "TEST" },
  statement: "balance_sheet",
  sourceTableType: "primary_statement",
  section: "current assets",
  xbrlTag: "OtherCurrentAssetsExtension",
  reportedLabel: "Other current assets",
  availableModelRows: ["Inventory", "Prepaid & Other Current Assets"]
});
assert.equal(lookup.scope, "approved_exact", "approved exact mappings must precede company history");

const availableModelRows = Object.keys(MODEL_ROW_DEFINITIONS);
const request = {
  company: { name: "Test Company", ticker: "TEST" },
  filing: { accession: "0000000001-26-000001" },
  fiscalPeriod: "1Q26",
  statement: "balance_sheet",
  sourceTableType: "primary_statement",
  reportedLineItemLabel: "Other current assets",
  cleanLabel: "Other current assets",
  xbrlTag: "OtherCurrentAssetsExtension",
  amount: 25_000_000,
  unit: "USD",
  periodType: "instant",
  section: "current assets",
  nearbyRows: [],
  isSubtotal: false,
  availableModelRows,
  modelRowDefinitions: modelRowDefinitionsForRows(availableModelRows),
  uncertaintyReason: "Extension concept requires classification."
};

(async () => {
  let llmCalls = 0;
  const classification = await classifyFinancialLineItem(request, {
    approvedMappings,
    llm: {
      enabled: true,
      apiKey: "unused",
      endpoint: "https://example.invalid",
      model: "unused",
      siteUrl: "http://localhost",
      appTitle: "test",
      fetchImpl: async () => {
        llmCalls += 1;
        throw new Error("Approved mappings must bypass the LLM.");
      }
    }
  });
  assert.equal(classification.recommended_model_row, "Prepaid & Other Current Assets");
  assert.equal(classification.classification_type, "approved exact mapping");
  assert.equal(classification.mapping_passed_validation, true);
  assert.equal(classification.llm_used, false);
  assert.equal(llmCalls, 0);
  console.log("Normalized architecture and approved mapping hierarchy checks passed.");
})().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
