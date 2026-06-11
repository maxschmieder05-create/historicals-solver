const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const Module = require("node:module");
const ts = require("typescript");

const repoRoot = path.resolve(__dirname, "..");
const sourcePath = path.join(repoRoot, "app", "api", "fill-model", "financial-line-item-classifier.ts");

function loadTypeScriptModule(file) {
  const source = fs.readFileSync(file, "utf8");
  const compiled = ts.transpileModule(source, {
    compilerOptions: {
      module: ts.ModuleKind.CommonJS,
      target: ts.ScriptTarget.ES2020,
      esModuleInterop: true
    }
  }).outputText;
  const mod = new Module(file, module);
  mod.filename = file;
  mod.paths = Module._nodeModulePaths(path.dirname(file));
  mod._compile(compiled, file);
  return mod.exports;
}

const {
  MODEL_ROW_DEFINITIONS,
  classifyFinancialLineItem,
  modelRowDefinitionsForRows,
  modelRowsMatch,
  lineItemNeedsClassification
} = loadTypeScriptModule(sourcePath);

const availableModelRows = Object.keys(MODEL_ROW_DEFINITIONS);
const modelRowDefinitions = modelRowDefinitionsForRows(availableModelRows);

function request(overrides) {
  return {
    company: { name: "Example Corp.", ticker: "EXM" },
    filing: { accession: "0000000000-26-000001", form: "10-Q", filingDate: "2026-05-01" },
    fiscalPeriod: "1Q26",
    statement: overrides.statement ?? "balance_sheet",
    sourceTableType: overrides.sourceTableType ?? "primary_statement",
    reportedLineItemLabel: overrides.label,
    cleanLabel: overrides.label,
    xbrlTag: overrides.xbrlTag,
    amount: overrides.amount ?? 100,
    unit: "USD",
    periodType: overrides.periodType ?? "instant",
    section: overrides.section,
    nearbyRows: overrides.nearbyRows ?? [],
    parentSubtotal: overrides.parentSubtotal,
    isSubtotal: false,
    availableModelRows: overrides.availableModelRows ?? availableModelRows,
    modelRowDefinitions,
    deterministicCandidate: overrides.deterministicCandidate,
    uncertaintyReason: overrides.uncertaintyReason ?? "ambiguous accounting label"
  };
}

async function classify(overrides) {
  return classifyFinancialLineItem(request(overrides), { llm: { enabled: false } });
}

(async () => {
  const convertibleNotes = await classify({
    label: "Short-term convertible senior notes",
    section: "current liabilities",
    deterministicCandidate: "Revolver"
  });
  assert.equal(convertibleNotes.recommended_model_row, "LT Debt (Incl. Current Portion)");
  assert.equal(convertibleNotes.is_debt, true);
  assert.equal(convertibleNotes.mapping_passed_validation, true);

  const currentMaturities = await classify({
    label: "Current maturities of long-term debt",
    section: "current liabilities",
    deterministicCandidate: "Revolver"
  });
  assert.equal(currentMaturities.recommended_model_row, "LT Debt (Incl. Current Portion)");

  const shortTermBorrowings = await classify({
    label: "Short-term borrowings",
    section: "current liabilities",
    deterministicCandidate: "Revolver"
  });
  assert.equal(shortTermBorrowings.recommended_model_row, "Revolver");

  const deferredIncome = await classify({
    label: "Deferred income",
    section: "current liabilities",
    deterministicCandidate: "Deferred Income Taxes"
  });
  assert.equal(deferredIncome.recommended_model_row, "Other Current Liabilities");
  assert.equal(deferredIncome.is_deferred_revenue_or_contract_liability, true);
  assert.equal(deferredIncome.is_deferred_tax, false);

  const deferredTax = await classify({
    label: "Deferred tax liabilities",
    section: "non-current liabilities",
    deterministicCandidate: "Other Non-Current Liabilities"
  });
  assert.equal(deferredTax.recommended_model_row, "Deferred Income Taxes");
  assert.equal(deferredTax.is_deferred_tax, true);

  const supplies = await classify({
    label: "Aircraft fuel, spare parts and supplies",
    section: "current assets",
    deterministicCandidate: "Prepaid & Other Current Assets"
  });
  assert.equal(supplies.recommended_model_row, "Inventory");

  const shortTermInvestments = await classify({
    label: "Short-term investments",
    section: "current assets",
    deterministicCandidate: "Prepaid & Other Current Assets",
    availableModelRows: availableModelRows.filter((row) => !/short[-\s]?term investments?|current investments?|marketable securities/i.test(row))
  });
  assert.equal(shortTermInvestments.recommended_model_row, "Cash & Cash Equivalents");

  const iprd = await classify({
    label: "Acquired in-process research and development",
    statement: "income_statement",
    section: "operating expenses",
    periodType: "duration",
    deterministicCandidate: "D&A"
  });
  assert.notEqual(iprd.recommended_model_row, "D&A");

  const cashFlowDa = await classify({
    label: "Depreciation and amortization",
    statement: "cash_flow",
    sourceTableType: "cash_flow_reconciliation",
    section: "unknown",
    periodType: "duration",
    deterministicCandidate: "D&A"
  });
  assert.equal(cashFlowDa.mapping_passed_validation, false);
  assert.equal(modelRowsMatch(cashFlowDa.recommended_model_row, "D&A"), false);

  const specialItems = await classify({
    label: "Special items, restructuring, impairment, and other charges",
    statement: "income_statement",
    section: "operating expenses",
    periodType: "duration",
    deterministicCandidate: "Other Non-Operating Income / Expense"
  });
  assert.equal(specialItems.recommended_model_row, "Other Operating Income / Expense");

  assert.equal(lineItemNeedsClassification(request({ label: "Deferred income", section: "current liabilities" })), true);

  const llmRequestedPayloads = [];
  const ambiguousLease = await classifyFinancialLineItem(
    request({
      label: "Other lease financing obligations",
      section: "non-current liabilities",
      deterministicCandidate: "Other Non-Current Liabilities",
      uncertaintyReason: "lease financing may represent debt or another long-term liability"
    }),
    {
      llm: {
        enabled: true,
        apiKey: "test-key",
        endpoint: "https://example.test/chat/completions",
        model: "test-model",
        siteUrl: "http://localhost:3000",
        appTitle: "Historicals Solver Test",
        timeoutMs: 100,
        fetchImpl: async (_url, init) => {
          const body = JSON.parse(init.body);
          llmRequestedPayloads.push(body);
          return {
            ok: true,
            json: async () => ({
              choices: [
                {
                  message: {
                    content: JSON.stringify({
                      source_line_item: "Other lease financing obligations",
                      recommended_model_row: "LT Debt (Incl. Current Portion)",
                      classification_type: "lease financing debt obligation",
                      is_current: false,
                      is_debt: true,
                      is_operating: false,
                      is_tax_related: false,
                      is_deferred_revenue_or_contract_liability: false,
                      is_deferred_tax: false,
                      is_subtotal: false,
                      should_exclude_from_other_bucket: true,
                      confidence: "high",
                      reason: "Lease financing obligations are debt-like capital structure liabilities.",
                      requires_validation: true
                    })
                  }
                }
              ]
            })
          };
        }
      }
    }
  );
  assert.equal(ambiguousLease.llm_used, true);
  assert.equal(ambiguousLease.recommended_model_row, "LT Debt (Incl. Current Portion)");
  assert.equal(llmRequestedPayloads.length, 1);
  assert.equal(llmRequestedPayloads[0].response_format.type, "json_schema");
  const llmUserPayload = JSON.parse(llmRequestedPayloads[0].messages[1].content);
  assert.equal(llmUserPayload.reportedLineItemLabel, "Other lease financing obligations");
  assert.equal(llmUserPayload.modelRowDefinitions["LT Debt (Incl. Current Portion)"].includes("Long-term debt instruments"), true);

  console.log("Financial line item classifier rules passed.");
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
