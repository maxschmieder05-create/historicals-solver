const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const Module = require("node:module");
const ts = require("typescript");

const repoRoot = path.resolve(__dirname, "..");
const sourcePath = path.join(repoRoot, "server", "fill-model", "financial-line-item-classifier.ts");

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

const {
  MODEL_ROW_DEFINITIONS,
  classifyFinancialLineItem,
  classifyFinancialStatementLineItems,
  classificationModelRowAssignmentForPrimaryStatement,
  modelRowDefinitionsForRows,
  modelRowsMatch,
  lineItemNeedsClassification
} = loadTypeScriptModule(sourcePath);

const availableModelRows = Object.keys(MODEL_ROW_DEFINITIONS);
const modelRowDefinitions = modelRowDefinitionsForRows(availableModelRows);
assert.equal(modelRowsMatch("Research & Development (R&D)", "R&D"), true);
assert.equal(modelRowsMatch("Selling, General & Administration (SG&A)", "SG&A"), true);
assert.equal(
  classificationModelRowAssignmentForPrimaryStatement(
    {
      source_line_item: "Research and development",
      recommended_action: "map",
      recommended_model_row: "R&D",
      recommended_model_row_mappings: [],
      explicit_zero_rows: [],
      classification_type: "research and development",
      is_current: null,
      is_debt: false,
      is_operating: true,
      is_tax_related: false,
      is_deferred_revenue_or_contract_liability: false,
      is_deferred_tax: false,
      is_subtotal: false,
      should_exclude_from_other_bucket: false,
      confidence: "high",
      reason: "R&D maps to the workbook R&D row.",
      requires_validation: false,
      requires_revalidation: false,
      llm_used: false,
      mapping_passed_validation: true
    },
    ["Research & Development (R&D)"]
  ).modelRow,
  "Research & Development (R&D)"
);
assert.equal(
  classificationModelRowAssignmentForPrimaryStatement(
    {
      source_line_item: "Marketing, general and administrative",
      recommended_action: "map",
      recommended_model_row: "SG&A",
      recommended_model_row_mappings: [],
      explicit_zero_rows: [],
      classification_type: "selling general and administrative",
      is_current: null,
      is_debt: false,
      is_operating: true,
      is_tax_related: false,
      is_deferred_revenue_or_contract_liability: false,
      is_deferred_tax: false,
      is_subtotal: false,
      should_exclude_from_other_bucket: false,
      confidence: "high",
      reason: "SG&A maps to the workbook SG&A row.",
      requires_validation: false,
      requires_revalidation: false,
      llm_used: false,
      mapping_passed_validation: true
    },
    ["Selling, General & Administration (SG&A)"]
  ).modelRow,
  "Selling, General & Administration (SG&A)"
);

function request(overrides) {
  return {
    company: { name: "Example Corp.", ticker: "EXM" },
    filing: { accession: "0000000000-26-000001", form: "10-Q", filingDate: "2026-05-01" },
    fiscalPeriod: "1Q26",
    statement: overrides.statement ?? "balance_sheet",
    sourceTableType: overrides.sourceTableType ?? "primary_statement",
    sourceRowKey: overrides.sourceRowKey,
    rowOrder: overrides.rowOrder,
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
    priorPeriodSourceLabels: overrides.priorPeriodSourceLabels ?? [],
    currentPeriodSourceLines: overrides.currentPeriodSourceLines ?? [],
    availableModelRows: overrides.availableModelRows ?? availableModelRows,
    modelRowDefinitions,
    deterministicCandidate: overrides.deterministicCandidate,
    uncertaintyReason: overrides.uncertaintyReason ?? "ambiguous accounting label",
    validationError: overrides.validationError ?? "",
    alreadyMappedRows: overrides.alreadyMappedRows ?? []
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
  assert.equal(
    classificationModelRowAssignmentForPrimaryStatement(currentMaturities, availableModelRows).modelRow,
    "LT Debt (Incl. Current Portion)"
  );

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

  const longTermUnearnedRevenue = await classify({
    label: "Long-term unearned revenue",
    xbrlTag: "ContractWithCustomerLiabilityNoncurrent",
    section: "current liabilities",
    deterministicCandidate: "Other Current Liabilities"
  });
  assert.equal(longTermUnearnedRevenue.recommended_model_row, "Other Non-Current Liabilities");
  assert.equal(longTermUnearnedRevenue.is_current, false);

  const longTermIncomeTaxes = await classify({
    label: "Long-term income taxes",
    xbrlTag: "AccruedIncomeTaxesNoncurrent",
    section: "current liabilities",
    deterministicCandidate: "Accrued Liabilities"
  });
  assert.equal(longTermIncomeTaxes.recommended_model_row, "Other Non-Current Liabilities");
  assert.equal(longTermIncomeTaxes.is_current, false);
  assert.equal(longTermIncomeTaxes.is_tax_related, true);

  const shortTermIncomeTaxes = await classify({
    label: "Short-term income taxes",
    xbrlTag: "AccruedIncomeTaxesCurrent",
    section: "current liabilities",
    deterministicCandidate: "Other Current Liabilities"
  });
  assert.equal(shortTermIncomeTaxes.recommended_model_row, "Accrued Liabilities");
  assert.equal(shortTermIncomeTaxes.is_current, true);
  assert.equal(shortTermIncomeTaxes.is_tax_related, true);

  const deferredTax = await classify({
    label: "Deferred tax liabilities",
    section: "non-current liabilities",
    deterministicCandidate: "Other Non-Current Liabilities"
  });
  assert.equal(deferredTax.recommended_model_row, "Deferred Income Taxes");
  assert.equal(deferredTax.is_deferred_tax, true);

  const deferredTaxAssetUnknownSection = await classify({
    label: "Deferred income tax assets",
    xbrlTag: "DeferredIncomeTaxAssetsNet",
    section: "unknown",
    deterministicCandidate: "Deferred Income Taxes"
  });
  assert.equal(deferredTaxAssetUnknownSection.recommended_model_row, "Other Non-Current Assets");
  assert.equal(deferredTaxAssetUnknownSection.is_deferred_tax, true);

  const supplies = await classify({
    label: "Aircraft fuel, spare parts and supplies",
    section: "current assets",
    deterministicCandidate: "Prepaid & Other Current Assets"
  });
  assert.equal(supplies.recommended_model_row, "Inventory");

  const workInProcess = await classify({
    label: "Work-in-process",
    section: "unknown",
    deterministicCandidate: "Unmapped / Needs Review"
  });
  assert.equal(workInProcess.recommended_model_row, "Inventory");

  const finishedGoods = await classify({
    label: "Finished goods",
    section: "unknown",
    deterministicCandidate: "Unmapped / Needs Review"
  });
  assert.equal(finishedGoods.recommended_model_row, "Inventory");

  const landBuildings = await classify({
    label: "Land, buildings and improvements",
    section: "unknown",
    deterministicCandidate: "Unmapped / Needs Review"
  });
  assert.equal(landBuildings.recommended_model_row, "PP&E, Net");

  const machineryEquipment = await classify({
    label: "Machinery and equipment",
    section: "unknown",
    deterministicCandidate: "Unmapped / Needs Review"
  });
  assert.equal(machineryEquipment.recommended_model_row, "PP&E, Net");

  const shortTermInvestments = await classify({
    label: "Short-term investments",
    section: "current assets",
    deterministicCandidate: "Prepaid & Other Current Assets",
    availableModelRows: availableModelRows.filter((row) => !/short[-\s]?term investments?|current investments?|marketable securities/i.test(row))
  });
  assert.equal(shortTermInvestments.recommended_model_row, "Prepaid & Other Current Assets");
  assert.equal(
    classificationModelRowAssignmentForPrimaryStatement(shortTermInvestments, availableModelRows).modelRow,
    "Prepaid & Other Current Assets"
  );

  const marketableSecurities = await classify({
    label: "Marketable securities",
    section: "current assets",
    deterministicCandidate: "Prepaid & Other Current Assets",
    availableModelRows: availableModelRows.filter((row) => !/short[-\s]?term investments?|current investments?|marketable securities/i.test(row))
  });
  assert.equal(marketableSecurities.recommended_model_row, "Prepaid & Other Current Assets");

  const iprd = await classify({
    label: "Acquired in-process research and development",
    statement: "income_statement",
    section: "operating expenses",
    periodType: "duration",
    deterministicCandidate: "D&A"
  });
  assert.notEqual(iprd.recommended_model_row, "D&A");

  const iprdImpairment = await classify({
    label: "In-process research and development impairments",
    statement: "income_statement",
    section: "operating expenses",
    periodType: "duration",
    deterministicCandidate: "R&D"
  });
  assert.equal(iprdImpairment.recommended_model_row, "Other Operating Income / Expense");

  const rdExcludingAcquiredInProcessCost = await classify({
    label: "Research and Development Expense (Excluding Acquired in Process Cost)",
    xbrlTag: "ResearchAndDevelopmentExpenseExcludingAcquiredInProcessCost",
    statement: "income_statement",
    section: "operating expenses",
    periodType: "duration",
    deterministicCandidate: "R&D"
  });
  assert.equal(rdExcludingAcquiredInProcessCost.recommended_model_row, "R&D");

  const belowOperatingRestructuring = await classify({
    label: "Restructuring",
    statement: "income_statement",
    section: "below operating income",
    periodType: "duration",
    deterministicCandidate: "Other Operating Income / Expense"
  });
  assert.equal(belowOperatingRestructuring.recommended_model_row, "Other Non-Operating Income / Expense");

  const accruedRebates = await classify({
    label: "Accrued rebates, returns and promotions",
    statement: "balance_sheet",
    section: "current liabilities",
    periodType: "instant",
    deterministicCandidate: "Other Current Liabilities"
  });
  assert.equal(accruedRebates.recommended_model_row, "Accrued Liabilities");

  const exactAccruedLiabilities = await classify({
    label: "Accrued Liabilities",
    xbrlTag: "AccruedLiabilitiesCurrent",
    statement: "balance_sheet",
    section: "current liabilities",
    periodType: "instant",
    deterministicCandidate: "Other Current Liabilities",
    uncertaintyReason: "Other buckets are allowed only when no better dedicated model row exists."
  });
  assert.equal(exactAccruedLiabilities.recommended_model_row, "Accrued Liabilities");
  assert.equal(exactAccruedLiabilities.mapping_passed_validation, true);

  const exactAccruedCurrentAndNoncurrentTag = await classify({
    label: "Accrued Liabilities",
    xbrlTag: "AccruedLiabilitiesCurrentAndNoncurrent",
    statement: "balance_sheet",
    section: "current liabilities",
    periodType: "instant",
    deterministicCandidate: "Other Current Liabilities",
    uncertaintyReason: "XBRL concept says current and noncurrent, but the primary statement places the line in current liabilities."
  });
  assert.equal(exactAccruedCurrentAndNoncurrentTag.recommended_model_row, "Accrued Liabilities");
  assert.equal(exactAccruedCurrentAndNoncurrentTag.mapping_passed_validation, true);

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

  const advertising = await classify({
    label: "Advertising expense",
    xbrlTag: "AdvertisingExpense",
    statement: "income_statement",
    section: "operating expenses",
    periodType: "duration",
    uncertaintyReason: ""
  });
  assert.equal(advertising.recommended_model_row, "SG&A");
  assert.equal(advertising.mapping_passed_validation, true);

  const salesAndMarketing = await classify({
    label: "Sales and marketing",
    statement: "income_statement",
    section: "operating expenses",
    periodType: "duration",
    uncertaintyReason: ""
  });
  assert.equal(salesAndMarketing.recommended_model_row, "SG&A");

  const costOfSales = await classify({
    label: "Cost of sales",
    statement: "income_statement",
    section: "operating expenses",
    periodType: "duration",
    uncertaintyReason: ""
  });
  assert.equal(costOfSales.recommended_model_row, "COGS / Cost of Goods Sold");

  assert.equal(lineItemNeedsClassification(request({ label: "Deferred income", section: "current liabilities" })), true);
  assert.equal(lineItemNeedsClassification(request({ label: "Marketable securities", section: "current assets" })), true);
  assert.equal(
    lineItemNeedsClassification(
      request({
        label: "Advertising expense",
        statement: "income_statement",
        section: "operating expenses",
        periodType: "duration",
        uncertaintyReason: ""
      })
    ),
    true
  );

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
                      recommended_action: "remap",
                      recommended_model_row: "LT Debt (Incl. Current Portion)",
                      recommended_model_row_mappings: [
                        {
                          source_line_item: "Other lease financing obligations",
                          model_row: "LT Debt (Incl. Current Portion)",
                          amount: 100,
                          reason: "Lease financing obligations are debt-like capital structure liabilities."
                        }
                      ],
                      explicit_zero_rows: [],
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
                      requires_validation: true,
                      requires_revalidation: true
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
  assert.equal(llmRequestedPayloads[0].response_format.json_schema.schema.properties.recommended_action.enum.includes("set_zero"), true);

  const statementBatchPayloads = [];
  const batchResult = await classifyFinancialStatementLineItems(
    [
      request({
        sourceRowKey: "row-cash",
        rowOrder: 1,
        label: "Cash and cash equivalents",
        xbrlTag: "CashAndCashEquivalentsAtCarryingValue",
        section: "current assets",
        uncertaintyReason: ""
      }),
      request({
        sourceRowKey: "row-investments",
        rowOrder: 2,
        label: "Short-term investments",
        xbrlTag: "ShortTermInvestments",
        section: "current assets",
        deterministicCandidate: "Prepaid & Other Current Assets"
      }),
      request({
        sourceRowKey: "row-advertising",
        rowOrder: 7,
        label: "Advertising expense",
        xbrlTag: "AdvertisingExpense",
        statement: "income_statement",
        section: "operating expenses",
        periodType: "duration",
        uncertaintyReason: ""
      }),
      request({
        sourceRowKey: "row-lease-financing",
        rowOrder: 12,
        label: "Other lease financing obligations",
        xbrlTag: "OtherLeaseFinancingObligations",
        section: "non-current liabilities",
        deterministicCandidate: "Other Non-Current Liabilities",
        uncertaintyReason: "lease financing may represent debt or another long-term liability"
      }),
      request({
        sourceRowKey: "row-notes-payable",
        rowOrder: 13,
        label: "Notes payable",
        xbrlTag: "NotesPayableCurrent",
        section: "current liabilities",
        deterministicCandidate: "Other Current Liabilities",
        uncertaintyReason: "notes payable could be current borrowings even if reported below other liabilities"
      })
    ],
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
          statementBatchPayloads.push(body);
          return {
            ok: true,
            json: async () => ({
              choices: [
                {
                  message: {
                    content: JSON.stringify({
                      classifications: [
                        {
                          source_row_key: "row-investments",
                          source_line_item: "Short-term investments",
                          recommended_action: "remap",
                          recommended_model_row: "Cash & Cash Equivalents",
                          recommended_model_row_mappings: [],
                          explicit_zero_rows: [],
                          classification_type: "cash and marketable securities current investment",
                          is_current: true,
                          is_debt: false,
                          is_operating: false,
                          is_tax_related: false,
                          is_deferred_revenue_or_contract_liability: false,
                          is_deferred_tax: false,
                          is_subtotal: false,
                          should_exclude_from_other_bucket: true,
                          confidence: "high",
                          reason: "Whole-statement context shows no dedicated current investments row, so short-term investments group with cash.",
                          requires_validation: true,
                          requires_revalidation: true
                        },
                        {
                          source_row_key: "row-lease-financing",
                          source_line_item: "Other lease financing obligations",
                          recommended_action: "remap",
                          recommended_model_row: "LT Debt (Incl. Current Portion)",
                          recommended_model_row_mappings: [],
                          explicit_zero_rows: [],
                          classification_type: "debt-like financing obligation",
                          is_current: false,
                          is_debt: true,
                          is_operating: false,
                          is_tax_related: false,
                          is_deferred_revenue_or_contract_liability: false,
                          is_deferred_tax: false,
                          is_subtotal: false,
                          should_exclude_from_other_bucket: true,
                          confidence: "high",
                          reason: "Whole-statement context shows this is a financing obligation below operating liabilities.",
                          requires_validation: true,
                          requires_revalidation: true
                        },
                        {
                          source_row_key: "row-notes-payable",
                          source_line_item: "Notes payable",
                          recommended_action: "remap",
                          recommended_model_row: "Revolver",
                          recommended_model_row_mappings: [],
                          explicit_zero_rows: [],
                          classification_type: "current borrowing facility",
                          is_current: true,
                          is_debt: true,
                          is_operating: false,
                          is_tax_related: false,
                          is_deferred_revenue_or_contract_liability: false,
                          is_deferred_tax: false,
                          is_subtotal: false,
                          should_exclude_from_other_bucket: true,
                          confidence: "high",
                          reason: "Notes payable current is a current borrowing source and belongs with Revolver/current borrowings.",
                          requires_validation: true,
                          requires_revalidation: true
                        }
                      ]
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
  assert.equal(statementBatchPayloads.length, 1);
  const statementPayload = JSON.parse(statementBatchPayloads[0].messages[1].content);
  assert.deepEqual(statementPayload.targetSourceRowKeys, ["row-investments", "row-lease-financing", "row-notes-payable"]);
  assert.equal(statementPayload.statementRows.length, 5);
  assert.equal(statementPayload.statementRows.some((row) => row.sourceRowKey === "row-cash" && row.target === false), true);
  assert.equal(statementPayload.statementRows.some((row) => row.sourceRowKey === "row-investments" && row.target === true && row.deterministicClassification), true);
  assert.equal(
    statementPayload.statementRows.some(
      (row) => row.sourceRowKey === "row-advertising" && row.target === false && row.deterministicClassification?.recommendedModelRow === "SG&A"
    ),
    true
  );
  assert.equal(statementPayload.modelRowDefinitions["SG&A"].includes("advertising"), true);
  assert.equal(statementBatchPayloads[0].response_format.json_schema.schema.properties.classifications.items.properties.source_row_key.type, "string");
  assert.equal(batchResult.llmCalls, 1);
  const investmentsClassification = batchResult.classifications.find((item) => item.request.sourceRowKey === "row-investments").classification;
  assert.equal(investmentsClassification.recommended_model_row, "Cash & Cash Equivalents");
  assert.equal(investmentsClassification.llm_used, true);
  assert.equal(batchResult.classifications.find((item) => item.request.sourceRowKey === "row-advertising").classification.recommended_model_row, "SG&A");
  assert.equal(batchResult.classifications.find((item) => item.request.sourceRowKey === "row-notes-payable").classification.recommended_model_row, "Revolver");
  const leaseFinancingClassification = batchResult.classifications.find((item) => item.request.sourceRowKey === "row-lease-financing").classification;
  const leaseFinancingAssignment = classificationModelRowAssignmentForPrimaryStatement(leaseFinancingClassification, availableModelRows);
  assert.equal(leaseFinancingAssignment.modelRow, "LT Debt (Incl. Current Portion)");
  assert.equal(leaseFinancingAssignment.llmUsed, true);
  assert.equal(leaseFinancingAssignment.reason.startsWith("LLM line-item classification:"), true);

  const lowConfidenceAssignment = classificationModelRowAssignmentForPrimaryStatement(
    {
      ...leaseFinancingClassification,
      confidence: "low"
    },
    availableModelRows
  );
  assert.equal(lowConfidenceAssignment, null);
  const deterministicAssignment = classificationModelRowAssignmentForPrimaryStatement(
    {
      ...leaseFinancingClassification,
      llm_used: false
    },
    availableModelRows
  );
  assert.equal(deterministicAssignment.modelRow, "LT Debt (Incl. Current Portion)");
  assert.equal(deterministicAssignment.llmUsed, false);
  assert.equal(deterministicAssignment.reason.startsWith("Validated line-item classification:"), true);

  console.log("Financial line item classifier rules passed.");
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
