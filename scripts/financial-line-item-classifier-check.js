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
  classificationSourceKeys,
  classificationPassesValidation,
  classificationModelRowAssignmentForPrimaryStatement,
  fullStatementLineItemNeedsAnalystPass,
  materialStatementLineItemNeedsAnalystPass,
  modelRowDefinitionsForRows,
  modelRowsMatch,
  lineItemNeedsClassification
} = loadTypeScriptModule(sourcePath);

const availableModelRows = Object.keys(MODEL_ROW_DEFINITIONS);
const modelRowDefinitions = modelRowDefinitionsForRows(availableModelRows);
const scopedClassificationKeys = classificationSourceKeys({
  period: "1Q26",
  accession: "0000000000-26-000001",
  xbrlTag: "ExampleExtensionConcept",
  label: "Example extension line",
  amount: 100
});
assert.ok(
  scopedClassificationKeys.every((key) => /1q26|000000000026000001/.test(key)),
  "Classification keys must remain period/accession scoped."
);
assert.ok(
  !scopedClassificationKeys.some((key) => key === "concept-label|exampleextensionconcept|exampleextensionline"),
  "A generic concept/label key must not leak one filing's classification into another period."
);
const firstSameConceptPresentationKeys = classificationSourceKeys({
  period: "FY25",
  accession: "0000000000-25-000001",
  xbrlTag: "UtilityOperatingCosts",
  label: "Purchased power",
  amount: 250
});
const secondSameConceptPresentationKeys = classificationSourceKeys({
  period: "FY25",
  accession: "0000000000-25-000001",
  xbrlTag: "UtilityOperatingCosts",
  label: "Operations and maintenance",
  amount: 250
});
const sameConceptPresentationStore = new Map();
firstSameConceptPresentationKeys.forEach((key) =>
  sameConceptPresentationStore.set(key, { recommended_model_row: "COGS / Cost of Goods Sold" })
);
const inheritedSameConceptDecision = secondSameConceptPresentationKeys
  .map((key) => sameConceptPresentationStore.get(key))
  .find(Boolean);
assert.equal(
  inheritedSameConceptDecision,
  undefined,
  "Distinct presentation rows using the same period and XBRL concept must not inherit each other's classification decision."
);
assert.equal(modelRowsMatch("Research & Development (R&D)", "R&D"), true);
assert.equal(modelRowsMatch("Selling, General & Administration (SG&A)", "SG&A"), true);
assert.equal(
  materialStatementLineItemNeedsAnalystPass(
    request({
      label: "Cash and cash equivalents",
      xbrlTag: "CashAndCashEquivalentsAtCarryingValue",
      section: "current assets",
      amount: 1_000_000,
      uncertaintyReason: ""
    })
  ),
  true
);
assert.equal(
  materialStatementLineItemNeedsAnalystPass(
    request({
      label: "Cash and cash equivalents",
      xbrlTag: "CashAndCashEquivalentsAtCarryingValue",
      section: "current assets",
      amount: 100,
      uncertaintyReason: ""
    })
  ),
  false
);
assert.equal(
  classificationModelRowAssignmentForPrimaryStatement(
    {
      source_line_item: "Interest income",
      recommended_action: "map",
      recommended_model_row: "Interest Income",
      recommended_model_row_mappings: [],
      explicit_zero_rows: [],
      classification_type: "interest income",
      is_current: null,
      is_debt: false,
      is_operating: false,
      is_tax_related: false,
      is_deferred_revenue_or_contract_liability: false,
      is_deferred_tax: false,
      is_subtotal: false,
      should_exclude_from_other_bucket: false,
      confidence: "high",
      reason: "Standalone interest income.",
      requires_validation: false,
      requires_revalidation: false,
      llm_used: false,
      mapping_passed_validation: true
    },
    ["Revenue", "Interest Expense", "Net Revenue"]
  ),
  null
);
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
    unit: overrides.unit ?? "USD",
    periodType: overrides.periodType ?? "instant",
    section: overrides.section,
    nearbyRows: overrides.nearbyRows ?? [],
    parentSubtotal: overrides.parentSubtotal,
    isSubtotal: overrides.isSubtotal ?? false,
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
  const trustedValidationCases = [
    {
      name: "R&D operating expense cannot map to Revenue even when prior validation and LLM flags claim success",
      request: request({
        label: "Research and development expense",
        xbrlTag: "ResearchAndDevelopmentExpense",
        statement: "income_statement",
        section: "operating expenses",
        periodType: "duration"
      }),
      row: "Revenue",
      expected: false
    },
    {
      name: "R&D operating expense maps to R&D",
      request: request({
        label: "Research and development expense",
        xbrlTag: "ResearchAndDevelopmentExpense",
        statement: "income_statement",
        section: "operating expenses",
        periodType: "duration"
      }),
      row: "R&D",
      expected: true
    },
    {
      name: "pension obligations cannot self-certify as deferred taxes through an LLM flag",
      request: request({
        label: "Pension obligations",
        xbrlTag: "PensionLiabilitiesNoncurrent",
        section: "non-current liabilities"
      }),
      row: "Deferred Income Taxes",
      expected: false
    },
    {
      name: "pension obligations fall back to other non-current liabilities when no pension row exists",
      request: request({
        label: "Pension obligations",
        xbrlTag: "PensionLiabilitiesNoncurrent",
        section: "non-current liabilities"
      }),
      row: "Other Non-Current Liabilities",
      expected: true
    },
    {
      name: "true deferred tax liabilities pass even when the LLM flag is false",
      request: request({
        label: "Deferred tax liabilities, non-current",
        xbrlTag: "DeferredTaxLiabilitiesNoncurrent",
        section: "non-current liabilities"
      }),
      row: "Deferred Income Taxes",
      expected: true
    },
    {
      name: "contract liabilities cannot map to deferred taxes",
      request: request({
        label: "Contract liabilities, non-current",
        xbrlTag: "ContractWithCustomerLiabilityNoncurrent",
        section: "non-current liabilities"
      }),
      row: "Deferred Income Taxes",
      expected: false
    },
    {
      name: "cash-flow reconciliation D&A cannot populate income-statement D&A",
      request: request({
        label: "Depreciation and amortization",
        xbrlTag: "DepreciationDepletionAndAmortization",
        statement: "cash_flow",
        sourceTableType: "cash_flow_reconciliation",
        section: "unknown",
        periodType: "duration"
      }),
      row: "D&A",
      expected: false
    },
    {
      name: "current maturities of long-term debt cannot map to Revolver",
      request: request({
        label: "Current maturities of long-term debt",
        xbrlTag: "LongTermDebtCurrent",
        section: "current liabilities"
      }),
      row: "Revolver",
      expected: false
    },
    {
      name: "current maturities of long-term debt map to debt including current portion",
      request: request({
        label: "Current maturities of long-term debt",
        xbrlTag: "LongTermDebtCurrent",
        section: "current liabilities"
      }),
      row: "LT Debt (Incl. Current Portion)",
      expected: true
    },
    {
      name: "generic current debt may map to debt including current portion when it is not identified as short-term borrowing",
      request: request({
        label: "Debt, Current",
        xbrlTag: "DebtCurrent",
        section: "current liabilities"
      }),
      row: "LT Debt (Incl. Current Portion)",
      expected: true
    },
    {
      name: "commercial paper maps to Revolver/current borrowings",
      request: request({
        label: "Commercial paper",
        xbrlTag: "CommercialPaper",
        section: "current liabilities"
      }),
      row: "Revolver",
      expected: true
    },
    {
      name: "inventory-like assets cannot be hidden in the current-assets catch-all",
      request: request({
        label: "Spare parts and supplies inventory",
        xbrlTag: "InventoryPartsAndSupplies",
        section: "current assets"
      }),
      row: "Prepaid & Other Current Assets",
      expected: false
    },
    {
      name: "inventory-like assets map to Inventory",
      request: request({
        label: "Spare parts and supplies inventory",
        xbrlTag: "InventoryPartsAndSupplies",
        section: "current assets"
      }),
      row: "Inventory",
      expected: true
    },
    {
      name: "non-current inventory cannot map to the current Inventory row",
      request: request({
        label: "Inventories classified in Other assets",
        xbrlTag: "InventoryNoncurrent",
        section: "non-current assets"
      }),
      row: "Inventory",
      expected: false
    },
    {
      name: "non-current inventory maps to Other Non-Current Assets",
      request: request({
        label: "Inventories classified in Other assets",
        xbrlTag: "InventoryNoncurrent",
        section: "non-current assets"
      }),
      row: "Other Non-Current Assets",
      expected: true
    },
    {
      name: "equity securities without a readily determinable fair value remain investment assets",
      request: request({
        label: "Equity Securities without Readily Determinable Fair Value, Amount",
        xbrlTag: "EquitySecuritiesWithoutReadilyDeterminableFairValueAmount",
        section: "unknown"
      }),
      row: "Other Non-Current Assets",
      expected: true
    },
    {
      name: "equity securities cannot map to common stock and APIC merely because the label contains equity",
      request: request({
        label: "Equity Securities without Readily Determinable Fair Value, Amount",
        xbrlTag: "EquitySecuritiesWithoutReadilyDeterminableFairValueAmount",
        section: "unknown"
      }),
      row: "Common Stock & APIC",
      expected: false
    },
    {
      name: "utility direct costs can map to COGS without a ticker-specific rule",
      request: request({
        label: "Fuel and purchased power expense",
        xbrlTag: "UtilityFuelAndPurchasedPower",
        statement: "income_statement",
        section: "operating expenses",
        periodType: "duration"
      }),
      row: "COGS / Cost of Goods Sold",
      expected: true
    },
    {
      name: "cost of revenue excluding D&A maps to COGS",
      request: request({
        label: "Cost of revenue, excluding depreciation and amortization",
        xbrlTag: "CostOfRevenueExcludingDepreciationDepletionAndAmortization",
        statement: "income_statement",
        section: "operating expenses",
        periodType: "duration"
      }),
      row: "COGS / Cost of Goods Sold",
      expected: true
    },
    {
      name: "cost of revenue excluding D&A cannot map to D&A",
      request: request({
        label: "Cost of revenue, exclusive of depreciation",
        xbrlTag: "CostOfRevenueExcludingDepreciationDepletionAndAmortization",
        statement: "income_statement",
        section: "operating expenses",
        periodType: "duration"
      }),
      row: "D&A",
      expected: false
    },
    {
      name: "net interest expense maps to Interest Expense",
      request: request({
        label: "Interest expense, net of interest income",
        xbrlTag: "InterestExpenseNonOperatingNet",
        statement: "income_statement",
        section: "below operating income",
        periodType: "duration"
      }),
      row: "Interest Expense",
      expected: true
    },
    {
      name: "net interest expense cannot map to Interest Income",
      request: request({
        label: "Interest expense, net of interest income",
        xbrlTag: "InterestExpenseNonOperatingNet",
        statement: "income_statement",
        section: "below operating income",
        periodType: "duration"
      }),
      row: "Interest Income",
      expected: false
    },
    {
      name: "ambiguous interest income expense net belongs in other non-operating",
      request: request({
        label: "Interest income (expense), net",
        xbrlTag: "InterestIncomeExpenseNonOperatingNet",
        statement: "income_statement",
        section: "below operating income",
        periodType: "duration"
      }),
      row: "Other Non-Operating Income / Expense",
      expected: true
    },
    {
      name: "ambiguous interest income expense net cannot self-select Interest Income",
      request: request({
        label: "Interest income (expense), net",
        xbrlTag: "InterestIncomeExpenseNonOperatingNet",
        statement: "income_statement",
        section: "below operating income",
        periodType: "duration"
      }),
      row: "Interest Income",
      expected: false
    },
    {
      name: "advertising cannot be hidden in other operating expense when SG&A exists",
      request: request({
        label: "Advertising expense",
        xbrlTag: "AdvertisingExpense",
        statement: "income_statement",
        section: "operating expenses",
        periodType: "duration"
      }),
      row: "Other Operating Income / Expense",
      expected: false
    },
    {
      name: "advertising maps to SG&A",
      request: request({
        label: "Advertising expense",
        xbrlTag: "AdvertisingExpense",
        statement: "income_statement",
        section: "operating expenses",
        periodType: "duration"
      }),
      row: "SG&A",
      expected: true
    }
  ];
  for (const testCase of trustedValidationCases) {
    const actual = classificationPassesValidation(testCase.request, {
      recommended_action: "map",
      recommended_model_row: testCase.row,
      is_deferred_tax: !testCase.expected,
      is_debt: !testCase.expected,
      mapping_passed_validation: true,
      reason: "Untrusted LLM assertion"
    });
    assert.equal(actual, testCase.expected, testCase.name);
  }
  const combinedNonOperatingRequest = request({
    label: "Interest and other income (loss), net",
    xbrlTag: "NonoperatingIncomeExpense",
    statement: "income_statement",
    section: "below operating income",
    periodType: "duration",
    isSubtotal: false
  });
  assert.equal(
    classificationPassesValidation(combinedNonOperatingRequest, {
      recommended_action: "exclude",
      recommended_model_row: "Other Non-Operating Income / Expense",
      is_subtotal: true,
      confidence: "high",
      mapping_passed_validation: true,
      reason: "This combined subtotal would double-count components that are reported and mapped separately."
    }),
    true,
    "a high-confidence LLM subtotal exclusion must survive when the parser missed the subtotal"
  );
  assert.equal(
    classificationPassesValidation(combinedNonOperatingRequest, {
      recommended_action: "exclude",
      recommended_model_row: "Other Non-Operating Income / Expense",
      is_subtotal: false,
      confidence: "high",
      mapping_passed_validation: true,
      reason: "Exclude without accounting evidence."
    }),
    false,
    "an unsupported exclusion must still fail closed"
  );

  const rejectedWholeStatementRevenueMapping = await classifyFinancialStatementLineItems(
    [
      request({
        sourceRowKey: "row-trusted-gate-rd",
        rowOrder: 1,
        label: "Research and development expense",
        xbrlTag: "ResearchAndDevelopmentExpense",
        statement: "income_statement",
        section: "operating expenses",
        periodType: "duration"
      })
    ],
    {
      llm: {
        enabled: true,
        apiKey: "test-key",
        endpoint: "https://example.test/chat/completions",
        model: "openai/gpt-5.2",
        siteUrl: "http://localhost:3000",
        appTitle: "Historicals Solver Test",
        timeoutMs: 100,
        fetchImpl: async () =>
          new Response(
            JSON.stringify({
              choices: [
                {
                  message: {
                    content: JSON.stringify({
                      classifications: [
                        {
                          source_row_key: "row-trusted-gate-rd",
                          recommended_action: "remap",
                          recommended_model_row: "Revenue",
                          confidence: "high",
                          reason: "Malicious semantic remap used to exercise the trusted gate."
                        }
                      ]
                    })
                  }
                }
              ]
            }),
            { status: 200 }
          )
      },
      statementAnalystPass: { enabled: true, coverage: "all_primary_rows" }
    }
  );
  const rejectedWholeStatementClassification = rejectedWholeStatementRevenueMapping.classifications[0].classification;
  assert.equal(rejectedWholeStatementClassification.recommended_model_row, "R&D");
  assert.match(rejectedWholeStatementClassification.reason, /rejected by accounting validation/i);
  assert.equal(rejectedWholeStatementClassification.llm_used, false);
  assert.equal(rejectedWholeStatementRevenueMapping.llmReviewedCount, 0);
  assert.equal(rejectedWholeStatementRevenueMapping.acceptedDecisionCount, 0);
  assert.deepEqual(rejectedWholeStatementRevenueMapping.unreviewedTargetKeys, ["row-trusted-gate-rd"]);
  assert.equal(rejectedWholeStatementRevenueMapping.llmTelemetry[0].affectedOutput, false);

  const incomeStatementCostSubtotal = {
    label: "Cost of sales, operating expenses, and other-net",
    concept: "CostOfSalesOperatingExpensesAndOtherNet"
  };
  const incomeRows = [
    "COGS / Cost of Goods Sold",
    "Selling, General & Administration (SG&A)",
    "Other Operating Income (Expense)"
  ];
  const subtotalChildCostOfSales = await classify({
    statement: "income_statement",
    periodType: "duration",
    label: "Cost of sales",
    xbrlTag: "CostOfGoodsAndServicesSold",
    section: "operating expenses",
    parentSubtotal: incomeStatementCostSubtotal,
    availableModelRows: incomeRows
  });
  assert.equal(subtotalChildCostOfSales.recommended_model_row, "COGS / Cost of Goods Sold");

  const marketingAdmin = await classify({
    statement: "income_statement",
    periodType: "duration",
    label: "Marketing, selling, and administrative",
    xbrlTag: "SellingGeneralAndAdministrativeExpense",
    section: "operating expenses",
    parentSubtotal: incomeStatementCostSubtotal,
    deterministicCandidate: "SG&A",
    availableModelRows: incomeRows
  });
  assert.equal(modelRowsMatch(marketingAdmin.recommended_model_row, "Selling, General & Administration (SG&A)"), true);

  const restructuringCharges = await classify({
    statement: "income_statement",
    periodType: "duration",
    label: "Asset impairment, restructuring, and other special charges",
    xbrlTag: "RestructuringSettlementAndImpairmentProvisions",
    section: "operating expenses",
    parentSubtotal: incomeStatementCostSubtotal,
    deterministicCandidate: "Other Operating Income / Expense",
    availableModelRows: incomeRows
  });
  assert.equal(modelRowsMatch(restructuringCharges.recommended_model_row, "Other Operating Income (Expense)"), true);

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

  const debtIssuanceCostAdjustment = await classify({
    label: "Less unamortized debt discounts and issuance costs",
    xbrlTag: "DeferredFinanceCostsNet",
    section: "unknown",
    deterministicCandidate: "Unmapped / Needs Review"
  });
  assert.equal(debtIssuanceCostAdjustment.recommended_model_row, "LT Debt (Incl. Current Portion)");
  assert.equal(debtIssuanceCostAdjustment.is_current, false);
  assert.equal(debtIssuanceCostAdjustment.is_debt, true);
  assert.equal(debtIssuanceCostAdjustment.mapping_passed_validation, true);

  const currentSelfInsuranceReserve = await classify({
    label: "Less: current portion",
    xbrlTag: "SelfInsuranceReserveCurrent",
    section: "current liabilities",
    deterministicCandidate: "LT Debt (Incl. Current Portion)"
  });
  assert.equal(currentSelfInsuranceReserve.recommended_model_row, "Accrued Liabilities");
  assert.equal(currentSelfInsuranceReserve.is_debt, false);
  assert.equal(currentSelfInsuranceReserve.mapping_passed_validation, true);

  const noncurrentSelfInsuranceReserve = await classify({
    label: "Insurance reserves, net of current portion",
    xbrlTag: "SelfInsuranceReserveNoncurrent",
    section: "current liabilities",
    deterministicCandidate: "LT Debt (Incl. Current Portion)"
  });
  assert.equal(noncurrentSelfInsuranceReserve.recommended_model_row, "Other Non-Current Liabilities");
  assert.equal(noncurrentSelfInsuranceReserve.is_current, false);
  assert.equal(noncurrentSelfInsuranceReserve.is_debt, false);
  assert.equal(noncurrentSelfInsuranceReserve.mapping_passed_validation, true);

  const noncurrentRestrictedCash = await classify({
    label: "Restricted cash and marketable securities",
    xbrlTag: "RestrictedCashAndInvestmentsNoncurrent",
    section: "unknown",
    deterministicCandidate: "Cash & Cash Equivalents"
  });
  assert.equal(noncurrentRestrictedCash.recommended_model_row, "Other Non-Current Assets");
  assert.equal(noncurrentRestrictedCash.is_current, false);
  assert.equal(noncurrentRestrictedCash.mapping_passed_validation, true);

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
  assert.equal(shortTermInvestments.recommended_model_row, "Cash & Cash Equivalents");
  assert.equal(
    classificationModelRowAssignmentForPrimaryStatement(shortTermInvestments, availableModelRows).modelRow,
    "Cash & Cash Equivalents"
  );

  const marketableSecurities = await classify({
    label: "Marketable securities",
    section: "current assets",
    deterministicCandidate: "Prepaid & Other Current Assets",
    availableModelRows: availableModelRows.filter((row) => !/short[-\s]?term investments?|current investments?|marketable securities/i.test(row))
  });
  assert.equal(marketableSecurities.recommended_model_row, "Cash & Cash Equivalents");

  const dedicatedShortTermInvestments = await classify({
    label: "Short-term investments",
    section: "current assets",
    deterministicCandidate: "Prepaid & Other Current Assets",
    availableModelRows: [...availableModelRows, "Short-Term Investments"]
  });
  assert.equal(dedicatedShortTermInvestments.recommended_model_row, "Short-Term Investments");

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

  const interestPayable = await classify({
    label: "Interest Payable, Current",
    xbrlTag: "InterestPayableCurrent",
    statement: "balance_sheet",
    section: "current liabilities",
    periodType: "instant",
    deterministicCandidate: "Other Current Liabilities",
    uncertaintyReason: "Current interest payable is a non-debt accrued liability line."
  });
  assert.equal(interestPayable.recommended_model_row, "Accrued Liabilities");
  assert.equal(interestPayable.mapping_passed_validation, true);

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

  const redeemableNciWithMezzanineRow = await classify({
    label: "Redeemable noncontrolling interests in subsidiaries",
    xbrlTag: "RedeemableNoncontrollingInterestEquityCarryingAmount",
    statement: "balance_sheet",
    section: "equity",
    periodType: "instant",
    deterministicCandidate: "Common Stock & APIC",
    uncertaintyReason: "Redeemable noncontrolling interests are mezzanine equity."
  });
  assert.equal(redeemableNciWithMezzanineRow.recommended_model_row, "Mezzanine Equity");
  assert.equal(redeemableNciWithMezzanineRow.mapping_passed_validation, true);

  const availableRowsWithoutMezzanine = availableModelRows.filter((row) => row !== "Mezzanine Equity");
  const redeemableNciWithoutMezzanineRow = await classify({
    label: "Redeemable noncontrolling interests in subsidiaries",
    xbrlTag: "RedeemableNoncontrollingInterestEquityCarryingAmount",
    statement: "balance_sheet",
    section: "equity",
    periodType: "instant",
    availableModelRows: availableRowsWithoutMezzanine,
    deterministicCandidate: "Common Stock & APIC",
    uncertaintyReason: "Redeemable noncontrolling interests are mezzanine equity, but the template has no mezzanine row."
  });
  assert.equal(redeemableNciWithoutMezzanineRow.recommended_model_row, "Other Non-Current Liabilities");
  assert.equal(redeemableNciWithoutMezzanineRow.mapping_passed_validation, true);
  assert.equal(
    classificationModelRowAssignmentForPrimaryStatement(redeemableNciWithoutMezzanineRow, availableRowsWithoutMezzanine).modelRow,
    "Other Non-Current Liabilities"
  );
  assert.equal(
    lineItemNeedsClassification(
      request({
        label: "Redeemable noncontrolling interests in subsidiaries",
        xbrlTag: "RedeemableNoncontrollingInterestEquityCarryingAmount",
        statement: "balance_sheet",
        section: "equity",
        periodType: "instant"
      })
    ),
    true
  );

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
  assert.equal(
    classificationPassesValidation(
      request({
        label: "Adjustment to withdrawal liability for multiemployer pension funds",
        xbrlTag: "PensionAndOtherPostretirementBenefitExpense",
        statement: "income_statement",
        section: "operating expenses",
        periodType: "duration"
      }),
      {
        recommended_model_row: "Other Non-Operating Income / Expense",
        is_deferred_tax: false,
        is_debt: false
      }
    ),
    false,
    "An LLM may not move a primary-statement operating expense below operating income."
  );

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

  const companyOperatedRestaurantExpenses = await classify({
    label: "Company-owned and operated restaurant expenses",
    xbrlTag: "CompanyOperatedRestaurantExpenses",
    statement: "income_statement",
    section: "operating expenses",
    periodType: "duration",
    uncertaintyReason: ""
  });
  assert.equal(companyOperatedRestaurantExpenses.recommended_model_row, "COGS / Cost of Goods Sold");
  assert.equal(companyOperatedRestaurantExpenses.mapping_passed_validation, true);

  const restaurantFoodAndPaper = await classify({
    label: "Food & paper",
    xbrlTag: "FoodAndPaperExpense",
    statement: "income_statement",
    section: "operating expenses",
    periodType: "duration",
    parentSubtotal: { label: "Company-owned and operated restaurant expenses" },
    uncertaintyReason: ""
  });
  assert.equal(restaurantFoodAndPaper.recommended_model_row, "COGS / Cost of Goods Sold");
  assert.equal(restaurantFoodAndPaper.mapping_passed_validation, true);

  const costExcludingDa = await classify({
    label: "Cost of revenue, excluding depreciation and amortization",
    xbrlTag: "CostOfRevenueExcludingDepreciationDepletionAndAmortization",
    statement: "income_statement",
    section: "operating expenses",
    periodType: "duration",
    uncertaintyReason: ""
  });
  assert.equal(costExcludingDa.recommended_model_row, "COGS / Cost of Goods Sold");

  const netInterestExpense = await classify({
    label: "Interest expense, net of interest income",
    xbrlTag: "InterestExpenseNonOperatingNet",
    statement: "income_statement",
    section: "below operating income",
    periodType: "duration",
    uncertaintyReason: ""
  });
  assert.equal(netInterestExpense.recommended_model_row, "Interest Expense");

  const ambiguousNetInterest = await classify({
    label: "Interest income (expense), net",
    xbrlTag: "InterestIncomeExpenseNonOperatingNet",
    statement: "income_statement",
    section: "below operating income",
    periodType: "duration",
    uncertaintyReason: ""
  });
  assert.equal(ambiguousNetInterest.recommended_model_row, "Other Non-Operating Income / Expense");

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
        model: "openai/gpt-5.2",
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
  assert.equal("temperature" in llmRequestedPayloads[0], false);
  assert.deepEqual(llmRequestedPayloads[0].reasoning, { effort: "low", exclude: true });
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
        model: "openai/gpt-5.2",
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
  assert.equal(
    statementPayload.statementRows.some(
      (row) =>
        row.sourceRowKey === "row-investments" &&
        row.target === true &&
        row.deterministicClassification?.recommendedModelRow === "Cash & Cash Equivalents"
    ),
    true
  );
  assert.equal(
    statementPayload.statementRows.some(
      (row) => row.sourceRowKey === "row-advertising" && row.target === false && row.deterministicClassification?.recommendedModelRow === "SG&A"
    ),
    true
  );
  assert.equal(statementPayload.modelRowDefinitions["SG&A"].includes("advertising"), true);
  assert.equal(statementBatchPayloads[0].response_format.json_schema.schema.properties.classifications.items.properties.source_row_key.type, "string");
  assert.equal("temperature" in statementBatchPayloads[0], false);
  assert.deepEqual(statementBatchPayloads[0].reasoning, { effort: "low", exclude: true });
  assert.equal(batchResult.llmCalls, 1);
  assert.equal(batchResult.llmAttempts, 1);
  assert.equal(batchResult.llmSuccessfulCompletions, 1);
  assert.equal(batchResult.llmTelemetry[0].status, "completed_validated");
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

  const fullAnalystPassPayloads = [];
  const fullAnalystPassResult = await classifyFinancialStatementLineItems(
    [
      request({
        sourceRowKey: "row-cash",
        rowOrder: 1,
        label: "Cash and cash equivalents",
        xbrlTag: "CashAndCashEquivalentsAtCarryingValue",
        section: "current assets",
        amount: 1_000_000,
        uncertaintyReason: ""
      }),
      request({
        sourceRowKey: "row-investments",
        rowOrder: 2,
        label: "Short-term investments",
        xbrlTag: "ShortTermInvestments",
        section: "current assets",
        amount: 2_000_000,
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
        amount: 3_000_000,
        uncertaintyReason: ""
      }),
      request({
        sourceRowKey: "row-lease-financing",
        rowOrder: 12,
        label: "Other lease financing obligations",
        xbrlTag: "OtherLeaseFinancingObligations",
        section: "non-current liabilities",
        amount: 4_000_000,
        deterministicCandidate: "Other Non-Current Liabilities",
        uncertaintyReason: "lease financing may represent debt or another long-term liability"
      }),
      request({
        sourceRowKey: "row-notes-payable",
        rowOrder: 13,
        label: "Notes payable",
        xbrlTag: "NotesPayableCurrent",
        section: "current liabilities",
        amount: 5_000_000,
        deterministicCandidate: "Other Current Liabilities",
        uncertaintyReason: "notes payable could be current borrowings even if reported below other liabilities"
      })
    ],
    {
      llm: {
        enabled: true,
        apiKey: "test-key",
        endpoint: "https://example.test/chat/completions",
        model: "openai/gpt-5.2",
        siteUrl: "http://localhost:3000",
        appTitle: "Historicals Solver Test",
        timeoutMs: 100,
        fetchImpl: async (_url, init) => {
          const body = JSON.parse(init.body);
          fullAnalystPassPayloads.push(body);
          return {
            ok: true,
            json: async () => ({
              choices: [
                {
                  message: {
                    content: JSON.stringify({
                      classifications: [
                        {
                          source_row_key: "row-cash",
                          source_line_item: "Cash and cash equivalents",
                          recommended_action: "map",
                          recommended_model_row: "Cash & Cash Equivalents",
                          recommended_model_row_mappings: [],
                          explicit_zero_rows: [],
                          classification_type: "cash",
                          is_current: true,
                          is_debt: false,
                          is_operating: false,
                          is_tax_related: false,
                          is_deferred_revenue_or_contract_liability: false,
                          is_deferred_tax: false,
                          is_subtotal: false,
                          should_exclude_from_other_bucket: true,
                          confidence: "high",
                          reason: "Cash maps to the cash row after reviewing the full statement.",
                          requires_validation: true,
                          requires_revalidation: true
                        },
                        {
                          source_row_key: "row-investments",
                          source_line_item: "Short-term investments",
                          recommended_action: "remap",
                          recommended_model_row: "Cash & Cash Equivalents",
                          recommended_model_row_mappings: [],
                          explicit_zero_rows: [],
                          classification_type: "cash-like current investment",
                          is_current: true,
                          is_debt: false,
                          is_operating: false,
                          is_tax_related: false,
                          is_deferred_revenue_or_contract_liability: false,
                          is_deferred_tax: false,
                          is_subtotal: false,
                          should_exclude_from_other_bucket: true,
                          confidence: "high",
                          reason: "No dedicated current investments row is available, so investments group with cash.",
                          requires_validation: true,
                          requires_revalidation: true
                        },
                        {
                          source_row_key: "row-advertising",
                          source_line_item: "Advertising expense",
                          recommended_action: "map",
                          recommended_model_row: "SG&A",
                          recommended_model_row_mappings: [],
                          explicit_zero_rows: [],
                          classification_type: "selling and marketing expense",
                          is_current: null,
                          is_debt: false,
                          is_operating: true,
                          is_tax_related: false,
                          is_deferred_revenue_or_contract_liability: false,
                          is_deferred_tax: false,
                          is_subtotal: false,
                          should_exclude_from_other_bucket: true,
                          confidence: "high",
                          reason: "Advertising is a selling and marketing operating expense and groups into SG&A.",
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
      },
      statementAnalystPass: {
        enabled: true,
        materialityThreshold: 500_000
      }
    }
  );
  assert.equal(fullAnalystPassPayloads.length, 1);
  const fullAnalystPayload = JSON.parse(fullAnalystPassPayloads[0].messages[1].content);
  assert.deepEqual(fullAnalystPayload.targetSourceRowKeys, [
    "row-cash",
    "row-investments",
    "row-advertising",
    "row-lease-financing",
    "row-notes-payable"
  ]);
  assert.equal(fullAnalystPayload.statementRows.every((row) => row.target === true), true);
  assert.equal(
    fullAnalystPayload.statementRows.every((row) => row.targetReason === "material_pre_fill_analyst_pass"),
    true
  );
  assert.equal(
    fullAnalystPayload.classificationGoal.includes("Deterministic candidates are evidence and validation guardrails"),
    true
  );
  assert.equal(fullAnalystPassResult.classifications.length, 5);
  assert.equal(fullAnalystPassResult.classifications.find((item) => item.request.sourceRowKey === "row-cash").classification.llm_used, true);
  assert.equal(
    fullAnalystPassResult.classifications.find((item) => item.request.sourceRowKey === "row-advertising").classification.recommended_model_row,
    "SG&A"
  );

  const completeCoveragePayloads = [];
  const completeCoverageResult = await classifyFinancialStatementLineItems(
    [
      request({
        sourceRowKey: "row-small-cash",
        rowOrder: 1,
        label: "Cash and cash equivalents",
        xbrlTag: "CashAndCashEquivalentsAtCarryingValue",
        section: "current assets",
        amount: 100,
        currentPeriodSourceLines: ["Cash and cash equivalents", "Accounts receivable, net", "Total assets"],
        uncertaintyReason: ""
      }),
      request({
        sourceRowKey: "row-small-receivables",
        rowOrder: 2,
        label: "Accounts receivable, net",
        xbrlTag: "AccountsReceivableNetCurrent",
        section: "current assets",
        amount: 200,
        currentPeriodSourceLines: ["Cash and cash equivalents", "Accounts receivable, net", "Total assets"],
        uncertaintyReason: ""
      })
    ],
    {
      llm: {
        enabled: true,
        apiKey: "test-key",
        endpoint: "https://example.test/chat/completions",
        model: "openai/gpt-5.6-terra-pro",
        siteUrl: "http://localhost:3000",
        appTitle: "Historicals Solver Test",
        timeoutMs: 100,
        fetchImpl: async (_url, init) => {
          const body = JSON.parse(init.body);
          completeCoveragePayloads.push(body);
          return new Response(
            JSON.stringify({
              choices: [
                {
                  message: {
                    content: JSON.stringify({
                      classifications: [
                        {
                          source_row_key: "row-small-cash",
                          recommended_action: "map",
                          recommended_model_row: "Cash & Cash Equivalents",
                          confidence: "high",
                          reason: "The primary statement presents a cash carrying amount."
                        },
                        {
                          source_row_key: "row-small-receivables",
                          recommended_action: "map",
                          recommended_model_row: "Accounts Receivable",
                          confidence: "high",
                          reason: "The primary statement presents net accounts receivable."
                        }
                      ]
                    })
                  }
                }
              ]
            }),
            { status: 200 }
          );
        }
      },
      statementAnalystPass: {
        enabled: true,
        materialityThreshold: 500_000,
        coverage: "all_primary_rows"
      }
    }
  );
  assert.equal(completeCoveragePayloads.length, 1);
  const completeCoveragePayload = JSON.parse(completeCoveragePayloads[0].messages[1].content);
  assert.equal(completeCoveragePayload.statementRows.every((row) => row.target), true);
  assert.equal(
    completeCoveragePayload.statementRows.every((row) => row.targetReason === "complete_primary_statement_coverage"),
    true
  );
  assert.equal(completeCoveragePayload.statementContexts.length, 1);
  assert.deepEqual(completeCoveragePayload.statementContexts[0].orderedSourceLines, [
    "Cash and cash equivalents",
    "Accounts receivable, net",
    "Total assets"
  ]);
  assert.deepEqual(
    Object.keys(completeCoveragePayloads[0].response_format.json_schema.schema.properties.classifications.items.properties).sort(),
    ["confidence", "reason", "recommended_action", "recommended_model_row", "source_row_key"]
  );
  assert.equal(completeCoverageResult.targetCount, 2);
  assert.equal(completeCoverageResult.llmReviewedCount, 2);
  assert.equal(completeCoverageResult.acceptedDecisionCount, 2);
  assert.deepEqual(completeCoverageResult.unreviewedTargetKeys, []);
  assert.equal(completeCoverageResult.classifications.every((item) => item.classification.llm_used), true);

  const tolerantCoverageResult = await classifyFinancialStatementLineItems(
    [
      request({
        sourceRowKey: "row-tolerant-advertising",
        rowOrder: 1,
        statement: "income_statement",
        periodType: "duration",
        label: "Advertising expense",
        xbrlTag: "AdvertisingExpense",
        section: "operating expenses",
        amount: 200,
        uncertaintyReason: ""
      }),
      request({
        sourceRowKey: "row-tolerant-debt",
        rowOrder: 2,
        label: "Current maturities of long-term debt",
        xbrlTag: "LongTermDebtCurrent",
        section: "current liabilities",
        amount: 100,
        uncertaintyReason: ""
      })
    ],
    {
      llm: {
        enabled: true,
        apiKey: "test-key",
        endpoint: "https://example.test/chat/completions",
        model: "openai/gpt-5.6-terra-pro",
        siteUrl: "http://localhost:3000",
        appTitle: "Historicals Solver Test",
        timeoutMs: 100,
        fetchImpl: async () =>
          new Response(
            JSON.stringify({
              choices: [
                {
                  message: {
                    content:
                      '```json\n[{"source_row_key":"row-tolerant-debt","action":"assign","recommendedModelRow":"PP&E, Net","confidence":"high","reason":"Incorrect provider recommendation used to exercise validation."},{"source_row_key":"row-tolerant-advertising","recommendedAction":"Map to model row: SG&A","modelRow":"SG&A","certainty":"certain","rationale":"Advertising is an SG&A operating expense."}]\n```'
                  }
                }
              ]
            }),
            { status: 200 }
          )
      },
      statementAnalystPass: { enabled: true, coverage: "all_primary_rows" }
    }
  );
  assert.equal(tolerantCoverageResult.llmReviewedCount, 1);
  assert.equal(tolerantCoverageResult.acceptedDecisionCount, 1);
  assert.deepEqual(tolerantCoverageResult.unreviewedTargetKeys, ["row-tolerant-debt"]);
  assert.equal(
    tolerantCoverageResult.classifications.find((item) => item.request.sourceRowKey === "row-tolerant-advertising").classification
      .recommended_model_row,
    "SG&A"
  );
  assert.equal(
    tolerantCoverageResult.classifications.find((item) => item.request.sourceRowKey === "row-tolerant-advertising").classification.llm_used,
    true
  );
  const guardedDebt = tolerantCoverageResult.classifications.find(
    (item) => item.request.sourceRowKey === "row-tolerant-debt"
  ).classification;
  assert.equal(guardedDebt.recommended_model_row, "LT Debt (Incl. Current Portion)");
  assert.equal(guardedDebt.reason.includes("rejected by accounting validation"), true);
  assert.equal(guardedDebt.llm_used, false);

  const lowConfidenceCoverageResult = await classifyFinancialStatementLineItems(
    [
      request({
        sourceRowKey: "row-low-confidence-advertising",
        rowOrder: 1,
        statement: "income_statement",
        periodType: "duration",
        label: "Advertising expense",
        xbrlTag: "AdvertisingExpense",
        section: "operating expenses",
        amount: 200,
        uncertaintyReason: ""
      })
    ],
    {
      llm: {
        enabled: true,
        apiKey: "test-key",
        endpoint: "https://example.test/chat/completions",
        model: "openai/gpt-5.6-terra-pro",
        siteUrl: "http://localhost:3000",
        appTitle: "Historicals Solver Test",
        timeoutMs: 100,
        fetchImpl: async () =>
          new Response(
            JSON.stringify({
              choices: [
                {
                  message: {
                    content: JSON.stringify({
                      classifications: [
                        {
                          source_row_key: "row-low-confidence-advertising",
                          recommended_action: "map",
                          recommended_model_row: "SG&A",
                          confidence: "low",
                          reason: "The provider was not confident enough to certify this decision."
                        }
                      ]
                    })
                  }
                }
              ]
            }),
            { status: 200 }
          )
      },
      statementAnalystPass: { enabled: true, coverage: "all_primary_rows" }
    }
  );
  const lowConfidenceFallback = lowConfidenceCoverageResult.classifications[0].classification;
  assert.equal(lowConfidenceFallback.recommended_model_row, "SG&A");
  assert.match(lowConfidenceFallback.reason, /LLM returned low confidence/i);
  assert.equal(lowConfidenceFallback.llm_used, false);
  assert.equal(lowConfidenceCoverageResult.llmReviewedCount, 0);
  assert.equal(lowConfidenceCoverageResult.acceptedDecisionCount, 0);
  assert.deepEqual(lowConfidenceCoverageResult.unreviewedTargetKeys, ["row-low-confidence-advertising"]);
  assert.equal(lowConfidenceCoverageResult.llmTelemetry[0].affectedOutput, false);

  let missingSourceKeyAttempts = 0;
  const missingSourceKeyResult = await classifyFinancialStatementLineItems(
    [
      request({
        sourceRowKey: "row-missing-key-advertising",
        rowOrder: 1,
        statement: "income_statement",
        periodType: "duration",
        label: "Advertising expense",
        xbrlTag: "AdvertisingExpense",
        section: "operating expenses",
        uncertaintyReason: ""
      }),
      request({
        sourceRowKey: "row-missing-key-debt",
        rowOrder: 2,
        label: "Current maturities of long-term debt",
        xbrlTag: "LongTermDebtCurrent",
        section: "current liabilities",
        uncertaintyReason: ""
      })
    ],
    {
      llm: {
        enabled: true,
        apiKey: "test-key",
        endpoint: "https://example.test/chat/completions",
        model: "openai/gpt-5.6-terra-pro",
        siteUrl: "http://localhost:3000",
        appTitle: "Historicals Solver Test",
        timeoutMs: 100,
        maxAttempts: 1,
        fetchImpl: async () => {
          missingSourceKeyAttempts += 1;
          return new Response(
            JSON.stringify({
              choices: [
                {
                  message: {
                    content: JSON.stringify({
                      classifications: [
                        {
                          recommended_action: "map",
                          recommended_model_row: "SG&A",
                          confidence: "high",
                          reason: "Advertising is an SG&A operating expense."
                        },
                        {
                          recommended_action: "map",
                          recommended_model_row: "LT Debt (Incl. Current Portion)",
                          confidence: "high",
                          reason: "Current maturities are part of long-term debt."
                        }
                      ]
                    })
                  }
                }
              ]
            }),
            { status: 200 }
          );
        }
      },
      statementAnalystPass: { enabled: true, coverage: "all_primary_rows" }
    }
  );
  assert.equal(missingSourceKeyAttempts, 1);
  assert.equal(missingSourceKeyResult.llmAttempts, 1);
  assert.equal(missingSourceKeyResult.acceptedDecisionCount, 0);
  assert.equal(missingSourceKeyResult.classifications.every((item) => item.classification.llm_used === false), true);
  assert.match(missingSourceKeyResult.warnings.join("\n"), /must include source_row_key/i);

  const structurallyValidLlmResult = await classifyFinancialStatementLineItems(
    [
      request({
        sourceRowKey: "row-restricted-cash",
        label: "Restricted cash and marketable securities",
        xbrlTag: "RestrictedCashAndInvestmentsNoncurrent",
        section: "unknown",
        deterministicCandidate: undefined,
        uncertaintyReason: ""
      })
    ],
    {
      llm: {
        enabled: true,
        apiKey: "test-key",
        endpoint: "https://example.test/chat/completions",
        model: "openai/gpt-5.6-terra-pro",
        siteUrl: "http://localhost:3000",
        appTitle: "Historicals Solver Test",
        timeoutMs: 100,
        fetchImpl: async () =>
          new Response(
            JSON.stringify({
              choices: [
                {
                  message: {
                    content: JSON.stringify({
                      classifications: [
                        {
                          source_row_key: "row-restricted-cash",
                          recommended_action: "map",
                          recommended_model_row: "Other Non-Current Assets",
                          confidence: "high",
                          reason: "The SEC concept is explicitly non-current and the model has no dedicated restricted-cash row."
                        }
                      ]
                    })
                  }
                }
              ]
            }),
            { status: 200 }
          )
      },
      statementAnalystPass: { enabled: true, coverage: "all_primary_rows" }
    }
  );
  assert.equal(structurallyValidLlmResult.acceptedDecisionCount, 1);
  assert.equal(structurallyValidLlmResult.classifications[0].classification.mapping_passed_validation, true);
  assert.equal(structurallyValidLlmResult.classifications[0].classification.recommended_model_row, "Other Non-Current Assets");

  const noncurrentLiabilityLlmResult = await classifyFinancialStatementLineItems(
    [
      request({
        sourceRowKey: "row-environmental-noncurrent",
        label: "Accrued Capping, Closure, Post-closure and Environmental Costs, Noncurrent",
        xbrlTag: "AccruedCappingClosurePostClosureAndEnvironmentalCostsNoncurrent",
        section: "current liabilities",
        deterministicCandidate: "Other Current Liabilities"
      })
    ],
    {
      llm: {
        enabled: true,
        apiKey: "test-key",
        endpoint: "https://example.test/chat/completions",
        model: "openai/gpt-5.6-terra-pro",
        siteUrl: "http://localhost:3000",
        appTitle: "Historicals Solver Test",
        timeoutMs: 100,
        fetchImpl: async () =>
          new Response(
            JSON.stringify({
              choices: [
                {
                  message: {
                    content: JSON.stringify({
                      classifications: [
                        {
                          source_row_key: "row-environmental-noncurrent",
                          recommended_action: "remap",
                          recommended_model_row: "Other Non-Current Liabilities",
                          confidence: "high",
                          reason: "The XBRL concept explicitly identifies this environmental obligation as non-current."
                        }
                      ]
                    })
                  }
                }
              ]
            }),
            { status: 200 }
          )
      },
      statementAnalystPass: { enabled: true, coverage: "all_primary_rows" }
    }
  );
  assert.equal(noncurrentLiabilityLlmResult.acceptedDecisionCount, 1);
  assert.equal(noncurrentLiabilityLlmResult.classifications[0].classification.mapping_passed_validation, true);

  let splitBatchAttempts = 0;
  const splitBatchRequests = Array.from({ length: 6 }, (_, index) =>
    request({
      sourceRowKey: `row-split-${index + 1}`,
      rowOrder: index + 1,
      statement: "income_statement",
      periodType: "duration",
      label: `Advertising expense ${index + 1}`,
      xbrlTag: `AdvertisingExpense${index + 1}`,
      section: "operating expenses"
    })
  );
  const splitBatchResult = await classifyFinancialStatementLineItems(splitBatchRequests, {
    llm: {
      enabled: true,
      apiKey: "test-key",
      endpoint: "https://example.test/chat/completions",
      model: "openai/gpt-5.6-terra-pro",
      siteUrl: "http://localhost:3000",
      appTitle: "Historicals Solver Test",
      timeoutMs: 100,
      maxAttempts: 3,
      fetchImpl: async (_url, init) => {
        splitBatchAttempts += 1;
        const body = JSON.parse(init.body);
        const payload = JSON.parse(body.messages[1].content);
        if (splitBatchAttempts === 1) {
          return new Response(
            JSON.stringify({
              choices: [
                {
                  message: {
                    content: JSON.stringify({
                      classifications: payload.targetSourceRowKeys.map((sourceRowKey) => ({
                        source_row_key: sourceRowKey,
                        recommended_action: "map",
                        ...(sourceRowKey === "row-split-1" ? {} : { recommended_model_row: "SG&A" }),
                        confidence: "high",
                        reason: "Advertising is an SG&A operating expense."
                      }))
                    })
                  }
                }
              ]
            }),
            { status: 200 }
          );
        }
        return new Response(
          JSON.stringify({
            choices: [
              {
                message: {
                  content: JSON.stringify({
                    classifications: payload.targetSourceRowKeys.map((sourceRowKey) => ({
                      source_row_key: sourceRowKey,
                      recommended_action: "map",
                      recommended_model_row: "SG&A",
                      confidence: "high",
                      reason: "Advertising is an SG&A operating expense."
                    }))
                  })
                }
              }
            ]
          }),
          { status: 200 }
        );
      }
    },
    statementAnalystPass: { enabled: true, coverage: "all_primary_rows" }
  });
  assert.equal(splitBatchAttempts, 2, "one malformed row must receive a focused retry without discarding valid peer decisions");
  assert.equal(splitBatchResult.llmAttempts, 2);
  assert.equal(splitBatchResult.llmReviewedCount, 6);
  assert.equal(splitBatchResult.acceptedDecisionCount, 6);
  assert.deepEqual(splitBatchResult.unreviewedTargetKeys, []);
  assert.equal(
    splitBatchResult.classifications.find((item) => item.request.sourceRowKey === "row-split-1").classification.llm_used,
    true
  );
  assert.equal(
    splitBatchResult.classifications
      .filter((item) => item.request.sourceRowKey !== "row-split-1")
      .every((item) => item.classification.llm_used),
    true
  );

  let partialBatchAttempts = 0;
  const partialBatchResult = await classifyFinancialStatementLineItems(
    [
      request({
        sourceRowKey: "row-partial-unclassified",
        rowOrder: 1,
        statement: "income_statement",
        periodType: "duration",
        label: "Unclassified operating extension line",
        xbrlTag: "ExampleUnclassifiedOperatingExtension",
        section: "operating expenses",
        uncertaintyReason: "No independently validated deterministic classification exists."
      }),
      request({
        sourceRowKey: "row-partial-advertising",
        rowOrder: 2,
        statement: "income_statement",
        periodType: "duration",
        label: "Advertising expense",
        xbrlTag: "AdvertisingExpense",
        section: "operating expenses",
        uncertaintyReason: ""
      })
    ],
    {
      llm: {
        enabled: true,
        apiKey: "test-key",
        endpoint: "https://example.test/chat/completions",
        model: "openai/gpt-5.6-terra-pro",
        siteUrl: "http://localhost:3000",
        appTitle: "Historicals Solver Test",
        timeoutMs: 100,
        maxAttempts: 1,
        fetchImpl: async () => {
          partialBatchAttempts += 1;
          return new Response(
            JSON.stringify({
              choices: [
                {
                  message: {
                    content: JSON.stringify({
                      classifications: [
                        {
                          source_row_key: "row-partial-unclassified",
                          recommended_action: "map",
                          recommended_model_row: "Revenue",
                          confidence: "high",
                          reason: "Intentionally invalid cross-section recommendation."
                        },
                        {
                          source_row_key: "row-partial-advertising",
                          recommended_action: "map",
                          recommended_model_row: "SG&A",
                          confidence: "high",
                          reason: "Advertising is an SG&A operating expense."
                        }
                      ]
                    })
                  }
                }
              ]
            }),
            { status: 200 }
          );
        }
      },
      statementAnalystPass: { enabled: true, coverage: "all_primary_rows" }
    }
  );
  assert.equal(partialBatchAttempts, 1);
  assert.equal(partialBatchResult.llmAttempts, 1);
  assert.equal(partialBatchResult.acceptedDecisionCount, 1, "valid peer decisions must survive one rejected row in the same batch");
  assert.deepEqual(partialBatchResult.unreviewedTargetKeys, ["row-partial-unclassified"]);
  assert.equal(
    partialBatchResult.classifications.find((item) => item.request.sourceRowKey === "row-partial-advertising").classification.llm_used,
    true
  );
  assert.equal(
    partialBatchResult.classifications.find((item) => item.request.sourceRowKey === "row-partial-unclassified").classification.llm_used,
    false
  );

  let omittedRowRetryAttempts = 0;
  const omittedRowRetryTargets = [];
  const omittedRowRetryResult = await classifyFinancialStatementLineItems(
    [
      request({
        sourceRowKey: "row-retry-other-operating",
        rowOrder: 1,
        statement: "income_statement",
        periodType: "duration",
        label: "Unclassified operating extension line",
        xbrlTag: "ExampleUnclassifiedOperatingExtension",
        section: "operating expenses",
        uncertaintyReason: "No independently validated deterministic classification exists."
      }),
      request({
        sourceRowKey: "row-retry-advertising",
        rowOrder: 2,
        statement: "income_statement",
        periodType: "duration",
        label: "Advertising expense",
        xbrlTag: "AdvertisingExpense",
        section: "operating expenses",
        uncertaintyReason: ""
      })
    ],
    {
      llm: {
        enabled: true,
        apiKey: "test-key",
        endpoint: "https://example.test/chat/completions",
        model: "openai/gpt-5.6-terra-pro",
        siteUrl: "http://localhost:3000",
        appTitle: "Historicals Solver Test",
        timeoutMs: 100,
        maxAttempts: 2,
        fetchImpl: async (_url, init) => {
          omittedRowRetryAttempts += 1;
          const payload = JSON.parse(init.body).messages.at(-1);
          const requestPayload = JSON.parse(payload.content);
          omittedRowRetryTargets.push(requestPayload.targetSourceRowKeys);
          const sourceRowKey = omittedRowRetryAttempts === 1 ? "row-retry-advertising" : "row-retry-other-operating";
          return new Response(
            JSON.stringify({
              choices: [
                {
                  message: {
                    content: JSON.stringify({
                      classifications: [
                        {
                          source_row_key: sourceRowKey,
                          recommended_action: "map",
                          recommended_model_row: omittedRowRetryAttempts === 1 ? "SG&A" : "Other Operating Income / Expense",
                          confidence: "high",
                          reason: "Accounting-substance assignment from the whole-statement context."
                        }
                      ]
                    })
                  }
                }
              ]
            }),
            { status: 200 }
          );
        }
      },
      statementAnalystPass: { enabled: true, coverage: "all_primary_rows" }
    }
  );
  assert.equal(omittedRowRetryAttempts, 2, "an omitted difficult row must receive a focused follow-up analyst pass");
  assert.deepEqual(omittedRowRetryTargets[1], ["row-retry-other-operating"]);
  assert.equal(omittedRowRetryResult.acceptedDecisionCount, 2);
  assert.deepEqual(omittedRowRetryResult.unreviewedTargetKeys, []);

  let invalidLargeBatchAttempts = 0;
  const invalidLargeBatchRequests = Array.from({ length: 12 }, (_, index) =>
    request({
      sourceRowKey: `row-bounded-invalid-${index + 1}`,
      rowOrder: index + 1,
      statement: "income_statement",
      periodType: "duration",
      label: `Unclassified operating line bounded ${index + 1}`,
      xbrlTag: `ExampleUnclassifiedOperatingLine${index + 1}`,
      section: "operating expenses",
      uncertaintyReason: "No independently validated deterministic classification exists for this extension line."
    })
  );
  const invalidLargeBatchResult = await classifyFinancialStatementLineItems(invalidLargeBatchRequests, {
    llm: {
      enabled: true,
      apiKey: "test-key",
      endpoint: "https://example.test/chat/completions",
      model: "openai/gpt-5.6-terra-pro",
      fallbackModels: ["deepseek/deepseek-v4-flash", "anthropic/claude-sonnet-4"],
      siteUrl: "http://localhost:3000",
      appTitle: "Historicals Solver Test",
      timeoutMs: 1_000,
      maxAttempts: 4,
      fetchImpl: async (_url, init) => {
        invalidLargeBatchAttempts += 1;
        const body = JSON.parse(init.body);
        const payload = JSON.parse(body.messages[1].content);
        return new Response(
          JSON.stringify({
            choices: [
              {
                message: {
                  content: JSON.stringify({
                    classifications: payload.targetSourceRowKeys.map((sourceRowKey) => ({
                      source_row_key: sourceRowKey,
                      recommended_action: "map",
                      confidence: "high",
                      reason: "Intentionally invalid because the required model row is missing."
                    }))
                  })
                }
              }
            ]
          }),
          { status: 200 }
        );
      }
    },
    statementAnalystPass: { enabled: true, coverage: "all_primary_rows" }
  });
  assert.equal(invalidLargeBatchAttempts, 4, "recursive split retries must share one hard provider-attempt budget");
  assert.equal(invalidLargeBatchResult.llmAttempts, 4);
  assert.equal(invalidLargeBatchResult.acceptedDecisionCount, 0);
  assert.equal(invalidLargeBatchResult.unreviewedTargetKeys.length, 12);
  assert.equal(invalidLargeBatchResult.classifications.every((item) => item.classification.llm_used === false), true);

  assert.equal(
    fullStatementLineItemNeedsAnalystPass(
      request({
        label: "Total current assets",
        xbrlTag: "AssetsCurrent",
        section: "current assets",
        isSubtotal: true
      })
    ),
    false
  );
  assert.equal(
    fullStatementLineItemNeedsAnalystPass(
      request({
        statement: "income_statement",
        periodType: "duration",
        label: "Income (Loss) Before Taxes",
        xbrlTag: "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
        section: "below operating income"
      })
    ),
    false,
    "extended pre-tax subtotal concepts must not be paid LLM mapping targets"
  );
  assert.equal(
    fullStatementLineItemNeedsAnalystPass(
      request({
        statement: "income_statement",
        periodType: "duration",
        label: "Less: Net Income Attributable to Noncontrolling Interests",
        xbrlTag: "NetIncomeLossAttributableToNoncontrollingInterest",
        section: "net income"
      })
    ),
    false,
    "post-net-income attribution lines must remain structural reconciliation rows"
  );
  for (const [label, xbrlTag] of [
    ["Land", "Land"],
    ["Buildings and Improvements, Gross", "BuildingsAndImprovementsGross"],
    ["Machinery and Equipment, Gross", "MachineryAndEquipmentGross"],
    ["Construction in Progress, Gross", "ConstructionInProgressGross"],
    ["Property, Plant and Equipment, Gross", "PropertyPlantAndEquipmentGross"]
  ]) {
    assert.equal(
      fullStatementLineItemNeedsAnalystPass(
        request({
          statement: "balance_sheet",
          periodType: "instant",
          label,
          xbrlTag,
          section: "non-current assets",
          currentPeriodSourceLines: [
            "Property, Plant and Equipment, at cost, net of accumulated depreciation",
            label
          ]
        })
      ),
      false,
      `${label} must remain PP&E support detail when net PP&E is already reported`
    );
  }
  assert.equal(
    fullStatementLineItemNeedsAnalystPass(
      request({
        label: "Assets, current",
        xbrlTag: "AssetsCurrent",
        section: "current assets"
      })
    ),
    false
  );
  assert.equal(
    fullStatementLineItemNeedsAnalystPass(
      request({
        statement: "income_statement",
        periodType: "duration",
        label: "Operating income",
        xbrlTag: "OperatingIncomeLoss",
        section: "operating expenses"
      })
    ),
    false
  );
  assert.equal(
    fullStatementLineItemNeedsAnalystPass(
      request({
        label: "Treasury stock shares",
        xbrlTag: "TreasuryStockCommonShares",
        section: "equity",
        unit: "shares"
      })
    ),
    false
  );
  assert.equal(
    fullStatementLineItemNeedsAnalystPass(
      request({
        statement: "income_statement",
        periodType: "duration",
        label: "Diluted earnings per share",
        xbrlTag: "EarningsPerShareDiluted",
        section: "net income"
      })
    ),
    false
  );
  assert.equal(
    fullStatementLineItemNeedsAnalystPass(
      request({
        statement: "income_statement",
        periodType: "duration",
        label: "Basic",
        xbrlTag: "EarningsPerShareBasic",
        section: "operating expenses",
        unit: "USD/shares"
      })
    ),
    false,
    "camel-case EPS concepts must stay out of full-statement mapping coverage even when the filing label is only Basic/Diluted"
  );

  const aggregateOperatingExpenseRequest = request({
    statement: "income_statement",
    periodType: "duration",
    label: "Operating expenses",
    xbrlTag: "acme:TotalOperatingCostsAndExpenses",
    section: "operating expenses",
    amount: 519_000_000,
    currentPeriodSourceLines: [
      "Cost of product and service sold",
      "Research and development expense",
      "Selling and marketing expense",
      "General and administrative expense",
      "Operating expenses",
      "Operating income"
    ],
    deterministicCandidate: "Other Operating Income / Expense"
  });
  assert.equal(lineItemNeedsClassification(aggregateOperatingExpenseRequest), false);
  assert.equal(materialStatementLineItemNeedsAnalystPass(aggregateOperatingExpenseRequest), false);
  assert.equal(fullStatementLineItemNeedsAnalystPass(aggregateOperatingExpenseRequest), false);
  assert.equal(
    classificationPassesValidation(aggregateOperatingExpenseRequest, {
      source_line_item: "Operating expenses",
      recommended_action: "merge_into_other",
      recommended_model_row: "Other Operating Income / Expense",
      recommended_model_row_mappings: [],
      explicit_zero_rows: [],
      classification_type: "LLM aggregate expense assignment",
      is_current: null,
      is_debt: false,
      is_operating: true,
      is_tax_related: false,
      is_deferred_revenue_or_contract_liability: false,
      is_deferred_tax: false,
      is_subtotal: false,
      should_exclude_from_other_bucket: false,
      confidence: "high",
      reason: "Map the aggregate to other operating expense.",
      requires_validation: true,
      requires_revalidation: true,
      llm_used: true,
      mapping_passed_validation: true
    }),
    false,
    "an aggregate operating-expense total must not become a detail-row assignment when its reported components are available"
  );
  let aggregateOperatingExpenseLlmCalls = 0;
  const aggregateOperatingExpenseCoverage = await classifyFinancialStatementLineItems([aggregateOperatingExpenseRequest], {
    llm: {
      enabled: true,
      apiKey: "test-key",
      endpoint: "https://openrouter.ai/api/v1/chat/completions",
      model: "deepseek/deepseek-v4-flash",
      siteUrl: "http://localhost:3000",
      appTitle: "Historicals Solver",
      fetchImpl: async () => {
        aggregateOperatingExpenseLlmCalls += 1;
        throw new Error("aggregate subtotal should never be sent to the LLM");
      }
    },
    statementAnalystPass: { enabled: true, coverage: "all_primary_rows" }
  });
  assert.equal(aggregateOperatingExpenseLlmCalls, 0);
  assert.equal(aggregateOperatingExpenseCoverage.targetCount, 0);
  assert.equal(aggregateOperatingExpenseCoverage.classifications.length, 0);

  const failedPayloads = [];
  const failedResult = await classifyFinancialStatementLineItems(
    [
      request({
        sourceRowKey: "row-other-debt",
        rowOrder: 3,
        label: "Other debt obligations",
        xbrlTag: "OtherDebtObligations",
        section: "non-current liabilities",
        deterministicCandidate: "Other Non-Current Liabilities",
        uncertaintyReason: "debt-like obligations require LLM accounting review"
      })
    ],
    {
      llm: {
        enabled: true,
        apiKey: "test-key",
        endpoint: "https://example.test/chat/completions",
        model: "openai/gpt-5.2",
        siteUrl: "http://localhost:3000",
        appTitle: "Historicals Solver Test",
        timeoutMs: 100,
        fetchImpl: async (_url, init) => {
          failedPayloads.push(JSON.parse(init.body));
          return {
            ok: false,
            status: 404,
            statusText: "Not Found",
            headers: { get: () => null },
            json: async () => ({ error: { message: "No endpoints found that can handle the requested parameters" } })
          };
        }
      }
    }
  );
  assert.equal(failedPayloads.length, 1);
  assert.equal(failedResult.llmCalls, 0);
  assert.equal(failedResult.llmAttempts, 1);
  assert.equal(failedResult.llmSuccessfulCompletions, 0);
  assert.equal(failedResult.llmTelemetry[0].status, "attempted_failed");
  assert.equal(failedResult.llmTelemetry[0].routingError, true);
  const failedClassification = failedResult.classifications[0].classification;
  assert.equal(failedClassification.llm_used, false);
  assert.equal(failedClassification.confidence, "low");
  assert.equal(failedClassification.llm_status, "attempted_failed");

  const usageResult = await classifyFinancialStatementLineItems(
    [
      request({
        sourceRowKey: "row-other-current",
        rowOrder: 4,
        label: "Other current notes payable",
        xbrlTag: "NotesPayableCurrent",
        section: "current liabilities",
        deterministicCandidate: "Other Current Liabilities",
        uncertaintyReason: "notes payable could be current borrowings"
      })
    ],
    {
      llm: {
        enabled: true,
        apiKey: "test-key",
        endpoint: "https://example.test/chat/completions",
        model: "openai/gpt-5.2",
        siteUrl: "http://localhost:3000",
        appTitle: "Historicals Solver Test",
        timeoutMs: 100,
        fetchImpl: async () => ({
          ok: true,
          status: 200,
          statusText: "OK",
          headers: { get: (name) => (name === "X-Generation-Id" ? "gen-test-123" : null) },
          json: async () => ({
            id: "gen-test-123",
            model: "openai/gpt-5.2",
            usage: {
              prompt_tokens: 111,
              completion_tokens: 22,
              total_tokens: 133,
              cost: 0.001,
              completion_tokens_details: { reasoning_tokens: 7 },
              prompt_tokens_details: { cached_tokens: 5 }
            },
            choices: [
              {
                message: {
                  content: JSON.stringify({
                    classifications: [
                      {
                        source_row_key: "row-other-current",
                        source_line_item: "Other current notes payable",
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
                        reason: "Notes payable current is a current borrowing source.",
                        requires_validation: true,
                        requires_revalidation: true
                      }
                    ]
                  })
                }
              }
            ]
          })
        })
      }
    }
  );
  assert.equal(usageResult.llmCalls, 1);
  assert.equal(usageResult.llmTelemetry[0].generationId, "gen-test-123");
  assert.equal(usageResult.llmTelemetry[0].usage.promptTokens, 111);
  assert.equal(usageResult.llmTelemetry[0].usage.completionTokens, 22);
  assert.equal(usageResult.llmTelemetry[0].usage.cost, 0.001);

  let malformedAttempts = 0;
  const malformedResult = await classifyFinancialStatementLineItems(
    [
      request({
        sourceRowKey: "row-bad",
        rowOrder: 5,
        label: "Other deferred balance",
        xbrlTag: "OtherDeferredBalance",
        section: "current liabilities",
        deterministicCandidate: "Other Current Liabilities",
        uncertaintyReason: "deferred balance requires review"
      })
    ],
    {
      llm: {
        enabled: true,
        apiKey: "test-key",
        endpoint: "https://example.test/chat/completions",
        model: "openai/gpt-5.2",
        siteUrl: "http://localhost:3000",
        appTitle: "Historicals Solver Test",
        timeoutMs: 100,
        fetchImpl: async () => {
          malformedAttempts += 1;
          return {
            ok: true,
            status: 200,
            statusText: "OK",
            headers: { get: () => null },
            json: async () => ({
              choices: [
                {
                  message: {
                    content: malformedAttempts === 1 ? "{not-json" : JSON.stringify({ classifications: [] })
                  }
                }
              ]
            })
          };
        }
      }
    }
  );
  assert.equal(malformedAttempts, 1);
  assert.equal(malformedResult.llmCalls, 0);
  assert.equal(malformedResult.llmAttempts, 1);
  assert.equal(malformedResult.llmTelemetry[0].repairAttempted, false);
  assert.equal(malformedResult.llmTelemetry[0].status, "needs_human_review");
  assert.equal(malformedResult.classifications[0].classification.confidence, "low");
  assert.equal(malformedResult.classifications[0].classification.llm_status, "needs_human_review");
  assert.deepEqual(malformedResult.unreviewedTargetKeys, ["row-bad"]);

  let deadlineAbortObserved = false;
  const deadlineStartedAt = Date.now();
  const deadlineResult = await classifyFinancialStatementLineItems(
    [
      request({
        sourceRowKey: "row-deadline",
        rowOrder: 6,
        label: "Other current notes payable",
        xbrlTag: "NotesPayableCurrent",
        section: "current liabilities",
        deterministicCandidate: "Other Current Liabilities",
        uncertaintyReason: "notes payable requires statement-level review"
      })
    ],
    {
      llm: {
        enabled: true,
        apiKey: "test-key",
        endpoint: "https://example.test/chat/completions",
        model: "openai/gpt-5.2",
        siteUrl: "http://localhost:3000",
        appTitle: "Historicals Solver Test",
        timeoutMs: 1_000,
        deadlineAt: Date.now() + 50,
        fetchImpl: async (_url, init) =>
          new Promise((_resolve, reject) => {
            const onAbort = () => {
              deadlineAbortObserved = true;
              reject(new Error("deadline abort"));
            };
            if (init.signal?.aborted) onAbort();
            else init.signal?.addEventListener("abort", onAbort, { once: true });
          })
      }
    }
  );
  assert.equal(deadlineAbortObserved, true, "the statement classifier must forward its global LLM deadline to the controller");
  assert.ok(Date.now() - deadlineStartedAt < 500, "the statement classifier must not outlive its global LLM deadline through recursive retries");
  assert.equal(deadlineResult.llmCalls, 0);
  assert.equal(deadlineResult.classifications[0].classification.llm_used, false);

  console.log("Financial line item classifier rules passed.");
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
