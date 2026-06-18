const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const Module = require("node:module");
const ts = require("typescript");

const repoRoot = path.resolve(__dirname, "..");
const sourcePath = path.join(repoRoot, "app", "api", "fill-model", "llm-workbench.ts");

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

const { buildLlmWorkbookToolbox, llmWorkbookToolboxSystemInstruction } = loadTypeScriptModule(sourcePath);

const toolbox = buildLlmWorkbookToolbox({
  company: { ticker: "TST", name: "Test Co.", cik: "0000000001" },
  periods: ["1Q25"],
  facts: [
    {
      period: "1Q25",
      periodType: "duration",
      concept: "RevenueFromContractWithCustomerExcludingAssessedTax",
      label: "Revenue",
      value: 500_000_000,
      valueMillions: 500,
      unit: "USD",
      accession: "0000000001-25-000001"
    },
    {
      period: "1Q25",
      periodType: "duration",
      concept: "InterestIncomeExpenseNonOperatingNet",
      label: "Interest income",
      value: 5_000_000,
      valueMillions: 5,
      unit: "USD",
      accession: "0000000001-25-000001"
    }
  ],
  statementRows: [
    {
      period: "1Q25",
      statementName: "Consolidated Statements of Operations",
      sourceTableType: "primary_statement",
      rowLabel: "Revenue",
      xbrlConcept: "RevenueFromContractWithCustomerExcludingAssessedTax",
      value: 500_000_000,
      valueMillions: 500,
      unit: "USD",
      accession: "0000000001-25-000001",
      rowOrder: 1
    }
  ],
  workbookCells: [
    {
      sheetName: "Model",
      cell: "F10",
      row: 10,
      column: 6,
      period: "1Q25",
      modelRowLabel: "Revenue",
      value: 500,
      isFormula: false,
      mappingType: "direct",
      validationStatus: "OK!",
      sourceLedgerStatus: "explicit_current_sec_source",
      sourceLineItemLabel: "Revenue",
      sourceXbrlTag: "RevenueFromContractWithCustomerExcludingAssessedTax"
    },
    {
      sheetName: "Model",
      cell: "F20",
      row: 20,
      column: 6,
      period: "1Q25",
      modelRowLabel: "Other Current Liabilities",
      value: 42,
      isFormula: false,
      mappingType: "skipped",
      validationStatus: "blocked: no source",
      sourceLedgerStatus: "stale_or_unsupported"
    }
  ],
  auditRows: [
    {
      period: "1Q25",
      cell: "Model!F20",
      modelRowLabel: "Other Current Liabilities",
      valueWritten: 42,
      mappingType: "skipped",
      validationStatus: "blocked: no source",
      conceptsUsed: "",
      notes: "No SEC source."
    }
  ],
  sourceLedgerRows: [
    {
      period: "1Q25",
      cell: "Model!F20",
      modelRow: 20,
      value: 42,
      sourceLineItemLabel: "",
      sourceXbrlTag: "",
      mappingStatus: "stale_or_unsupported",
      sourceStatement: "balance",
      sourceTableType: "primary_statement",
      classificationReason: ""
    }
  ],
  balanceSheetAssignments: [
    {
      period: "1Q25",
      sourceLineItemLabel: "Deferred revenue",
      sourceXbrlTag: "ContractWithCustomerLiabilityCurrent",
      amount: 42_000_000,
      amountMillions: 42,
      assignedModelRow: "Other Current Liabilities",
      assignmentStatus: "mapped_to_model_row",
      classificationReason: "Deferred revenue is an operating liability.",
      validationStatus: "OK!",
      side: "liabilities_and_equity",
      sourceSection: "current liabilities"
    }
  ],
  incomeStatementAssignments: [
    {
      period: "1Q25",
      sourceLineItemLabel: "Advertising expense",
      sourceXbrlTag: "AdvertisingExpense",
      sourceAmount: 12_000_000,
      sourceAmountMillions: 12,
      modelAmount: -12_000_000,
      modelAmountMillions: -12,
      assignedModelRow: "SG&A",
      assignmentStatus: "grouped_into_model_row",
      classificationReason: "Advertising expense is an SG&A operating expense.",
      validationStatus: "OK!",
      sourceSection: "operating expenses"
    }
  ],
  validationFailures: ["Model!F20 1Q25: hardcoded historical value has no current-company source ledger support."],
  warnings: ["Example warning"],
  limits: {
    maxFacts: 1,
    maxWorkbookCells: 10,
    maxStatementRows: 10,
    maxAuditRows: 10,
    maxSourceLedgerRows: 10,
    maxBalanceSheetAssignments: 10,
    maxIncomeStatementAssignments: 10
  }
});

assert.equal(toolbox.protocol, "historicals-solver.llm-workbench.v1");
assert.equal(toolbox.verificationGate.status, "failed");
assert.equal(toolbox.verificationGate.failClosed, true);
assert.match(toolbox.verificationGate.rule, /Do not approve or return/);
assert.deepEqual(
  toolbox.toolOutputs.map((tool) => tool.name),
  [
    "sec.get_company_facts",
    "sec.get_financial_statement_rows",
    "workbook.inspect_output_cells",
    "workbook.trace_source_ledger",
    "workbook.get_mapping_audit",
    "workbook.get_balance_sheet_assignment_ledger",
    "workbook.get_income_statement_assignment_ledger",
    "workbook.validate_return"
  ]
);
assert.equal(toolbox.omittedCounts.facts, 1);
assert.equal(toolbox.toolOutputs[0].output.facts[0].concept, "RevenueFromContractWithCustomerExcludingAssessedTax");
assert.equal(toolbox.toolOutputs[2].output.cells[0].sourceLedgerStatus, "stale_or_unsupported");
assert.equal(toolbox.toolOutputs[6].output.rows[0].assignedModelRow, "SG&A");
assert.match(llmWorkbookToolboxSystemInstruction(), /MCP-style tool outputs/);
assert.match(llmWorkbookToolboxSystemInstruction(), /workbook\.validate_return/);

console.log("LLM workbook workbench guard passed.");
