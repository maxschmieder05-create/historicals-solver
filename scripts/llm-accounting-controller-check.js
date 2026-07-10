const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const Module = require("node:module");
const ts = require("typescript");
const ExcelJS = require("exceljs");

const repoRoot = path.resolve(__dirname, "..");

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

const { requestAccountingJson } = loadTypeScriptModule(path.join(repoRoot, "server", "fill-model", "llm-accounting-controller.ts"));
const { classificationSourceKeys } = loadTypeScriptModule(path.join(repoRoot, "server", "fill-model", "financial-line-item-classifier.ts"));
const { addLlmMappingReviewSheet, applyLlmMappingReviewCorrections } = loadTypeScriptModule(
  path.join(repoRoot, "server", "fill-model", "fill-model-service.ts")
);

async function main() {
  let calls = 0;
  const result = await requestAccountingJson({
    purpose: "workbook_mapping_review",
    apiKey: "test-key",
    endpoint: "https://openrouter.ai/api/v1/chat/completions",
    model: "deepseek/deepseek-chat-v3.1",
    fallbackModels: ["deepseek/deepseek-v3.1-terminus"],
    siteUrl: "http://localhost:3000",
    appTitle: "Historicals Solver",
    messages: [{ role: "user", content: "{}" }],
    jsonSchema: { name: "test", schema: { type: "object" } },
    maxTokens: 50,
    timeoutMs: 1000,
    fetchImpl: async (_url, init) => {
      calls += 1;
      const body = JSON.parse(init.body);
      if (calls === 1) {
        assert.equal(body.model, "deepseek/deepseek-chat-v3.1");
        return new Response(JSON.stringify({ id: "gen-empty", choices: [{ message: { content: "" } }] }), { status: 200 });
      }
      assert.equal(body.model, "deepseek/deepseek-v3.1-terminus");
      return new Response(JSON.stringify({ id: "gen-ok", choices: [{ message: { content: "{\"ok\":true}" } }] }), { status: 200 });
    },
    validate: (value) => (value && value.ok === true ? { ok: true, value, validated: true, affectedOutput: true } : { ok: false, error: "bad" })
  });

  assert.equal(result.status, "completed_validated");
  assert.equal(result.value.ok, true);
  assert.equal(calls, 2);
  assert.equal(result.attemptTelemetry.length, 2);
  assert.equal(result.attemptTelemetry[0].model, "deepseek/deepseek-chat-v3.1");
  assert.equal(result.attemptTelemetry[0].status, "attempted_failed");
  assert.equal(result.attemptTelemetry[1].model, "deepseek/deepseek-v3.1-terminus");
  assert.equal(result.attemptTelemetry[1].status, "completed_validated");

  let quotaCalls = 0;
  const quotaResult = await requestAccountingJson({
    purpose: "statement_line_item_classification",
    apiKey: "test-key",
    endpoint: "https://openrouter.ai/api/v1/chat/completions",
    model: "deepseek/deepseek-v3.1-terminus",
    fallbackModels: ["deepseek/deepseek-r1-0528"],
    siteUrl: "http://localhost:3000",
    appTitle: "Historicals Solver",
    messages: [{ role: "user", content: "{}" }],
    jsonSchema: { name: "test", schema: { type: "object" } },
    maxTokens: 50,
    timeoutMs: 1000,
    fetchImpl: async (_url, init) => {
      quotaCalls += 1;
      const body = JSON.parse(init.body);
      if (quotaCalls === 1) {
        assert.equal(body.model, "deepseek/deepseek-v3.1-terminus");
        return new Response(
          JSON.stringify({ error: { message: "Key limit exceeded (total limit)." } }),
          { status: 403, statusText: "Forbidden" }
        );
      }
      assert.equal(body.model, "deepseek/deepseek-r1-0528");
      return new Response(JSON.stringify({ id: "gen-quota-ok", choices: [{ message: { content: "{\"ok\":true}" } }] }), { status: 200 });
    },
    validate: (value) => (value && value.ok === true ? { ok: true, value, validated: true, affectedOutput: true } : { ok: false, error: "bad" })
  });

  assert.equal(quotaResult.status, "attempted_failed");
  assert.equal(quotaCalls, 1);
  assert.equal(quotaResult.attemptTelemetry[0].httpStatus, 403);
  assert.equal(quotaResult.attemptTelemetry.length, 1);

  const bodyTimeoutResult = await requestAccountingJson({
    purpose: "statement_line_item_classification",
    apiKey: "test-key",
    endpoint: "https://openrouter.ai/api/v1/chat/completions",
    model: "nvidia/nemotron-3-super-120b-a12b:free",
    siteUrl: "http://localhost:3000",
    appTitle: "Historicals Solver",
    messages: [{ role: "user", content: "{}" }],
    jsonSchema: { name: "test", schema: { type: "object" } },
    maxTokens: 50,
    maxAttemptsPerModel: 1,
    timeoutMs: 5,
    fetchImpl: async () => ({
      ok: true,
      status: 200,
      statusText: "OK",
      headers: { get: () => null },
      json: async () => new Promise(() => {})
    }),
    validate: () => ({ ok: false, error: "unreachable" })
  });
  assert.equal(bodyTimeoutResult.status, "attempted_failed");
  assert.equal(bodyTimeoutResult.error.includes("timed out after 5ms"), true);

  const workbook = new ExcelJS.Workbook();
  addLlmMappingReviewSheet(workbook, [
    {
      company: "Test Co.",
      ticker: "TST",
      reviewerModel: "deepseek/deepseek-r1-0528",
      llmStatus: "attempted_failed",
      llmAttempts: 1,
      llmSuccessfulCompletions: 0,
      llmValidatedCompletions: 0,
      llmAffectedOutputDecisions: 0,
      llmGenerationId: "",
      llmUsageSummary: "",
      reviewStatus: "skipped",
      coverageSummary: "LLM mapping review attempted_failed (No endpoints found).",
      verifiedDecisionCount: 0,
      severity: "warning",
      issueType: "needs_human_review",
      period: "",
      sourceLineItemLabel: "",
      sourceXbrlTag: "",
      currentModelRow: "",
      recommendedModelRow: "",
      reason: "Router failed before inference.",
      reusableRule: "Run deterministic workbook validation and source-ledger checks when LLM review is unavailable.",
      evidence: []
    }
  ]);

  const sheet = workbook.getWorksheet("LLM Mapping Review");
  assert.ok(sheet);
  const headers = sheet.getRow(1).values.slice(1);
  assert.ok(headers.includes("LLM status"));
  assert.ok(headers.includes("LLM attempts"));
  assert.ok(headers.includes("LLM completions"));
  assert.ok(headers.includes("LLM generation id"));
  assert.equal(sheet.getRow(2).getCell(headers.indexOf("LLM status") + 1).value, "attempted_failed");
  assert.equal(sheet.getRow(2).getCell(headers.indexOf("LLM attempts") + 1).value, 1);
  assert.equal(sheet.getRow(2).getCell(headers.indexOf("review status") + 1).value, "skipped");

  const source = {
    period: "1Q26",
    accession: "0000000000-26-000001",
    xbrlTag: "OtherAccruedLiabilitiesCurrent",
    label: "Accrued rebates and returns",
    amount: 125_000_000
  };
  const classification = {
    source_line_item: source.label,
    recommended_action: "merge_into_other",
    recommended_model_row: "Other Current Liabilities",
    recommended_model_row_mappings: [],
    explicit_zero_rows: [],
    classification_type: "other current liability",
    is_current: true,
    is_debt: false,
    is_operating: true,
    is_tax_related: false,
    is_deferred_revenue_or_contract_liability: false,
    is_deferred_tax: false,
    is_subtotal: false,
    should_exclude_from_other_bucket: false,
    confidence: "high",
    reason: "Initial statement classification.",
    requires_validation: true,
    requires_revalidation: true,
    llm_used: true,
    llm_status: "completed_validated",
    mapping_passed_validation: true
  };
  const classifications = new Map();
  classificationSourceKeys(source).forEach((key) => classifications.set(key, classification));
  const issue = {
    severity: "error",
    issueType: "wrong_model_row",
    period: source.period,
    sourceLineItemLabel: source.label,
    sourceXbrlTag: source.xbrlTag,
    currentModelRow: "Other Current Liabilities",
    recommendedModelRow: "Accrued Liabilities",
    reason: "The source is an accrued operating liability and the template has a dedicated accrued-liabilities row.",
    reusableRule: "Prefer a dedicated accrued-liabilities row over a generic current-liability bucket.",
    evidence: ["Primary balance sheet / current liabilities"]
  };
  const balanceAssignment = {
    fiscalPeriod: source.period,
    sourceFilingAccession: source.accession,
    sourceStatement: "Consolidated Balance Sheets",
    sourceLineItemLabel: source.label,
    amount: source.amount,
    sourceXbrlTag: source.xbrlTag,
    assignedModelRow: "Other Current Liabilities",
    assignmentStatus: "grouped_into_model_row",
    classificationReason: classification.reason,
    llmUsed: true,
    validationStatus: "OK!",
    side: "liabilities_and_equity",
    sourceSection: "current liabilities",
    sourceRowKey: "row-1"
  };
  const correction = applyLlmMappingReviewCorrections({
    issues: [issue],
    classifications,
    balanceAssignments: [balanceAssignment],
    incomeAssignments: [],
    availableModelRows: ["Accrued Liabilities", "Other Current Liabilities"]
  });
  assert.equal(correction.changed, true);
  assert.equal(correction.repairs.length, 1);
  const corrected = classifications.get(classificationSourceKeys(source)[0]);
  assert.equal(corrected.recommended_model_row, "Accrued Liabilities");
  assert.equal(corrected.recommended_action, "remap");
  assert.equal(corrected.mapping_passed_validation, true);
  assert.equal(corrected.reason.startsWith("LLM workbook reviewer correction:"), true);

  const rejectedCorrection = applyLlmMappingReviewCorrections({
    issues: [{ ...issue, recommendedModelRow: "PP&E, Net" }],
    classifications,
    balanceAssignments: [balanceAssignment],
    incomeAssignments: [],
    availableModelRows: ["PP&E, Net", "Other Current Liabilities"]
  });
  assert.equal(rejectedCorrection.changed, false);
  assert.equal(rejectedCorrection.rejected.some((item) => /section\/side validation/.test(item)), true);

  console.log("LLM accounting controller workbook and fallback guards passed.");
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
