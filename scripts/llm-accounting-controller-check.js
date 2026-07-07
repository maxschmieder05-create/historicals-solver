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
const { addLlmMappingReviewSheet } = loadTypeScriptModule(path.join(repoRoot, "server", "fill-model", "fill-model-service.ts"));

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

  console.log("LLM accounting controller workbook and fallback guards passed.");
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
