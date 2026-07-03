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

const { addLlmMappingReviewSheet } = loadTypeScriptModule(path.join(repoRoot, "server", "fill-model", "fill-model-service.ts"));

const workbook = new ExcelJS.Workbook();
addLlmMappingReviewSheet(workbook, [
  {
    company: "Test Co.",
    ticker: "TST",
    reviewerModel: "openai/gpt-5.2",
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

console.log("LLM accounting controller workbook status guard passed.");
