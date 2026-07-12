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

const {
  requestAccountingJson,
  aggregateAccountingTelemetry,
  __llmAccountingControllerTestHooks
} = loadTypeScriptModule(path.join(repoRoot, "server", "fill-model", "llm-accounting-controller.ts"));
const { classificationSourceKeys } = loadTypeScriptModule(path.join(repoRoot, "server", "fill-model", "financial-line-item-classifier.ts"));
const { addLlmMappingReviewSheet, applyLlmMappingReviewCorrections, __fillModelServiceTestHooks } = loadTypeScriptModule(
  path.join(repoRoot, "server", "fill-model", "fill-model-service.ts")
);

async function checkLlmMappingReviewAvailabilityPolicy() {
  const envNames = [
    "LLM_MAPPING_ENABLED",
    "LLM_MAPPING_REVIEW_ENABLED",
    "LLM_MAPPING_REVIEW_BLOCKING",
    "OPENROUTER_API_KEY",
    "OPENAI_API_KEY",
    "ACCOUNTING_LLM_API_KEY"
  ];
  const previousEnv = Object.fromEntries(envNames.map((name) => [name, process.env[name]]));
  const company = { cik: "0000000000", ticker: "TST", title: "Test Co." };
  const debug = {
    step() {},
    warn() {},
    error() {}
  };
  const runEarlyReview = (state) =>
    __fillModelServiceTestHooks.runLlmMappingReview(
      company,
      [],
      [],
      [],
      [],
      [],
      [],
      [],
      {},
      state,
      debug
    );

  try {
    process.env.LLM_MAPPING_ENABLED = "false";
    process.env.LLM_MAPPING_REVIEW_ENABLED = "true";
    for (const keyName of ["OPENROUTER_API_KEY", "OPENAI_API_KEY", "ACCOUNTING_LLM_API_KEY"]) {
      delete process.env[keyName];
    }

    process.env.LLM_MAPPING_REVIEW_BLOCKING = "true";
    const missingCredentialBlocking = await runEarlyReview(__fillModelServiceTestHooks.createLlmMappingState());
    assert.equal(missingCredentialBlocking.rows[0].llmStatus, "unavailable");
    assert.equal(missingCredentialBlocking.warnings.length, 1);
    assert.deepEqual(
      missingCredentialBlocking.blockingErrors,
      missingCredentialBlocking.warnings,
      "blocking review mode must fail closed when its endpoint credential is unavailable"
    );

    process.env.LLM_MAPPING_REVIEW_BLOCKING = "false";
    const missingCredentialAdvisory = await runEarlyReview(__fillModelServiceTestHooks.createLlmMappingState());
    assert.equal(missingCredentialAdvisory.warnings.length, 1);
    assert.deepEqual(
      missingCredentialAdvisory.blockingErrors,
      [],
      "review unavailability must remain advisory unless blocking mode is explicitly enabled"
    );

    process.env.LLM_MAPPING_REVIEW_ENABLED = "false";
    process.env.LLM_MAPPING_REVIEW_BLOCKING = "true";
    const explicitlyDisabled = await runEarlyReview(__fillModelServiceTestHooks.createLlmMappingState());
    assert.equal(explicitlyDisabled.rows[0].llmStatus, "disabled");
    assert.deepEqual(explicitlyDisabled.warnings, []);
    assert.deepEqual(
      explicitlyDisabled.blockingErrors,
      [],
      "an explicitly disabled review must not become a blocking failure"
    );

    process.env.LLM_MAPPING_REVIEW_ENABLED = "true";
    process.env.LLM_MAPPING_REVIEW_BLOCKING = "true";
    process.env.LLM_MAPPING_ENABLED = "true";
    process.env.OPENROUTER_API_KEY = "test-key";
    process.env.OPENAI_API_KEY = "test-key";
    process.env.ACCOUNTING_LLM_API_KEY = "test-key";
    const analystState = __fillModelServiceTestHooks.createLlmMappingState();
    assert.equal(analystState.analystMode, true, "LLM analyst mode must own the workflow when mapping and review are available");
    assert.equal(__fillModelServiceTestHooks.llmAnalystControlsValidation(analystState), true);
    assert.ok(analystState.reservedReviewCalls > 0, "classification must reserve calls for validation recovery");
    const reviewToolbox = (blockingFailures) => ({ verificationGate: { blockingFailures } });
    assert.ok(
      __fillModelServiceTestHooks.llmMappingReviewItemLimit(reviewToolbox([])) >
        __fillModelServiceTestHooks.llmMappingReviewItemLimit(reviewToolbox(["failed accounting equation"])),
      "validation recovery must use a smaller, context-safe evidence workbench than final review"
    );
    assert.equal(__fillModelServiceTestHooks.llmMappingReviewItemLimit(reviewToolbox(["failed accounting equation"])), 220);
    analystState.attempts = analystState.maxCalls - analystState.reservedReviewCalls;
    assert.equal(__fillModelServiceTestHooks.llmMappingAttemptsRemaining(analystState), 0);
    assert.equal(
      __fillModelServiceTestHooks.llmMappingCanUse(analystState, 1),
      false,
      "pre-fill classification must stop before consuming the review/recovery reserve"
    );

    const exhaustedState = __fillModelServiceTestHooks.createLlmMappingState();
    exhaustedState.attempts = exhaustedState.maxCalls;
    const exhaustedBudgetBlocking = await runEarlyReview(exhaustedState);
    assert.equal(exhaustedBudgetBlocking.rows[0].llmStatus, "unavailable");
    assert.equal(exhaustedBudgetBlocking.warnings.length, 1);
    assert.deepEqual(
      exhaustedBudgetBlocking.blockingErrors,
      exhaustedBudgetBlocking.warnings,
      "blocking review mode must fail closed when the shared request budget is exhausted before review"
    );
  } finally {
    for (const name of envNames) {
      if (previousEnv[name] === undefined) delete process.env[name];
      else process.env[name] = previousEnv[name];
    }
  }
}

async function main() {
  await checkLlmMappingReviewAvailabilityPolicy();
  for (const falsyValue of [false, 0]) {
    let falsyCalls = 0;
    const falsyResult = await requestAccountingJson({
      purpose: "row_mapping",
      apiKey: "test-key",
      endpoint: "https://openrouter.ai/api/v1/chat/completions",
      model: "deepseek/deepseek-v4-flash",
      siteUrl: "http://localhost:3000",
      appTitle: "Historicals Solver",
      messages: [{ role: "user", content: "Return the requested scalar." }],
      jsonSchema: { name: "falsy", schema: {} },
      maxTokens: 10,
      timeoutMs: 1000,
      fetchImpl: async () => {
        falsyCalls += 1;
        return new Response(
          JSON.stringify({ choices: [{ message: { content: JSON.stringify(falsyValue) } }] }),
          { status: 200 }
        );
      },
      validate: (value) =>
        Object.is(value, falsyValue)
          ? { ok: true, value, validated: true, affectedOutput: true }
          : { ok: false, error: "falsy scalar mismatch" }
    });
    assert.equal(falsyResult.status, "completed_validated");
    assert.equal(falsyResult.value, falsyValue);
    assert.equal(falsyCalls, 1);
  }

  const repairBodies = [];
  let repairCalls = 0;
  const repairedResult = await requestAccountingJson({
    purpose: "workbook_mapping_review",
    apiKey: "test-key",
    endpoint: "https://openrouter.ai/api/v1/chat/completions",
    model: "deepseek/deepseek-v4-flash",
    siteUrl: "http://localhost:3000",
    appTitle: "Historicals Solver",
    messages: [{ role: "user", content: "Return a validated object." }],
    jsonSchema: { name: "repair_test", schema: { type: "object" } },
    maxTokens: 50,
    timeoutMs: 1000,
    repair: { enabled: true, instruction: "Repair the invalid object." },
    fetchImpl: async (_url, init) => {
      repairCalls += 1;
      repairBodies.push(JSON.parse(init.body));
      if (repairCalls === 1) {
        return new Response(
          JSON.stringify({
            id: "gen-invalid",
            choices: [{ message: { content: '{"ok":false}' } }],
            usage: { prompt_tokens: 3, completion_tokens: 2, total_tokens: 5, cost: 0.01 }
          }),
          { status: 200 }
        );
      }
      return new Response(
        JSON.stringify({
          id: "gen-repaired",
          choices: [{ message: { content: '{"ok":true}' } }],
          usage: { prompt_tokens: 4, completion_tokens: 1, total_tokens: 5, cost: 0.02 }
        }),
        { status: 200 }
      );
    },
    validate: (value) =>
      value && value.ok === true
        ? { ok: true, value, validated: true, affectedOutput: true }
        : { ok: false, error: "validation rejected payload", needsHumanReview: true }
  });
  assert.equal(repairedResult.status, "repaired");
  assert.equal(repairCalls, 2);
  assert.equal(repairBodies[1].messages.at(-2).role, "assistant");
  assert.equal(repairBodies[1].messages.at(-2).content, '{"ok":false}');
  assert.match(repairBodies[1].messages.at(-1).content, /validation rejected payload/);
  assert.equal(repairBodies[1].response_format.type, "json_schema");
  assert.equal(repairedResult.attemptTelemetry.length, 2);
  assert.equal(repairedResult.attemptTelemetry[0].status, "needs_human_review");
  assert.equal(repairedResult.attemptTelemetry[0].completed, true);
  assert.equal(repairedResult.attemptTelemetry[0].validated, false);
  assert.equal(repairedResult.attemptTelemetry[1].status, "repaired");
  const repairedAggregate = aggregateAccountingTelemetry(repairedResult.attemptTelemetry);
  assert.equal(repairedAggregate.attempts, 2);
  assert.equal(repairedAggregate.successfulCompletions, 2);
  assert.equal(repairedAggregate.validatedCompletions, 1);
  assert.equal(repairedAggregate.failedAttempts, 1);
  assert.equal(repairedAggregate.usage.promptTokens, 7);
  assert.equal(repairedAggregate.usage.completionTokens, 3);
  assert.equal(repairedAggregate.usage.totalTokens, 10);
  assert.equal(repairedAggregate.usage.cost, 0.03);

  let oneAttemptRepairCalls = 0;
  const oneAttemptRepairResult = await requestAccountingJson({
    purpose: "workbook_mapping_review",
    apiKey: "test-key",
    endpoint: "https://openrouter.ai/api/v1/chat/completions",
    model: "deepseek/deepseek-v4-flash",
    siteUrl: "http://localhost:3000",
    appTitle: "Historicals Solver",
    messages: [{ role: "user", content: "Return a validated object." }],
    jsonSchema: { name: "bounded_repair_test", schema: { type: "object" } },
    maxTokens: 50,
    maxTotalAttempts: 1,
    timeoutMs: 1000,
    repair: { enabled: true, instruction: "Repair the invalid object." },
    fetchImpl: async () => {
      oneAttemptRepairCalls += 1;
      return new Response(JSON.stringify({ choices: [{ message: { content: '{"ok":false}' } }] }), { status: 200 });
    },
    validate: (value) =>
      value && value.ok === true
        ? { ok: true, value, validated: true, affectedOutput: true }
        : { ok: false, error: "validation rejected payload", needsHumanReview: true }
  });
  assert.equal(oneAttemptRepairCalls, 1, "a repair request must not exceed the caller's remaining attempt budget");
  assert.equal(oneAttemptRepairResult.attemptTelemetry.length, 1);
  assert.equal(oneAttemptRepairResult.status, "needs_human_review");

  let gptRequestBody;
  await requestAccountingJson({
    purpose: "row_mapping",
    apiKey: "test-key",
    endpoint: "https://openrouter.ai/api/v1/chat/completions",
    model: "openai/gpt-5.2",
    siteUrl: "http://localhost:3000",
    appTitle: "Historicals Solver",
    messages: [{ role: "user", content: "{}" }],
    jsonSchema: { name: "token_parameter", schema: { type: "object" } },
    maxTokens: 77,
    timeoutMs: 1000,
    fetchImpl: async (_url, init) => {
      gptRequestBody = JSON.parse(init.body);
      return new Response(JSON.stringify({ choices: [{ message: { content: '{"ok":true}' } }] }), { status: 200 });
    },
    validate: (value) => (value && value.ok === true ? { ok: true, value, validated: true } : { ok: false, error: "bad" })
  });
  assert.equal(gptRequestBody.max_completion_tokens, 77);
  assert.equal("max_tokens" in gptRequestBody, false);

  const fallbackCapabilities = __llmAccountingControllerTestHooks.staticCapabilitiesForModel("unlisted/provider-model");
  let probeAborted = false;
  let probeUrl = "";
  let probeHeaders;
  const probeStartedAt = Date.now();
  const probedCapabilities = await __llmAccountingControllerTestHooks.probeModelCapabilities(
    "unlisted/provider-model",
    async (url, init) => {
      probeUrl = String(url);
      probeHeaders = init.headers;
      return new Promise((_resolve, reject) => {
        init.signal.addEventListener("abort", () => {
          probeAborted = true;
          reject(new Error("probe aborted"));
        });
      });
    },
    "probe-secret",
    20,
    fallbackCapabilities
  );
  assert.match(probeUrl, /^https:\/\/openrouter\.ai\//);
  assert.equal(probeHeaders.Authorization, "Bearer probe-secret");
  assert.equal(probeAborted, true);
  assert.ok(Date.now() - probeStartedAt < 250);
  assert.deepEqual(probedCapabilities, fallbackCapabilities);

  const previousCapabilityCacheTtl = process.env.OPENROUTER_MODEL_CAPABILITY_CACHE_TTL_MS;
  const previousCapabilityFailureCacheTtl = process.env.OPENROUTER_MODEL_CAPABILITY_FAILURE_CACHE_TTL_MS;
  try {
    process.env.OPENROUTER_MODEL_CAPABILITY_CACHE_TTL_MS = "1000";
    process.env.OPENROUTER_MODEL_CAPABILITY_FAILURE_CACHE_TTL_MS = "200";
    __llmAccountingControllerTestHooks.clearModelCapabilitiesCache();

    let capabilityProbeCalls = 0;
    let releaseFailedProbe;
    const capabilityFetch = async () => {
      capabilityProbeCalls += 1;
      if (capabilityProbeCalls === 1) {
        return new Promise((resolve) => {
          releaseFailedProbe = () =>
            resolve(
              new Response(JSON.stringify({ error: { message: "temporary model-catalog outage" } }), {
                status: 503,
                statusText: "Service Unavailable"
              })
            );
        });
      }
      return new Response(
        JSON.stringify({
          data: [
            {
              id: "unlisted/retryable-provider-model",
              supported_parameters: ["response_format", "structured_outputs", "temperature", "max_tokens"]
            }
          ]
        }),
        { status: 200 }
      );
    };
    const retryableFallback = __llmAccountingControllerTestHooks.staticCapabilitiesForModel(
      "unlisted/retryable-provider-model"
    );
    const firstCapabilityProbe = __llmAccountingControllerTestHooks.capabilitiesForModel(
      "unlisted/retryable-provider-model",
      capabilityFetch,
      true,
      "probe-secret",
      1000
    );
    const dedupedCapabilityProbe = __llmAccountingControllerTestHooks.capabilitiesForModel(
      "unlisted/retryable-provider-model",
      capabilityFetch,
      true,
      "probe-secret",
      1000
    );
    assert.equal(capabilityProbeCalls, 1, "concurrent capability lookups must share one in-flight probe");
    releaseFailedProbe();
    assert.deepEqual(await firstCapabilityProbe, retryableFallback);
    assert.deepEqual(await dedupedCapabilityProbe, retryableFallback);

    const cachedFailure = await __llmAccountingControllerTestHooks.capabilitiesForModel(
      "unlisted/retryable-provider-model",
      capabilityFetch,
      true,
      "probe-secret",
      1000
    );
    assert.deepEqual(cachedFailure, retryableFallback);
    assert.equal(capabilityProbeCalls, 1, "a transient probe failure should be briefly deduped inside its retry TTL");
    const failureCacheState = __llmAccountingControllerTestHooks.modelCapabilitiesCacheState();
    assert.equal(failureCacheState.length, 1);
    assert.ok(Number.isFinite(failureCacheState[0].expiresAt), "a failed probe must never be cached for process lifetime");

    await new Promise((resolve) => setTimeout(resolve, Math.max(0, failureCacheState[0].expiresAt - Date.now()) + 10));
    const recoveredCapabilities = await __llmAccountingControllerTestHooks.capabilitiesForModel(
      "unlisted/retryable-provider-model",
      capabilityFetch,
      true,
      "probe-secret",
      1000
    );
    assert.equal(capabilityProbeCalls, 2, "an expired failure entry must trigger a fresh model-capability probe");
    assert.equal(recoveredCapabilities.source, "router");
    assert.equal(recoveredCapabilities.supportsStructuredOutputs, true);

    const cachedSuccess = await __llmAccountingControllerTestHooks.capabilitiesForModel(
      "unlisted/retryable-provider-model",
      capabilityFetch,
      true,
      "probe-secret",
      1000
    );
    assert.deepEqual(cachedSuccess, recoveredCapabilities);
    assert.equal(capabilityProbeCalls, 2, "a successful capability probe should remain cached inside its success TTL");
  } finally {
    __llmAccountingControllerTestHooks.clearModelCapabilitiesCache();
    if (previousCapabilityCacheTtl === undefined) delete process.env.OPENROUTER_MODEL_CAPABILITY_CACHE_TTL_MS;
    else process.env.OPENROUTER_MODEL_CAPABILITY_CACHE_TTL_MS = previousCapabilityCacheTtl;
    if (previousCapabilityFailureCacheTtl === undefined) delete process.env.OPENROUTER_MODEL_CAPABILITY_FAILURE_CACHE_TTL_MS;
    else process.env.OPENROUTER_MODEL_CAPABILITY_FAILURE_CACHE_TTL_MS = previousCapabilityFailureCacheTtl;
  }

  let finalAttemptCalls = 0;
  const retryCallTimes = [];
  const finalAttemptResult = await requestAccountingJson({
    purpose: "statement_line_item_classification",
    apiKey: "test-key",
    endpoint: "https://openrouter.ai/api/v1/chat/completions",
    model: "deepseek/deepseek-v4-flash",
    siteUrl: "http://localhost:3000",
    appTitle: "Historicals Solver",
    messages: [{ role: "user", content: "{}" }],
    jsonSchema: { name: "final_attempt", schema: { type: "object" } },
    maxTokens: 25,
    maxAttemptsPerModel: 2,
    timeoutMs: 1000,
    fetchImpl: async () => {
      finalAttemptCalls += 1;
      retryCallTimes.push(Date.now());
      if (finalAttemptCalls === 1) {
        return new Response(JSON.stringify({ choices: [{ message: { content: '{"ok":false}' } }] }), { status: 200 });
      }
      return new Response(JSON.stringify({ error: { message: "final provider unavailable" } }), {
        status: 503,
        statusText: "Service Unavailable"
      });
    },
    validate: (value) =>
      value && value.ok === true
        ? { ok: true, value, validated: true }
        : { ok: false, error: "validation rejected payload", needsHumanReview: true }
  });
  assert.equal(finalAttemptCalls, 2);
  assert.equal(finalAttemptResult.status, "attempted_failed");
  assert.equal(finalAttemptResult.telemetry.httpStatus, 503);
  assert.equal(finalAttemptResult.error, "final provider unavailable");
  assert.equal(finalAttemptResult.attemptTelemetry.length, 2);
  assert.ok(retryCallTimes[1] > retryCallTimes[0], "same-model retries should use bounded backoff with jitter");

  let boundedFallbackCalls = 0;
  const boundedFallbackResult = await requestAccountingJson({
    purpose: "statement_line_item_classification",
    apiKey: "test-key",
    endpoint: "https://openrouter.ai/api/v1/chat/completions",
    model: "deepseek/deepseek-v4-flash",
    fallbackModels: ["deepseek/deepseek-v3.1-terminus", "anthropic/claude-sonnet-4"],
    siteUrl: "http://localhost:3000",
    appTitle: "Historicals Solver",
    messages: [{ role: "user", content: "{}" }],
    jsonSchema: { name: "bounded_fallback", schema: { type: "object" } },
    maxTokens: 25,
    maxAttemptsPerModel: 2,
    maxTotalAttempts: 3,
    timeoutMs: 1000,
    fetchImpl: async () => {
      boundedFallbackCalls += 1;
      return new Response(JSON.stringify({ error: { message: "provider unavailable" } }), {
        status: 503,
        statusText: "Service Unavailable"
      });
    },
    validate: () => ({ ok: false, error: "not reached", needsHumanReview: true })
  });
  assert.equal(boundedFallbackCalls, 3, "same-model retries and fallback models must share one total attempt ceiling");
  assert.equal(boundedFallbackResult.attemptTelemetry.filter((item) => item.attempted).length, 3);

  let transportFailureCalls = 0;
  const transportFailureRecovery = await requestAccountingJson({
    purpose: "statement_line_item_classification",
    apiKey: "test-key",
    endpoint: "https://openrouter.ai/api/v1/chat/completions",
    model: "deepseek/deepseek-v4-flash",
    fallbackModels: ["anthropic/claude-sonnet-4"],
    siteUrl: "http://localhost:3000",
    appTitle: "Historicals Solver",
    messages: [{ role: "user", content: "{}" }],
    jsonSchema: { name: "transport_failure_recovery", schema: { type: "object" } },
    maxTokens: 25,
    maxAttemptsPerModel: 1,
    maxTotalAttempts: 2,
    timeoutMs: 1000,
    fetchImpl: async (_url, init) => {
      transportFailureCalls += 1;
      const body = JSON.parse(init.body);
      if (transportFailureCalls === 1) {
        assert.equal(body.model, "deepseek/deepseek-v4-flash");
        throw new TypeError("fetch failed");
      }
      assert.equal(body.model, "anthropic/claude-sonnet-4");
      return new Response(JSON.stringify({ choices: [{ message: { content: '{"ok":true}' } }] }), { status: 200 });
    },
    validate: (value) =>
      value && value.ok === true
        ? { ok: true, value, validated: true, affectedOutput: true }
        : { ok: false, error: "not recovered", needsHumanReview: true }
  });
  assert.equal(transportFailureCalls, 2, "transient transport errors should advance to a configured fallback model");

  let timeoutFallbackCalls = 0;
  const timeoutFallbackStartedAt = Date.now();
  const timeoutFallbackRecovery = await requestAccountingJson({
    purpose: "workbook_mapping_review",
    apiKey: "test-key",
    endpoint: "https://openrouter.ai/api/v1/chat/completions",
    model: "openai/gpt-5.6-sol-pro",
    fallbackModels: ["anthropic/claude-sonnet-5"],
    siteUrl: "http://localhost:3000",
    appTitle: "Historicals Solver",
    messages: [{ role: "user", content: "{}" }],
    jsonSchema: { name: "timeout_fallback_recovery", schema: { type: "object" } },
    maxTokens: 25,
    maxAttemptsPerModel: 1,
    maxTotalAttempts: 2,
    timeoutMs: 160,
    fetchImpl: async (_url, init) => {
      timeoutFallbackCalls += 1;
      const body = JSON.parse(init.body);
      if (body.model === "openai/gpt-5.6-sol-pro") {
        return new Promise((_resolve, reject) => {
          init.signal.addEventListener("abort", () => reject(new Error("preferred model timed out")), { once: true });
        });
      }
      assert.equal(body.model, "anthropic/claude-sonnet-5");
      return new Response(JSON.stringify({ choices: [{ message: { content: '{"ok":true}' } }] }), { status: 200 });
    },
    validate: (value) =>
      value && value.ok === true
        ? { ok: true, value, validated: true, affectedOutput: true }
        : { ok: false, error: "not recovered", needsHumanReview: true }
  });
  assert.equal(timeoutFallbackRecovery.status, "completed_validated");
  assert.equal(timeoutFallbackCalls, 2, "a preferred-model timeout must preserve time for the configured fallback");
  assert.ok(Date.now() - timeoutFallbackStartedAt < 300, "timeout fallback must remain inside the shared deadline");
  assert.equal(transportFailureRecovery.status, "completed_validated");
  assert.equal(transportFailureRecovery.attemptTelemetry.length, 2);
  assert.match(transportFailureRecovery.attemptTelemetry[0].errorMessage, /fetch failed/i);

  let invalidProviderRequestCalls = 0;
  const invalidProviderRequest = await requestAccountingJson({
    purpose: "statement_line_item_classification",
    apiKey: "test-key",
    endpoint: "https://openrouter.ai/api/v1/chat/completions",
    model: "deepseek/deepseek-v4-flash",
    fallbackModels: ["anthropic/claude-sonnet-4"],
    siteUrl: "http://localhost:3000",
    appTitle: "Historicals Solver",
    messages: [{ role: "user", content: "{}" }],
    jsonSchema: { name: "invalid_provider_request", schema: { type: "object" } },
    maxTokens: 25,
    maxAttemptsPerModel: 2,
    timeoutMs: 1000,
    fetchImpl: async () => {
      invalidProviderRequestCalls += 1;
      return new Response(JSON.stringify({ error: { message: "Invalid JSON schema parameter" } }), {
        status: 400,
        statusText: "Bad Request"
      });
    },
    validate: () => ({ ok: false, error: "not reached", needsHumanReview: true })
  });
  assert.equal(invalidProviderRequest.status, "attempted_failed");
  assert.equal(invalidProviderRequest.telemetry.httpStatus, 400);
  assert.equal(invalidProviderRequestCalls, 1, "non-transient provider 4xx errors must not retry or advance to fallbacks");

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
  assert.match(
    bodyTimeoutResult.error,
    /(?:timed out after [1-5]ms|exceeded its total 5ms request budget)/,
    "the complete controller request, including body parsing, must stay inside the total timeout budget"
  );

  const cancellationController = new AbortController();
  let cancellationFetchStarted;
  const cancellationFetchStartedPromise = new Promise((resolve) => {
    cancellationFetchStarted = resolve;
  });
  const cancellationStartedAt = Date.now();
  const cancelledLlmPromise = requestAccountingJson({
    purpose: "statement_line_item_classification",
    apiKey: "test-key",
    endpoint: "https://openrouter.ai/api/v1/chat/completions",
    model: "deepseek/deepseek-v4-flash",
    siteUrl: "http://localhost:3000",
    appTitle: "Historicals Solver",
    messages: [{ role: "user", content: "{}" }],
    jsonSchema: { name: "cancel_test", schema: { type: "object" } },
    maxTokens: 50,
    maxAttemptsPerModel: 2,
    timeoutMs: 30_000,
    signal: cancellationController.signal,
    fetchImpl: async () => {
      cancellationFetchStarted();
      return new Promise(() => {});
    },
    validate: () => ({ ok: false, error: "unreachable" })
  });
  await cancellationFetchStartedPromise;
  cancellationController.abort();
  const cancelledLlmResult = await cancelledLlmPromise;
  assert.equal(cancelledLlmResult.status, "attempted_failed");
  assert.match(cancelledLlmResult.error, /cancelled because the workbook fill request was cancelled/i);
  assert.ok(Date.now() - cancellationStartedAt < 500, "client cancellation must stop an in-flight LLM wait without trying fallbacks or repairs");

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

  const unclassifiedDebtSource = {
    period: "3Q26",
    accession: "0000000000-26-000002",
    xbrlTag: "DeferredFinanceCostsNet",
    label: "Less unamortized debt discounts and issuance costs",
    amount: 14000000
  };
  const unclassifiedDebtAssignment = {
    fiscalPeriod: unclassifiedDebtSource.period,
    sourceFilingAccession: unclassifiedDebtSource.accession,
    sourceStatement: "Consolidated Balance Sheets",
    sourceLineItemLabel: unclassifiedDebtSource.label,
    amount: unclassifiedDebtSource.amount,
    sourceXbrlTag: unclassifiedDebtSource.xbrlTag,
    assignedModelRow: "",
    assignmentStatus: "explicitly_excluded_with_reason",
    classificationReason: "Could not determine asset, liability, or equity section for this primary balance sheet line.",
    llmUsed: false,
    validationStatus: "OK!",
    side: "unknown",
    sourceSection: "unknown",
    sourceRowKey: "row-2"
  };
  const missingClassificationStore = new Map();
  const missingClassificationCorrection = applyLlmMappingReviewCorrections({
    issues: [
      {
        severity: "error",
        issueType: "unsupported_exclusion",
        period: unclassifiedDebtSource.period,
        sourceLineItemLabel: unclassifiedDebtSource.label,
        sourceXbrlTag: unclassifiedDebtSource.xbrlTag,
        currentModelRow: "",
        recommendedModelRow: "LT Debt (Incl. Current Portion)",
        reason: "Debt issuance costs are a long-term debt carrying-value adjustment.",
        reusableRule: "Classify debt carrying-value adjustments with long-term debt using statement context.",
        evidence: ["Primary balance sheet debt context"]
      }
    ],
    classifications: missingClassificationStore,
    balanceAssignments: [unclassifiedDebtAssignment],
    incomeAssignments: [],
    availableModelRows: ["LT Debt (Incl. Current Portion)"]
  });
  assert.equal(missingClassificationCorrection.changed, true);
  const createdClassification = missingClassificationStore.get(classificationSourceKeys(unclassifiedDebtSource)[0]);
  assert.equal(createdClassification.recommended_model_row, "LT Debt (Incl. Current Portion)");
  assert.equal(createdClassification.llm_used, true);
  assert.equal(createdClassification.mapping_passed_validation, true);

  const subtotalSource = {
    period: "3Q26",
    accession: "0000000000-26-000003",
    xbrlTag: "NonoperatingIncomeExpense",
    label: "Interest and other income (loss), net",
    amount: 79000000
  };
  const subtotalClassification = {
    ...classification,
    source_line_item: subtotalSource.label,
    recommended_action: "merge_into_other",
    recommended_model_row: "Other Non-Operating Income / Expense",
    classification_type: "combined non-operating line",
    is_current: null,
    is_operating: false,
    reason: "Initial whole-statement assignment."
  };
  const subtotalClassifications = new Map();
  classificationSourceKeys(subtotalSource).forEach((key) => subtotalClassifications.set(key, subtotalClassification));
  const subtotalExclusion = applyLlmMappingReviewCorrections({
    issues: [
      {
        severity: "error",
        issueType: "unsupported_exclusion",
        period: subtotalSource.period,
        sourceLineItemLabel: subtotalSource.label,
        sourceXbrlTag: subtotalSource.xbrlTag,
        currentModelRow: "Other Non-Operating Income / Expense",
        recommendedModelRow: "Other Non-Operating Income / Expense",
        reason: "The combined line is a subtotal of separately mapped interest income, interest expense, and other income components.",
        reusableRule: "Exclude a reported subtotal when its disclosed components are mapped separately.",
        evidence: ["Primary income statement and failed EBIT-to-pre-tax bridge"]
      }
    ],
    classifications: subtotalClassifications,
    balanceAssignments: [],
    incomeAssignments: [
      {
        fiscalPeriod: subtotalSource.period,
        sourceFilingAccession: subtotalSource.accession,
        sourceStatement: "Consolidated Statements of Operations",
        sourceLineItemLabel: subtotalSource.label,
        sourceAmount: subtotalSource.amount,
        sourceXbrlTag: subtotalSource.xbrlTag,
        assignedModelRow: "Other Non-Operating Income / Expense",
        sourceSection: "below operating income"
      }
    ],
    availableModelRows: ["Other Non-Operating Income / Expense"]
  });
  assert.equal(subtotalExclusion.changed, true, "the LLM reviewer must be able to exclude a double-counted subtotal");
  const excludedSubtotal = subtotalClassifications.get(classificationSourceKeys(subtotalSource)[0]);
  assert.equal(excludedSubtotal.recommended_action, "exclude");
  assert.equal(excludedSubtotal.is_subtotal, true);
  assert.equal(excludedSubtotal.llm_used, true);
  assert.match(excludedSubtotal.reason, /reviewer exclusion/i);

  console.log("LLM accounting controller workbook and fallback guards passed.");
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
