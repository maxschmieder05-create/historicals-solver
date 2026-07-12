export type AccountingLlmStatus =
  | "disabled"
  | "unavailable"
  | "attempted_failed"
  | "completed_unvalidated"
  | "completed_validated"
  | "repaired"
  | "needs_human_review";

export type AccountingLlmPurpose =
  | "line_item_classification"
  | "statement_line_item_classification"
  | "row_mapping"
  | "workbook_mapping_review";

export type AccountingLlmMessage = {
  role: "system" | "user" | "assistant";
  content: string;
};

export type AccountingLlmUsage = {
  promptTokens?: number;
  completionTokens?: number;
  totalTokens?: number;
  reasoningTokens?: number;
  cachedTokens?: number;
  cacheWriteTokens?: number;
  cost?: number;
  upstreamInferenceCost?: number | null;
};

export type AccountingLlmTelemetry = {
  purpose: AccountingLlmPurpose;
  status: AccountingLlmStatus;
  attempted: boolean;
  completed: boolean;
  validated: boolean;
  affectedOutput: boolean;
  repairAttempted: boolean;
  model: string;
  responseModel?: string;
  provider?: string;
  requestFormat?: "json_schema" | "json_object";
  capabilitySource?: string;
  endpoint: string;
  httpStatus?: number;
  statusText?: string;
  errorMessage?: string;
  routingError?: boolean;
  generationId?: string;
  usage?: AccountingLlmUsage;
  routerMetadata?: unknown;
  durationMs: number;
};

export type AccountingLlmValidationResult<T> =
  | {
      ok: true;
      value: T;
      validated: boolean;
      affectedOutput?: boolean;
    }
  | {
      ok: false;
      error: string;
      needsHumanReview?: boolean;
    };

export type AccountingLlmRequest<T> = {
  purpose: AccountingLlmPurpose;
  apiKey: string;
  endpoint: string;
  model: string;
  fallbackModels?: string[];
  siteUrl: string;
  appTitle: string;
  messages: AccountingLlmMessage[];
  jsonSchema: unknown;
  maxTokens: number;
  maxAttemptsPerModel?: number;
  maxTotalAttempts?: number;
  reasoningEffort?: "max" | "xhigh" | "high" | "medium" | "low" | "minimal" | "none";
  timeoutMs: number;
  deadlineAt?: number;
  signal?: AbortSignal;
  enabled?: boolean;
  sessionId?: string;
  fetchImpl?: typeof fetch;
  validate: (value: unknown) => AccountingLlmValidationResult<T>;
  repair?: {
    enabled: boolean;
    instruction?: string;
  };
};

export type AccountingLlmResult<T> = {
  status: AccountingLlmStatus;
  value?: T;
  rawText?: string;
  error?: string;
  telemetry: AccountingLlmTelemetry;
  attemptTelemetry?: AccountingLlmTelemetry[];
};

type OpenRouterModelCapabilities = {
  supportsResponseFormat: boolean;
  supportsStructuredOutputs: boolean;
  supportsTemperature: boolean;
  supportsMaxCompletionTokens: boolean;
  supportsMaxTokens: boolean;
  source: "static" | "router" | "unknown";
};

type OpenRouterRequestShape = {
  model: string;
  messages: AccountingLlmMessage[];
  provider?: {
    require_parameters?: boolean;
    sort?: string;
    order?: string[];
    max_price?: { prompt: number; completion: number };
  };
  response_format?: unknown;
  max_tokens?: number;
  max_completion_tokens?: number;
  temperature?: number;
  reasoning?: {
    effort: "max" | "xhigh" | "high" | "medium" | "low" | "minimal" | "none";
    exclude: boolean;
  };
  session_id?: string;
};

const OPENROUTER_MODELS_URL = "https://openrouter.ai/api/v1/models";
const OPENROUTER_MAX_PROMPT_PRICE_PER_MILLION = boundedPositivePrice(
  process.env.OPENROUTER_MAX_PROMPT_PRICE_PER_MILLION,
  0.5
);
const OPENROUTER_MAX_COMPLETION_PRICE_PER_MILLION = boundedPositivePrice(
  process.env.OPENROUTER_MAX_COMPLETION_PRICE_PER_MILLION,
  2.1
);
const OPENROUTER_PROVIDER_ORDER = (process.env.OPENROUTER_PROVIDER_ORDER || "")
  .split(",")
  .map((provider) => provider.trim())
  .filter(Boolean);
const DEFAULT_MODEL_CAPABILITY_CACHE_TTL_MS = 15 * 60_000;
const DEFAULT_MODEL_CAPABILITY_FAILURE_CACHE_TTL_MS = 10_000;
const MAX_MODEL_CAPABILITY_CACHE_TTL_MS = 24 * 60 * 60_000;

type ModelCapabilitiesCacheEntry = {
  promise: Promise<OpenRouterModelCapabilities>;
  // An in-flight probe never expires so concurrent workbook requests share it.
  // Once it settles, router results and fallback results receive separate TTLs.
  expiresAt: number;
};

const modelCapabilitiesCache = new Map<string, ModelCapabilitiesCacheEntry>();

export function emptyAccountingLlmTelemetry(input: {
  purpose: AccountingLlmPurpose;
  model: string;
  endpoint: string;
  status: AccountingLlmStatus;
  errorMessage?: string;
}): AccountingLlmTelemetry {
  return {
    purpose: input.purpose,
    status: input.status,
    attempted: false,
    completed: false,
    validated: false,
    affectedOutput: false,
    repairAttempted: false,
    model: input.model,
    endpoint: input.endpoint,
    errorMessage: input.errorMessage,
    durationMs: 0
  };
}

export async function requestAccountingJson<T>(request: AccountingLlmRequest<T>): Promise<AccountingLlmResult<T>> {
  if (request.signal?.aborted) return accountingLlmCancelledResult(request);
  const modelCandidates = uniqueModels([request.model, ...(request.fallbackModels ?? [])]);
  const attemptTelemetry: AccountingLlmTelemetry[] = [];
  let lastResult: AccountingLlmResult<T> | null = null;
  const maxAttemptsPerModel = Math.max(1, Math.min(4, Math.floor(request.maxAttemptsPerModel ?? 1)));
  const maxTotalAttempts = Number.isFinite(request.maxTotalAttempts)
    ? Math.max(0, Math.floor(request.maxTotalAttempts!))
    : Number.POSITIVE_INFINITY;
  const configuredTimeoutMs = Number.isFinite(request.timeoutMs) && request.timeoutMs > 0 ? request.timeoutMs : 1_000;
  const requestDeadlineAt = Math.min(
    Number.isFinite(request.deadlineAt) && (request.deadlineAt ?? 0) > 0 ? request.deadlineAt! : Number.POSITIVE_INFINITY,
    Date.now() + configuredTimeoutMs
  );

  for (let modelIndex = 0; modelIndex < modelCandidates.length; modelIndex += 1) {
    const model = modelCandidates[modelIndex];
    for (let attempt = 0; attempt < maxAttemptsPerModel; attempt += 1) {
      const remainingAttempts = maxTotalAttempts - attemptedTelemetryCount(attemptTelemetry);
      if (remainingAttempts <= 0) return lastResult ?? accountingLlmAttemptBudgetResult(request, attemptTelemetry);
      if (attempt > 0) await retryBackoff(attempt, configuredTimeoutMs, requestDeadlineAt, request.signal);
      if (request.signal?.aborted) {
        const cancelled = accountingLlmCancelledResult(request);
        return { ...cancelled, attemptTelemetry: attemptTelemetry.length ? attemptTelemetry : [cancelled.telemetry] };
      }
      const remainingMs = requestDeadlineAt - Date.now();
      if (remainingMs <= 0) return lastResult ? { ...lastResult, attemptTelemetry } : accountingLlmDeadlineResult(request, attemptTelemetry);
      // A timeout from the preferred provider must leave enough of the shared
      // request budget for a configured fallback. Without this reserve, the
      // first model can consume the entire deadline and the fallback is only
      // nominally configured.
      const hasLaterFallback = modelIndex < modelCandidates.length - 1;
      const modelAttemptMs = hasLaterFallback ? Math.max(1, Math.floor((remainingMs * 2) / 3)) : remainingMs;
      const result = await requestAccountingJsonForModel({
        ...request,
        model,
        maxTotalAttempts: remainingAttempts,
        deadlineAt: requestDeadlineAt,
        timeoutMs: Math.max(1, Math.min(configuredTimeoutMs, modelAttemptMs))
      });
      attemptTelemetry.push(...(result.attemptTelemetry ?? [result.telemetry]));
      if (accountingLlmResultHasValue(result)) return { ...result, attemptTelemetry };
      lastResult = result;
      if (request.signal?.aborted) return { ...result, attemptTelemetry };
      if (attemptedTelemetryCount(attemptTelemetry) >= maxTotalAttempts) return { ...result, attemptTelemetry };
      if (!llmResultEligibleForFallback(result)) break;
    }
    if (lastResult && !llmResultEligibleForFallback(lastResult)) break;
  }

  if (lastResult) return { ...lastResult, attemptTelemetry };
  if (maxTotalAttempts <= 0) return accountingLlmAttemptBudgetResult(request, attemptTelemetry);
  return requestAccountingJsonForModel(request);
}

function attemptedTelemetryCount(items: AccountingLlmTelemetry[]) {
  return items.filter((item) => item.attempted).length;
}

function accountingLlmResultHasValue<T>(result: AccountingLlmResult<T>) {
  return Object.prototype.hasOwnProperty.call(result, "value");
}

async function retryBackoff(attempt: number, timeoutMs: number, deadlineAt: number, signal?: AbortSignal) {
  const boundedTimeout = Number.isFinite(timeoutMs) && timeoutMs > 0 ? timeoutMs : 1_000;
  const exponentialMs = 50 * 2 ** Math.max(0, attempt - 1);
  const jitter = 0.75 + Math.random() * 0.5;
  const remainingMs = Math.max(0, deadlineAt - Date.now());
  const delayMs = Math.max(0, Math.min(1_000, Math.floor(boundedTimeout / 4), Math.round(exponentialMs * jitter), remainingMs));
  if (!delayMs) return;
  await new Promise<void>((resolve) => {
    const timeout = setTimeout(finish, delayMs);
    const onAbort = () => finish();
    function finish() {
      clearTimeout(timeout);
      signal?.removeEventListener("abort", onAbort);
      resolve();
    }
    if (signal?.aborted) finish();
    else signal?.addEventListener("abort", onAbort, { once: true });
  });
}

function accountingLlmDeadlineResult<T>(
  request: AccountingLlmRequest<T>,
  attemptTelemetry: AccountingLlmTelemetry[] = []
): AccountingLlmResult<T> {
  const error = `LLM ${request.purpose} exceeded its total ${request.timeoutMs}ms request budget`;
  const telemetry = emptyAccountingLlmTelemetry({
    purpose: request.purpose,
    model: request.model,
    endpoint: request.endpoint,
    status: "attempted_failed",
    errorMessage: error
  });
  return { status: "attempted_failed", error, telemetry, attemptTelemetry: attemptTelemetry.length ? attemptTelemetry : [telemetry] };
}

function accountingLlmAttemptBudgetResult<T>(
  request: AccountingLlmRequest<T>,
  attemptTelemetry: AccountingLlmTelemetry[] = []
): AccountingLlmResult<T> {
  const error = `LLM ${request.purpose} exhausted its bounded request-attempt budget`;
  const telemetry = emptyAccountingLlmTelemetry({
    purpose: request.purpose,
    model: request.model,
    endpoint: request.endpoint,
    status: "attempted_failed",
    errorMessage: error
  });
  return { status: "attempted_failed", error, telemetry, attemptTelemetry: attemptTelemetry.length ? attemptTelemetry : [telemetry] };
}

function accountingLlmCancelledResult<T>(request: AccountingLlmRequest<T>): AccountingLlmResult<T> {
  const error = `LLM ${request.purpose} cancelled because the workbook fill request was cancelled`;
  const telemetry = emptyAccountingLlmTelemetry({
    purpose: request.purpose,
    model: request.model,
    endpoint: request.endpoint,
    status: "attempted_failed",
    errorMessage: error
  });
  return { status: "attempted_failed", error, telemetry, attemptTelemetry: [telemetry] };
}

async function requestAccountingJsonForModel<T>(request: AccountingLlmRequest<T>): Promise<AccountingLlmResult<T>> {
  const startedAt = Date.now();
  if (request.signal?.aborted) return accountingLlmCancelledResult(request);
  if (request.enabled === false) {
    const telemetry = emptyAccountingLlmTelemetry({
      purpose: request.purpose,
      model: request.model,
      endpoint: request.endpoint,
      status: "disabled",
      errorMessage: "LLM accounting controller disabled"
    });
    return { status: "disabled", error: telemetry.errorMessage, telemetry };
  }
  if (!request.apiKey) {
    const telemetry = emptyAccountingLlmTelemetry({
      purpose: request.purpose,
      model: request.model,
      endpoint: request.endpoint,
      status: "unavailable",
      errorMessage: "OPENROUTER_API_KEY is not set"
    });
    return { status: "unavailable", error: telemetry.errorMessage, telemetry };
  }

  const fetchImpl = request.fetchImpl ?? fetch;
  const capabilityTimeoutMs = accountingLlmRemainingMs(request);
  if (capabilityTimeoutMs <= 0) return accountingLlmDeadlineResult(request);
  const capabilities = await capabilitiesForModel(
    request.model,
    fetchImpl,
    !request.fetchImpl && isOpenRouterUrl(request.endpoint),
    request.apiKey,
    capabilityTimeoutMs
  );
  if (request.signal?.aborted) return accountingLlmCancelledResult(request);
  const shape = buildOpenRouterRequestShape(request, capabilities);
  const primaryTimeoutMs = accountingLlmRemainingMs(request);
  if (primaryTimeoutMs <= 0) return accountingLlmDeadlineResult(request);
  const primary = await postOpenRouterJson({ ...request, timeoutMs: primaryTimeoutMs }, shape, fetchImpl, startedAt, false, capabilities);
  if (!primary.ok) return primary.result;

  const parsed = parseAndValidateResponse(primary.text, request.validate);
  if (parsed.ok) {
    const status: AccountingLlmStatus = parsed.validation.validated ? "completed_validated" : "completed_unvalidated";
    return {
      status,
      value: parsed.validation.value,
      rawText: primary.text,
      telemetry: {
        ...primary.telemetry,
        status,
        completed: true,
        validated: parsed.validation.validated,
        affectedOutput: parsed.validation.affectedOutput ?? parsed.validation.validated
      }
    };
  }

  const primaryFailureStatus: AccountingLlmStatus = parsed.needsHumanReview ? "needs_human_review" : "attempted_failed";
  const primaryFailureTelemetry: AccountingLlmTelemetry = {
    ...primary.telemetry,
    status: primaryFailureStatus,
    completed: true,
    validated: false,
    affectedOutput: false,
    errorMessage: parsed.error
  };

  const repairFitsAttemptBudget =
    !Number.isFinite(request.maxTotalAttempts) || Math.max(0, Math.floor(request.maxTotalAttempts!)) >= 2;
  if (request.repair?.enabled && repairFitsAttemptBudget) {
    const repairShape = buildRepairRequestShape(request, shape, primary.text, parsed.error);
    const repairTimeoutMs = accountingLlmRemainingMs(request);
    if (repairTimeoutMs <= 0) {
      return {
        status: primaryFailureStatus,
        rawText: primary.text,
        error: parsed.error,
        telemetry: primaryFailureTelemetry,
        attemptTelemetry: [primaryFailureTelemetry]
      };
    }
    const repair = await postOpenRouterJson({ ...request, timeoutMs: repairTimeoutMs }, repairShape, fetchImpl, Date.now(), true, capabilities);
    if (!repair.ok) {
      const repairTelemetry = {
        ...repair.result.telemetry,
        repairAttempted: true,
        errorMessage: repair.result.telemetry.errorMessage || parsed.error
      };
      return {
        ...repair.result,
        telemetry: repairTelemetry,
        attemptTelemetry: [primaryFailureTelemetry, repairTelemetry]
      };
    }
    const repaired = parseAndValidateResponse(repair.text, request.validate);
    if (repaired.ok) {
      const status: AccountingLlmStatus = repaired.validation.validated ? "repaired" : "completed_unvalidated";
      const repairTelemetry: AccountingLlmTelemetry = {
        ...repair.telemetry,
        status,
        completed: true,
        validated: repaired.validation.validated,
        affectedOutput: repaired.validation.affectedOutput ?? repaired.validation.validated,
        repairAttempted: true
      };
      return {
        status,
        value: repaired.validation.value,
        rawText: repair.text,
        telemetry: repairTelemetry,
        attemptTelemetry: [primaryFailureTelemetry, repairTelemetry]
      };
    }
    const status: AccountingLlmStatus = repaired.needsHumanReview ? "needs_human_review" : "attempted_failed";
    const repairTelemetry: AccountingLlmTelemetry = {
      ...repair.telemetry,
      status,
      completed: true,
      validated: false,
      affectedOutput: false,
      repairAttempted: true,
      errorMessage: repaired.error
    };
    return {
      status,
      rawText: repair.text,
      error: repaired.error,
      telemetry: repairTelemetry,
      attemptTelemetry: [primaryFailureTelemetry, repairTelemetry]
    };
  }

  return {
    status: primaryFailureStatus,
    rawText: primary.text,
    error: parsed.error,
    telemetry: primaryFailureTelemetry
  };
}

function accountingLlmRemainingMs(request: { timeoutMs: number; deadlineAt?: number }) {
  const timeoutMs = Number.isFinite(request.timeoutMs) && request.timeoutMs > 0 ? request.timeoutMs : 1_000;
  if (!request.deadlineAt || !Number.isFinite(request.deadlineAt)) return timeoutMs;
  return Math.max(0, Math.min(timeoutMs, request.deadlineAt - Date.now()));
}

function uniqueModels(models: string[]) {
  const seen = new Set<string>();
  return models
    .map((model) => model.trim())
    .filter(Boolean)
    .filter((model) => {
      const key = model.toLowerCase();
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
}

function llmResultEligibleForFallback<T>(result: AccountingLlmResult<T>) {
  if (result.status !== "attempted_failed" && result.status !== "needs_human_review") return false;
  const error = result.error || result.telemetry.errorMessage || "";
  const httpStatus = result.telemetry.httpStatus;
  if ((!httpStatus || [401, 403].includes(httpStatus)) && /key limit|quota|billing|credits?|auth|api key|forbidden/i.test(error)) return false;
  if (httpStatus && [408, 409, 425, 429, 500, 502, 503, 504].includes(httpStatus)) return true;
  if (httpStatus && httpStatus >= 400 && httpStatus < 500) return false;
  return Boolean(
    !error ||
      /no endpoints|routing|provider|timed out|fetch failed|network error|socket(?: hang up)?|econnreset|econnrefused|enotfound|eai_again|connection (?:reset|closed|terminated)|other side closed|capacity|temporarily unavailable|overloaded|did not include text output|json|schema|parse|validation|omitted target|unexpected source_row_key|duplicate source_row_key/i.test(
        error
      )
  );
}

export function aggregateAccountingTelemetry(items: AccountingLlmTelemetry[]) {
  const statusCounts = items.reduce<Record<AccountingLlmStatus, number>>(
    (counts, item) => {
      counts[item.status] += 1;
      return counts;
    },
    {
      disabled: 0,
      unavailable: 0,
      attempted_failed: 0,
      completed_unvalidated: 0,
      completed_validated: 0,
      repaired: 0,
      needs_human_review: 0
    }
  );
  const usage = items.reduce<AccountingLlmUsage>(
    (total, item) => addUsage(total, item.usage),
    {}
  );
  return {
    attempts: items.filter((item) => item.attempted).length,
    successfulCompletions: items.filter((item) => item.completed).length,
    validatedCompletions: items.filter((item) => item.validated).length,
    failedAttempts: items.filter((item) => item.attempted && (!item.completed || !item.validated)).length,
    affectedOutputDecisions: items.filter((item) => item.affectedOutput).length,
    statuses: statusCounts,
    models: Array.from(new Set(items.map((item) => item.responseModel || item.model).filter(Boolean))),
    usage
  };
}

export function sanitizeAccountingTelemetry(item: AccountingLlmTelemetry) {
  return {
    purpose: item.purpose,
    status: item.status,
    attempted: item.attempted,
    completed: item.completed,
    validated: item.validated,
    affectedOutput: item.affectedOutput,
    repairAttempted: item.repairAttempted,
    model: item.model,
    responseModel: item.responseModel,
    provider: item.provider,
    requestFormat: item.requestFormat,
    capabilitySource: item.capabilitySource,
    httpStatus: item.httpStatus,
    statusText: item.statusText,
    errorMessage: item.errorMessage,
    routingError: item.routingError,
    generationId: item.generationId,
    usage: item.usage,
    durationMs: item.durationMs
  };
}

function buildOpenRouterRequestShape<T>(
  request: AccountingLlmRequest<T>,
  capabilities: OpenRouterModelCapabilities
): OpenRouterRequestShape {
  const responseFormat = capabilities.supportsStructuredOutputs
    ? {
        type: "json_schema",
        json_schema: request.jsonSchema
      }
    : capabilities.supportsResponseFormat
      ? { type: "json_object" }
      : undefined;
  const shape: OpenRouterRequestShape = {
    model: request.model,
    messages: request.messages,
    session_id: request.sessionId,
    provider: {
      ...(responseFormat ? { require_parameters: true } : {}),
      ...(OPENROUTER_PROVIDER_ORDER.length ? { order: OPENROUTER_PROVIDER_ORDER } : {}),
      // Avoid `sort`, which disables OpenRouter's health-aware provider load
      // balancing. An explicit order remains available as an opt-in override.
      max_price: {
        prompt: OPENROUTER_MAX_PROMPT_PRICE_PER_MILLION,
        completion: OPENROUTER_MAX_COMPLETION_PRICE_PER_MILLION
      }
    },
    response_format: responseFormat,
    reasoning: request.reasoningEffort ? { effort: request.reasoningEffort, exclude: true } : undefined
  };

  if (capabilities.supportsMaxCompletionTokens) {
    shape.max_completion_tokens = request.maxTokens;
  } else if (capabilities.supportsMaxTokens) {
    shape.max_tokens = request.maxTokens;
  }
  if (capabilities.supportsTemperature) shape.temperature = 0;
  return stripUndefined(shape);
}

function buildRepairRequestShape<T>(
  request: AccountingLlmRequest<T>,
  originalShape: OpenRouterRequestShape,
  originalResponse: string,
  validationError: string
): OpenRouterRequestShape {
  const repairInstruction =
    request.repair?.instruction ||
    "Repair the previous response so it is valid JSON matching the requested schema. Return only the corrected JSON object.";
  return {
    ...originalShape,
    messages: [
      ...request.messages,
      {
        role: "assistant",
        content: originalResponse
      },
      {
        role: "user",
        content: `${repairInstruction}\n\nValidation error: ${validationError}`
      }
    ]
  };
}

async function postOpenRouterJson<T>(
  request: AccountingLlmRequest<T>,
  shape: OpenRouterRequestShape,
  fetchImpl: typeof fetch,
  startedAt: number,
  repairAttempted: boolean,
  capabilities: OpenRouterModelCapabilities
): Promise<
  | { ok: true; text: string; telemetry: AccountingLlmTelemetry }
  | { ok: false; result: AccountingLlmResult<T> }
> {
  const controller = new AbortController();
  let timeout: ReturnType<typeof setTimeout> | undefined;
  let onExternalAbort: (() => void) | undefined;
  try {
    const { response, body } = await Promise.race([
      (async () => {
        const response = await fetchImpl(request.endpoint, {
          method: "POST",
          signal: controller.signal,
          headers: {
            "Content-Type": "application/json",
            Authorization: `Bearer ${request.apiKey}`,
            "HTTP-Referer": request.siteUrl,
            "X-OpenRouter-Title": request.appTitle,
            "X-Title": request.appTitle,
            "X-OpenRouter-Metadata": "enabled",
            ...(request.sessionId ? { "x-session-id": request.sessionId } : {})
          },
          body: JSON.stringify(shape)
        });
        const body = await response.json().catch(() => null);
        return { response, body };
      })(),
      new Promise<never>((_, reject) => {
        timeout = setTimeout(() => {
          controller.abort();
          reject(new Error(`OpenRouter ${request.purpose} timed out after ${request.timeoutMs}ms`));
        }, request.timeoutMs);
      }),
      new Promise<never>((_, reject) => {
        if (!request.signal) return;
        onExternalAbort = () => {
          controller.abort();
          reject(new Error(`LLM ${request.purpose} cancelled because the workbook fill request was cancelled`));
        };
        if (request.signal.aborted) onExternalAbort();
        else request.signal.addEventListener("abort", onExternalAbort, { once: true });
      })
    ]);
    const telemetry = telemetryFromOpenRouterResponse(
      request,
      shape,
      response,
      body,
      Date.now() - startedAt,
      repairAttempted,
      capabilities
    );
    if (!response.ok) {
      const message = openRouterErrorMessage(body, response);
      return {
        ok: false,
        result: {
          status: "attempted_failed",
          error: message,
          telemetry: {
            ...telemetry,
            status: "attempted_failed",
            completed: false,
            validated: false,
            affectedOutput: false,
            errorMessage: message,
            routingError: response.status === 404 || /no endpoints|routing|provider/i.test(message)
          }
        }
      };
    }

    const text = responseOutputText(body);
    if (!text) {
      const message = "OpenRouter response did not include text output";
      return {
        ok: false,
        result: {
          status: "attempted_failed",
          error: message,
          telemetry: {
            ...telemetry,
            status: "attempted_failed",
            completed: false,
            validated: false,
            affectedOutput: false,
            errorMessage: message
          }
        }
      };
    }
    return { ok: true, text, telemetry };
  } catch (error) {
    const message = request.signal?.aborted
      ? `LLM ${request.purpose} cancelled because the workbook fill request was cancelled`
      : error instanceof Error
        ? error.message
        : String(error);
    return {
      ok: false,
      result: {
        status: "attempted_failed",
        error: message,
        telemetry: {
          purpose: request.purpose,
          status: "attempted_failed",
          attempted: true,
          completed: false,
          validated: false,
          affectedOutput: false,
          repairAttempted,
          model: request.model,
          endpoint: request.endpoint,
          requestFormat: responseFormatType(shape),
          capabilitySource: capabilities.source,
          errorMessage: message,
          routingError: /no endpoints|routing|provider/i.test(message),
          durationMs: Date.now() - startedAt
        }
      }
    };
  } finally {
    if (timeout) clearTimeout(timeout);
    if (onExternalAbort) request.signal?.removeEventListener("abort", onExternalAbort);
  }
}

function parseAndValidateResponse<T>(
  text: string,
  validate: (value: unknown) => AccountingLlmValidationResult<T>
): { ok: true; validation: Extract<AccountingLlmValidationResult<T>, { ok: true }> } | { ok: false; error: string; needsHumanReview: boolean } {
  let parsed: unknown;
  try {
    parsed = parseLlmJson(text);
  } catch (error) {
    return {
      ok: false,
      error: error instanceof Error ? error.message : "LLM JSON parse failed",
      needsHumanReview: true
    };
  }
  const validation = validate(parsed);
  if (validation.ok) return { ok: true, validation };
  return { ok: false, error: validation.error, needsHumanReview: validation.needsHumanReview !== false };
}

function parseLlmJson(text: string) {
  const trimmed = text.trim();
  try {
    return JSON.parse(trimmed);
  } catch {
    // Some compatible providers wrap otherwise-valid structured output in a
    // markdown fence despite response_format. Strip only the wrapper and keep
    // schema validation as the authority for the parsed value.
  }
  const unfenced = trimmed
    .replace(/^```(?:json)?\s*/i, "")
    .replace(/\s*```$/i, "")
    .trim();
  try {
    return JSON.parse(unfenced);
  } catch {
    const objectStart = unfenced.indexOf("{");
    const arrayStart = unfenced.indexOf("[");
    const start = [objectStart, arrayStart].filter((index) => index >= 0).sort((a, b) => a - b)[0];
    if (start === undefined) throw new Error("LLM response did not contain a JSON object or array");
    const open = unfenced[start];
    const close = open === "[" ? "]" : "}";
    const end = unfenced.lastIndexOf(close);
    if (end <= start) throw new Error("LLM response contained incomplete JSON");
    return JSON.parse(unfenced.slice(start, end + 1));
  }
}

function telemetryFromOpenRouterResponse<T>(
  request: AccountingLlmRequest<T>,
  shape: OpenRouterRequestShape,
  response: Response,
  body: any,
  durationMs: number,
  repairAttempted: boolean,
  capabilities: OpenRouterModelCapabilities
): AccountingLlmTelemetry {
  const generationId = response.headers?.get?.("X-Generation-Id") || body?.id || undefined;
  const metadata = body?.openrouter_metadata;
  return {
    purpose: request.purpose,
    status: "attempted_failed",
    attempted: true,
    completed: false,
    validated: false,
    affectedOutput: false,
    repairAttempted,
    model: request.model,
    responseModel: typeof body?.model === "string" ? body.model : undefined,
    provider: providerFromOpenRouterMetadata(metadata),
    requestFormat: responseFormatType(shape),
    capabilitySource: capabilities.source,
    endpoint: request.endpoint,
    httpStatus: response.status,
    statusText: response.statusText,
    generationId,
    usage: usageFromOpenRouterBody(body),
    routerMetadata: metadata,
    durationMs
  };
}

async function capabilitiesForModel(
  model: string,
  fetchImpl: typeof fetch,
  allowProbe: boolean,
  apiKey: string,
  timeoutMs: number
): Promise<OpenRouterModelCapabilities> {
  const fromStatic = staticCapabilitiesForModel(model);
  if (!allowProbe || fromStatic.source === "static" || process.env.OPENROUTER_MODEL_CAPABILITY_PROBE === "0") return fromStatic;
  if (!isOpenRouterUrl(OPENROUTER_MODELS_URL)) return fromStatic;
  const now = Date.now();
  const cached = modelCapabilitiesCache.get(model);
  if (cached && cached.expiresAt > now) return cached.promise;
  if (cached) modelCapabilitiesCache.delete(model);

  const promise = probeModelCapabilities(model, fetchImpl, apiKey, timeoutMs, fromStatic);
  const entry: ModelCapabilitiesCacheEntry = {
    promise,
    expiresAt: Number.POSITIVE_INFINITY
  };
  modelCapabilitiesCache.set(model, entry);
  void promise.then(
    (capabilities) => {
      // Do not mutate a newer entry if this probe was evicted/replaced while it
      // was settling. A router result is safe to reuse longer; a timeout,
      // malformed response, missing model, or other fallback is retried soon.
      if (modelCapabilitiesCache.get(model) !== entry) return;
      entry.expiresAt =
        Date.now() +
        modelCapabilityCacheTtlMs(
          capabilities.source === "router"
            ? "OPENROUTER_MODEL_CAPABILITY_CACHE_TTL_MS"
            : "OPENROUTER_MODEL_CAPABILITY_FAILURE_CACHE_TTL_MS",
          capabilities.source === "router"
            ? DEFAULT_MODEL_CAPABILITY_CACHE_TTL_MS
            : DEFAULT_MODEL_CAPABILITY_FAILURE_CACHE_TTL_MS
        );
    },
    () => {
      // probeModelCapabilities currently resolves to its fallback on failure,
      // but keep unexpected rejections retryable rather than poisoning cache.
      if (modelCapabilitiesCache.get(model) === entry) modelCapabilitiesCache.delete(model);
    }
  );
  return promise;
}

function modelCapabilityCacheTtlMs(environmentKey: string, fallbackMs: number) {
  const configured = Number(process.env[environmentKey]);
  if (!Number.isFinite(configured) || configured < 0) return fallbackMs;
  return Math.min(MAX_MODEL_CAPABILITY_CACHE_TTL_MS, Math.floor(configured));
}

async function probeModelCapabilities(
  model: string,
  fetchImpl: typeof fetch,
  apiKey: string,
  timeoutMs: number,
  fallback: OpenRouterModelCapabilities
): Promise<OpenRouterModelCapabilities> {
  if (!isOpenRouterUrl(OPENROUTER_MODELS_URL)) return fallback;
  const controller = new AbortController();
  const boundedRequestMs = Number.isFinite(timeoutMs) && timeoutMs > 0 ? timeoutMs : 2_000;
  const probeTimeoutMs = Math.max(1, Math.min(1_500, Math.floor(boundedRequestMs / 4) || boundedRequestMs));
  let timeout: ReturnType<typeof setTimeout> | undefined;
  try {
    const body = await Promise.race([
      (async () => {
        const response = await fetchImpl(`${OPENROUTER_MODELS_URL}?supported_parameters=response_format`, {
          signal: controller.signal,
          headers: {
            Accept: "application/json",
            ...(apiKey ? { Authorization: `Bearer ${apiKey}` } : {})
          }
        });
        if (!response.ok) return null;
        return response.json().catch(() => null);
      })(),
      new Promise<null>((resolve) => {
        timeout = setTimeout(() => {
          controller.abort();
          resolve(null);
        }, probeTimeoutMs);
      })
    ]);
    const found = Array.isArray(body?.data) ? body.data.find((item: any) => item?.id === model) : null;
    if (!found || !Array.isArray(found.supported_parameters)) return fallback;
    const params = new Set(found.supported_parameters.filter((item: unknown): item is string => typeof item === "string"));
    return {
      supportsResponseFormat: params.has("response_format"),
      supportsStructuredOutputs: params.has("structured_outputs"),
      supportsTemperature: params.has("temperature"),
      supportsMaxCompletionTokens: params.has("max_completion_tokens"),
      supportsMaxTokens: params.has("max_tokens"),
      source: "router"
    };
  } catch {
    return fallback;
  } finally {
    if (timeout) clearTimeout(timeout);
  }
}

function isOpenRouterUrl(value: string) {
  try {
    const url = new URL(value);
    return url.protocol === "https:" && (url.hostname === "openrouter.ai" || url.hostname.endsWith(".openrouter.ai"));
  } catch {
    return false;
  }
}

function staticCapabilitiesForModel(model: string): OpenRouterModelCapabilities {
  const normalized = model.toLowerCase();
  if (/^openai\/gpt-5(?:\.|$)/.test(normalized)) {
    return {
      supportsResponseFormat: true,
      supportsStructuredOutputs: true,
      supportsTemperature: false,
      supportsMaxCompletionTokens: true,
      supportsMaxTokens: false,
      source: "static"
    };
  }
  if (/^openai\/gpt-4o|^openai\/o[34]|^anthropic\/claude|^google\/gemini|^openrouter\/(?:auto|free)/.test(normalized)) {
    return {
      supportsResponseFormat: true,
      supportsStructuredOutputs: true,
      supportsTemperature: !/^openai\/o[34]/.test(normalized),
      supportsMaxCompletionTokens: /^openai\/o[34]/.test(normalized),
      supportsMaxTokens: !/^openai\/o[34]/.test(normalized),
      source: "static"
    };
  }
  if (/^(deepseek|qwen|moonshotai|z-ai)\//.test(normalized)) {
    return {
      supportsResponseFormat: true,
      supportsStructuredOutputs: true,
      supportsTemperature: true,
      supportsMaxCompletionTokens: false,
      supportsMaxTokens: true,
      source: "static"
    };
  }
  return {
    supportsResponseFormat: true,
    supportsStructuredOutputs: false,
    supportsTemperature: false,
    supportsMaxCompletionTokens: false,
    supportsMaxTokens: true,
    source: "unknown"
  };
}

function responseOutputText(body: any) {
  if (typeof body?.output_text === "string") return body.output_text;
  const content = body?.choices?.[0]?.message?.content;
  if (typeof content === "string") return content.trim();
  if (content && typeof content === "object" && !Array.isArray(content)) return JSON.stringify(content);
  if (Array.isArray(content)) {
    return content
      .map((item) => (typeof item?.text === "string" ? item.text : typeof item?.content === "string" ? item.content : ""))
      .join("")
      .trim();
  }
  const reasoning = body?.choices?.[0]?.message?.reasoning;
  if (typeof reasoning === "string" && /[\[{]/.test(reasoning)) return reasoning.trim();
  const chunks: string[] = [];
  for (const item of body?.output ?? []) {
    if (item?.type !== "message") continue;
    for (const contentItem of item.content ?? []) {
      if (contentItem?.type === "output_text" && typeof contentItem.text === "string") chunks.push(contentItem.text);
    }
  }
  return chunks.join("").trim();
}

function usageFromOpenRouterBody(body: any): AccountingLlmUsage | undefined {
  const usage = body?.usage;
  if (!usage || typeof usage !== "object") return undefined;
  return stripUndefined({
    promptTokens: numeric(usage.prompt_tokens),
    completionTokens: numeric(usage.completion_tokens),
    totalTokens: numeric(usage.total_tokens),
    reasoningTokens: numeric(usage.completion_tokens_details?.reasoning_tokens),
    cachedTokens: numeric(usage.prompt_tokens_details?.cached_tokens),
    cacheWriteTokens: numeric(usage.prompt_tokens_details?.cache_write_tokens),
    cost: numeric(usage.cost),
    upstreamInferenceCost: numericOrNull(usage.cost_details?.upstream_inference_cost)
  });
}

function addUsage(left: AccountingLlmUsage, right: AccountingLlmUsage | undefined): AccountingLlmUsage {
  if (!right) return left;
  return stripUndefined({
    promptTokens: sumOptional(left.promptTokens, right.promptTokens),
    completionTokens: sumOptional(left.completionTokens, right.completionTokens),
    totalTokens: sumOptional(left.totalTokens, right.totalTokens),
    reasoningTokens: sumOptional(left.reasoningTokens, right.reasoningTokens),
    cachedTokens: sumOptional(left.cachedTokens, right.cachedTokens),
    cacheWriteTokens: sumOptional(left.cacheWriteTokens, right.cacheWriteTokens),
    cost: sumOptional(left.cost, right.cost),
    upstreamInferenceCost: sumOptional(left.upstreamInferenceCost ?? undefined, right.upstreamInferenceCost ?? undefined)
  });
}

function responseFormatType(shape: OpenRouterRequestShape): "json_schema" | "json_object" | undefined {
  const value = shape.response_format;
  if (!value || typeof value !== "object") return undefined;
  const type = (value as { type?: unknown }).type;
  return type === "json_schema" || type === "json_object" ? type : undefined;
}

function openRouterErrorMessage(body: any, response: Response) {
  return (
    body?.error?.message ||
    body?.message ||
    (typeof body?.error === "string" ? body.error : "") ||
    `${response.status} ${response.statusText}`.trim()
  );
}

function providerFromOpenRouterMetadata(metadata: any) {
  if (!metadata || typeof metadata !== "object") return undefined;
  const candidates = [
    metadata.provider_name,
    metadata.provider,
    metadata.provider_slug,
    metadata.route?.provider_name,
    metadata.route?.provider,
    metadata.final_provider,
    metadata.final_provider_name
  ];
  return candidates.find((item): item is string => typeof item === "string" && item.length > 0);
}

function numeric(value: unknown) {
  return typeof value === "number" && Number.isFinite(value) ? value : undefined;
}

function numericOrNull(value: unknown) {
  if (value === null) return null;
  return numeric(value);
}

function sumOptional(a: number | undefined, b: number | undefined) {
  if (a === undefined && b === undefined) return undefined;
  return (a ?? 0) + (b ?? 0);
}

function boundedPositivePrice(value: string | undefined, fallback: number) {
  const parsed = Number(value ?? fallback);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : fallback;
}

function stripUndefined<T extends Record<string, any>>(value: T): T {
  for (const key of Object.keys(value)) {
    if (value[key] === undefined) delete value[key];
  }
  return value;
}

export const __llmAccountingControllerTestHooks = {
  capabilitiesForModel,
  clearModelCapabilitiesCache: () => modelCapabilitiesCache.clear(),
  modelCapabilitiesCacheState: () =>
    Array.from(modelCapabilitiesCache.entries()).map(([model, entry]) => ({
      model,
      expiresAt: entry.expiresAt
    })),
  probeModelCapabilities,
  staticCapabilitiesForModel
};
