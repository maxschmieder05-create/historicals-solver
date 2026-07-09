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
  timeoutMs: number;
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
  provider?: { require_parameters?: boolean; sort?: string };
  response_format?: unknown;
  max_tokens?: number;
  max_completion_tokens?: number;
  temperature?: number;
  session_id?: string;
};

const OPENROUTER_MODELS_URL = "https://openrouter.ai/api/v1/models";
const modelCapabilitiesCache = new Map<string, Promise<OpenRouterModelCapabilities>>();

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
  const modelCandidates = uniqueModels([request.model, ...(request.fallbackModels ?? [])]);
  const attemptTelemetry: AccountingLlmTelemetry[] = [];
  let lastResult: AccountingLlmResult<T> | null = null;

  for (const model of modelCandidates) {
    const result = await requestAccountingJsonForModel({ ...request, model });
    attemptTelemetry.push(...(result.attemptTelemetry ?? [result.telemetry]));
    if (result.value) return { ...result, attemptTelemetry };
    lastResult = result;
    if (!llmResultEligibleForFallback(result)) break;
  }

  if (lastResult) return { ...lastResult, attemptTelemetry };
  return requestAccountingJsonForModel(request);
}

async function requestAccountingJsonForModel<T>(request: AccountingLlmRequest<T>): Promise<AccountingLlmResult<T>> {
  const startedAt = Date.now();
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
  const capabilities = await capabilitiesForModel(request.model, fetchImpl, !request.fetchImpl);
  const shape = buildOpenRouterRequestShape(request, capabilities);
  const primary = await postOpenRouterJson(request, shape, fetchImpl, startedAt, false, capabilities);
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

  if (request.repair?.enabled) {
    const repairShape = buildRepairRequestShape(request, capabilities, shape, parsed.error);
    const repair = await postOpenRouterJson(request, repairShape, fetchImpl, startedAt, true, capabilities);
    if (!repair.ok) {
      return {
        ...repair.result,
        telemetry: {
          ...repair.result.telemetry,
          repairAttempted: true,
          errorMessage: repair.result.telemetry.errorMessage || parsed.error
        }
      };
    }
    const repaired = parseAndValidateResponse(repair.text, request.validate);
    if (repaired.ok) {
      const status: AccountingLlmStatus = repaired.validation.validated ? "repaired" : "completed_unvalidated";
      return {
        status,
        value: repaired.validation.value,
        rawText: repair.text,
        telemetry: {
          ...repair.telemetry,
          status,
          completed: true,
          validated: repaired.validation.validated,
          affectedOutput: repaired.validation.affectedOutput ?? repaired.validation.validated,
          repairAttempted: true
        }
      };
    }
    const status: AccountingLlmStatus = repaired.needsHumanReview ? "needs_human_review" : "attempted_failed";
    return {
      status,
      rawText: repair.text,
      error: repaired.error,
      telemetry: {
        ...repair.telemetry,
        status,
        completed: true,
        validated: false,
        affectedOutput: false,
        repairAttempted: true,
        errorMessage: repaired.error
      }
    };
  }

  const status: AccountingLlmStatus = parsed.needsHumanReview ? "needs_human_review" : "attempted_failed";
  return {
    status,
    rawText: primary.text,
    error: parsed.error,
    telemetry: {
      ...primary.telemetry,
      status,
      completed: true,
      validated: false,
      affectedOutput: false,
      errorMessage: parsed.error
    }
  };
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
  if (httpStatus && [408, 409, 425, 429, 500, 502, 503, 504].includes(httpStatus)) return true;
  if (httpStatus === 403 && /key limit|quota|rate limit|billing|credits?|capacity|provider/i.test(error)) return true;
  return Boolean(
    !error ||
      /no endpoints|routing|provider|timed out|key limit|quota|rate limit|billing|credits?|capacity|temporarily unavailable|overloaded|did not include text output|json|schema|parse|validation|omitted target|unexpected source_row_key|duplicate source_row_key/i.test(
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
    failedAttempts: items.filter((item) => item.attempted && !item.completed).length,
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
    provider: responseFormat ? { require_parameters: true, sort: "throughput" } : { sort: "throughput" },
    response_format: responseFormat
  };

  if (capabilities.supportsMaxCompletionTokens && !capabilities.supportsMaxTokens) {
    shape.max_completion_tokens = request.maxTokens;
  } else {
    shape.max_tokens = request.maxTokens;
  }
  if (capabilities.supportsTemperature) shape.temperature = 0;
  return stripUndefined(shape);
}

function buildRepairRequestShape<T>(
  request: AccountingLlmRequest<T>,
  capabilities: OpenRouterModelCapabilities,
  originalShape: OpenRouterRequestShape,
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
        role: "user",
        content: `${repairInstruction}\n\nValidation error: ${validationError}`
      }
    ],
    response_format: capabilities.supportsResponseFormat ? { type: "json_object" } : originalShape.response_format,
    provider: capabilities.supportsResponseFormat ? { require_parameters: true, sort: "throughput" } : originalShape.provider
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
  try {
    const response = await Promise.race([
      fetchImpl(request.endpoint, {
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
      }),
      new Promise<Response>((_, reject) => {
        timeout = setTimeout(() => {
          controller.abort();
          reject(new Error(`OpenRouter ${request.purpose} timed out after ${request.timeoutMs}ms`));
        }, request.timeoutMs);
      })
    ]);
    const body = await response.json().catch(() => null);
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
    const message = error instanceof Error ? error.message : String(error);
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
  }
}

function parseAndValidateResponse<T>(
  text: string,
  validate: (value: unknown) => AccountingLlmValidationResult<T>
): { ok: true; validation: Extract<AccountingLlmValidationResult<T>, { ok: true }> } | { ok: false; error: string; needsHumanReview: boolean } {
  let parsed: unknown;
  try {
    parsed = JSON.parse(text);
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

async function capabilitiesForModel(model: string, fetchImpl: typeof fetch, allowProbe: boolean): Promise<OpenRouterModelCapabilities> {
  const fromStatic = staticCapabilitiesForModel(model);
  if (!allowProbe || fromStatic.source === "static" || process.env.OPENROUTER_MODEL_CAPABILITY_PROBE === "0") return fromStatic;
  if (!/^https:\/\/openrouter\.ai\//i.test(OPENROUTER_MODELS_URL)) return fromStatic;
  const cached = modelCapabilitiesCache.get(model);
  if (cached) return cached;
  const promise = fetchImpl(`${OPENROUTER_MODELS_URL}?supported_parameters=response_format`, {
    headers: { Accept: "application/json" }
  })
    .then(async (response) => {
      if (!response.ok) return fromStatic;
      const body = await response.json().catch(() => null);
      const found = Array.isArray(body?.data) ? body.data.find((item: any) => item?.id === model) : null;
      if (!found || !Array.isArray(found.supported_parameters)) return fromStatic;
      const params = new Set(found.supported_parameters.filter((item: unknown): item is string => typeof item === "string"));
      return {
        supportsResponseFormat: params.has("response_format"),
        supportsStructuredOutputs: params.has("structured_outputs"),
        supportsTemperature: params.has("temperature"),
        supportsMaxCompletionTokens: params.has("max_completion_tokens"),
        supportsMaxTokens: params.has("max_tokens"),
        source: "router" as const
      };
    })
    .catch(() => fromStatic);
  modelCapabilitiesCache.set(model, promise);
  return promise;
}

function staticCapabilitiesForModel(model: string): OpenRouterModelCapabilities {
  const normalized = model.toLowerCase();
  if (/^openai\/gpt-5(?:\.|$)/.test(normalized)) {
    return {
      supportsResponseFormat: true,
      supportsStructuredOutputs: true,
      supportsTemperature: false,
      supportsMaxCompletionTokens: true,
      supportsMaxTokens: true,
      source: "static"
    };
  }
  if (/^openai\/gpt-4o|^openai\/o[34]|^anthropic\/claude|^google\/gemini|^openrouter\/(?:auto|free)/.test(normalized)) {
    return {
      supportsResponseFormat: true,
      supportsStructuredOutputs: true,
      supportsTemperature: !/^openai\/o[34]/.test(normalized),
      supportsMaxCompletionTokens: /^openai\/o[34]/.test(normalized),
      supportsMaxTokens: true,
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

function stripUndefined<T extends Record<string, any>>(value: T): T {
  for (const key of Object.keys(value)) {
    if (value[key] === undefined) delete value[key];
  }
  return value;
}
