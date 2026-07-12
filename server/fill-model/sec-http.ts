type SecFetchParser<T> = (response: Response) => Promise<T>;

export type SecFetchOptions = {
  timeoutMs?: number;
  maxAttempts?: number;
  retryBaseDelayMs?: number;
  retryMaxDelayMs?: number;
  fetchImpl?: typeof fetch;
  env?: NodeJS.ProcessEnv;
};

export type SecFetchResult<T> = {
  response: Response;
  data: T | null;
};

export type SecCompanyDirectoryEntry = {
  cik: string;
  ticker: string;
  title: string;
};

const DEFAULT_SEC_FETCH_TIMEOUT_MS = 20_000;
const DEFAULT_SEC_FETCH_MAX_ATTEMPTS = 3;
const DEFAULT_SEC_RETRY_BASE_DELAY_MS = 300;
const DEFAULT_SEC_RETRY_MAX_DELAY_MS = 5_000;
const DEVELOPMENT_SEC_USER_AGENT = "HistoricalsSolver/0.1 local-development";
const SEC_USER_AGENT_CONFIGURATION_MESSAGE =
  "SEC_USER_AGENT must identify the application and include a real, monitored operator email address.";

export class SecFetchTimeoutError extends Error {
  readonly url: string;
  readonly timeoutMs: number;

  constructor(url: string, timeoutMs: number) {
    super(`SEC request timed out after ${timeoutMs} ms: ${url}`);
    this.name = "SecFetchTimeoutError";
    this.url = url;
    this.timeoutMs = timeoutMs;
  }
}

export class SecUserAgentConfigurationError extends Error {
  constructor(message = SEC_USER_AGENT_CONFIGURATION_MESSAGE) {
    super(message);
    this.name = "SecUserAgentConfigurationError";
  }
}

export function resolveSecUserAgent(env: NodeJS.ProcessEnv = process.env) {
  const configured = env.SEC_USER_AGENT?.trim() ?? "";
  if (configured) {
    if (!validConfiguredSecUserAgent(configured)) {
      throw new SecUserAgentConfigurationError(
        `${SEC_USER_AGENT_CONFIGURATION_MESSAGE} Placeholder and reserved example addresses are not accepted.`
      );
    }
    return configured;
  }
  if (env.NODE_ENV === "production") {
    throw new SecUserAgentConfigurationError(
      `${SEC_USER_AGENT_CONFIGURATION_MESSAGE} It is required before production EDGAR requests are allowed.`
    );
  }
  return DEVELOPMENT_SEC_USER_AGENT;
}

export function secRequestHeaders(
  headers: HeadersInit | undefined,
  accept: string | undefined,
  env: NodeJS.ProcessEnv = process.env
) {
  const result = new Headers(headers);
  const suppliedUserAgent = result.get("User-Agent")?.trim();
  const userAgent =
    suppliedUserAgent === DEVELOPMENT_SEC_USER_AGENT && env.NODE_ENV !== "production" && !env.SEC_USER_AGENT
      ? DEVELOPMENT_SEC_USER_AGENT
      : suppliedUserAgent
        ? resolveSecUserAgent({ ...env, SEC_USER_AGENT: suppliedUserAgent })
        : resolveSecUserAgent(env);
  result.set("User-Agent", userAgent);
  if (accept) result.set("Accept", accept);
  return Object.fromEntries(result.entries());
}

export async function fetchSecJson<T = unknown>(
  url: string,
  init: RequestInit = {},
  options: SecFetchOptions = {}
): Promise<SecFetchResult<T>> {
  return fetchSecParsed<T>(url, init, options, async (response) => response.json() as Promise<T>);
}

export async function fetchSecText(
  url: string,
  init: RequestInit = {},
  options: SecFetchOptions = {}
): Promise<SecFetchResult<string>> {
  return fetchSecParsed<string>(url, init, options, (response) => response.text());
}

async function fetchSecParsed<T>(
  url: string,
  init: RequestInit,
  options: SecFetchOptions,
  parse: SecFetchParser<T>
): Promise<SecFetchResult<T>> {
  const env = options.env ?? process.env;
  const timeoutMs = boundedPositiveInteger(
    options.timeoutMs ?? env.SEC_FETCH_TIMEOUT_MS,
    DEFAULT_SEC_FETCH_TIMEOUT_MS,
    1,
    120_000
  );
  const maxAttempts = boundedPositiveInteger(
    options.maxAttempts ?? env.SEC_FETCH_MAX_ATTEMPTS,
    DEFAULT_SEC_FETCH_MAX_ATTEMPTS,
    1,
    5
  );
  const retryBaseDelayMs = boundedPositiveInteger(
    options.retryBaseDelayMs ?? env.SEC_FETCH_RETRY_BASE_DELAY_MS,
    DEFAULT_SEC_RETRY_BASE_DELAY_MS,
    1,
    10_000
  );
  const retryMaxDelayMs = boundedPositiveInteger(
    options.retryMaxDelayMs ?? env.SEC_FETCH_RETRY_MAX_DELAY_MS,
    DEFAULT_SEC_RETRY_MAX_DELAY_MS,
    retryBaseDelayMs,
    30_000
  );
  const fetchImpl = options.fetchImpl ?? fetch;
  const callerSignal = init.signal ?? undefined;
  const headers = secRequestHeaders(init.headers, undefined, env);
  let lastError: unknown;

  for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
    throwIfCallerAborted(callerSignal);
    const controller = new AbortController();
    const timeoutError = new SecFetchTimeoutError(url, timeoutMs);
    let timeoutHandle: ReturnType<typeof setTimeout> | undefined;
    let rejectCallerAbort: ((reason: unknown) => void) | undefined;
    const timeoutGate = new Promise<never>((_resolve, reject) => {
      timeoutHandle = setTimeout(() => {
        controller.abort(timeoutError);
        reject(timeoutError);
      }, timeoutMs);
    });
    const callerAbortGate = new Promise<never>((_resolve, reject) => {
      rejectCallerAbort = reject;
    });
    const onCallerAbort = () => {
      const reason = callerAbortReason(callerSignal);
      controller.abort(reason);
      rejectCallerAbort?.(reason);
    };
    callerSignal?.addEventListener("abort", onCallerAbort, { once: true });
    if (callerSignal?.aborted) onCallerAbort();

    try {
      const response = await Promise.race([
        fetchImpl(url, { ...init, headers, signal: controller.signal }),
        timeoutGate,
        callerAbortGate
      ]);
      if (retryableSecStatus(response.status) && attempt < maxAttempts) {
        cancelResponseBody(response);
        const retryDelayMs = retryDelayForResponse(response, attempt, retryBaseDelayMs, retryMaxDelayMs);
        if (timeoutHandle) {
          clearTimeout(timeoutHandle);
          timeoutHandle = undefined;
        }
        await cancellableDelay(retryDelayMs, callerSignal);
        continue;
      }
      if (!response.ok) return { response, data: null };
      const data = await Promise.race([parse(response), timeoutGate, callerAbortGate]);
      return { response, data };
    } catch (error) {
      if (callerSignal?.aborted) throw callerAbortReason(callerSignal);
      const effectiveError = controller.signal.reason instanceof SecFetchTimeoutError ? controller.signal.reason : error;
      lastError = effectiveError;
      if (attempt >= maxAttempts || !retryableSecError(effectiveError)) throw effectiveError;
      const retryDelayMs = Math.min(retryMaxDelayMs, retryBaseDelayMs * 2 ** (attempt - 1));
      await cancellableDelay(retryDelayMs, callerSignal);
    } finally {
      if (timeoutHandle) clearTimeout(timeoutHandle);
      callerSignal?.removeEventListener("abort", onCallerAbort);
      rejectCallerAbort = undefined;
    }
  }

  throw lastError instanceof Error ? lastError : new Error(`SEC request failed: ${url}`);
}

export function selectSecCompanyMatch(query: string, directory: SecCompanyDirectoryEntry[]) {
  const normalizedQuery = normalizeCompanyLookup(query);
  if (!normalizedQuery) throw new Error("Enter a ticker, CIK, or company name containing letters or numbers.");
  const normalizedCik = normalizedCikQuery(query);

  const exactTicker = directory.filter((company) => normalizeCompanyLookup(company.ticker) === normalizedQuery);
  if (exactTicker.length) return singleExactMatch(query, exactTicker, "ticker");

  const exactTitle = directory.filter((company) => normalizeCompanyLookup(company.title) === normalizedQuery);
  if (exactTitle.length) return singleExactMatch(query, exactTitle, "company name");

  if (normalizedCik) {
    const exactCik = directory.filter((company) => normalizeCikValue(company.cik) === normalizedCik);
    if (exactCik.length) return singleExactMatch(query, exactCik, "CIK");
  }

  const partialTitles = directory.filter((company) => normalizeCompanyLookup(company.title).includes(normalizedQuery));
  if (partialTitles.length === 1) return partialTitles[0];
  if (partialTitles.length > 1) throw ambiguousCompanyMatchError(query, partialTitles);
  throw new Error(`No SEC company match found for "${query}".`);
}

function validConfiguredSecUserAgent(value: string) {
  if (value.length < 10 || value.length > 250) return false;
  const email = value.match(/[A-Z0-9._%+-]+@([A-Z0-9.-]+\.[A-Z]{2,})/i);
  if (!email) return false;
  const normalized = value.toLowerCase();
  const domain = email[1].toLowerCase();
  if (/\b(?:contact|your[-_. ]?email|email)@example\.(?:com|org|net)\b/.test(normalized)) return false;
  if (/^(?:example\.(?:com|org|net)|localhost|invalid)$/.test(domain)) return false;
  if (/\.(?:example|invalid|localhost|test)$/.test(domain) || /(?:^|\.)your-domain\./.test(domain)) return false;
  return true;
}

function retryableSecStatus(status: number) {
  return status === 408 || status === 425 || status === 429 || status >= 500;
}

function retryableSecError(error: unknown) {
  if (error instanceof SecFetchTimeoutError) return true;
  if (!(error instanceof Error)) return true;
  return error.name === "AbortError" || error.name === "TypeError" || /network|fetch|socket|timed?\s*out|reset/i.test(error.message);
}

function retryDelayForResponse(response: Response, attempt: number, baseMs: number, maxMs: number) {
  const retryAfter = response.headers?.get?.("Retry-After")?.trim();
  if (retryAfter) {
    const seconds = Number(retryAfter);
    if (Number.isFinite(seconds) && seconds >= 0) return Math.min(maxMs, Math.max(baseMs, seconds * 1_000));
    const dateMs = Date.parse(retryAfter);
    if (Number.isFinite(dateMs)) return Math.min(maxMs, Math.max(baseMs, dateMs - Date.now()));
  }
  return Math.min(maxMs, baseMs * 2 ** (attempt - 1));
}

function cancelResponseBody(response: Response) {
  void response.body?.cancel().catch(() => {
    // The next bounded attempt remains safe if a mocked or already-consumed
    // response body cannot be cancelled.
  });
}

function cancellableDelay(ms: number, signal?: AbortSignal) {
  if (ms <= 0) return Promise.resolve();
  throwIfCallerAborted(signal);
  return new Promise<void>((resolve, reject) => {
    const timeout = setTimeout(finish, ms);
    const onAbort = () => {
      clearTimeout(timeout);
      signal?.removeEventListener("abort", onAbort);
      reject(callerAbortReason(signal));
    };
    function finish() {
      signal?.removeEventListener("abort", onAbort);
      resolve();
    }
    signal?.addEventListener("abort", onAbort, { once: true });
  });
}

function throwIfCallerAborted(signal?: AbortSignal): asserts signal is AbortSignal | undefined {
  if (signal?.aborted) throw callerAbortReason(signal);
}

function callerAbortReason(signal?: AbortSignal) {
  if (signal?.reason instanceof Error) return signal.reason;
  const error = new Error("SEC request was cancelled by the caller.");
  error.name = "AbortError";
  return error;
}

function boundedPositiveInteger(value: unknown, fallback: number, min: number, max: number) {
  const parsed = Number(value);
  if (!Number.isFinite(parsed) || parsed <= 0) return fallback;
  return Math.min(max, Math.max(min, Math.floor(parsed)));
}

function normalizeCompanyLookup(value: string) {
  return String(value ?? "").toLowerCase().replace(/[^a-z0-9]/g, "");
}

function normalizedCikQuery(query: string) {
  const trimmed = String(query ?? "").trim();
  if (!/^(?:cik\s*)?\d{1,10}$/i.test(trimmed)) return "";
  return normalizeCikValue(trimmed);
}

function normalizeCikValue(value: string) {
  const digits = String(value ?? "").replace(/\D/g, "");
  return digits ? digits.padStart(10, "0") : "";
}

function singleExactMatch(query: string, matches: SecCompanyDirectoryEntry[], kind: string) {
  const uniqueMatches = uniqueCompanyMatches(matches);
  if (uniqueMatches.length === 1) return uniqueMatches[0];
  throw ambiguousCompanyMatchError(query, uniqueMatches, `More than one SEC issuer uses that exact ${kind}.`);
}

function ambiguousCompanyMatchError(query: string, matches: SecCompanyDirectoryEntry[], detail = "The company-name fragment is not unique.") {
  const suggestions = uniqueCompanyMatches(matches)
    .slice(0, 6)
    .map((company) => `${company.ticker || "no ticker"} — ${company.title} (CIK ${normalizeCikValue(company.cik)})`)
    .join("; ");
  return new Error(
    `Multiple SEC companies match "${query}". ${detail} Enter an exact ticker, full legal name, or CIK.${suggestions ? ` Matches include: ${suggestions}.` : ""}`
  );
}

function uniqueCompanyMatches(matches: SecCompanyDirectoryEntry[]) {
  const byCik = new Map<string, SecCompanyDirectoryEntry>();
  for (const company of matches) {
    const cik = normalizeCikValue(company.cik);
    if (cik && !byCik.has(cik)) byCik.set(cik, { ...company, cik });
  }
  return Array.from(byCik.values());
}
