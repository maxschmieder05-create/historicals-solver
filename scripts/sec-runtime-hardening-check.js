const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const Module = require("node:module");
const ts = require("typescript");

const repoRoot = path.resolve(__dirname, "..");
const sourcePath = path.join(repoRoot, "server", "fill-model", "sec-http.ts");
const validProductionEnv = {
  NODE_ENV: "production",
  SEC_USER_AGENT: "HistoricalsSolver SEC integration sec-operations@historicalssolver.dev"
};

function compileTypeScript(source) {
  return ts.transpileModule(source, {
    compilerOptions: {
      module: ts.ModuleKind.CommonJS,
      target: ts.ScriptTarget.ES2020,
      esModuleInterop: true
    }
  }).outputText;
}

function loadTypeScriptModule(file) {
  const source = fs.readFileSync(file, "utf8");
  const compiled = compileTypeScript(source);
  const mod = new Module(file, module);
  mod.filename = file;
  mod.paths = Module._nodeModulePaths(path.dirname(file));
  mod._compile(compiled, file);
  return mod.exports;
}

async function main() {
  const {
    SecFetchTimeoutError,
    fetchSecJson,
    resolveSecUserAgent,
    secRequestHeaders,
    selectSecCompanyMatch
  } = loadTypeScriptModule(sourcePath);

  assert.throws(
    () => resolveSecUserAgent({ NODE_ENV: "production" }),
    /SEC_USER_AGENT.*required.*production/i,
    "Production SEC access must fail clearly when no accountable identity is configured."
  );
  assert.throws(
    () => resolveSecUserAgent({ NODE_ENV: "production", SEC_USER_AGENT: "HistoricalsSolver/0.1 contact@example.com" }),
    /placeholder|reserved example/i,
    "The former example.com placeholder must never pass production identity validation."
  );
  assert.equal(resolveSecUserAgent(validProductionEnv), validProductionEnv.SEC_USER_AGENT);
  assert.equal(
    new Headers(secRequestHeaders({}, "application/json", validProductionEnv)).get("user-agent"),
    validProductionEnv.SEC_USER_AGENT
  );
  let unconfiguredProductionFetches = 0;
  await assert.rejects(
    fetchSecJson("https://data.sec.gov/submissions/CIK0000320193.json", {}, {
      env: { NODE_ENV: "production" },
      fetchImpl: async () => {
        unconfiguredProductionFetches += 1;
        return new Response("{}", { status: 200 });
      }
    }),
    /SEC_USER_AGENT.*required.*production/i
  );
  assert.equal(unconfiguredProductionFetches, 0, "A production process with no SEC identity must fail before opening a network request.");

  const retryStatuses = [500, 429, 200];
  const attemptSignals = [];
  const retryResult = await fetchSecJson(
    "https://data.sec.gov/api/xbrl/companyfacts/CIK0000320193.json",
    {},
    {
      env: validProductionEnv,
      timeoutMs: 200,
      maxAttempts: 3,
      retryBaseDelayMs: 1,
      retryMaxDelayMs: 2,
      fetchImpl: async (_url, init = {}) => {
        attemptSignals.push(init.signal);
        const status = retryStatuses.shift();
        assert.equal(new Headers(init.headers).get("user-agent"), validProductionEnv.SEC_USER_AGENT);
        return new Response(status === 200 ? JSON.stringify({ entityName: "Apple Inc." }) : "retry", {
          status,
          headers: status === 429 ? { "Retry-After": "0" } : undefined
        });
      }
    }
  );
  assert.equal(retryResult.response.status, 200);
  assert.deepEqual(retryResult.data, { entityName: "Apple Inc." });
  assert.equal(attemptSignals.length, 3, "Transient 5xx and 429 responses should be retried only up to the configured bound.");
  assert.equal(new Set(attemptSignals).size, 3, "Every attempt should have its own timeout controller.");

  let permanentFailureAttempts = 0;
  const permanentFailure = await fetchSecJson("https://www.sec.gov/files/company_tickers.json", {}, {
    env: validProductionEnv,
    timeoutMs: 200,
    maxAttempts: 3,
    fetchImpl: async () => {
      permanentFailureAttempts += 1;
      return new Response("not found", { status: 404 });
    }
  });
  assert.equal(permanentFailure.response.status, 404);
  assert.equal(permanentFailure.data, null);
  assert.equal(permanentFailureAttempts, 1, "Permanent 4xx responses must fail immediately instead of consuming the retry budget.");

  let timedOutAttempts = 0;
  const timedOutSignals = [];
  await assert.rejects(
    fetchSecJson("https://data.sec.gov/submissions/CIK0000320193.json", {}, {
      env: validProductionEnv,
      timeoutMs: 15,
      maxAttempts: 2,
      retryBaseDelayMs: 1,
      retryMaxDelayMs: 1,
      fetchImpl: (_url, init = {}) => {
        timedOutAttempts += 1;
        timedOutSignals.push(init.signal);
        return new Promise((_resolve, reject) => {
          const onAbort = () => reject(init.signal?.reason ?? Object.assign(new Error("aborted"), { name: "AbortError" }));
          init.signal?.addEventListener("abort", onAbort, { once: true });
          if (init.signal?.aborted) onAbort();
        });
      }
    }),
    (error) => error instanceof SecFetchTimeoutError,
    "A hung SEC response must terminate after a bounded timeout instead of hanging the workbook request."
  );
  assert.equal(timedOutAttempts, 2, "Timeouts should retry only to the configured attempt bound.");
  assert.ok(timedOutSignals.every((signal) => signal?.aborted), "Each timed-out network attempt must abort its underlying fetch.");

  const callerController = new AbortController();
  let cancellationAttempts = 0;
  let cancellationSignal;
  const cancellation = fetchSecJson("https://www.sec.gov/Archives/edgar/data/320193/test.htm", {
    signal: callerController.signal
  }, {
    env: validProductionEnv,
    timeoutMs: 1_000,
    maxAttempts: 3,
    fetchImpl: (_url, init = {}) => {
      cancellationAttempts += 1;
      cancellationSignal = init.signal;
      return new Promise((_resolve, reject) => {
        const onAbort = () => reject(init.signal?.reason ?? Object.assign(new Error("aborted"), { name: "AbortError" }));
        init.signal?.addEventListener("abort", onAbort, { once: true });
      });
    }
  });
  const callerAbort = Object.assign(new Error("client disconnected"), { name: "AbortError" });
  callerController.abort(callerAbort);
  await assert.rejects(cancellation, (error) => error === callerAbort);
  assert.equal(cancellationAttempts, 1, "Caller cancellation must not start a retry.");
  assert.equal(cancellationSignal?.aborted, true, "Caller cancellation must reach the active SEC fetch.");

  const companies = [
    { cik: "0000320193", ticker: "AAPL", title: "Apple Inc." },
    { cik: "0001604028", ticker: "APLE", title: "Apple Hospitality REIT, Inc." },
    { cik: "0000789019", ticker: "MSFT", title: "Microsoft Corporation" }
  ];
  assert.equal(selectSecCompanyMatch("AAPL", companies).cik, "0000320193", "Exact ticker lookup must remain supported.");
  assert.equal(selectSecCompanyMatch("Apple Inc.", companies).cik, "0000320193", "Exact legal-name lookup must remain supported.");
  assert.equal(selectSecCompanyMatch("CIK 320193", companies).cik, "0000320193", "Exact CIK lookup must remain supported.");
  assert.equal(selectSecCompanyMatch("Microsoft", companies).cik, "0000789019", "A unique partial legal-name match remains usable.");
  assert.throws(
    () => selectSecCompanyMatch("Apple", companies),
    /Multiple SEC companies match.*exact ticker.*full legal name.*CIK/i,
    "An ambiguous partial legal-name match must fail closed instead of selecting the first issuer."
  );

  console.log("SEC runtime identity, timeout, retry, cancellation, and issuer-match hardening passed.");
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
