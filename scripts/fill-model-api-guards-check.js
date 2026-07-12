const assert = require("node:assert/strict");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const Module = require("node:module");
const ts = require("typescript");

const repoRoot = path.resolve(__dirname, "..");
const routePath = path.join(repoRoot, "app", "api", "fill-model", "route.ts");
const servicePath = path.join(repoRoot, "server", "fill-model", "fill-model-service.ts");
const testLockPath = path.join(repoRoot, "tmp", `fill-model-api-guards-${process.pid}.lock`);
process.env.FILL_MODEL_PROCESS_LOCK_PATH = testLockPath;
process.on("exit", () => fs.rmSync(testLockPath, { force: true }));

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
  require.extensions[".ts"] = (mod, filename) => {
    mod._compile(compileTypeScript(fs.readFileSync(filename, "utf8")), filename);
  };
}

function loadRoute(fillModelWorkbook, env = {}) {
  registerTypeScriptRequire();
  const priorEnvironment = {};
  for (const [key, value] of Object.entries(env)) {
    priorEnvironment[key] = process.env[key];
    if (value === undefined) delete process.env[key];
    else process.env[key] = String(value);
  }

  const serviceModule = new Module(servicePath, module);
  serviceModule.filename = servicePath;
  serviceModule.loaded = true;
  serviceModule.exports = {
    fillModelWorkbook,
    fillModelErrorDetails: (error) => ({
      message: error instanceof Error ? error.message : String(error),
      status: 500
    })
  };
  require.cache[servicePath] = serviceModule;

  try {
    const routeModule = new Module(routePath, module);
    routeModule.filename = routePath;
    routeModule.paths = Module._nodeModulePaths(path.dirname(routePath));
    routeModule._compile(compileTypeScript(fs.readFileSync(routePath, "utf8")), routePath);
    return routeModule.exports;
  } finally {
    delete require.cache[servicePath];
    for (const [key, value] of Object.entries(priorEnvironment)) {
      if (value === undefined) delete process.env[key];
      else process.env[key] = value;
    }
  }
}

function successfulFill() {
  return {
    output: Buffer.from("filled workbook"),
    outputName: "TEST_historicals_filled.xlsx",
    summary: {
      companyName: "Test Company",
      ticker: "TEST",
      periods: ["2025"],
      filledCells: 1,
      commentsAdded: 1,
      warnings: [],
      llm: { telemetry: [{ errorMessage: "x".repeat(100_000) }] },
      internalOnly: "must not be serialized into a response header"
    }
  };
}

function workbookRequest({
  name = "template.xlsx",
  contents = "PK workbook",
  ticker = "TEST",
  authorization,
  contentLength,
  signal
} = {}) {
  const form = new FormData();
  if (ticker !== null) form.set("ticker", ticker);
  if (name !== null) {
    form.set(
      "file",
      new File([contents], name, {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      })
    );
  }
  const headers = {};
  if (authorization) headers.authorization = authorization;
  if (contentLength !== undefined) headers["content-length"] = String(contentLength);
  return new Request("http://localhost/api/fill-model", {
    method: "POST",
    headers,
    body: form,
    signal
  });
}

function streamingRequest(chunks, { signal, contentType = "multipart/form-data; boundary=test-boundary", onCancel } = {}) {
  let index = 0;
  const body = new ReadableStream({
    pull(controller) {
      if (index >= chunks.length) {
        controller.close();
        return;
      }
      controller.enqueue(new TextEncoder().encode(chunks[index]));
      index += 1;
    },
    cancel(reason) {
      onCancel?.(reason);
    }
  });
  return new Request("http://localhost/api/fill-model", {
    method: "POST",
    headers: { "content-type": contentType },
    body,
    signal,
    duplex: "half"
  });
}

async function responseJson(response) {
  return JSON.parse(await response.text());
}

async function main() {
  let acceptedCalls = 0;
  let acceptedSignal;
  const acceptedRoute = loadRoute(async (input) => {
    acceptedCalls += 1;
    acceptedSignal = input.signal;
    return successfulFill();
  }, {
    FILL_MODEL_MAX_UPLOAD_BYTES: undefined,
    FILL_MODEL_MAX_CONCURRENT_REQUESTS: undefined,
    HISTORICALS_API_KEY: undefined
  });

  const acceptedRequest = workbookRequest({ name: "template.XLSX" });
  const accepted = await acceptedRoute.POST(acceptedRequest);
  assert.equal(accepted.status, 200, ".xlsx uploads should reach the fill service");
  assert.equal(acceptedCalls, 1);
  assert.equal(acceptedSignal, acceptedRequest.signal, "the route must pass the live client signal into the fill service");
  assert.equal(accepted.headers.get("x-output-filename"), "TEST_historicals_filled.xlsx");
  assert.equal(accepted.headers.get("cache-control"), "no-store");
  const browserSummary = JSON.parse(decodeURIComponent(accepted.headers.get("x-fill-summary")));
  assert.equal(browserSummary.companyName, "Test Company");
  assert.equal(browserSummary.llm, undefined, "large internal LLM telemetry must not be serialized into the browser response header");
  assert.equal(browserSummary.internalOnly, undefined, "only browser-visible fill summary fields belong in the response header");
  assert.ok(accepted.headers.get("x-fill-summary").length < 2_000, "the browser summary header must remain compact");
  assert.equal(await accepted.text(), "filled workbook");

  const noisySummaryRoute = loadRoute(async () => ({
    ...successfulFill(),
    summary: {
      ...successfulFill().summary,
      warnings: Array.from({ length: 100 }, (_unused, index) => `Warning ${index}: ${"x".repeat(2_000)}`)
    }
  }));
  const noisySummaryResponse = await noisySummaryRoute.POST(workbookRequest());
  assert.equal(noisySummaryResponse.status, 200);
  const noisySummaryHeader = noisySummaryResponse.headers.get("x-fill-summary");
  assert.ok(noisySummaryHeader.length < 8_000, "workbook-controlled warning text must never create an oversized response header");
  const noisyBrowserSummary = JSON.parse(decodeURIComponent(noisySummaryHeader));
  assert.equal(noisyBrowserSummary.warnings.length, 12);
  assert.ok(noisyBrowserSummary.warnings.every((warning) => warning.length <= 240));

  const rejectedMacro = await acceptedRoute.POST(workbookRequest({ name: "template.xlsm" }));
  assert.equal(rejectedMacro.status, 400, ".xlsm uploads must fail before the fill service");
  assert.match((await responseJson(rejectedMacro)).error, /\.xlsm files are rejected/i);
  assert.equal(acceptedCalls, 1);

  const rejectedMissingTicker = await acceptedRoute.POST(workbookRequest({ ticker: "" }));
  assert.equal(rejectedMissingTicker.status, 400);
  assert.match((await responseJson(rejectedMissingTicker)).error, /ticker or company name/i);
  assert.equal(acceptedCalls, 1);

  const rejectedEmpty = await acceptedRoute.POST(workbookRequest({ contents: "" }));
  assert.equal(rejectedEmpty.status, 400);
  assert.match((await responseJson(rejectedEmpty)).error, /empty/i);
  assert.equal(acceptedCalls, 1);

  const sizeRoute = loadRoute(async () => {
    throw new Error("oversized input reached fill service");
  }, {
    FILL_MODEL_MAX_UPLOAD_BYTES: "8",
    FILL_MODEL_MAX_MULTIPART_OVERHEAD_BYTES: "512",
    FILL_MODEL_MAX_CONCURRENT_REQUESTS: undefined,
    HISTORICALS_API_KEY: undefined
  });
  const rejectedByHeader = await sizeRoute.POST(workbookRequest({ contentLength: 521 }));
  assert.equal(rejectedByHeader.status, 413, "declared oversized requests must fail before multipart parsing");
  assert.match((await responseJson(rejectedByHeader)).error, /upload exceeds/i);

  const rejectedByFileSize = await sizeRoute.POST(workbookRequest({ contents: "123456789" }));
  assert.equal(rejectedByFileSize.status, 413, "oversized files must fail even without a content-length header");
  assert.match((await responseJson(rejectedByFileSize)).error, /upload exceeds/i);

  let oversizedChunkedCancelled = false;
  const chunkedSizeRoute = loadRoute(async () => {
    throw new Error("oversized chunked input reached fill service");
  }, {
    FILL_MODEL_MAX_UPLOAD_BYTES: "8",
    FILL_MODEL_MAX_MULTIPART_OVERHEAD_BYTES: "4",
    FILL_MODEL_MAX_CONCURRENT_REQUESTS: undefined,
    HISTORICALS_API_KEY: undefined
  });
  const rejectedChunked = await chunkedSizeRoute.POST(
    streamingRequest(["12345678", "90123456", "unread trailing body"], {
      onCancel: () => {
        oversizedChunkedCancelled = true;
      }
    })
  );
  assert.equal(rejectedChunked.status, 413, "chunked requests must be bounded while the body stream is read");
  assert.match((await responseJson(rejectedChunked)).error, /upload exceeds/i);
  assert.equal(oversizedChunkedCancelled, true, "the route must cancel the remaining request stream after the byte limit is crossed");

  let timedOutUploadCancelled = false;
  const uploadTimeoutRoute = loadRoute(async () => {
    throw new Error("timed-out upload reached fill service");
  }, {
    FILL_MODEL_UPLOAD_READ_TIMEOUT_MS: "10",
    FILL_MODEL_MAX_UPLOAD_BYTES: "1024",
    FILL_MODEL_MAX_MULTIPART_OVERHEAD_BYTES: "1024",
    FILL_MODEL_MAX_CONCURRENT_REQUESTS: undefined,
    HISTORICALS_API_KEY: undefined
  });
  const stalledBody = new ReadableStream({
    pull() {
      return new Promise(() => {});
    },
    cancel() {
      timedOutUploadCancelled = true;
    }
  });
  const timedOutUpload = await uploadTimeoutRoute.POST(
    new Request("http://localhost/api/fill-model", {
      method: "POST",
      headers: { "content-type": "multipart/form-data; boundary=stalled" },
      body: stalledBody,
      duplex: "half"
    })
  );
  assert.equal(timedOutUpload.status, 408, "a stalled chunked upload must release the route within the upload deadline");
  assert.equal(timedOutUploadCancelled, true);

  let authenticatedCalls = 0;
  const authenticatedRoute = loadRoute(async () => {
    authenticatedCalls += 1;
    return successfulFill();
  }, {
    HISTORICALS_API_KEY: "release-secret",
    FILL_MODEL_MAX_UPLOAD_BYTES: undefined,
    FILL_MODEL_MAX_CONCURRENT_REQUESTS: undefined
  });
  process.env.HISTORICALS_API_KEY = "release-secret";
  try {
    const missingAuth = await authenticatedRoute.POST(workbookRequest());
    assert.equal(missingAuth.status, 401);
    assert.equal(missingAuth.headers.get("www-authenticate"), 'Bearer realm="Historicals Solver"');
    assert.match((await responseJson(missingAuth)).error, /access key/i);
    const wrongAuth = await authenticatedRoute.POST(workbookRequest({ authorization: "Bearer wrong" }));
    assert.equal(wrongAuth.status, 401);
    const validAuth = await authenticatedRoute.POST(workbookRequest({ authorization: "Bearer release-secret" }));
    assert.equal(validAuth.status, 200);
    assert.equal(authenticatedCalls, 1);
  } finally {
    delete process.env.HISTORICALS_API_KEY;
  }

  const priorNodeEnv = process.env.NODE_ENV;
  const priorUnauthenticatedFlag = process.env.ALLOW_UNAUTHENTICATED_FILL;
  delete process.env.HISTORICALS_API_KEY;
  process.env.NODE_ENV = "production";
  delete process.env.ALLOW_UNAUTHENTICATED_FILL;
  const productionAuthRoute = loadRoute(async () => successfulFill(), {
    HISTORICALS_API_KEY: undefined,
    ALLOW_UNAUTHENTICATED_FILL: undefined
  });
  try {
    const unconfiguredProduction = await productionAuthRoute.POST(workbookRequest());
    assert.equal(unconfiguredProduction.status, 503, "production must fail closed when the deployment access key is missing");
    assert.match((await responseJson(unconfiguredProduction)).error, /not configured for workbook uploads/i);

    process.env.ALLOW_UNAUTHENTICATED_FILL = "true";
    const explicitlyOpenProduction = await productionAuthRoute.POST(workbookRequest());
    assert.equal(explicitlyOpenProduction.status, 200, "an intentionally isolated deployment may explicitly opt into unauthenticated fills");
  } finally {
    if (priorNodeEnv === undefined) delete process.env.NODE_ENV;
    else process.env.NODE_ENV = priorNodeEnv;
    if (priorUnauthenticatedFlag === undefined) delete process.env.ALLOW_UNAUTHENTICATED_FILL;
    else process.env.ALLOW_UNAUTHENTICATED_FILL = priorUnauthenticatedFlag;
  }

  const alreadyAborted = new AbortController();
  alreadyAborted.abort();
  const rejectedCancelled = await acceptedRoute.POST(workbookRequest({ signal: alreadyAborted.signal }));
  assert.equal(rejectedCancelled.status, 499, "a request cancelled before upload parsing must fail without invoking the fill service");
  assert.equal(acceptedCalls, 1);

  let serviceStarted;
  let propagatedSignal;
  const cancellableRoute = loadRoute(
    async (input) => {
      propagatedSignal = input.signal;
      serviceStarted?.();
      return new Promise((resolve) => input.signal.addEventListener("abort", () => resolve(successfulFill()), { once: true }));
    },
    {
      FILL_MODEL_MAX_UPLOAD_BYTES: undefined,
      FILL_MODEL_MAX_CONCURRENT_REQUESTS: undefined,
      HISTORICALS_API_KEY: undefined
    }
  );
  const serviceStartedPromise = new Promise((resolve) => {
    serviceStarted = resolve;
  });
  const inFlightController = new AbortController();
  const cancelledInFlightPromise = cancellableRoute.POST(workbookRequest({ signal: inFlightController.signal }));
  await serviceStartedPromise;
  inFlightController.abort();
  const cancelledInFlight = await cancelledInFlightPromise;
  assert.equal(propagatedSignal.aborted, true, "client cancellation must reach a running fill service");
  assert.equal(cancelledInFlight.status, 499, "the route must not return a workbook completed after client cancellation");

  let releaseFirst;
  let concurrentCalls = 0;
  const concurrencyRoute = loadRoute(async () => {
    concurrentCalls += 1;
    if (concurrentCalls > 1) return successfulFill();
    return new Promise((resolve) => {
      releaseFirst = () => resolve(successfulFill());
    });
  }, {
    FILL_MODEL_MAX_CONCURRENT_REQUESTS: "1",
    FILL_MODEL_MAX_UPLOAD_BYTES: undefined,
    HISTORICALS_API_KEY: undefined
  });

  const firstRequest = concurrencyRoute.POST(workbookRequest());
  const busyResponse = await concurrencyRoute.POST(workbookRequest());
  assert.equal(busyResponse.status, 429, "a second fill must be rejected while the single worker is occupied");
  assert.equal(busyResponse.headers.get("retry-after"), "15");
  assert.match((await responseJson(busyResponse)).error, /already running/i);
  const secondWorkerRoute = loadRoute(async () => successfulFill(), {
    FILL_MODEL_MAX_CONCURRENT_REQUESTS: "1",
    FILL_MODEL_MAX_UPLOAD_BYTES: undefined,
    HISTORICALS_API_KEY: undefined
  });
  const crossWorkerBusyResponse = await secondWorkerRoute.POST(workbookRequest());
  assert.equal(crossWorkerBusyResponse.status, 429, "the process lock must reject a fill from a second route worker");
  while (!releaseFirst) await new Promise((resolve) => setImmediate(resolve));
  releaseFirst();
  assert.equal((await firstRequest).status, 200);

  const afterRelease = await concurrencyRoute.POST(workbookRequest());
  assert.equal(afterRelease.status, 200, "the concurrency slot must be released in finally");
  assert.equal(concurrentCalls, 2);

  fs.writeFileSync(
    testLockPath,
    JSON.stringify({ pid: 99_999_999, startedAt: new Date().toISOString(), token: "crashed-worker" })
  );
  const crashedWorkerRecoveryRoute = loadRoute(async () => successfulFill(), {
    FILL_MODEL_MAX_CONCURRENT_REQUESTS: "1",
    FILL_MODEL_PROCESS_LOCK_STALE_MS: String(60 * 60 * 1000),
    FILL_MODEL_MAX_UPLOAD_BYTES: undefined,
    HISTORICALS_API_KEY: undefined
  });
  const recoveredAfterCrash = await crashedWorkerRecoveryRoute.POST(workbookRequest());
  assert.equal(
    recoveredAfterCrash.status,
    200,
    "a recent lock owned by a dead worker must be recovered immediately instead of wedging the API until the age timeout"
  );
  assert.equal(fs.existsSync(testLockPath), false, "the recovered worker must release its replacement lock after the request");

  fs.writeFileSync(
    testLockPath,
    JSON.stringify({ pid: process.pid, hostname: os.hostname(), startedAt: new Date(0).toISOString(), token: "live-long-running-worker" })
  );
  const oldTimestamp = new Date(Date.now() - 24 * 60 * 60 * 1000);
  fs.utimesSync(testLockPath, oldTimestamp, oldTimestamp);
  const liveLongRunningLockRoute = loadRoute(async () => {
    throw new Error("a live long-running lock was stolen");
  }, {
    FILL_MODEL_MAX_CONCURRENT_REQUESTS: "1",
    FILL_MODEL_PROCESS_LOCK_STALE_MS: "1",
    FILL_MODEL_MAX_UPLOAD_BYTES: undefined,
    HISTORICALS_API_KEY: undefined
  });
  const liveLongRunningBusy = await liveLongRunningLockRoute.POST(workbookRequest());
  assert.equal(liveLongRunningBusy.status, 429, "lock age alone must never steal a lock from a live local worker");
  fs.rmSync(testLockPath, { force: true });

  console.log("Fill-model API upload/auth/concurrency guards passed.");
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
