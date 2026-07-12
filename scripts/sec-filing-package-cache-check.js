const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const Module = require("node:module");
const ts = require("typescript");

const repoRoot = path.resolve(__dirname, "..");
const sourcePath = path.join(repoRoot, "server", "fill-model", "sec-filing-package.ts");
const requestHeaders = {
  "User-Agent": "HistoricalsSolver SEC regression sec-tests@historicalssolver.dev"
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

function registerTypeScriptRequire() {
  if (require.extensions[".ts"]) return;
  require.extensions[".ts"] = (mod, file) => {
    mod._compile(compileTypeScript(fs.readFileSync(file, "utf8")), file);
  };
}

function loadTypeScriptModule(file, environment) {
  registerTypeScriptRequire();
  const previous = Object.fromEntries(Object.keys(environment).map((key) => [key, process.env[key]]));
  Object.assign(process.env, environment);
  try {
    const source = fs.readFileSync(file, "utf8");
    const compiled = compileTypeScript(source);
    const mod = new Module(`${file}?cache-check=${Math.random()}`, module);
    mod.filename = file;
    mod.paths = Module._nodeModulePaths(path.dirname(file));
    mod._compile(compiled, file);
    return mod.exports;
  } finally {
    for (const [key, value] of Object.entries(previous)) {
      if (value === undefined) delete process.env[key];
      else process.env[key] = value;
    }
  }
}

function filing(index) {
  return {
    cik: "1",
    accessionNumber: `0000000001-26-${String(index).padStart(6, "0")}`,
    form: "10-Q",
    primaryDocument: "test.htm"
  };
}

async function main() {
  const { fetchSecFilingPackageSupport, fetchSecXbrlFrame, __secFilingPackageTestHooks: hooks } = loadTypeScriptModule(
    sourcePath,
    {
      SEC_ARCHIVE_MIN_INTERVAL_MS: "25",
      SEC_ARCHIVE_RETRY_BASE_DELAY_MS: "1",
      SEC_FILING_RESPONSE_CACHE_MAX_ENTRIES: "2",
      SEC_FILING_PACKAGE_CACHE_MAX_ENTRIES: "2",
      SEC_FILING_CACHE_TTL_MS: "1000"
    }
  );

  let now = 0;
  const unitCache = hooks.createBoundedTtlLruCache(2, 100, () => now);
  unitCache.set("a", "first");
  unitCache.set("b", "second");
  assert.equal(unitCache.get("a"), "first");
  unitCache.set("c", "third");
  assert.equal(unitCache.get("b"), undefined, "The least-recently-used cache entry was not evicted.");
  assert.deepEqual(unitCache.keys(), ["a", "c"]);
  now = 100;
  assert.equal(unitCache.size, 0, "Expired entries remained visible after their TTL elapsed.");

  const originalFetch = global.fetch;
  try {
    hooks.clearCaches();
    const fetchCounts = new Map();
    global.fetch = async (url) => {
      const key = String(url);
      fetchCounts.set(key, (fetchCounts.get(key) || 0) + 1);
      return new Response(`body:${key}`, { status: 200 });
    };

    const urls = [1, 2, 3].map((index) => `https://example.test/artifact-${index}`);
    await hooks.fetchSecText(urls[0], requestHeaders, "text/plain");
    await hooks.fetchSecText(urls[1], requestHeaders, "text/plain");
    await hooks.fetchSecText(urls[0], requestHeaders, "text/plain");
    await hooks.fetchSecText(urls[2], requestHeaders, "text/plain");
    assert.deepEqual(
      hooks.cacheState().completedResponseTextKeys,
      [urls[0], urls[2]],
      "Completed response text did not use its configured LRU bound."
    );
    await hooks.fetchSecText(urls[1], requestHeaders, "text/plain");
    assert.equal(fetchCounts.get(urls[1]), 2, "An evicted response was unexpectedly retained.");
    assert.equal(hooks.cacheState().inFlightResponseText, 0, "Settled response promises remained in the in-flight cache.");

    hooks.clearCaches();
    let transientAttempts = 0;
    global.fetch = async () => {
      transientAttempts += 1;
      if (transientAttempts <= 3) return new Response("temporarily unavailable", { status: 503 });
      return new Response("recovered", { status: 200 });
    };
    const transientUrl = "https://example.test/transient-artifact";
    assert.equal(await hooks.fetchSecText(transientUrl, requestHeaders, "text/plain"), null);
    assert.deepEqual(
      hooks.cacheState().completedResponseTextKeys,
      [],
      "A failed response was retained as a completed cache entry."
    );
    assert.equal(await hooks.fetchSecText(transientUrl, requestHeaders, "text/plain"), "recovered");
    assert.equal(await hooks.fetchSecText(transientUrl, requestHeaders, "text/plain"), "recovered");
    assert.equal(transientAttempts, 4, "A transient failure either poisoned the cache or a successful retry was not reused.");

    hooks.clearCaches();
    let jsonAttempts = 0;
    global.fetch = async () => {
      jsonAttempts += 1;
      return new Response(jsonAttempts === 1 ? "{malformed" : JSON.stringify({ recovered: true }), { status: 200 });
    };
    const firstJson = await fetchSecXbrlFrame("us-gaap", "Assets", "USD", "CY2025Q4I", requestHeaders);
    assert.equal(firstJson, null);
    assert.deepEqual(hooks.cacheState().completedResponseTextKeys, [], "Malformed JSON left its raw response permanently cached.");
    const recoveredJson = await fetchSecXbrlFrame("us-gaap", "Assets", "USD", "CY2025Q4I", requestHeaders);
    assert.deepEqual(recoveredJson, { recovered: true });
    assert.deepEqual(
      await fetchSecXbrlFrame("us-gaap", "Assets", "USD", "CY2025Q4I", requestHeaders),
      { recovered: true }
    );
    assert.equal(jsonAttempts, 2, "Malformed SEC JSON was negative-cached or a valid JSON result was not reused.");

    hooks.clearCaches();
    let packageFetches = 0;
    global.fetch = async (url) => {
      packageFetches += 1;
      const value = String(url);
      if (value.endsWith("/index.json")) {
        return new Response(JSON.stringify({ directory: { item: [{ name: "test.htm", type: "text/html" }] } }), { status: 200 });
      }
      return new Response("<html><body></body></html>", { status: 200 });
    };
    for (const index of [1, 2, 3]) {
      const result = await fetchSecFilingPackageSupport([filing(index)], requestHeaders);
      assert.equal(result.packages.length, 1);
    }
    assert.deepEqual(
      hooks.cacheState().completedPackageKeys,
      ["1:000000000126000002", "1:000000000126000003"],
      "Parsed SEC filing packages exceeded their configured LRU bound."
    );
    const beforeReload = packageFetches;
    const reloaded = await fetchSecFilingPackageSupport([filing(1)], requestHeaders);
    assert.equal(reloaded.packages.length, 1);
    assert.ok(packageFetches > beforeReload, "An evicted filing package was not reloaded.");
    assert.equal(hooks.cacheState().completedPackageKeys.length, 2);
    assert.equal(hooks.cacheState().inFlightPackages, 0, "Settled package promises remained in the in-flight cache.");

    hooks.clearCaches();
    global.fetch = async (url) => {
      await new Promise((resolve) => setTimeout(resolve, 100));
      return new Response(`pending:${String(url)}`, { status: 200 });
    };
    const concurrent = [1, 2, 3, 4].map((index) =>
      hooks.fetchSecText(`https://example.test/concurrent-${index}`, requestHeaders, "text/plain")
    );
    assert.equal(hooks.cacheState().inFlightResponseText, 2, "The process-global in-flight response map exceeded its bound.");
    await Promise.all(concurrent);
    assert.equal(hooks.cacheState().inFlightResponseText, 0);
    assert.equal(hooks.cacheState().completedResponseTextKeys.length, 2);
  } finally {
    global.fetch = originalFetch;
  }

  console.log("SEC filing-package bounded LRU/TTL caches and retryable failure handling passed.");
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
