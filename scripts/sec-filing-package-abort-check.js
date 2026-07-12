const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const Module = require("node:module");
const ts = require("typescript");

const repoRoot = path.resolve(__dirname, "..");
const sourcePath = path.join(repoRoot, "server", "fill-model", "sec-filing-package.ts");

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

async function main() {
  const { fetchSecFilingPackageSupport } = loadTypeScriptModule(sourcePath);
  const filing = {
    cik: "9999999",
    accessionNumber: "0009999999-26-000001",
    form: "10-Q",
    primaryDocument: "test.htm"
  };
  const originalFetch = global.fetch;
  let fetchCount = 0;
  let started;
  const fetchStarted = new Promise((resolve) => {
    started = resolve;
  });
  let underlyingSignal;

  try {
    global.fetch = (_url, init = {}) => {
      fetchCount += 1;
      underlyingSignal = init.signal;
      started();
      return new Promise((_resolve, reject) => {
        const rejectAborted = () => reject(init.signal?.reason ?? Object.assign(new Error("aborted"), { name: "AbortError" }));
        init.signal?.addEventListener("abort", rejectAborted, { once: true });
        if (init.signal?.aborted) rejectAborted();
      });
    };

    const controller = new AbortController();
    const abortedLoad = fetchSecFilingPackageSupport(
      [filing],
      { "User-Agent": "HistoricalsSolver SEC regression sec-tests@historicalssolver.dev" },
      controller.signal
    );
    await fetchStarted;
    controller.abort();

    await assert.rejects(
      abortedLoad,
      (error) => error?.name === "AbortError",
      "Caller cancellation must reject instead of becoming an SEC filing-package warning."
    );
    assert.equal(underlyingSignal?.aborted, true, "Caller cancellation must abort the active SEC network request.");
    assert.equal(fetchCount, 1);

    global.fetch = async () => {
      fetchCount += 1;
      return {
        ok: true,
        status: 200,
        text: async () => JSON.stringify({ directory: { item: [] } })
      };
    };

    const retry = await fetchSecFilingPackageSupport(
      [filing],
      { "User-Agent": "HistoricalsSolver SEC regression sec-tests@historicalssolver.dev" }
    );
    assert.equal(retry.packages.length, 0);
    assert.deepEqual(retry.warnings, [
      "SEC filing package could not be loaded for accession 0009999999-26-000001."
    ]);
    assert.equal(fetchCount, 2, "An aborted request must not poison the cache used by a later caller.");

    const cachedRetry = await fetchSecFilingPackageSupport(
      [filing],
      { "User-Agent": "HistoricalsSolver SEC regression sec-tests@historicalssolver.dev" }
    );
    assert.deepEqual(cachedRetry, retry);
    assert.equal(fetchCount, 2, "A completed cache entry must remain reusable by callers that omit an AbortSignal.");
  } finally {
    global.fetch = originalFetch;
  }

  console.log("SEC filing-package abort propagation passed.");
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
