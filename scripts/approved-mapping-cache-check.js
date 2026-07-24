const assert = require("node:assert/strict");
const fs = require("node:fs/promises");
const fsSync = require("node:fs");
const Module = require("node:module");
const os = require("node:os");
const path = require("node:path");
const ts = require("typescript");

const repoRoot = path.resolve(__dirname, "..");
const sourcePath = path.join(repoRoot, "server", "fill-model", "approved-mapping-cache.ts");

function loadTypeScriptModule(file) {
  const source = fsSync.readFileSync(file, "utf8");
  const compiled = ts.transpileModule(source, {
    compilerOptions: {
      esModuleInterop: true,
      module: ts.ModuleKind.CommonJS,
      target: ts.ScriptTarget.ES2020
    },
    fileName: file
  }).outputText;
  const mod = new Module(file, module);
  mod.filename = file;
  mod.paths = Module._nodeModulePaths(path.dirname(file));
  mod._compile(compiled, file);
  return mod.exports;
}

function approvedMapping(overrides = {}) {
  return {
    scope: "approved_exact",
    statement: "balance_sheet",
    xbrlTag: "OtherCurrentAssetsExtension",
    modelRow: "Prepaid & Other Current Assets",
    action: "merge_into_other",
    explanation: "Reviewed SEC concept mapping.",
    approvedBy: "controller@example.test",
    approvedAt: "2026-07-24T12:00:00.000Z",
    ...overrides
  };
}

function lookup(overrides = {}) {
  return {
    company: { name: "Alpha Corporation", ticker: "" },
    statement: "balance_sheet",
    sourceTableType: "primary_statement",
    section: "current assets",
    xbrlTag: "OtherCurrentAssetsExtension",
    reportedLabel: "Other current assets",
    availableModelRows: ["Inventory", "Prepaid & Other Current Assets", "Other Non-Current Assets"],
    ...overrides
  };
}

async function main() {
  const { cacheApprovedMapping, findApprovedMapping, loadApprovedMappingCache } = loadTypeScriptModule(sourcePath);

  const nameOnlyCompanyMapping = approvedMapping({
    scope: "company_historical",
    companyTicker: "",
    companyName: "Alpha Corporation"
  });
  assert.equal(
    findApprovedMapping({ version: 1, mappings: [nameOnlyCompanyMapping] }, lookup())?.companyName,
    "Alpha Corporation",
    "A name-bound company mapping should remain usable when the issuer has no ticker."
  );
  assert.equal(
    findApprovedMapping(
      { version: 1, mappings: [nameOnlyCompanyMapping] },
      lookup({ company: { name: "Beta Corporation", ticker: "" } })
    ),
    null,
    "Blank tickers must not make a company-specific mapping match a different issuer."
  );

  const broadMapping = approvedMapping({
    xbrlTag: undefined,
    reportedLabel: "Other current assets",
    modelRow: "Inventory",
    approvedAt: "2026-07-24T14:00:00.000Z"
  });
  const conceptAndContextMapping = approvedMapping({
    reportedLabel: "Other current assets",
    sourceTableType: "primary_statement",
    section: "current assets",
    approvedAt: "2026-07-23T14:00:00.000Z"
  });
  assert.equal(
    findApprovedMapping({ version: 1, mappings: [broadMapping, conceptAndContextMapping] }, lookup())?.modelRow,
    "Prepaid & Other Current Assets",
    "The most specific approved mapping should win even when a broader mapping was approved more recently."
  );

  const tempRoot = await fs.mkdtemp(path.join(os.tmpdir(), "approved-mapping-cache-"));
  const cachePath = path.join(tempRoot, "nested", "approved_mappings.json");
  try {
    const firstEmpty = await loadApprovedMappingCache(path.join(tempRoot, "missing-a.json"));
    firstEmpty.mappings.push(approvedMapping());
    const secondEmpty = await loadApprovedMappingCache(path.join(tempRoot, "missing-b.json"));
    assert.equal(secondEmpty.mappings.length, 0, "Missing cache reads must not share mutable empty state.");

    await cacheApprovedMapping(approvedMapping(), cachePath);
    await Promise.all(
      ["ExtensionOne", "ExtensionTwo", "ExtensionThree"].map((xbrlTag, index) =>
        cacheApprovedMapping(
          approvedMapping({
            xbrlTag,
            reportedLabel: undefined,
            modelRow: index === 0 ? "Inventory" : "Other Non-Current Assets"
          }),
          cachePath
        )
      )
    );

    let persisted = await loadApprovedMappingCache(cachePath);
    assert.equal(persisted.mappings.length, 4, "Concurrent cache updates must retain every distinct approved mapping.");

    await cacheApprovedMapping(
      approvedMapping({ modelRow: "Inventory", action: "map", explanation: "Replacement approval." }),
      cachePath
    );
    persisted = await loadApprovedMappingCache(cachePath);
    assert.equal(persisted.mappings.length, 4, "Updating one mapping identity must replace rather than duplicate it.");
    assert.equal(
      persisted.mappings.find((mapping) => mapping.xbrlTag === "OtherCurrentAssetsExtension")?.explanation,
      "Replacement approval."
    );
    assert.deepEqual(
      (await fs.readdir(path.dirname(cachePath))).filter((name) => name.endsWith(".tmp")),
      [],
      "Atomic cache writes must clean up owned temporary files."
    );

    await fs.writeFile(
      cachePath,
      `${JSON.stringify(
        {
          version: 1,
          mappings: [
            approvedMapping(),
            approvedMapping({ action: "set_zero" }),
            approvedMapping({ statement: 42 }),
            approvedMapping({ xbrlTag: "   " }),
            approvedMapping({ approvedAt: "not-a-timestamp" })
          ]
        },
        null,
        2
      )}\n`
    );
    const filtered = await loadApprovedMappingCache(cachePath);
    assert.equal(
      filtered.mappings.length,
      1,
      "Unsupported actions and malformed string fields must not enter the trusted mapping cache."
    );
    await assert.rejects(
      cacheApprovedMapping(approvedMapping({ action: "set_zero" }), cachePath),
      /incomplete or invalid/i,
      "Unsupported actions must also be rejected on write."
    );

    console.log("Approved mapping matching, validation, and atomic cache persistence checks passed.");
  } finally {
    await fs.rm(tempRoot, { recursive: true, force: true });
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
