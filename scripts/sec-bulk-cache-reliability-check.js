const assert = require("node:assert/strict");
const crypto = require("node:crypto");
const fs = require("node:fs/promises");
const fsSync = require("node:fs");
const Module = require("node:module");
const os = require("node:os");
const path = require("node:path");
const ts = require("typescript");
const JSZip = require("jszip");

const repoRoot = path.resolve(__dirname, "..");
const sourcePath = path.join(repoRoot, "server", "fill-model", "sec-bulk.ts");
const source = fsSync.readFileSync(sourcePath, "utf8");
const compiled = ts.transpileModule(source, {
  compilerOptions: {
    esModuleInterop: true,
    module: ts.ModuleKind.CommonJS,
    target: ts.ScriptTarget.ES2020
  },
  fileName: sourcePath
}).outputText;

function loadHooks(cacheDir) {
  const overrides = {
    SEC_BULK_CACHE_DIR: cacheDir,
    SEC_BULK_ORPHAN_MAX_AGE_MS: "1000",
    SEC_BULK_ORPHAN_MAX_FILES: "2",
    SEC_BULK_REFRESH_LOCK_STALE_MS: "60000",
    SEC_BULK_LOCK_POLL_MS: "5",
    SEC_BULK_METADATA_LOCK_WAIT_MS: "3000",
    SEC_BULK_ZIP_DIRECTORY_CACHE_MAX_ENTRIES: "2",
    SEC_BULK_REFRESH_DISABLED: "1"
  };
  const previous = Object.fromEntries(Object.keys(overrides).map((key) => [key, process.env[key]]));
  Object.assign(process.env, overrides);
  try {
    const mod = new Module(`${sourcePath}?instance=${Math.random()}`, module);
    mod.filename = sourcePath;
    mod.paths = Module._nodeModulePaths(path.dirname(sourcePath));
    mod._compile(compiled, sourcePath);
    return mod.exports.__secBulkTestHooks;
  } finally {
    for (const [key, value] of Object.entries(previous)) {
      if (value === undefined) delete process.env[key];
      else process.env[key] = value;
    }
  }
}

async function archiveBuffer(payload, cik = "0000000001") {
  const zip = new JSZip();
  zip.file(`CIK${cik}.json`, `${JSON.stringify(payload)}\n`);
  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}

async function archiveEntriesBuffer(entries) {
  const zip = new JSZip();
  for (const [name, payload] of Object.entries(entries)) zip.file(name, `${JSON.stringify(payload)}\n`);
  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}

function hash(buffer) {
  return crypto.createHash("sha256").update(buffer).digest("hex");
}

async function assertNoOwnedTempFiles(cacheDir, filename) {
  const names = await fs.readdir(cacheDir);
  assert.deepEqual(
    names.filter((name) => name.startsWith(`${filename}.`) && name.endsWith(".tmp")),
    [],
    `${filename} left an owned temporary download behind.`
  );
}

async function testFailedDownloadPreservesArchiveAndCleansTemp(cacheDir, hooks) {
  const archivePath = path.join(cacheDir, "companyfacts.zip");
  const original = await archiveBuffer({ version: "original", facts: {} });
  await fs.writeFile(archivePath, original);

  await assert.rejects(
    hooks.refreshArchive("companyfacts", {}, {
      force: true,
      fetchImpl: async () => new Response(Buffer.from("not a zip archive"), { status: 200 })
    }),
    /ZIP|central directory/i
  );

  assert.equal(hash(await fs.readFile(archivePath)), hash(original), "An invalid download replaced the last valid SEC archive.");
  await assertNoOwnedTempFiles(cacheDir, "companyfacts.zip");
  assert.equal(await fileExists(`${archivePath}.refresh.lock`), false, "Failed refresh left its archive lock behind.");

  const failingBody = new ReadableStream({
    start(controller) {
      controller.enqueue(Buffer.from("partial download"));
      controller.error(new Error("simulated response stream failure"));
    }
  });
  await assert.rejects(
    hooks.refreshArchive("companyfacts", {}, {
      force: true,
      fetchImpl: async () => new Response(failingBody, { status: 200 })
    }),
    /simulated response stream failure/i
  );
  assert.equal(hash(await fs.readFile(archivePath)), hash(original), "A failed response stream replaced the last valid SEC archive.");
  await assertNoOwnedTempFiles(cacheDir, "companyfacts.zip");
  assert.equal(await fileExists(`${archivePath}.refresh.lock`), false, "Stream failure left its archive lock behind.");
}

async function testSuccessfulValidatedReplacement(cacheDir, hooks) {
  const archivePath = path.join(cacheDir, "companyfacts.zip");
  const replacement = await archiveBuffer({ cik: 1, version: "replacement", facts: { ok: true } });
  const result = await hooks.refreshArchive("companyfacts", {}, {
    force: true,
    fetchImpl: async () => new Response(replacement, { status: 200 })
  });
  assert.equal(result, "refreshed");

  const zip = await hooks.openZip(archivePath);
  const payload = JSON.parse(await zip.readText("CIK0000000001.json"));
  assert.equal(payload.version, "replacement");
  const metadata = await hooks.readMetadata();
  assert.equal(metadata.archives.companyfacts.path, archivePath);
  await assertNoOwnedTempFiles(cacheDir, "companyfacts.zip");
  assert.equal(await fileExists(`${archivePath}.refresh.lock`), false);
  assert.equal(await fileExists(path.join(cacheDir, "metadata.json.lock")), false);
  assert.deepEqual((await fs.readdir(cacheDir)).filter((name) => /^metadata\.json\..*\.tmp$/.test(name)), []);
}

async function testOrphanRetention(cacheDir, hooks) {
  const now = Date.now();
  const validArchivePath = path.join(cacheDir, "submissions.zip");
  const validArchive = await archiveBuffer({ name: "valid submissions", filings: {} });
  await fs.writeFile(validArchivePath, validArchive);
  await fs.writeFile(path.join(cacheDir, "unrelated-user-file.tmp"), "preserve me");

  for (let index = 0; index < 6; index += 1) {
    const filePath = path.join(cacheDir, `submissions.zip.${index}.tmp`);
    await fs.writeFile(filePath, `orphan-${index}`);
    const ageMs = index < 2 ? (index + 1) * 100 : 10_000 + index;
    const timestamp = new Date(now - ageMs);
    await fs.utimes(filePath, timestamp, timestamp);
  }

  const result = await hooks.cleanupOrphanTempFiles("submissions.zip", { maxAgeMs: 1000, maxFiles: 2, nowMs: now });
  assert.equal(result.scanned, 6);
  assert.equal(result.removed.length, 4);
  assert.equal(result.retained.length, 2);
  assert.equal(hash(await fs.readFile(validArchivePath)), hash(validArchive), "Orphan cleanup touched the valid archive.");
  assert.equal(await fs.readFile(path.join(cacheDir, "unrelated-user-file.tmp"), "utf8"), "preserve me");
  const remaining = (await fs.readdir(cacheDir)).filter((name) => /^submissions\.zip\..*\.tmp$/.test(name));
  assert.equal(remaining.length, 2, "Orphan retention did not enforce the configured count bound.");
}

async function testCrossInstanceRefreshDedupe(cacheDir, firstHooks, secondHooks) {
  const replacement = await archiveBuffer({ cik: 1, version: "cross-instance", facts: { ok: true } });
  let fetchCount = 0;
  let releaseFetch;
  let markFetchStarted;
  const fetchStarted = new Promise((resolve) => {
    markFetchStarted = resolve;
  });
  const fetchGate = new Promise((resolve) => {
    releaseFetch = resolve;
  });
  const fetchImpl = async () => {
    fetchCount += 1;
    markFetchStarted();
    await fetchGate;
    return new Response(replacement, { status: 200 });
  };

  const firstRefresh = firstHooks.refreshArchive("companyfacts", {}, { force: true, fetchImpl });
  await fetchStarted;
  const secondResult = await secondHooks.refreshArchive("companyfacts", {}, {
    force: true,
    fetchImpl: async () => {
      fetchCount += 1;
      return new Response(replacement, { status: 200 });
    }
  });
  assert.equal(secondResult, "busy", "A second module instance did not honor the shared archive lock.");
  releaseFetch();
  assert.equal(await firstRefresh, "refreshed");
  assert.equal(fetchCount, 1, "Cross-instance archive locking allowed a duplicate download.");
}

async function testAtomicMetadataMerge(cacheDir, firstHooks, secondHooks) {
  await Promise.all([
    firstHooks.writeMetadata("companyfacts", path.join(cacheDir, "companyfacts.zip"), "https://example.test/companyfacts.zip"),
    secondHooks.writeMetadata("submissions", path.join(cacheDir, "submissions.zip"), "https://example.test/submissions.zip")
  ]);
  const metadata = JSON.parse(await fs.readFile(path.join(cacheDir, "metadata.json"), "utf8"));
  assert.equal(metadata.archives.companyfacts.path, path.join(cacheDir, "companyfacts.zip"));
  assert.equal(metadata.archives.submissions.path, path.join(cacheDir, "submissions.zip"));
  assert.equal(await fileExists(path.join(cacheDir, "metadata.json.lock")), false);
  assert.deepEqual((await fs.readdir(cacheDir)).filter((name) => /^metadata\.json\..*\.tmp$/.test(name)), []);
}

async function testStaleLockRecovery(cacheDir, hooks) {
  const lockPath = path.join(cacheDir, "stale-test.lock");
  await fs.writeFile(lockPath, `${JSON.stringify({ token: "abandoned" })}\n`);
  const old = new Date(Date.now() - 60_000);
  await fs.utimes(lockPath, old, old);
  const lock = await hooks.acquireFileLock(lockPath, { waitMs: 0, staleMs: 1000, pollMs: 5 });
  assert.ok(lock, "A stale lock was not recovered.");
  assert.notEqual(lock.token, "abandoned");
  await lock.release();
  assert.equal(await fileExists(lockPath), false);
  assert.deepEqual((await fs.readdir(cacheDir)).filter((name) => name.includes("stale-test.lock") && name.endsWith(".stale")), []);
}

async function testBoundedFingerprintAwareZipCache(cacheDir, hooks) {
  for (let index = 0; index < 4; index += 1) {
    const filePath = path.join(cacheDir, `cache-${index}.zip`);
    await fs.writeFile(filePath, await archiveBuffer({ version: `cache-${index}` }, String(index + 10).padStart(10, "0")));
    await hooks.openZip(filePath);
  }
  assert.equal(hooks.zipDirectoryCacheSize(), 2, "ZIP directory cache exceeded its configured LRU bound.");

  const targetPath = path.join(cacheDir, "cache-3.zip");
  const cik = "0000000013";
  const updated = await archiveBuffer({ version: "cache-updated", padding: "x".repeat(200) }, cik);
  await fs.writeFile(targetPath, updated);
  const future = new Date(Date.now() + 5000);
  await fs.utimes(targetPath, future, future);
  const refreshedZip = await hooks.openZip(targetPath);
  const payload = JSON.parse(await refreshedZip.readText(`CIK${cik}.json`));
  assert.equal(payload.version, "cache-updated", "ZIP cache reused a directory after the file fingerprint changed.");
  assert.ok(hooks.zipDirectoryCacheSize() <= 2);
}

async function testStaleDirectoryRejectsReplacementArchive(cacheDir, hooks) {
  const archivePath = path.join(cacheDir, "companyfacts-race.zip");
  const cik = "0000000099";
  await fs.writeFile(archivePath, await archiveBuffer({ version: "old-index", facts: { old: true } }, cik));
  const staleDirectory = await hooks.openZip(archivePath);

  const replacementPath = path.join(cacheDir, "companyfacts-race.replacement.zip");
  await fs.writeFile(
    replacementPath,
    await archiveBuffer({ version: "new-archive", facts: { replacement: true }, padding: "x".repeat(500) }, cik)
  );
  await fs.rename(replacementPath, archivePath);

  await assert.rejects(
    staleDirectory.readText("CIK0000000100.json"),
    /changed or was replaced/i,
    "A stale directory was allowed to report a missing entry without validating the replacement archive."
  );
  await assert.rejects(
    staleDirectory.readText(`CIK${cik}.json`),
    /changed or was replaced/i,
    "A directory indexed from the old archive was allowed to read offsets from its replacement."
  );

  const currentDirectory = await hooks.openZip(archivePath);
  const payload = JSON.parse(await currentDirectory.readText(`CIK${cik}.json`));
  assert.equal(payload.version, "new-archive", "The ZIP reader did not recover with a fresh index after archive replacement.");
}

async function testIssuerIdentityValidationAndCacheInvalidation(cacheDir, hooks) {
  const requestedCik = "0000000001";
  const companyFactsPath = path.join(cacheDir, "companyfacts.zip");
  const submissionsPath = path.join(cacheDir, "submissions.zip");
  hooks.resetInMemoryState();

  await fs.writeFile(
    companyFactsPath,
    await archiveBuffer({ cik: 2, entityName: "Wrong issuer", facts: { "us-gaap": {} } }, requestedCik)
  );
  await fs.writeFile(
    submissionsPath,
    await archiveBuffer({ cik: "0000000002", name: "Wrong issuer", filings: { recent: {}, files: [] } }, requestedCik)
  );

  const swapped = await hooks.loadSecBulkSupport(requestedCik, {});
  assert.equal(swapped.companyFacts, null, "Swapped companyfacts content was accepted under the requested issuer.");
  assert.equal(swapped.submissions, null, "Swapped submissions content was accepted under the requested issuer.");
  assert.equal(
    swapped.companyFacts ? "sec_bulk_companyfacts" : "sec_live_companyfacts",
    "sec_live_companyfacts",
    "Rejected issuer content could still be stamped with bulk-companyfacts provenance downstream."
  );
  assert.equal(
    swapped.submissions ? "sec_bulk_submissions" : "sec_live_submissions",
    "sec_live_submissions",
    "Rejected issuer content could still be stamped with bulk-submissions provenance downstream."
  );
  assert.match(swapped.warnings.join("\n"), /companyfacts payload declared CIK 0000000002.*CIK 0000000001 was requested/i);
  assert.match(swapped.warnings.join("\n"), /submissions payload declared CIK 0000000002.*CIK 0000000001 was requested/i);
  assert.equal(
    hooks.zipDirectoryCacheSize(),
    0,
    "Issuer-integrity failures did not invalidate the cached ZIP directory before live fallback."
  );

  await fs.writeFile(
    companyFactsPath,
    await archiveBuffer({ cik: 1, entityName: "Correct issuer", facts: { "us-gaap": {} } }, requestedCik)
  );
  await fs.writeFile(
    submissionsPath,
    await archiveBuffer({ cik: requestedCik, name: "Correct issuer", filings: { recent: {}, files: [] } }, requestedCik)
  );
  const valid = await hooks.loadSecBulkSupport("1", {});
  assert.equal(valid.companyFacts.entityName, "Correct issuer", "Numeric companyfacts CIK normalization rejected the correct issuer.");
  assert.equal(valid.submissions.name, "Correct issuer", "Zero-padded submissions CIK normalization rejected the correct issuer.");
  assert.equal(valid.companyFacts ? "sec_bulk_companyfacts" : "sec_live_companyfacts", "sec_bulk_companyfacts");
  assert.equal(valid.submissions ? "sec_bulk_submissions" : "sec_live_submissions", "sec_bulk_submissions");

  hooks.resetInMemoryState();
  await fs.writeFile(companyFactsPath, await archiveBuffer({ entityName: "Missing CIK", facts: {} }, requestedCik));
  await fs.writeFile(submissionsPath, await archiveBuffer({ cik: "not-a-cik", filings: {} }, requestedCik));
  const malformed = await hooks.loadSecBulkSupport(requestedCik, {});
  assert.equal(malformed.companyFacts, null, "A companyfacts root with no declared CIK was accepted.");
  assert.equal(malformed.submissions, null, "A submissions root with a malformed declared CIK was accepted.");
  assert.match(malformed.warnings.join("\n"), /companyfacts payload did not declare a CIK/i);
  assert.match(malformed.warnings.join("\n"), /submissions payload declared an invalid CIK/i);
  assert.equal(hooks.zipDirectoryCacheSize(), 0, "Malformed issuer roots remained in the ZIP directory cache.");
}

async function testSubmissionHistoryPathBinding(cacheDir, hooks) {
  const requestedCik = "0000000001";
  const submissionsPath = path.join(cacheDir, "submissions.zip");
  const historyName = `CIK${requestedCik}-submissions-001.json`;
  hooks.resetInMemoryState();
  await fs.writeFile(
    submissionsPath,
    await archiveEntriesBuffer({
      [`CIK${requestedCik}.json`]: { cik: requestedCik, filings: { recent: {}, files: [{ name: historyName }] } },
      [historyName]: { accessionNumber: ["0000000001-00-000001"], form: ["10-K"] }
    })
  );

  const validHistory = await hooks.readSecBulkSubmissionFile(historyName, {}, requestedCik);
  assert.deepEqual(validHistory.form, ["10-K"], "A correctly issuer-bound SEC submissions history file was rejected.");
  const wrongIssuerHistory = await hooks.readSecBulkSubmissionFile(historyName, {}, "0000000002");
  assert.equal(wrongIssuerHistory, null, "A submissions history path was accepted for a different requested issuer.");
  assert.equal(hooks.zipDirectoryCacheSize(), 0, "A mismatched submissions history path did not invalidate the ZIP directory cache.");
}

async function testSwappedDownloadCannotReplaceValidCache(cacheDir, hooks) {
  const archivePath = path.join(cacheDir, "companyfacts.zip");
  const requestedCik = "0000000001";
  const original = await archiveBuffer({ cik: 1, version: "known-good", facts: {} }, requestedCik);
  const swapped = await archiveBuffer({ cik: 2, version: "swapped", facts: {} }, requestedCik);
  await fs.writeFile(archivePath, original);

  await assert.rejects(
    hooks.refreshArchive("companyfacts", {}, {
      force: true,
      fetchImpl: async () => new Response(swapped, { status: 200 })
    }),
    /issuer identity/i,
    "A downloaded archive whose entry name and payload CIK disagreed was accepted."
  );
  assert.equal(hash(await fs.readFile(archivePath)), hash(original), "A swapped-CIK download replaced the last known-good bulk cache.");
  await assertNoOwnedTempFiles(cacheDir, "companyfacts.zip");
}

async function fileExists(filePath) {
  try {
    await fs.access(filePath);
    return true;
  } catch {
    return false;
  }
}

async function main() {
  const cacheDir = await fs.mkdtemp(path.join(os.tmpdir(), "sec-bulk-cache-reliability-"));
  const firstHooks = loadHooks(cacheDir);
  const secondHooks = loadHooks(cacheDir);
  try {
    await testFailedDownloadPreservesArchiveAndCleansTemp(cacheDir, firstHooks);
    await testSuccessfulValidatedReplacement(cacheDir, firstHooks);
    await testOrphanRetention(cacheDir, firstHooks);
    await testCrossInstanceRefreshDedupe(cacheDir, firstHooks, secondHooks);
    await testAtomicMetadataMerge(cacheDir, firstHooks, secondHooks);
    await testStaleLockRecovery(cacheDir, firstHooks);
    await testBoundedFingerprintAwareZipCache(cacheDir, firstHooks);
    await testStaleDirectoryRejectsReplacementArchive(cacheDir, firstHooks);
    await testIssuerIdentityValidationAndCacheInvalidation(cacheDir, firstHooks);
    await testSubmissionHistoryPathBinding(cacheDir, firstHooks);
    await testSwappedDownloadCannotReplaceValidCache(cacheDir, firstHooks);
    console.log("SEC bulk cache cleanup, validation, locking, metadata, identity, and bounded-cache guards passed.");
  } finally {
    await fs.rm(cacheDir, { recursive: true, force: true });
  }
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
