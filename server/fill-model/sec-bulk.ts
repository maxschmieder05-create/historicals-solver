import { createWriteStream } from "node:fs";
import fs from "node:fs/promises";
import path from "node:path";
import { Readable } from "node:stream";
import { pipeline } from "node:stream/promises";
import { inflateRawSync } from "node:zlib";

type SecBulkArchiveKind = "companyfacts" | "submissions";

type SecBulkArchiveConfig = {
  filename: string;
  url: string;
  configuredPathEnv: string;
};

export type SecBulkArchiveStatus = {
  path?: string;
  source: "cache" | "configured_file" | "missing";
  lastRefreshAt?: string;
  stale: boolean;
  warning?: string;
};

export type SecBulkSupport = {
  companyFacts?: any;
  submissions?: any;
  archives: Partial<Record<SecBulkArchiveKind, SecBulkArchiveStatus>>;
  warnings: string[];
  latestRefreshAt?: string;
};

type ZipEntry = {
  name: string;
  compressionMethod: number;
  compressedSize: number;
  localHeaderOffset: number;
};

type SecBulkMetadata = {
  archives?: Partial<Record<SecBulkArchiveKind, { lastRefreshAt: string; path: string; sourceUrl: string }>>;
};

type ZipDirectoryCacheEntry = {
  fingerprint: string;
  promise: Promise<ZipDirectory>;
};

type FileLockHandle = {
  lockPath: string;
  token: string;
  release: () => Promise<void>;
};

type FileLockOptions = {
  waitMs?: number;
  staleMs?: number;
  pollMs?: number;
};

type OrphanCleanupOptions = {
  maxAgeMs?: number;
  maxFiles?: number;
  nowMs?: number;
  excludePaths?: string[];
};

type OrphanCleanupResult = {
  scanned: number;
  removed: string[];
  retained: string[];
};

type RefreshArchiveOptions = {
  force?: boolean;
  fetchImpl?: typeof fetch;
};

type BulkJsonReadExpectation =
  | {
      type: "issuer_root";
      expectedCik: string;
    }
  | {
      type: "submission_history";
      expectedCik: string;
    };

const SEC_BULK_REFRESH_INTERVAL_MS = Number(process.env.SEC_BULK_REFRESH_INTERVAL_MS || 24 * 60 * 60 * 1000);
const SEC_BULK_CACHE_DIR = process.env.SEC_BULK_CACHE_DIR || path.join(process.cwd(), "tmp", "sec-bulk");
const SEC_BULK_METADATA_FILE = path.join(SEC_BULK_CACHE_DIR, "metadata.json");
const SEC_BULK_DISABLED = process.env.SEC_BULK_DISABLED === "true" || process.env.SEC_BULK_DISABLED === "1";
const SEC_BULK_REFRESH_DISABLED = process.env.SEC_BULK_REFRESH_DISABLED === "true" || process.env.SEC_BULK_REFRESH_DISABLED === "1";
const SEC_BULK_ORPHAN_MAX_AGE_MS = positiveEnvNumber(process.env.SEC_BULK_ORPHAN_MAX_AGE_MS, 24 * 60 * 60 * 1000);
const SEC_BULK_ORPHAN_MAX_FILES = nonNegativeEnvInteger(process.env.SEC_BULK_ORPHAN_MAX_FILES, 4);
const SEC_BULK_REFRESH_LOCK_STALE_MS = positiveEnvNumber(process.env.SEC_BULK_REFRESH_LOCK_STALE_MS, 24 * 60 * 60 * 1000);
const SEC_BULK_LOCK_POLL_MS = positiveEnvNumber(process.env.SEC_BULK_LOCK_POLL_MS, 100);
const SEC_BULK_METADATA_LOCK_WAIT_MS = positiveEnvNumber(process.env.SEC_BULK_METADATA_LOCK_WAIT_MS, 15_000);
const SEC_BULK_ZIP_DIRECTORY_CACHE_MAX_ENTRIES = positiveEnvInteger(process.env.SEC_BULK_ZIP_DIRECTORY_CACHE_MAX_ENTRIES, 4);

const ARCHIVES: Record<SecBulkArchiveKind, SecBulkArchiveConfig> = {
  companyfacts: {
    filename: "companyfacts.zip",
    url: "https://www.sec.gov/Archives/edgar/daily-index/xbrl/companyfacts.zip",
    configuredPathEnv: "SEC_BULK_COMPANYFACTS_ZIP"
  },
  submissions: {
    filename: "submissions.zip",
    url: "https://www.sec.gov/Archives/edgar/daily-index/bulkdata/submissions.zip",
    configuredPathEnv: "SEC_BULK_SUBMISSIONS_ZIP"
  }
};

const zipDirectoryCache = new Map<string, ZipDirectoryCacheEntry>();
const refreshPromises = new Map<SecBulkArchiveKind, Promise<void>>();
let metadataWriteQueue = Promise.resolve();

export async function loadSecBulkSupport(cik: string, headers: Record<string, string>): Promise<SecBulkSupport> {
  if (SEC_BULK_DISABLED) {
    return {
      archives: {},
      warnings: ["SEC bulk support is disabled by SEC_BULK_DISABLED; using live SEC APIs only."]
    };
  }

  const expectedCik = requireNormalizedSecCik(cik, "requested SEC issuer CIK");
  const cikFile = `CIK${expectedCik}.json`;
  const [companyFacts, submissions] = await Promise.all([
    readBulkJson("companyfacts", cikFile, headers, { type: "issuer_root", expectedCik }),
    readBulkJson("submissions", cikFile, headers, { type: "issuer_root", expectedCik })
  ]);

  const statuses = [companyFacts.status, submissions.status].filter(Boolean) as SecBulkArchiveStatus[];
  const latestRefreshAt = latestTimestamp(statuses.map((status) => status.lastRefreshAt).filter(Boolean) as string[]);
  const warnings = unique([companyFacts.warning, submissions.warning, companyFacts.status?.warning, submissions.status?.warning].filter(Boolean) as string[]);

  return {
    companyFacts: companyFacts.payload,
    submissions: submissions.payload,
    archives: {
      companyfacts: companyFacts.status,
      submissions: submissions.status
    },
    warnings,
    latestRefreshAt
  };
}

export async function readSecBulkSubmissionFile(fileName: string, headers: Record<string, string>, expectedCik: string) {
  const result = await readBulkJson("submissions", fileName, headers, {
    type: "submission_history",
    expectedCik: requireNormalizedSecCik(expectedCik, "requested SEC issuer CIK")
  });
  return result.payload ?? null;
}

async function readBulkJson(
  kind: SecBulkArchiveKind,
  fileName: string,
  headers: Record<string, string>,
  expectation: BulkJsonReadExpectation
) {
  const status = await ensureArchive(kind, headers);
  if (!status.path) return { payload: null, status, warning: status.warning };

  try {
    assertBulkEntryNameMatchesExpectation(kind, fileName, expectation);
    const text = await readCurrentZipText(status.path, fileName);
    if (!text) return { payload: null, status, warning: `${ARCHIVES[kind].filename} did not contain ${fileName}; live SEC API fallback will be used.` };
    const payload = JSON.parse(text);
    validateSecBulkPayload(kind, payload, expectation.expectedCik, {
      requireDeclaredCik: expectation.type === "issuer_root",
      requireIssuerRootShape: expectation.type === "issuer_root"
    });
    return { payload, status, warning: undefined };
  } catch (error) {
    invalidateZipDirectoryCache(status.path);
    if (status.source === "cache") scheduleRefresh(kind, headers, true);
    const warning = `Could not read ${fileName} from ${ARCHIVES[kind].filename}: ${error instanceof Error ? error.message : "unknown error"}. Live SEC API fallback will be used.`;
    return { payload: null, status, warning };
  }
}

async function readCurrentZipText(zipPath: string, name: string) {
  let identityError: ZipArchiveIdentityChangedError | null = null;
  for (let attempt = 0; attempt < 2; attempt += 1) {
    try {
      const zip = await openZip(zipPath);
      return await zip.readText(name);
    } catch (error) {
      if (!(error instanceof ZipArchiveIdentityChangedError)) throw error;
      identityError = error;
      invalidateZipDirectoryCache(zipPath);
    }
  }
  throw identityError ?? new Error(`Could not read ${name} from ${path.basename(zipPath)}.`);
}

async function ensureArchive(kind: SecBulkArchiveKind, headers: Record<string, string>): Promise<SecBulkArchiveStatus> {
  const config = ARCHIVES[kind];
  const configuredPath = process.env[config.configuredPathEnv];
  if (configuredPath && (await fileExists(configuredPath))) {
    const stat = await fs.stat(configuredPath);
    return {
      path: configuredPath,
      source: "configured_file",
      lastRefreshAt: stat.mtime.toISOString(),
      stale: isStale(stat.mtimeMs),
      warning: isStale(stat.mtimeMs) ? `${config.configuredPathEnv} is older than the SEC bulk refresh interval; live SEC API fallback remains enabled for recent filings.` : undefined
    };
  }

  const archivePath = path.join(SEC_BULK_CACHE_DIR, config.filename);
  const exists = await fileExists(archivePath);
  if (exists) {
    const stat = await fs.stat(archivePath);
    const metadata = await readMetadata();
    const metadataRefreshAt = metadata.archives?.[kind]?.lastRefreshAt;
    const metadataRefreshMs = metadataRefreshAt ? Date.parse(metadataRefreshAt) : Number.NaN;
    const refreshAt = new Date(Math.max(stat.mtimeMs, Number.isFinite(metadataRefreshMs) ? metadataRefreshMs : 0));
    if (isStale(refreshAt.getTime())) scheduleRefresh(kind, headers);
    return {
      path: archivePath,
      source: "cache",
      lastRefreshAt: refreshAt.toISOString(),
      stale: isStale(refreshAt.getTime()),
      warning: isStale(refreshAt.getTime()) ? `${config.filename} cache is stale; a background SEC bulk refresh was started and live SEC API fallback remains enabled.` : undefined
    };
  }

  scheduleRefresh(kind, headers);
  return {
    source: "missing",
    stale: true,
    warning: `${config.filename} cache is not available yet; a background SEC bulk download was started and live SEC API fallback will be used for this request.`
  };
}

function scheduleRefresh(kind: SecBulkArchiveKind, headers: Record<string, string>, force = false) {
  if (SEC_BULK_REFRESH_DISABLED || refreshPromises.has(kind)) return;
  const promise = refreshArchive(kind, headers, { force })
    .then(() => undefined)
    .catch(() => {
      // Existing archives and live SEC APIs are still usable if a background refresh fails.
    })
    .finally(() => {
      refreshPromises.delete(kind);
    });
  refreshPromises.set(kind, promise);
}

async function refreshArchive(kind: SecBulkArchiveKind, headers: Record<string, string>, options: RefreshArchiveOptions = {}) {
  const config = ARCHIVES[kind];
  await fs.mkdir(SEC_BULK_CACHE_DIR, { recursive: true });
  const archivePath = path.join(SEC_BULK_CACHE_DIR, config.filename);
  const refreshLock = await acquireFileLock(`${archivePath}.refresh.lock`, {
    waitMs: 0,
    staleMs: SEC_BULK_REFRESH_LOCK_STALE_MS,
    pollMs: SEC_BULK_LOCK_POLL_MS
  });
  if (!refreshLock) return "busy" as const;

  let tempPath: string | null = null;
  try {
    if (!options.force && !(await archiveNeedsRefresh(kind, archivePath))) return "fresh" as const;
    await cleanupOrphanTempFiles(config.filename);

    tempPath = uniqueTempPath(archivePath);
    const response = await (options.fetchImpl ?? fetch)(config.url, { headers });
    if (!response.ok || !response.body) {
      throw new Error(`SEC bulk download failed for ${config.filename}: ${response.status} ${response.statusText}`);
    }
    await pipeline(Readable.fromWeb(response.body as any), createWriteStream(tempPath, { flags: "wx" }));
    await validateSecBulkArchive(tempPath, kind);

    invalidateZipDirectoryCache(archivePath);
    await fs.rename(tempPath, archivePath);
    tempPath = null;
    invalidateZipDirectoryCache(archivePath);
    await writeMetadata(kind, archivePath, config.url);
    return "refreshed" as const;
  } finally {
    if (tempPath) await removeFileIfPresent(tempPath);
    await refreshLock.release();
  }
}

async function readMetadata(): Promise<SecBulkMetadata> {
  try {
    return JSON.parse(await fs.readFile(SEC_BULK_METADATA_FILE, "utf8")) as SecBulkMetadata;
  } catch {
    return {};
  }
}

function writeMetadata(kind: SecBulkArchiveKind, archivePath: string, sourceUrl: string) {
  const operation = metadataWriteQueue.catch(() => undefined).then(async () => {
    await fs.mkdir(SEC_BULK_CACHE_DIR, { recursive: true });
    const lock = await acquireFileLock(`${SEC_BULK_METADATA_FILE}.lock`, {
      waitMs: SEC_BULK_METADATA_LOCK_WAIT_MS,
      staleMs: SEC_BULK_REFRESH_LOCK_STALE_MS,
      pollMs: SEC_BULK_LOCK_POLL_MS
    });
    if (!lock) throw new Error("Timed out waiting for the SEC bulk metadata lock.");
    try {
      const metadata = await readMetadata();
      metadata.archives = {
        ...(metadata.archives ?? {}),
        [kind]: {
          lastRefreshAt: new Date().toISOString(),
          path: archivePath,
          sourceUrl
        }
      };
      await writeFileAtomically(SEC_BULK_METADATA_FILE, `${JSON.stringify(metadata, null, 2)}\n`);
    } finally {
      await lock.release();
    }
  });
  metadataWriteQueue = operation.then(() => undefined, () => undefined);
  return operation;
}

async function openZip(zipPath: string) {
  const stat = await fs.stat(zipPath);
  const fingerprint = zipFileFingerprint(stat);
  const cached = zipDirectoryCache.get(zipPath);
  if (cached?.fingerprint === fingerprint) {
    touchZipDirectoryCacheEntry(zipPath, cached);
    return cached.promise;
  }
  if (cached) zipDirectoryCache.delete(zipPath);

  const entry: ZipDirectoryCacheEntry = {
    fingerprint,
    promise: ZipDirectory.open(zipPath)
  };
  zipDirectoryCache.set(zipPath, entry);
  trimZipDirectoryCache();
  try {
    return await entry.promise;
  } catch (error) {
    if (zipDirectoryCache.get(zipPath) === entry) zipDirectoryCache.delete(zipPath);
    throw error;
  }
}

class ZipDirectory {
  private constructor(
    private readonly zipPath: string,
    private readonly archiveFingerprint: string,
    private readonly entries: Map<string, ZipEntry>
  ) {}

  static async open(zipPath: string) {
    const handle = await fs.open(zipPath, "r");
    try {
      const stat = await handle.stat();
      const archiveFingerprint = zipFileFingerprint(stat);
      if (stat.size < 22) throw new Error("ZIP file is too small to contain an end-of-central-directory record.");
      const tailLength = Math.min(stat.size, 66_000);
      const tail = Buffer.alloc(tailLength);
      await readExactly(handle, tail, stat.size - tailLength);
      const eocdOffset = findEndOfCentralDirectory(tail);
      if (eocdOffset < 0) throw new Error("ZIP central directory was not found.");

      const rawEntryCount = tail.readUInt16LE(eocdOffset + 10);
      const expectedEntryCount = rawEntryCount === 0xffff ? null : rawEntryCount;
      const centralDirectorySize = tail.readUInt32LE(eocdOffset + 12);
      const centralDirectoryOffset = tail.readUInt32LE(eocdOffset + 16);
      if (centralDirectorySize === 0xffffffff || centralDirectoryOffset === 0xffffffff) {
        throw new Error("ZIP64 archives are not supported by the lightweight SEC bulk reader.");
      }
      if (centralDirectoryOffset > stat.size || centralDirectorySize > stat.size - centralDirectoryOffset) {
        throw new Error("ZIP central directory points outside the archive.");
      }

      const centralDirectory = Buffer.alloc(centralDirectorySize);
      await readExactly(handle, centralDirectory, centralDirectoryOffset);
      assertZipArchiveIdentity(zipPath, archiveFingerprint, await handle.stat());
      return new ZipDirectory(zipPath, archiveFingerprint, parseCentralDirectory(centralDirectory, expectedEntryCount, stat.size));
    } finally {
      await handle.close();
    }
  }

  async readText(name: string) {
    const handle = await fs.open(this.zipPath, "r");
    try {
      const stat = await handle.stat();
      assertZipArchiveIdentity(this.zipPath, this.archiveFingerprint, stat);
      const entry = this.entries.get(name);
      if (!entry) return null;
      const localHeader = Buffer.alloc(30);
      await readExactly(handle, localHeader, entry.localHeaderOffset);
      if (localHeader.readUInt32LE(0) !== 0x04034b50) throw new Error(`Invalid local ZIP header for ${name}.`);
      const localCompressionMethod = localHeader.readUInt16LE(8);
      if (localCompressionMethod !== entry.compressionMethod) {
        throw new Error(`Local ZIP compression method did not match the directory entry for ${name}.`);
      }
      const fileNameLength = localHeader.readUInt16LE(26);
      const extraLength = localHeader.readUInt16LE(28);
      const localFileName = Buffer.alloc(fileNameLength);
      await readExactly(handle, localFileName, entry.localHeaderOffset + 30);
      if (localFileName.toString("utf8") !== entry.name) {
        throw new Error(`Local ZIP file name did not match the directory entry for ${name}.`);
      }
      const dataOffset = entry.localHeaderOffset + 30 + fileNameLength + extraLength;
      if (dataOffset > stat.size || entry.compressedSize > stat.size - dataOffset) {
        throw new Error(`Compressed ZIP data for ${name} points outside the archive.`);
      }
      const compressed = Buffer.alloc(entry.compressedSize);
      await readExactly(handle, compressed, dataOffset);
      assertZipArchiveIdentity(this.zipPath, this.archiveFingerprint, await handle.stat());
      if (entry.compressionMethod === 0) return compressed.toString("utf8");
      if (entry.compressionMethod === 8) return inflateRawSync(compressed).toString("utf8");
      throw new Error(`Unsupported ZIP compression method ${entry.compressionMethod} for ${name}.`);
    } finally {
      await handle.close();
    }
  }

  validationEntryNames(pattern: RegExp, preferredNames: string[] = [], limit = 20) {
    const matching = Array.from(this.entries.values()).filter((entry) => pattern.test(entry.name));
    const preferred = preferredNames.filter((name) => this.entries.has(name));
    const smallestNonempty = matching
      .filter((entry) => entry.compressedSize >= 16)
      .sort((a, b) => a.compressedSize - b.compressedSize || a.name.localeCompare(b.name))
      .map((entry) => entry.name);
    const fallback = matching.sort((a, b) => a.name.localeCompare(b.name)).map((entry) => entry.name);
    return Array.from(new Set([...preferred, ...smallestNonempty, ...fallback])).slice(0, Math.max(1, limit));
  }

  get entryCount() {
    return this.entries.size;
  }
}

class ZipArchiveIdentityChangedError extends Error {
  constructor(zipPath: string) {
    super(`${path.basename(zipPath)} changed or was replaced while its ZIP directory index was in use.`);
    this.name = "ZipArchiveIdentityChangedError";
  }
}

function parseCentralDirectory(buffer: Buffer, expectedEntryCount: number | null, archiveSize: number) {
  const entries = new Map<string, ZipEntry>();
  let offset = 0;
  let index = 0;
  while (offset + 46 <= buffer.length && (expectedEntryCount === null || index < expectedEntryCount)) {
    if (buffer.readUInt32LE(offset) !== 0x02014b50) {
      if (expectedEntryCount === null) break;
      throw new Error(`ZIP central directory entry ${index + 1} is invalid or truncated.`);
    }
    const compressionMethod = buffer.readUInt16LE(offset + 10);
    const compressedSize = buffer.readUInt32LE(offset + 20);
    const fileNameLength = buffer.readUInt16LE(offset + 28);
    const extraLength = buffer.readUInt16LE(offset + 30);
    const commentLength = buffer.readUInt16LE(offset + 32);
    const localHeaderOffset = buffer.readUInt32LE(offset + 42);
    const nextOffset = offset + 46 + fileNameLength + extraLength + commentLength;
    if (nextOffset > buffer.length) throw new Error(`ZIP central directory entry ${index + 1} exceeds the directory bounds.`);
    if (localHeaderOffset >= archiveSize) throw new Error(`ZIP central directory entry ${index + 1} points outside the archive.`);
    const name = buffer.subarray(offset + 46, offset + 46 + fileNameLength).toString("utf8");
    entries.set(name, { name, compressionMethod, compressedSize, localHeaderOffset });
    offset = nextOffset;
    index += 1;
  }
  if (expectedEntryCount !== null && index !== expectedEntryCount) {
    throw new Error(`ZIP central directory ended after ${index} of ${expectedEntryCount} expected entries.`);
  }
  return entries;
}

async function validateSecBulkArchive(filePath: string, kind: SecBulkArchiveKind) {
  const zip = await ZipDirectory.open(filePath);
  if (!zip.entryCount) throw new Error(`${ARCHIVES[kind].filename} contained no ZIP entries.`);
  const sampleNames = zip.validationEntryNames(/^CIK\d{10}\.json$/, ["CIK0000320193.json", "CIK0000789019.json"]);
  if (!sampleNames.length) throw new Error(`${ARCHIVES[kind].filename} did not contain a CIK JSON entry.`);
  for (const sampleName of sampleNames) {
    const sampleText = await zip.readText(sampleName);
    if (!sampleText) continue;
    try {
      const sample = JSON.parse(sampleText);
      const expectedCik = rootCikFromEntryName(sampleName);
      if (!expectedCik) continue;
      validateSecBulkPayload(kind, sample, expectedCik, {
        requireDeclaredCik: true,
        requireIssuerRootShape: true
      });
      return;
    } catch {
      // Try another small CIK entry before rejecting the complete archive.
    }
  }
  throw new Error(`${ARCHIVES[kind].filename} CIK entries did not match the expected ${kind} payload shape and issuer identity.`);
}

function validateSecBulkPayload(
  kind: SecBulkArchiveKind,
  payload: unknown,
  expectedCik: string,
  options: { requireDeclaredCik: boolean; requireIssuerRootShape: boolean }
) {
  if (!payload || typeof payload !== "object" || Array.isArray(payload)) {
    throw new SecBulkIssuerIdentityError(`${kind} payload was not a JSON object.`);
  }
  const record = payload as Record<string, unknown>;
  if (options.requireIssuerRootShape) {
    const expectedObject = kind === "companyfacts" ? record.facts : record.filings;
    if (!expectedObject || typeof expectedObject !== "object" || Array.isArray(expectedObject)) {
      throw new SecBulkIssuerIdentityError(`${kind} payload did not contain its required ${kind === "companyfacts" ? "facts" : "filings"} object.`);
    }
  }

  if (!("cik" in record)) {
    if (options.requireDeclaredCik) throw new SecBulkIssuerIdentityError(`${kind} payload did not declare a CIK.`);
    return;
  }
  const declaredCik = normalizeSecCik(record.cik);
  if (!declaredCik) throw new SecBulkIssuerIdentityError(`${kind} payload declared an invalid CIK.`);
  if (declaredCik !== expectedCik) {
    throw new SecBulkIssuerIdentityError(`${kind} payload declared CIK ${declaredCik}, but CIK ${expectedCik} was requested.`);
  }
}

function assertBulkEntryNameMatchesExpectation(
  kind: SecBulkArchiveKind,
  fileName: string,
  expectation: BulkJsonReadExpectation
) {
  if (expectation.type === "issuer_root") {
    const entryCik = rootCikFromEntryName(fileName);
    if (entryCik !== expectation.expectedCik) {
      throw new SecBulkIssuerIdentityError(
        `${kind} entry ${fileName} did not belong to requested CIK ${expectation.expectedCik}.`
      );
    }
    return;
  }
  const entryCik = submissionHistoryCikFromEntryName(fileName);
  if (kind !== "submissions" || entryCik !== expectation.expectedCik) {
    throw new SecBulkIssuerIdentityError(
      `submissions history entry ${fileName} did not belong to requested CIK ${expectation.expectedCik}.`
    );
  }
}

function rootCikFromEntryName(fileName: string) {
  const match = fileName.match(/^CIK(\d{10})\.json$/);
  return match ? normalizeSecCik(match[1]) : null;
}

function submissionHistoryCikFromEntryName(fileName: string) {
  const match = fileName.match(/^CIK(\d{10})-submissions-\d+\.json$/);
  return match ? normalizeSecCik(match[1]) : null;
}

function requireNormalizedSecCik(value: unknown, label: string) {
  const normalized = normalizeSecCik(value);
  if (!normalized) throw new SecBulkIssuerIdentityError(`${label} was invalid.`);
  return normalized;
}

function normalizeSecCik(value: unknown) {
  let digits: string;
  if (typeof value === "number") {
    if (!Number.isSafeInteger(value) || value <= 0) return null;
    digits = String(value);
  } else if (typeof value === "string") {
    digits = value.trim();
    if (!/^\d{1,10}$/.test(digits)) return null;
  } else {
    return null;
  }
  const significant = digits.replace(/^0+/, "");
  if (!significant || significant.length > 10) return null;
  return significant.padStart(10, "0");
}

class SecBulkIssuerIdentityError extends Error {
  constructor(message: string) {
    super(`SEC bulk issuer identity validation failed: ${message}`);
    this.name = "SecBulkIssuerIdentityError";
  }
}

async function readExactly(handle: Awaited<ReturnType<typeof fs.open>>, buffer: Buffer, position: number) {
  let offset = 0;
  while (offset < buffer.length) {
    const { bytesRead } = await handle.read(buffer, offset, buffer.length - offset, position + offset);
    if (!bytesRead) throw new Error("Unexpected end of file while reading ZIP data.");
    offset += bytesRead;
  }
}

async function archiveNeedsRefresh(kind: SecBulkArchiveKind, archivePath: string) {
  if (!(await fileExists(archivePath))) return true;
  const stat = await fs.stat(archivePath);
  const metadata = await readMetadata();
  const metadataRefreshAt = metadata.archives?.[kind]?.lastRefreshAt;
  const metadataRefreshMs = metadataRefreshAt ? Date.parse(metadataRefreshAt) : Number.NaN;
  const refreshMs = Math.max(stat.mtimeMs, Number.isFinite(metadataRefreshMs) ? metadataRefreshMs : 0);
  return isStale(refreshMs);
}

async function cleanupOrphanTempFiles(filename: string, options: OrphanCleanupOptions = {}): Promise<OrphanCleanupResult> {
  const maxAgeMs = nonNegativeNumber(options.maxAgeMs, SEC_BULK_ORPHAN_MAX_AGE_MS);
  const maxFiles = nonNegativeInteger(options.maxFiles, SEC_BULK_ORPHAN_MAX_FILES);
  const nowMs = options.nowMs ?? Date.now();
  const excluded = new Set((options.excludePaths ?? []).map((item) => path.resolve(item)));
  let names: string[] = [];
  try {
    names = await fs.readdir(SEC_BULK_CACHE_DIR);
  } catch {
    return { scanned: 0, removed: [], retained: [] };
  }

  const prefix = `${filename}.`;
  const candidates = (
    await Promise.all(
      names
        .filter((name) => name.startsWith(prefix) && name.endsWith(".tmp"))
        .map(async (name) => {
          const filePath = path.join(SEC_BULK_CACHE_DIR, name);
          if (excluded.has(path.resolve(filePath))) return null;
          try {
            const stat = await fs.stat(filePath);
            return stat.isFile() ? { filePath, mtimeMs: stat.mtimeMs } : null;
          } catch {
            return null;
          }
        })
    )
  )
    .filter((item): item is { filePath: string; mtimeMs: number } => Boolean(item))
    .sort((a, b) => b.mtimeMs - a.mtimeMs || a.filePath.localeCompare(b.filePath));

  const retainedByCount = new Set(candidates.slice(0, maxFiles).map((item) => item.filePath));
  const removed: string[] = [];
  const retained: string[] = [];
  for (const candidate of candidates) {
    const tooOld = nowMs - candidate.mtimeMs > maxAgeMs;
    if (tooOld || !retainedByCount.has(candidate.filePath)) {
      if (await removeFileIfPresent(candidate.filePath)) removed.push(candidate.filePath);
    } else {
      retained.push(candidate.filePath);
    }
  }
  return { scanned: candidates.length, removed, retained };
}

async function acquireFileLock(lockPath: string, options: FileLockOptions = {}): Promise<FileLockHandle | null> {
  const waitMs = nonNegativeNumber(options.waitMs, 0);
  const staleMs = positiveNumber(options.staleMs, SEC_BULK_REFRESH_LOCK_STALE_MS);
  const pollMs = positiveNumber(options.pollMs, SEC_BULK_LOCK_POLL_MS);
  const deadline = Date.now() + waitMs;
  await fs.mkdir(path.dirname(lockPath), { recursive: true });

  while (true) {
    await quarantineStaleLock(lockPath, staleMs);
    const token = `${process.pid}-${Date.now()}-${Math.random().toString(36).slice(2, 10)}`;
    try {
      const handle = await fs.open(lockPath, "wx", 0o600);
      try {
        await handle.writeFile(`${JSON.stringify({ token, pid: process.pid, createdAt: new Date().toISOString() })}\n`, "utf8");
        await handle.sync();
      } catch (error) {
        await handle.close().catch(() => undefined);
        await removeFileIfPresent(lockPath);
        throw error;
      }
      await handle.close();
      let released = false;
      return {
        lockPath,
        token,
        release: async () => {
          if (released) return;
          released = true;
          const currentToken = await lockToken(lockPath);
          if (currentToken !== token) return;
          const releasePath = `${lockPath}.${token}.release`;
          try {
            await fs.rename(lockPath, releasePath);
            await removeFileIfPresent(releasePath);
          } catch (error) {
            if (!isFileSystemError(error, "ENOENT")) throw error;
          }
        }
      };
    } catch (error) {
      if (!isFileSystemError(error, "EEXIST")) throw error;
      if (Date.now() >= deadline) return null;
      await sleep(Math.min(pollMs, Math.max(1, deadline - Date.now())));
    }
  }
}

async function quarantineStaleLock(lockPath: string, staleMs: number) {
  let stat;
  try {
    stat = await fs.stat(lockPath);
  } catch (error) {
    if (isFileSystemError(error, "ENOENT")) return false;
    throw error;
  }
  if (Date.now() - stat.mtimeMs <= staleMs) return false;
  const quarantinePath = `${lockPath}.${process.pid}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}.stale`;
  try {
    await fs.rename(lockPath, quarantinePath);
    await removeFileIfPresent(quarantinePath);
    return true;
  } catch (error) {
    if (isFileSystemError(error, "ENOENT")) return false;
    throw error;
  }
}

async function lockToken(lockPath: string) {
  try {
    const parsed = JSON.parse(await fs.readFile(lockPath, "utf8"));
    return typeof parsed?.token === "string" ? parsed.token : null;
  } catch {
    return null;
  }
}

async function writeFileAtomically(filePath: string, contents: string | Buffer) {
  await fs.mkdir(path.dirname(filePath), { recursive: true });
  const tempPath = uniqueTempPath(filePath);
  let handle: Awaited<ReturnType<typeof fs.open>> | null = null;
  try {
    handle = await fs.open(tempPath, "wx", 0o600);
    await handle.writeFile(contents);
    await handle.sync();
    await handle.close();
    handle = null;
    await fs.rename(tempPath, filePath);
  } finally {
    if (handle) await handle.close().catch(() => undefined);
    await removeFileIfPresent(tempPath);
  }
}

function uniqueTempPath(filePath: string) {
  return `${filePath}.${process.pid}-${Date.now()}-${Math.random().toString(36).slice(2, 10)}.tmp`;
}

async function removeFileIfPresent(filePath: string) {
  try {
    await fs.rm(filePath, { force: true });
    return true;
  } catch (error) {
    if (isFileSystemError(error, "ENOENT")) return false;
    throw error;
  }
}

function invalidateZipDirectoryCache(zipPath: string) {
  zipDirectoryCache.delete(zipPath);
}

function touchZipDirectoryCacheEntry(zipPath: string, entry: ZipDirectoryCacheEntry) {
  zipDirectoryCache.delete(zipPath);
  zipDirectoryCache.set(zipPath, entry);
}

function trimZipDirectoryCache() {
  while (zipDirectoryCache.size > SEC_BULK_ZIP_DIRECTORY_CACHE_MAX_ENTRIES) {
    const oldest = zipDirectoryCache.keys().next().value as string | undefined;
    if (!oldest) break;
    zipDirectoryCache.delete(oldest);
  }
}

function zipFileFingerprint(stat: { dev: number | bigint; ino: number | bigint; size: number; mtimeMs: number; ctimeMs: number }) {
  return `${String(stat.dev)}:${String(stat.ino)}:${stat.size}:${stat.mtimeMs}:${stat.ctimeMs}`;
}

function assertZipArchiveIdentity(
  zipPath: string,
  expectedFingerprint: string,
  stat: { dev: number | bigint; ino: number | bigint; size: number; mtimeMs: number; ctimeMs: number }
) {
  if (zipFileFingerprint(stat) !== expectedFingerprint) throw new ZipArchiveIdentityChangedError(zipPath);
}

function isFileSystemError(error: unknown, code: string) {
  return Boolean(error && typeof error === "object" && "code" in error && (error as { code?: unknown }).code === code);
}

function sleep(ms: number) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function positiveEnvNumber(value: string | undefined, fallback: number) {
  return positiveNumber(value === undefined ? undefined : Number(value), fallback);
}

function positiveEnvInteger(value: string | undefined, fallback: number) {
  return Math.max(1, Math.floor(positiveEnvNumber(value, fallback)));
}

function nonNegativeEnvInteger(value: string | undefined, fallback: number) {
  return nonNegativeInteger(value === undefined ? undefined : Number(value), fallback);
}

function positiveNumber(value: number | undefined, fallback: number) {
  return typeof value === "number" && Number.isFinite(value) && value > 0 ? value : fallback;
}

function nonNegativeNumber(value: number | undefined, fallback: number) {
  return typeof value === "number" && Number.isFinite(value) && value >= 0 ? value : fallback;
}

function nonNegativeInteger(value: number | undefined, fallback: number) {
  return Math.max(0, Math.floor(nonNegativeNumber(value, fallback)));
}

function findEndOfCentralDirectory(buffer: Buffer) {
  for (let offset = buffer.length - 22; offset >= 0; offset -= 1) {
    if (buffer.readUInt32LE(offset) === 0x06054b50) return offset;
  }
  return -1;
}

async function fileExists(filePath: string) {
  try {
    await fs.access(filePath);
    return true;
  } catch {
    return false;
  }
}

function isStale(mtimeMs: number) {
  return Date.now() - mtimeMs > SEC_BULK_REFRESH_INTERVAL_MS;
}

function latestTimestamp(values: string[]) {
  return values.sort().at(-1);
}

function unique(values: string[]) {
  return Array.from(new Set(values));
}

export const __secBulkTestHooks = {
  acquireFileLock,
  cleanupOrphanTempFiles,
  invalidateZipDirectoryCache,
  openZip,
  loadSecBulkSupport,
  normalizeSecCik,
  readBulkJson,
  readSecBulkSubmissionFile,
  readMetadata,
  refreshArchive,
  validateSecBulkArchive,
  validateSecBulkPayload,
  writeFileAtomically,
  writeMetadata,
  zipDirectoryCacheSize: () => zipDirectoryCache.size,
  resetInMemoryState: () => {
    zipDirectoryCache.clear();
    refreshPromises.clear();
    metadataWriteQueue = Promise.resolve();
  }
};
