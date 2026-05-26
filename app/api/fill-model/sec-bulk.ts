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

const SEC_BULK_REFRESH_INTERVAL_MS = Number(process.env.SEC_BULK_REFRESH_INTERVAL_MS || 24 * 60 * 60 * 1000);
const SEC_BULK_CACHE_DIR = process.env.SEC_BULK_CACHE_DIR || path.join(process.cwd(), "tmp", "sec-bulk");
const SEC_BULK_METADATA_FILE = path.join(SEC_BULK_CACHE_DIR, "metadata.json");
const SEC_BULK_DISABLED = process.env.SEC_BULK_DISABLED === "true" || process.env.SEC_BULK_DISABLED === "1";
const SEC_BULK_REFRESH_DISABLED = process.env.SEC_BULK_REFRESH_DISABLED === "true" || process.env.SEC_BULK_REFRESH_DISABLED === "1";

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

const zipDirectoryCache = new Map<string, Promise<ZipDirectory>>();
const refreshPromises = new Map<SecBulkArchiveKind, Promise<void>>();

export async function loadSecBulkSupport(cik: string, headers: Record<string, string>): Promise<SecBulkSupport> {
  if (SEC_BULK_DISABLED) {
    return {
      archives: {},
      warnings: ["SEC bulk support is disabled by SEC_BULK_DISABLED; using live SEC APIs only."]
    };
  }

  const cikFile = `CIK${cik}.json`;
  const [companyFacts, submissions] = await Promise.all([
    readBulkJson("companyfacts", cikFile, headers),
    readBulkJson("submissions", cikFile, headers)
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

export async function readSecBulkSubmissionFile(fileName: string, headers: Record<string, string>) {
  const result = await readBulkJson("submissions", fileName, headers);
  return result.payload ?? null;
}

async function readBulkJson(kind: SecBulkArchiveKind, fileName: string, headers: Record<string, string>) {
  const status = await ensureArchive(kind, headers);
  if (!status.path) return { payload: null, status, warning: status.warning };

  try {
    const zip = await openZip(status.path);
    const text = await zip.readText(fileName);
    if (!text) return { payload: null, status, warning: `${ARCHIVES[kind].filename} did not contain ${fileName}; live SEC API fallback will be used.` };
    return { payload: JSON.parse(text), status, warning: undefined };
  } catch (error) {
    const warning = `Could not read ${fileName} from ${ARCHIVES[kind].filename}: ${error instanceof Error ? error.message : "unknown error"}. Live SEC API fallback will be used.`;
    return { payload: null, status, warning };
  }
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
    const refreshAt = metadataRefreshAt ? new Date(metadataRefreshAt) : stat.mtime;
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

function scheduleRefresh(kind: SecBulkArchiveKind, headers: Record<string, string>) {
  if (SEC_BULK_REFRESH_DISABLED || refreshPromises.has(kind)) return;
  const promise = refreshArchive(kind, headers)
    .catch(() => {
      // Existing archives and live SEC APIs are still usable if a background refresh fails.
    })
    .finally(() => {
      refreshPromises.delete(kind);
      zipDirectoryCache.delete(path.join(SEC_BULK_CACHE_DIR, ARCHIVES[kind].filename));
    });
  refreshPromises.set(kind, promise);
}

async function refreshArchive(kind: SecBulkArchiveKind, headers: Record<string, string>) {
  const config = ARCHIVES[kind];
  await fs.mkdir(SEC_BULK_CACHE_DIR, { recursive: true });
  const archivePath = path.join(SEC_BULK_CACHE_DIR, config.filename);
  const tempPath = `${archivePath}.${Date.now()}.tmp`;
  const response = await fetch(config.url, { headers });
  if (!response.ok || !response.body) throw new Error(`SEC bulk download failed for ${config.filename}: ${response.status} ${response.statusText}`);
  await pipeline(Readable.fromWeb(response.body as any), createWriteStream(tempPath));
  await fs.rename(tempPath, archivePath);
  await writeMetadata(kind, archivePath, config.url);
}

async function readMetadata(): Promise<SecBulkMetadata> {
  try {
    return JSON.parse(await fs.readFile(SEC_BULK_METADATA_FILE, "utf8")) as SecBulkMetadata;
  } catch {
    return {};
  }
}

async function writeMetadata(kind: SecBulkArchiveKind, archivePath: string, sourceUrl: string) {
  const metadata = await readMetadata();
  metadata.archives = {
    ...(metadata.archives ?? {}),
    [kind]: {
      lastRefreshAt: new Date().toISOString(),
      path: archivePath,
      sourceUrl
    }
  };
  await fs.mkdir(SEC_BULK_CACHE_DIR, { recursive: true });
  await fs.writeFile(SEC_BULK_METADATA_FILE, `${JSON.stringify(metadata, null, 2)}\n`);
}

async function openZip(zipPath: string) {
  let cached = zipDirectoryCache.get(zipPath);
  if (!cached) {
    cached = ZipDirectory.open(zipPath);
    zipDirectoryCache.set(zipPath, cached);
  }
  return cached;
}

class ZipDirectory {
  private constructor(
    private readonly zipPath: string,
    private readonly entries: Map<string, ZipEntry>
  ) {}

  static async open(zipPath: string) {
    const handle = await fs.open(zipPath, "r");
    try {
      const stat = await handle.stat();
      const tailLength = Math.min(stat.size, 66_000);
      const tail = Buffer.alloc(tailLength);
      await handle.read(tail, 0, tailLength, stat.size - tailLength);
      const eocdOffset = findEndOfCentralDirectory(tail);
      if (eocdOffset < 0) throw new Error("ZIP central directory was not found.");

      const centralDirectorySize = tail.readUInt32LE(eocdOffset + 12);
      const centralDirectoryOffset = tail.readUInt32LE(eocdOffset + 16);
      if (centralDirectorySize === 0xffffffff || centralDirectoryOffset === 0xffffffff) {
        throw new Error("ZIP64 archives are not supported by the lightweight SEC bulk reader.");
      }

      const centralDirectory = Buffer.alloc(centralDirectorySize);
      await handle.read(centralDirectory, 0, centralDirectorySize, centralDirectoryOffset);
      return new ZipDirectory(zipPath, parseCentralDirectory(centralDirectory));
    } finally {
      await handle.close();
    }
  }

  async readText(name: string) {
    const entry = this.entries.get(name);
    if (!entry) return null;
    const handle = await fs.open(this.zipPath, "r");
    try {
      const localHeader = Buffer.alloc(30);
      await handle.read(localHeader, 0, localHeader.length, entry.localHeaderOffset);
      if (localHeader.readUInt32LE(0) !== 0x04034b50) throw new Error(`Invalid local ZIP header for ${name}.`);
      const fileNameLength = localHeader.readUInt16LE(26);
      const extraLength = localHeader.readUInt16LE(28);
      const dataOffset = entry.localHeaderOffset + 30 + fileNameLength + extraLength;
      const compressed = Buffer.alloc(entry.compressedSize);
      await handle.read(compressed, 0, entry.compressedSize, dataOffset);
      if (entry.compressionMethod === 0) return compressed.toString("utf8");
      if (entry.compressionMethod === 8) return inflateRawSync(compressed).toString("utf8");
      throw new Error(`Unsupported ZIP compression method ${entry.compressionMethod} for ${name}.`);
    } finally {
      await handle.close();
    }
  }
}

function parseCentralDirectory(buffer: Buffer) {
  const entries = new Map<string, ZipEntry>();
  let offset = 0;
  while (offset + 46 <= buffer.length) {
    if (buffer.readUInt32LE(offset) !== 0x02014b50) break;
    const compressionMethod = buffer.readUInt16LE(offset + 10);
    const compressedSize = buffer.readUInt32LE(offset + 20);
    const fileNameLength = buffer.readUInt16LE(offset + 28);
    const extraLength = buffer.readUInt16LE(offset + 30);
    const commentLength = buffer.readUInt16LE(offset + 32);
    const localHeaderOffset = buffer.readUInt32LE(offset + 42);
    const name = buffer.subarray(offset + 46, offset + 46 + fileNameLength).toString("utf8");
    entries.set(name, { name, compressionMethod, compressedSize, localHeaderOffset });
    offset += 46 + fileNameLength + extraLength + commentLength;
  }
  return entries;
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
