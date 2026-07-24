import { randomUUID } from "node:crypto";
import { mkdir, readFile, rename, rm, writeFile } from "node:fs/promises";
import path from "node:path";

export type ApprovedMappingScope = "approved_exact" | "company_historical";

export type ApprovedFinancialMapping = {
  scope: ApprovedMappingScope;
  companyTicker?: string;
  companyName?: string;
  statement: string;
  sourceTableType?: string;
  section?: string;
  xbrlTag?: string;
  reportedLabel?: string;
  modelRow: string;
  action?: "map" | "remap" | "merge_into_other" | "exclude";
  explanation: string;
  approvedBy: string;
  approvedAt: string;
};

export type ApprovedMappingCache = {
  version: 1;
  mappings: ApprovedFinancialMapping[];
};

export type ApprovedMappingLookup = {
  company: { name: string; ticker: string };
  statement: string;
  sourceTableType: string;
  section: string;
  xbrlTag?: string;
  reportedLabel: string;
  availableModelRows: string[];
};

const APPROVED_MAPPING_ACTIONS = new Set<NonNullable<ApprovedFinancialMapping["action"]>>([
  "map",
  "remap",
  "merge_into_other",
  "exclude"
]);
const mappingWriteQueues = new Map<string, Promise<void>>();

export async function loadApprovedMappingCache(filePath = approvedMappingCachePath()): Promise<ApprovedMappingCache> {
  try {
    const parsed = JSON.parse(await readFile(filePath, "utf8")) as Partial<ApprovedMappingCache> | null;
    if (!parsed || parsed.version !== 1 || !Array.isArray(parsed.mappings)) return emptyApprovedMappingCache();
    return { version: 1, mappings: parsed.mappings.filter(validApprovedMapping) };
  } catch (error) {
    if ((error as NodeJS.ErrnoException)?.code === "ENOENT") return emptyApprovedMappingCache();
    throw error;
  }
}

export function findApprovedMapping(cache: ApprovedMappingCache | undefined, request: ApprovedMappingLookup) {
  if (!cache?.mappings.length) return null;
  const mappings = cache.mappings.filter(validApprovedMapping);
  const exact = bestMatchingMapping(
    mappings.filter((mapping) => mapping.scope === "approved_exact"),
    request,
    false
  );
  if (exact) return exact;
  return bestMatchingMapping(
    mappings.filter((mapping) => mapping.scope === "company_historical"),
    request,
    true
  );
}

export async function cacheApprovedMapping(
  mapping: ApprovedFinancialMapping,
  filePath = approvedMappingCachePath()
) {
  if (!validApprovedMapping(mapping)) throw new Error("Approved mapping is incomplete or invalid.");
  const resolvedPath = path.resolve(filePath);
  return enqueueMappingWrite(resolvedPath, async () => {
    const cache = await loadApprovedMappingCache(resolvedPath);
    const identity = approvedMappingIdentity(mapping);
    const mappings = cache.mappings.filter((item) => approvedMappingIdentity(item) !== identity);
    mappings.push(mapping);
    const next: ApprovedMappingCache = { version: 1, mappings };
    await writeApprovedMappingCache(resolvedPath, next);
    return next;
  });
}

export function approvedMappingCachePath() {
  return path.resolve(process.env.APPROVED_MAPPING_CACHE_PATH || path.join(process.cwd(), "config", "approved_mappings.json"));
}

function mappingMatches(mapping: ApprovedFinancialMapping, request: ApprovedMappingLookup, requireCompany: boolean) {
  if (requireCompany) {
    const mappingTicker = normalize(mapping.companyTicker ?? "");
    const mappingName = normalize(mapping.companyName ?? "");
    const tickerMatches = Boolean(mappingTicker && mappingTicker === normalize(request.company.ticker));
    const nameMatches = Boolean(mappingName && mappingName === normalize(request.company.name));
    if (!tickerMatches && !nameMatches) return false;
  }
  if (normalize(mapping.statement) !== normalize(request.statement)) return false;
  if (mapping.sourceTableType && normalize(mapping.sourceTableType) !== normalize(request.sourceTableType)) return false;
  if (mapping.section && normalize(mapping.section) !== normalize(request.section)) return false;
  if (mapping.xbrlTag && normalize(mapping.xbrlTag) !== normalize(request.xbrlTag ?? "")) return false;
  if (mapping.reportedLabel && normalize(mapping.reportedLabel) !== normalize(request.reportedLabel)) return false;
  if (!mapping.xbrlTag && !mapping.reportedLabel) return false;
  return request.availableModelRows.some((row) => normalize(row) === normalize(mapping.modelRow));
}

function validApprovedMapping(mapping: unknown): mapping is ApprovedFinancialMapping {
  if (!mapping || typeof mapping !== "object") return false;
  const item = mapping as Partial<ApprovedFinancialMapping>;
  if (item.scope !== "approved_exact" && item.scope !== "company_historical") return false;
  if (
    !nonEmptyString(item.statement) ||
    !nonEmptyString(item.modelRow) ||
    !nonEmptyString(item.explanation) ||
    !nonEmptyString(item.approvedBy) ||
    !validTimestamp(item.approvedAt)
  ) {
    return false;
  }
  const optionalFields = [
    item.companyTicker,
    item.companyName,
    item.sourceTableType,
    item.section,
    item.xbrlTag,
    item.reportedLabel
  ];
  if (!optionalFields.every(optionalString)) {
    return false;
  }
  if (!nonEmptyString(item.xbrlTag) && !nonEmptyString(item.reportedLabel)) return false;
  if (item.action !== undefined && !APPROVED_MAPPING_ACTIONS.has(item.action)) return false;
  return item.scope !== "company_historical" || nonEmptyString(item.companyTicker) || nonEmptyString(item.companyName);
}

function nonEmptyString(value: unknown): value is string {
  return typeof value === "string" && Boolean(value.trim());
}

function optionalString(value: unknown): value is string | undefined {
  return value === undefined || typeof value === "string";
}

function validTimestamp(value: unknown): value is string {
  return nonEmptyString(value) && Number.isFinite(Date.parse(value));
}

function bestMatchingMapping(
  mappings: ApprovedFinancialMapping[],
  request: ApprovedMappingLookup,
  requireCompany: boolean
) {
  return (
    mappings
      .filter((mapping) => mappingMatches(mapping, request, requireCompany))
      .map((mapping, index) => ({ mapping, index, specificity: mappingSpecificity(mapping) }))
      .sort(
        (left, right) =>
          right.specificity - left.specificity ||
          approvedAtTimestamp(right.mapping) - approvedAtTimestamp(left.mapping) ||
          right.index - left.index
      )[0]?.mapping ?? null
  );
}

function mappingSpecificity(mapping: ApprovedFinancialMapping) {
  return (
    (mapping.xbrlTag ? 8 : 0) +
    (mapping.reportedLabel ? 4 : 0) +
    (mapping.sourceTableType ? 2 : 0) +
    (mapping.section ? 1 : 0)
  );
}

function approvedAtTimestamp(mapping: ApprovedFinancialMapping) {
  const timestamp = Date.parse(mapping.approvedAt);
  return Number.isFinite(timestamp) ? timestamp : 0;
}

async function writeApprovedMappingCache(filePath: string, cache: ApprovedMappingCache) {
  await mkdir(path.dirname(filePath), { recursive: true });
  const temporaryPath = `${filePath}.${process.pid}-${randomUUID()}.tmp`;
  try {
    await writeFile(temporaryPath, `${JSON.stringify(cache, null, 2)}\n`, { encoding: "utf8", flag: "wx" });
    await rename(temporaryPath, filePath);
  } finally {
    await rm(temporaryPath, { force: true }).catch(() => undefined);
  }
}

async function enqueueMappingWrite<T>(filePath: string, write: () => Promise<T>) {
  const previous = mappingWriteQueues.get(filePath) ?? Promise.resolve();
  const operation = previous.catch(() => undefined).then(write);
  const queueEntry = operation.then(
    () => undefined,
    () => undefined
  );
  mappingWriteQueues.set(filePath, queueEntry);
  try {
    return await operation;
  } finally {
    if (mappingWriteQueues.get(filePath) === queueEntry) mappingWriteQueues.delete(filePath);
  }
}

function emptyApprovedMappingCache(): ApprovedMappingCache {
  return { version: 1, mappings: [] };
}

function approvedMappingIdentity(mapping: ApprovedFinancialMapping) {
  return [
    mapping.scope,
    mapping.scope === "company_historical" ? mapping.companyTicker || mapping.companyName || "" : "",
    mapping.statement,
    mapping.sourceTableType ?? "",
    mapping.section ?? "",
    mapping.xbrlTag ?? "",
    mapping.reportedLabel ?? ""
  ]
    .map(normalize)
    .join("|");
}

function normalize(value: string) {
  return value.trim().toLowerCase().replace(/[^a-z0-9]+/g, "");
}
