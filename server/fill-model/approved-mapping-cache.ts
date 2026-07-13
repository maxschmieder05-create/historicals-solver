import { readFile, rename, writeFile } from "node:fs/promises";
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

const EMPTY_CACHE: ApprovedMappingCache = { version: 1, mappings: [] };

export async function loadApprovedMappingCache(filePath = approvedMappingCachePath()): Promise<ApprovedMappingCache> {
  try {
    const parsed = JSON.parse(await readFile(filePath, "utf8")) as Partial<ApprovedMappingCache>;
    if (parsed.version !== 1 || !Array.isArray(parsed.mappings)) return EMPTY_CACHE;
    return { version: 1, mappings: parsed.mappings.filter(validApprovedMapping) };
  } catch (error) {
    if ((error as NodeJS.ErrnoException)?.code === "ENOENT") return EMPTY_CACHE;
    throw error;
  }
}

export function findApprovedMapping(cache: ApprovedMappingCache | undefined, request: ApprovedMappingLookup) {
  if (!cache?.mappings.length) return null;
  const exact = cache.mappings.find(
    (mapping) => mapping.scope === "approved_exact" && mappingMatches(mapping, request, false)
  );
  if (exact) return exact;
  return (
    cache.mappings.find(
      (mapping) => mapping.scope === "company_historical" && mappingMatches(mapping, request, true)
    ) ?? null
  );
}

export async function cacheApprovedMapping(
  mapping: ApprovedFinancialMapping,
  filePath = approvedMappingCachePath()
) {
  if (!validApprovedMapping(mapping)) throw new Error("Approved mapping is incomplete or invalid.");
  const cache = await loadApprovedMappingCache(filePath);
  const identity = approvedMappingIdentity(mapping);
  const mappings = cache.mappings.filter((item) => approvedMappingIdentity(item) !== identity);
  mappings.push(mapping);
  const next: ApprovedMappingCache = { version: 1, mappings };
  const temporaryPath = `${filePath}.${process.pid}.tmp`;
  await writeFile(temporaryPath, `${JSON.stringify(next, null, 2)}\n`, "utf8");
  await rename(temporaryPath, filePath);
  return next;
}

export function approvedMappingCachePath() {
  return path.resolve(process.env.APPROVED_MAPPING_CACHE_PATH || path.join(process.cwd(), "config", "approved_mappings.json"));
}

function mappingMatches(mapping: ApprovedFinancialMapping, request: ApprovedMappingLookup, requireCompany: boolean) {
  if (requireCompany) {
    const tickerMatches = normalize(mapping.companyTicker ?? "") === normalize(request.company.ticker);
    const nameMatches = normalize(mapping.companyName ?? "") === normalize(request.company.name);
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
  return Boolean(
    (item.scope === "approved_exact" || item.scope === "company_historical") &&
      item.statement &&
      (item.xbrlTag || item.reportedLabel) &&
      item.modelRow &&
      item.explanation &&
      item.approvedBy &&
      item.approvedAt &&
      (item.scope !== "company_historical" || item.companyTicker || item.companyName)
  );
}

function approvedMappingIdentity(mapping: ApprovedFinancialMapping) {
  return [
    mapping.scope,
    mapping.companyTicker ?? mapping.companyName ?? "",
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
