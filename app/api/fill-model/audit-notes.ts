export type AuditNoteFactSource = {
  concept?: string;
  label?: string;
  value?: number | null;
  note?: string;
  sourceLayer?: string;
};

export type AuditNoteResolvedValue = {
  value?: number | null;
  sources: AuditNoteFactSource[];
  note?: string;
  classification?: string;
  includedLineItems?: string[];
};

const SUBTOTAL_CONCEPTS = new Set([
  "assets",
  "assetscurrent",
  "assetsnoncurrent",
  "liabilities",
  "liabilitiescurrent",
  "liabilitiesnoncurrent",
  "liabilitiesandstockholdersequity",
  "liabilitiesandstockholdersequityincludingportionattributabletononcontrollinginterest",
  "stockholdersequity",
  "stockholdersequityincludingportionattributabletononcontrollinginterest",
  "stockholdersequityincludingportionattributabletononcontrollinginterestandtemporaryequity",
  "temporaryequityandstockholdersequity",
  "partnerscapital",
  "commitmentsandcontingencies"
]);

export function lineItemMappingSentence(modelRow: string, resolved: AuditNoteResolvedValue, mode?: "maps" | "includes") {
  const labels = sourceLineItemLabels(resolved);
  return lineItemSentence(modelRow, labels, mode ?? (labels.length === 1 && (resolved.classification ?? "direct") === "direct" ? "maps" : "includes"));
}

export function lineItemSentence(modelRow: string, labels: string[], mode: "maps" | "includes" = labels.length === 1 ? "maps" : "includes") {
  const cleanModelRow = cleanLineItemLabel(modelRow) || "Model row";
  const cleanLabels = reportableLineItemLabels(labels);
  if (!cleanLabels.length) return "";
  const effectiveMode = mode === "includes" ? "includes" : cleanLabels.length === 1 ? "maps" : mode;
  const verb = effectiveMode === "maps" ? "maps directly to" : "includes";
  return `${cleanModelRow} ${verb} ${humanList(cleanLabels)}.`;
}

export function sourceLineItemLabels(resolved: AuditNoteResolvedValue) {
  if (resolved.includedLineItems?.length) return reportableLineItemLabels(resolved.includedLineItems);
  if (isResidualOrBridgeMapping(resolved)) return [];

  const labels: string[] = [];
  const seen = new Set<string>();
  for (const source of resolved.sources) {
    if (!isReportableLineItemSource(source)) continue;
    const label = sourceLineItemLabel(source);
    if (!label) continue;
    const key = normalize(label);
    if (!key || seen.has(key)) continue;
    seen.add(key);
    labels.push(label);
  }
  return labels;
}

export function sourceLineItemLabel(source: AuditNoteFactSource) {
  if (!isReportableLineItemSource(source)) return "";

  const preferred = readableSourceLabel(source.label ?? "", source.concept);
  if (preferred) return preferred;

  const fallback = readableSourceLabel(source.note ?? "", source.concept);
  if (fallback) return fallback;

  const concept = cleanLineItemLabel(source.concept ?? "");
  if (!concept || isSubtotalConcept(concept)) return "";
  return humanizeConceptForLineItem(concept);
}

export function reportableLineItemLabels(labels: string[]) {
  return uniqueByNormalizedLabel(
    labels
      .map(cleanLineItemLabel)
      .filter((label) => label && !isTechnicalLineItemLabel(label) && !isBroadSubtotalLineItemLabel(label) && !isExclusionSummaryLabel(label))
  );
}

export function humanList(items: string[]) {
  if (items.length <= 1) return items[0] ?? "";
  if (items.length === 2) return `${items[0]} and ${items[1]}`;
  return `${items.slice(0, -1).join(", ")}, and ${items.at(-1)}`;
}

export function uniqueByNormalizedLabel(labels: string[]) {
  const seen = new Set<string>();
  const result: string[] = [];
  for (const label of labels) {
    const key = normalize(label);
    if (!key || seen.has(key)) continue;
    seen.add(key);
    result.push(label);
  }
  return result;
}

export function cleanLineItemLabel(label: string) {
  return label
    .replace(/\s+/g, " ")
    .replace(/^["'`]+|["'`]+$/g, "")
    .replace(/^\s*(?:EDGAR|SEC|us-gaap|ifrs-full|dei)\s*[:\-]\s*/i, "")
    .replace(/\s*\((?:in\s+)?(?:millions|thousands|USD \$|shares|except per share data)\)\s*$/i, "")
    .trim();
}

export function isTechnicalLineItemLabel(label: string) {
  return /\b(parser|mapped|mapping|calculated|derived|validation|accession|period|formula|residual|resolved|carried forward|SEC validation|LLM-assisted|Balance Sheet Check|Segment Analysis|abstract)\b/i.test(
    label
  );
}

export function normalizedLineItemComment(text: string) {
  const normalized = text.replace(/\r\n/g, "\n").trim();
  if (!normalized || hasLegacyCommentDetail(normalized)) return "";
  if (!/^[^\n]+ (?:includes|maps directly to)(?::\n- |\s+).+/i.test(normalized)) return "";
  if (commentContainsUnreportableLineItem(normalized)) return "";
  return normalized;
}

export function hasLegacyCommentDetail(text: string) {
  return /\b(?:EDGAR:|Period:|Value:|Confidence:|parser|timestamp|accession|fragment|Filing says:|LLM-assisted|SEC validation|formula|residual|calculated|needs review|mapped to directly)\b/i.test(
    text
  );
}

function isReportableLineItemSource(source: AuditNoteFactSource) {
  if (source.value === 0 || source.sourceLayer === "model" || source.sourceLayer === "derived") return false;
  if (isSubtotalConcept(source.concept ?? "")) return false;
  const label = cleanLineItemLabel(source.label ?? "");
  if (label && (isTechnicalLineItemLabel(label) || isBroadSubtotalLineItemLabel(label) || isExclusionSummaryLabel(label))) return false;
  return true;
}

function readableSourceLabel(label: string, concept?: string) {
  const clean = cleanLineItemLabel(label);
  if (!clean || isTechnicalLineItemLabel(clean) || isBroadSubtotalLineItemLabel(clean) || isExclusionSummaryLabel(clean)) return "";
  if (!looksLikeXbrlTag(clean, concept) && clean !== concept) return clean;
  const readableConcept = humanizeConceptForLineItem(concept ?? clean);
  return readableConcept && !isBroadSubtotalLineItemLabel(readableConcept) ? readableConcept : "";
}

function isResidualOrBridgeMapping(resolved: AuditNoteResolvedValue) {
  if (resolved.sources.length <= 1) return false;
  const note = `${resolved.note ?? ""}`.toLowerCase();
  if (/\b(?:less|excluding|exclude|excludes|residual|balancing|bridge|difference)\b/.test(note)) return true;
  if (/\b(?:derived|calculated)\s+from\b/.test(note)) return true;
  return resolved.sources.some((source) => isSubtotalConcept(source.concept ?? "") || source.sourceLayer === "derived");
}

function isSubtotalConcept(concept: string) {
  const normalized = normalize(concept);
  if (!normalized) return false;
  if (SUBTOTAL_CONCEPTS.has(normalized)) return true;
  if (/abstract$/.test(normalized)) return true;
  if (/^(?:total)?(?:assets|liabilities|stockholdersequity|shareholdersequity|partnercapital)$/.test(normalized)) return true;
  if (/^(?:total)?(?:current|noncurrent)(?:assets|liabilities)$/.test(normalized)) return true;
  return false;
}

function isBroadSubtotalLineItemLabel(label: string) {
  const normalized = normalize(label);
  if (!normalized) return false;
  if (SUBTOTAL_CONCEPTS.has(normalized)) return true;
  if (/^total(?:current|noncurrent)?(?:assets|liabilities|equity|stockholdersequity|shareholdersequity)$/.test(normalized)) {
    return true;
  }
  if (/^(?:current|noncurrent)?(?:assets|liabilities)$/.test(normalized)) return true;
  if (/^(?:assets|liabilities)(?:current|noncurrent)?$/.test(normalized)) return true;
  if (/^(?:consolidated)?(?:balanceSheets?|statements?OfFinancialPosition|currentAssets|currentLiabilities)$/.test(normalized)) return true;
  return false;
}

function isExclusionSummaryLabel(label: string) {
  return /\b(?:excluding|less separately|less the|net of separately|from the primary consolidated|subtotal)\b/i.test(label);
}

function commentContainsUnreportableLineItem(comment: string) {
  const match = comment.match(/^[^\n]+ (?:includes|maps directly to)(?::\n- |\s+)([\s\S]+?)\.?$/i);
  if (!match) return false;
  return splitCommentLineItems(match[1]).some((label) => {
    const clean = cleanLineItemLabel(label);
    return Boolean(clean && (isBroadSubtotalLineItemLabel(clean) || isExclusionSummaryLabel(clean) || isTechnicalLineItemLabel(clean)));
  });
}

function splitCommentLineItems(text: string) {
  if (text.includes("\n- ")) return text.split(/\n-\s+/).map((item) => item.trim()).filter(Boolean);
  return text
    .replace(/\s+and\s+/gi, ", ")
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean);
}

function looksLikeXbrlTag(label: string, concept?: string) {
  const compact = label.replace(/[^A-Za-z0-9]/g, "");
  if (!compact) return false;
  if (concept && normalize(label) === normalize(concept)) return true;
  return !/\s/.test(label) && /[a-z][A-Z]/.test(label);
}

function humanizeConceptForLineItem(concept: string) {
  return concept
    .replace(/^(?:us-gaap|ifrs-full|dei):/i, "")
    .replace(/([a-z])([A-Z])/g, "$1 $2")
    .replace(/\bAnd\b/g, "and")
    .replace(/\bOf\b/g, "of")
    .replace(/\bFor\b/g, "for")
    .replace(/\bNet\b/g, "net")
    .replace(/\bLoss\b/g, "loss")
    .replace(/\bExpense\b/g, "expense")
    .replace(/\bIncome\b/g, "income")
    .replace(/\s+/g, " ")
    .trim();
}

function normalize(input: string) {
  return input.toLowerCase().replace(/[^a-z0-9]/g, "");
}
