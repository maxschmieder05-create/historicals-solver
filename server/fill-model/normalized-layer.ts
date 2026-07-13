export type NormalizedPeriodType = "instant" | "quarterly" | "year_to_date" | "annual";

export type NormalizedUnitFamily = "currency" | "shares" | "per_share" | "ratio" | "unknown";

export type NormalizedSourceFactInput = {
  concept: string;
  label?: string;
  value: number;
  taxonomy?: string;
  unit?: string;
  cik?: string;
  accession?: string;
  start?: string;
  end?: string;
  period: string;
  sourcePeriod?: string;
  periodType?: NormalizedPeriodType;
  sourceLayer?: string;
};

export type NormalizedFact = {
  id: string;
  sourceFactId: string;
  concept: string;
  label: string;
  taxonomy: string;
  value: number;
  unit: string;
  unitFamily: NormalizedUnitFamily;
  cik: string;
  accession: string;
  startDate: string;
  endDate: string;
  period: string;
  sourcePeriod: string;
  periodType: NormalizedPeriodType | "";
  sourceLayer: string;
  derived: boolean;
};

export type CellFormulaPolicy = "hardcode" | "preserve" | "replace_with_reported_actual" | "protected";

export type CellAssignmentRequest = {
  sheetName: string;
  cell: string;
  modelCategory: string;
  modelRow: string;
  statement: "income" | "balance" | "cash_flow" | "segment" | "support";
  expectedPeriodType: "duration" | "instant";
  expectedUnitFamilies: NormalizedUnitFamily[];
  period: string;
  value: number;
  facts: NormalizedFact[];
  mappingType: "direct" | "derived" | "grouped" | "residual";
  formulaPolicy: CellFormulaPolicy;
  projected: boolean;
  reportedPeriod: boolean;
};

export type AuthorizedCellAssignment = CellAssignmentRequest & {
  authorized: boolean;
  reasons: string[];
  directSourceFactIds: string[];
};

export function normalizeSourceFact(input: NormalizedSourceFactInput): NormalizedFact {
  const taxonomy = input.taxonomy?.trim() ?? "";
  const unit = input.unit?.trim() ?? "";
  const accession = normalizeKey(input.accession ?? "");
  const sourceFactId = stableFactId([
    input.cik ?? "",
    accession,
    taxonomy,
    input.concept,
    input.start ?? "",
    input.end ?? "",
    unit,
    String(input.value)
  ]);
  return {
    id: stableFactId([sourceFactId, input.period, input.periodType ?? ""]),
    sourceFactId,
    concept: input.concept,
    label: input.label?.trim() || input.concept,
    taxonomy,
    value: input.value,
    unit,
    unitFamily: normalizedUnitFamily(unit),
    cik: input.cik ?? "",
    accession,
    startDate: input.start ?? "",
    endDate: input.end ?? "",
    period: input.period,
    sourcePeriod: input.sourcePeriod ?? input.period,
    periodType: input.periodType ?? "",
    sourceLayer: input.sourceLayer ?? "",
    derived: input.sourceLayer === "derived"
  };
}

export function authorizeCellAssignment(request: CellAssignmentRequest): AuthorizedCellAssignment {
  const reasons: string[] = [];
  if (!request.sheetName || !request.cell || !request.modelRow || !request.period) {
    reasons.push("Target sheet, cell, model row, and period are required.");
  }
  if (!Number.isFinite(request.value)) reasons.push("Target value is not finite.");
  if (!request.facts.length) reasons.push("No normalized SEC fact supports the target value.");
  if (request.formulaPolicy === "protected") reasons.push("Protected formula/check cells cannot be overwritten.");
  if (request.projected && request.formulaPolicy !== "replace_with_reported_actual") {
    reasons.push("Projected cells require an actual reported filing and explicit actualization authorization.");
  }
  if (request.formulaPolicy === "replace_with_reported_actual" && !request.reportedPeriod) {
    reasons.push("A formula may be replaced only when the target period has a reported SEC filing.");
  }

  const finalFacts = finalOutputFacts(request.facts);
  for (const fact of finalFacts) {
    if (fact.period && fact.period !== request.period && !fact.derived) {
      reasons.push(`${fact.concept} belongs to ${fact.period}, not ${request.period}.`);
    }
    if (!periodTypeCompatible(request.expectedPeriodType, fact.periodType, fact.derived)) {
      reasons.push(`${fact.concept} has ${fact.periodType || "unknown"} period type; ${request.modelRow} expects ${request.expectedPeriodType}.`);
    }
    if (
      request.expectedUnitFamilies.length &&
      fact.unitFamily !== "unknown" &&
      !request.expectedUnitFamilies.includes(fact.unitFamily)
    ) {
      reasons.push(`${fact.concept} uses ${fact.unit || fact.unitFamily}, which is incompatible with ${request.modelRow}.`);
    }
  }

  const directSourceFactIds =
    request.mappingType === "direct"
      ? unique(finalFacts.filter((fact) => !fact.derived).map((fact) => fact.sourceFactId).filter(Boolean))
      : [];
  return { ...request, authorized: reasons.length === 0, reasons: unique(reasons), directSourceFactIds };
}

export function validateAuthorizedCellAssignments(assignments: AuthorizedCellAssignment[]) {
  const errors: string[] = [];
  for (const assignment of assignments) {
    if (!assignment.authorized) {
      errors.push(
        `${assignment.sheetName}!${assignment.cell} ${assignment.period}: unauthorized normalized assignment (${assignment.reasons.join(" ")})`
      );
    }
  }

  const directAllocations = new Map<string, AuthorizedCellAssignment>();
  for (const assignment of assignments.filter((item) => item.authorized)) {
    for (const sourceFactId of assignment.directSourceFactIds) {
      const key = [assignment.statement, assignment.period, sourceFactId].join("|");
      const prior = directAllocations.get(key);
      if (!prior) {
        directAllocations.set(key, assignment);
        continue;
      }
      if (prior.modelCategory === assignment.modelCategory || prior.modelRow === assignment.modelRow) continue;
      errors.push(
        `${assignment.statement} ${assignment.period}: source fact ${sourceFactId} was directly assigned to both "${prior.modelRow}" (${prior.sheetName}!${prior.cell}) and "${assignment.modelRow}" (${assignment.sheetName}!${assignment.cell}).`
      );
    }
  }
  return unique(errors);
}

export function normalizedUnitFamily(unit: string): NormalizedUnitFamily {
  const normalized = unit.trim().toLowerCase().replace(/\s+/g, "");
  if (!normalized) return "unknown";
  if (normalized === "shares" || normalized.endsWith(":shares")) return "shares";
  if (normalized.includes("/shares") || normalized.includes("pershare")) return "per_share";
  if (normalized === "pure" || normalized === "percent" || normalized === "%") return "ratio";
  if (/^(?:iso4217:)?[a-z]{3}$/.test(normalized) || /^(?:usd|eur|gbp|jpy|cad|aud|chf|cny)$/.test(normalized)) return "currency";
  return "unknown";
}

function finalOutputFacts(facts: NormalizedFact[]) {
  const first = facts[0];
  if (first?.derived) return [first];
  return facts.filter((fact) => !fact.derived);
}

function periodTypeCompatible(expected: "duration" | "instant", actual: NormalizedFact["periodType"], derived: boolean) {
  if (expected === "instant") return actual === "instant";
  if (derived) return actual === "quarterly" || actual === "annual";
  return actual === "quarterly" || actual === "annual";
}

function stableFactId(parts: string[]) {
  return parts.map((part) => normalizeKey(part)).join("|");
}

function normalizeKey(value: string) {
  return value.trim().toLowerCase().replace(/[^a-z0-9.:/_-]+/g, "");
}

function unique<T>(items: T[]) {
  return Array.from(new Set(items));
}
