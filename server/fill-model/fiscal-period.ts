import { normalizedUnitFamily } from "./normalized-layer";

export type FiscalDerivationFact = {
  unit?: string;
  cik?: string;
  periodType?: "instant" | "quarterly" | "year_to_date" | "annual";
};

/**
 * Quarterly subtraction is valid only for additive duration flows. Instant
 * balances, share counts, per-share values, and ratios are deliberately
 * excluded even when SEC metadata happens to expose a duration context.
 */
export function quarterlyFlowDerivationAllowed(concept: string, source?: FiscalDerivationFact) {
  if (/WeightedAverage|EarningsPerShare|SharesOutstanding|ShareCount|PerShare|Instant|BalanceSheet/i.test(concept)) return false;
  if (source?.periodType === "instant") return false;
  if (source?.unit && normalizedUnitFamily(source.unit) !== "currency") return false;
  return true;
}

export function quarterlyFlowInputsCompatible(sources: FiscalDerivationFact[]) {
  if (!sources.length || sources.some((source) => source.periodType === "instant")) return false;
  const units = unique(sources.map((source) => source.unit ?? "").filter(Boolean));
  if (units.length > 1) return false;
  if (units.length && normalizedUnitFamily(units[0]) !== "currency") return false;
  const ciks = unique(sources.map((source) => source.cik ?? "").filter(Boolean));
  return ciks.length <= 1;
}

function unique<T>(items: T[]) {
  return Array.from(new Set(items));
}
