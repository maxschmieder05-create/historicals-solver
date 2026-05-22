export type LiabilityTemplateRow = {
  label: string;
  statement?: string;
  concepts?: string[];
};

export type TemplateMappingContext = {
  hasCurrentDebtRow: boolean;
  hasDebtInclCurrentPortionRow: boolean;
  hasNonCurrentLeaseLiabilityRow: boolean;
  hasPensionLiabilityRow: boolean;
};

export const NONCURRENT_DEBT_CONCEPTS = ["LongTermDebtNoncurrent", "LongTermDebt"];

export const COMBINED_NONCURRENT_DEBT_AND_LEASE_CONCEPTS = [
  "LongTermDebtAndFinanceLeaseObligationsNoncurrent",
  "LongTermDebtAndCapitalLeaseObligationsNoncurrent"
];

export const NONCURRENT_LEASE_LIABILITY_CONCEPTS = [
  "OperatingLeaseLiabilityNoncurrent",
  "FinanceLeaseLiabilityNoncurrent",
  "LesseeOperatingLeaseLiabilityNoncurrent"
];

export const PENSION_LIABILITY_CONCEPTS = [
  "PensionAndOtherPostretirementDefinedBenefitPlansLiabilityNoncurrent",
  "DefinedBenefitPensionPlanLiabilitiesNoncurrent",
  "OtherPostretirementDefinedBenefitPlanLiabilitiesNoncurrent"
];

const CURRENT_DEBT_CONCEPTS = ["DebtCurrent", "LongTermDebtCurrent", "CurrentPortionOfLongTermDebt", "ShortTermBorrowings", "ShortTermBorrowingsCurrent"];

function normalizeLiabilityLabel(input: string) {
  return input.toLowerCase().replace(/[^a-z0-9]/g, "");
}

function rowHasConcept(row: LiabilityTemplateRow, concepts: string[]) {
  return row.concepts?.some((concept) => concepts.includes(concept)) ?? false;
}

function rowMatchesLabel(row: LiabilityTemplateRow, patterns: RegExp[]) {
  const label = normalizeLiabilityLabel(row.label);
  return patterns.some((pattern) => pattern.test(label));
}

export function buildLiabilityTemplateMappingContext(rows: LiabilityTemplateRow[]): TemplateMappingContext {
  const balanceRows = rows.filter((row) => !row.statement || row.statement === "balance");
  const hasAnyConcept = (concepts: string[]) => balanceRows.some((row) => rowHasConcept(row, concepts));
  const hasAnyLabel = (patterns: RegExp[]) => balanceRows.some((row) => rowMatchesLabel(row, patterns));

  return {
    hasCurrentDebtRow:
      hasAnyConcept(CURRENT_DEBT_CONCEPTS) ||
      hasAnyLabel([
        /^currentdebt$/,
        /^shorttermdebt$/,
        /^shorttermborrowings$/,
        /^currentportionoflongtermdebt$/,
        /^currentmaturitiesoflongtermdebt$/
      ]),
    hasDebtInclCurrentPortionRow: hasAnyLabel([
      /^ltdebtinclcurrentportion$/,
      /^longtermdebtinclcurrentportion$/,
      /^longtermdebtincludingcurrentportion$/,
      /^longtermdebtincludingcurrentmaturities$/,
      /^borrowingsincludingcurrentportion$/,
      /^totaldebt$/
    ]),
    hasNonCurrentLeaseLiabilityRow:
      hasAnyConcept(NONCURRENT_LEASE_LIABILITY_CONCEPTS) ||
      hasAnyLabel([/^(operating|finance)?leaseliabilities$/, /^(operating|finance)?leaseobligations$/]),
    hasPensionLiabilityRow:
      hasAnyConcept(PENSION_LIABILITY_CONCEPTS) ||
      hasAnyLabel([/^pensionliabilities$/, /^pensionandotherpostretirementliabilities$/, /^postretirementliabilities$/])
  };
}

export function currentDebtBelongsInAccruedLiabilities(context: TemplateMappingContext) {
  return !context.hasCurrentDebtRow;
}

export function otherNonCurrentLiabilityResidualExclusions(context: TemplateMappingContext) {
  return {
    alwaysExcludeCurrentLiabilities: true,
    alwaysExcludeNonCurrentDebt: true,
    alwaysExcludeDeferredTaxLiabilities: true,
    excludeLeaseLiabilities: context.hasNonCurrentLeaseLiabilityRow,
    excludePensionLiabilities: context.hasPensionLiabilityRow
  };
}
