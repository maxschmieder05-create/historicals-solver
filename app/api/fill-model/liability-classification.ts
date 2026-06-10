export type LiabilityTemplateRow = {
  label: string;
  statement?: string;
  concepts?: string[];
};

export type TemplateMappingContext = {
  hasCurrentInvestmentRow: boolean;
  hasCurrentDebtRow: boolean;
  hasCurrentDebtMaturitiesRow: boolean;
  hasShortTermBorrowingsRow: boolean;
  hasCurrentLiabilitiesExcludingDebtRow: boolean;
  hasOtherCurrentLiabilityRow: boolean;
  hasDebtInclCurrentPortionRow: boolean;
  hasDeferredTaxLiabilityRow: boolean;
  hasNonCurrentLeaseLiabilityRow: boolean;
  hasPensionLiabilityRow: boolean;
};

export const NONCURRENT_DEBT_CONCEPTS = [
  "LongTermDebtNoncurrent",
  "LongTermDebt",
  "LongTermDebtAndCapitalLeaseObligations",
  "LongTermNotesAndLoans",
  "LongTermNotesPayable"
];

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

export const DEFERRED_TAX_LIABILITY_CONCEPTS = [
  "DeferredIncomeTaxLiabilitiesNet",
  "DeferredTaxLiabilitiesNoncurrent",
  "DeferredTaxLiabilities"
];

export const CURRENT_DEBT_CONCEPTS = [
  "DebtCurrent",
  "LongTermDebtCurrent",
  "CurrentPortionOfLongTermDebt",
  "CurrentMaturitiesOfLongTermDebt",
  "ShortTermBorrowings",
  "ShortTermBorrowingsCurrent",
  "OtherShortTermBorrowings",
  "CurrentBorrowings",
  "NotesPayableCurrent",
  "LoansPayableCurrent",
  "CommercialPaper",
  "CommercialPaperCurrent",
  "RevolvingCreditFacility",
  "RevolvingCreditFacilityCurrent",
  "CreditFacilityCurrent",
  "LineOfCreditFacilityCurrentBorrowings",
  "LongTermDebtAndFinanceLeaseObligationsCurrent",
  "LongTermDebtAndCapitalLeaseObligationsCurrent",
  "FinanceLeaseLiabilityCurrent",
  "CapitalLeaseObligationsCurrent"
];
const CURRENT_DEBT_MATURITY_CONCEPTS = [
  "DebtCurrent",
  "LongTermDebtCurrent",
  "CurrentPortionOfLongTermDebt",
  "CurrentMaturitiesOfLongTermDebt",
  "LongTermDebtAndFinanceLeaseObligationsCurrent",
  "LongTermDebtAndCapitalLeaseObligationsCurrent",
  "FinanceLeaseLiabilityCurrent",
  "CapitalLeaseObligationsCurrent"
];
const SHORT_TERM_BORROWING_CONCEPTS = [
  "ShortTermBorrowings",
  "ShortTermBorrowingsCurrent",
  "OtherShortTermBorrowings",
  "CurrentBorrowings",
  "CommercialPaper",
  "CommercialPaperCurrent",
  "LineOfCreditFacilityCurrentBorrowings",
  "RevolvingCreditFacility",
  "RevolvingCreditFacilityCurrent",
  "CreditFacilityCurrent",
  "NotesPayableCurrent",
  "LoansPayableCurrent"
];
const CURRENT_INVESTMENT_CONCEPTS = [
  "MarketableSecuritiesCurrent",
  "ShortTermInvestments",
  "OtherShortTermInvestments",
  "AvailableForSaleSecuritiesDebtSecuritiesCurrent",
  "DebtSecuritiesAvailableForSaleCurrent"
];

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
  const hasCurrentDebtMaturitiesRow =
    hasAnyConcept(CURRENT_DEBT_MATURITY_CONCEPTS) ||
    hasAnyLabel([
      /^currentdebt$/,
      /^debtcurrent$/,
      /^currentportionoflongtermdebt$/,
      /^currentmaturitiesoflongtermdebt$/,
      /^debtduewithinoneyear$/,
      /^currentfinanceleaseliabilities$/,
      /^financeleaseliabilitiescurrent$/
    ]);
  const hasShortTermBorrowingsRow =
    hasAnyConcept(SHORT_TERM_BORROWING_CONCEPTS) ||
    hasAnyLabel([
      /^shorttermdebt$/,
      /^shorttermborrowings$/,
      /^currentborrowings$/,
      /^commercialpaper$/,
      /^revolver$/,
      /^revolvingcreditfacility$/,
      /^creditfacilitycurrent$/,
      /^lineofcredit$/,
      /^lineofcreditborrowings$/,
      /^notespayablecurrent$/,
      /^currentnotespayable$/,
      /^loanspayablecurrent$/,
      /^currentloanspayable$/
    ]);

  return {
    hasCurrentInvestmentRow:
      hasAnyConcept(CURRENT_INVESTMENT_CONCEPTS) ||
      hasAnyLabel([
        /^shortterminvestments$/,
        /^marketablesecurities$/,
        /^currentmarketablesecurities$/,
        /^investmentsecurities$/,
        /^availableforsalesecurities$/,
        /^treasurysecurities$/
      ]),
    hasCurrentDebtRow: hasCurrentDebtMaturitiesRow || hasShortTermBorrowingsRow,
    hasCurrentDebtMaturitiesRow,
    hasShortTermBorrowingsRow,
    hasCurrentLiabilitiesExcludingDebtRow: hasAnyLabel([
      /^totalcurrentliabilitiesexcldebt$/,
      /^totalcurrentliabilitiesexcludingdebt$/,
      /^totalcurrentliabilitieslessdebt$/,
      /^totalcurrentliabilitiesnetofdebt$/,
      /^totalcurrentliabilitieswithoutdebt$/,
      /^currentliabilitiesexcldebt$/,
      /^currentliabilitiesexcludingdebt$/,
      /^currentliabilitieslessdebt$/,
      /^currentliabilitiesnetofdebt$/,
      /^currentliabilitieswithoutdebt$/
    ]),
    hasOtherCurrentLiabilityRow: hasAnyLabel([/^othercurrentliabilities$/, /^othercurrentliabs$/]),
    hasDebtInclCurrentPortionRow: hasAnyLabel([
      /^ltdebtinclcurrentportion$/,
      /^longtermdebtinclcurrentportion$/,
      /^longtermdebtincludingcurrentportion$/,
      /^longtermdebtincludingcurrentmaturities$/,
      /^borrowingsincludingcurrentportion$/,
      /^totaldebt$/
    ]),
    hasDeferredTaxLiabilityRow:
      hasAnyConcept(DEFERRED_TAX_LIABILITY_CONCEPTS) ||
      hasAnyLabel([
        /^deferredincometaxes$/,
        /^deferredincometaxliabilities$/,
        /^deferredtaxliabilities$/,
        /^deferredtaxes$/
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
  if (context.hasCurrentLiabilitiesExcludingDebtRow) return false;
  return !context.hasCurrentDebtRow;
}

export function otherNonCurrentLiabilityResidualExclusions(context: TemplateMappingContext) {
  return {
    alwaysExcludeCurrentLiabilities: true,
    alwaysExcludeNonCurrentDebt: true,
    excludeDeferredTaxLiabilities: context.hasDeferredTaxLiabilityRow,
    excludeLeaseLiabilities: context.hasNonCurrentLeaseLiabilityRow,
    excludePensionLiabilities: context.hasPensionLiabilityRow
  };
}
