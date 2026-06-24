export type BalanceSheetResolverState =
  | "direct_sec_sourced"
  | "residual_calculated"
  | "explicit_zero_not_applicable"
  | "unresolved_failure";

export type BalanceSheetSide = "assets" | "liabilities_and_equity";

export type BalanceSheetRowFamily =
  | "current_assets"
  | "non_current_assets"
  | "current_liabilities"
  | "non_current_liabilities"
  | "equity"
  | "totals";

export type BalanceSheetRowKind = "component" | "catch_all" | "subtotal" | "total";

export type BalanceSheetSourceSection =
  | "current assets"
  | "non-current assets"
  | "current liabilities"
  | "non-current liabilities"
  | "equity"
  | "unknown";

export type BalanceSheetRowDefinition = {
  canonical: string;
  family: BalanceSheetRowFamily;
  side: BalanceSheetSide;
  kind: BalanceSheetRowKind;
  aliases: string[];
  sourceAliases?: string[];
  tags: string[];
  residualEligible?: boolean;
};

export type BalanceSheetResolvedSource = {
  concept?: string;
  label?: string;
  value?: number | null;
  sourceLayer?: string;
  note?: string;
};

export type BalanceSheetResolvedCellInput = {
  modelRow: string;
  value: number | null;
  sources: BalanceSheetResolvedSource[];
  classification?: string;
  note?: string;
};

export type BalanceSheetResolvedCellState = {
  state: BalanceSheetResolverState;
  canonicalModelRow: string | null;
  residualFormula: string;
  reason: string;
};

export const BALANCE_SHEET_ROW_DEFINITIONS: BalanceSheetRowDefinition[] = [
  {
    canonical: "Cash & Cash Equivalents",
    family: "current_assets",
    side: "assets",
    kind: "component",
    aliases: ["Cash and Cash Equivalents", "Cash and Equivalents", "Cash", "Cash & Short-Term Investments", "Cash and Short-Term Investments"],
    tags: [
      "CashAndCashEquivalentsAtCarryingValue",
      "CashAndDueFromBanks",
      "InterestBearingDepositsInBanks",
      "InterestBearingDepositsInBanksAndOtherFinancialInstitutions",
      "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents",
      "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalentsIncludingDisposalGroupAndDiscontinuedOperations"
    ],
    sourceAliases: ["Cash and due from banks", "Due from banks", "Interest-bearing deposits in banks", "Interest-bearing deposits in other banks"]
  },
  {
    canonical: "Short-Term Investments",
    family: "current_assets",
    side: "assets",
    kind: "component",
    aliases: [
      "Short Term Investments",
      "Marketable Securities",
      "Current Marketable Securities",
      "Investment Securities",
      "Available-for-Sale Securities",
      "Treasury Securities"
    ],
    tags: [
      "MarketableSecuritiesCurrent",
      "ShortTermInvestments",
      "OtherShortTermInvestments",
      "DebtAndEquitySecuritiesAvailableForSale",
      "AvailableForSaleSecuritiesDebtSecuritiesCurrent",
      "DebtSecuritiesAvailableForSaleCurrent"
    ]
  },
  {
    canonical: "Accounts Receivable",
    family: "current_assets",
    side: "assets",
    kind: "component",
    aliases: ["Accounts Receivable, Net", "Trade Receivables", "Fees Receivable", "Receivables"],
    tags: ["AccountsReceivableNetCurrent", "AccountsReceivableNet", "TradeAccountsReceivableNetCurrent", "ReceivablesNetCurrent"]
  },
  {
    canonical: "Inventory",
    family: "current_assets",
    side: "assets",
    kind: "component",
    aliases: ["Inventories"],
    sourceAliases: ["Parts and supplies", "Spare parts", "Raw materials", "Work-in-process", "Finished goods"],
    tags: ["InventoryNet", "InventoryFinishedGoods", "InventoryWorkInProcess", "InventoryRawMaterialsAndSupplies", "InventoryPartsAndSupplies"]
  },
  {
    canonical: "Prepaid & Other Current Assets",
    family: "current_assets",
    side: "assets",
    kind: "catch_all",
    residualEligible: true,
    aliases: ["Prepaid and Other Current Assets", "Other Current Assets", "Prepaids and Other Current Assets"],
    tags: ["OtherAssetsCurrent", "OtherCurrentAssets", "PrepaidExpenseAndOtherAssetsCurrent", "PrepaidExpenseCurrent"]
  },
  {
    canonical: "PP&E, Net",
    family: "non_current_assets",
    side: "assets",
    kind: "component",
    aliases: ["Property Plant and Equipment Net", "Property, Plant and Equipment, Net", "Property and Equipment, Net", "Real Estate Investments"],
    tags: ["PropertyPlantAndEquipmentNet", "PropertyAndEquipmentNet", "RealEstateInvestmentPropertyNet", "RealEstateInvestments"]
  },
  {
    canonical: "Intangible Assets, Net",
    family: "non_current_assets",
    side: "assets",
    kind: "component",
    aliases: ["Intangibles, Net", "Intangible Assets"],
    tags: ["FiniteLivedIntangibleAssetsNet", "IntangibleAssetsNetExcludingGoodwill", "IndefiniteLivedIntangibleAssets", "Trademarks"]
  },
  {
    canonical: "Goodwill",
    family: "non_current_assets",
    side: "assets",
    kind: "component",
    aliases: [],
    tags: ["Goodwill"]
  },
  {
    canonical: "Other Non-Current Assets",
    family: "non_current_assets",
    side: "assets",
    kind: "catch_all",
    residualEligible: true,
    aliases: ["Other Long-Term Assets", "Other LT Assets", "Other Assets and Loans", "Other Noncurrent Assets"],
    tags: ["OtherAssetsNoncurrent", "LongTermInvestments", "InvestmentsNoncurrent", "OperatingLeaseRightOfUseAsset", "OperatingLeaseRightOfUseAssetNet"]
  },
  {
    canonical: "Accounts Payable",
    family: "current_liabilities",
    side: "liabilities_and_equity",
    kind: "component",
    aliases: ["Accounts Payable and Accrued Liabilities", "Accounts Payable & Accrued Liabilities", "Pharmacy Costs Payable", "Trade Payables"],
    tags: ["AccountsPayableCurrent", "AccountsPayableTradeCurrent", "AccountsPayableAndAccruedLiabilitiesCurrent"]
  },
  {
    canonical: "Accrued Liabilities",
    family: "current_liabilities",
    side: "liabilities_and_equity",
    kind: "component",
    aliases: ["Accrued Expenses", "Accrued Expenses and Other", "Accrued Expenses and Other Current Liabilities"],
    sourceAliases: ["Accrued and other current liabilities", "Other accrued expenses and liabilities", "Accrued operating liabilities"],
    tags: ["AccruedLiabilitiesCurrent", "AccruedIncomeTaxesCurrent", "EmployeeRelatedLiabilitiesCurrent", "OtherAccruedLiabilitiesCurrent"]
  },
  {
    canonical: "Other Current Liabilities",
    family: "current_liabilities",
    side: "liabilities_and_equity",
    kind: "catch_all",
    residualEligible: true,
    aliases: ["Other Current Liabs", "Other Liabilities, Current", "Other Current Liability", "Other Liabilities Current"],
    sourceAliases: [
      "Accrued and other current liabilities",
      "Accrued expenses and other current liabilities",
      "Other accrued expenses and liabilities",
      "Contract liabilities, current",
      "Deferred revenue, current",
      "Current operating lease liabilities",
      "Income taxes payable"
    ],
    tags: [
      "OtherLiabilitiesCurrent",
      "OtherCurrentLiabilities",
      "ContentLiabilitiesCurrent",
      "ContractWithCustomerLiabilityCurrent",
      "DeferredRevenueCurrent",
      "DeferredIncomeCurrent",
      "UnearnedRevenueCurrent",
      "CustomerAdvancesAndDepositsCurrent",
      "OperatingLeaseLiabilityCurrent",
      "OperatingLeaseLiabilitiesCurrent",
      "LesseeOperatingLeaseLiabilityCurrent",
      "IncomeTaxesPayableCurrent",
      "TaxesPayableCurrent"
    ]
  },
  {
    canonical: "Revolver",
    family: "current_liabilities",
    side: "liabilities_and_equity",
    kind: "component",
    aliases: ["Short-Term Borrowings", "Short Term Borrowings", "Short-Term Debt", "Short Term Debt", "Current Borrowings", "Line of Credit"],
    tags: ["ShortTermBorrowings", "ShortTermBorrowingsCurrent", "CurrentBorrowings", "CommercialPaper", "RevolvingCreditFacilityCurrent"]
  },
  {
    canonical: "LT Debt (Incl. Current Portion)",
    family: "non_current_liabilities",
    side: "liabilities_and_equity",
    kind: "component",
    aliases: ["Long Term Debt", "Long-Term Debt", "Senior Notes", "Borrowings"],
    tags: ["DebtLongtermAndShorttermCombinedAmount", "LongTermDebtCurrent", "LongTermDebtNoncurrent", "LongTermDebt", "CurrentPortionOfLongTermDebt"]
  },
  {
    canonical: "Current Portion of Long-Term Debt",
    family: "current_liabilities",
    side: "liabilities_and_equity",
    kind: "component",
    aliases: ["Current Maturities of Long-Term Debt", "Debt Due Within One Year", "Current Debt Maturities"],
    tags: ["LongTermDebtCurrent", "CurrentPortionOfLongTermDebt", "CurrentMaturitiesOfLongTermDebt"]
  },
  {
    canonical: "Total Debt",
    family: "non_current_liabilities",
    side: "liabilities_and_equity",
    kind: "component",
    aliases: ["Debt", "Debt and Finance Lease Obligations"],
    tags: ["DebtLongtermAndShorttermCombinedAmount", "DebtAndCapitalLeaseObligations", "LongTermDebtAndFinanceLeaseObligations"]
  },
  {
    canonical: "Deferred Income Taxes",
    family: "non_current_liabilities",
    side: "liabilities_and_equity",
    kind: "component",
    aliases: ["Deferred Tax Liabilities", "Deferred Taxes", "Deferred Income Tax Liabilities"],
    tags: ["DeferredIncomeTaxLiabilitiesNet", "DeferredTaxLiabilitiesNoncurrent", "DeferredTaxLiabilities"]
  },
  {
    canonical: "Lease Liabilities",
    family: "non_current_liabilities",
    side: "liabilities_and_equity",
    kind: "component",
    aliases: ["Operating Lease Liabilities", "Finance Lease Liabilities", "Lease Obligations"],
    tags: ["OperatingLeaseLiabilityNoncurrent", "FinanceLeaseLiabilityNoncurrent", "LesseeOperatingLeaseLiabilityNoncurrent"]
  },
  {
    canonical: "Pension Liabilities",
    family: "non_current_liabilities",
    side: "liabilities_and_equity",
    kind: "component",
    aliases: ["Pension and Other Postretirement Liabilities", "Postretirement Liabilities"],
    tags: ["PensionAndOtherPostretirementDefinedBenefitPlansLiabilityNoncurrent", "DefinedBenefitPensionPlanLiabilitiesNoncurrent"]
  },
  {
    canonical: "Other Non-Current Liabilities",
    family: "non_current_liabilities",
    side: "liabilities_and_equity",
    kind: "catch_all",
    residualEligible: true,
    aliases: ["Other Long-Term Liabilities", "Other LT Liabilities", "Other Noncurrent Liabilities"],
    tags: [
      "OtherLiabilitiesNoncurrent",
      "AccruedIncomeTaxesNoncurrent",
      "LongTermIncomeTaxesPayable",
      "DeferredRevenueNoncurrent",
      "DeferredIncomeNoncurrent",
      "ContractWithCustomerLiabilityNoncurrent",
      "AssetRetirementObligationsNoncurrent"
    ]
  },
  {
    canonical: "Common Stock & APIC",
    family: "equity",
    side: "liabilities_and_equity",
    kind: "catch_all",
    aliases: ["Common Stock and APIC", "Common Stock and Additional Paid-In Capital", "Common Stock & Additional Paid-In Capital", "Additional Paid-In Capital"],
    tags: ["CommonStocksIncludingAdditionalPaidInCapital", "AdditionalPaidInCapitalCommonStocks", "AdditionalPaidInCapitalCommonStock", "AdditionalPaidInCapital"]
  },
  {
    canonical: "Retained Earnings",
    family: "equity",
    side: "liabilities_and_equity",
    kind: "component",
    aliases: ["Accumulated Deficit", "Reinvested Earnings"],
    tags: ["RetainedEarningsAccumulatedDeficit"]
  },
  {
    canonical: "Treasury Stock",
    family: "equity",
    side: "liabilities_and_equity",
    kind: "component",
    aliases: ["Treasury & Preferred Stock", "Preferred Stock"],
    tags: ["TreasuryStockCommonValue", "TreasuryStockValue", "PreferredStockValue"]
  },
  {
    canonical: "Accumulated Other Comprehensive Income (AOCI)",
    family: "equity",
    side: "liabilities_and_equity",
    kind: "component",
    aliases: ["Accumulated Other Comprehensive Income", "Accumulated Other Comprehensive Income (Loss)", "AOCI"],
    tags: ["AccumulatedOtherComprehensiveIncomeLossNetOfTax"]
  },
  {
    canonical: "Noncontrolling Interests",
    family: "equity",
    side: "liabilities_and_equity",
    kind: "component",
    aliases: ["Non-Controlling Interests", "Minority Interest"],
    tags: ["MinorityInterest", "NoncontrollingInterestInConsolidatedEntity"]
  },
  {
    canonical: "Mezzanine Equity",
    family: "equity",
    side: "liabilities_and_equity",
    kind: "component",
    aliases: ["Redeemable Noncontrolling Interests", "Redeemable NCI"],
    tags: ["RedeemableNoncontrollingInterestEquityCarryingAmount"]
  },
  {
    canonical: "Total Current Assets",
    family: "totals",
    side: "assets",
    kind: "subtotal",
    aliases: [],
    tags: ["AssetsCurrent"]
  },
  {
    canonical: "Total Non-Current Assets",
    family: "totals",
    side: "assets",
    kind: "subtotal",
    aliases: ["Total Noncurrent Assets", "Total Long-Term Assets", "Total Long Term Assets"],
    tags: []
  },
  {
    canonical: "Total Assets",
    family: "totals",
    side: "assets",
    kind: "total",
    aliases: ["Assets"],
    tags: ["Assets"]
  },
  {
    canonical: "Total Current Liabilities",
    family: "totals",
    side: "liabilities_and_equity",
    kind: "subtotal",
    aliases: ["Total Current Liabilities (Excl. Debt)", "Total Current Liabilities Excl. Debt", "Current Liabilities Excluding Debt"],
    tags: ["LiabilitiesCurrent"]
  },
  {
    canonical: "Total Non-Current Liabilities",
    family: "totals",
    side: "liabilities_and_equity",
    kind: "subtotal",
    aliases: ["Total Noncurrent Liabilities", "Total Long-Term Liabilities", "Total Long Term Liabilities"],
    tags: ["LiabilitiesNoncurrent"]
  },
  {
    canonical: "Total Liabilities",
    family: "totals",
    side: "liabilities_and_equity",
    kind: "total",
    aliases: [],
    tags: ["Liabilities"]
  },
  {
    canonical: "Total Shareholders' Equity",
    family: "totals",
    side: "liabilities_and_equity",
    kind: "subtotal",
    aliases: ["Total Shareholder's Equity", "Total Shareholders Equity", "Total Stockholders' Equity", "Total Stockholders Equity"],
    tags: ["StockholdersEquity"]
  },
  {
    canonical: "Total Equity",
    family: "totals",
    side: "liabilities_and_equity",
    kind: "subtotal",
    aliases: [],
    tags: ["StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest"]
  },
  {
    canonical: "Total Liabilities & Shareholder's Equity",
    family: "totals",
    side: "liabilities_and_equity",
    kind: "total",
    aliases: ["Total Liabilities and Shareholder's Equity", "Total Liabilities and Stockholders' Equity", "Total Liabilities and Equity"],
    tags: ["LiabilitiesAndStockholdersEquity", "Assets"]
  }
];

const ROWS_BY_CANONICAL = new Map(BALANCE_SHEET_ROW_DEFINITIONS.map((row) => [normalizeBalanceSheetKey(row.canonical), row]));

export function normalizeBalanceSheetKey(input: string) {
  return input.toLowerCase().replace(/[^a-z0-9]/g, "");
}

export function balanceSheetRowDefinitionForCanonical(row: string) {
  return ROWS_BY_CANONICAL.get(normalizeBalanceSheetKey(row)) ?? null;
}

export function balanceSheetRowDefinitionForLabel(label: string, options: { includeSourceAliases?: boolean } = {}) {
  const key = normalizeBalanceSheetKey(label);
  if (!key) return null;
  for (const row of BALANCE_SHEET_ROW_DEFINITIONS) {
    const aliases = balanceSheetAliasesForDefinition(row, options);
    if (aliases.some((alias) => normalizeBalanceSheetKey(alias) === key)) return row;
  }
  return null;
}

export function canonicalBalanceSheetRow(label: string, options: { includeSourceAliases?: boolean } = {}) {
  return balanceSheetRowDefinitionForLabel(label, options)?.canonical ?? null;
}

export function balanceSheetRowAliases(row: string, options: { includeSourceAliases?: boolean } = {}) {
  const definition = balanceSheetRowDefinitionForCanonical(row) ?? balanceSheetRowDefinitionForLabel(row, options);
  return definition ? balanceSheetAliasesForDefinition(definition, options) : [row];
}

export function balanceSheetRowsEquivalent(a: string, b: string) {
  const left = canonicalBalanceSheetRow(a);
  const right = canonicalBalanceSheetRow(b);
  if (left && right) return normalizeBalanceSheetKey(left) === normalizeBalanceSheetKey(right);
  if (left) return balanceSheetRowAliases(left).some((alias) => normalizeBalanceSheetKey(alias) === normalizeBalanceSheetKey(b));
  if (right) return balanceSheetRowAliases(right).some((alias) => normalizeBalanceSheetKey(alias) === normalizeBalanceSheetKey(a));
  return normalizeBalanceSheetKey(a) === normalizeBalanceSheetKey(b);
}

export function balanceSheetRowMatchesSourceAlias(modelRow: string, sourceLabelOrTag: string) {
  const definition = balanceSheetRowDefinitionForLabel(modelRow) ?? balanceSheetRowDefinitionForCanonical(modelRow);
  if (!definition) return balanceSheetRowsEquivalent(modelRow, sourceLabelOrTag);
  const sourceKey = normalizeBalanceSheetKey(sourceLabelOrTag);
  if (!sourceKey) return false;
  return balanceSheetAliasesForDefinition(definition, { includeSourceAliases: true }).some((alias) => normalizeBalanceSheetKey(alias) === sourceKey);
}

export function balanceSheetRowsForMetric(metricName: string) {
  const definition = balanceSheetRowDefinitionForLabel(metricName, { includeSourceAliases: true });
  if (definition) return balanceSheetAliasesForDefinition(definition, { includeSourceAliases: false });
  const normalized = normalizeBalanceSheetKey(metricName);
  for (const row of BALANCE_SHEET_ROW_DEFINITIONS) {
    if (normalizeBalanceSheetKey(row.canonical).includes(normalized) || normalized.includes(normalizeBalanceSheetKey(row.canonical))) {
      return balanceSheetAliasesForDefinition(row, { includeSourceAliases: false });
    }
  }
  return [metricName];
}

export function balanceSheetSectionCompatible(
  modelRow: string,
  sourceSection: BalanceSheetSourceSection | string | null | undefined,
  source: { label?: string; tag?: string; concept?: string } = {}
) {
  const definition = balanceSheetRowDefinitionForLabel(modelRow) ?? balanceSheetRowDefinitionForCanonical(modelRow);
  if (!definition || definition.family === "totals" || definition.family === "equity") return true;
  const evidence = classifyBalanceSheetSourceSection(sourceSection, source);
  if (evidence === "unknown") return true;
  if (definition.canonical === "LT Debt (Incl. Current Portion)" && evidence === "current liabilities" && sourceLooksLikeCurrentLongTermDebtPortion(source)) return true;
  if (definition.canonical === "Total Debt" && evidence === "current liabilities" && sourceLooksLikeCurrentLongTermDebtPortion(source)) return true;
  if (definition.family === "current_assets") return evidence === "current assets";
  if (definition.family === "non_current_assets") return evidence === "non-current assets";
  if (definition.family === "current_liabilities") return evidence === "current liabilities";
  if (definition.family === "non_current_liabilities") return evidence === "non-current liabilities";
  return true;
}

function sourceLooksLikeCurrentLongTermDebtPortion(source: { label?: string; tag?: string; concept?: string }) {
  const text = `${source.label ?? ""} ${source.tag ?? ""} ${source.concept ?? ""}`.replace(/([a-z])([A-Z])/g, "$1 $2").toLowerCase();
  return /\bcurrent maturit|\bcurrent portion\b.*\blong[-\s]?term debt|\blong[-\s]?term debt\b.*\bcurrent\b|\bconvertible\b.*\bnotes?\b|\bsenior notes?\b/.test(text);
}

export function classifyBalanceSheetSourceSection(
  sourceSection: BalanceSheetSourceSection | string | null | undefined,
  source: { label?: string; tag?: string; concept?: string } = {}
): BalanceSheetSourceSection {
  const section = String(sourceSection ?? "").toLowerCase();
  if (section === "current assets" || section === "non-current assets" || section === "current liabilities" || section === "non-current liabilities" || section === "equity") {
    return section as BalanceSheetSourceSection;
  }

  const text = `${source.label ?? ""} ${source.tag ?? ""} ${source.concept ?? ""}`.replace(/([a-z])([A-Z])/g, "$1 $2").toLowerCase();
  const compact = normalizeBalanceSheetKey(text);
  if (/cashandduefrombanks|duefrombanks|interestbearingdeposits?inbanks?/.test(compact)) return "current assets";
  if (/shortterminvestments?|marketablesecurities|availableforsalesecurities|debtandequitysecurities|investmentsecurities/.test(compact)) {
    return /current/.test(compact) ? "current assets" : "non-current assets";
  }
  if (/stockholders?equity|shareholders?equity|retainedearnings|treasurystock|aoci|noncontrollinginterest|commonstock|additionalpaidincapital/.test(compact)) return "equity";
  const isCurrent = /\bcurrent\b/.test(text) || /current/.test(compact);
  const isNonCurrent = /\bnon[-\s]?current\b|\blong[-\s]?term\b/.test(text) || /noncurrent|longterm/.test(compact);
  const isAsset = /\bassets?\b/.test(text) || /asset/.test(compact);
  const isLiability = /\bliabilit(?:y|ies)\b|\bpayable\b|\bdebt\b|\bborrowings?\b|\bobligations?\b/.test(text) || /liabilit|payable|debt|borrowing|obligation/.test(compact);
  if (isCurrent && isAsset && !isLiability) return "current assets";
  if (isNonCurrent && isAsset && !isLiability) return "non-current assets";
  if (isCurrent && isLiability) return "current liabilities";
  if (isNonCurrent && isLiability) return "non-current liabilities";
  return "unknown";
}

export function balanceSheetLineLooksSubtotalLike(label: string, tag = "") {
  const text = `${label} ${tag}`.replace(/([a-z])([A-Z])/g, "$1 $2").toLowerCase();
  if (/\b(total|subtotal)\b/.test(text)) return true;
  if (/\band other\b|\bincluding\b|\bcomprised of\b|\bconsists of\b/.test(text)) {
    return /\b(current|non[-\s]?current|long[-\s]?term|assets?|liabilit(?:y|ies)|equity|accrued|accounts payable|debt|lease|tax|deferred)\b/.test(text);
  }
  return false;
}

export function classifyBalanceSheetResolution(input: BalanceSheetResolvedCellInput): BalanceSheetResolvedCellState {
  const definition = balanceSheetRowDefinitionForLabel(input.modelRow) ?? balanceSheetRowDefinitionForCanonical(input.modelRow);
  const canonicalModelRow = definition?.canonical ?? null;
  if (input.value === null || !Number.isFinite(input.value)) {
    return {
      state: "unresolved_failure",
      canonicalModelRow,
      residualFormula: "",
      reason: "No SEC-backed direct value, residual, or explicit zero was resolved for this balance-sheet row."
    };
  }

  const sources = input.sources ?? [];
  const nonModelSources = sources.filter((source) => source.sourceLayer !== "model");
  const text = `${input.classification ?? ""} ${input.note ?? ""} ${sources.map((source) => `${source.concept ?? ""} ${source.label ?? ""} ${source.note ?? ""}`).join(" ")}`;
  if (Math.abs(input.value) <= 0.0001 && (!nonModelSources.length || /not reported|no separate|no current sec source|not applicable|explicit(?:ly)? zero/i.test(text))) {
    return {
      state: "explicit_zero_not_applicable",
      canonicalModelRow,
      residualFormula: "",
      reason: input.note || "The current SEC filing did not disclose an applicable balance for this row."
    };
  }

  const residualFormula = balanceSheetResidualFormula(input.modelRow, sources);
  const looksResidual =
    input.classification === "residual" ||
    /\bresidual\b|\bderived\b|\bcalculated\b|\bless\b|\bexcluding\b|\bminus\b/i.test(text) ||
    Boolean(definition?.residualEligible && sources.length > 1);
  if (looksResidual && definition?.residualEligible === true && nonModelSources.length) {
    return {
      state: "residual_calculated",
      canonicalModelRow,
      residualFormula,
      reason: input.note || "Calculated as an SEC-sourced subtotal less SEC-sourced components."
    };
  }

  if (nonModelSources.length) {
    return {
      state: "direct_sec_sourced",
      canonicalModelRow,
      residualFormula: "",
      reason: input.note || "Mapped directly to SEC-sourced balance-sheet line item support."
    };
  }

  return {
    state: "unresolved_failure",
    canonicalModelRow,
    residualFormula: "",
    reason: "The row has a value but no SEC source, residual formula, or explicit-zero support."
  };
}

export function balanceSheetResidualFormula(modelRow: string, sources: BalanceSheetResolvedSource[]) {
  if (sources.length < 2) return "";
  const [total, ...components] = sources;
  const totalLabel = balanceSheetSourceName(total);
  const componentLabels = components.map(balanceSheetSourceName).filter(Boolean);
  if (!totalLabel || !componentLabels.length) return "";
  return `${modelRow} = ${totalLabel} - ${componentLabels.join(" - ")}`;
}

function balanceSheetAliasesForDefinition(definition: BalanceSheetRowDefinition, options: { includeSourceAliases?: boolean }) {
  return [definition.canonical, ...definition.aliases, ...(options.includeSourceAliases ? definition.sourceAliases ?? [] : [])];
}

function balanceSheetSourceName(source: BalanceSheetResolvedSource | undefined) {
  if (!source) return "";
  return source.label || source.concept || "";
}
