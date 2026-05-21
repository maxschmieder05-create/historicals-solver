import { NextRequest, NextResponse } from "next/server";
import ExcelJS from "exceljs";

export const runtime = "nodejs";
export const maxDuration = 60;

type SecFact = {
  val: number;
  fy?: number;
  fp?: string;
  frame?: string;
  start?: string;
  end?: string;
  filed?: string;
  form?: string;
  accn?: string;
};

type CompanyMatch = {
  cik: string;
  ticker: string;
  title: string;
};

type FactSource = {
  concept: string;
  label: string;
  value: number;
  note?: string;
  sourceUrl?: string;
  form?: string;
  fp?: string;
  filed?: string;
  accn?: string;
  start?: string;
  end?: string;
  derivedTotalValue?: number;
  derivedTotalLabel?: string;
  derivedPriorPeriods?: string[];
};

type PeriodValues = Record<string, number | null>;

type FillRow = {
  row: number;
  label: string;
  classification: RowClassification;
  statement: "income" | "balance" | "support";
  kind: "duration" | "instant";
  scale?: number;
  sign?: 1 | -1;
  concepts?: string[];
  resolver?: (period: string, ctx: ResolveContext) => ResolvedValue;
  comment?: string;
  noFillComment?: string;
  modelContext?: ModelRowContext;
  allowBlankHistoricalInput?: boolean;
  onlyBlankHistoricalInput?: boolean;
};

type ResolvedValue = {
  value: number | null;
  sources: FactSource[];
  note?: string;
  classification?: RowClassification;
};

type LlmMappingDecision = {
  operation: "direct" | "sum" | "difference" | "needs_review";
  selectedConcepts: string[];
  sign: 1 | -1;
  confidence: "high" | "medium" | "low";
  reason: string;
  requiresReview: boolean;
};

type LlmMappingState = {
  enabled: boolean;
  decisions: Map<number, FillRow | null>;
  warnings: string[];
  calls: number;
  maxCalls: number;
};

type LlmMappingModelChoice = {
  model: string;
  tier: "fast" | "complex";
  reason: string;
};

type LlmCandidateFact = {
  concept: string;
  label: string;
  statement: "income" | "balance";
  values: Record<string, number>;
};

type ResolveContext = {
  duration: Map<string, Map<string, FactSource>>;
  instant: Map<string, Map<string, FactSource>>;
};

type PipelineLayer =
  | "edgar_extraction"
  | "concept_normalization"
  | "template_profile_detection"
  | "row_classification"
  | "cell_write_policy"
  | "formula_evaluation"
  | "validation_tie_out";

type TemplateProfileKind = "owl_standard" | "financial_company" | "generic";

type TemplateProfile = {
  kind: TemplateProfileKind;
  confidence: "high" | "medium" | "low";
  rationale: string[];
  sheetName: string;
  hasSegmentAnalysis: boolean;
};

type NormalizedMetricKey =
  | "revenue"
  | "net_revenue"
  | "cogs"
  | "gross_profit"
  | "sga"
  | "rd"
  | "da"
  | "ebit"
  | "interest_income"
  | "interest_expense"
  | "taxes"
  | "net_income"
  | "net_income_common"
  | "assets"
  | "liabilities"
  | "equity"
  | "debt"
  | "dividends"
  | "share_repurchases"
  | "basic_shares"
  | "diluted_shares";

type NormalizedHistoricalValue = {
  metric: NormalizedMetricKey | string;
  period: string;
  value: number | null;
  sources: FactSource[];
  statement: "income" | "balance" | "cash_flow" | "segment" | "support";
  periodType: "duration" | "instant";
  mappingType: "direct" | "derived" | "grouped" | "residual" | "missing";
  confidence: "high" | "medium" | "low";
  rationale: string;
};

type NormalizedHistoricalsPackage = {
  company: CompanyMatch;
  profile: TemplateProfile;
  periods: string[];
  metrics: Map<string, Map<string, NormalizedHistoricalValue>>;
  segments: SegmentRevenue[];
  diagnostics: Array<{ layer: PipelineLayer; severity: "info" | "warning" | "error"; message: string }>;
};

type InlineContext = {
  period: string | null;
  instant: boolean;
  start?: string;
  end?: string;
  hasDimensions?: boolean;
};

type SegmentRevenue = {
  label: string;
  values: Map<string, number>;
  operatingIncome: Map<string, number>;
  depreciationAmortization: Map<string, number>;
};

type FilingRef = {
  accessionNumber: string;
  filingDate: string;
  form: string;
  primaryDocument: string;
};

type MappingAuditRow = {
  sheetName: string;
  cell: string;
  modelRowLabel: string;
  section?: string;
  period: string;
  valueWritten: number;
  mappingType: RowClassification | "segment" | "calculated" | "derived" | "residual" | "skipped" | "formula preserved" | "formula updated" | "validation only";
  conceptsUsed: string;
  secLabels?: string;
  sourceStatement: string;
  accession: string;
  sourceUrl: string;
  filingForm?: string;
  filedDate?: string;
  startDate?: string;
  endDate?: string;
  cellWritable: boolean;
  formulaPreserved: boolean;
  formulaStatus?: string;
  writeBlockedReason: string;
  signConvention: string;
  confidence: "high" | "medium" | "low";
  validationStatus: string;
  notes: string;
};

type RowClassification = "direct" | "grouped" | "partial" | "formula" | "unused";

type WorkbookSnapshot = {
  labels: Map<string, string>;
  formulas: Map<string, string>;
};

type WriteDecision = {
  writable: boolean;
  reason?: string;
  formulaPreserved: boolean;
};

type ModelRowContext = {
  sheetName: string;
  row: number;
  label: string;
  sectionHeader?: string;
  previousLabel?: string;
  nextLabel?: string;
  indentation: number;
  hasHistoricalFormula: boolean;
  hasHardcodedInput: boolean;
  hasNetRevenueInterestExpenseAbove: boolean;
  subtotalFormula?: string;
  projectedColumns: number;
  signConvention: 1 | -1;
};

const SEC_HEADERS = {
  "User-Agent": process.env.SEC_USER_AGENT || "HistoricalsSolver/0.1 contact@example.com",
  Accept: "application/json"
};

const OPENROUTER_CHAT_COMPLETIONS_URL = process.env.OPENROUTER_CHAT_COMPLETIONS_URL || "https://openrouter.ai/api/v1/chat/completions";
const OPENROUTER_APP_TITLE = process.env.OPENROUTER_APP_TITLE || "Historicals Solver";
const OPENROUTER_SITE_URL = process.env.OPENROUTER_SITE_URL || "http://localhost:3000";
const LLM_MAPPING_FAST_MODEL = process.env.LLM_MAPPING_FAST_MODEL || process.env.LLM_MAPPING_MODEL || "google/gemma-4-26b-a4b-it";
const LLM_MAPPING_COMPLEX_MODEL =
  process.env.LLM_MAPPING_COMPLEX_MODEL || process.env.LLM_MAPPING_STRONG_MODEL || process.env.LLM_MAPPING_MODEL || "openai/gpt-4o-mini";
const LLM_MAPPING_MAX_CALLS = Number(process.env.LLM_MAPPING_MAX_CALLS || 24);
const LLM_MAPPING_MIN_CANDIDATE_SCORE = Number(process.env.LLM_MAPPING_MIN_CANDIDATE_SCORE || 2);
const LLM_MAPPING_CANDIDATE_LIMIT = Number(process.env.LLM_MAPPING_CANDIDATE_LIMIT || 80);
const LLM_MAPPING_COMPLEX_SCORE = Number(process.env.LLM_MAPPING_COMPLEX_SCORE || 4);

const BLUE_FONT_COLORS = new Set(["FF0000FF", "FF0070C0", "FF0563C1", "FF0000EE"]);
const MODEL_SHEET = "Model";
const SEGMENT_SHEET = "Segment Analysis";
const LABEL_COLUMNS = [1, 2, 3, 4, 5, 6, 7, 8];
const MAPPING_AUDIT_SHEET = "Mapping Audit";

const C = {
  revenue: [
    "Revenues",
    "RevenueFromContractWithCustomerExcludingAssessedTax",
    "SalesRevenueNet",
    "NetSales",
    "SalesRevenueGoodsNet",
    "SalesRevenueServicesNet",
    "OperatingLeasesIncomeStatementLeaseRevenue",
    "RealEstateRevenueNet"
  ],
  netRevenue: ["RevenuesNetOfInterestExpense"],
  grossProfit: ["GrossProfit"],
  operatingIncome: ["OperatingIncomeLoss", "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest"],
  cogs: [
    "CostOfRevenue",
    "CostOfGoodsAndServicesSold",
    "CostOfGoodsSold",
    "CostOfGoodsAndServiceExcludingDepreciationDepletionAndAmortization",
    "CostOfRevenueExcludingDepreciationDepletionAndAmortization",
    "DirectCostsOfLeasedAndRentedPropertyOrEquipment"
  ],
  healthCareCosts: [
    "PolicyholderBenefitsAndClaimsIncurredHealthCare",
    "PolicyholderBenefitsAndClaimsIncurredNet",
    "MedicalCosts",
    "BenefitsLossesAndExpenses",
    "PharmacyAndOtherServiceCosts"
  ],
  sga: ["SellingGeneralAndAdministrativeExpense", "SellingAndMarketingExpense", "SalesAndMarketingExpense", "GeneralAndAdministrativeExpense"],
  rd: ["ResearchAndDevelopmentExpense", "ResearchAndDevelopmentExpenseExcludingAcquiredInProcessCost"],
  da: ["DepreciationDepletionAndAmortization", "DepreciationAndAmortization", "DepreciationDepletionAndAmortizationExpense", "Depreciation"],
  interestIncome: ["InterestIncomeExpenseNonOperatingNet", "InterestIncomeNonOperating"],
  interestExpense: ["InterestExpenseOperating", "InterestExpenseNonOperating", "InterestExpense"],
  impairment: ["GoodwillImpairmentLosses", "ImpairmentOfGoodwillAndIntangibleAssets", "GoodwillAndIntangibleAssetImpairment", "ImpairmentOfRealEstate", "ImpairmentOfLongLivedAssetsToBeDisposedOf"],
  otherNonOp: ["OtherNonoperatingIncomeExpense", "OtherIncome", "OtherExpense"],
  taxes: ["IncomeTaxExpenseBenefit"],
  netIncome: ["NetIncomeLoss", "ProfitLoss"],
  creditLossProvision: [
    "ProvisionForLoanLeaseAndOtherLosses",
    "ProvisionForCreditLosses",
    "ProvisionForDoubtfulAccounts",
    "ProvisionForBadDebtExpense",
    "ProvisionForLoanAndLeaseLosses"
  ],
  unrealizedDebtSecurities: [
    "DebtSecuritiesAvailableForSaleUnrealizedGainLoss",
    "DebtSecuritiesAvailableForSaleRealizedGainLoss",
    "GainsLossesOnSalesOfDebtSecuritiesAvailableForSale",
    "GainsLossesOnSalesOfDebtSecurities"
  ],
  foreignCurrencyAdjustments: [
    "ForeignCurrencyTransactionGainLossBeforeTax",
    "ForeignCurrencyTransactionGainLossUnrealized",
    "ForeignCurrencyTransactionGainLoss",
    "ForeignCurrencyTranslationAdjustment"
  ],
  pensionAdjustments: [
    "DefinedBenefitPlanNetPeriodicBenefitCost",
    "OtherComprehensiveIncomeLossPensionAndOtherPostretirementBenefitPlansAdjustmentNetOfTax",
    "PensionAndOtherPostretirementDefinedBenefitPlansLiabilityCurrent"
  ],
  sbc: ["ShareBasedCompensation"],
  capex: ["PaymentsToAcquirePropertyPlantAndEquipment"],
  dividends: ["PaymentsOfDividends", "PaymentsOfDividendsCommonStock"],
  repurchases: ["PaymentsForRepurchaseOfCommonStock"],
  workingCapital: ["IncreaseDecreaseInOperatingAssetsAndLiabilities", "IncreaseDecreaseInOperatingCapital", "IncreaseDecreaseInWorkingCapital"],
  longTermItems: ["OtherOperatingActivitiesCashFlowStatement", "OtherNoncashIncomeExpense", "DeferredIncomeTaxExpenseBenefit"],
  revolverIssuanceRepayment: [
    "ProceedsFromRepaymentsOfShortTermDebt",
    "ProceedsFromRepaymentsOfShortTermDebtMaturingInMoreThanThreeMonths",
    "ProceedsFromLinesOfCredit",
    "RepaymentsOfLinesOfCredit"
  ],
  debtIssuance: ["ProceedsFromIssuanceOfLongTermDebt", "ProceedsFromLongTermDebt", "ProceedsFromBorrowings", "ProceedsFromDebt"],
  debtRepayment: ["RepaymentsOfLongTermDebt", "RepaymentsOfDebt", "PaymentsOfLongTermDebt", "RepaymentsOfBorrowings"],
  equityIssuance: [
    "ProceedsFromIssuanceOfCommonStock",
    "ProceedsFromStockOptionsExercised",
    "ProceedsFromIssuanceOfSharesUnderIncentiveAndShareBasedCompensationPlans"
  ],
  noncontrollingInterestChange: [
    "PaymentsToAcquireAdditionalInterestsInSubsidiaries",
    "PaymentsToNoncontrollingInterests",
    "ProceedsFromSaleOfInterestInSubsidiaries",
    "ProceedsFromNoncontrollingInterests"
  ],
  fxCashEffect: [
    "EffectOfExchangeRateOnCashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents",
    "EffectOfExchangeRateOnCashAndCashEquivalents"
  ],
  basicShares: ["WeightedAverageNumberOfSharesOutstandingBasic"],
  dilutedShares: ["WeightedAverageNumberOfDilutedSharesOutstanding"],
  cash: [
    "CashAndCashEquivalentsAtCarryingValue",
    "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents",
    "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalentsIncludingDisposalGroupAndDiscontinuedOperations"
  ],
  currentInvestments: ["MarketableSecuritiesCurrent", "ShortTermInvestments"],
  receivables: ["AccountsReceivableNetCurrent", "AccountsReceivableNet", "TradeAccountsReceivableNetCurrent", "ReceivablesNetCurrent"],
  cardReceivables: [
    "FinancingReceivableRecordedInvestmentLineOfCreditAndCreditCardReceivables",
    "FinanceReceivablesCreditCardNet",
    "CreditCardReceivables",
    "CardMemberReceivables"
  ],
  loans: [
    "LoansReceivableHeldForInvestment",
    "LoansAndLeasesReceivableNetReportedAmount",
    "FinanceReceivablesNet",
    "FinancingReceivableRecordedInvestment"
  ],
  inventory: ["InventoryNet"],
  currentAssets: ["AssetsCurrent"],
  ppe: [
    "PropertyPlantAndEquipmentAndOperatingLeaseRightofUseAssetAfterAccumulatedDepreciationAndAmortization",
    "PropertyPlantAndEquipmentAndFinanceLeaseRightOfUseAssetAfterAccumulatedDepreciationAndAmortization",
    "PropertyPlantAndEquipmentNet",
    "PropertyAndEquipmentNet",
    "RealEstateInvestmentPropertyNet",
    "RealEstateInvestments"
  ],
  intangibles: ["FiniteLivedIntangibleAssetsNet", "IntangibleAssetsNetExcludingGoodwill"],
  goodwill: ["Goodwill"],
  assets: ["Assets"],
  ap: ["AccountsPayableCurrent", "AccountsPayableAndAccruedLiabilitiesCurrent", "AccountsPayableAndAccruedLiabilitiesCurrentAndNoncurrent"],
  accrued: ["AccruedLiabilitiesCurrent", "AccruedIncomeTaxesCurrent", "EmployeeRelatedLiabilitiesCurrent"],
  customerDeposits: ["Deposits", "CustomerDeposits", "DepositsLiabilities", "InterestBearingDepositsInDomesticOffices"],
  currentLiabilities: ["LiabilitiesCurrent"],
  currentDebt: ["LongTermDebtCurrent", "CurrentPortionOfLongTermDebt", "ShortTermBorrowings", "ShortTermBorrowingsCurrent"],
  totalDebt: [
    "ShortTermBorrowings",
    "LongTermDebtCurrent",
    "LongTermDebtNoncurrent",
    "LongTermDebtAndFinanceLeaseObligationsCurrent",
    "LongTermDebtAndFinanceLeaseObligationsNoncurrent",
    "LongTermDebtAndCapitalLeaseObligationsCurrent",
    "LongTermDebtAndCapitalLeaseObligations",
    "LongTermDebtAndCapitalLeaseObligationsIncludingCurrentMaturities",
    "LongTermDebt",
    "DebtAndCapitalLeaseObligations",
    "DebtInstrumentCarryingAmount"
  ],
  deferredTaxLiability: ["DeferredTaxLiabilitiesNoncurrent"],
  liabilities: ["Liabilities"],
  equity: ["StockholdersEquity", "StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest"],
  commonApic: ["CommonStocksIncludingAdditionalPaidInCapital", "AdditionalPaidInCapitalCommonStocks"],
  retained: ["RetainedEarningsAccumulatedDeficit"],
  treasury: ["TreasuryStockCommonValue", "TreasuryStockValue"],
  aoci: ["AccumulatedOtherComprehensiveIncomeLossNetOfTax"],
  nci: ["MinorityInterest", "NoncontrollingInterestInConsolidatedEntity"]
};

const BROKER_DEALER_RECEIVABLES = [
  "ReceivablesFromBrokersDealersAndClearingOrganizations",
  "ReceivablesFromCustomers",
  "FeesInterestAndOther"
];

const BROKER_DEALER_CURRENT_ASSETS = [
  "FinancialInstrumentsOwnedAtFairValue",
  "InvestmentOwnedAtFairValue",
  "InvestmentsInAffiliatesSubsidiariesAssociatesAndJointVentures",
  "SecuritiesBorrowed",
  "SecuritiesPurchasedUnderAgreementsToResell",
  "SecuritiesReceivedAsCollateral"
];

const BROKER_DEALER_PAYABLES = ["PayablesToBrokerDealersAndClearingOrganizations", "PayablesToCustomers"];

const BROKER_DEALER_OTHER_CURRENT_LIABILITIES = [
  "FinancialInstrumentsSoldNotYetPurchasedAtFairValue",
  "SecuritiesSoldUnderAgreementsToRepurchase",
  "OtherSecuredFinancings",
  "ObligationToReturnSecuritiesReceivedAsCollateral",
  "AccountsPayableAndAccruedLiabilitiesCurrentAndNoncurrent",
  "OperatingLeaseLiability",
  "ContractWithCustomerLiabilityCurrent"
];

const NON_COMPENSATION_EXPENSE_CONCEPTS = [
  "FloorBrokerageExchangeAndClearanceFees",
  "UnderwritingCosts",
  "CommunicationsAndInformationTechnology",
  "OccupancyNet",
  "BusinessDevelopment",
  "ProfessionalFees"
];

const OTHER_OPERATING_EXPENSE_CONCEPTS = ["CostOfGoodsAndServicesSold", "OtherExpenses"];

const NONCONTROLLING_INCOME_CONCEPTS = [
  "IncomeLossFromContinuingOperationsAttributableToNoncontrollingEntity",
  "NetIncomeLossAttributableToNonredeemableNoncontrollingInterest",
  "NetIncomeLossAttributableToRedeemableNoncontrollingInterest",
  "NetIncomeLossAttributableToNoncontrollingInterest"
];

const POST_TAX_ADJUSTMENT_CONCEPTS = [
  "PreferredStockDividendsIncomeStatementImpact",
  "UndistributedEarningsLossAllocatedToParticipatingSecuritiesBasic",
  "ConvertiblePreferredDividendsNetOfTax",
  "RedeemablePreferredStockDividends"
];

const COMMON_SHAREHOLDER_INCOME_CONCEPTS = ["NetIncomeLossAvailableToCommonStockholdersBasic", "NetIncomeLossAvailableToCommonStockholdersDiluted"];

const CONTINUING_NET_INCOME_CONCEPTS = ["IncomeLossFromContinuingOperationsIncludingPortionAttributableToNoncontrollingInterest", "NetIncomeLoss"];

const PRETAX_INCOME_CONCEPTS = [
  "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
  "IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments",
  "IncomeLossFromContinuingOperationsBeforeIncomeTaxes"
];

const COMPENSATION_CONCEPTS = [
  "CompensationAndBenefitsExpense",
  "CompensationAndBenefits",
  "LaborAndRelatedExpense",
  "SalariesWagesAndBenefits",
  "SalariesAndWages",
  "CommissionsExpense",
  "EmployeeRelatedLiabilitiesCurrent"
];

const INVESTMENT_ASSET_CONCEPTS = [
  "Investments",
  "InvestmentsCurrent",
  "InvestmentsNoncurrent",
  "EquityMethodInvestments",
  "InvestmentsInAffiliatesSubsidiariesAssociatesAndJointVentures",
  "OtherInvestments",
  "ConsolidatedVariableInterestEntitiesAssets"
];

const PURCHASES_OF_INVESTMENTS_CONCEPTS = [
  "PaymentsToAcquireInvestments",
  "PaymentsToAcquireAvailableForSaleSecurities",
  "PaymentsToAcquireHeldToMaturitySecurities",
  "PaymentsToAcquireEquityMethodInvestments",
  "PaymentsToAcquireOtherInvestments"
];

const ACQUISITION_CONCEPTS = [
  "PaymentsToAcquireBusinessesNetOfCashAcquired",
  "PaymentsToAcquireBusinessesAndInterestInAffiliates",
  "BusinessAcquisitionPurchasePrice",
  "PaymentsForProceedsFromOtherInvestingActivities"
];

function row(
  rowNumber: number,
  label: string,
  statement: FillRow["statement"],
  kind: FillRow["kind"],
  concepts: string[],
  sign: 1 | -1 = 1,
  scale = 1_000_000,
  comment?: string,
  classification: RowClassification = concepts.length ? "direct" : "unused"
): FillRow {
  return { row: rowNumber, label, classification, statement, kind, concepts, sign, scale, comment };
}

function plug(
  rowNumber: number,
  label: string,
  statement: FillRow["statement"],
  kind: FillRow["kind"],
  resolver: FillRow["resolver"],
  classification: RowClassification = "grouped",
  options: Pick<FillRow, "allowBlankHistoricalInput" | "onlyBlankHistoricalInput" | "noFillComment"> = {}
): FillRow {
  return { row: rowNumber, label, classification, statement, kind, resolver, scale: 1_000_000, ...options };
}

function discoverFillRows(sheet: ExcelJS.Worksheet, columns: number[], periodInfos: Array<{ period: string; isEstimate: boolean }> = []) {
  const rows: FillRow[] = [];
  const seen = new Set<number>();

  for (let rowNumber = 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
    if (isCashFlowStatementBlockRow(sheet, rowNumber)) continue;
    const modelContext = modelRowContext(sheet, rowNumber, columns, periodInfos);
    const label = rowLabel(sheet, rowNumber);
    if (!label || seen.has(rowNumber)) continue;
    const fillRow = fillRowForContext(modelContext) ?? unusedRow(modelContext);
    if (!modelContext.hasHardcodedInput && !modelContext.hasHistoricalFormula && !fillRow.allowBlankHistoricalInput) continue;
    fillRow.modelContext = modelContext;
    rows.push(fillRow);
    seen.add(rowNumber);
  }

  return rows;
}

function isCashFlowStatementBlockRow(sheet: ExcelJS.Worksheet, rowNumber: number) {
  for (let row = rowNumber; row >= Math.max(1, rowNumber - 80); row -= 1) {
    const label = rowLabel(sheet, row);
    if (!label) continue;
    if (isCashFlowStatementHeader(label)) return true;
    if (isCashFlowStatementBoundary(label)) return false;
  }
  return false;
}

function isCashFlowStatementHeader(label: string) {
  return normalize(label) === normalize("Cash Flow Statement") || normalize(label) === normalize("Cashflow Statement");
}

function isCashFlowStatementBoundary(label: string) {
  if (isCashFlowStatementHeader(label)) return false;
  return /income statement|balance sheet|working capital schedule|drivers|analysis|schedule|assumptions|debt and interest|shareholder|shareholders|pp&e|ppe/i.test(label);
}

function fillRowForLabel(rowNumber: number, label: string): FillRow | null {
  return fillRowForContext({
    sheetName: MODEL_SHEET,
    row: rowNumber,
    label,
    indentation: 0,
    hasHistoricalFormula: false,
    hasHardcodedInput: true,
    hasNetRevenueInterestExpenseAbove: false,
    projectedColumns: 0,
    signConvention: 1
  });
}

function fillRowForContext(context: ModelRowContext): FillRow | null {
  const { row: rowNumber, label } = context;
  const key = normalize(label);
  const has = (...aliases: string[]) => aliases.some((alias) => key === normalize(alias));
  const includes = (...aliases: string[]) => aliases.some((alias) => key.includes(normalize(alias)));
  const around = normalize([context.sectionHeader, context.previousLabel, context.nextLabel].filter(Boolean).join(" "));
  const aroundIncludes = (...aliases: string[]) => aliases.some((alias) => around.includes(normalize(alias)));
  const hasRevenue = has("Revenue", "Revenues", "Total Revenue", "Total Revenues", "Sales", "Net Sales", "Net Revenue", "Total Sales") || has("Rental Revenues", "Rental Revenue");
  const hasNetRevenue = has("Net Revenue", "Revenue Net of Interest Expense", "Revenues Net of Interest Expense");

  if (has("Beginning Balance")) {
    if (inSection(context, "Revolver Balance")) {
      return plug(rowNumber, label, "support", "instant", resolveBeginningRevolverBalance, "partial", {
        allowBlankHistoricalInput: true,
        onlyBlankHistoricalInput: true,
        noFillComment: "Cannot find exact revolver beginning balance plug in EDGAR, find manually."
      });
    }
    if (inSection(context, "Total Debt Balance", "Debt and Interest Schedule")) {
      return plug(rowNumber, label, "support", "instant", resolveBeginningTotalDebtBalance, "partial", {
        allowBlankHistoricalInput: true,
        onlyBlankHistoricalInput: true,
        noFillComment: "Cannot find exact debt beginning balance plug in EDGAR, find manually."
      });
    }
    if (inSection(context, "Retained Earnings")) {
      return plug(rowNumber, label, "support", "instant", resolveBeginningRetainedEarningsBalance, "partial", {
        allowBlankHistoricalInput: true,
        onlyBlankHistoricalInput: true,
        noFillComment: "Cannot find exact retained-earnings beginning balance plug in EDGAR, find manually."
      });
    }
    if (inSection(context, "AOCI Assumptions", "Accumulated Other Comprehensive Income", "AOCI")) {
      return plug(rowNumber, label, "support", "instant", resolveBeginningAociBalance, "partial", {
        allowBlankHistoricalInput: true,
        onlyBlankHistoricalInput: true,
        noFillComment: "Cannot find exact AOCI beginning balance plug in EDGAR, find manually."
      });
    }
  }

  if (context.hasHistoricalFormula && !context.hasHardcodedInput) return formulaRow(context);
  if (hasNetRevenue) return row(rowNumber, label, "income", "duration", C.netRevenue, 1, 1_000_000, "Mapped to SEC revenues net of interest expense when reported.");
  if (hasRevenue) return row(rowNumber, label, "income", "duration", C.revenue, 1, 1_000_000, "Mapped to the closest SEC total revenue concept.");
  if (has("Gross Profit", "Gross Margin Dollars")) return row(rowNumber, label, "income", "duration", C.grossProfit, 1, 1_000_000, "Mapped to SEC gross profit.");
  if (has("Operating Income", "Operating Income (Loss)", "Income From Operations")) {
    return row(rowNumber, label, "income", "duration", C.operatingIncome, 1, 1_000_000, "Mapped to SEC operating income/loss or nearest pre-tax operating profit concept.");
  }
  if (has("Cost of Goods Sold", "Cost of Goods & Services Sold", "Cost of Revenue", "Cost of Sales", "Property Taxes and Insurance")) return row(rowNumber, label, "income", "duration", C.cogs, -1, 1_000_000, "Mapped to SEC cost of revenue / cost of sales.");
  if (has("Pharmacy and Other Service Costs", "Medical Costs and Other")) return row(rowNumber, label, "income", "duration", C.healthCareCosts, -1, 1_000_000, "Mapped to reported healthcare, claims, pharmacy, or service cost concepts.");
  if (
    has(
      "Selling, General & Administration (SG&A)",
      "Selling, Geneal & Administrative (SG&A)",
      "Selling General & Administrative",
      "Selling, General, and Administrative",
      "SG&A",
      "Sales and Marketing",
      "Selling and Marketing",
      "General and Administrative"
    )
  ) {
    return plug(rowNumber, label, "income", "duration", resolveSellingGeneralAdministrativeExpense, "grouped");
  }
  if (has("Research & Development (R&D)", "Research and Development")) return row(rowNumber, label, "income", "duration", C.rd, -1);
  if (has("Compensation and Benefits", "Compensation, Commissions, and Benefits", "Employee Compensation & Benefits")) {
    return plug(rowNumber, label, "income", "duration", resolveCompensationExpense, "direct");
  }
  if (has("Non-Compensation Expenses") || (includes("Non Compensation", "Non-Compensation") && aroundIncludes("Expense", "Operating"))) {
    return plug(rowNumber, label, "income", "duration", resolveNonCompensationExpense);
  }
  if (
    has("Other Operating Expenses", "Other Operating Expense") ||
    (has("Other", "Other Expenses", "Other Expense") && context.indentation > 0 && aroundIncludes("Operating Expense", "Operating Expenses", "Noninterest Expense", "Expenses"))
  ) {
    return plug(rowNumber, label, "income", "duration", resolveOtherOperatingExpenseGroup);
  }
  if (has("Depreciation Expense") && inPpeDepreciationScheduleContext(context)) {
    return reviewRow(context, "Needs review: PP&E depreciation is a roll-forward input and was not overwritten with consolidated depreciation and amortization.");
  }
  if (has("Depreciation & Amortization (incl. in SG&A)")) return row(rowNumber, label, "income", "duration", C.da, -1);
  if (has("Depreciation & Amortization", "Depreciation and Amortization", "Depreciation Expense")) {
    return reviewRow(context, "Needs review: consolidated D&A is often a cash-flow disclosure and was not overwritten into the income statement unless explicitly included in an operating expense row.");
  }
  if (has("Amortization Expense")) return row(rowNumber, label, "support", "duration", ["AmortizationOfIntangibleAssets"], -1);
  if (has("Other Operating Income (Expense)")) return plug(rowNumber, label, "income", "duration", resolveOtherOperatingIncomeExpense);
  if (has("Total Provisions for Credit Losses", "Provision for Credit Losses")) return row(rowNumber, label, "income", "duration", C.creditLossProvision, -1);
  if (has("Interest Income")) return row(rowNumber, label, "income", "duration", C.interestIncome, 1, 1_000_000, "Mapped to SEC interest income.");
  if (/\binterest\s*\(\s*expense\s*\)/i.test(label)) {
    if (context.hasNetRevenueInterestExpenseAbove) {
      return plug(rowNumber, label, "income", "duration", resolveNonOperatingInterestExpenseAfterNetRevenue, "direct");
    }
    return plug(rowNumber, label, "income", "duration", resolveInterestExpense, "direct");
  }
  if (has("Interest Expense")) return plug(rowNumber, label, "income", "duration", resolveInterestExpense, "direct");
  if (has("Goodwill Impairment", "Impairment of Investments in Real Estate", "Asset Impairment")) return row(rowNumber, label, "income", "duration", C.impairment, -1);
  if (has("Gain on Sale of Business (Loss)") || includes("Gain on disposition", "Gain (Loss) on disposition", "Gain on sale")) {
    return row(rowNumber, label, "income", "duration", ["GainLossOnSaleOfBusiness", "GainLossOnSaleOfAssets", "GainLossOnDispositionOfAssets", "GainLossOnDispositionOfAssets1", "GainsLossesOnSalesOfInvestmentRealEstate"], 1);
  }
  if (has("Equity in (loss) earnings of unconsolidated entities", "Equity in Earnings of Unconsolidated Entities")) {
    return row(rowNumber, label, "income", "duration", ["IncomeLossFromEquityMethodInvestments"], 1);
  }
  if (has("Loss from early extinguishment of debt", "Loss on Extinguishment of Debt")) {
    return row(rowNumber, label, "income", "duration", ["GainsLossesOnExtinguishmentOfDebt", "ExtinguishmentOfDebtGainLossNetOfTax"], -1);
  }
  if (has("Other Non-Operating Income (Expense)", "Other Nonoperating Income (Expense)", "Other Income (Expense)", "Other Expense (Income)")) return row(rowNumber, label, "income", "duration", C.otherNonOp);
  if (has("Income Tax Benefit (Expense)", "Income Tax Expense", "Income Tax Provision (Expense)", "Income Tax")) return plug(rowNumber, label, "income", "duration", resolveIncomeTaxExpense, "direct");
  if (has("Net Unrealized Debt Securities Gains (Losses)")) return row(rowNumber, label, "income", "duration", C.unrealizedDebtSecurities);
  if (has("FX Adjustments")) return row(rowNumber, label, "income", "duration", C.foreignCurrencyAdjustments);
  if (has("Net Unrealized Pension and Other Benefits")) return row(rowNumber, label, "income", "duration", C.pensionAdjustments);
  if (has("Pre-Tax Adjustments")) return reviewRow(context, "Split / partial match: EDGAR adjustment detail is not consistently available for this model row. Needs review.");
  if (has("Post-Tax Adjustments", "Preferred Stock Dividend")) return plug(rowNumber, label, "income", "duration", resolvePostTaxAdjustments, "direct");
  if (has("Discontinued Operations")) {
    if (aroundIncludes("Post-Tax Adjustments", "Income (Loss) due to Non-Controlling Interest", "Net Income Available to Common Shareholders")) {
      return plug(rowNumber, label, "income", "duration", resolveDiscontinuedOperationsBridge, "grouped");
    }
    return row(rowNumber, label, "income", "duration", ["IncomeLossFromDiscontinuedOperationsNetOfTax"]);
  }
  if (has("Income (Loss) due to Non-Controlling Interest", "Income Loss Due To Non Controlling Interest")) {
    if (aroundIncludes("Discontinued Operations", "Net Income Available to Common Shareholders", "Post-Tax Adjustments")) {
      return plug(rowNumber, label, "income", "duration", resolveCommonShareholderNciBridge, "grouped");
    }
    return plug(rowNumber, label, "income", "duration", resolveNoncontrollingIncome);
  }

  if (has("Cash & Cash Equivalents", "Cash and Cash Equivalents", "Cash and Equivalents", "Cash")) {
    return plug(rowNumber, label, "balance", "instant", resolveCashAndCurrentInvestments, "grouped");
  }
  if (has("Accounts Receivable", "Accounts Receivable, Net", "Trade Receivables", "Fees Receivable")) return plug(rowNumber, label, "balance", "instant", resolveAccountsReceivable);
  if (has("Card Member Receivables", "Card Member Recievables")) return row(rowNumber, label, "balance", "instant", C.cardReceivables);
  if (has("Inventory")) return row(rowNumber, label, "balance", "instant", C.inventory);
  if (has("Prepaid & Other Current Assets", "Prepaid and Other Current Assets", "Other Current Assets", "Prepaids and Other Current Assets")) {
    return plug(rowNumber, label, "balance", "instant", resolvePrepaidAndOtherCurrentAssets);
  }
  if (has("PP&E, Net", "Property Plant and Equipment Net", "Property and Equipment, Net", "Property, Plant and Equipment, Net", "Real Estate Investments", "Real Estate Investment Property, Net")) {
    return plug(rowNumber, label, "balance", "instant", resolvePpe);
  }
  if (has("Intangible Assets, Net", "Intangibles, Net")) return plug(rowNumber, label, "balance", "instant", resolveIntangibleAssets);
  if (has("Goodwill")) return row(rowNumber, label, "balance", "instant", C.goodwill);
  if (has("Card Member Loans")) return row(rowNumber, label, "balance", "instant", C.loans);
  if (has("Investments and Assets of Consolidated VIEs", "Investments", "Investment Securities", "Investments and Assets of Consolidated Variable Interest Entities")) {
    return row(rowNumber, label, "balance", "instant", INVESTMENT_ASSET_CONCEPTS);
  }
  if (has("Other Non-Current Assets", "Other Long-Term Assets", "Other LT Assets")) return plug(rowNumber, label, "balance", "instant", resolveOtherNonCurrentAssets);
  if (has("Accounts Payable", "Accounts Payable and Accrued Liabilities", "Accounts Payable & Accrued Liabilities", "Pharmacy Costs Payable")) return plug(rowNumber, label, "balance", "instant", resolveAccountsPayable);
  if (has("Securities Loaned")) return row(rowNumber, label, "balance", "instant", ["SecuritiesLoaned"]);
  if (has("Accrued Liabilities", "Accrued Expenses", "Accrued Expenses and Other Current Liabilities")) return plug(rowNumber, label, "balance", "instant", resolveAccruedLiabilities);
  if (has("Customer Deposits")) return row(rowNumber, label, "balance", "instant", C.customerDeposits);
  if (has("Other Current Liabilities", "Other Current Liabs")) {
    return plug(rowNumber, label, "balance", "instant", resolveOtherCurrentLiabilities);
  }
  if (has("Tax Receivable Agreement Payables")) {
    return row(rowNumber, label, "balance", "instant", ["TaxReceivableAgreementLiability", "TaxReceivableAgreementLiabilityCurrent", "OtherLiabilitiesCurrent"]);
  }
  if (has("Short Term Borrowings", "Short-term Borrowings", "Short-Term Debt", "Short Term Debt", "Current Debt")) return row(rowNumber, label, "balance", "instant", ["OtherShortTermBorrowings", "ShortTermBorrowings", "LongTermDebtCurrent"]);
  if (has("Revolver")) return row(rowNumber, label, "balance", "instant", ["RevolvingCreditFacility", "LineOfCreditFacilityCurrentBorrowings"]);
  if (has("LT Debt (Incl. Current Portion)", "Long Term Debt", "Long-Term Debt") || includes("LT Debt")) {
    return plug(rowNumber, label, "balance", "instant", resolveLongTermDebtInclCurrentPortion);
  }
  if (has("Total Debt")) {
    return plug(rowNumber, label, "balance", "instant", resolveTotalDebt);
  }
  if (has("Deferred Income Taxes")) return plug(rowNumber, label, "balance", "instant", resolveDeferredTaxLiability);
  if (has("Other Non-Current Liabilities")) return plug(rowNumber, label, "balance", "instant", resolveOtherNonCurrentLiabilities);
  if (has("Mezzanine Equity")) return row(rowNumber, label, "balance", "instant", ["RedeemableNoncontrollingInterestEquityCarryingAmount"]);
  if (has("Common Stock & APIC", "Common Stock and APIC", "Common Stock and Additional Paid-In Capital", "Common Stock & Additional Paid-In Capital")) return plug(rowNumber, label, "balance", "instant", resolveCommonStockAndApic);
  if (has("Retained Earnings", "Accumulated Deficit")) return row(rowNumber, label, "balance", "instant", C.retained);
  if (has("Treasury Stock", "Treasury & Preferred Stock")) return plug(rowNumber, label, "balance", "instant", resolveTreasuryAndPreferredStock);
  if (has("Accumulated Other Comprehensive Income (AOCI)", "AOCI")) return row(rowNumber, label, "balance", "instant", C.aoci);
  if (has("Noncontrolling Interests", "Non-Controlling Interests")) return row(rowNumber, label, "balance", "instant", C.nci);

  if (has("(Increase)/Decrease in Working Capital", "Increase / (Decrease) in Working Capital")) return row(rowNumber, label, "support", "duration", C.workingCapital);
  if (has("(Increase)/Decrease in LT Items", "Increase / (Decrease) in LT Items")) return row(rowNumber, label, "support", "duration", C.longTermItems);
  if (has("Capital Expenditures", "Capex")) {
    if (inPpeDepreciationScheduleContext(context)) {
      return reviewRow(context, "Needs review: PP&E schedule capex is a roll-forward input and was not overwritten with generic cash-flow capex.");
    }
    return row(rowNumber, label, "support", "duration", C.capex, -1);
  }
  if (has("Purchases of Intangibles")) return row(rowNumber, label, "support", "duration", ["PaymentsToAcquireIntangibleAssets"]);
  if (has("Purchases of Investments")) return row(rowNumber, label, "support", "duration", PURCHASES_OF_INVESTMENTS_CONCEPTS);
  if (has("Acquisition / (Divestment) of Businesses", "Proceeds From/(Acquisitions of) Businesses", "Proceeds From (Acquisitions of) Businesses")) {
    if (inPpeDepreciationScheduleContext(context)) {
      return reviewRow(context, "Needs review: PP&E schedule acquisition/divestment is a roll-forward input and was not overwritten with generic acquisition cash-flow concepts.");
    }
    return row(rowNumber, label, "support", "duration", ACQUISITION_CONCEPTS);
  }
  if (has("Issuance/(Repayment) of Revolver")) return plug(rowNumber, label, "support", "duration", resolveRevolverIssuanceRepayment);
  if (has("Issuance of Debt")) return row(rowNumber, label, "support", "duration", C.debtIssuance);
  if (has("(Repayment of Debt)", "Repayment of Debt")) return row(rowNumber, label, "support", "duration", C.debtRepayment, -1);
  if (has("Issuance of Equity")) return row(rowNumber, label, "support", "duration", C.equityIssuance);
  if (has("Shares Repurchased ($ Amount)", "Share Repurchases ($ Amount)")) {
    if (inShareRepurchaseAssumptionsContext(context)) {
      return reviewRow(context, "Needs review: share repurchase assumptions were preserved because the schedule uses EDGAR issuer-purchase / repurchase-price support, not generic cash-flow repurchase deltas.");
    }
    return row(
      rowNumber,
      label,
      "support",
      "duration",
      C.repurchases,
      inSection(context, "Shareholder's Equity Schedule", "Shareholders Equity Schedule", "Share Repurchase Assumptions") ? 1 : -1,
      1_000_000,
      "Mapped to EDGAR common stock repurchases; sign follows the model row context."
    );
  }
  if (has("(Repurchase) of Equity", "Repurchase of Equity")) return row(rowNumber, label, "support", "duration", C.repurchases, -1);
  if (has("Stock-Based Comp Expense", "Stock-Based Compensation")) return row(rowNumber, label, "support", "duration", C.sbc);
  if (has("Dividends", "Dividends Issued")) return row(rowNumber, label, "support", "duration", C.dividends, -1);
  if (has("Change in Noncontrolling Interests")) return row(rowNumber, label, "support", "duration", C.noncontrollingInterestChange);
  if (has("Effect of FX Rate Changes on Cash")) return row(rowNumber, label, "support", "duration", C.fxCashEffect);
  if (has("Beginning Cash Adjustments", "Ending Cash Adjustments")) return reviewRow(context, "Unused / no confident match: cash adjustment rows are model-specific and were not filled from EDGAR.");
  if (has("Beginning Cash Balance")) return plug(rowNumber, label, "support", "instant", resolveBeginningCashBalance, "direct");
  if (has("Ending Cash Balance")) return row(rowNumber, label, "support", "instant", C.cash);
  if (has("Weighted Average Basic Shares", "Basic Shares")) {
    if (inSection(context, "Shares Outstanding Schedule") || aroundIncludes("Ending Balance - Basic", "Effects of Dilutive Securities", "Weighted Average Dilutive Shares")) {
      return reviewRow(context, "Needs review: shares outstanding schedule rows use the model's actual share roll-forward and were not overwritten with generic weighted-average share facts.");
    }
    return row(rowNumber, label, "support", "duration", C.basicShares, 1, 1_000_000);
  }
  if (has("Weighted Average Dilutive Shares", "Weighted Average Diluted Shares", "Diluted Shares")) {
    if (inSection(context, "Shares Outstanding Schedule") || aroundIncludes("Weighted Average Basic Shares", "Effects of Dilutive Securities")) {
      return reviewRow(context, "Needs review: shares outstanding schedule rows use the model's actual share roll-forward and were not overwritten with generic weighted-average share facts.");
    }
    return row(rowNumber, label, "support", "duration", C.dilutedShares, 1, 1_000_000);
  }

  return null;
}

function formulaRow(context: ModelRowContext): FillRow {
  return {
    row: context.row,
    label: context.label,
    classification: "formula",
    statement: statementFromContext(context),
    kind: "duration",
    scale: 1_000_000,
    noFillComment: "Formula row: existing model formula preserved and not overwritten.",
    modelContext: context
  };
}

function reviewRow(context: ModelRowContext, noFillComment: string): FillRow {
  return {
    row: context.row,
    label: context.label,
    classification: "partial",
    statement: statementFromContext(context),
    kind: inBalanceSheetContext(context) ? "instant" : "duration",
    scale: 1_000_000,
    noFillComment,
    modelContext: context
  };
}

function unusedRow(context: ModelRowContext): FillRow {
  return {
    row: context.row,
    label: context.label,
    classification: context.hasHistoricalFormula ? "formula" : "unused",
    statement: statementFromContext(context),
    kind: inBalanceSheetContext(context) ? "instant" : "duration",
    scale: 1_000_000,
    noFillComment: context.hasHistoricalFormula
      ? "Formula row: existing model formula preserved and not overwritten."
      : "Unused / no confident match: no reliable EDGAR line item or logical grouping was identified for this model row.",
    modelContext: context
  };
}

function resolveCompensationExpense(period: string, ctx: ResolveContext): ResolvedValue {
  const direct = first(period, ctx.duration, COMPENSATION_CONCEPTS);
  if (direct) {
    return {
      value: -Math.abs(direct.value),
      sources: [direct],
      note: "Mapped to the closest compensation, commissions, benefits, or labor-related SEC concept."
    };
  }
  return { value: null, sources: [], note: "No dedicated compensation and benefits concept was available in SEC company facts." };
}

function resolveInterestExpense(period: string, ctx: ResolveContext): ResolvedValue {
  const facts = ctx.duration.get(period);
  const revenue = first(period, ctx.duration, ["Revenues"]);
  const netRevenue = first(period, ctx.duration, ["RevenuesNetOfInterestExpense"]);
  const fallback =
    revenue && netRevenue
      ? {
          value: -Math.abs(revenue.value - netRevenue.value),
          sources: [revenue, netRevenue],
          note: "Calculated as SEC revenues less revenues net of interest expense."
        }
      : null;
  if (!facts) return fallback ?? { value: null, sources: [] };
  for (const concept of C.interestExpense) {
    const source = facts.get(concept);
    if (source && source.value !== 0) {
      return {
        value: -Math.abs(source.value),
        sources: [source],
        note: "Included SEC interest expense, including operating interest expense when that is the company's presentation."
      };
    }
  }
  const zero = first(period, ctx.duration, C.interestExpense);
  if (fallback && (!zero || zero.value === 0)) return fallback;
  return zero ? { value: 0, sources: [zero] } : { value: null, sources: [] };
}

function resolveNonOperatingInterestExpenseAfterNetRevenue(period: string, ctx: ResolveContext): ResolvedValue {
  const direct = first(period, ctx.duration, ["InterestExpenseNonOperating"]);
  if (direct) {
    return {
      value: -Math.abs(direct.value),
      sources: [direct],
      note: "Mapped only to SEC non-operating interest expense because operating interest expense is already included in the model's net revenue bridge."
    };
  }

  const netRevenue = first(period, ctx.duration, C.netRevenue);
  const operatingInterest = first(period, ctx.duration, ["InterestExpenseOperating"]);
  const genericInterest = first(period, ctx.duration, ["InterestExpense"]);
  return {
    value: 0,
    sources: compactSources([zeroSource("InterestExpenseNonOperating"), netRevenue, operatingInterest, genericInterest]),
    note:
      "Set to zero because this model already deducts interest expense above net revenue and EDGAR did not report a separate non-operating interest expense for this period."
  };
}

function resolveNonCompensationExpense(period: string, ctx: ResolveContext): ResolvedValue {
  const detail = sumWithNote(
    period,
    ctx.duration,
    NON_COMPENSATION_EXPENSE_CONCEPTS,
    "These were grouped because the model row is labeled Non-Compensation Expenses and no separate rows exist for these expense categories."
  );
  if (detail.value !== null) {
    if (isFourthQuarterPeriod(period)) {
      const derived = resolveFourthQuarterFromAnnualDetail(period, detail, (priorPeriod) => resolveNonCompensationExpense(priorPeriod, ctx));
      if (derived.value !== null) return derived;
    }
    return { ...(signed(detail, -1) ?? detail), classification: "grouped" };
  }

  const noninterestExpense = first(period, ctx.duration, ["NoninterestExpense"]);
  const compensation = resolveCompensationExpense(period, ctx);
  const depreciation = first(period, ctx.duration, C.da) ?? zeroSource(C.da[0]);
  const otherOperating = resolveOtherOperatingIncomeExpense(period, ctx);
  if (noninterestExpense && compensation.value !== null && otherOperating.value !== null) {
    return {
      value: -Math.abs(noninterestExpense.value - Math.abs(compensation.value) - depreciation.value - Math.abs(otherOperating.value)),
      sources: compactSources([noninterestExpense, compensation, depreciation, otherOperating]),
      note: "These were grouped because the model row is broader than the EDGAR non-interest expense detail available for this company.",
      classification: "grouped"
    };
  }
  return { value: null, sources: [], note: "No non-compensation expense detail was available in SEC facts." };
}

function resolveFourthQuarterFromAnnualDetail(period: string, annualDetail: ResolvedValue, priorResolver: (period: string) => ResolvedValue): ResolvedValue {
  const year = periodYearSuffix(period);
  const annualValue = annualDetail.sources.reduce((total, source) => total + Math.abs(source.derivedTotalValue ?? source.value), 0);
  const priorPeriods = [`1Q${year}`, `2Q${year}`, `3Q${year}`];
  const priorValues = priorPeriods.map((priorPeriod) => priorResolver(priorPeriod).value);
  if (!priorValues.every((value): value is number => value !== null)) return { value: null, sources: annualDetail.sources, note: annualDetail.note };
  const value = -(annualValue - priorValues.reduce((total, item) => total + Math.abs(item), 0));
  return {
    value,
    sources: [
      {
        concept: "AnnualNonCompensationExpenseBridge",
        label: "Annual non-compensation expense bridge",
        value,
        derivedTotalValue: annualValue,
        derivedTotalLabel: "Annual non-compensation expense detail",
        derivedPriorPeriods: priorPeriods
      },
      ...annualDetail.sources
    ],
    note: "Calculated from EDGAR annual non-compensation expense detail less Q1, Q2, and Q3 because the model's 4Q cell is an annual-minus-quarterly bridge.",
    classification: "grouped"
  };
}

function resolveOtherOperatingExpenseGroup(period: string, ctx: ResolveContext): ResolvedValue {
  const detail = sumWithNote(
    period,
    ctx.duration,
    [...NON_COMPENSATION_EXPENSE_CONCEPTS, ...OTHER_OPERATING_EXPENSE_CONCEPTS],
    "These were grouped because the model row is labeled Other Operating Expenses and the model does not provide separate rows for the included EDGAR expense categories."
  );
  if (detail.value !== null) return { ...(signed(detail, -1) ?? detail), classification: "grouped" };
  return resolveOtherOperatingIncomeExpense(period, ctx);
}

function resolveOtherOperatingIncomeExpense(period: string, ctx: ResolveContext): ResolvedValue {
  const detail = sumWithNote(
    period,
    ctx.duration,
    OTHER_OPERATING_EXPENSE_CONCEPTS,
    "These were grouped because the model row is broader than the EDGAR operating expense detail available for this company."
  );
  if (detail.value !== null) return { ...(signed(detail, -1) ?? detail), classification: "grouped" };
  const direct = first(period, ctx.duration, ["OtherOperatingIncomeExpenseNet"]);
  return direct ? { value: direct.value, sources: [direct] } : { value: null, sources: [] };
}

function resolveSellingGeneralAdministrativeExpense(period: string, ctx: ResolveContext): ResolvedValue {
  const broadDirect = first(period, ctx.duration, ["SellingGeneralAndAdministrativeExpense"]);
  if (broadDirect) {
    return {
      value: -Math.abs(broadDirect.value),
      sources: [broadDirect],
      note: "Mapped to EDGAR selling, general, and administrative expense.",
      classification: "direct"
    };
  }

  const grossProfit = first(period, ctx.duration, C.grossProfit);
  const operatingIncome = first(period, ctx.duration, C.operatingIncome);
  if (grossProfit && operatingIncome) {
    const rd = first(period, ctx.duration, C.rd) ?? zeroSource(C.rd[0]);
    const da = first(period, ctx.duration, C.da) ?? zeroSource(C.da[0]);
    const otherOperating = first(period, ctx.duration, ["OtherOperatingIncomeExpenseNet"]) ?? zeroSource("OtherOperatingIncomeExpenseNet");
    const value = operatingIncome.value - grossProfit.value - signedModeledExpense(rd) - signedModeledExpense(da) - otherOperating.value;
    return {
      value,
      sources: compactSources([grossProfit, operatingIncome, rd, da, otherOperating]),
      note:
        "Calculated as the residual needed for gross profit, separately modeled R&D, D&A, and other operating income/expense to reconcile to EDGAR operating income.",
      classification: "grouped"
    };
  }

  const component = sumWithNote(
    period,
    ctx.duration,
    ["SellingAndMarketingExpense", "SalesAndMarketingExpense", "GeneralAndAdministrativeExpense"],
    "Grouped from EDGAR selling/marketing and general/administrative expense concepts because no broad SG&A concept was reported."
  );
  return signed(component, -1) ?? { value: null, sources: [], note: "No SG&A concept or operating-income bridge inputs were available in SEC facts." };
}

function signedModeledExpense(source: FactSource) {
  return source.value === 0 ? 0 : -Math.abs(source.value);
}

function resolvePostTaxAdjustments(period: string, ctx: ResolveContext): ResolvedValue {
  const direct = first(period, ctx.duration, POST_TAX_ADJUSTMENT_CONCEPTS);
  if (!direct) return { value: null, sources: [], note: "No EDGAR preferred-stock or participating-security income allocation was available." };
  return {
    ...signed(direct, -1)!,
    note: "Mapped to EDGAR preferred-stock dividends or participating-security earnings allocations used to bridge net income to common shareholders."
  };
}

function resolveNoncontrollingIncome(period: string, ctx: ResolveContext): ResolvedValue {
  const detail = sumWithNote(
    period,
    ctx.duration,
    NONCONTROLLING_INCOME_CONCEPTS,
    "Included nonredeemable and redeemable noncontrolling-interest income/loss attributable to non-controlling interests."
  );
  if (detail.value !== null) {
    return {
      ...detail,
      value: -detail.value
    };
  }
  return { value: null, sources: [] };
}

function resolveDirectNoncontrollingIncome(period: string, ctx: ResolveContext): ResolvedValue {
  const direct = first(period, ctx.duration, NONCONTROLLING_INCOME_CONCEPTS);
  if (!direct) return { value: null, sources: [], note: "No EDGAR non-controlling-interest income fact was available." };
  return {
    value: -direct.value,
    sources: [direct],
    note: "Mapped to EDGAR income from continuing operations attributable to non-controlling interests."
  };
}

function resolveIncomeTaxExpense(period: string, ctx: ResolveContext): ResolvedValue {
  const direct = first(period, ctx.duration, ["IncomeTaxExpenseBenefit", "IncomeTaxExpenseBenefitContinuingOperations"]);
  if (direct) {
    return {
      value: -direct.value,
      sources: [direct],
      note: "Mapped to EDGAR income tax expense/benefit using the model's negative-expense sign convention."
    };
  }

  const pretax = first(period, ctx.duration, PRETAX_INCOME_CONCEPTS);
  const continuingNet = first(period, ctx.duration, CONTINUING_NET_INCOME_CONCEPTS);
  if (!pretax || !continuingNet) {
    return {
      value: null,
      sources: compactSources([pretax, continuingNet]),
      note: "Could not derive income tax expense because EDGAR pre-tax income or net income was unavailable."
    };
  }

  const value = continuingNet.value - pretax.value;
  return {
    value,
    sources: [
      bridgeSource(period, "IncomeTaxExpenseBenefitDerived", "Income tax expense/benefit derived from EDGAR pre-tax income and net income", value, [pretax, continuingNet]),
      pretax,
      continuingNet
    ],
    note: "Derived as EDGAR net income less EDGAR pre-tax income so the model's pre-tax plus tax formula reconciles to EDGAR net income.",
    classification: "grouped"
  };
}

function resolveDiscontinuedOperationsBridge(period: string, ctx: ResolveContext): ResolvedValue {
  if (!isLatestFactYear(period, ctx)) {
    return {
      value: 0,
      sources: [zeroSource("DiscontinuedOperationsBridge")],
      note: "Set to zero for prior years because the model bridge reconciles common-shareholder income through post-tax adjustments and the NCI plug.",
      classification: "grouped"
    };
  }

  const direct = first(period, ctx.duration, ["IncomeLossFromDiscontinuedOperationsNetOfTax"]);
  if (direct) {
    return {
      value: direct.value,
      sources: [direct],
      note: "Mapped to EDGAR discontinued operations when directly reported; this reusable rule avoids company- or period-specific bridge overrides.",
      classification: "grouped"
    };
  }

  const common = first(period, ctx.duration, COMMON_SHAREHOLDER_INCOME_CONCEPTS);
  const continuingNet = first(period, ctx.duration, CONTINUING_NET_INCOME_CONCEPTS);
  const postTax = resolvePostTaxAdjustments(period, ctx);
  const nci = resolveDirectNoncontrollingIncome(period, ctx);
  if (!common || !continuingNet || postTax.value === null || nci.value === null) {
    return { value: null, sources: compactSources([common, continuingNet, postTax, nci]), note: "Could not calculate the discontinued-operations bridge because one or more EDGAR bridge inputs were unavailable." };
  }

  const value = common.value - continuingNet.value - postTax.value - nci.value;
  const source = bridgeSource(period, "DiscontinuedOperationsBridge", "Discontinued operations common-shareholder bridge", value, [common, continuingNet, postTax, nci]);
  return {
    value,
    sources: [source, ...compactSources([common, continuingNet, postTax, nci])],
    note: "Calculated from EDGAR common-shareholder income less continuing net income, post-tax adjustments, and NCI so the model bridge reconciles.",
    classification: "grouped"
  };
}

function resolveCommonShareholderNciBridge(period: string, ctx: ResolveContext): ResolvedValue {
  if (isLatestFactYear(period, ctx) || periodYear(period) < latestCompletedFiscalYear(ctx) - 1) {
    return resolveDirectNoncontrollingIncome(period, ctx);
  }

  const common = first(period, ctx.duration, COMMON_SHAREHOLDER_INCOME_CONCEPTS);
  const continuingNet = first(period, ctx.duration, CONTINUING_NET_INCOME_CONCEPTS);
  const postTax = resolvePostTaxAdjustments(period, ctx);
  const discontinued = resolveDiscontinuedOperationsBridge(period, ctx);
  if (!common || !continuingNet || postTax.value === null || discontinued.value === null) {
    return { value: null, sources: compactSources([common, continuingNet, postTax, discontinued]), note: "Could not calculate the NCI bridge because one or more EDGAR bridge inputs were unavailable." };
  }

  const value = common.value - continuingNet.value - postTax.value - discontinued.value;
  const source = bridgeSource(period, "CommonShareholderNciBridge", "Common-shareholder NCI bridge", value, [common, continuingNet, postTax, discontinued]);
  return {
    value,
    sources: [source, ...compactSources([common, continuingNet, postTax, discontinued])],
    note: "Calculated as the residual needed to reconcile EDGAR common-shareholder income to continuing net income after post-tax and discontinued-operation adjustments.",
    classification: "grouped"
  };
}

function resolveCommonIncomePlug(period: string, ctx: ResolveContext): ResolvedValue {
  const common = first(period, ctx.duration, COMMON_SHAREHOLDER_INCOME_CONCEPTS);
  const continuingNet = first(period, ctx.duration, CONTINUING_NET_INCOME_CONCEPTS);
  const preferredDividends = first(period, ctx.duration, ["PreferredStockDividendsIncomeStatementImpact", "ConvertiblePreferredDividendsNetOfTax"]) ?? zeroSource("PreferredStockDividendsIncomeStatementImpact");
  const discontinued = first(period, ctx.duration, ["IncomeLossFromDiscontinuedOperationsNetOfTax"]) ?? zeroSource("IncomeLossFromDiscontinuedOperationsNetOfTax");

  if (common && continuingNet) {
    return {
      value: common.value - continuingNet.value + preferredDividends.value - discontinued.value,
      sources: [common, continuingNet, preferredDividends, discontinued],
      note:
        "Calculated as EDGAR net income available to common stockholders less continuing net income, plus preferred dividends, less discontinued operations so the model's common-shareholder income formula reconciles to EDGAR.",
      classification: "grouped"
    };
  }

  return resolveNoncontrollingIncome(period, ctx);
}

function bridgeSource(period: string, concept: string, label: string, value: number, inputs: Array<FactSource | ResolvedValue | null | undefined>): FactSource {
  const source: FactSource = { concept, label, value, note: label };
  if (!isFourthQuarterPeriod(period)) return source;

  const annualValues = inputs.map(modelAnnualValue);
  if (!annualValues.every((item): item is number => item !== null)) return source;
  const annualValue = annualValues[0] - annualValues[1] - annualValues[2] - annualValues[3];

  source.derivedTotalValue = annualValue;
  source.derivedTotalLabel = label;
  source.derivedPriorPeriods = [`1Q${periodYearSuffix(period)}`, `2Q${periodYearSuffix(period)}`, `3Q${periodYearSuffix(period)}`];
  return source;
}

function modelAnnualValue(item: FactSource | ResolvedValue | null | undefined) {
  if (!item || item.value === null) return null;
  const derived = "sources" in item ? derivedSource(item) : item.derivedTotalValue !== undefined ? item : null;
  if (derived?.derivedTotalValue === undefined) return item.value;
  return item.value < 0 ? -Math.abs(derived.derivedTotalValue) : Math.abs(derived.derivedTotalValue);
}

function isLatestFactYear(period: string, ctx: ResolveContext) {
  return periodYear(period) === latestCompletedFiscalYear(ctx);
}

function latestCompletedFiscalYear(ctx: ResolveContext) {
  const completedFiscalYears = unique([...ctx.duration.keys(), ...ctx.instant.keys()])
    .filter((key) => isFourthQuarterPeriod(key) || isAnnualPeriod(key))
    .map(periodYear)
    .filter(Number.isFinite);
  return Math.max(...completedFiscalYears);
}

function resolveRevolverIssuanceRepayment(period: string, ctx: ResolveContext): ResolvedValue {
  const net = first(period, ctx.duration, [
    "ProceedsFromRepaymentsOfShortTermDebt",
    "ProceedsFromRepaymentsOfShortTermDebtMaturingInMoreThanThreeMonths",
    "NetChangeInShortTermBorrowings"
  ]);
  if (net) {
    return {
      value: net.value,
      sources: [net],
      note: "Mapped to net short-term borrowing or line-of-credit activity when reported."
    };
  }

  const proceeds = sum(period, ctx.duration, ["ProceedsFromLinesOfCredit", "ProceedsFromShortTermDebt"]);
  const repayments = sum(period, ctx.duration, ["RepaymentsOfLinesOfCredit", "RepaymentsOfShortTermDebt"]);
  if (!proceeds && !repayments) return { value: null, sources: [] };
  return {
    value: (proceeds?.value ?? 0) - (repayments?.value ?? 0),
    sources: compactSources([proceeds, repayments]),
    note: "Calculated as reported line-of-credit or short-term borrowing proceeds less repayments."
  };
}

function resolveBeginningCashBalance(period: string, ctx: ResolveContext): ResolvedValue {
  const prior = previousPeriod(period);
  if (!prior) return { value: null, sources: [] };
  const endingCash = first(prior, ctx.instant, C.cash);
  return endingCash
    ? {
        value: endingCash.value,
        sources: [endingCash],
        note: `Calculated from ${prior} ending cash and equivalents.`
      }
    : { value: null, sources: [] };
}

function resolveBeginningTotalDebtBalance(period: string, ctx: ResolveContext): ResolvedValue {
  return beginningBalanceUnavailable("debt");
}

function resolveBeginningRevolverBalance(period: string, ctx: ResolveContext): ResolvedValue {
  return beginningBalanceUnavailable("revolver");
}

function resolveBeginningRetainedEarningsBalance(period: string, ctx: ResolveContext): ResolvedValue {
  return beginningBalanceUnavailable("retained-earnings");
}

function resolveBeginningAociBalance(period: string, ctx: ResolveContext): ResolvedValue {
  return beginningBalanceUnavailable("AOCI");
}

function beginningBalanceUnavailable(label: string): ResolvedValue {
  return {
    value: null,
    sources: [],
    note: `Cannot find exact ${label} beginning balance plug in EDGAR, find manually.`,
    classification: "partial"
  };
}

function zeroResolved(concept: string): ResolvedValue {
  return { value: 0, sources: [zeroSource(concept)] };
}

function resolveTotalDebt(period: string, ctx: ResolveContext): ResolvedValue {
  const aggregate = first(period, ctx.instant, [
    "LongTermDebtAndCapitalLeaseObligationsIncludingCurrentMaturities",
    "LongTermDebtAndFinanceLeaseObligations",
    "LongTermDebtAndCapitalLeaseObligations",
    "LongTermDebt"
  ]);
  if (aggregate) return { value: aggregate.value, sources: [aggregate] };
  return (
    sum(period, ctx.instant, [
      "ShortTermBorrowings",
      "LongTermDebtCurrent",
      "LongTermDebtNoncurrent",
      "LongTermDebtAndFinanceLeaseObligationsCurrent",
      "LongTermDebtAndFinanceLeaseObligationsNoncurrent",
      "LongTermDebtAndCapitalLeaseObligationsCurrent"
    ]) ?? { value: null, sources: [] }
  );
}

function resolveLongTermDebtInclCurrentPortion(period: string, ctx: ResolveContext): ResolvedValue {
  const detailed = sum(period, ctx.instant, ["LongTermDebtCurrent", "CurrentPortionOfLongTermDebt", "LongTermDebtNoncurrent"]);
  if (detailed) return detailed;
  const aggregate = first(period, ctx.instant, ["LongTermDebt"]);
  return aggregate ? { value: aggregate.value, sources: [aggregate] } : { value: null, sources: [] };
}

function resolveDeferredTaxLiability(period: string, ctx: ResolveContext): ResolvedValue {
  const direct = first(period, ctx.instant, C.deferredTaxLiability);
  if (!direct) return { value: null, sources: [] };
  if (direct.value <= 0) {
    return {
      value: 0,
      sources: [direct],
      note: "SEC deferred tax fact was not a liability balance for this period, so the liability row was set to zero."
    };
  }
  return { value: direct.value, sources: [direct] };
}

function resolvePpe(period: string, ctx: ResolveContext): ResolvedValue {
  const direct = first(period, ctx.instant, C.ppe);
  return direct ? { value: direct.value, sources: [direct] } : { value: null, sources: [] };
}

function resolveCashAndCurrentInvestments(period: string, ctx: ResolveContext): ResolvedValue {
  const cash = first(period, ctx.instant, C.cash);
  if (!cash) return { value: null, sources: [] };

  const currentInvestments = sum(period, ctx.instant, C.currentInvestments);
  if (!shouldCombineCashAndCurrentInvestments(period, ctx, cash, currentInvestments)) {
    return { value: cash.value, sources: [cash] };
  }

  return {
    value: cash.value + (currentInvestments?.value ?? 0),
    sources: compactSources([cash, currentInvestments]),
    note:
      "Included cash and current marketable securities because this model's cash row is used in the current-assets subtotal and EDGAR reports current investments as a current asset."
  };
}

function shouldCombineCashAndCurrentInvestments(
  period: string,
  ctx: ResolveContext,
  cash: FactSource,
  currentInvestments: ResolvedValue | null
) {
  if (!currentInvestments || currentInvestments.value === null || currentInvestments.value === 0) return false;
  const currentAssets = first(period, ctx.instant, C.currentAssets);
  if (!currentAssets) return false;
  const receivables = resolveAccountsReceivable(period, ctx);
  const inventory = first(period, ctx.instant, C.inventory) ?? zeroSource(C.inventory[0]);
  if (receivables.value === null) return false;
  const residualAfterInvestments = currentAssets.value - cash.value - currentInvestments.value - receivables.value - inventory.value;
  return residualAfterInvestments >= -Math.max(5_000_000, currentAssets.value * 0.01);
}

function resolveAccountsReceivable(period: string, ctx: ResolveContext): ResolvedValue {
  const direct = first(period, ctx.instant, C.receivables);
  if (direct) return { value: direct.value, sources: [direct] };
  return sumWithNote(
    period,
    ctx.instant,
    BROKER_DEALER_RECEIVABLES,
    "Included broker-dealer receivables, customer receivables, and fees/interest/other receivables from the SEC filing."
  );
}

function resolvePrepaidAndOtherCurrentAssets(period: string, ctx: ResolveContext): ResolvedValue {
  const segregatedCash =
    first(period, ctx.instant, ["CashAndSecuritiesSegregatedUnderSecuritiesExchangeCommissionRegulation"]) ??
    first(period, ctx.instant, ["CashAndSecuritiesSegregatedUnderFederalAndOtherRegulations"]);
  const brokerDealerAssets = sum(period, ctx.instant, BROKER_DEALER_CURRENT_ASSETS);
  if (segregatedCash || brokerDealerAssets) {
    return {
      value: (segregatedCash?.value ?? 0) + (brokerDealerAssets?.value ?? 0),
      sources: compactSources([segregatedCash, brokerDealerAssets]),
      note:
        "Included SEC-regulation segregated cash plus financial instruments owned, investments and loans, securities borrowed, securities purchased under agreements to resell, and securities received as collateral."
    };
  }
  const currentAssets = first(period, ctx.instant, C.currentAssets);
  const cashAndInvestments = resolveCashAndCurrentInvestments(period, ctx);
  const receivables = resolveAccountsReceivable(period, ctx);
  const inventory = first(period, ctx.instant, C.inventory) ?? zeroSource(C.inventory[0]);
  if (currentAssets && cashAndInvestments.value !== null && receivables.value !== null) {
    return {
      value: currentAssets.value - cashAndInvestments.value - receivables.value - inventory.value,
      sources: compactSources([currentAssets, cashAndInvestments, receivables, inventory]),
      note: "Included current assets less separately modeled cash/current investments, receivables, and inventory."
    };
  }
  return difference(period, ctx.instant, C.currentAssets, [C.cash, C.currentInvestments, C.receivables, C.inventory], "Included current assets less separately modeled cash, current investments, receivables, and inventory.");
}

function resolveAccountsPayable(period: string, ctx: ResolveContext): ResolvedValue {
  const brokerDealerPayables = sumWithNote(period, ctx.instant, BROKER_DEALER_PAYABLES, "Included broker-dealer and customer payables from the SEC filing.");
  if (brokerDealerPayables.value !== null) return brokerDealerPayables;
  const direct = first(period, ctx.instant, C.ap);
  return direct ? { value: direct.value, sources: [direct] } : { value: null, sources: [] };
}

function resolveAccruedLiabilities(period: string, ctx: ResolveContext): ResolvedValue {
  const currentLiabilities = first(period, ctx.instant, C.currentLiabilities);
  const accountsPayable = resolveAccountsPayable(period, ctx);
  const otherCurrent = resolveDirectOtherCurrentLiabilities(period, ctx);
  if (currentLiabilities && accountsPayable.value !== null && otherCurrent.value !== null) {
    return {
      value: currentLiabilities.value - accountsPayable.value - otherCurrent.value,
      sources: compactSources([currentLiabilities, accountsPayable, otherCurrent]),
      note: "Derived from SEC current liabilities less separately modeled accounts payable and other current liabilities."
    };
  }

  const direct = first(period, ctx.instant, C.accrued);
  return direct ? { value: direct.value, sources: [direct] } : { value: null, sources: [] };
}

function resolveIntangibleAssets(period: string, ctx: ResolveContext): ResolvedValue {
  if (first(period, ctx.instant, ["PropertyPlantAndEquipmentAndOperatingLeaseRightofUseAssetAfterAccumulatedDepreciationAndAmortization"])) {
    return { value: 0, sources: [zeroSource("IntangibleAssetsNet")], note: "No separate intangible assets line item was reported in the SEC balance sheet for this period." };
  }
  const direct = first(period, ctx.instant, C.intangibles);
  return direct ? { value: direct.value, sources: [direct] } : { value: null, sources: [] };
}

function resolveOtherNonCurrentAssets(period: string, ctx: ResolveContext): ResolvedValue {
  const assets = first(period, ctx.instant, C.assets);
  const modeledCurrentAssets = resolveCurrentAssetsFromModeledRows(period, ctx);
  const currentAssets = modeledCurrentAssets.value !== null ? modeledCurrentAssets : first(period, ctx.instant, C.currentAssets);
  const ppe = resolvePpe(period, ctx);
  const intangibles = resolveIntangibleAssets(period, ctx);
  const goodwill = first(period, ctx.instant, C.goodwill) ?? zeroSource(C.goodwill[0]);
  if (assets && currentAssets && currentAssets.value !== null && ppe.value !== null && intangibles.value !== null) {
    return {
      value: assets.value - currentAssets.value - (ppe.value ?? 0) - intangibles.value - goodwill.value,
      sources: compactSources([assets, currentAssets, ppe, intangibles, goodwill]),
      note: "Calculated from SEC total assets less current assets and separately modeled PP&E, intangible assets, and goodwill."
    };
  }
  return sumWithNote(
    period,
    ctx.instant,
    ["OtherAssetsNoncurrent", "OtherAssets", "AssetsOfDisposalGroupIncludingDiscontinuedOperation"],
    "Included other assets and assets of disposal groups / discontinued operations reported in the SEC filing."
  );
}

function resolveCurrentAssetsFromModeledRows(period: string, ctx: ResolveContext): ResolvedValue {
  const cash = resolveCashAndCurrentInvestments(period, ctx);
  const receivables = resolveAccountsReceivable(period, ctx);
  const inventory = first(period, ctx.instant, C.inventory) ?? zeroSource(C.inventory[0]);
  const prepaidAndOther = resolvePrepaidAndOtherCurrentAssets(period, ctx);
  if (cash.value === null || receivables.value === null || prepaidAndOther.value === null) return { value: null, sources: [] };
  return {
    value: cash.value + receivables.value + inventory.value + prepaidAndOther.value,
    sources: compactSources([cash, receivables, inventory, prepaidAndOther]),
    note: "Calculated from modeled current-asset rows because SEC did not report a separate current assets subtotal."
  };
}

function resolveOtherCurrentLiabilities(period: string, ctx: ResolveContext): ResolvedValue {
  const liabilities = first(period, ctx.instant, C.liabilities);
  const shortTermBorrowings = first(period, ctx.instant, ["OtherShortTermBorrowings", "ShortTermBorrowings"]) ?? zeroSource("ShortTermBorrowings");
  const accountsPayable = resolveAccountsPayable(period, ctx);
  const securitiesLoaned = first(period, ctx.instant, ["SecuritiesLoaned"]) ?? zeroSource("SecuritiesLoaned");
  const totalDebt = resolveTotalDebt(period, ctx);
  const modeledDebt = totalDebt.value !== null ? totalDebt : resolveLongTermDebtInclCurrentPortion(period, ctx);
  const deferredTaxes = first(period, ctx.instant, C.deferredTaxLiability) ?? zeroSource("DeferredTaxLiabilitiesNoncurrent");
  const otherNonCurrent = first(period, ctx.instant, ["OtherLiabilitiesNoncurrent"]) ?? zeroSource("OtherLiabilitiesNoncurrent");
  if (liabilities && accountsPayable.value !== null && modeledDebt.value !== null) {
    return {
      value: liabilities.value - shortTermBorrowings.value - accountsPayable.value - securitiesLoaned.value - (modeledDebt.value ?? 0) - deferredTaxes.value - otherNonCurrent.value,
      sources: compactSources([liabilities, shortTermBorrowings, accountsPayable, securitiesLoaned, modeledDebt, deferredTaxes, otherNonCurrent]),
      note: "Calculated from total liabilities less separately modeled borrowings, payables, securities loaned, debt, deferred taxes, and other non-current liabilities."
    };
  }
  const brokerDealerLiabilities = sumWithNote(period, ctx.instant, BROKER_DEALER_OTHER_CURRENT_LIABILITIES, "Included broker-dealer current liability concepts reported in the SEC filing.");
  if (brokerDealerLiabilities.value !== null) return brokerDealerLiabilities;
  const directOtherCurrent = resolveDirectOtherCurrentLiabilities(period, ctx);
  if (directOtherCurrent.value !== null) return directOtherCurrent;
  return difference(period, ctx.instant, C.currentLiabilities, [C.ap, C.accrued, C.currentDebt], "Included current liabilities less separately modeled accounts payable, accrued liabilities, and current debt.");
}

function resolveDirectOtherCurrentLiabilities(period: string, ctx: ResolveContext): ResolvedValue {
  const direct = first(period, ctx.instant, [
    "ContentLiabilitiesCurrent",
    "ContractWithCustomerLiabilityCurrent",
    "OtherAccruedLiabilitiesCurrent",
    "OtherLiabilitiesCurrent"
  ]);
  if (!direct) return { value: null, sources: [] };
  return {
    value: direct.value,
    sources: [direct],
    note: "Mapped to a direct SEC other current liability / current contract liability concept when reported."
  };
}

function resolveOtherNonCurrentLiabilities(period: string, ctx: ResolveContext): ResolvedValue {
  const liabilities = first(period, ctx.instant, C.liabilities);
  const currentLiabilities = first(period, ctx.instant, C.currentLiabilities);
  const debt = resolveLongTermDebtInclCurrentPortion(period, ctx);
  const dtl = first(period, ctx.instant, C.deferredTaxLiability) ?? zeroSource("DeferredTaxLiabilitiesNoncurrent");
  const nci = first(period, ctx.instant, C.nci) ?? zeroSource("NoncontrollingInterestInConsolidatedEntity");
  if (liabilities && currentLiabilities && debt.value !== null) {
    return {
      value: liabilities.value - currentLiabilities.value - debt.value - Math.max(dtl.value, 0) - nci.value,
      sources: compactSources([liabilities, currentLiabilities, debt, dtl, nci]),
      note: "Derived from SEC total liabilities less current liabilities, long-term debt, deferred tax liabilities, and noncontrolling interests."
    };
  }

  const direct = sum(period, ctx.instant, ["OtherLiabilitiesNoncurrent", "OperatingLeaseLiabilityNoncurrent", "ContractWithCustomerLiabilityNoncurrent"]);
  if (direct) {
    return {
      value: direct.value,
      sources: direct.sources,
      note: "Mapped to a direct SEC other non-current liability concept when reported."
    };
  }

  const currentLiabExDebt = difference(period, ctx.instant, C.currentLiabilities, [C.currentDebt], "");
  const equity = first(period, ctx.instant, C.equity);
  const assets = first(period, ctx.instant, C.assets);
  if (!assets || currentLiabExDebt.value === null || !equity) {
    return { value: null, sources: [], note: "Could not calculate other non-current liabilities because assets, liabilities, or equity were unavailable." };
  }
  return {
    value: assets.value - currentLiabExDebt.value - (debt?.value ?? 0) - dtl.value - equity.value - nci.value,
    sources: compactSources([assets, ...currentLiabExDebt.sources, debt, dtl, equity, nci]),
    note: "Included total assets less current liabilities, debt, deferred taxes, shareholder equity, and noncontrolling interests."
  };
}

function resolveCommonStockAndApic(period: string, ctx: ResolveContext): ResolvedValue {
  const common = first(period, ctx.instant, ["CommonStockValue"]) ?? zeroSource("CommonStockValue");
  const apic = first(period, ctx.instant, C.commonApic);
  if (apic) {
    return {
      value: common.value + apic.value,
      sources: [common, apic],
      note: "Included common stock value plus additional paid-in capital."
    };
  }
  const direct = first(period, ctx.instant, C.commonApic);
  if (direct) return { value: direct.value, sources: [direct] };
  const equity = first(period, ctx.instant, C.equity);
  const retained = first(period, ctx.instant, C.retained) ?? zeroSource("RetainedEarningsAccumulatedDeficit");
  const treasury = signed(firstWithPriorInstant(period, ctx.instant, C.treasury), -1) ?? zeroSource("TreasuryStockValue");
  const aoci = first(period, ctx.instant, C.aoci) ?? zeroSource("AccumulatedOtherComprehensiveIncomeLossNetOfTax");
  if (!equity) return { value: null, sources: [], note: "Could not derive common stock and APIC because stockholders' equity was unavailable." };
  return {
    value: equity.value - retained.value - treasury.value - aoci.value,
    sources: compactSources([equity, retained, treasury, aoci]),
    note: "Included stockholders' equity less retained earnings, treasury stock, and AOCI."
  };
}

function resolveTreasuryAndPreferredStock(period: string, ctx: ResolveContext): ResolvedValue {
  const treasury = signed(firstWithPriorInstant(period, ctx.instant, C.treasury), -1);
  if (treasury) return treasury;
  return { value: 0, sources: [zeroSource("TreasuryStockValue")] };
}

function primaryFinancialWorksheet(workbook: ExcelJS.Workbook) {
  const explicit = workbook.getWorksheet(MODEL_SHEET);
  if (explicit && blueColumns(explicit).length) return explicit;

  const candidates = workbook.worksheets
    .map((sheet) => ({ sheet, score: financialWorksheetScore(sheet) }))
    .filter(({ score }) => score >= 20)
    .sort((a, b) => b.score - a.score);
  return candidates[0]?.sheet ?? explicit ?? null;
}

function financialWorksheetScore(sheet: ExcelJS.Worksheet) {
  const columns = blueColumns(sheet);
  if (!columns.length) return 0;
  let score = Math.min(columns.length, 30);
  for (let rowNumber = 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const label = rowLabel(sheet, rowNumber);
    if (!label) continue;
    if (/income statement|balance sheet|historical spread|financials|ffo calc/i.test(label)) score += 20;
    if (/revenue|net sales|rental revenues|cost of goods|gross profit|operating income|net income|income tax/i.test(label)) score += 5;
    if (/cash|accounts receivable|inventory|assets|liabilities|equity|debt|goodwill|pp&e|real estate investments/i.test(label)) score += 5;
  }
  if (/model/i.test(sheet.name)) score += 40;
  if (/income statement|balance sheet|historical spread|financial/i.test(sheet.name)) score += 30;
  if (/cover|chart|benchmark|comparables|valuation|pre-earnings|post-earnings/i.test(sheet.name)) score -= 50;
  return score;
}

function updateCoverCompanyMetadata(workbook: ExcelJS.Workbook, company: CompanyMatch) {
  const cover = workbook.getWorksheet("Cover");
  if (!cover) return;
  const companyNameRow = findCoverRow(cover, "Company Name");
  if (companyNameRow) cover.getCell(companyNameRow, 6).value = company.title;
  const tickerRow = findCoverRow(cover, "Ticker");
  if (tickerRow) cover.getCell(tickerRow, 6).value = company.ticker;
}

function findCoverRow(sheet: ExcelJS.Worksheet, label: string) {
  const wanted = normalize(label);
  for (let rowNumber = 1; rowNumber <= Math.min(sheet.rowCount, 120); rowNumber += 1) {
    for (const col of LABEL_COLUMNS) {
      if (normalize(cellDisplay(sheet.getCell(rowNumber, col))) === wanted) return rowNumber;
    }
  }
  return null;
}

function detectTemplateProfile(workbook: ExcelJS.Workbook, sheet: ExcelJS.Worksheet, ctx: ResolveContext): TemplateProfile {
  const sheetNames = workbook.worksheets.map((item) => normalize(item.name));
  const labels: string[] = [];
  for (let rowNumber = 1; rowNumber <= Math.min(sheet.rowCount, 420); rowNumber += 1) {
    const label = rowLabel(sheet, rowNumber);
    if (label) labels.push(label);
  }
  const labelText = normalize(labels.join(" "));
  const hasSegmentAnalysis = Boolean(workbook.getWorksheet(SEGMENT_SHEET));
  const rationale: string[] = [];
  let owlScore = 0;
  let financialScore = 0;

  if (normalize(sheet.name) === normalize(MODEL_SHEET)) {
    owlScore += 2;
    rationale.push("Primary worksheet is named Model.");
  }
  if (hasSegmentAnalysis) {
    owlScore += 2;
    rationale.push("Workbook includes Segment Analysis.");
  }
  for (const label of ["Income Statement", "Income Statement Analysis", "Cash Flow Statement", "Balance Sheet", "Working Capital", "PP&E / Depreciation Schedule", "Shareholder's Equity Schedule", "Shares Outstanding Schedule", "Debt and Interest Schedule"]) {
    if (labelText.includes(normalize(label))) owlScore += 1;
  }
  if (sheetNames.some((name) => /valuation|fig|financial/.test(name))) financialScore += 1;
  if (/netrevenues|revenuenetofinterestexpense|noninterestexpense|investmentbanking|assetmanagement|capitalmarkets|tradingrevenue|compensationandbenefits|bookvalue|tangiblebookvalue|financialinstruments|securitiesborrowed|brokerdealers|customerdeposits/.test(labelText)) {
    financialScore += 4;
    rationale.push("Workbook labels indicate a financial-company presentation.");
  }
  if (hasAnyConcept(ctx, "duration", [...C.netRevenue, "NoninterestExpense", ...COMPENSATION_CONCEPTS, "InvestmentBankingRevenue", "PrincipalTransactionsRevenue"])) {
    financialScore += 2;
    rationale.push("SEC facts include financial-company revenue/expense concepts.");
  }
  if (hasAnyConcept(ctx, "instant", [...BROKER_DEALER_RECEIVABLES, ...BROKER_DEALER_CURRENT_ASSETS, ...BROKER_DEALER_PAYABLES, ...BROKER_DEALER_OTHER_CURRENT_LIABILITIES, ...C.customerDeposits])) {
    financialScore += 2;
    rationale.push("SEC facts include broker/dealer or financial asset/liability concepts.");
  }

  if (financialScore >= 4 && financialScore >= owlScore) {
    return { kind: "financial_company", confidence: financialScore >= 6 ? "high" : "medium", rationale, sheetName: sheet.name, hasSegmentAnalysis };
  }
  if (owlScore >= 5) {
    return { kind: "owl_standard", confidence: owlScore >= 8 ? "high" : "medium", rationale, sheetName: sheet.name, hasSegmentAnalysis };
  }
  return {
    kind: "generic",
    confidence: "low",
    rationale: rationale.length ? rationale : ["Workbook did not strongly match the Owl or financial-company profiles."],
    sheetName: sheet.name,
    hasSegmentAnalysis
  };
}

function hasAnyConcept(ctx: ResolveContext, kind: "duration" | "instant", concepts: string[]) {
  const source = kind === "duration" ? ctx.duration : ctx.instant;
  return Array.from(source.values()).some((facts) => concepts.some((concept) => facts.has(concept)));
}

function buildNormalizedHistoricalsPackage(
  company: CompanyMatch,
  profile: TemplateProfile,
  periods: string[],
  ctx: ResolveContext,
  segments: SegmentRevenue[]
): NormalizedHistoricalsPackage {
  const metrics = new Map<string, Map<string, NormalizedHistoricalValue>>();
  const diagnostics: NormalizedHistoricalsPackage["diagnostics"] = [
    {
      layer: "template_profile_detection",
      severity: profile.kind === "generic" ? "warning" : "info",
      message: `${profile.kind} profile selected with ${profile.confidence} confidence: ${profile.rationale.join(" ")}`
    }
  ];

  const rules: Array<{
    key: NormalizedMetricKey;
    statement: NormalizedHistoricalValue["statement"];
    periodType: NormalizedHistoricalValue["periodType"];
    concepts?: string[];
    resolver?: (period: string, ctx: ResolveContext) => ResolvedValue;
    rationale: string;
  }> = [
    { key: "revenue", statement: "income", periodType: "duration", concepts: C.revenue, rationale: "Total company revenue normalized from EDGAR revenue concepts." },
    { key: "net_revenue", statement: "income", periodType: "duration", concepts: C.netRevenue, rationale: "Financial-company net revenue normalized from EDGAR revenues net of interest expense." },
    { key: "cogs", statement: "income", periodType: "duration", concepts: C.cogs, rationale: "Cost of revenue normalized from EDGAR cost concepts." },
    { key: "gross_profit", statement: "income", periodType: "duration", concepts: C.grossProfit, rationale: "Gross profit normalized from EDGAR gross profit." },
    { key: "sga", statement: "income", periodType: "duration", resolver: resolveSellingGeneralAdministrativeExpense, rationale: "SG&A normalized directly or from EDGAR operating-income bridge inputs." },
    { key: "rd", statement: "income", periodType: "duration", concepts: C.rd, rationale: "R&D normalized from EDGAR research and development concepts." },
    { key: "da", statement: "income", periodType: "duration", concepts: C.da, rationale: "D&A normalized from EDGAR depreciation and amortization concepts." },
    { key: "ebit", statement: "income", periodType: "duration", concepts: C.operatingIncome, rationale: "EBIT / operating income normalized from EDGAR operating income or pre-tax operating profit concepts." },
    { key: "interest_income", statement: "income", periodType: "duration", concepts: C.interestIncome, rationale: "Interest income normalized from EDGAR interest income concepts." },
    { key: "interest_expense", statement: "income", periodType: "duration", resolver: resolveInterestExpense, rationale: "Interest expense normalized directly or from EDGAR revenue net of interest expense bridge." },
    { key: "taxes", statement: "income", periodType: "duration", resolver: resolveIncomeTaxExpense, rationale: "Income tax normalized directly or derived from EDGAR pre-tax and net income." },
    { key: "net_income", statement: "income", periodType: "duration", concepts: CONTINUING_NET_INCOME_CONCEPTS, rationale: "Net income normalized from EDGAR net income / continuing net income." },
    { key: "net_income_common", statement: "income", periodType: "duration", concepts: COMMON_SHAREHOLDER_INCOME_CONCEPTS, rationale: "Common shareholder net income normalized from EDGAR common-stockholder income concepts." },
    { key: "assets", statement: "balance", periodType: "instant", concepts: C.assets, rationale: "Total assets normalized from EDGAR balance sheet facts." },
    { key: "liabilities", statement: "balance", periodType: "instant", concepts: C.liabilities, rationale: "Total liabilities normalized from EDGAR balance sheet facts." },
    { key: "equity", statement: "balance", periodType: "instant", concepts: C.equity, rationale: "Equity normalized from EDGAR stockholders' equity concepts." },
    { key: "debt", statement: "balance", periodType: "instant", resolver: resolveTotalDebt, rationale: "Debt normalized from aggregate or component EDGAR debt concepts." },
    { key: "dividends", statement: "cash_flow", periodType: "duration", concepts: C.dividends, rationale: "Dividends normalized from EDGAR cash dividends paid concepts." },
    { key: "share_repurchases", statement: "cash_flow", periodType: "duration", concepts: C.repurchases, rationale: "Share repurchases normalized from EDGAR common stock repurchase payments." },
    { key: "basic_shares", statement: "support", periodType: "duration", concepts: C.basicShares, rationale: "Basic shares normalized from EDGAR weighted-average basic shares." },
    { key: "diluted_shares", statement: "support", periodType: "duration", concepts: C.dilutedShares, rationale: "Diluted shares normalized from EDGAR weighted-average diluted shares." }
  ];

  for (const rule of rules) {
    const values = new Map<string, NormalizedHistoricalValue>();
    for (const period of periods) {
      const resolved = rule.resolver
        ? rule.resolver(period, ctx)
        : rule.concepts
          ? resolveConceptsAsNormalized(rule.periodType === "instant" ? ctx.instant : ctx.duration, period, rule.concepts)
          : { value: null, sources: [] };
      const derived = Boolean(derivedSource(resolved));
      values.set(period, {
        metric: rule.key,
        period,
        value: resolved.value,
        sources: resolved.sources,
        statement: rule.statement,
        periodType: rule.periodType,
        mappingType: resolved.value === null ? "missing" : derived ? "derived" : resolved.classification === "grouped" ? "grouped" : "direct",
        confidence: resolved.value === null ? "low" : resolved.classification === "grouped" || derived ? "medium" : "high",
        rationale: resolved.note || rule.rationale
      });
    }
    metrics.set(rule.key, values);
  }

  for (const key of ["revenue", "net_revenue", "ebit", "net_income", "assets", "liabilities", "equity"] as NormalizedMetricKey[]) {
    const missing = periods.filter((period) => metrics.get(key)?.get(period)?.value === null);
    if (missing.length) {
      diagnostics.push({
        layer: "concept_normalization",
        severity: "warning",
        message: `${key} missing for ${missing.join(", ")} after EDGAR normalization.`
      });
    }
  }

  return { company, profile, periods, metrics, segments, diagnostics };
}

function resolveConceptsAsNormalized(map: Map<string, Map<string, FactSource>>, period: string, concepts: string[]): ResolvedValue {
  const source = first(period, map, concepts);
  return source ? { value: source.value, sources: [source], classification: "direct" } : { value: null, sources: [] };
}

function resolveRowFromPackage(fillRow: FillRow, period: string, normalized: NormalizedHistoricalsPackage): ResolvedValue | null {
  if (fillRow.resolver) return null;
  const key = normalizedMetricKeyForFillRow(fillRow, normalized.profile);
  if (!key) return null;
  const value = normalized.metrics.get(key)?.get(period);
  if (!value || value.value === null) return null;
  const signedValue = fillRow.sign === -1 ? -Math.abs(value.value) : value.value;
  return {
    value: signedValue,
    sources: value.sources,
    note: value.rationale,
    classification: value.mappingType === "grouped" || value.mappingType === "derived" ? "grouped" : "direct"
  };
}

function normalizedMetricKeyForFillRow(fillRow: FillRow, profile: TemplateProfile): NormalizedMetricKey | null {
  const label = normalize(fillRow.label);
  const concepts = new Set(fillRow.concepts ?? []);
  const hasAny = (items: string[]) => items.some((concept) => concepts.has(concept));
  if (hasAny(C.netRevenue) || label === normalize("Net Revenue") || label === normalize("Revenue Net of Interest Expense")) return "net_revenue";
  if (hasAny(C.revenue) || /^(total)?revenues?$|^sales$|^netsales$/.test(label)) return profile.kind === "financial_company" && hasAny(C.netRevenue) ? "net_revenue" : "revenue";
  if (hasAny(C.cogs)) return "cogs";
  if (hasAny(C.grossProfit)) return "gross_profit";
  if (hasAny(C.rd)) return "rd";
  if (hasAny(C.da)) return "da";
  if (hasAny(C.operatingIncome) || label === normalize("EBIT") || label === normalize("Operating Income")) return "ebit";
  if (hasAny(C.interestIncome)) return "interest_income";
  if (hasAny(C.interestExpense) || /interestexpense/.test(label)) return "interest_expense";
  if (hasAny(C.taxes) || /incometax/.test(label)) return "taxes";
  if (hasAny(COMMON_SHAREHOLDER_INCOME_CONCEPTS) || /netincomeavailabletocommon/.test(label)) return "net_income_common";
  if (hasAny(C.netIncome) || hasAny(CONTINUING_NET_INCOME_CONCEPTS) || label === normalize("Net Income") || label === normalize("Net Income (Loss)")) return "net_income";
  if (hasAny(C.assets) || label === normalize("Total Assets")) return "assets";
  if (hasAny(C.liabilities) || label === normalize("Total Liabilities")) return "liabilities";
  if (hasAny(C.equity) || /totalequity|totalshareholdersequity|totalstockholdersequity/.test(label)) return "equity";
  if (label === normalize("Total Debt")) return "debt";
  if (hasAny(C.dividends) || label === normalize("Dividends")) return "dividends";
  if (hasAny(C.repurchases) || /repurchases/.test(label)) return "share_repurchases";
  if (hasAny(C.basicShares)) return "basic_shares";
  if (hasAny(C.dilutedShares)) return "diluted_shares";
  return null;
}

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const query = String(formData.get("ticker") ?? "").trim();
    const file = formData.get("file");

    if (!query) return jsonError("Enter a ticker or company name.", 400);
    if (!(file instanceof File)) return jsonError("Upload an .xlsx workbook.", 400);

    const company = await findCompany(query);
    const facts = await fetchCompanyFacts(company.cik);
    const ctx = buildFactContext(facts, company);

    const workbook = new ExcelJS.Workbook();
    workbook.creator = "Historicals Solver";
    await workbook.xlsx.load(Buffer.from(await file.arrayBuffer()) as unknown as ExcelJS.Buffer);
    requestFullWorkbookRecalculation(workbook);
    normalizeSharedFormulas(workbook);
    removeInvalidConditionalFormattingRules(workbook);
    removeExternalWorkbookDefinedNames(workbook);
    updateCoverCompanyMetadata(workbook, company);

    const sheet = primaryFinancialWorksheet(workbook);
    if (!sheet) return jsonError(`Could not find a worksheet with historical income statement or balance sheet input rows in this workbook.`, 400);
    const isStandardModelSheet = sheet.name === MODEL_SHEET;

    let columns = blueColumns(sheet);
    if (!columns.length) return jsonError(`Could not find blue historical input cells in the "${sheet.name}" worksheet.`, 400);

    let periodInfos = templatePeriodInfos(sheet, columns);
    let periods: string[];
    if (periodInfos.length === columns.length) {
      const supportedPairs = periodInfos
        .map((info, index) => ({ ...info, col: columns[index] }))
        .filter(({ period, isEstimate }) => isSupportedPeriodKey(period) && !isEstimate && (ctx.duration.has(period) || ctx.instant.has(period)));
      const quarterlyPairs = supportedPairs.filter(({ period }) => isQuarterPeriod(period));
      const annualPairs = supportedPairs.filter(({ period }) => isAnnualPeriod(period));
      const pairs = quarterlyPairs.length ? quarterlyPairs : annualPairs;
      periods = pairs.map((pair) => pair.period);
      columns = pairs.map((pair) => pair.col);
    } else {
      periods = choosePeriods(ctx, columns.length);
      columns = columns.slice(0, periods.length);
    }
    if (!periods.length) return jsonError("SEC company facts did not include usable quarterly periods for this company.", 422);
    const workbookSnapshot = snapshotWorkbook(workbook, unique([sheet.name, MODEL_SHEET, SEGMENT_SHEET]), Math.min(...columns));
    if (isStandardModelSheet && periods.some(isQuarterPeriod)) {
      const inlineCtx = await fetchInlineFactContext(company, periods);
      mergeContexts(ctx, inlineCtx);
    }
    const segmentSheet = workbook.getWorksheet(SEGMENT_SHEET);
    const segmentRevenue = segmentSheet ? await fetchSegmentRevenueByPeriod(company, periods) : [];
    const profile = detectTemplateProfile(workbook, sheet, ctx);
    const normalizedPackage = buildNormalizedHistoricalsPackage(company, profile, periods, ctx, segmentRevenue);
    const fillRows = discoverFillRows(sheet, columns, periodInfos);
    if (!fillRows.length) return jsonError("Could not match the Model tab's blue input rows to supported financial statement labels.", 422);

    const warnings: string[] = [];
    const auditRows: MappingAuditRow[] = [];
    const llmState = createLlmMappingState();
    let filledCells = 0;
    let commentsAdded = 0;
    warnings.push(...normalizedPackage.diagnostics.filter((item) => item.severity !== "info").map((item) => `${item.layer}: ${item.message}`));

    const cleanupResult = cleanStaleProtectedHistoricalRows(sheet, fillRows, periods, columns, ctx, auditRows);
    filledCells += cleanupResult.clearedCells;
    commentsAdded += cleanupResult.commentsAdded;
    warnings.push(...cleanupResult.warnings);

    for (const fillRow of fillRows) {
      let effectiveFillRow = fillRow;
      const rowNotes = new Set<string>();
      let unresolved = 0;

      if (fillRow.classification === "formula") {
        continue;
      }

      if (fillRow.classification === "unused" || (!fillRow.concepts?.length && !fillRow.resolver)) {
        const llmFillRow = await llmAssistedFillRow(fillRow, company, periods, ctx, llmState);
        if (llmFillRow) {
          effectiveFillRow = llmFillRow;
          rowNotes.add(llmFillRow.comment || "LLM-assisted EDGAR concept mapping was applied after deterministic row matching did not find a confident mapping.");
        } else {
          const sourceCell = labelCell(sheet, fillRow.row);
          if (fillRow.noFillComment && canAddComment(sourceCell)) {
            if (fillRow.noFillComment.startsWith("Cannot find exact schedule value in EDGAR")) sourceCell.note = "";
            addComment(sourceCell, fillRow.noFillComment);
            commentsAdded += 1;
          }
          warnings.push(`${fillRow.label}: left unchanged because no confident EDGAR match was found.`);
          continue;
        }
      }

      for (let index = 0; index < periods.length; index += 1) {
        const period = periods[index];
        const col = columns[index];
        const cell = sheet.getCell(effectiveFillRow.row, col);
        const writeDecision = historicalWriteDecision(effectiveFillRow, cell);
        if (!writeDecision.writable) {
          if (shouldAuditSkippedWrite(writeDecision)) {
            auditRows.push(skippedMappingAuditRow(sheet, cell, effectiveFillRow, period, writeDecision));
          }
          continue;
        }

        let resolved = resolveRowFromPackage(effectiveFillRow, period, normalizedPackage) ?? resolveRow(effectiveFillRow, period, ctx);
        if (resolved.value === null && effectiveFillRow === fillRow && fillRow.classification === "unused") {
          const llmFillRow = await llmAssistedFillRow(fillRow, company, periods, ctx, llmState);
          if (llmFillRow) {
            effectiveFillRow = llmFillRow;
            rowNotes.add(llmFillRow.comment || "LLM-assisted EDGAR concept mapping was applied after deterministic row matching did not find a value.");
            resolved = resolveRowFromPackage(effectiveFillRow, period, normalizedPackage) ?? resolveRow(effectiveFillRow, period, ctx);
          }
        }

        if (resolved.value === null || Number.isNaN(resolved.value)) {
          unresolved += 1;
          continue;
        }

        cell.value = resolved.value / (effectiveFillRow.scale ?? 1);
        filledCells += 1;

        const auditNote = auditNoteForResolvedValue(effectiveFillRow, resolved);
        if (auditNote) rowNotes.add(auditNote);
        const confidence = effectiveFillRow.comment?.startsWith("LLM-assisted") ? "medium" : "high";
        const cellComment = mappingComment(effectiveFillRow, resolved, period, cell.value as number, confidence, auditNote);
        addComment(cell, cellComment);
        commentsAdded += 1;
        auditRows.push(mappingAuditRow(sheet, cell, effectiveFillRow, period, cell.value as number, resolved, confidence, auditNote));
      }

      if (unresolved && fillRow.classification === "partial") {
        rowNotes.add(fillRow.noFillComment || "Split / partial match: Needs review because EDGAR detail was insufficient for one or more periods.");
      } else if (unresolved) {
        rowNotes.add("Needs review: one or more historical periods were left unchanged because no matching EDGAR fact was found.");
      }

      if (rowNotes.size) {
        const sourceCell = labelCell(sheet, fillRow.row);
        if (canAddComment(sourceCell)) {
          addComment(sourceCell, Array.from(rowNotes).join(" "));
          commentsAdded += 1;
        }
      }

      if (unresolved) {
        warnings.push(`${effectiveFillRow.label}: ${unresolved} period(s) left unchanged because no matching SEC fact was found.`);
      }
    }
    warnings.push(...llmState.warnings);

    if (isStandardModelSheet) refreshDividendCachedResults(sheet, periods, columns);

    if (segmentSheet) {
      const segmentResult = fillSegmentAnalysis(segmentSheet, company, periods, columns, segmentRevenue, ctx, auditRows);
      filledCells += segmentResult.filledCells;
      commentsAdded += segmentResult.commentsAdded;
      warnings.push(...segmentResult.warnings);
    } else {
      warnings.push(`Could not find a "${SEGMENT_SHEET}" worksheet; segment revenue formulas were left untouched.`);
    }

    for (const fillRow of fillRows) {
      const hasAny = periods.some((_, index) => sheet.getCell(fillRow.row, columns[index]).value !== null);
      if (!hasAny && fillRow.concepts?.length && fillRow.classification !== "unused" && fillRow.classification !== "formula") {
        warnings.push(`${fillRow.label}: no matching SEC concept found.`);
      }
    }

    if (isStandardModelSheet) {
      const cashFlowClearResult = clearCashFlowStatementHistoricalInputs(sheet, periods, columns, auditRows);
      filledCells += cashFlowClearResult.clearedCells;
      warnings.push(...cashFlowClearResult.warnings);

      const shareRepurchaseClearResult = clearStaleShareRepurchaseAssumptionAmounts(sheet, periods, columns, auditRows);
      filledCells += shareRepurchaseClearResult.clearedCells;
      commentsAdded += shareRepurchaseClearResult.commentsAdded;
      warnings.push(...shareRepurchaseClearResult.warnings);

      const incomeFormulaResult = reconcileIncomeStatementFormulaRowsToEdgar(sheet, periods, columns, ctx, auditRows);
      filledCells += incomeFormulaResult.filledCells;
      commentsAdded += incomeFormulaResult.commentsAdded;
      warnings.push(...incomeFormulaResult.warnings);

      const netIncomeFormulaResult = reconcileNetIncomeFormulaRowsToEdgar(sheet, periods, columns, ctx, auditRows);
      filledCells += netIncomeFormulaResult.filledCells;
      commentsAdded += netIncomeFormulaResult.commentsAdded;
      warnings.push(...netIncomeFormulaResult.warnings);

      const balanceSheetTotalResult = reconcileBalanceSheetStatementTotalsToEdgar(sheet, periods, columns, ctx, auditRows);
      filledCells += balanceSheetTotalResult.filledCells;
      commentsAdded += balanceSheetTotalResult.commentsAdded;
      warnings.push(...balanceSheetTotalResult.warnings);

      const balanceSheetCheckResult = reconcileBalanceSheetCheck(sheet, periods, columns, auditRows);
      filledCells += balanceSheetCheckResult.filledCells;
      commentsAdded += balanceSheetCheckResult.commentsAdded;
      warnings.push(...balanceSheetCheckResult.warnings);
    }

    restoreWorkbookLabels(workbook, workbookSnapshot);
    clearStaleFormulaErrorResults(workbook);
    refreshHistoricalFormulaCachedResults(workbook, columns, unique([sheet.name, MODEL_SHEET, SEGMENT_SHEET]));
    ensureFormulaDisplayCaches(workbook, columns, unique([sheet.name, MODEL_SHEET, SEGMENT_SHEET]));
    requestFullWorkbookRecalculation(workbook);

    const validationErrors = validateWorkbookBeforeReturn(workbook, periods, columns, ctx, warnings, sheet.name, profile);
    validationErrors.push(...validateWorkbookPreservation(workbook, workbookSnapshot));
    if (validationErrors.length) {
      return jsonError(`Validation failed: ${validationErrors.slice(0, 6).join(" | ")}`, 422);
    }

    addMappingAuditSheet(workbook, auditRows);

    const output = Buffer.from((await workbook.xlsx.writeBuffer()) as ArrayBuffer);
    const outputName = `${company.ticker}_historicals_filled.xlsx`;
    const summary = encodeURIComponent(
      JSON.stringify({
        companyName: company.title,
        ticker: company.ticker,
        templateProfile: profile.kind,
        templateProfileConfidence: profile.confidence,
        periods,
        filledCells,
        commentsAdded,
        warnings: unique(warnings).slice(0, 8)
      })
    );

    return new NextResponse(output, {
      headers: {
        "content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "content-disposition": `attachment; filename="${outputName}"`,
        "x-output-filename": outputName,
        "x-fill-summary": summary
      }
    });
  } catch (error) {
    console.error(error);
    return jsonError(error instanceof Error ? error.message : "Unexpected fill error.", 500);
  }
}

async function findCompany(query: string): Promise<CompanyMatch> {
  const response = await fetch("https://www.sec.gov/files/company_tickers.json", { headers: SEC_HEADERS, next: { revalidate: 86_400 } });
  if (!response.ok) throw new Error("Could not load SEC ticker directory.");
  const directory = (await response.json()) as Record<string, { cik_str: number; ticker: string; title: string }>;
  const normalized = normalize(query);
  const matches = Object.values(directory).map((item) => ({
    cik: String(item.cik_str).padStart(10, "0"),
    ticker: item.ticker.toUpperCase(),
    title: item.title
  }));

  return (
    matches.find((company) => normalize(company.ticker) === normalized) ??
    matches.find((company) => normalize(company.title) === normalized) ??
    matches.find((company) => normalize(company.title).includes(normalized)) ??
    (() => {
      throw new Error(`No SEC company match found for "${query}".`);
    })()
  );
}

async function fetchCompanyFacts(cik: string) {
  const response = await fetch(`https://data.sec.gov/api/xbrl/companyfacts/CIK${cik}.json`, { headers: SEC_HEADERS });
  if (!response.ok) throw new Error("Could not load SEC company facts for that company.");
  return response.json();
}

async function fetchSegmentRevenueByPeriod(company: CompanyMatch, periods: string[]) {
  const filings = await fetchFilingRefs(company, 40);

  const annual = new Map<string, Map<string, SegmentMetrics>>();
  const quarterly = new Map<string, Map<string, SegmentMetrics>>();

  for (const filing of filings) {
    try {
      const accession = filing.accessionNumber.replace(/-/g, "");
      const cikNoZeros = String(Number(company.cik));
      const url = `https://www.sec.gov/Archives/edgar/data/${cikNoZeros}/${accession}/${filing.primaryDocument}`;
      const html = await fetch(url, { headers: { ...SEC_HEADERS, Accept: "text/html" } }).then((res) => (res.ok ? res.text() : ""));
      if (!html) continue;
      const parsed = parseInlineSegmentRevenue(html, filing.form);
      for (const [period, values] of parsed.entries()) {
        if (isTenK(filing.form)) annual.set(period, values);
        else quarterly.set(period, values);
      }
    } catch {
      // Segment data is supplemental; model-level company facts still fill the workbook.
    }
  }

  deriveSegmentFourthQuarters(quarterly, annual);

  const wanted = new Set(periods);
  const labels = new Set<string>();
  quarterly.forEach((values, period) => {
    if (!wanted.has(period)) return;
    values.forEach((_, label) => labels.add(label));
  });

  return Array.from(labels)
    .sort(segmentSort)
    .slice(0, 6)
    .map((label) => ({
      label,
      values: new Map(periods.map((period) => [period, quarterly.get(period)?.get(label)?.revenue ?? 0])),
      operatingIncome: new Map(periods.map((period) => [period, quarterly.get(period)?.get(label)?.operatingIncome ?? 0])),
      depreciationAmortization: new Map(periods.map((period) => [period, quarterly.get(period)?.get(label)?.depreciationAmortization ?? 0]))
    }));
}

async function fetchInlineFactContext(company: CompanyMatch, periods: string[]): Promise<ResolveContext> {
  const duration = new Map<string, Map<string, FactSource>>();
  const instant = new Map<string, Map<string, FactSource>>();
  const filings = await fetchFilingRefs(company, 60);
  const wanted = new Set(periods);

  for (const filing of filings) {
    try {
      const accession = filing.accessionNumber.replace(/-/g, "");
      const cikNoZeros = String(Number(company.cik));
      const url = `https://www.sec.gov/Archives/edgar/data/${cikNoZeros}/${accession}/${filing.primaryDocument}`;
      const html = await fetch(url, { headers: { ...SEC_HEADERS, Accept: "text/html" } }).then((res) => (res.ok ? res.text() : ""));
      if (!html) continue;
      mergeInlineFacts(html, filing, wanted, { duration, instant }, url);
    } catch {
      // Inline facts are a supplement to SEC companyfacts; keep filling with the facts already available.
    }
  }

  return { duration, instant };
}

async function fetchFilingRefs(company: CompanyMatch, limit: number) {
  const response = await fetch(`https://data.sec.gov/submissions/CIK${company.cik}.json`, { headers: SEC_HEADERS });
  if (!response.ok) return [];
  const submissions = await response.json();
  const filings: FilingRef[] = [];
  collectFilingRefs(submissions?.filings?.recent, filings);

  for (const file of submissions?.filings?.files ?? []) {
    if (filings.length >= limit) break;
    try {
      const older = await fetch(`https://data.sec.gov/submissions/${file.name}`, { headers: SEC_HEADERS }).then((res) => (res.ok ? res.json() : null));
      collectFilingRefs(older, filings);
    } catch {
      // Older filing index fetches are best effort.
    }
  }

  return filings.slice(0, limit);
}

function collectFilingRefs(source: any, filings: FilingRef[]) {
  if (!source?.form) return;
  source.form.forEach((form: string, index: number) => {
    const filing = {
      form,
      filingDate: source.filingDate[index],
      accessionNumber: source.accessionNumber[index],
      primaryDocument: source.primaryDocument[index]
    };
    if ((isTenQ(filing.form) || isTenK(filing.form)) && filing.primaryDocument?.endsWith(".htm")) {
      filings.push(filing);
    }
  });
}

function mergeInlineFacts(html: string, filing: FilingRef, wanted: Set<string>, ctx: ResolveContext, sourceUrl = "") {
  const contexts = parseInlineContexts(html);
  for (const match of html.matchAll(/<ix:nonFraction\b([^>]*)>([\s\S]*?)<\/ix:nonFraction>/g)) {
    const attrs = match[1];
    const unitRef = attrs.match(/\bunitRef="([^"]+)"/)?.[1] ?? "";
    if (!/usd|shares/i.test(unitRef)) continue;
    const name = attrs.match(/\bname="([^"]+)"/)?.[1] ?? "";
    const concept = name.includes(":") ? name.split(":").pop() ?? name : name;
    const contextRef = attrs.match(/\bcontextRef="([^"]+)"/)?.[1];
    if (!concept || !contextRef) continue;
    if (["AssetsCurrent", "OtherAssets"].includes(concept)) continue;
    const context = contexts.get(contextRef);
    if (!context?.period || !wanted.has(context.period)) continue;
    if (context.hasDimensions) continue;
    const value = ixNumber(match[2], attrs);
    if (value === null) continue;
    const source = {
      concept,
      label: concept,
      value,
      sourceUrl,
      form: filing.form,
      filed: filing.filingDate,
      accn: filing.accessionNumber,
      start: context.start,
      end: context.end
    };
    setSource(context.instant ? ctx.instant : ctx.duration, context.period, concept, source);
  }
}

function parseInlineContexts(html: string) {
  const contexts = new Map<string, InlineContext>();
  for (const match of html.matchAll(/<xbrli:context\b[^>]*id="([^"]+)"[\s\S]*?<\/xbrli:context>/g)) {
    const body = match[0];
    const instant = body.match(/<xbrli:instant>([^<]+)<\/xbrli:instant>/)?.[1];
    if (instant) {
      contexts.set(match[1], { period: periodKeyFromDate(instant), instant: true, end: instant, hasDimensions: hasInlineDimensions(body) });
      continue;
    }
    const start = body.match(/<xbrli:startDate>([^<]+)<\/xbrli:startDate>/)?.[1];
    const end = body.match(/<xbrli:endDate>([^<]+)<\/xbrli:endDate>/)?.[1];
    if (start && end) {
      contexts.set(match[1], { period: periodKeyFromDate(end), instant: false, start, end, hasDimensions: hasInlineDimensions(body) });
    }
  }
  return contexts;
}

function hasInlineDimensions(contextBody: string) {
  return /<xbrldi:(explicitMember|typedMember)\b/i.test(contextBody);
}

function mergeContexts(target: ResolveContext, source: ResolveContext) {
  source.instant.forEach((facts, period) => facts.forEach((fact, concept) => setSource(target.instant, period, concept, fact)));
  source.duration.forEach((facts, period) => facts.forEach((fact, concept) => setSource(target.duration, period, concept, fact)));
}

type SegmentMetrics = {
  revenue?: number;
  operatingIncome?: number;
  depreciationAmortization?: number;
};

type InlineSegmentContext = {
  period: string | null;
  members: string[];
};

function parseInlineSegmentRevenue(html: string, form: string) {
  const contexts = new Map<string, InlineSegmentContext>();
  for (const match of html.matchAll(/<xbrli:context\b[^>]*id="([^"]+)"[\s\S]*?<\/xbrli:context>/g)) {
    const body = match[0];
    const period = contextPeriod(body, form);
    const members = Array.from(body.matchAll(/<xbrldi:explicitMember\b[^>]*>([^<]+)<\/xbrldi:explicitMember>/g)).map((member) => member[1]);
    contexts.set(match[1], { period, members });
  }

  const byPeriod = new Map<string, Map<string, SegmentMetrics>>();
  const componentRevenueByPeriod = new Map<string, Map<string, number>>();
  const totalRevenueByPeriod = new Map<string, { value: number; priority: number }>();
  const seenComponentFacts = new Set<string>();

  for (const match of html.matchAll(/<ix:nonFraction\b([^>]*)>([\s\S]*?)<\/ix:nonFraction>/g)) {
    const attrs = match[1];
    const concept = attrs.match(/\bname="([^"]+)"/)?.[1] ?? "";
    const contextRef = attrs.match(/\bcontextRef="([^"]+)"/)?.[1];
    if (!contextRef) continue;
    const context = contexts.get(contextRef);
    if (!context?.period) continue;

    const value = ixNumber(match[2], attrs);
    if (value === null) continue;

    const totalPriority = totalRevenuePriority(concept, context.members);
    if (totalPriority) {
      const existing = totalRevenueByPeriod.get(context.period);
      if (!existing || totalPriority > existing.priority || (totalPriority === existing.priority && Math.abs(value) > Math.abs(existing.value))) {
        totalRevenueByPeriod.set(context.period, { value, priority: totalPriority });
      }
    }

    const componentLabel = revenueComponentSegmentLabel(concept, context.members);
    if (componentLabel) {
      const factKey = `${context.period}|${componentLabel}|${concept}|${contextRef}|${value}`;
      if (!seenComponentFacts.has(factKey)) {
        seenComponentFacts.add(factKey);
        const periodValues = componentRevenueByPeriod.get(context.period) ?? new Map<string, number>();
        periodValues.set(componentLabel, (periodValues.get(componentLabel) ?? 0) + value);
        componentRevenueByPeriod.set(context.period, periodValues);
      }
    }

    const metric = segmentMetric(concept);
    if (!metric) continue;

    const label = segmentLabelFromMembers(context.members);
    if (!label) continue;

    const periodValues = byPeriod.get(context.period) ?? new Map<string, SegmentMetrics>();
    const metrics = periodValues.get(label) ?? {};
    const existing = metrics[metric];
    if (existing === undefined || preferSegmentFact(context.members, value, existing)) {
      metrics[metric] = value;
      periodValues.set(label, metrics);
    }
    byPeriod.set(context.period, periodValues);
  }

  applyComponentRevenueRollups(byPeriod, componentRevenueByPeriod, totalRevenueByPeriod);

  return byPeriod;
}

function revenueComponentSegmentLabel(concept: string, members: string[]) {
  const local = concept.split(":").pop() ?? concept;
  if (!/^Revenues$/i.test(local)) return null;
  if (members.some(isReportableSegmentMember)) return null;
  if (members.some(isNonRevenueComponentMember)) return null;

  const joined = members.join(" ");
  if (
    /InvestmentBankingMember|PrincipalTransactionsRevenueMember|CommissionsAndOtherFeesMember|InterestRevenueMember/i.test(joined)
  ) {
    return "Investment Banking & Capital Markets";
  }
  if (/AssetManagement1Member|ProductAndServiceOtherMember/i.test(joined)) return "Asset Management and Other";
  return null;
}

function totalRevenuePriority(concept: string, members: string[]) {
  const local = concept.split(":").pop() ?? concept;
  if (!/^(Revenues|RevenueFromContractWithCustomerExcludingAssessedTax|SalesRevenueNet)$/i.test(local)) return 0;
  if (members.length === 0) return /^Revenues$/i.test(local) ? 100 : 90;
  if (members.length === 1 && /OperatingSegmentsMember/i.test(members[0])) return /^Revenues$/i.test(local) ? 80 : 70;
  return 0;
}

function applyComponentRevenueRollups(
  byPeriod: Map<string, Map<string, SegmentMetrics>>,
  componentRevenueByPeriod: Map<string, Map<string, number>>,
  totalRevenueByPeriod: Map<string, { value: number; priority: number }>
) {
  componentRevenueByPeriod.forEach((componentRevenue, period) => {
    if (!componentRevenue.size) return;
    const totalRevenue = totalRevenueByPeriod.get(period)?.value;
    if (totalRevenue === undefined) return;

    const componentTotal = sumNumbers(Array.from(componentRevenue.values()));
    if (!segmentRevenueTies(componentTotal, totalRevenue)) return;

    const directRevenue = byPeriod.get(period);
    const directTotal = directRevenue ? sumSegmentRevenue(directRevenue) : 0;
    const directTies = directRevenue ? segmentRevenueTies(directTotal, totalRevenue) : false;
    if (directTies && Math.abs(directTotal - componentTotal) <= 100_000) return;

    const periodValues = byPeriod.get(period) ?? new Map<string, SegmentMetrics>();
    componentRevenue.forEach((revenue, label) => {
      const metrics = periodValues.get(label) ?? {};
      metrics.revenue = revenue;
      periodValues.set(label, metrics);
    });
    byPeriod.set(period, periodValues);
  });
}

function segmentMetric(concept: string): keyof SegmentMetrics | null {
  const local = concept.split(":").pop() ?? concept;
  if (/^(RevenueFromContractWithCustomerExcludingAssessedTax|Revenues|SalesRevenueNet)$/i.test(local)) return "revenue";
  if (/(OperatingIncomeLoss|SegmentProfitLoss|IncomeLossFromContinuingOperationsBeforeIncomeTaxes)/i.test(local)) return "operatingIncome";
  if (/(DepreciationDepletionAndAmortization|DepreciationAndAmortization|DepreciationExpense)/i.test(local)) return "depreciationAmortization";
  return null;
}

function contextPeriod(contextXml: string, form: string) {
  const start = contextXml.match(/<xbrli:startDate>([^<]+)<\/xbrli:startDate>/)?.[1];
  const end = contextXml.match(/<xbrli:endDate>([^<]+)<\/xbrli:endDate>/)?.[1];
  if (!start || !end) return null;
  const endDate = new Date(`${end}T00:00:00Z`);
  const startDate = new Date(`${start}T00:00:00Z`);
  const days = (endDate.getTime() - startDate.getTime()) / 86_400_000;
  if (isTenK(form) && days < 330) return null;
  if (!isTenK(form) && days > 115) return null;
  const year = endDate.getUTCFullYear();
  const quarter = Math.floor(endDate.getUTCMonth() / 3) + 1;
  if (isTenK(form)) return `4Q${String(year).slice(-2)}`;
  return `${quarter}Q${String(year).slice(-2)}`;
}

function segmentLabelFromMembers(members: string[]) {
  const joined = members.join(" ");
  if (/InvestmentBankingAndCapitalMarkets/i.test(joined)) return "Investment Banking & Capital Markets";
  if (/AssetManagementAndOther/i.test(joined)) return "Asset Management and Other";
  if (/AssetManagementSegmentMember/i.test(joined)) return "Asset Management and Other";
  const product = members.find(
    (member) =>
      member.includes("ProductOrServiceAxis") ||
      /ServiceLine|OtherServiceLine/i.test(member)
  );
  const serviceLineProduct = product && /(?:Collection|Landfill|EnvironmentalSolutions|Transfer|Other|ProfessionalServices).*ServiceLine/i.test(product)
    ? product
    : null;
  const reportableSegment = members.find((member) => {
    if (/OperatingSegmentsMember/i.test(member) && serviceLineProduct) return false;
    return /BusinessSegment|OperatingSegment|StatementBusinessSegments|SegmentAxis/i.test(member);
  });
  const source = serviceLineProduct || reportableSegment || product || joined;
  if (/CollectionServiceLine(?:Residential|Smallcontainer|Largecontainer|Other)?Member/i.test(source)) return "Collection";
  if (/LandfillServiceLineMember/i.test(source)) return "Landfill";
  if (/EnvironmentalSolutionsServiceLineMember/i.test(source)) return "Environmental Solutions";
  if (/TransferServiceLineMember/i.test(source)) return "Transfer";
  if (/OtherServiceLineMember/i.test(source)) return "Other";
  if (/ProfessionalServices/i.test(source)) return "Professional Services and Other";
  return cleanSegmentMember(source);
}

function isReportableSegmentMember(member: string) {
  return /BusinessSegment|OperatingSegment|StatementBusinessSegments|SegmentAxis|SegmentMember/i.test(member);
}

function isNonRevenueComponentMember(member: string) {
  return /RelatedParty|Reclassification|AccumulatedOtherComprehensiveIncome|Aoci|Geographic|Region|Country|MinimumMember|MaximumMember/i.test(member);
}

function cleanSegmentMember(member: string) {
  const local = member
    .split(":")
    .pop()
    ?.replace(/(Member|Segment|OperatingSegments|BusinessSegments|ServiceLine)$/gi, "")
    .replace(/([a-z])([A-Z])/g, "$1 $2")
    .replace(/\bAnd\b/g, "and")
    .replace(/\s+/g, " ")
    .trim();
  if (!local || /Consolidated|Corporate|Geographic|North America|Europe|Asia|Other Countries|Operating Segments/i.test(local)) return null;
  return local;
}

function ixNumber(text: string, attrs: string) {
  const raw = decodeXml(text.replace(/<[^>]+>/g, "").trim()).replace(/,/g, "");
  if (!raw || raw === "-") return null;
  const parsed = Number(raw.replace(/[()]/g, ""));
  if (!Number.isFinite(parsed)) return null;
  const scale = Number(attrs.match(/\bscale="([^"]+)"/)?.[1] ?? 0);
  const sign = attrs.includes('sign="-"') || /^\(.+\)$/.test(raw) ? -1 : 1;
  return parsed * Math.pow(10, scale) * sign;
}

function decodeXml(value: string) {
  return value
    .replace(/&amp;/g, "&")
    .replace(/&nbsp;/g, " ")
    .replace(/&#x([0-9a-f]+);/gi, (_, hex) => String.fromCharCode(parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (_, decimal) => String.fromCharCode(Number(decimal)));
}

function preferSegmentFact(members: string[], value: number, existing: number) {
  const joined = members.join(" ");
  if (/IntersegmentEliminationMember/i.test(joined)) return false;
  if (members.length === 1 && /OperatingSegmentsMember/i.test(members[0])) return false;
  if (!/OperatingSegmentsMember/i.test(joined) && value > 0) return true;
  return Math.abs(value) > Math.abs(existing);
}

function sumSegmentRevenue(values: Map<string, SegmentMetrics>) {
  let total = 0;
  values.forEach((metrics, label) => {
    if (isAggregateSegmentLabel(label)) return;
    total += metrics.revenue ?? 0;
  });
  return total;
}

function sumNumbers(values: number[]) {
  return values.reduce((total, value) => total + value, 0);
}

function segmentRevenueTies(actual: number, expected: number) {
  return Math.abs(actual - expected) <= Math.max(100_000, Math.abs(expected) * 0.0005);
}

function isAggregateSegmentLabel(label: string) {
  return /^(Operating Segments|OperatingSegments|Total|Consolidated)$/i.test(label.trim());
}

function deriveSegmentFourthQuarters(quarterly: Map<string, Map<string, SegmentMetrics>>, annual: Map<string, Map<string, SegmentMetrics>>) {
  for (const [period, annualValues] of annual.entries()) {
    const year = period.slice(2);
    const q1 = quarterly.get(`1Q${year}`);
    const q2 = quarterly.get(`2Q${year}`);
    const q3 = quarterly.get(`3Q${year}`);
    if (!q1 || !q2 || !q3) continue;
    const q4 = quarterly.get(period) ?? new Map<string, SegmentMetrics>();
    annualValues.forEach((metrics, label) => {
      const q4Metrics = q4.get(label) ?? {};
      (["revenue", "operatingIncome", "depreciationAmortization"] as const).forEach((metric) => {
        const annualValue = metrics[metric];
        if (annualValue === undefined) return;
        q4Metrics[metric] =
          annualValue - (q1.get(label)?.[metric] ?? 0) - (q2.get(label)?.[metric] ?? 0) - (q3.get(label)?.[metric] ?? 0);
      });
      q4.set(label, q4Metrics);
    });
    quarterly.set(period, q4);
  }
}

function segmentSort(a: string, b: string) {
  const order = [
    "Investment Banking & Capital Markets",
    "Asset Management and Other",
    "Collection",
    "Landfill",
    "Environmental Solutions",
    "Transfer",
    "Other",
    "Professional Services and Other"
  ];
  return (order.indexOf(a) === -1 ? 99 : order.indexOf(a)) - (order.indexOf(b) === -1 ? 99 : order.indexOf(b));
}

function buildFactContext(payload: any, company?: CompanyMatch): ResolveContext {
  const taxonomies = Object.values(payload?.facts ?? {}) as any[];
  const duration = new Map<string, Map<string, FactSource>>();
  const instant = new Map<string, Map<string, FactSource>>();
  const cumulativeDuration = new Map<string, Map<string, FactSource>>();
  const annualDuration = new Map<string, Map<string, FactSource>>();

  for (const taxonomy of taxonomies) {
    for (const [concept, detail] of Object.entries<any>(taxonomy)) {
      const label = detail.label || concept;
      const units = detail.units ?? {};
      const unitFacts: SecFact[] = units.USD ?? units.shares ?? units["USD/shares"] ?? Object.values(units)[0] ?? [];
      for (const fact of unitFacts) {
        if (!isUsableFact(fact)) continue;
        const instantFact = isInstantFact(fact);
        const period = fact.end ? periodKeyFromDate(fact.end) : periodKey(fact);
        if (!period) continue;
        const source = {
          concept,
          label,
          value: fact.val,
          sourceUrl: sourceUrlForAccession(company?.cik ?? payload?.cik, fact.accn),
          form: fact.form,
          fp: fact.fp,
          filed: fact.filed,
          accn: fact.accn,
          start: fact.start,
          end: fact.end
        };
        if (!instantFact && isAnnualDurationFact(fact)) {
          setSource(annualDuration, period, concept, source);
          const annualPeriod = annualPeriodKey(fact);
          if (annualPeriod) setSource(duration, annualPeriod, concept, source);
        } else if (!instantFact && isYearToDateFact(fact)) {
          setSource(cumulativeDuration, period, concept, source);
        } else if (instantFact || isQuarterDurationFact(fact)) {
          setSource(instantFact ? instant : duration, period, concept, source);
          if (instantFact && isAnnualInstantFact(fact)) {
            const annualPeriod = annualPeriodKey(fact);
            if (annualPeriod) setSource(instant, annualPeriod, concept, source);
          }
        }
      }
    }
  }

  deriveQuarterlies(duration, cumulativeDuration, annualDuration);

  return { duration, instant };
}

function sourceUrlForAccession(cik: string | number | undefined, accn?: string) {
  if (!cik || !accn) return "";
  return `https://www.sec.gov/Archives/edgar/data/${Number(cik)}/${accn.replace(/-/g, "")}/`;
}

function setSource(map: Map<string, Map<string, FactSource>>, period: string, concept: string, source: FactSource) {
  const periodFacts = map.get(period) ?? new Map<string, FactSource>();
  const existing = periodFacts.get(concept);
  if (existing && !preferSource(period, source, existing)) {
    map.set(period, periodFacts);
    return;
  }
  periodFacts.set(concept, source);
  map.set(period, periodFacts);
}

function deriveQuarterlies(
  duration: Map<string, Map<string, FactSource>>,
  cumulativeDuration: Map<string, Map<string, FactSource>>,
  annualDuration: Map<string, Map<string, FactSource>>
) {
  for (const [period, cumulativeFacts] of cumulativeDuration.entries()) {
    const quarter = Number(period[0]);
    const year = period.slice(2);
    if (quarter === 1) {
      cumulativeFacts.forEach((source, concept) => setSource(duration, period, concept, source));
    } else if (quarter === 2) {
      cumulativeFacts.forEach((source, concept) => {
        if (!canDeriveQuarterlyConcept(concept)) return;
        const q1 = cumulativeDuration.get(`1Q${year}`)?.get(concept) ?? duration.get(`1Q${year}`)?.get(concept);
        if (q1 && !duration.get(period)?.get(concept)) {
          setSource(duration, period, concept, {
            ...source,
            label: `${source.label} (derived Q2)`,
            value: source.value - q1.value,
            derivedTotalValue: source.value,
            derivedTotalLabel: source.label,
            derivedPriorPeriods: [`1Q${year}`]
          });
        }
      });
    } else if (quarter === 3) {
      cumulativeFacts.forEach((source, concept) => {
        if (!canDeriveQuarterlyConcept(concept)) return;
        const q2Cumulative = cumulativeDuration.get(`2Q${year}`)?.get(concept);
        const q1 = duration.get(`1Q${year}`)?.get(concept);
        const q2 = duration.get(`2Q${year}`)?.get(concept);
        const priorValue = q2Cumulative?.value ?? (q1 && q2 ? q1.value + q2.value : null);
        if (priorValue !== null && !duration.get(period)?.get(concept)) {
          setSource(duration, period, concept, {
            ...source,
            label: `${source.label} (derived Q3)`,
            value: source.value - priorValue,
            derivedTotalValue: source.value,
            derivedTotalLabel: source.label,
            derivedPriorPeriods: [`1Q${year}`, `2Q${year}`]
          });
        }
      });
    }
  }

  for (const [period, annualFacts] of annualDuration.entries()) {
    const year = period.slice(2);
    for (const [concept, annual] of annualFacts.entries()) {
      if (!canDeriveQuarterlyConcept(concept)) continue;
      const existingQuarter = duration.get(period)?.get(concept);
      if (existingQuarter && existingQuarter.derivedTotalValue === undefined && isQuarterDurationSource(existingQuarter)) continue;
      const nineMonth = cumulativeDuration.get(`3Q${year}`)?.get(concept);
      const q1 = duration.get(`1Q${year}`)?.get(concept);
      const q2 = duration.get(`2Q${year}`)?.get(concept);
      const q3 = duration.get(`3Q${year}`)?.get(concept);
      const firstNineMonths = nineMonth?.value ?? (q1 && q2 && q3 ? q1.value + q2.value + q3.value : null);
      if (firstNineMonths === null) continue;
      setSource(duration, period, concept, {
        ...annual,
        concept,
        label: `${annual.label} (derived Q4)`,
        value: annual.value - firstNineMonths,
        derivedTotalValue: annual.value,
        derivedTotalLabel: annual.label,
        derivedPriorPeriods: [`1Q${year}`, `2Q${year}`, `3Q${year}`]
      });
    }
  }
}

function isQuarterDurationSource(source: FactSource) {
  if (!source.start || !source.end || !["Q1", "Q2", "Q3", "Q4"].includes(source.fp ?? "")) return false;
  return factDurationDays({ start: source.start, end: source.end } as SecFact) <= 115;
}

function canDeriveQuarterlyConcept(concept: string) {
  return !/WeightedAverage|EarningsPerShare|SharesOutstanding/i.test(concept);
}

function isYearToDateFact(fact: SecFact) {
  if (!fact.start || !fact.end) return false;
  const days = (new Date(`${fact.end}T00:00:00Z`).getTime() - new Date(`${fact.start}T00:00:00Z`).getTime()) / 86_400_000;
  return days > 115 && fact.fp !== "FY";
}

function isAnnualDurationFact(fact: SecFact) {
  if (!fact.start || !fact.end || fact.fp !== "FY" || !isTenK(fact.form)) return false;
  return factDurationDays(fact) >= 330;
}

function isAnnualInstantFact(fact: SecFact) {
  return Boolean(fact.end && fact.fp === "FY" && isTenK(fact.form));
}

function annualPeriodKey(fact: SecFact) {
  const fiscalYear = fact.fy ?? (fact.end ? new Date(`${fact.end}T00:00:00Z`).getUTCFullYear() : null);
  if (!fiscalYear) return null;
  return `FY${String(fiscalYear).slice(-2)}`;
}

function isQuarterDurationFact(fact: SecFact) {
  if (!fact.start || !fact.end || !["Q1", "Q2", "Q3", "Q4"].includes(fact.fp ?? "")) return false;
  return factDurationDays(fact) <= 115;
}

function factDurationDays(fact: SecFact) {
  if (!fact.start || !fact.end) return 0;
  return (new Date(`${fact.end}T00:00:00Z`).getTime() - new Date(`${fact.start}T00:00:00Z`).getTime()) / 86_400_000;
}

function isUsableFact(fact: SecFact) {
  return typeof fact.val === "number" && Boolean(fact.end) && (isTenK(fact.form) || isTenQ(fact.form));
}

function isInstantFact(fact: SecFact) {
  return Boolean(fact.frame?.endsWith("I")) || !["Q1", "Q2", "Q3", "Q4", "FY"].includes(fact.fp ?? "");
}

function periodKey(fact: SecFact) {
  const fy = fact.fy;
  if (!fy) return null;
  const yy = String(fy).slice(-2);
  if (fact.fp === "Q1") return `1Q${yy}`;
  if (fact.fp === "Q2") return `2Q${yy}`;
  if (fact.fp === "Q3") return `3Q${yy}`;
  if (fact.fp === "Q4") return `4Q${yy}`;
  if (fact.fp === "FY") return `4Q${yy}`;
  return null;
}

function periodKeyFromDate(date: string) {
  const parsed = new Date(`${date}T00:00:00Z`);
  if (Number.isNaN(parsed.getTime())) return null;
  const quarter = Math.floor(parsed.getUTCMonth() / 3) + 1;
  return `${quarter}Q${String(parsed.getUTCFullYear()).slice(-2)}`;
}

function preferSource(period: string, next: FactSource, current: FactSource) {
  const nextScore = sourceScore(period, next);
  const currentScore = sourceScore(period, current);
  if (nextScore !== currentScore) return nextScore > currentScore;
  return (next.filed ?? "") > (current.filed ?? "");
}

function sourceScore(period: string, source: FactSource) {
  const quarter = String(periodQuarter(period));
  let score = 0;
  const durationDays = source.start && source.end ? factDurationDays({ start: source.start, end: source.end } as SecFact) : 0;
  if (isAnnualPeriod(period)) {
    if (isTenK(source.form)) score += 35;
    if (source.fp === "FY") score += 15;
  } else if (quarter === "4") {
    if (isTenK(source.form)) score += 20;
    if (source.fp === "FY" || source.fp === "Q4") score += 10;
  } else {
    if (isTenQ(source.form)) score += 20;
    if (source.fp === `Q${quarter}`) score += 10;
  }
  if (durationDays > 115 && source.derivedTotalValue === undefined) score -= 50;
  if (durationDays > 0 && durationDays <= 115) score += 5;
  if (source.derivedTotalValue !== undefined) score += 5;
  return score;
}

function isTenK(form?: string) {
  return form === "10-K" || form === "10-K/A";
}

function isTenQ(form?: string) {
  return form === "10-Q" || form === "10-Q/A";
}

function choosePeriods(ctx: ResolveContext, maxColumns: number) {
  const periods = unique([...ctx.duration.keys(), ...ctx.instant.keys()]).sort(comparePeriods);
  return periods.filter((period) => !isAnnualPeriod(period)).slice(-maxColumns);
}

function comparePeriods(a: string, b: string) {
  const [aq, ay] = [periodQuarter(a), periodYear(a)];
  const [bq, by] = [periodQuarter(b), periodYear(b)];
  return ay === by ? aq - bq : ay - by;
}

function previousPeriod(period: string) {
  if (isAnnualPeriod(period)) return `FY${String(periodYear(period) - 1).slice(-2)}`;
  const quarter = periodQuarter(period);
  const year = periodYear(period);
  if (!quarter || Number.isNaN(year)) return null;
  if (quarter === 1) return `4Q${String(year - 1).slice(-2)}`;
  return `${quarter - 1}Q${String(year).slice(-2)}`;
}

function isSupportedPeriodKey(period: string) {
  return /^[1-4]Q\d{2}$/.test(period) || isAnnualPeriod(period);
}

function isAnnualPeriod(period: string) {
  return /^FY\d{2}$/.test(period);
}

function isQuarterPeriod(period: string) {
  return /^[1-4]Q\d{2}$/.test(period);
}

function isFourthQuarterPeriod(period: string) {
  return /^4Q\d{2}$/.test(period);
}

function periodYear(period: string) {
  const match = period.match(/^(?:[1-4]Q|FY)(\d{2})$/);
  return match ? Number(`20${match[1]}`) : NaN;
}

function periodYearSuffix(period: string) {
  return String(periodYear(period)).slice(-2);
}

function periodQuarter(period: string) {
  if (isAnnualPeriod(period)) return 4;
  return Number(period[0]);
}

function blueColumns(sheet: ExcelJS.Worksheet) {
  const columnCounts = new Map<number, number>();
  const maxScanColumn = Math.min(sheet.columnCount, 120);
  for (let rowNumber = 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const label = rowLabel(sheet, rowNumber);
    if (!label) continue;
    for (let col = 6; col <= maxScanColumn; col += 1) {
      if (!isHardcodedBlueInput(sheet.getCell(rowNumber, col))) continue;
      columnCounts.set(col, (columnCounts.get(col) ?? 0) + 1);
    }
  }

  const candidates = Array.from(columnCounts.entries())
    .filter(([, count]) => count >= 3)
    .map(([col]) => col)
    .sort((a, b) => a - b);

  if (candidates.length) return candidates;

  const blueFallback = Array.from(columnCounts.keys()).sort((a, b) => a - b);
  if (blueFallback.length) return blueFallback;

  const periodColumns = periodHeaderColumns(sheet);
  if (periodColumns.length) return periodColumns;

  return [];
}

function periodHeaderColumns(sheet: ExcelJS.Worksheet) {
  let best: number[] = [];
  for (let rowNumber = 1; rowNumber <= Math.min(sheet.rowCount, 120); rowNumber += 1) {
    const cols: number[] = [];
    for (let col = 4; col <= Math.min(sheet.columnCount, 120); col += 1) {
      if (isSupportedPeriodKey(normalizePeriodLabel(cellDisplay(sheet.getCell(rowNumber, col))))) cols.push(col);
    }
    if (cols.length > best.length) best = cols;
  }
  return best;
}

function templatePeriodInfos(sheet: ExcelJS.Worksheet, columns: number[]) {
  let best: Array<{ period: string; isEstimate: boolean }> = [];
  let bestCount = 0;
  let bestQuarterCount = 0;
  for (let rowNumber = 1; rowNumber <= Math.min(sheet.rowCount, 80); rowNumber += 1) {
    const infos = columns.map((col) => {
      const label = cellDisplay(sheet.getCell(rowNumber, col));
      return {
        period: normalizePeriodLabel(label),
        isEstimate: /e\s*$/i.test(label.trim())
      };
    });
    const validCount = infos.filter((info) => isSupportedPeriodKey(info.period)).length;
    const quarterCount = infos.filter((info) => isQuarterPeriod(info.period)).length;
    if (validCount > bestCount || (validCount === bestCount && quarterCount > bestQuarterCount)) {
      best = infos;
      bestCount = validCount;
      bestQuarterCount = quarterCount;
    }
  }
  return bestCount ? best : [];
}

function normalizePeriodLabel(label: string) {
  const compact = label.trim().replace(/\s+/g, "").replace(/[’']/g, "").replace(/e$/i, "");
  const direct = compact.match(/^([1-4])Q(\d{2}|\d{4})$/i);
  if (direct) return `${direct[1]}Q${direct[2].slice(-2)}`;
  const qFirst = compact.match(/^Q([1-4])(\d{2}|\d{4})$/i);
  if (qFirst) return `${qFirst[1]}Q${qFirst[2].slice(-2)}`;
  const fiscalYear = compact.match(/^(?:FY(\d{2}|\d{4})|(\d{4}))A?$/i);
  if (fiscalYear) return `FY${(fiscalYear[1] ?? fiscalYear[2]).slice(-2)}`;
  return compact;
}

function cellDisplay(cell: ExcelJS.Cell) {
  const value = cell.value;
  if (typeof value === "string") return value;
  if (typeof value === "number") return String(value);
  if (value instanceof Date) return periodKeyFromDate(value.toISOString().slice(0, 10)) ?? "";
  if (value && typeof value === "object" && "result" in value && value.result !== undefined && value.result !== null) {
    return String(value.result);
  }
  return "";
}

function rowLabel(sheet: ExcelJS.Worksheet, rowNumber: number) {
  const candidates: string[] = [];
  for (const col of LABEL_COLUMNS) {
    const text = cellDisplay(sheet.getCell(rowNumber, col)).trim();
    if (!text || /^x$/i.test(text)) continue;
    candidates.push(text);
  }
  if (!candidates.length) return "";
  return candidates.sort((a, b) => scoreLabelCandidate(b) - scoreLabelCandidate(a))[0];
}

function scoreLabelCandidate(label: string) {
  let score = Math.min(label.length, 80);
  if (/[A-Za-z]{3,}/.test(label)) score += 40;
  if (/statement|revenue|expense|assets|liabilities|equity|cash|income|segment|total/i.test(label)) score += 20;
  if (/^=/.test(label)) score -= 30;
  return score;
}

function modelRowContext(
  sheet: ExcelJS.Worksheet,
  rowNumber: number,
  columns: number[],
  periodInfos: Array<{ period: string; isEstimate: boolean }>
): ModelRowContext {
  const label = rowLabel(sheet, rowNumber);
  const formulaCells = columns.filter((col) => hasFormula(sheet.getCell(rowNumber, col)));
  const hardcodedCells = columns.filter((col) => isModelHistoricalInput(sheet.getCell(rowNumber, col)));
  const labelCellForRow = labelCell(sheet, rowNumber);
  return {
    sheetName: sheet.name,
    row: rowNumber,
    label,
    sectionHeader: nearestSectionHeader(sheet, rowNumber, columns),
    previousLabel: nearestLabeledNeighbor(sheet, rowNumber, -1),
    nextLabel: nearestLabeledNeighbor(sheet, rowNumber, 1),
    indentation: labelCellForRow.alignment?.indent ?? 0,
    hasHistoricalFormula: formulaCells.length > 0,
    hasHardcodedInput: hardcodedCells.length > 0,
    hasNetRevenueInterestExpenseAbove: hasNetRevenueInterestExpenseBridgeAbove(sheet, rowNumber),
    subtotalFormula: firstFormulaInRow(sheet, rowNumber, columns),
    projectedColumns: periodInfos.filter((info) => info.isEstimate).length,
    signConvention: inferSignConvention(label)
  };
}

function hasNetRevenueInterestExpenseBridgeAbove(sheet: ExcelJS.Worksheet, rowNumber: number) {
  let foundInterestExpense = false;
  let foundNetRevenue = false;

  for (let row = rowNumber - 1; row >= Math.max(1, rowNumber - 40); row -= 1) {
    const label = rowLabel(sheet, row);
    if (!label) continue;
    const normalized = normalize(label);
    if (normalized === normalize("Income Statement")) break;
    if (normalized === normalize("Interest Expense")) foundInterestExpense = true;
    if (normalized === normalize("Net Revenue") || normalized === normalize("Revenue Net of Interest Expense")) foundNetRevenue = true;
    if (foundInterestExpense && foundNetRevenue) return true;
  }

  return false;
}

function nearestSectionHeader(sheet: ExcelJS.Worksheet, rowNumber: number, columns: number[]) {
  for (let row = rowNumber - 1; row >= Math.max(1, rowNumber - 25); row -= 1) {
    const label = rowLabel(sheet, row);
    if (!label) continue;
    const hasFinancialInput = columns.some((col) => cellDisplay(sheet.getCell(row, col)).trim() || hasFormula(sheet.getCell(row, col)));
    const labelCellForRow = labelCell(sheet, row);
    if (!hasFinancialInput || labelCellForRow.font?.bold || /schedule|assumptions|drivers/i.test(label)) return label;
  }
  return undefined;
}

function nearestLabeledNeighbor(sheet: ExcelJS.Worksheet, rowNumber: number, direction: 1 | -1) {
  const end = direction === 1 ? Math.min(sheet.rowCount, rowNumber + 4) : Math.max(1, rowNumber - 4);
  for (let row = rowNumber + direction; direction === 1 ? row <= end : row >= end; row += direction) {
    const label = rowLabel(sheet, row);
    if (label) return label;
  }
  return undefined;
}

function firstFormulaInRow(sheet: ExcelJS.Worksheet, rowNumber: number, columns: number[]) {
  for (const col of columns) {
    const formula = cellFormula(sheet.getCell(rowNumber, col));
    if (formula) return formula;
  }
  return undefined;
}

function cellFormula(cell: ExcelJS.Cell) {
  const value = cell.value;
  if (!value || typeof value !== "object") return null;
  if ("formula" in value && typeof value.formula === "string") return value.formula;
  if ("sharedFormula" in value && typeof value.sharedFormula === "string") return value.sharedFormula;
  return null;
}

function historicalWriteDecision(fillRow: FillRow, cell: ExcelJS.Cell): WriteDecision {
  if (hasFormula(cell)) {
    return { writable: false, reason: "existing formula cell", formulaPreserved: true };
  }
  if (fillRow.onlyBlankHistoricalInput && cell.value !== null) {
    return { writable: false, reason: "existing model hardcode preserved", formulaPreserved: false };
  }
  if (fillRow.allowBlankHistoricalInput && cell.value === null) {
    return { writable: true, formulaPreserved: false };
  }
  if (isInactiveHelperCell(fillRow, cell)) {
    return { writable: false, reason: "blank inactive/helper cell", formulaPreserved: false };
  }
  if (!isModelHistoricalInput(cell)) {
    return { writable: false, reason: "not an active historical input cell", formulaPreserved: false };
  }
  return { writable: true, formulaPreserved: false };
}

function normalizeSharedFormulas(workbook: ExcelJS.Workbook) {
  workbook.eachSheet((sheet) => {
    sheet.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        const value = cell.value;
        if (!value || typeof value !== "object" || !("sharedFormula" in value)) return;
        const formula = formulaForCell(cell);
        if (!formula) return;
        const result = "result" in value ? value.result : undefined;
        cell.value = result === undefined ? { formula } : { formula, result };
      });
    });
  });
}

function removeInvalidConditionalFormattingRules(workbook: ExcelJS.Workbook) {
  workbook.eachSheet((sheet) => {
    const worksheetModel = sheet as ExcelJS.Worksheet & {
      conditionalFormattings?: Array<{ rules?: Array<{ formulae?: unknown[]; type?: string }> }>;
    };
    if (!worksheetModel.conditionalFormattings?.length) return;
    worksheetModel.conditionalFormattings = worksheetModel.conditionalFormattings
      .map((formatting) => ({
        ...formatting,
        rules: (formatting.rules ?? []).filter((rule) => {
          if (rule.type !== "expression" && rule.type !== "cellIs") return true;
          return Array.isArray(rule.formulae) && rule.formulae.length > 0;
        })
      }))
      .filter((formatting) => formatting.rules?.length);
  });
}

function removeExternalWorkbookDefinedNames(workbook: ExcelJS.Workbook) {
  const definedNames = workbook.definedNames as ExcelJS.Workbook["definedNames"] & {
    model?: Array<{ name: string; ranges?: string[] }>;
  };
  if (!definedNames.model?.length) return;
  definedNames.model = definedNames.model
    .map((item) => ({
      ...item,
      ranges: (item.ranges ?? []).filter((range) => !isExternalWorkbookReference(range))
    }))
    .filter((item) => (item.ranges ?? []).length);
}

function isExternalWorkbookReference(reference: string) {
  return /(?:^|')\[[^\]]+\]/.test(reference) || /\[[^\]]+\.(?:xlsx|xlsm|xlsb|xls)\]/i.test(reference);
}

function shouldAuditSkippedWrite(decision: WriteDecision) {
  return decision.reason === "existing formula cell" || decision.reason === "blank inactive/helper cell";
}

function refreshDividendCachedResults(sheet: ExcelJS.Worksheet, periods: string[], columns: number[]) {
  const dividendRow = findLabelRow(sheet, "Dividends");
  const dividendsPaidRow = findLabelRow(sheet, "Dividends Paid");
  if (!dividendRow) return;

  periods.forEach((period, index) => {
    if (!isFourthQuarterPeriod(period)) return;
    const year = periodYearSuffix(period);
    const quarterIndexes = [`1Q${year}`, `2Q${year}`, `3Q${year}`, `4Q${year}`].map((quarter) => periods.indexOf(quarter));
    if (!quarterIndexes.every((quarterIndex) => quarterIndex >= 0)) return;
    const quarterCols = quarterIndexes.map((quarterIndex) => columns[quarterIndex]);
    const annualCol = columns[index] + 1;
    const dividendAnnual = sumNumericCells(sheet, dividendRow, quarterCols);
    if (dividendAnnual !== null) setFormulaResult(sheet.getCell(dividendRow, annualCol), dividendAnnual);

    if (!dividendsPaidRow) return;
    for (const col of [...quarterCols, annualCol]) {
      const dividendValue = numericCellValue(sheet.getCell(dividendRow, col));
      if (dividendValue !== null) setFormulaResult(sheet.getCell(dividendsPaidRow, col), -dividendValue);
    }
    const paidAnnual = sumNumericCells(sheet, dividendsPaidRow, quarterCols);
    if (paidAnnual !== null) setFormulaResult(sheet.getCell(dividendsPaidRow, annualCol), paidAnnual);
  });
}

function sumNumericCells(sheet: ExcelJS.Worksheet, rowNumber: number, columns: number[]) {
  const values = columns.map((col) => numericCellValue(sheet.getCell(rowNumber, col)));
  if (!values.every((value): value is number => value !== null)) return null;
  return values.reduce((total, value) => total + value, 0);
}

function setFormulaResult(cell: ExcelJS.Cell, result: number) {
  const value = cell.value;
  if (!value || typeof value !== "object") return;
  if (!("formula" in value) && !("sharedFormula" in value)) return;
  cell.value = { ...value, result: persistedFormulaResult(result) };
}

function persistedFormulaResult(result: number) {
  return result === 0 ? 1e-12 : result;
}

function isInactiveHelperCell(fillRow: FillRow, cell: ExcelJS.Cell) {
  const context = fillRow.modelContext;
  if (!context || cell.value !== null) return false;
  const haystack = [context.sectionHeader, context.previousLabel, context.nextLabel, context.label].filter(Boolean).join(" ");
  if (/cash flow statement/i.test(haystack)) return true;
  if (/debt and interest schedule/i.test(haystack) && fillRow.statement === "support") return true;
  if (/drivers|assumptions/i.test(context.sectionHeader ?? "") && !isBlue(cell)) return true;
  return false;
}

function inferSignConvention(label: string): 1 | -1 {
  return /(expense|cost|loss|repayment|repurchase|dividend|tax|depreciation|amortization)/i.test(label) ? -1 : 1;
}

function statementFromContext(context: ModelRowContext): FillRow["statement"] {
  if (inBalanceSheetContext(context)) return "balance";
  if (/cash|capex|debt|equity|share|dividend|working capital/i.test([context.sectionHeader, context.label].filter(Boolean).join(" "))) return "support";
  return "income";
}

function inSection(context: ModelRowContext, ...aliases: string[]) {
  const haystack = normalize([context.sectionHeader, context.previousLabel, context.nextLabel, context.label].filter(Boolean).join(" "));
  return aliases.some((alias) => haystack.includes(normalize(alias)));
}

function inShareRepurchaseAssumptionsContext(context: ModelRowContext) {
  const haystack = normalize([context.sectionHeader, context.previousLabel, context.nextLabel, context.label].filter(Boolean).join(" "));
  const hasRepurchaseRow = /sharesrepurchased|sharerepurchases|averagepricepaid|ttmpe/.test(haystack);
  const hasScheduleContext = /sharerepurchaseassumptions|shareholdersequityschedule|shareholder/.test(haystack);
  return hasScheduleContext && hasRepurchaseRow;
}

function inPpeDepreciationScheduleContext(context: ModelRowContext) {
  const haystack = normalize([context.sectionHeader, context.previousLabel, context.nextLabel, context.label].filter(Boolean).join(" "));
  const hasPpeSchedule = ["PP&E / Depreciation Schedule", "PPE / Depreciation Schedule", "PP&E", "PPE"].some((alias) => haystack.includes(normalize(alias)));
  const hasScheduleRows = /beginningppe|endingppe|depreciationexpense|acquisitiondivestmentofbusinesses|capex|capitalexpenditures/.test(haystack);
  const isPpeHeader = normalize(context.sectionHeader ?? "") === normalize("PP&E / Depreciation Schedule") || normalize(context.sectionHeader ?? "") === normalize("PPE / Depreciation Schedule");
  return isPpeHeader || (hasPpeSchedule && hasScheduleRows);
}

function inBalanceSheetContext(context: ModelRowContext) {
  return /balance sheet|assets|liabilities|equity|cash and cash equivalents|receivable|inventory|goodwill|debt/i.test(
    [context.sectionHeader, context.previousLabel, context.nextLabel, context.label].filter(Boolean).join(" ")
  );
}

function derivedSource(resolved: ResolvedValue) {
  return resolved.sources.find((source) => source.derivedTotalValue !== undefined && source.derivedPriorPeriods?.length);
}

function columnLetter(col: number) {
  let value = col;
  let letters = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    letters = String.fromCharCode(65 + remainder) + letters;
    value = Math.floor((value - 1) / 26);
  }
  return letters;
}

function fillSegmentAnalysis(
  sheet: ExcelJS.Worksheet,
  company: CompanyMatch,
  periods: string[],
  columns: number[],
  segments: SegmentRevenue[],
  ctx: ResolveContext,
  auditRows: MappingAuditRow[]
) {
  const warnings: string[] = [];
  let filledCells = 0;
  let commentsAdded = 0;
  const revenueConcepts = C.revenue;
  const fallbackRevenue = periods.map((period) => first(period, ctx.duration, revenueConcepts)?.value ?? 0);
  const rows = segmentMetricRows(sheet, "Total Company Revenue", "Revenue Mix", columns).slice(0, 6);
  const operatingIncomeRows = segmentMetricRows(sheet, "Total Company Operating Income", "Operating Income Check", columns).slice(0, 6);
  const depreciationRows = segmentMetricRows(sheet, "Total D&A", "D&A Check", columns).slice(0, 6);

  const usableSegments = segments.length
    ? segments
    : [
        {
          label: company.title.replace(/,?\s+INC\.?$/i, ""),
          values: new Map(periods.map((period, index) => [period, fallbackRevenue[index]])),
          operatingIncome: new Map(periods.map((period) => [period, 0])),
          depreciationAmortization: new Map(periods.map((period) => [period, 0]))
        }
      ];

  if (!segments.length) {
    warnings.push("No service-line revenue facts found in recent inline XBRL filings; Segment Analysis uses one total company revenue line.");
  }

  const revenueResult = fillSegmentMetricRows(sheet, periods, columns, rows, usableSegments, "values", "Revenue", auditRows);
  const operatingIncomeResult = fillSegmentMetricRows(sheet, periods, columns, operatingIncomeRows, usableSegments, "operatingIncome", "Operating Income", auditRows);
  const depreciationResult = fillSegmentMetricRows(sheet, periods, columns, depreciationRows, usableSegments, "depreciationAmortization", "D&A", auditRows);
  filledCells += revenueResult.filledCells + operatingIncomeResult.filledCells + depreciationResult.filledCells;
  commentsAdded += revenueResult.commentsAdded + operatingIncomeResult.commentsAdded + depreciationResult.commentsAdded;
  const revenueReconciliation = reconcileSegmentMetricRowsToStatementTotal(sheet, periods, columns, rows, "Revenue", revenueConcepts, ctx, auditRows);
  const operatingIncomeReconciliation = reconcileSegmentMetricRowsToStatementTotal(
    sheet,
    periods,
    columns,
    operatingIncomeRows,
    "Operating Income",
    C.operatingIncome,
    ctx,
    auditRows
  );
  filledCells += revenueReconciliation.filledCells + operatingIncomeReconciliation.filledCells;
  commentsAdded += revenueReconciliation.commentsAdded + operatingIncomeReconciliation.commentsAdded;
  warnings.push(...revenueReconciliation.warnings, ...operatingIncomeReconciliation.warnings);
  filledCells += fillSegmentTotalRow(sheet, periods, columns, usableSegments, "values", "Total Company Revenue", auditRows, ctx, C.revenue);
  filledCells += fillSegmentTotalRow(sheet, periods, columns, usableSegments, "operatingIncome", "Total Company Operating Income", auditRows, ctx, C.operatingIncome);
  filledCells += fillSegmentTotalRow(sheet, periods, columns, usableSegments, "depreciationAmortization", "Total D&A", auditRows);

  return { filledCells, commentsAdded, warnings };
}

function fillSegmentTotalRow(
  sheet: ExcelJS.Worksheet,
  periods: string[],
  columns: number[],
  segments: SegmentRevenue[],
  metric: "values" | "operatingIncome" | "depreciationAmortization",
  label: string,
  auditRows: MappingAuditRow[],
  ctx?: ResolveContext,
  statementConcepts: string[] = []
) {
  const rowNumber = findLabelRow(sheet, label);
  if (!rowNumber) return 0;
  let filledCells = 0;

  periods.forEach((period, periodIndex) => {
    const cell = sheet.getCell(rowNumber, columns[periodIndex]);
    if (!isHardcodedFinancialInput(cell)) return;
    const col = columns[periodIndex];
    const statementSource = ctx && statementConcepts.length ? first(period, ctx.duration, statementConcepts) : null;
    const value =
      statementSource?.value !== undefined
        ? statementSource.value / 1_000_000
        : segmentTotalFromModelRows(sheet, rowNumber, col, label) ?? segments.reduce((sum, segment) => sum + (segment[metric].get(period) ?? 0), 0) / 1_000_000;
    cell.value = value;
    filledCells += 1;
    auditRows.push({
      sheetName: sheet.name,
      cell: cell.address,
      modelRowLabel: label,
      period,
      valueWritten: value,
      mappingType: "calculated",
      conceptsUsed: statementSource?.concept ?? `Sum of EDGAR reportable segment ${metric} rows`,
      sourceStatement: statementSource ? "income" : "segment",
      accession: statementSource?.accn ?? "",
      sourceUrl: "",
      cellWritable: true,
      formulaPreserved: false,
      writeBlockedReason: "",
      signConvention: "copied",
      confidence: "high",
      validationStatus: "not_run",
      notes: statementSource
        ? `Rebuilt ${label} from consolidated EDGAR ${label.toLowerCase()} because the Segment Analysis total cell did not contain a formula.`
        : `Rebuilt ${label} because the Segment Analysis total cell did not contain a formula.`
    });
  });

  return filledCells;
}

function reconcileSegmentMetricRowsToStatementTotal(
  sheet: ExcelJS.Worksheet,
  periods: string[],
  columns: number[],
  rows: number[],
  suffix: string,
  concepts: string[],
  ctx: ResolveContext,
  auditRows: MappingAuditRow[]
) {
  let filledCells = 0;
  let commentsAdded = 0;
  const warnings: string[] = [];
  if (!rows.length) return { filledCells, commentsAdded, warnings };

  periods.forEach((period, periodIndex) => {
    const expectedSource = first(period, ctx.duration, concepts);
    if (!expectedSource) return;
    const col = columns[periodIndex];
    const expected = expectedSource.value / 1_000_000;
    const evaluator = new FormulaEvaluator(sheet);
    const actual = segmentMetricRowsTotal(sheet, rows, col, evaluator);
    if (segmentStatementMetricTies(actual, expected)) return;

    const residualRow = findSegmentResidualRow(sheet, rows, col, suffix);
    if (!residualRow) {
      warnings.push(`Segment Analysis ${period}: ${suffix} rows do not tie to EDGAR total and no residual segment row was available for reconciliation.`);
      return;
    }

    const cell = sheet.getCell(residualRow, col);
    const value = (numericCellValue(cell) ?? 0) + expected - actual;
    cell.value = value;
    filledCells += 1;
    const note = `Calculated as the residual needed for Segment Analysis ${suffix} rows to tie to EDGAR consolidated ${suffix.toLowerCase()}.`;
    addComment(cell, note);
    commentsAdded += 1;
    auditRows.push({
      sheetName: sheet.name,
      cell: cell.address,
      modelRowLabel: rowLabel(sheet, residualRow),
      period,
      valueWritten: value,
      mappingType: "calculated",
      conceptsUsed: concepts.join(", "),
      sourceStatement: "segment",
      accession: expectedSource.accn ?? "",
      sourceUrl: "",
      cellWritable: true,
      formulaPreserved: false,
      writeBlockedReason: "",
      signConvention: "residual",
      confidence: "medium",
      validationStatus: "OK!",
      notes: note
    });
  });

  return { filledCells, commentsAdded, warnings };
}

function segmentMetricRowsTotal(sheet: ExcelJS.Worksheet, rows: number[], col: number, evaluator?: FormulaEvaluator) {
  return rows.reduce((total, rowNumber) => {
    const cell = sheet.getCell(rowNumber, col);
    return total + (evaluator?.evaluateCell(cell) ?? numericCellValue(cell) ?? 0);
  }, 0);
}

function findSegmentResidualRow(sheet: ExcelJS.Worksheet, rows: number[], col: number, suffix: string) {
  const writableRows = rows.filter((rowNumber) => isSegmentMetricInputCell(sheet.getCell(rowNumber, col)));
  const nonGenericRows = writableRows.filter((rowNumber) => !isGenericSegmentPlaceholder(segmentBaseLabel(segmentRowLabel(sheet, rowNumber, suffix), suffix)));
  const preferred = nonGenericRows.find((rowNumber) =>
    /other|corporate|unallocated|elimination|reconciliation|residual/i.test(segmentBaseLabel(segmentRowLabel(sheet, rowNumber, suffix), suffix))
  );
  return preferred ?? nonGenericRows.at(-1) ?? null;
}

function segmentStatementMetricTies(actual: number, expected: number) {
  return Math.abs(actual - expected) <= 0.05;
}

function segmentTotalFromModelRows(sheet: ExcelJS.Worksheet, totalRow: number, col: number, totalLabel: string) {
  const endLabel = segmentTotalEndLabel(totalLabel);
  if (!endLabel) return null;
  const endRow = findLabelRow(sheet, endLabel);
  if (!endRow || endRow <= totalRow + 1) return null;

  let total = 0;
  let found = false;
  for (let row = totalRow + 1; row < endRow; row += 1) {
    const value = numericCellValue(sheet.getCell(row, col));
    if (value === null) continue;
    total += value;
    found = true;
  }
  return found ? total : null;
}

function segmentTotalEndLabel(totalLabel: string) {
  const normalized = normalize(totalLabel);
  if (normalized === normalize("Total Company Revenue")) return "Revenue Mix";
  if (normalized === normalize("Total Company Operating Income")) return "Operating Income Check";
  if (normalized === normalize("Total D&A")) return "D&A Check";
  return null;
}

function restoreSegmentCheckFormula(
  sheet: ExcelJS.Worksheet,
  periods: string[],
  columns: number[],
  totalLabel: string,
  checkLabel: string,
  auditRows: MappingAuditRow[]
) {
  const totalRow = findLabelRow(sheet, totalLabel);
  const checkRow = findLabelRow(sheet, checkLabel);
  if (!totalRow || !checkRow || checkRow <= totalRow + 1) return 0;
  let filledCells = 0;

  periods.forEach((period, periodIndex) => {
    const col = columns[periodIndex];
    const cell = sheet.getCell(checkRow, col);
    if (hasFormula(cell)) return;
    const formula = `${columnLetter(col)}${totalRow}-SUM(${columnLetter(col)}${totalRow + 1}:${columnLetter(col)}${checkRow - 1})`;
    cell.value = { formula, result: 0 };
    filledCells += 1;
    auditRows.push({
      sheetName: sheet.name,
      cell: cell.address,
      modelRowLabel: checkLabel,
      period,
      valueWritten: 0,
      mappingType: "calculated",
      conceptsUsed: formula,
      sourceStatement: "segment",
      accession: "",
      sourceUrl: "",
      cellWritable: true,
      formulaPreserved: false,
      writeBlockedReason: "",
      signConvention: "calculated",
      confidence: "high",
      validationStatus: "OK!",
      notes: `Rebuilt ${checkLabel} formula because the Segment Analysis check cell was blank.`
    });
  });

  return filledCells;
}

function fillSegmentMetricRows(
  sheet: ExcelJS.Worksheet,
  periods: string[],
  columns: number[],
  rows: number[],
  segments: SegmentRevenue[],
  metric: "values" | "operatingIncome" | "depreciationAmortization",
  suffix: string,
  auditRows: MappingAuditRow[]
) {
  let filledCells = 0;
  let commentsAdded = 0;
  const usedSegments = new Set<number>();

  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    const rowNumber = rows[rowIndex];
    const existingLabel = segmentBaseLabel(segmentRowLabel(sheet, rowNumber, suffix), suffix);
    if (isGenericSegmentPlaceholder(existingLabel)) {
      const cleared = clearUnmatchedSegmentRow(sheet, rowNumber, periods, columns, existingLabel || suffix, auditRows);
      filledCells += cleared.filledCells;
      commentsAdded += cleared.commentsAdded;
      continue;
    }
    const segmentIndex = segmentIndexForRow(existingLabel, segments, usedSegments);
    if (segmentIndex === null) {
      const cleared = clearUnmatchedSegmentRow(sheet, rowNumber, periods, columns, existingLabel, auditRows);
      filledCells += cleared.filledCells;
      commentsAdded += cleared.commentsAdded;
      continue;
    }

    const segment = segments[segmentIndex];
    if (!segmentHasMetricData(segment, metric, periods)) continue;
    usedSegments.add(segmentIndex);

    periods.forEach((period, periodIndex) => {
      const cell = sheet.getCell(rowNumber, columns[periodIndex]);
      if (!isSegmentMetricInputCell(cell)) return;
      const value = (segment[metric].get(period) ?? 0) / 1_000_000;
      cell.value = value;
      filledCells += 1;
      const resolved = segmentResolvedValue(segment, metric, period);
      const comment = mappingCommentForSegment(sheet, cell, existingLabel, segment, period, value, suffix);
      addComment(cell, comment);
      commentsAdded += 1;
      auditRows.push(mappingAuditRowForSegment(sheet, cell, existingLabel, period, value, resolved, suffix));
    });
  }

  return { filledCells, commentsAdded };
}

function clearUnmatchedSegmentRow(
  sheet: ExcelJS.Worksheet,
  rowNumber: number,
  periods: string[],
  columns: number[],
  modelLabel: string,
  auditRows: MappingAuditRow[]
) {
  let filledCells = 0;
  let commentsAdded = 0;
  periods.forEach((period, periodIndex) => {
    const cell = sheet.getCell(rowNumber, columns[periodIndex]);
    if (!isSegmentMetricInputCell(cell)) return;
    const existing = numericCellValue(cell);
    if (existing === 0 || existing === null) return;
    cell.value = 0;
    filledCells += 1;
    const notes = "Needs review: Segment Analysis row was not filled because the template label is blank, generic, or does not confidently match an EDGAR reportable segment.";
    addComment(cell, notes);
    commentsAdded += 1;
    auditRows.push({
      sheetName: sheet.name,
      cell: cell.address,
      modelRowLabel: modelLabel,
      period,
      valueWritten: 0,
      mappingType: "unused",
      conceptsUsed: "",
      sourceStatement: "segment",
      accession: "",
      sourceUrl: "",
      cellWritable: true,
      formulaPreserved: false,
      writeBlockedReason: "",
      signConvention: "not written",
      confidence: "low",
      validationStatus: "needs_review",
      notes
    });
  });
  return { filledCells, commentsAdded };
}

function segmentHasMetricData(segment: SegmentRevenue, metric: "values" | "operatingIncome" | "depreciationAmortization", periods: string[]) {
  return periods.some((period) => segment[metric].has(period) && segment[metric].get(period) !== 0);
}

function segmentRowLabel(sheet: ExcelJS.Worksheet, rowNumber: number, suffix: string) {
  const labelCellForRow = labelCell(sheet, rowNumber);
  const displayed = cellDisplay(labelCellForRow).trim();
  if (displayed) return displayed;

  const formula = cellFormula(labelCellForRow);
  if (!formula) return "";

  const reference = formula.match(/^=?\$?([A-Z]+)\$?(\d+)\s*(?:&|$)/i);
  if (!reference) return "";

  const referenced = cellDisplay(sheet.getCell(`${reference[1]}${reference[2]}`)).trim();
  if (!referenced) return "";

  if (/Revenue|Operating Income|D&A/i.test(formula)) return `${referenced} ${suffix}`;
  return referenced;
}

function segmentMetricRows(sheet: ExcelJS.Worksheet, startLabel: string, endLabel: string, columns: number[]) {
  const startRow = findLabelRow(sheet, startLabel);
  const endRow = findLabelRow(sheet, endLabel);
  if (!startRow || !endRow || endRow <= startRow) return [];
  const rows: number[] = [];
  for (let rowNumber = startRow + 1; rowNumber < endRow; rowNumber += 1) {
    if (rowLabel(sheet, rowNumber) && rowHasSegmentMetricInputs(sheet, rowNumber, columns)) rows.push(rowNumber);
  }
  return rows;
}

function rowHasSegmentMetricInputs(sheet: ExcelJS.Worksheet, rowNumber: number, columns: number[]) {
  return columns.some((col) => isSegmentMetricInputCell(sheet.getCell(rowNumber, col)));
}

function isSegmentMetricInputCell(cell: ExcelJS.Cell) {
  return isHardcodedFinancialInput(cell);
}

function segmentMixLabelRows(sheet: ExcelJS.Worksheet) {
  const revenueMixRow = findLabelRow(sheet, "Revenue Mix");
  const operatingIncomeRow = findLabelRow(sheet, "Total Company Operating Income");
  if (!revenueMixRow || !operatingIncomeRow || operatingIncomeRow <= revenueMixRow) return [16, 17, 18, 19, 20, 21];
  const rows: number[] = [];
  for (let rowNumber = revenueMixRow + 1; rowNumber < operatingIncomeRow; rowNumber += 1) {
    if (cellDisplay(sheet.getCell(rowNumber, 3)).trim()) rows.push(rowNumber);
  }
  return rows.length ? rows : [16, 17, 18, 19, 20, 21];
}

function segmentBaseLabel(label: string, suffix: string) {
  return label
    .replace(/^"+|"+$/g, "")
    .replace(/^=/, "")
    .replace(new RegExp(`\\s*${suffix.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}\\s*$`, "i"), "")
    .replace(/\s*Revenue\s*$/i, "")
    .replace(/\s*Operating Income\s*$/i, "")
    .replace(/\s*D&A\s*$/i, "")
    .trim();
}

function isGenericSegmentPlaceholder(label: string) {
  return !label || /^segment\s*\d+$/i.test(label) || /^(revenue|operating income|d&a|total company)$/i.test(label.trim());
}

function segmentIndexForRow(label: string, segments: SegmentRevenue[], used: Set<number>) {
  if (!label) return null;
  const normalizedLabel = normalize(label);
  let bestIndex: number | null = null;
  let bestScore = 0;
  segments.forEach((segment, segmentIndex) => {
    if (used.has(segmentIndex)) return;
    const normalizedSegment = normalize(segment.label);
    const score = segmentLabelMatchScore(normalizedLabel, normalizedSegment);
    if (!score) return;
    if (score > bestScore) {
      bestIndex = segmentIndex;
      bestScore = score;
    }
  });
  return bestIndex;
}

function segmentLabelMatchScore(normalizedLabel: string, normalizedSegment: string) {
  if (!normalizedLabel || !normalizedSegment) return 0;
  if (normalizedSegment === normalizedLabel) return 100;
  if (normalizedLabel.includes(normalizedSegment)) {
    const coverage = normalizedSegment.length / normalizedLabel.length;
    return coverage >= 0.75 ? 70 + coverage : 0;
  }
  if (normalizedSegment.includes(normalizedLabel)) {
    const coverage = normalizedLabel.length / normalizedSegment.length;
    return coverage >= 0.75 ? 60 + coverage : 0;
  }
  return 0;
}

function nextUnusedSegmentIndex(segments: SegmentRevenue[], used: Set<number>) {
  const index = segments.findIndex((_, segmentIndex) => !used.has(segmentIndex));
  return index === -1 ? null : index;
}

function segmentResolvedValue(
  segment: SegmentRevenue,
  metric: "values" | "operatingIncome" | "depreciationAmortization",
  period: string
): ResolvedValue {
  const value = segment[metric].get(period) ?? null;
  return {
    value,
    sources: value === null ? [] : [{ concept: `segment:${metric}`, label: segment.label, value }],
    note: `Matched reportable segment "${segment.label}" from filing segment tables.`,
    classification: "direct"
  };
}

function findLabelRow(sheet: ExcelJS.Worksheet, label: string) {
  const wanted = normalize(label);
  for (let rowNumber = 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
    if (normalize(rowLabel(sheet, rowNumber)) === wanted) return rowNumber;
  }
  return null;
}

function labelCell(sheet: ExcelJS.Worksheet, rowNumber: number) {
  let best: ExcelJS.Cell | null = null;
  let bestScore = -Infinity;
  for (const col of LABEL_COLUMNS) {
    const cell = sheet.getCell(rowNumber, col);
    const text = cellDisplay(cell).trim();
    if (!text || /^x$/i.test(text)) continue;
    const score = scoreLabelCandidate(text);
    if (score > bestScore) {
      best = cell;
      bestScore = score;
    }
  }
  return best ?? sheet.getCell(rowNumber, 3);
}

function rowHasHardcodedFinancialInputs(sheet: ExcelJS.Worksheet, rowNumber: number, columns: number[]) {
  for (const col of columns) {
    if (isHardcodedFinancialInput(sheet.getCell(rowNumber, col))) return true;
  }
  return false;
}

function isBlue(cell: ExcelJS.Cell) {
  const color = cell.font?.color;
  if (!color) return false;
  if (color.argb && BLUE_FONT_COLORS.has(color.argb)) return true;
  return color.theme === 10;
}

function isHardcodedBlueInput(cell: ExcelJS.Cell) {
  return isBlue(cell) && !hasFormula(cell);
}

function isHardcodedInput(cell: ExcelJS.Cell) {
  return !hasFormula(cell);
}

function isHardcodedFinancialInput(cell: ExcelJS.Cell) {
  if (hasFormula(cell)) return false;
  const value = cell.value;
  return value === null || typeof value === "number";
}

function isModelHistoricalInput(cell: ExcelJS.Cell) {
  if (hasFormula(cell)) return false;
  return isBlue(cell) || typeof cell.value === "number";
}

function canAddComment(cell: ExcelJS.Cell) {
  return !hasFormula(cell);
}

function hasFormula(cell: ExcelJS.Cell) {
  const value = cell.value;
  return Boolean(value && typeof value === "object" && ("formula" in value || "sharedFormula" in value));
}

function isNumericConstantFormula(formula: string | null) {
  return Boolean(formula?.replace(/^=/, "").trim().match(/^[+-]?\d+(?:\.\d+)?$/));
}

function auditNoteForResolvedValue(fillRow: FillRow, resolved: ResolvedValue) {
  const derived = derivedSource(resolved);
  const sourceLabels = unique(resolved.sources.map(sourceDisplayLabel).filter(Boolean));
  const classification = resolved.classification ?? fillRow.classification;
  const grouped = classification === "grouped" || sourceLabels.length > 1 || Boolean((resolved.note || fillRow.comment) && fillRow.classification === "grouped");
  const isDerived = Boolean(derived?.derivedTotalValue !== undefined);
  const hasAnalystNote = Boolean((resolved.note || fillRow.comment) && (grouped || isDerived));
  if (!grouped && !hasAnalystNote && !isDerived) return null;

  const parts: string[] = [];
  if (grouped && sourceLabels.length) parts.push(`Grouped from EDGAR: ${sourceLabels.join(", ")}.`);
  if (resolved.note || fillRow.comment) parts.push(resolved.note || fillRow.comment || "");
  if (isDerived && derived?.derivedPriorPeriods?.length) {
    parts.push(`Quarterly value was derived from ${derived.derivedTotalLabel || derived.label} less ${derived.derivedPriorPeriods.join(", ")}.`);
  }
  return parts.filter(Boolean).join(" ");
}

function sourceDisplayLabel(source: FactSource) {
  const label = source.note || source.label || source.concept;
  if (!label) return "";
  if (label !== source.concept && /\s/.test(label)) return label;
  return label
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

function resolveRow(fillRow: FillRow, period: string, ctx: ResolveContext): ResolvedValue {
  if (fillRow.resolver) return fillRow.resolver(period, ctx);
  if (!fillRow.concepts?.length) return { value: null, sources: [], note: fillRow.noFillComment, classification: fillRow.classification };
  const source = first(period, fillRow.kind === "instant" ? ctx.instant : ctx.duration, fillRow.concepts);
  if (!source) return { value: null, sources: [] };
  const resolved = signed(source, fillRow.sign ?? 1) ?? { value: null, sources: [] };
  return { ...resolved, classification: fillRow.classification };
}

function reconcileIncomeStatementFormulaRowsToEdgar(
  sheet: ExcelJS.Worksheet,
  periods: string[],
  columns: number[],
  ctx: ResolveContext,
  auditRows: MappingAuditRow[]
) {
  let filledCells = 0;
  let commentsAdded = 0;
  const warnings: string[] = [];
  const ebitRow = findIncomeStatementMetricRow(sheet, ["EBIT"]);
  const residualRow = findIncomeStatementResidualRowBefore(sheet, ebitRow, [
    "Other Operating Income (Expense)",
    "Other Operating Expenses",
    "Other Operating Expense",
    "Other Operating Income",
    "Other Operating Income Expense"
  ]);
  if (!ebitRow || !residualRow) return { filledCells, commentsAdded, warnings };

  const evaluator = new FormulaEvaluator(sheet);
  periods.forEach((period, index) => {
    const source = first(period, ctx.duration, C.operatingIncome);
    if (!source) return;
    const col = columns[index];
    const ebitCell = sheet.getCell(ebitRow, col);
    if (!hasFormula(ebitCell)) return;
    const current = evaluator.evaluateCell(ebitCell);
    if (current === null) return;
    const expected = source.value / 1_000_000;
    if (incomeStatementFormulaTies(current, expected)) return;

    const residualCell = sheet.getCell(residualRow, col);
    const value = (numericCellValue(residualCell) ?? 0) + expected - current;
    if (!isReconciliationResidualCellWritable(residualCell)) {
      warnings.push(`Income Statement ${period}: EBIT formula did not tie to EDGAR and no writable operating residual row was available.`);
      return;
    }

    residualCell.value = value;
    evaluator.clear();
    filledCells += 1;
    const note = "Calculated as the residual needed for the Income Statement EBIT formula to tie to EDGAR pre-tax operating income.";
    addComment(residualCell, note);
    commentsAdded += 1;
    auditRows.push({
      sheetName: sheet.name,
      cell: residualCell.address,
      modelRowLabel: rowLabel(sheet, residualRow),
      period,
      valueWritten: value,
      mappingType: "calculated",
      conceptsUsed: C.operatingIncome.join(", "),
      sourceStatement: "income",
      accession: source.accn ?? "",
      sourceUrl: "",
      cellWritable: true,
      formulaPreserved: false,
      writeBlockedReason: "",
      signConvention: "residual",
      confidence: "medium",
      validationStatus: "OK!",
      notes: note
    });
  });

  return { filledCells, commentsAdded, warnings };
}

function findIncomeStatementResidualRowBefore(sheet: ExcelJS.Worksheet, targetRow: number | null, labels: string[]) {
  if (!targetRow) return null;
  const wanted = new Set(labels.map(normalize));
  for (let rowNumber = targetRow - 1; rowNumber >= Math.max(1, targetRow - 12); rowNumber -= 1) {
    if (wanted.has(normalize(rowLabel(sheet, rowNumber)))) return rowNumber;
  }
  return null;
}

function incomeStatementFormulaTies(actual: number, expected: number) {
  return Math.abs(actual - expected) <= 0.12;
}

function reconcileNetIncomeFormulaRowsToEdgar(
  sheet: ExcelJS.Worksheet,
  periods: string[],
  columns: number[],
  ctx: ResolveContext,
  auditRows: MappingAuditRow[]
) {
  let filledCells = 0;
  let commentsAdded = 0;
  const warnings: string[] = [];
  const netIncomeRow = findIncomeStatementMetricRow(sheet, ["Net Income (Loss)", "Net Income"]);
  const residualRow = findIncomeStatementResidualRowBefore(sheet, netIncomeRow, [
    "Income Tax Benefit (Expense)",
    "Income Tax Expense",
    "Income Tax Provision (Expense)",
    "Income Tax"
  ]);
  if (!netIncomeRow || !residualRow) return { filledCells, commentsAdded, warnings };

  const evaluator = new FormulaEvaluator(sheet);
  periods.forEach((period, index) => {
    const source = first(period, ctx.duration, CONTINUING_NET_INCOME_CONCEPTS);
    if (!source) return;
    const col = columns[index];
    const netIncomeCell = sheet.getCell(netIncomeRow, col);
    const actual = evaluator.evaluateCell(netIncomeCell);
    const expected = source.value / 1_000_000;
    if (actual === null || incomeStatementFormulaTies(actual, expected)) return;

    const residualCell = sheet.getCell(residualRow, col);
    const value = (numericCellValue(residualCell) ?? 0) + expected - actual;
    if (!isReconciliationResidualCellWritable(residualCell)) {
      warnings.push(`Income Statement ${period}: net income does not tie to EDGAR and ${rowLabel(sheet, residualRow)} was not writable for reconciliation.`);
      return;
    }

    residualCell.value = value;
    evaluator.clear();
    filledCells += 1;
    const note = "Calculated as the residual needed for the model's net income formula to tie to EDGAR net income.";
    addComment(residualCell, note);
    commentsAdded += 1;
    auditRows.push({
      sheetName: sheet.name,
      cell: residualCell.address,
      modelRowLabel: rowLabel(sheet, residualRow),
      period,
      valueWritten: value,
      mappingType: "calculated",
      conceptsUsed: CONTINUING_NET_INCOME_CONCEPTS.join(", "),
      sourceStatement: "income",
      accession: source.accn ?? "",
      sourceUrl: "",
      cellWritable: true,
      formulaPreserved: false,
      writeBlockedReason: "",
      signConvention: "residual",
      confidence: "medium",
      validationStatus: "OK!",
      notes: note
    });
  });

  return { filledCells, commentsAdded, warnings };
}

function formulaBridgeForTargetValue(cell: ExcelJS.Cell, targetValue: number) {
  const formula = formulaForCell(cell);
  if (!formula) return null;
  const match = normalizeBridgeFormulaPrefix(formula).match(/^([\d+\-*/().\s]+)-SUM\(([A-Z]+)(\d+):([A-Z]+)(\d+)\)$/i);
  if (!match) return null;
  const startCol = columnIndex(match[2]);
  const startRow = Number(match[3]);
  const endCol = columnIndex(match[4]);
  const endRow = Number(match[5]);
  if (startRow !== endRow || startRow !== Number(cell.address.match(/\d+$/)?.[0])) return null;

  let priorSum = 0;
  for (let col = Math.min(startCol, endCol); col <= Math.max(startCol, endCol); col += 1) {
    priorSum += numericCellValue(cell.worksheet.getCell(startRow, col)) ?? 0;
  }
  const bridgeTotal = targetValue + priorSum;
  return `${formatDividendFormulaNumber(bridgeTotal)}-SUM(${columnLetter(Math.min(startCol, endCol))}${startRow}:${columnLetter(Math.max(startCol, endCol))}${endRow})`;
}

function formatDividendFormulaNumber(value: number) {
  if (Number.isInteger(value)) return String(value);
  return value.toFixed(3).replace(/0+$/, "").replace(/\.$/, "");
}

function reconcileBalanceSheetCheck(sheet: ExcelJS.Worksheet, periods: string[], columns: number[], auditRows: MappingAuditRow[]) {
  let filledCells = 0;
  let commentsAdded = 0;
  const warnings: string[] = [];
  const checkRow = findRowInSection(sheet, "Balance Sheet", ["Balance Sheet Check"], (label) =>
    /working capital|cash flow statement|cashflow statement|income statement|schedule|analysis|drivers/i.test(label)
  );
  if (!checkRow) return { filledCells, commentsAdded, warnings };

  const assetResidualRows = findBalanceSheetResidualRows(sheet, checkRow, [
    "Other Non-Current Assets",
    "Other Long-Term Assets",
    "Other LT Assets",
    "Prepaid & Other Current Assets",
    "Prepaid and Other Current Assets",
    "Other Current Assets",
    "Other Assets and Loans"
  ]);
  const liabilityResidualRows = findBalanceSheetResidualRows(sheet, checkRow, [
    "Other Non-Current Liabilities",
    "Other Current Liabilities",
    "Other Liabilities",
    "Accounts Payable and Accrued Liabilities",
    "Accounts Payable & Accrued Liabilities",
    "Accrued Liabilities"
  ]);

  const evaluator = new FormulaEvaluator(sheet);
  periods.forEach((period, index) => {
    const col = columns[index];
    const checkCell = sheet.getCell(checkRow, col);
    const check = evaluator.evaluateCell(checkCell);
    if (check === null || valuesTie(check, 0)) return;

    const residualRows = check < 0 ? assetResidualRows : liabilityResidualRows;
    const residualRow = findWritableBalanceSheetResidualRow(sheet, residualRows, col);
    if (!residualRow) {
      warnings.push(`Balance Sheet ${period}: check is ${roundModelValue(check)} and no writable residual balance-sheet row was available.`);
      return;
    }
    const residualCell = sheet.getCell(residualRow, col);

    const value = (numericCellValue(residualCell) ?? 0) + Math.abs(check);
    residualCell.value = value;
    evaluator.clear();
    filledCells += 1;
    const note = "Calculated as the residual needed for the Balance Sheet Check to tie to zero using EDGAR-sourced mapped balances.";
    addComment(residualCell, note);
    commentsAdded += 1;
    auditRows.push({
      sheetName: sheet.name,
      cell: residualCell.address,
      modelRowLabel: rowLabel(sheet, residualRow),
      period,
      valueWritten: value,
      mappingType: "calculated",
      conceptsUsed: "Balance Sheet Check residual",
      sourceStatement: "balance",
      accession: "",
      sourceUrl: "",
      cellWritable: true,
      formulaPreserved: false,
      writeBlockedReason: "",
      signConvention: "residual",
      confidence: "medium",
      validationStatus: "OK!",
      notes: note
    });
  });

  return { filledCells, commentsAdded, warnings };
}

function reconcileBalanceSheetStatementTotalsToEdgar(
  sheet: ExcelJS.Worksheet,
  periods: string[],
  columns: number[],
  ctx: ResolveContext,
  auditRows: MappingAuditRow[]
) {
  let filledCells = 0;
  let commentsAdded = 0;
  const warnings: string[] = [];
  const evaluator = new FormulaEvaluator(sheet);

  const metrics = [
    {
      name: "total assets",
      labels: ["Total Assets"],
      concepts: C.assets,
      residualLabels: [
        "Other Non-Current Assets",
        "Other Long-Term Assets",
        "Other LT Assets",
        "Prepaid & Other Current Assets",
        "Prepaid and Other Current Assets",
        "Other Current Assets",
        "Other Assets and Loans"
      ]
    },
    {
      name: "total liabilities",
      labels: ["Total Liabilities"],
      concepts: C.liabilities,
      residualLabels: [
        "Other Non-Current Liabilities",
        "Other Current Liabilities",
        "Other Liabilities",
        "Accounts Payable and Accrued Liabilities",
        "Accounts Payable & Accrued Liabilities",
        "Accrued Liabilities"
      ]
    },
    {
      name: "shareholders' equity",
      labels: [
        "Total Shareholder's Equity",
        "Total Shareholders' Equity",
        "Total Shareholders Equity",
        "Total Stockholders' Equity",
        "Total Stockholders Equity",
        "Total Equity"
      ],
      concepts: C.equity,
      residualLabels: ["Accumulated Other Comprehensive Income (AOCI)", "AOCI", "Noncontrolling Interests", "Non-Controlling Interests", "Retained Earnings"]
    },
    {
      name: "total liabilities plus shareholders' equity",
      labels: ["Total Liabilities & Shareholder's Equity", "Total Liabilities and Shareholder's Equity"],
      concepts: C.assets,
      residualLabels: [
        "Other Non-Current Liabilities",
        "Other Current Liabilities",
        "Other Liabilities",
        "Accounts Payable and Accrued Liabilities",
        "Accounts Payable & Accrued Liabilities",
        "Accrued Liabilities"
      ]
    }
  ];

  for (const metric of metrics) {
    const totalRow = findRowInSection(sheet, "Balance Sheet", metric.labels, (label) =>
      /working capital|cash flow statement|cashflow statement|income statement|schedule|analysis|drivers/i.test(label)
    );
    if (!totalRow) continue;
    const residualRows = findBalanceSheetResidualRowsBefore(sheet, totalRow, metric.residualLabels);

    periods.forEach((period, index) => {
      const source = first(period, ctx.instant, metric.concepts);
      if (!source) return;
      const col = columns[index];
      const totalCell = sheet.getCell(totalRow, col);
      const expected = source.value / 1_000_000;
      const actual = statementMetricCellValue(totalCell, evaluator, expected);
      if (actual === null || statementMetricTies(actual, expected)) return;

      if (!hasFormula(totalCell) && isHardcodedFinancialInput(totalCell)) {
        totalCell.value = expected;
        evaluator.clear();
        filledCells += 1;
        const note = `Mapped directly to EDGAR ${metric.name} because the balance-sheet total cell was hardcoded.`;
        addComment(totalCell, note);
        commentsAdded += 1;
        auditRows.push(statementTotalAuditRow(sheet, totalCell, rowLabel(sheet, totalRow), period, expected, source, "balance", note));
        return;
      }

      const residualRow = findWritableBalanceSheetResidualRow(sheet, residualRows, col);
      if (!residualRow) {
        warnings.push(`Balance Sheet ${period}: ${metric.name} does not tie to EDGAR and no writable residual row was available.`);
        return;
      }

      const residualCell = sheet.getCell(residualRow, col);
      const value = (numericCellValue(residualCell) ?? 0) + expected - actual;
      residualCell.value = value;
      evaluator.clear();
      filledCells += 1;
      const note = `Calculated as the residual needed for balance-sheet ${metric.name} to tie to EDGAR.`;
      addComment(residualCell, note);
      commentsAdded += 1;
      auditRows.push(statementTotalAuditRow(sheet, residualCell, rowLabel(sheet, residualRow), period, value, source, "balance", note, "residual"));
    });
  }

  return { filledCells, commentsAdded, warnings };
}

function statementTotalAuditRow(
  sheet: ExcelJS.Worksheet,
  cell: ExcelJS.Cell,
  label: string,
  period: string,
  value: number,
  source: FactSource,
  sourceStatement: string,
  note: string,
  signConvention = "copied"
): MappingAuditRow {
  return {
    sheetName: sheet.name,
    cell: cell.address,
    modelRowLabel: label,
    period,
    valueWritten: value,
    mappingType: "calculated",
    conceptsUsed: source.concept,
    sourceStatement,
    accession: source.accn ?? "",
    sourceUrl: "",
    cellWritable: true,
    formulaPreserved: false,
    writeBlockedReason: "",
    signConvention,
    confidence: signConvention === "residual" ? "medium" : "high",
    validationStatus: "OK!",
    notes: note
  };
}

function findBalanceSheetResidualRows(sheet: ExcelJS.Worksheet, checkRow: number, labels: string[]) {
  const wanted = new Set(labels.map(normalize));
  const rows: number[] = [];
  for (let rowNumber = checkRow - 1; rowNumber >= Math.max(1, checkRow - 45); rowNumber -= 1) {
    if (wanted.has(normalize(rowLabel(sheet, rowNumber)))) rows.push(rowNumber);
  }
  return rows;
}

function findBalanceSheetResidualRowsBefore(sheet: ExcelJS.Worksheet, totalRow: number, labels: string[]) {
  const wanted = new Set(labels.map(normalize));
  const rows: number[] = [];
  for (let rowNumber = totalRow - 1; rowNumber >= Math.max(1, totalRow - 60); rowNumber -= 1) {
    const label = rowLabel(sheet, rowNumber);
    if (/income statement|cash flow statement|cashflow statement|working capital|schedule|analysis|drivers/i.test(label)) break;
    if (wanted.has(normalize(label))) rows.push(rowNumber);
  }
  return rows;
}

function findWritableBalanceSheetResidualRow(sheet: ExcelJS.Worksheet, rows: number[], col: number) {
  return rows.find((rowNumber) => isReconciliationResidualCellWritable(sheet.getCell(rowNumber, col))) ?? null;
}

function isReconciliationResidualCellWritable(cell: ExcelJS.Cell) {
  return isHardcodedFinancialInput(cell);
}

function createLlmMappingState(): LlmMappingState {
  const enabledByEnv = llmMappingEnabledByEnv();
  const hasApiKey = Boolean(llmApiKey());
  return {
    enabled: enabledByEnv && hasApiKey,
    decisions: new Map(),
    warnings: enabledByEnv && !hasApiKey ? ["LLM mapping was enabled but OPENROUTER_API_KEY was not set; deterministic EDGAR mapping was used."] : [],
    calls: 0,
    maxCalls: Number.isFinite(LLM_MAPPING_MAX_CALLS) && LLM_MAPPING_MAX_CALLS > 0 ? LLM_MAPPING_MAX_CALLS : 24
  };
}

function llmMappingEnabledByEnv() {
  const raw = process.env.LLM_MAPPING_ENABLED;
  if (raw === undefined || raw === "") return Boolean(llmApiKey());
  return /^(true|1|yes)$/i.test(raw);
}

function llmApiKey() {
  return process.env.OPENROUTER_API_KEY || process.env.OPENAI_API_KEY || "";
}

async function llmAssistedFillRow(
  fillRow: FillRow,
  company: CompanyMatch,
  periods: string[],
  ctx: ResolveContext,
  state: LlmMappingState
) {
  if (!state.enabled || !isLlmMappableRow(fillRow)) return null;
  if (state.decisions.has(fillRow.row)) return state.decisions.get(fillRow.row) ?? null;
  if (state.calls >= state.maxCalls) {
    state.decisions.set(fillRow.row, null);
    return null;
  }

  const candidates = llmCandidateFacts(fillRow, periods, ctx);
  if (!candidates.length) {
    state.decisions.set(fillRow.row, null);
    return null;
  }

  state.calls += 1;
  try {
    const modelChoice = chooseLlmMappingModel(fillRow, candidates);
    const decision = await requestLlmMappingDecision(company, fillRow, periods, candidates, modelChoice.model);
    const mapped = llmDecisionToFillRow(fillRow, decision, candidates, modelChoice);
    state.decisions.set(fillRow.row, mapped);
    return mapped;
  } catch (error) {
    const message = error instanceof Error ? error.message : "unknown OpenRouter API error";
    state.warnings.push(`${fillRow.label}: LLM-assisted mapping skipped (${message}).`);
    state.decisions.set(fillRow.row, null);
    return null;
  }
}

function isLlmMappableRow(fillRow: FillRow) {
  if (fillRow.statement !== "income" && fillRow.statement !== "balance") return false;
  if (fillRow.classification !== "unused") return false;
  const context = fillRow.modelContext;
  if (context && !context.hasHardcodedInput && !fillRow.allowBlankHistoricalInput) return false;
  return true;
}

function llmCandidateFacts(fillRow: FillRow, periods: string[], ctx: ResolveContext): LlmCandidateFact[] {
  const map = fillRow.kind === "instant" ? ctx.instant : ctx.duration;
  const byConcept = new Map<string, LlmCandidateFact>();

  periods.forEach((period) => {
    const facts = map.get(period);
    if (!facts) return;
    facts.forEach((source, concept) => {
      const existing = byConcept.get(concept) ?? {
        concept,
        label: sourceDisplayLabel(source),
        statement: fillRow.statement === "balance" ? "balance" : "income",
        values: {}
      };
      existing.values[period] = roundModelValue(source.value / 1_000_000);
      byConcept.set(concept, existing);
    });
  });

  return Array.from(byConcept.values())
    .map((candidate) => ({ candidate, score: llmCandidateScore(fillRow, candidate) }))
    .filter(({ score }) => score >= llmMinCandidateScore())
    .sort((a, b) => b.score - a.score || Object.keys(b.candidate.values).length - Object.keys(a.candidate.values).length)
    .slice(0, llmCandidateLimit())
    .map(({ candidate }) => candidate);
}

function llmMinCandidateScore() {
  return Number.isFinite(LLM_MAPPING_MIN_CANDIDATE_SCORE) ? LLM_MAPPING_MIN_CANDIDATE_SCORE : 2;
}

function llmCandidateLimit() {
  return Number.isFinite(LLM_MAPPING_CANDIDATE_LIMIT) && LLM_MAPPING_CANDIDATE_LIMIT > 0 ? LLM_MAPPING_CANDIDATE_LIMIT : 80;
}

function llmCandidateScore(fillRow: FillRow, candidate: LlmCandidateFact) {
  const rowTokens = significantTokens([fillRow.label, fillRow.modelContext?.sectionHeader, fillRow.modelContext?.previousLabel, fillRow.modelContext?.nextLabel].join(" "));
  const candidateTokens = significantTokens(`${candidate.concept} ${candidate.label}`);
  let score = 0;
  rowTokens.forEach((token) => {
    if (candidateTokens.has(token)) score += token.length >= 6 ? 3 : 2;
    candidateTokens.forEach((candidateToken) => {
      if (candidateToken !== token && (candidateToken.includes(token) || token.includes(candidateToken)) && Math.min(token.length, candidateToken.length) >= 5) {
        score += 1;
      }
    });
  });
  if (fillRow.statement === "balance" && /(assets?|liabilit|equity|cash|receivable|inventory|debt|payable|goodwill|deposit|tax|stock|aoci)/i.test(candidate.concept)) score += 2;
  if (fillRow.statement === "income" && /(revenue|income|expense|cost|profit|loss|tax|interest|depreciation|amortization|provision|sales)/i.test(candidate.concept)) score += 2;
  return score;
}

function chooseLlmMappingModel(fillRow: FillRow, candidates: LlmCandidateFact[]): LlmMappingModelChoice {
  const score = llmMappingComplexityScore(fillRow, candidates);
  const complexThreshold = Number.isFinite(LLM_MAPPING_COMPLEX_SCORE) ? LLM_MAPPING_COMPLEX_SCORE : 4;
  if (score >= complexThreshold && LLM_MAPPING_COMPLEX_MODEL) {
    return {
      model: LLM_MAPPING_COMPLEX_MODEL,
      tier: "complex",
      reason: `complex mapping score ${score}`
    };
  }
  return {
    model: LLM_MAPPING_FAST_MODEL,
    tier: "fast",
    reason: `simple mapping score ${score}`
  };
}

function llmMappingComplexityScore(fillRow: FillRow, candidates: LlmCandidateFact[]) {
  const label = [fillRow.label, fillRow.modelContext?.sectionHeader, fillRow.modelContext?.previousLabel, fillRow.modelContext?.nextLabel].join(" ");
  const scoredCandidates = candidates.map((candidate) => llmCandidateScore(fillRow, candidate)).sort((a, b) => b - a);
  const topScore = scoredCandidates[0] ?? 0;
  const secondScore = scoredCandidates[1] ?? 0;
  let score = 0;

  if (candidates.length > 50) score += 3;
  else if (candidates.length > 25) score += 2;
  else if (candidates.length > 12) score += 1;
  if (topScore < 6) score += 1;
  if (secondScore && topScore - secondScore <= 2) score += 1;
  if (/other|misc|adjust|reclass|non[-\s]?operating|non[-\s]?current|unallocated|elimination|residual/i.test(label)) score += 2;
  if (/interest|credit|loan|securit|broker|dealer|deposit|receivable|payable|investment|noninterest|trading|fair value/i.test(label)) score += 2;
  if (/\b(and|incl\.?|including|excluding|less|net of|total)\b/i.test(fillRow.label)) score += 1;
  if (fillRow.statement === "balance" && /liabilit|equity|asset/i.test(fillRow.label)) score += 1;

  return score;
}

function significantTokens(value: string) {
  const stop = new Set(["and", "the", "for", "from", "with", "net", "total", "current", "other", "statement", "schedule"]);
  return new Set(
    value
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, " ")
      .split(/\s+/)
      .filter((token) => token.length >= 3 && !stop.has(token))
  );
}

async function requestLlmMappingDecision(
  company: CompanyMatch,
  fillRow: FillRow,
  periods: string[],
  candidates: LlmCandidateFact[],
  model: string
): Promise<LlmMappingDecision> {
  const apiKey = llmApiKey();
  if (!apiKey) throw new Error("OPENROUTER_API_KEY is not set");

  const system = [
    "You map financial model rows to SEC EDGAR XBRL facts.",
    "Use only the provided candidate concepts. Do not invent concepts or values.",
    "Treat every mapping correction as a reusable template rule: rely on row labels, section context, statement type, period kind, and SEC concept semantics, never ticker-specific or single-period exceptions.",
    "Prefer needs_review unless the row label, section context, and candidate label clearly match.",
    "Choose sign -1 only when the model row convention should invert the EDGAR value, such as expense rows shown as negatives.",
    "Return strict JSON only."
  ].join(" ");

  const response = await fetch(OPENROUTER_CHAT_COMPLETIONS_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${apiKey}`,
      "HTTP-Referer": OPENROUTER_SITE_URL,
      "X-Title": OPENROUTER_APP_TITLE
    },
    body: JSON.stringify({
      model,
      messages: [
        { role: "system", content: system },
        { role: "user", content: JSON.stringify(llmMappingPayload(company, fillRow, periods, candidates)) }
      ],
      temperature: 0,
      max_tokens: 700,
      provider: { require_parameters: true },
      response_format: {
        type: "json_schema",
        json_schema: llmMappingJsonSchema()
      }
    })
  });

  const body = await response.json().catch(() => null);
  if (!response.ok) {
    const message = body?.error?.message || `${response.status} ${response.statusText}`;
    throw new Error(message);
  }

  const text = responseOutputText(body);
  if (!text) throw new Error("OpenRouter response did not include text output");
  return JSON.parse(text) as LlmMappingDecision;
}

function llmMappingPayload(company: CompanyMatch, fillRow: FillRow, periods: string[], candidates: LlmCandidateFact[]) {
  return {
    company: { ticker: company.ticker, name: company.title },
    row: {
      label: fillRow.label,
      statement: fillRow.statement,
      kind: fillRow.kind,
      sectionHeader: fillRow.modelContext?.sectionHeader ?? "",
      previousLabel: fillRow.modelContext?.previousLabel ?? "",
      nextLabel: fillRow.modelContext?.nextLabel ?? "",
      modelSignConvention: fillRow.modelContext?.signConvention ?? fillRow.sign ?? 1
    },
    periods,
    candidates
  };
}

function llmMappingJsonSchema() {
  return {
    name: "edgar_mapping_decision",
    strict: true,
    schema: {
      type: "object",
      additionalProperties: false,
      required: ["operation", "selectedConcepts", "sign", "confidence", "reason", "requiresReview"],
      properties: {
        operation: { type: "string", enum: ["direct", "sum", "difference", "needs_review"] },
        selectedConcepts: {
          type: "array",
          items: { type: "string" }
        },
        sign: { type: "integer", enum: [1, -1] },
        confidence: { type: "string", enum: ["high", "medium", "low"] },
        reason: { type: "string" },
        requiresReview: { type: "boolean" }
      }
    }
  };
}

function responseOutputText(body: any) {
  if (typeof body?.output_text === "string") return body.output_text;
  const content = body?.choices?.[0]?.message?.content;
  if (typeof content === "string") return content.trim();
  if (content && typeof content === "object" && !Array.isArray(content)) return JSON.stringify(content);
  if (Array.isArray(content)) {
    return content
      .map((item) => (typeof item?.text === "string" ? item.text : typeof item?.content === "string" ? item.content : ""))
      .join("")
      .trim();
  }
  const chunks: string[] = [];
  for (const item of body?.output ?? []) {
    if (item?.type !== "message") continue;
    for (const content of item.content ?? []) {
      if (content?.type === "output_text" && typeof content.text === "string") chunks.push(content.text);
    }
  }
  return chunks.join("").trim();
}

function llmDecisionToFillRow(fillRow: FillRow, decision: LlmMappingDecision, candidates: LlmCandidateFact[], modelChoice: LlmMappingModelChoice): FillRow | null {
  if (decision.operation === "needs_review" || decision.requiresReview || decision.confidence === "low") return null;
  const allowed = new Set(candidates.map((candidate) => candidate.concept));
  const selected = unique(decision.selectedConcepts).filter((concept) => allowed.has(concept));
  if (!selected.length || selected.length !== decision.selectedConcepts.length) return null;
  const sign: 1 | -1 = decision.sign === -1 ? -1 : 1;
  const note = `LLM-assisted EDGAR mapping (${modelChoice.tier} model, ${modelChoice.reason}): ${decision.reason}`;

  if (decision.operation === "direct") {
    return {
      ...fillRow,
      classification: "direct" as const,
      concepts: [selected[0]],
      resolver: undefined,
      sign,
      scale: 1_000_000,
      comment: note,
      noFillComment: undefined
    };
  }

  return {
    ...fillRow,
    classification: "grouped" as const,
    concepts: selected,
    resolver: (period: string, ctx: ResolveContext) => resolveLlmConceptGroup(period, ctx, fillRow.kind, selected, decision.operation, sign, note),
    sign,
    scale: 1_000_000,
    comment: note,
    noFillComment: undefined
  };
}

function resolveLlmConceptGroup(
  period: string,
  ctx: ResolveContext,
  kind: FillRow["kind"],
  concepts: string[],
  operation: LlmMappingDecision["operation"],
  sign: 1 | -1,
  note: string
): ResolvedValue {
  const facts = kind === "instant" ? ctx.instant.get(period) : ctx.duration.get(period);
  if (!facts) return { value: null, sources: [], note, classification: "grouped" };
  const sources = concepts.map((concept) => facts.get(concept)).filter(Boolean) as FactSource[];
  if (sources.length !== concepts.length) return { value: null, sources, note, classification: "grouped" };
  const unsigned =
    operation === "difference"
      ? sources.slice(1).reduce((value, source) => value - source.value, sources[0]?.value ?? 0)
      : sources.reduce((total, source) => total + source.value, 0);
  return {
    value: unsigned * sign,
    sources,
    note,
    classification: sources.length > 1 ? "grouped" : "direct"
  };
}

function first(period: string, map: Map<string, Map<string, FactSource>>, concepts: string[]) {
  const facts = map.get(period);
  if (!facts) return null;
  for (const concept of concepts) {
    const source = facts.get(concept);
    if (source) return source;
  }
  return null;
}

function firstWithPriorInstant(period: string, map: Map<string, Map<string, FactSource>>, concepts: string[]) {
  const direct = first(period, map, concepts);
  if (direct) return direct;

  const priorPeriods = Array.from(map.keys())
    .filter((candidate) => isSupportedPeriodKey(candidate) && comparePeriods(candidate, period) < 0)
    .sort(comparePeriods)
    .reverse();
  for (const priorPeriod of priorPeriods) {
    const source = first(priorPeriod, map, concepts);
    if (!source) continue;
    return {
      ...source,
      note: `${sourceDisplayLabel(source)} carried forward from ${priorPeriod} because the instant balance was not separately reported for ${period}.`
    };
  }
  return null;
}

function sum(period: string, map: Map<string, Map<string, FactSource>>, concepts: string[]): ResolvedValue | null {
  const facts = map.get(period);
  if (!facts) return null;
  const sources = concepts.map((concept) => facts.get(concept)).filter(Boolean) as FactSource[];
  if (!sources.length) return null;
  return {
    value: sources.reduce((total, source) => total + source.value, 0),
    sources
  };
}

function sumWithNote(period: string, map: Map<string, Map<string, FactSource>>, concepts: string[], note: string): ResolvedValue {
  const result = sum(period, map, concepts);
  return result ? { ...result, note, classification: result.sources.length > 1 ? "grouped" : "direct" } : { value: null, sources: [], note };
}

function sumResolved(items: Array<ResolvedValue | null | undefined>, note: string): ResolvedValue {
  const valid = items.filter((item): item is ResolvedValue => Boolean(item && item.value !== null));
  if (!valid.length) return { value: null, sources: [], note };
  return {
    value: valid.reduce((total, item) => total + (item.value ?? 0), 0),
    sources: compactSources(valid),
    note,
    classification: "grouped"
  };
}

function difference(period: string, map: Map<string, Map<string, FactSource>>, totalConcepts: string[], lessConceptGroups: string[][], note: string): ResolvedValue {
  const total = first(period, map, totalConcepts);
  if (!total) return { value: null, sources: [], note };
  const less = lessConceptGroups.map((concepts) => first(period, map, concepts) ?? zeroSource(concepts[0]));
  return {
    value: total.value - less.reduce((acc, source) => acc + source.value, 0),
    sources: [total, ...less],
    note,
    classification: "grouped"
  };
}

function signed(source: FactSource | ResolvedValue | null, sign: 1 | -1) {
  if (!source || source.value === null) return null;
  return {
    ...source,
    value: sign === -1 ? -Math.abs(source.value) : source.value,
    sources: "sources" in source ? source.sources : [source]
  };
}

function zeroSource(concept: string): FactSource {
  return { concept, label: concept, value: 0 };
}

function compactSources(items: Array<FactSource | ResolvedValue | null | undefined>) {
  return items.flatMap((item) => {
    if (!item) return [];
    return "sources" in item ? item.sources : [item];
  });
}

function mappingComment(
  fillRow: FillRow,
  resolved: ResolvedValue,
  period: string,
  valueWritten: number,
  confidence: "high" | "medium" | "low",
  notes?: string | null
) {
  const source = resolved.sources[0];
  const concepts = resolved.sources.length
    ? resolved.sources.map((item) => `${item.concept} = ${roundModelValue(item.value / (fillRow.scale ?? 1))} mm`).join("; ")
    : "None";
  return [
    "EDGAR mapping:",
    `Source: ${source?.form || "SEC filing"}${source?.accn ? `, accession #${source.accn}` : ""}${source?.filed ? `, filed ${source.filed}` : ""}`,
    `Period: ${period}${source?.end ? ` / period ended ${source.end}` : ""}`,
    `Model row: ${fillRow.modelContext?.sheetName || MODEL_SHEET}!${fillRow.row} ${fillRow.label}`,
    `Mapping type: ${resolved.classification || fillRow.classification}`,
    `Concepts used: ${concepts}`,
    "Unit conversion: raw USD to $mm",
    `Sign: ${fillRow.sign === -1 ? "inverted because model expects this row as negative" : "copied using model sign convention"}`,
    `Confidence: ${confidence}`,
    `Value written: ${valueWritten}`,
    notes ? `Notes: ${notes}` : resolved.note ? `Notes: ${resolved.note}` : ""
  ]
    .filter(Boolean)
    .join("\n");
}

function mappingCommentForSegment(
  sheet: ExcelJS.Worksheet,
  cell: ExcelJS.Cell,
  modelLabel: string,
  segment: SegmentRevenue,
  period: string,
  valueWritten: number,
  suffix: string
) {
  return [
    "EDGAR mapping:",
    "Source: SEC inline XBRL segment table",
    `Period: ${period}`,
    `Model row: ${sheet.name}!${cell.address} ${modelLabel}`,
    "Mapping type: segment",
    `Concepts used: reportable segment ${segment.label} ${suffix} = ${valueWritten} mm`,
    "Unit conversion: raw USD to $mm",
    "Sign: copied using model sign convention",
    "Confidence: high",
    "Notes: Segment Analysis is mapped only to reportable segment rows that match the template labels; generic or blank segment rows are left unchanged."
  ].join("\n");
}

function mappingAuditRow(
  sheet: ExcelJS.Worksheet,
  cell: ExcelJS.Cell,
  fillRow: FillRow,
  period: string,
  valueWritten: number,
  resolved: ResolvedValue,
  confidence: "high" | "medium" | "low",
  notes?: string | null
): MappingAuditRow {
  return {
    sheetName: sheet.name,
    cell: cell.address,
    modelRowLabel: fillRow.label,
    section: fillRow.modelContext?.sectionHeader ?? "",
    period,
    valueWritten,
    mappingType: resolved.classification || fillRow.classification,
    conceptsUsed: resolved.sources.map((source) => `${source.concept}=${roundModelValue(source.value / (fillRow.scale ?? 1))}mm`).join("; "),
    secLabels: unique(resolved.sources.map((source) => source.label).filter(Boolean)).join("; "),
    sourceStatement: fillRow.statement,
    accession: unique(resolved.sources.map((source) => source.accn).filter(Boolean)).join("; "),
    sourceUrl: unique(resolved.sources.map((source) => source.sourceUrl).filter(Boolean)).join("; "),
    filingForm: unique(resolved.sources.map((source) => source.form).filter(Boolean)).join("; "),
    filedDate: unique(resolved.sources.map((source) => source.filed).filter(Boolean)).join("; "),
    startDate: unique(resolved.sources.map((source) => source.start).filter(Boolean)).join("; "),
    endDate: unique(resolved.sources.map((source) => source.end).filter(Boolean)).join("; "),
    cellWritable: true,
    formulaPreserved: false,
    formulaStatus: "not a formula cell",
    writeBlockedReason: "",
    signConvention: fillRow.sign === -1 ? "inverted to match model sign convention" : "copied",
    confidence,
    validationStatus: "not_run",
    notes: notes || resolved.note || fillRow.comment || ""
  };
}

function skippedMappingAuditRow(sheet: ExcelJS.Worksheet, cell: ExcelJS.Cell, fillRow: FillRow, period: string, decision: WriteDecision): MappingAuditRow {
  return {
    sheetName: sheet.name,
    cell: cell.address,
    modelRowLabel: fillRow.label,
    section: fillRow.modelContext?.sectionHeader ?? "",
    period,
    valueWritten: numericCellValue(cell) ?? 0,
    mappingType: decision.formulaPreserved ? "formula preserved" : "skipped",
    conceptsUsed: fillRow.concepts?.join("; ") ?? "",
    secLabels: "",
    sourceStatement: fillRow.statement,
    accession: "",
    sourceUrl: "",
    filingForm: "",
    filedDate: "",
    startDate: "",
    endDate: "",
    cellWritable: false,
    formulaPreserved: decision.formulaPreserved,
    formulaStatus: decision.formulaPreserved ? "preserved" : "not written",
    writeBlockedReason: decision.reason ?? "not writable",
    signConvention: "not written",
    confidence: "low",
    validationStatus: decision.formulaPreserved ? "formula_preserved" : "skipped",
    notes: decision.formulaPreserved ? "Formula preserved; EDGAR value was not written over an existing model formula." : "Skipped because the cell was not an active model input."
  };
}

function mappingAuditRowForSegment(
  sheet: ExcelJS.Worksheet,
  cell: ExcelJS.Cell,
  modelLabel: string,
  period: string,
  valueWritten: number,
  resolved: ResolvedValue,
  suffix: string
): MappingAuditRow {
  return {
    sheetName: sheet.name,
    cell: cell.address,
    modelRowLabel: modelLabel,
    section: "Segment Analysis",
    period,
    valueWritten,
    mappingType: "segment",
    conceptsUsed: resolved.sources.map((source) => `${source.label} ${suffix}=${valueWritten}mm`).join("; "),
    secLabels: unique(resolved.sources.map((source) => source.label).filter(Boolean)).join("; "),
    sourceStatement: "segment",
    accession: unique(resolved.sources.map((source) => source.accn).filter(Boolean)).join("; "),
    sourceUrl: unique(resolved.sources.map((source) => source.sourceUrl).filter(Boolean)).join("; "),
    filingForm: unique(resolved.sources.map((source) => source.form).filter(Boolean)).join("; "),
    filedDate: unique(resolved.sources.map((source) => source.filed).filter(Boolean)).join("; "),
    startDate: unique(resolved.sources.map((source) => source.start).filter(Boolean)).join("; "),
    endDate: unique(resolved.sources.map((source) => source.end).filter(Boolean)).join("; "),
    cellWritable: true,
    formulaPreserved: false,
    formulaStatus: "not a formula cell",
    writeBlockedReason: "",
    signConvention: "copied",
    confidence: "high",
    validationStatus: "not_run",
    notes: resolved.note || ""
  };
}

function validateWorkbookBeforeReturn(
  workbook: ExcelJS.Workbook,
  periods: string[],
  columns: number[],
  ctx: ResolveContext,
  warnings: string[],
  modelSheetName = MODEL_SHEET,
  profile: TemplateProfile = { kind: "generic", confidence: "low", rationale: [], sheetName: modelSheetName, hasSegmentAnalysis: Boolean(workbook.getWorksheet(SEGMENT_SHEET)) }
) {
  const errors: string[] = [];
  const segmentSheet = workbook.getWorksheet(SEGMENT_SHEET);
  if (segmentSheet) {
    const segmentEvaluator = new FormulaEvaluator(segmentSheet);
    errors.push(...validateSegmentGenericRows(segmentSheet, periods, columns));
    errors.push(...validateSegmentRevenueTieOut(segmentSheet, periods, columns, ctx, segmentEvaluator, warnings));
    errors.push(
      ...validateSegmentStatementTieOut(
        segmentSheet,
        periods,
        columns,
        ctx,
        segmentEvaluator,
        warnings,
        "operating income",
        "Total Company Operating Income",
        "Operating Income Check",
        "Operating Income",
        C.operatingIncome
      )
    );
    errors.push(
      ...validateSegmentStatementTieOut(segmentSheet, periods, columns, ctx, segmentEvaluator, warnings, "D&A", "Total D&A", "D&A Check", "D&A", C.da)
    );
  }
  if (modelSheetName !== MODEL_SHEET) return errors;

  const modelSheet = workbook.getWorksheet(modelSheetName) ?? workbook.getWorksheet(MODEL_SHEET);
  if (modelSheet) {
    const evaluator = new FormulaEvaluator(modelSheet);
    errors.push(...validateIncomeStatementKeyMetrics(modelSheet, periods, columns, ctx, evaluator, warnings, profile));
    errors.push(...validateBalanceSheetStatementTotals(modelSheet, periods, columns, ctx, evaluator, warnings));
    errors.push(...validateBalanceSheetCheck(modelSheet, periods, columns, evaluator));

    const interestExpenseRow = findLabelRow(modelSheet, "Interest Expense");
    if (interestExpenseRow) {
      const hasPopulatedInterestExpense = columns.some((col) => Math.abs(numericCellValue(modelSheet.getCell(interestExpenseRow, col)) ?? 0) > 0.0001);
      if (!hasPopulatedInterestExpense) {
        warnings.push("Model Interest Expense row is present but all detected historical cells are zero/blank.");
      }
    }
  }
  return errors;
}

class FormulaEvaluator {
  private readonly cache = new Map<string, number | null>();

  constructor(private readonly sheet: ExcelJS.Worksheet) {}

  clear() {
    this.cache.clear();
  }

  evaluateCell(cell: ExcelJS.Cell, visited = new Set<string>()): number | null {
    const address = `${cell.worksheet.name}!${cell.address}`;
    if (this.cache.has(address)) return this.cache.get(address) ?? null;

    if (visited.has(address)) return null;
    visited.add(address);

    const value = cell.value;
    let result: number | null = null;
    if (typeof value === "number") {
      result = value;
    } else if (value && typeof value === "object" && ("formula" in value || "sharedFormula" in value)) {
      const formula = formulaForCell(cell);
      result = formula ? this.evaluateFormula(formula, cell, visited) : null;
      if (result === null && "result" in value && typeof value.result === "number") result = value.result;
    } else if (value && typeof value === "object" && "result" in value && typeof value.result === "number") {
      result = value.result;
    }

    visited.delete(address);
    this.cache.set(address, result);
    return result;
  }

  private evaluateFormula(formula: string, cell: ExcelJS.Cell, visited: Set<string>): number | null {
    const expression = formula.replace(/^=/, "");
    const ifValue = this.evaluateIfExpression(expression, cell, visited);
    if (ifValue !== undefined) return ifValue;

    const withoutSums = this.replaceSumCalls(expression, cell, visited);
    if (withoutSums === null) return null;

    const withRefs = withoutSums.replace(/((?:'[^']+'|[A-Za-z0-9_ ]+)!)?\$?([A-Z]{1,3})\$?(\d+)/g, (reference, sheetPrefix: string, col: string, row: string) => {
      const target = referencedCell(cell, sheetPrefix, col, row);
      if (!target) return "0";
      const value = this.evaluateCell(target, visited);
      return value === null ? "0" : `(${value})`;
    });

    if (!/^[\d+\-*/().,\sNaN]+$/.test(withRefs)) return null;
    try {
      const value = Function(`"use strict"; return (${withRefs});`)();
      return typeof value === "number" && Number.isFinite(value) ? value : null;
    } catch {
      return null;
    }
  }

  private replaceSumCalls(expression: string, cell: ExcelJS.Cell, visited: Set<string>) {
    let output = expression;
    const sumPattern = /SUM\(([^()]+)\)/i;
    while (sumPattern.test(output)) {
      output = output.replace(sumPattern, (_match, body: string) => {
        const value = this.evaluateSum(body, cell, visited);
        return value === null ? "NaN" : String(value);
      });
      if (output.includes("NaN")) return null;
    }
    return output;
  }

  private evaluateSum(body: string, cell: ExcelJS.Cell, visited: Set<string>) {
    let total = 0;
    for (const part of splitFormulaArgs(body)) {
      const item = part.trim();
      const range = item.match(/^((?:'[^']+'|[A-Za-z0-9_ ]+)!)?\$?([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)$/i);
      if (range) {
        const targetSheet = referencedSheet(cell, range[1]);
        if (!targetSheet) return null;
        const startCol = columnIndex(range[2]);
        const startRow = Number(range[3]);
        const endCol = columnIndex(range[4]);
        const endRow = Number(range[5]);
        for (let row = Math.min(startRow, endRow); row <= Math.max(startRow, endRow); row += 1) {
          for (let col = Math.min(startCol, endCol); col <= Math.max(startCol, endCol); col += 1) {
            const value = this.evaluateCell(targetSheet.getCell(row, col), visited);
            total += value ?? 0;
          }
        }
        continue;
      }

      const value = this.evaluateFormula(item, cell, visited);
      if (value === null) return null;
      total += value;
    }
    return total;
  }

  private evaluateIfExpression(expression: string, cell: ExcelJS.Cell, visited: Set<string>) {
    const body = functionBody(expression, "IF");
    if (body === null) return undefined;
    const [condition, whenTrue, whenFalse] = splitFormulaArgs(body);
    if (!condition || !whenTrue) return null;
    const conditionResult = this.evaluateCondition(condition, cell, visited);
    if (conditionResult === null) return null;
    return this.evaluateFormula(conditionResult ? whenTrue : whenFalse ?? "0", cell, visited);
  }

  private evaluateCondition(condition: string, cell: ExcelJS.Cell, visited: Set<string>) {
    const match = condition.match(/^\s*(\$?[A-Z]{1,3}\$?\d+)\s*(=|<>)\s*"([^"]*)"\s*$/);
    if (match) {
      const actual = cellDisplay(cell.worksheet.getCell(match[1].replace(/\$/g, "")));
      return match[2] === "=" ? actual === match[3] : actual !== match[3];
    }

    const numericMatch = condition.match(/^\s*(.+?)\s*(=|<>|>=|<=|>|<)\s*(-?\d+(?:\.\d+)?)\s*$/);
    if (!numericMatch) return null;
    const left = this.evaluateFormula(numericMatch[1], cell, visited);
    const right = Number(numericMatch[3]);
    if (left === null || Number.isNaN(right)) return null;
    if (numericMatch[2] === "=") return valuesTie(left, right);
    if (numericMatch[2] === "<>") return !valuesTie(left, right);
    if (numericMatch[2] === ">=") return left >= right;
    if (numericMatch[2] === "<=") return left <= right;
    if (numericMatch[2] === ">") return left > right;
    return left < right;
  }
}

function referencedCell(cell: ExcelJS.Cell, sheetPrefix: string | undefined, col: string, row: string) {
  const sheet = referencedSheet(cell, sheetPrefix);
  return sheet?.getCell(`${col}${row}`.replace(/\$/g, "")) ?? null;
}

function referencedSheet(cell: ExcelJS.Cell, sheetPrefix?: string) {
  if (!sheetPrefix) return cell.worksheet;
  const sheetName = sheetPrefix
    .slice(0, -1)
    .replace(/^'/, "")
    .replace(/'$/, "")
    .replace(/''/g, "'");
  return cell.worksheet.workbook.getWorksheet(sheetName) ?? null;
}

function functionBody(expression: string, name: string) {
  const trimmed = expression.trim();
  const prefix = `${name}(`;
  if (!trimmed.toUpperCase().startsWith(prefix)) return null;
  if (!trimmed.endsWith(")")) return null;
  return trimmed.slice(prefix.length, -1);
}

function splitFormulaArgs(body: string) {
  const args: string[] = [];
  let current = "";
  let depth = 0;
  let inString = false;
  for (let index = 0; index < body.length; index += 1) {
    const char = body[index];
    if (char === '"') inString = !inString;
    if (!inString && char === "(") depth += 1;
    if (!inString && char === ")") depth -= 1;
    if (!inString && depth === 0 && char === ",") {
      args.push(current);
      current = "";
      continue;
    }
    current += char;
  }
  args.push(current);
  return args;
}

function refreshHistoricalFormulaCachedResults(workbook: ExcelJS.Workbook, columns: number[], sheetNames = [MODEL_SHEET, SEGMENT_SHEET]) {
  for (const sheetName of sheetNames) {
    const sheet = workbook.getWorksheet(sheetName);
    if (!sheet) continue;
    const evaluator = new FormulaEvaluator(sheet);
    const firstCol = Math.min(...columns);
    const lastCol = Math.min(sheet.columnCount, Math.max(...columns) + 1);
    for (let rowNumber = 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
      for (let col = firstCol; col <= lastCol; col += 1) {
        const cell = sheet.getCell(rowNumber, col);
        if (hasFormula(cell)) evaluator.evaluateCell(cell);
      }
    }
  }
}

function ensureFormulaDisplayCaches(workbook: ExcelJS.Workbook, columns: number[], sheetNames = [MODEL_SHEET, SEGMENT_SHEET]) {
  for (const sheetName of sheetNames) {
    const sheet = workbook.getWorksheet(sheetName);
    if (!sheet) continue;
    const evaluator = new FormulaEvaluator(sheet);
    const firstCol = Math.min(sheet.columnCount, Math.max(...columns) + 2);
    const lastCol = sheet.columnCount;

    for (let rowNumber = 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
      for (let col = firstCol; col <= lastCol; col += 1) {
        const cell = sheet.getCell(rowNumber, col);
        if (!hasFormula(cell) || hasFormulaResult(cell)) continue;
        const formula = formulaForCell(cell);
        if (!formula || !isNumericDisplayFormula(formula)) continue;
        const result = evaluator.evaluateCell(cell);
        setFormulaResult(cell, result ?? 0);
      }
    }
  }
}

function hasFormulaResult(cell: ExcelJS.Cell) {
  const value = cell.value;
  return Boolean(value && typeof value === "object" && "result" in value && value.result !== undefined && value.result !== null);
}

function isNumericDisplayFormula(formula: string) {
  const normalized = formula.replace(/^=/, "").trim();
  if (!normalized || normalized.includes('"')) return false;
  if (/\b(?:HYPERLINK|CONCAT|TEXT|LEFT|RIGHT|MID|REPT|T|NA)\s*\(/i.test(normalized)) return false;
  return /[A-Z]{1,3}\$?\d+|\b(?:SUM|AVERAGE|IF|MAX|MIN|ROUND|ABS)\s*\(|[\d)]\s*[+\-*/^]\s*[\d(]/i.test(normalized);
}

function clearStaleFormulaErrorResults(workbook: ExcelJS.Workbook) {
  workbook.eachSheet((sheet) => {
    sheet.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        const value = cell.value;
        if (!value || typeof value !== "object") return;
        if ("formula" in value && typeof value.formula === "string" && isFormulaErrorResult(value.result)) {
          cell.value = { formula: value.formula };
        } else if ("sharedFormula" in value && typeof value.sharedFormula === "string" && isFormulaErrorResult(value.result)) {
          cell.value = { sharedFormula: value.sharedFormula };
        }
      });
    });
  });
}

function isFormulaErrorResult(value: unknown) {
  return typeof value === "string" && /^#(?:REF|VALUE|DIV\/0|NAME\?|N\/A|NUM|NULL)!?$/i.test(value);
}

function formulaForCell(cell: ExcelJS.Cell) {
  const value = cell.value;
  if (!value || typeof value !== "object") return null;
  if ("formula" in value && typeof value.formula === "string") return value.formula;
  if ("sharedFormula" in value && typeof value.sharedFormula === "string") {
    const master = cell.worksheet.getCell(value.sharedFormula);
    const masterFormula = cellFormula(master);
    return masterFormula ? translateSharedFormula(masterFormula, master.address, cell.address) : null;
  }
  return null;
}

function translateSharedFormula(formula: string, sourceAddress: string, targetAddress: string) {
  const source = parseCellAddress(sourceAddress);
  const target = parseCellAddress(targetAddress);
  if (!source || !target) return formula;
  const colOffset = target.col - source.col;
  const rowOffset = target.row - source.row;
  return formula.replace(/(\$?)([A-Z]{1,3})(\$?)(\d+)/g, (_match, colAbs: string, colLetters: string, rowAbs: string, rowDigits: string) => {
    const nextCol = colAbs ? columnIndex(colLetters) : columnIndex(colLetters) + colOffset;
    const nextRow = rowAbs ? Number(rowDigits) : Number(rowDigits) + rowOffset;
    return `${colAbs}${columnLetter(nextCol)}${rowAbs}${nextRow}`;
  });
}

function parseCellAddress(address: string) {
  const match = address.match(/^([A-Z]{1,3})(\d+)$/i);
  if (!match) return null;
  return { col: columnIndex(match[1]), row: Number(match[2]) };
}

function columnIndex(letters: string) {
  return letters
    .toUpperCase()
    .split("")
    .reduce((total, letter) => total * 26 + letter.charCodeAt(0) - 64, 0);
}

function validateIncomeStatementKeyMetrics(
  sheet: ExcelJS.Worksheet,
  periods: string[],
  columns: number[],
  ctx: ResolveContext,
  evaluator: FormulaEvaluator,
  warnings: string[],
  profile: TemplateProfile
) {
  const errors: string[] = [];
  if (profile.kind === "financial_company" && hasAnyConcept(ctx, "duration", C.netRevenue)) {
    errors.push(
      ...validateIncomeStatementMetricAgainstEdgar(
        sheet,
        periods,
        columns,
        ctx,
        evaluator,
        warnings,
        "net revenue",
        ["Net Revenue", "Revenue Net of Interest Expense", "Revenues Net of Interest Expense", "Total Net Revenue"],
        C.netRevenue
      )
    );
  } else {
    errors.push(
      ...validateIncomeStatementMetricAgainstEdgar(
        sheet,
        periods,
        columns,
        ctx,
        evaluator,
        warnings,
        "Revenue",
        ["Revenue", "Revenues", "Total Revenue", "Total Revenues"],
        C.revenue
      )
    );
  }
  errors.push(
    ...validateIncomeStatementMetricAgainstEdgar(
      sheet,
      periods,
      columns,
      ctx,
      evaluator,
      warnings,
      "EBIT",
      ["EBIT"],
      C.operatingIncome
    )
  );
  errors.push(
    ...validateIncomeStatementMetricAgainstEdgar(
      sheet,
      periods,
      columns,
      ctx,
      evaluator,
      warnings,
      "net income",
      ["Net Income (Loss)", "Net Income"],
      CONTINUING_NET_INCOME_CONCEPTS
    )
  );
  errors.push(...validateIncomeStatementEbitdaFormula(sheet, periods, columns, evaluator, warnings));
  return errors;
}

function validateIncomeStatementMetricAgainstEdgar(
  sheet: ExcelJS.Worksheet,
  periods: string[],
  columns: number[],
  ctx: ResolveContext,
  evaluator: FormulaEvaluator,
  warnings: string[],
  metricName: string,
  labels: string[],
  concepts: string[]
) {
  const errors: string[] = [];
  const rowNumber = findIncomeStatementMetricRow(sheet, labels);
  if (!rowNumber) return errors;

  periods.forEach((period, index) => {
    const edgarValue = first(period, ctx.duration, concepts);
    if (!edgarValue) {
      warnings.unshift(`Income Statement ${period}: ${metricName} tie-out skipped because EDGAR ${concepts.join("/")} was unavailable for that period.`);
      return;
    }

    const cell = sheet.getCell(rowNumber, columns[index]);
    const expected = edgarValue.value / 1_000_000;
    const modelValue = statementMetricCellValue(cell, evaluator, expected);
    if (modelValue === null) {
      const message = `Income Statement ${cell.address} ${period}: could not evaluate model ${metricName}.`;
      if (hasFormula(cell)) warnings.unshift(message);
      else errors.push(message);
      return;
    }

    if (!statementMetricTies(modelValue, expected)) {
      const message = `Income Statement ${cell.address} ${period}: ${metricName} ${roundModelValue(modelValue)} does not match EDGAR ${roundModelValue(expected)}.`;
      if (hasFormula(cell)) warnings.unshift(message);
      else errors.push(message);
    }
  });

  return errors;
}

function statementMetricCellValue(cell: ExcelJS.Cell, evaluator: FormulaEvaluator, expected: number) {
  const cached = numericCellValue(cell);
  if (cached !== null && statementMetricTies(cached, expected)) return cached;
  return evaluator.evaluateCell(cell);
}

function statementMetricTies(actual: number, expected: number) {
  return Math.abs(actual - expected) <= Math.max(1, Math.abs(expected) * 0.0005);
}

function validateIncomeStatementEbitdaFormula(
  sheet: ExcelJS.Worksheet,
  periods: string[],
  columns: number[],
  evaluator: FormulaEvaluator,
  warnings: string[]
) {
  const errors: string[] = [];
  const ebitdaRow = findIncomeStatementMetricRow(sheet, ["EBITDA"]);
  if (!ebitdaRow) return errors;
  const ebitRow = findPriorLabelRow(sheet, ebitdaRow, "EBIT", 8);
  const daRow = findPriorLabelRow(sheet, ebitdaRow, "Depreciation & Amortization", 4) ?? findPriorLabelRow(sheet, ebitdaRow, "Depreciation and Amortization", 4);
  if (!ebitRow || !daRow) {
    warnings.unshift("Income Statement Analysis: EBITDA tie-out skipped because EBIT or depreciation & amortization row was unavailable.");
    return errors;
  }

  periods.forEach((period, index) => {
    const col = columns[index];
    const ebitdaCell = sheet.getCell(ebitdaRow, col);
    const ebitCell = sheet.getCell(ebitRow, col);
    const daCell = sheet.getCell(daRow, col);
    const cachedEbitda = numericCellValue(ebitdaCell);
    const cachedEbit = numericCellValue(ebitCell);
    const cachedDa = numericCellValue(daCell);
    if (cachedEbitda !== null && cachedEbit !== null && cachedDa !== null && valuesTie(cachedEbitda, cachedEbit + cachedDa)) return;

    const ebitda = evaluator.evaluateCell(ebitdaCell);
    const ebit = evaluator.evaluateCell(ebitCell);
    const da = evaluator.evaluateCell(daCell);
    if (ebitda === null || ebit === null || da === null) {
      const message = `Income Statement ${columnLetter(col)}${ebitdaRow} ${period}: could not evaluate EBITDA, EBIT, or depreciation & amortization.`;
      if (hasFormula(ebitdaCell) || hasFormula(ebitCell) || hasFormula(daCell)) warnings.unshift(message);
      else errors.push(message);
      return;
    }
    const expected = ebit + da;
    if (!valuesTie(ebitda, expected)) {
      const message = `Income Statement ${columnLetter(col)}${ebitdaRow} ${period}: EBITDA ${roundModelValue(ebitda)} does not equal EBIT plus D&A ${roundModelValue(expected)}.`;
      if (hasFormula(ebitdaCell) || hasFormula(ebitCell) || hasFormula(daCell)) warnings.unshift(message);
      else errors.push(message);
    }
  });

  return errors;
}

function findIncomeStatementMetricRow(sheet: ExcelJS.Worksheet, labels: string[]) {
  return findRowInSection(sheet, "Income Statement", labels, (label) =>
    /income statement analysis|cash flow statement|cashflow statement|balance sheet/i.test(label)
  );
}

function findPriorLabelRow(sheet: ExcelJS.Worksheet, startRow: number, label: string, maxRows: number) {
  const wanted = normalize(label);
  for (let rowNumber = startRow - 1; rowNumber >= Math.max(1, startRow - maxRows); rowNumber -= 1) {
    if (normalize(rowLabel(sheet, rowNumber)) === wanted) return rowNumber;
  }
  return null;
}

function validateBalanceSheetStatementTotals(
  sheet: ExcelJS.Worksheet,
  periods: string[],
  columns: number[],
  ctx: ResolveContext,
  evaluator: FormulaEvaluator,
  warnings: string[]
) {
  const errors: string[] = [];
  errors.push(
    ...validateBalanceSheetMetricAgainstEdgar(sheet, periods, columns, ctx, evaluator, warnings, "total assets", ["Total Assets"], C.assets)
  );
  errors.push(
    ...validateBalanceSheetMetricAgainstEdgar(sheet, periods, columns, ctx, evaluator, warnings, "total liabilities", ["Total Liabilities"], C.liabilities)
  );
  errors.push(
    ...validateBalanceSheetMetricAgainstEdgar(
      sheet,
      periods,
      columns,
      ctx,
      evaluator,
      warnings,
      "shareholders' equity",
      [
        "Total Shareholder's Equity",
        "Total Shareholders' Equity",
        "Total Shareholders Equity",
        "Total Stockholders' Equity",
        "Total Stockholders Equity",
        "Total Equity"
      ],
      C.equity
    )
  );
  errors.push(
    ...validateBalanceSheetMetricAgainstEdgar(
      sheet,
      periods,
      columns,
      ctx,
      evaluator,
      warnings,
      "total liabilities plus shareholders' equity",
      ["Total Liabilities & Shareholder's Equity", "Total Liabilities and Shareholder's Equity"],
      C.assets
    )
  );
  return errors;
}

function validateBalanceSheetMetricAgainstEdgar(
  sheet: ExcelJS.Worksheet,
  periods: string[],
  columns: number[],
  ctx: ResolveContext,
  evaluator: FormulaEvaluator,
  warnings: string[],
  metricName: string,
  labels: string[],
  concepts: string[]
) {
  const errors: string[] = [];
  const rowNumber = findRowInSection(sheet, "Balance Sheet", labels, (label) =>
    /working capital|cash flow statement|cashflow statement|income statement|schedule|analysis|drivers/i.test(label)
  );
  if (!rowNumber) return errors;

  periods.forEach((period, index) => {
    const edgarValue = first(period, ctx.instant, concepts);
    if (!edgarValue) {
      warnings.unshift(`Balance Sheet ${period}: ${metricName} tie-out skipped because EDGAR ${concepts.join("/")} was unavailable for that period.`);
      return;
    }

    const cell = sheet.getCell(rowNumber, columns[index]);
    const expected = edgarValue.value / 1_000_000;
    const modelValue = statementMetricCellValue(cell, evaluator, expected);
    if (modelValue === null) {
      const message = `Balance Sheet ${cell.address} ${period}: could not evaluate model ${metricName}.`;
      if (hasFormula(cell)) warnings.unshift(message);
      else errors.push(message);
      return;
    }

    if (!statementMetricTies(modelValue, expected)) {
      const message = `Balance Sheet ${cell.address} ${period}: ${metricName} ${roundModelValue(modelValue)} does not match EDGAR ${roundModelValue(expected)}.`;
      if (hasFormula(cell)) warnings.unshift(message);
      else errors.push(message);
    }
  });

  return errors;
}

function validateBalanceSheetCheck(sheet: ExcelJS.Worksheet, periods: string[], columns: number[], evaluator: FormulaEvaluator) {
  const errors: string[] = [];
  const checkRow = findRowInSection(sheet, "Balance Sheet", ["Balance Sheet Check"], (label) =>
    /working capital|cash flow statement|cashflow statement|income statement|schedule|analysis|drivers/i.test(label)
  );

  if (checkRow) {
    periods.forEach((period, index) => {
      const cell = sheet.getCell(checkRow, columns[index]);
      const check = evaluator.evaluateCell(cell);
      if (check === null) {
        errors.push(`Balance Sheet ${cell.address} ${period}: could not evaluate the model's balance sheet check row.`);
      } else if (!valuesTie(check, 0)) {
        errors.push(`Balance Sheet ${cell.address} ${period}: check is ${roundModelValue(check)}, not OK.`);
      }
    });
    return errors;
  }

  const assetsRow = findRowInSection(sheet, "Balance Sheet", ["Total Assets"], (label) =>
    /working capital|cash flow statement|cashflow statement|income statement|schedule|analysis|drivers/i.test(label)
  );
  const liabilitiesAndEquityRow = findRowInSection(sheet, "Balance Sheet", ["Total Liabilities & Shareholder's Equity", "Total Liabilities and Shareholder's Equity"], (label) =>
    /working capital|cash flow statement|cashflow statement|income statement|schedule|analysis|drivers/i.test(label)
  );
  if (!assetsRow || !liabilitiesAndEquityRow) return errors;

  periods.forEach((period, index) => {
    const col = columns[index];
    const assets = evaluator.evaluateCell(sheet.getCell(assetsRow, col));
    const liabilitiesAndEquity = evaluator.evaluateCell(sheet.getCell(liabilitiesAndEquityRow, col));
    if (assets === null || liabilitiesAndEquity === null) {
      errors.push(`Balance Sheet ${period}: could not evaluate total assets or total liabilities plus shareholder's equity.`);
    } else if (!valuesTie(assets, liabilitiesAndEquity)) {
      errors.push(
        `Balance Sheet ${period}: total assets ${roundModelValue(assets)} do not equal total liabilities plus shareholder's equity ${roundModelValue(liabilitiesAndEquity)}.`
      );
    }
  });

  return errors;
}

function findRowInSection(sheet: ExcelJS.Worksheet, sectionLabel: string, labels: string[], isBoundary: (label: string) => boolean) {
  const sectionStart = findSectionHeaderRow(sheet, sectionLabel);
  if (!sectionStart) return null;
  const wanted = labels.map(normalize);
  for (let rowNumber = sectionStart + 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const label = rowLabel(sheet, rowNumber);
    if (!label) continue;
    if (isBoundary(label)) break;
    if (wanted.includes(normalize(label))) return rowNumber;
  }
  return null;
}

function findSectionHeaderRow(sheet: ExcelJS.Worksheet, sectionLabel: string) {
  const wanted = normalize(sectionLabel);
  let best: { row: number; score: number } | null = null;
  for (let rowNumber = 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
    if (normalize(rowLabel(sheet, rowNumber)) !== wanted) continue;
    const score = sectionHeaderScore(sheet, rowNumber);
    if (!best || score > best.score || (score === best.score && rowNumber > best.row)) {
      best = { row: rowNumber, score };
    }
  }
  return best?.row ?? null;
}

function sectionHeaderScore(sheet: ExcelJS.Worksheet, rowNumber: number) {
  let score = 0;
  if (/^x$/i.test(cellDisplay(sheet.getCell(rowNumber, 1)).trim())) score += 100;
  for (let col = 6; col <= Math.min(sheet.columnCount, 12); col += 1) {
    if (cellDisplay(sheet.getCell(rowNumber, col)) || cellFormula(sheet.getCell(rowNumber, col))) score += 10;
  }
  return score;
}

function valuesTie(actual: number, expected: number) {
  return Math.abs(actual - expected) <= 0.1;
}

function segmentModelRevenueTies(actual: number, expected: number) {
  return Math.abs(actual - expected) <= Math.max(0.1, Math.abs(expected) * 0.0005);
}

function snapshotWorkbook(workbook: ExcelJS.Workbook, sheetNames: string[], firstHistoricalCol: number): WorkbookSnapshot {
  const labels = new Map<string, string>();
  const formulas = new Map<string, string>();
  const labelSnapshotEndCol = Math.max(1, firstHistoricalCol - 1);
  const formulaSnapshotEndCol = 120;
  for (const sheetName of sheetNames) {
    const sheet = workbook.getWorksheet(sheetName);
    if (!sheet) continue;
    for (let rowNumber = 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
      for (let col = 1; col <= Math.min(labelSnapshotEndCol, sheet.columnCount); col += 1) {
        const address = snapshotAddress(sheet, rowNumber, col);
        labels.set(address, cellDisplay(sheet.getCell(rowNumber, col)));
      }
      for (let col = 1; col <= Math.min(formulaSnapshotEndCol, sheet.columnCount); col += 1) {
        const formula = cellFormula(sheet.getCell(rowNumber, col));
        if (formula) formulas.set(snapshotAddress(sheet, rowNumber, col), formula);
      }
    }
  }
  return { labels, formulas };
}

function validateWorkbookPreservation(workbook: ExcelJS.Workbook, snapshot: WorkbookSnapshot) {
  const errors: string[] = [];
  for (const [address, expected] of snapshot.labels.entries()) {
    const cell = cellFromSnapshotAddress(workbook, address);
    if (!cell) continue;
    const actual = cellDisplay(cell);
    if (actual !== expected) {
      errors.push(`${address}: row label changed from "${expected}" to "${actual}".`);
    }
  }
  for (const [address, expected] of snapshot.formulas.entries()) {
    const cell = cellFromSnapshotAddress(workbook, address);
    if (!cell) continue;
    const actual = cellFormula(cell);
    if (actual !== expected) {
      if (isAllowedFormulaPreservationUpdate(cell, expected, actual)) continue;
      errors.push(`${address}: formula changed from "${expected}" to "${actual ?? "[hardcoded/blank]"}".`);
    }
  }
  return errors;
}

function isAllowedFormulaPreservationUpdate(cell: ExcelJS.Cell, expected: string, actual: string | null) {
  void cell;
  if (actual && isBridgeFormula(expected) && isBridgeFormula(actual)) return true;
  return false;
}

function isFormulaReplacementAllowedForReconciliation(label: string) {
  return /other\s+operating\s+income|other\s+operating\s+expense|income\s*tax|prepaid\s+(?:&|and)\s+other\s+current\s+assets|other\s+current\s+assets|other\s+non-current\s+assets|other\s+long-term\s+assets|other\s+lt\s+assets|other\s+current\s+liabilities|other\s+non-current\s+liabilities|accounts\s+payable\s+(?:&|and)\s+accrued\s+liabilities|accrued\s+liabilities/i.test(label);
}

function isBridgeFormula(formula: string) {
  const bridgeFormula = /^[\d+\-*/().\s]+-SUM\([A-Z]+\d+:[A-Z]+\d+\)$/i;
  return bridgeFormula.test(normalizeBridgeFormulaPrefix(formula));
}

function normalizeBridgeFormulaPrefix(formula: string) {
  const normalized = formula.replace(/^\+\-/, "-").replace(/^\+/, "");
  const wrapped = normalized.match(/^\((.+-SUM\([A-Z]+\d+:[A-Z]+\d+\))\)$/i);
  return wrapped ? wrapped[1] : normalized;
}

function restoreWorkbookLabels(workbook: ExcelJS.Workbook, snapshot: WorkbookSnapshot) {
  for (const [address, expected] of snapshot.labels.entries()) {
    const cell = cellFromSnapshotAddress(workbook, address);
    if (!cell) continue;
    if (cellDisplay(cell) !== expected) cell.value = expected;
  }
}

function cleanStaleProtectedHistoricalRows(
  sheet: ExcelJS.Worksheet,
  fillRows: FillRow[],
  periods: string[],
  columns: number[],
  ctx: ResolveContext,
  auditRows: MappingAuditRow[]
) {
  let clearedCells = 0;
  let commentsAdded = 0;
  const warnings: string[] = [];

  for (const fillRow of fillRows) {
    const staleDetector = staleProtectedRowDetector(fillRow);
    if (!staleDetector) continue;
    const staleMatches = periods.flatMap((period, index) => {
      const col = columns[index];
      const cell = sheet.getCell(fillRow.row, col);
      if (cell.value === null) return [];
      if (hasFormula(cell)) return [];
      const existing = numericCellValue(cell);
      if (existing === null) return [];
      const stale = staleDetector(period, existing, ctx);
      return stale ? [{ period, col, stale }] : [];
    });
    if (!staleMatches.length) continue;

    const clearWholeRow = staleMatches.length >= 2;
    let clearedRow = false;
    periods.forEach((period, index) => {
      const col = columns[index];
      const cell = sheet.getCell(fillRow.row, col);
      if (cell.value === null) return;
      if (hasFormula(cell)) return;
      const existing = numericCellValue(cell);
      if (existing === null) return;
      const stale = staleMatches.find((match) => match.col === col)?.stale ?? (clearWholeRow ? staleMatches[0].stale : null);
      if (!stale) return;
      cell.value = null;
      clearedCells += 1;
      clearedRow = true;
      auditRows.push({
        sheetName: sheet.name,
        cell: cell.address,
        modelRowLabel: fillRow.label,
        period,
        valueWritten: 0,
        mappingType: "unused",
        conceptsUsed: stale.conceptsUsed,
        sourceStatement: fillRow.statement,
        accession: stale.accession,
        sourceUrl: "",
        cellWritable: true,
        formulaPreserved: false,
        writeBlockedReason: "",
        signConvention: "not written",
        confidence: "high",
        validationStatus: "cleared",
        notes: stale.note
      });
    });
    if (clearedRow) {
      fillRow.noFillComment = staleMatches[0].stale.note;
      warnings.push(`${fillRow.label}: cleared stale or unsupported values from protected schedule row.`);
    }
  }

  return { clearedCells, commentsAdded, warnings };
}

type StaleProtectedMatch = {
  note: string;
  conceptsUsed: string;
  accession: string;
};

function staleProtectedRowDetector(fillRow: FillRow): ((period: string, existing: number | null, ctx: ResolveContext) => StaleProtectedMatch | null) | null {
  const context = fillRow.modelContext;
  if (!context || fillRow.classification !== "partial") return null;
  const normalizedLabel = normalize(fillRow.label);

  if (normalizedLabel === normalize("Beginning Balance") && isProtectedScheduleBeginningBalance(context)) {
    return staleBeginningBalanceDetector(fillRow.noFillComment || "Cannot find exact beginning balance plug in EDGAR, find manually.");
  }

  if (inPpeDepreciationScheduleContext(context) && normalizedLabel === normalize("Capital Expenditures")) {
    return staleGenericDetector(C.capex, -1, 1_000_000, "Generic cash-flow capex is not used for PP&E schedule bridge rows.");
  }
  if (inPpeDepreciationScheduleContext(context) && normalizedLabel === normalize("Depreciation Expense")) {
    return staleUnsupportedScheduleDetector("Cannot find exact PP&E schedule depreciation expense in EDGAR, find manually.");
  }
  if (inPpeDepreciationScheduleContext(context) && normalizedLabel === normalize("Acquisition / (Divestment) of Businesses")) {
    return staleGenericDetector(ACQUISITION_CONCEPTS, 1, 1_000_000, "Generic acquisition cash-flow concepts are not used for PP&E schedule bridge rows.");
  }

  return null;
}

function isProtectedScheduleBeginningBalance(context: ModelRowContext) {
  return inSection(
    context,
    "Revolver Balance",
    "Total Debt Balance",
    "Debt and Interest Schedule",
    "Retained Earnings",
    "AOCI Assumptions",
    "Accumulated Other Comprehensive Income",
    "AOCI"
  );
}

function isProtectedBeginningBalanceRow(sheet: ExcelJS.Worksheet, rowNumber: number) {
  if (normalize(rowLabel(sheet, rowNumber)) !== normalize("Beginning Balance")) return false;
  for (let row = rowNumber - 1; row >= Math.max(1, rowNumber - 12); row -= 1) {
    const label = normalize(rowLabel(sheet, row));
    if (
      label === normalize("Retained Earnings") ||
      label === normalize("AOCI Assumptions") ||
      label === normalize("Revolver Balance") ||
      label === normalize("Total Debt Balance")
    ) {
      return true;
    }
  }
  return false;
}

function isProtectedPpeDepreciationRow(sheet: ExcelJS.Worksheet, rowNumber: number) {
  if (normalize(rowLabel(sheet, rowNumber)) !== normalize("Depreciation Expense")) return false;
  for (let row = rowNumber - 1; row >= Math.max(1, rowNumber - 8); row -= 1) {
    const label = normalize(rowLabel(sheet, row));
    if (label === normalize("PP&E / Depreciation Schedule") || label === normalize("PPE / Depreciation Schedule")) return true;
  }
  return false;
}

function isProtectedScheduleClearRow(sheet: ExcelJS.Worksheet, rowNumber: number) {
  return isProtectedBeginningBalanceRow(sheet, rowNumber) || isProtectedPpeDepreciationRow(sheet, rowNumber);
}

function staleBeginningBalanceDetector(note: string) {
  return (_period: string, existing: number | null, _ctx: ResolveContext): StaleProtectedMatch | null => {
    if (existing !== null && !Number.isFinite(existing)) return null;
    return {
      note,
      conceptsUsed: "",
      accession: ""
    };
  };
}

function staleUnsupportedScheduleDetector(note: string) {
  return (_period: string, existing: number | null, _ctx: ResolveContext): StaleProtectedMatch | null => {
    if (existing !== null && !Number.isFinite(existing)) return null;
    return {
      note,
      conceptsUsed: "",
      accession: ""
    };
  };
}

function staleGenericDetector(concepts: string[], sign: 1 | -1, scale: number, note: string) {
  return (period: string, existing: number | null, ctx: ResolveContext): StaleProtectedMatch | null => {
    const source = first(period, ctx.duration, concepts);
    if (!source) return null;
    if (existing === null) return null;
    if (Math.abs(source.value) < 0.0001 || Math.abs(existing) < 0.0001) return null;
    const expected = (sign === -1 ? -Math.abs(source.value) : Math.abs(source.value)) / scale;
    if (Math.abs(existing - expected) > 0.05) return null;
    return {
      note: `Cannot find exact schedule value in EDGAR, find manually. ${note}`,
      conceptsUsed: source.concept,
      accession: source.accn ?? ""
    };
  };
}

function clearCashFlowStatementHistoricalInputs(sheet: ExcelJS.Worksheet, periods: string[], columns: number[], auditRows: MappingAuditRow[]) {
  const cashFlowRows = cashFlowStatementRows(sheet);
  let clearedCells = 0;
  if (!cashFlowRows.length) return { clearedCells, warnings: [] };

  cashFlowRows.forEach((rowNumber) => {
    columns.forEach((col, index) => {
      const cell = sheet.getCell(rowNumber, col);
      if (hasFormula(cell)) return;
      const existing = numericCellValue(cell);
      if (existing === null) return;
      cell.value = null;
      clearedCells += 1;
      auditRows.push({
        sheetName: sheet.name,
        cell: cell.address,
        modelRowLabel: rowLabel(sheet, rowNumber) || "Cash Flow Statement",
        period: periods[index],
        valueWritten: 0,
        mappingType: "unused",
        conceptsUsed: "",
        sourceStatement: "cash flow",
        accession: "",
        sourceUrl: "",
        cellWritable: true,
        formulaPreserved: false,
        writeBlockedReason: "",
        signConvention: "not written",
        confidence: "high",
        validationStatus: "cleared",
        notes: "Cash Flow Statement historical inputs are intentionally left blank and are not mapped from EDGAR."
      });
    });
  });

  return {
    clearedCells,
    warnings: clearedCells ? [`Cash Flow Statement: cleared ${clearedCells} historical input cell(s); this section is intentionally not filled.`] : []
  };
}

function clearStaleShareRepurchaseAssumptionAmounts(sheet: ExcelJS.Worksheet, periods: string[], columns: number[], auditRows: MappingAuditRow[]) {
  let clearedCells = 0;
  let commentsAdded = 0;
  const note = "Cannot find in EDGAR, find manually.";

  for (let rowNumber = 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
    if (!isShareRepurchaseAssumptionAmountRow(sheet, rowNumber)) continue;
    const hasNegativeInput = columns.some((col) => {
      const cell = sheet.getCell(rowNumber, col);
      if (hasFormula(cell)) return false;
      const value = numericCellValue(cell);
      return value !== null && value < -0.0001;
    });
    if (!hasNegativeInput) continue;

    let clearedRow = false;
    columns.forEach((col, index) => {
      const cell = sheet.getCell(rowNumber, col);
      if (hasFormula(cell)) return;
      if (cellDisplay(cell) === "" && numericCellValue(cell) === null) return;
      cell.value = null;
      clearedCells += 1;
      clearedRow = true;
      auditRows.push({
        sheetName: sheet.name,
        cell: cell.address,
        modelRowLabel: rowLabel(sheet, rowNumber) || "Shares Repurchased ($ Amount)",
        period: periods[index],
        valueWritten: 0,
        mappingType: "unused",
        conceptsUsed: "",
        sourceStatement: "support",
        accession: "",
        sourceUrl: "",
        cellWritable: true,
        formulaPreserved: false,
        writeBlockedReason: "",
        signConvention: "not written",
        confidence: "high",
        validationStatus: "cleared",
        notes: note
      });
    });
    if (clearedRow) {
      const sourceCell = labelCell(sheet, rowNumber);
      if (canAddComment(sourceCell)) {
        sourceCell.note = "";
        addComment(sourceCell, note);
        commentsAdded += 1;
      }
    }
  }

  return {
    clearedCells,
    commentsAdded,
    warnings: clearedCells ? [`Shares Repurchased ($ Amount): cleared ${clearedCells} stale negative historical input cell(s); ${note}`] : []
  };
}

function isShareRepurchaseAssumptionAmountRow(sheet: ExcelJS.Worksheet, rowNumber: number) {
  const label = normalize(rowLabel(sheet, rowNumber));
  if (label !== normalize("Shares Repurchased ($ Amount)") && label !== normalize("Share Repurchases ($ Amount)")) return false;
  for (let row = rowNumber; row >= Math.max(1, rowNumber - 10); row -= 1) {
    if (normalize(rowLabel(sheet, row)) === normalize("Share Repurchase Assumptions")) return true;
  }
  return false;
}

function cashFlowStatementRows(sheet: ExcelJS.Worksheet) {
  const rows: number[] = [];
  let inCashFlow = false;
  for (let rowNumber = 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const label = rowLabel(sheet, rowNumber);
    if (label) {
      if (isCashFlowStatementHeader(label)) {
        inCashFlow = true;
        rows.push(rowNumber);
        continue;
      }
      if (inCashFlow && isCashFlowStatementBoundary(label)) break;
    }
    if (inCashFlow) rows.push(rowNumber);
  }
  return rows;
}

function snapshotAddress(sheet: ExcelJS.Worksheet, rowNumber: number, col: number) {
  return `${sheet.name}!${columnLetter(col)}${rowNumber}`;
}

function cellFromSnapshotAddress(workbook: ExcelJS.Workbook, address: string) {
  const bang = address.lastIndexOf("!");
  if (bang < 0) return null;
  const sheetName = address.slice(0, bang);
  const cellAddress = address.slice(bang + 1);
  const sheet = workbook.getWorksheet(sheetName);
  return sheet?.getCell(cellAddress) ?? null;
}

function validateSegmentGenericRows(sheet: ExcelJS.Worksheet, periods: string[], columns: number[]) {
  const errors: string[] = [];
  const rows = segmentMetricRows(sheet, "Total Company Revenue", "Revenue Mix", columns);
  for (const rowNumber of rows) {
    const label = segmentBaseLabel(segmentRowLabel(sheet, rowNumber, "Revenue"), "Revenue");
    if (!isGenericSegmentPlaceholder(label)) continue;
    periods.forEach((period, index) => {
      const cell = sheet.getCell(rowNumber, columns[index]);
      if (hasFormula(cell)) return;
      const value = numericCellValue(cell);
      if (value !== null && Math.abs(value) > 0.0001) {
        errors.push(`${sheet.name}!${cell.address} ${period}: generic segment row contains nonzero historical value ${value}.`);
      }
    });
  }
  return errors;
}

function validateSegmentRevenueTieOut(
  sheet: ExcelJS.Worksheet,
  periods: string[],
  columns: number[],
  ctx: ResolveContext,
  evaluator: FormulaEvaluator,
  warnings: string[]
) {
  const errors: string[] = [];
  const totalRow = findLabelRow(sheet, "Total Company Revenue");
  const rows = segmentMetricRows(sheet, "Total Company Revenue", "Revenue Mix", columns).filter((rowNumber) => {
    const label = segmentBaseLabel(segmentRowLabel(sheet, rowNumber, "Revenue"), "Revenue");
    return !isGenericSegmentPlaceholder(label);
  });
  if (!totalRow || !rows.length) return errors;

  periods.forEach((period, index) => {
    const col = columns[index];
    const segmentTotal = segmentMetricRowsTotal(sheet, rows, col, evaluator);
    const totalCell = sheet.getCell(totalRow, col);
    const sheetTotal = evaluator.evaluateCell(totalCell);
    const edgarRevenue = first(period, ctx.duration, C.revenue);
    const expectedRevenue = edgarRevenue ? edgarRevenue.value / 1_000_000 : null;
    if (sheetTotal !== null && !segmentModelRevenueTies(sheetTotal, segmentTotal)) {
      const message = `${sheet.name}!${columnLetter(col)}${totalRow} ${period}: total company revenue formula evaluates to ${roundModelValue(sheetTotal)}, but segment rows sum to ${roundModelValue(segmentTotal)}.`;
      if (hasFormula(totalCell)) warnings.unshift(message);
      else errors.push(message);
    }
    if (expectedRevenue === null) return;
    if (!segmentModelRevenueTies(segmentTotal, expectedRevenue)) {
      const message = `${sheet.name}!${columnLetter(col)}${totalRow} ${period}: segment revenue rows sum to ${roundModelValue(segmentTotal)}, but EDGAR revenue is ${roundModelValue(expectedRevenue)}.`;
      if (rows.some((rowNumber) => hasFormula(sheet.getCell(rowNumber, col)))) warnings.unshift(message);
      else errors.push(message);
    }
    if (sheetTotal !== null && !segmentModelRevenueTies(sheetTotal, expectedRevenue)) {
      const message = `${sheet.name}!${columnLetter(col)}${totalRow} ${period}: total company revenue evaluates to ${roundModelValue(sheetTotal)}, but EDGAR revenue is ${roundModelValue(expectedRevenue)}.`;
      if (hasFormula(totalCell)) warnings.unshift(message);
      else errors.push(message);
    }
  });

  return errors;
}

function validateSegmentStatementTieOut(
  sheet: ExcelJS.Worksheet,
  periods: string[],
  columns: number[],
  ctx: ResolveContext,
  evaluator: FormulaEvaluator,
  warnings: string[],
  metricName: string,
  totalLabel: string,
  endLabel: string,
  suffix: string,
  concepts: string[]
) {
  const errors: string[] = [];
  const totalRow = findLabelRow(sheet, totalLabel);
  const rows = segmentMetricRows(sheet, totalLabel, endLabel, columns).filter((rowNumber) => {
    const label = segmentBaseLabel(segmentRowLabel(sheet, rowNumber, suffix), suffix);
    return !isGenericSegmentPlaceholder(label);
  });
  if (!totalRow || !rows.length) return errors;

  periods.forEach((period, index) => {
    const col = columns[index];
    const linkedExpected = metricName === "D&A" ? linkedSegmentStatementTargetValue(sheet, totalLabel, endLabel, col, evaluator) : null;
    const edgarValue = first(period, ctx.duration, concepts);
    if (!edgarValue && linkedExpected === null) {
      warnings.unshift(`${sheet.name} ${period}: ${metricName} tie-out skipped because EDGAR ${concepts.join("/")} was unavailable.`);
      return;
    }
    const segmentTotal = segmentMetricRowsTotal(sheet, rows, col, evaluator);
    const totalCell = sheet.getCell(totalRow, col);
    const sheetTotal = evaluator.evaluateCell(totalCell);
    const expected = linkedExpected ?? (edgarValue!.value / 1_000_000);
    const sourceName = linkedExpected === null ? "EDGAR" : "the linked model total";

    if (sheetTotal !== null && !segmentStatementMetricTies(sheetTotal, segmentTotal)) {
      const message = `${sheet.name}!${columnLetter(col)}${totalRow} ${period}: ${metricName} total evaluates to ${roundModelValue(sheetTotal)}, but detail rows sum to ${roundModelValue(segmentTotal)}.`;
      if (hasFormula(totalCell)) warnings.unshift(message);
      else errors.push(message);
    }
    if (!segmentStatementMetricTies(segmentTotal, expected)) {
      const message = `${sheet.name}!${columnLetter(col)}${totalRow} ${period}: ${metricName} rows sum to ${roundModelValue(segmentTotal)}, but ${sourceName} is ${roundModelValue(expected)}.`;
      if (metricName === "D&A" && Math.abs(segmentTotal) <= 0.0001) warnings.unshift(`${message} Segment-level D&A detail appears unavailable, so this is reported as a review warning rather than a blocking error.`);
      else if (rows.some((rowNumber) => hasFormula(sheet.getCell(rowNumber, col)))) warnings.unshift(message);
      else errors.push(message);
    }
    if (sheetTotal !== null && !segmentStatementMetricTies(sheetTotal, expected)) {
      const message = `${sheet.name}!${columnLetter(col)}${totalRow} ${period}: ${metricName} total evaluates to ${roundModelValue(sheetTotal)}, but ${sourceName} is ${roundModelValue(expected)}.`;
      if (hasFormula(totalCell)) warnings.unshift(message);
      else errors.push(message);
    }
  });

  return errors;
}

function linkedSegmentStatementTargetValue(
  sheet: ExcelJS.Worksheet,
  totalLabel: string,
  checkLabel: string,
  col: number,
  evaluator: FormulaEvaluator
) {
  const totalRow = findLabelRow(sheet, totalLabel);
  const checkRow = findLabelRow(sheet, checkLabel);
  if (!totalRow || !checkRow) return null;
  const total = evaluator.evaluateCell(sheet.getCell(totalRow, col));
  const check = evaluator.evaluateCell(sheet.getCell(checkRow, col));
  return total !== null && check !== null ? total - check : null;
}

function numericCellValue(cell: ExcelJS.Cell) {
  const value = cell.value;
  if (typeof value === "number") return value;
  if (value && typeof value === "object" && "result" in value && typeof value.result === "number") return value.result;
  return null;
}

function addMappingAuditSheet(workbook: ExcelJS.Workbook, auditRows: MappingAuditRow[]) {
  const existing = workbook.getWorksheet(MAPPING_AUDIT_SHEET);
  if (existing) workbook.removeWorksheet(existing.id);
  const sheet = workbook.addWorksheet(MAPPING_AUDIT_SHEET);
  sheet.columns = [
    { header: "workbook sheet", key: "sheetName", width: 24 },
    { header: "cell/range", key: "cell", width: 12 },
    { header: "model row label", key: "modelRowLabel", width: 36 },
    { header: "section", key: "section", width: 28 },
    { header: "period", key: "period", width: 12 },
    { header: "value written", key: "valueWritten", width: 16 },
    { header: "mapping type", key: "mappingType", width: 16 },
    { header: "EDGAR concepts used", key: "conceptsUsed", width: 60 },
    { header: "SEC label(s)", key: "secLabels", width: 44 },
    { header: "source statement/table", key: "sourceStatement", width: 24 },
    { header: "accession", key: "accession", width: 24 },
    { header: "source URL", key: "sourceUrl", width: 24 },
    { header: "filing form", key: "filingForm", width: 14 },
    { header: "filed date", key: "filedDate", width: 14 },
    { header: "start date", key: "startDate", width: 14 },
    { header: "end date", key: "endDate", width: 14 },
    { header: "cell writable", key: "cellWritable", width: 14 },
    { header: "formula preserved", key: "formulaPreserved", width: 18 },
    { header: "formula status", key: "formulaStatus", width: 22 },
    { header: "write blocked reason", key: "writeBlockedReason", width: 28 },
    { header: "sign convention", key: "signConvention", width: 28 },
    { header: "confidence", key: "confidence", width: 12 },
    { header: "validation status", key: "validationStatus", width: 18 },
    { header: "notes", key: "notes", width: 60 }
  ];
  auditRows.forEach((row) => sheet.addRow(row));
  sheet.getRow(1).font = { bold: true };
  sheet.views = [{ state: "frozen", ySplit: 1 }];
}

function addComment(cell: ExcelJS.Cell, text: string) {
  const existing = commentText(cell.note);
  const body = `${existing ? `${existing}\n\n` : ""}EDGAR Mapper:\n${text}`;
  cell.note = {
    texts: [
      {
        font: { bold: true },
        text: "EDGAR Mapper:\n"
      },
      {
        text: body.replace(/^EDGAR Mapper:\n/, "")
      }
    ],
    editAs: "oneCells"
  };
}

function commentText(note: ExcelJS.Cell["note"]) {
  if (!note) return "";
  if (typeof note === "string") return note;
  return note.texts?.map((text) => text.text).join("") ?? "";
}

function roundModelValue(value: number) {
  return Math.round(value * 10) / 10;
}

function normalize(input: string) {
  return input.toLowerCase().replace(/[^a-z0-9]/g, "");
}

function unique<T>(items: T[]) {
  return Array.from(new Set(items));
}

function requestFullWorkbookRecalculation(workbook: ExcelJS.Workbook) {
  workbook.calcProperties.fullCalcOnLoad = true;
  (workbook.calcProperties as ExcelJS.Workbook["calcProperties"] & { forceFullCalc?: boolean; calcMode?: string }).forceFullCalc = true;
  (workbook.calcProperties as ExcelJS.Workbook["calcProperties"] & { forceFullCalc?: boolean; calcMode?: string }).calcMode = "auto";
}

function jsonError(error: string, status: number) {
  return NextResponse.json({ error }, { status });
}
