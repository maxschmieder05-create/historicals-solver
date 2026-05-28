import { NextRequest, NextResponse } from "next/server";
import ExcelJS from "exceljs";
import JSZip from "jszip";
import {
  COMBINED_NONCURRENT_DEBT_AND_LEASE_CONCEPTS,
  NONCURRENT_DEBT_CONCEPTS,
  NONCURRENT_LEASE_LIABILITY_CONCEPTS,
  PENSION_LIABILITY_CONCEPTS,
  TemplateMappingContext,
  buildLiabilityTemplateMappingContext,
  currentDebtBelongsInAccruedLiabilities
} from "./liability-classification";
import { loadSecBulkSupport, readSecBulkSubmissionFile, type SecBulkSupport } from "./sec-bulk";

export const runtime = "nodejs";
export const maxDuration = 900;

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

type MarketQuote = {
  currentPrice: number | null;
  fiftyTwoWeekHigh: number | null;
  fiftyTwoWeekLow: number | null;
};

type FactSource = {
  concept: string;
  label: string;
  value: number;
  note?: string;
  sourceUrl?: string;
  cik?: string;
  unit?: string;
  taxonomy?: string;
  sourceLayer?: "sec_bulk_companyfacts" | "sec_live_companyfacts" | "sec_inline_xbrl" | "derived" | "model";
  form?: string;
  fp?: string;
  filed?: string;
  accn?: string;
  start?: string;
  end?: string;
  periodKey?: string;
  periodType?: "instant" | "quarterly" | "year_to_date" | "annual";
  reportDate?: string;
  isAmendment?: boolean;
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

type FilingCommentaryEvidence = {
  period: string;
  text: string;
  topics: string[];
  sourceUrl?: string;
  form?: string;
  filed?: string;
  accn?: string;
};

type ResolveContext = {
  duration: Map<string, Map<string, FactSource>>;
  instant: Map<string, Map<string, FactSource>>;
  commentary?: Map<string, FilingCommentaryEvidence[]>;
  template?: TemplateMappingContext;
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
  | "pretax_income"
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
  family?: string;
  aggregate?: boolean;
  values: Map<string, number>;
  operatingIncome: Map<string, number>;
  depreciationAmortization: Map<string, number>;
};

type FilingRef = {
  accessionNumber: string;
  filingDate: string;
  form: string;
  primaryDocument: string;
  reportDate?: string;
  sourceLayer?: "sec_bulk_submissions" | "sec_live_submissions";
};

type FactContextOptions = {
  filingMetadata?: Map<string, FilingRef>;
  sourceLayer?: FactSource["sourceLayer"];
};

type SecWriteValidation = {
  status: "ok" | "warning" | "blocked";
  confidence: "high" | "medium" | "low";
  notes: string[];
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
  protectedCells: Map<string, ProtectedCellSnapshot>;
};

type WriteDecision = {
  writable: boolean;
  reason?: string;
  formulaPreserved: boolean;
};

type ProtectedCellSnapshot = {
  value: ExcelJS.CellValue;
  note: ExcelJS.Cell["note"];
  fingerprint: string;
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
const SEC_ARCHIVE_MIN_INTERVAL_MS = 150;
const secArchiveHtmlCache = new Map<string, string>();
let lastSecArchiveFetchAt = 0;
let secArchiveBlockedUntil = 0;

class SecArchiveRateLimitError extends Error {
  constructor() {
    super("SEC filing archive is temporarily rate limited. Please wait for the SEC cooldown and rerun so Segment Analysis can be filled from filing segment tables.");
    this.name = "SecArchiveRateLimitError";
  }
}

const TOTAL_REVENUE_CONCEPTS = [
  "Revenues",
  "RevenueFromContractWithCustomerExcludingAssessedTax",
  "SalesRevenueNet",
  "NetSales",
  "OperatingLeasesIncomeStatementLeaseRevenue",
  "RealEstateRevenueNet"
];

const REVENUE_COMPONENT_CONCEPTS = [
  "SalesRevenueGoodsNet",
  "SalesRevenueServicesNet",
  "ProductRevenue",
  "ServiceRevenue",
  "SubscriptionRevenue",
  "SubscriptionAndSupportRevenue",
  "CloudServicesAndLicenseSupportRevenue",
  "LicenseRevenue",
  "AdvertisingRevenue",
  "RevenueFromProducts",
  "RevenueFromServices"
];

const C = {
  revenue: TOTAL_REVENUE_CONCEPTS,
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
  interestIncome: ["InterestIncomeNonOperating", "InvestmentIncomeInterest"],
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
  currentInvestments: [
    "MarketableSecuritiesCurrent",
    "ShortTermInvestments",
    "OtherShortTermInvestments",
    "AvailableForSaleSecuritiesDebtSecuritiesCurrent",
    "DebtSecuritiesAvailableForSaleCurrent"
  ],
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
  intangibles: [
    "FiniteLivedIntangibleAssetsNet",
    "IntangibleAssetsNetExcludingGoodwill",
    "ContentAssetsNetNoncurrent",
    "ContentAssetsNet",
    "LicensedContentNet",
    "LicensedContentAssetsNet",
    "FilmCostsNet",
    "TelevisionProgrammingAndProductionCostsNet",
    "ProgramRightsNet"
  ],
  goodwill: ["Goodwill"],
  assets: ["Assets"],
  ap: ["AccountsPayableCurrent", "AccountsPayableTradeCurrent", "AccountsPayableAndAccruedLiabilitiesCurrent", "AccountsPayableAndAccruedLiabilitiesCurrentAndNoncurrent"],
  accrued: ["AccruedLiabilitiesCurrent", "AccruedIncomeTaxesCurrent", "EmployeeRelatedLiabilitiesCurrent"],
  customerDeposits: ["Deposits", "CustomerDeposits", "DepositsLiabilities", "InterestBearingDepositsInDomesticOffices"],
  currentLiabilities: ["LiabilitiesCurrent"],
  currentDebt: ["DebtCurrent", "LongTermDebtCurrent", "CurrentPortionOfLongTermDebt", "ShortTermBorrowings", "ShortTermBorrowingsCurrent"],
  totalDebt: [
    "DebtLongtermAndShorttermCombinedAmount",
    "DebtCurrent",
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

const INLINE_INSTANT_EXTENSION_CONCEPTS = new Set([
  ...C.intangibles,
  "ContentLiabilitiesCurrent",
  "ContentLiabilitiesNoncurrent"
]);

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
  "AccountsPayableAndAccruedLiabilitiesCurrentAndNoncurrent"
];

const NON_COMPENSATION_EXPENSE_CONCEPTS = [
  "FloorBrokerageExchangeAndClearanceFees",
  "UnderwritingCosts",
  "CommunicationsAndInformationTechnology",
  "OccupancyNet",
  "BusinessDevelopment",
  "ProfessionalFees"
];

const OTHER_OPERATING_EXPENSE_CONCEPTS = ["OtherExpenses"];

const FULFILLMENT_EXPENSE_CONCEPTS = [
  "FulfillmentExpense",
  "FulfillmentCosts",
  "FulfillmentCost",
  "FulfillmentAndShippingExpense"
];

const SALES_MARKETING_EXPENSE_CONCEPTS = [
  "SellingAndMarketingExpense",
  "SalesAndMarketingExpense",
  "MarketingExpense",
  "AdvertisingExpense"
];

const GENERAL_ADMINISTRATIVE_EXPENSE_CONCEPTS = [
  "GeneralAndAdministrativeExpense",
  "SellingGeneralAndAdministrativeExpense"
];

const TECHNOLOGY_CONTENT_RD_CONCEPTS = [
  ...C.rd,
  "TechnologyAndContentExpense",
  "TechnologyAndInfrastructureExpense",
  "TechnologyExpense",
  "TechnologyAndDevelopmentExpense",
  "ProductDevelopmentExpense",
  "ProductDevelopmentAndTechnologyExpense"
];

const INCOME_STATEMENT_DA_CONCEPTS = [
  "DepreciationAndAmortizationExpense",
  "DepreciationDepletionAndAmortizationExpense",
  "DepreciationExpense"
];

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

const CONTINUING_NET_INCOME_CONCEPTS = ["NetIncomeLoss", "ProfitLoss", "IncomeLossFromContinuingOperationsIncludingPortionAttributableToNoncontrollingInterest"];

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

const TOTAL_DEBT_AGGREGATE_CONCEPTS = [
  "DebtLongtermAndShorttermCombinedAmount",
  "LongTermDebtAndCapitalLeaseObligationsIncludingCurrentMaturities",
  "LongTermDebtAndFinanceLeaseObligations",
  "LongTermDebtAndCapitalLeaseObligations",
  "DebtAndCapitalLeaseObligations",
  "LongTermDebt"
];

const TOTAL_DEBT_INCLUDING_CURRENT_CONCEPTS = [
  "DebtLongtermAndShorttermCombinedAmount",
  "LongTermDebtAndCapitalLeaseObligationsIncludingCurrentMaturities",
  "DebtAndCapitalLeaseObligations"
];

const CASH_ROW_COMBINABLE_CURRENT_INVESTMENT_CONCEPTS = new Set([
  "AvailableForSaleSecuritiesDebtSecuritiesCurrent",
  "DebtSecuritiesAvailableForSaleCurrent",
  "OtherShortTermInvestments"
]);

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
  options: Pick<FillRow, "allowBlankHistoricalInput" | "onlyBlankHistoricalInput" | "noFillComment" | "comment"> = {}
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
  if (hasRevenue) {
    return plug(rowNumber, label, "income", "duration", resolveTotalRevenue, "direct", {
      comment: "Mapped to SEC consolidated total revenue; disaggregated product/service revenue is used only as a reviewed fallback when no consolidated total concept is available."
    });
  }
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
  if (has("Research & Development (R&D)", "Research and Development")) return plug(rowNumber, label, "income", "duration", resolveResearchDevelopmentExpense, "grouped");
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
  if (has("Depreciation & Amortization (incl. in SG&A)")) return plug(rowNumber, label, "income", "duration", resolveIncomeStatementDepreciationAmortization, "direct");
  if (has("Depreciation & Amortization", "Depreciation and Amortization", "Depreciation Expense")) {
    return plug(rowNumber, label, "income", "duration", resolveIncomeStatementDepreciationAmortization, "direct");
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
  if (has("Pre-Tax Adjustments")) return plug(rowNumber, label, "income", "duration", resolvePreTaxAdjustments, "grouped");
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
  if (has("Goodwill")) return plug(rowNumber, label, "balance", "instant", resolveGoodwill);
  if (has("Card Member Loans")) return row(rowNumber, label, "balance", "instant", C.loans);
  if (has("Investments and Assets of Consolidated VIEs", "Investments", "Investment Securities", "Investments and Assets of Consolidated Variable Interest Entities")) {
    return row(rowNumber, label, "balance", "instant", INVESTMENT_ASSET_CONCEPTS);
  }
  if (has("Other Non-Current Assets", "Other Long-Term Assets", "Other LT Assets", "Other Assets and Loans")) return plug(rowNumber, label, "balance", "instant", resolveOtherNonCurrentAssets);
  if (has("Accounts Payable", "Accounts Payable and Accrued Liabilities", "Accounts Payable & Accrued Liabilities", "Pharmacy Costs Payable")) return plug(rowNumber, label, "balance", "instant", resolveAccountsPayable);
  if (has("Securities Loaned")) return row(rowNumber, label, "balance", "instant", ["SecuritiesLoaned"]);
  if (has("Accrued Liabilities", "Accrued Expenses", "Accrued Expenses and Other", "Accrued Expenses and Other Current Liabilities")) {
    return plug(rowNumber, label, "balance", "instant", resolveAccruedLiabilities);
  }
  if (has("Customer Deposits")) return row(rowNumber, label, "balance", "instant", C.customerDeposits);
  if (has("Other Current Liabilities", "Other Current Liabs")) {
    return plug(rowNumber, label, "balance", "instant", resolveOtherCurrentLiabilities);
  }
  if (has("Tax Receivable Agreement Payables")) {
    return row(rowNumber, label, "balance", "instant", ["TaxReceivableAgreementLiability", "TaxReceivableAgreementLiabilityCurrent", "OtherLiabilitiesCurrent"]);
  }
  if (has("Short Term Borrowings", "Short-term Borrowings", "Short-Term Debt", "Short Term Debt", "Current Debt")) return row(rowNumber, label, "balance", "instant", ["OtherShortTermBorrowings", "ShortTermBorrowings", "LongTermDebtCurrent"]);
  if (has("Revolver")) return row(rowNumber, label, "balance", "instant", ["RevolvingCreditFacility", "LineOfCreditFacilityCurrentBorrowings"]);
  if (has("LT Debt (Incl. Current Portion)", "Long Term Debt", "Long-Term Debt", "Senior Notes", "Borrowings") || includes("LT Debt")) {
    return plug(rowNumber, label, "balance", "instant", resolveLongTermDebtInclCurrentPortion);
  }
  if (has("Total Debt")) {
    return plug(rowNumber, label, "balance", "instant", resolveTotalDebt);
  }
  if (has("Deferred Income Taxes")) return plug(rowNumber, label, "balance", "instant", resolveDeferredTaxLiability);
  if (has("Lease Liabilities", "Operating Lease Liabilities", "Finance Lease Liabilities", "Lease Obligations")) {
    return plug(rowNumber, label, "balance", "instant", resolveNonCurrentLeaseLiabilities);
  }
  if (has("Pension Liabilities", "Pension and Other Postretirement Liabilities", "Postretirement Liabilities")) {
    return plug(rowNumber, label, "balance", "instant", resolvePensionLiabilities);
  }
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

function resolveTotalRevenue(period: string, ctx: ResolveContext): ResolvedValue {
  const direct = first(period, ctx.duration, TOTAL_REVENUE_CONCEPTS);
  if (direct) {
    return {
      value: direct.value,
      sources: [direct],
      note: "Mapped to a consolidated SEC total revenue concept.",
      classification: "direct"
    };
  }

  const facts = ctx.duration.get(period);
  if (!facts) return { value: null, sources: [], note: "No consolidated total revenue fact was available for this period." };

  const components = REVENUE_COMPONENT_CONCEPTS.map((concept) => facts.get(concept))
    .filter((source): source is FactSource => Boolean(source))
    .filter((source) => !isSegmentLikeRevenueSource(source));
  const uniqueComponents = uniqueFactSources(components);
  if (!uniqueComponents.length) return { value: null, sources: [], note: "No consolidated or component revenue facts were available for this period." };

  if (uniqueComponents.length === 1) {
    const source = uniqueComponents[0];
    return {
      value: source.value,
      sources: [source],
      note:
        "Used the only available SEC revenue component as a fallback because no consolidated total revenue concept was reported for this period. This is kept at medium confidence for review.",
      classification: "partial"
    };
  }

  return {
    value: uniqueComponents.reduce((total, source) => total + source.value, 0),
    sources: uniqueComponents,
    note:
      "Derived consolidated revenue by summing SEC product/service/subscription revenue components because no consolidated total revenue concept was reported for this period.",
    classification: "grouped"
  };
}

function isSegmentLikeRevenueSource(source: FactSource) {
  return /segment|externalcustomer|geographic|member/i.test(`${source.concept} ${source.label}`);
}

function uniqueFactSources(sources: FactSource[]) {
  const seen = new Set<string>();
  return sources.filter((source) => {
    const key = `${source.concept}|${source.accn ?? ""}|${source.start ?? ""}|${source.end ?? ""}|${source.value}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
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
    [...FULFILLMENT_EXPENSE_CONCEPTS, ...NON_COMPENSATION_EXPENSE_CONCEPTS, ...OTHER_OPERATING_EXPENSE_CONCEPTS],
    "These were grouped because the model row is labeled Other Operating Expenses and the model does not provide separate rows for the included EDGAR expense categories."
  );
  if (detail.value !== null) return { ...(signed(detail, -1) ?? detail), classification: "grouped" };
  return resolveOtherOperatingIncomeExpense(period, ctx);
}

function resolveOtherOperatingIncomeExpense(period: string, ctx: ResolveContext): ResolvedValue {
  const operatingIncome = first(period, ctx.duration, C.operatingIncome);
  const grossProfit = first(period, ctx.duration, C.grossProfit);
  if (operatingIncome && grossProfit) {
    const sga = resolveSellingGeneralAdministrativeExpense(period, ctx);
    const rd = resolveResearchDevelopmentExpense(period, ctx);
    const da = resolveIncomeStatementDepreciationAmortization(period, ctx);
    if (sga.value !== null && rd.value !== null && da.value !== null) {
      const value = operatingIncome.value - grossProfit.value - sga.value - rd.value - da.value;
      return {
        value,
        sources: compactSources([operatingIncome, grossProfit, sga, rd, da]),
        note:
          "Calculated as the residual operating bucket needed for revenue less mapped operating expenses to reconcile to EDGAR operating income. This captures operating expense lines that do not have a dedicated model row.",
        classification: "grouped"
      };
    }
  }

  const direct = first(period, ctx.duration, ["OtherOperatingIncomeExpenseNet"]);
  return direct
    ? { value: direct.value, sources: [direct], note: "Mapped to EDGAR other operating income/expense, net.", classification: "direct" }
    : { value: null, sources: [], note: "No direct EDGAR other operating income/expense concept was available." };
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

  const component = sumWithNote(
    period,
    ctx.duration,
    [...SALES_MARKETING_EXPENSE_CONCEPTS, "GeneralAndAdministrativeExpense"],
    "Grouped from EDGAR selling/marketing and general/administrative expense concepts because no broad SG&A concept was reported."
  );
  return signed(component, -1) ?? { value: null, sources: [], note: "No SG&A concept or operating-income bridge inputs were available in SEC facts." };
}

function resolveResearchDevelopmentExpense(period: string, ctx: ResolveContext): ResolvedValue {
  const direct = first(period, ctx.duration, TECHNOLOGY_CONTENT_RD_CONCEPTS);
  if (direct) {
    return {
      value: -Math.abs(direct.value),
      sources: [direct],
      note: "Mapped to EDGAR research, development, technology, or technology/content operating expense.",
      classification: TECHNOLOGY_CONTENT_RD_CONCEPTS.includes(direct.concept) && !C.rd.includes(direct.concept) ? "grouped" : "direct"
    };
  }
  return {
    value: 0,
    sources: [zeroSource("ResearchDevelopmentTechnologyExpenseNotReported")],
    note: "No standalone R&D, technology, or technology/content operating expense line was reported for this period.",
    classification: "grouped"
  };
}

function resolveIncomeStatementDepreciationAmortization(period: string, ctx: ResolveContext): ResolvedValue {
  const direct = first(period, ctx.duration, INCOME_STATEMENT_DA_CONCEPTS);
  if (direct) {
    return {
      value: -Math.abs(direct.value),
      sources: [direct],
      note: "Mapped only to D&A concepts commonly presented as income-statement operating expense lines.",
      classification: "direct"
    };
  }

  return {
    value: 0,
    sources: [zeroSource("IncomeStatementDepreciationAmortizationNotReported")],
    note:
      "Set to zero because no standalone income-statement D&A expense line was reported. Cash-flow D&A disclosures are often embedded in cost of revenue, fulfillment, SG&A, technology/content, or other operating expense and were not double-counted above EBIT.",
    classification: "grouped"
  };
}

function resolvePreTaxAdjustments(period: string, ctx: ResolveContext): ResolvedValue {
  const direct = sumWithNote(
    period,
    ctx.duration,
    [
      "GainsLossesOnExtinguishmentOfDebt",
      "GainLossOnSaleOfBusiness",
      "GainLossOnSaleOfAssets",
      "GainLossOnDispositionOfAssets",
      "GainsLossesOnSalesOfInvestmentRealEstate"
    ],
    "Grouped from EDGAR pre-tax adjustment concepts because the model row bridges GAAP net income to adjusted net income."
  );
  if (direct.value !== null) {
    return {
      value: -direct.value,
      sources: direct.sources,
      note: "Reversed EDGAR pre-tax gains/losses so the adjusted net income bridge excludes those separately disclosed items.",
      classification: "grouped"
    };
  }

  return {
    value: 0,
    sources: [zeroSource("PreTaxAdjustmentsNotReported")],
    note: "No EDGAR-supported pre-tax adjustment item was reported for this period, so the adjusted net income bridge uses zero instead of preserving stale model hardcodes.",
    classification: "grouped"
  };
}

function resolvePostTaxAdjustments(period: string, ctx: ResolveContext): ResolvedValue {
  const direct = first(period, ctx.duration, POST_TAX_ADJUSTMENT_CONCEPTS);
  if (direct) {
    return {
      ...signed(direct, -1)!,
      note: "Mapped to EDGAR preferred-stock dividends or participating-security earnings allocations used to bridge net income to common shareholders."
    };
  }

  const preTaxAdjustment = resolvePreTaxAdjustments(period, ctx);
  const taxEffect = taxEffectForPreTaxAdjustment(period, ctx, preTaxAdjustment);
  if (taxEffect.value !== null) return taxEffect;

  return {
    value: 0,
    sources: [zeroSource("PostTaxAdjustmentsNotReported")],
    note: "No EDGAR-supported post-tax adjustment item was reported for this period, so the adjusted net income bridge uses zero instead of preserving stale model hardcodes.",
    classification: "grouped"
  };
}

function taxEffectForPreTaxAdjustment(period: string, ctx: ResolveContext, preTaxAdjustment: ResolvedValue): ResolvedValue {
  if (preTaxAdjustment.value === null || preTaxAdjustment.value === 0) return { value: null, sources: preTaxAdjustment.sources };
  const pretax = first(period, ctx.duration, PRETAX_INCOME_CONCEPTS);
  const taxExpense = resolveIncomeTaxExpense(period, ctx);
  if (!pretax || taxExpense.value === null || pretax.value === 0) {
    return {
      value: null,
      sources: compactSources([preTaxAdjustment, pretax, taxExpense]),
      note: "Could not tax-effect pre-tax adjustments because EDGAR pre-tax income or income tax expense was unavailable."
    };
  }

  const taxRate = Math.max(0, Math.min(0.5, Math.abs(taxExpense.value) / Math.abs(pretax.value)));
  const value = -preTaxAdjustment.value * taxRate;
  return {
    value,
    sources: [
      bridgeSource(period, "TaxEffectOnPreTaxAdjustments", "Tax effect on pre-tax adjustments", value, [preTaxAdjustment, taxExpense, pretax]),
      ...compactSources([preTaxAdjustment, taxExpense, pretax])
    ],
    note: "Calculated as the EDGAR tax effect on modeled pre-tax adjustments using the reported effective tax rate for the period.",
    classification: "grouped"
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
  const pretax = resolvePreTaxIncome(period, ctx);
  const continuingNet = first(period, ctx.duration, CONTINUING_NET_INCOME_CONCEPTS);
  const derived = pretax.value !== null && continuingNet
    ? {
        value: continuingNet.value - pretax.value,
        sources: [
          bridgeSource(period, "IncomeTaxExpenseBenefitDerived", "Income tax expense/benefit derived from EDGAR pre-tax income and net income", continuingNet.value - pretax.value, [pretax, continuingNet]),
          ...pretax.sources,
          continuingNet
        ],
        note: "Derived as EDGAR net income less EDGAR pre-tax income so the model's pre-tax plus tax formula reconciles to EDGAR net income.",
        classification: "grouped" as RowClassification
      }
    : null;
  if (direct) {
    if (derived && !statementMetricTies(-direct.value / 1_000_000, derived.value / 1_000_000)) {
      return {
        ...derived,
        sources: [derived.sources[0], direct, ...derived.sources.slice(1)],
        note: "Derived from EDGAR pre-tax income and net income because the direct income tax fact did not reconcile the income statement for this period."
      };
    }
    return {
      value: -direct.value,
      sources: [direct],
      note: "Mapped to EDGAR income tax expense/benefit using the model's negative-expense sign convention."
    };
  }

  if (!derived) {
    return {
      value: null,
      sources: compactSources([pretax, continuingNet]),
      note: "Could not derive income tax expense because EDGAR pre-tax income or net income was unavailable."
    };
  }

  return derived;
}

function resolvePreTaxIncome(period: string, ctx: ResolveContext): ResolvedValue {
  const direct = first(period, ctx.duration, [
    "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
    "IncomeLossFromContinuingOperationsBeforeIncomeTaxes"
  ]);
  if (direct) return { value: direct.value, sources: [direct], classification: "direct" };

  const beforeEquity = first(period, ctx.duration, [
    "IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments"
  ]);
  const equityMethod = first(period, ctx.duration, ["IncomeLossFromEquityMethodInvestments"]);
  if (beforeEquity && equityMethod) {
    const value = beforeEquity.value + equityMethod.value;
    return {
      value,
      sources: [
        bridgeSource(period, "IncomeLossFromContinuingOperationsBeforeIncomeTaxesIncludingEquityMethodInvestments", "Pre-tax income including equity-method income", value, [beforeEquity, equityMethod]),
        beforeEquity,
        equityMethod
      ],
      note: "Derived from EDGAR pre-tax income before equity-method investments plus EDGAR equity-method income so the model pre-tax and net-income bridge includes equity-method items.",
      classification: "grouped"
    };
  }

  const fallback = first(period, ctx.duration, PRETAX_INCOME_CONCEPTS);
  return fallback ? { value: fallback.value, sources: [fallback], classification: "direct" } : { value: null, sources: [] };
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
  const source: FactSource = { concept, label, value, note: label, sourceLayer: "derived", periodKey: period };
  if (!isFourthQuarterPeriod(period)) return source;
  if (inputs.length < 4) return source;

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

function resolveCurrentDebt(period: string, ctx: ResolveContext): ResolvedValue {
  const direct = reportedCurrentDebt(period, ctx);
  return direct ?? { value: 0, sources: [zeroSource(C.currentDebt[0])], note: "No separate current debt concept was reported for this period." };
}

function reportedCurrentDebt(period: string, ctx: ResolveContext): ResolvedValue | null {
  const aggregate = first(period, ctx.instant, ["DebtCurrent"]);
  if (aggregate) return { value: aggregate.value, sources: [aggregate] };

  const currentMaturities = first(period, ctx.instant, ["LongTermDebtCurrent", "CurrentPortionOfLongTermDebt"]);
  const shortTermBorrowings = first(period, ctx.instant, ["ShortTermBorrowings", "ShortTermBorrowingsCurrent"]);
  if (!currentMaturities && !shortTermBorrowings) return null;
  return {
    value: (currentMaturities?.value ?? 0) + (shortTermBorrowings?.value ?? 0),
    sources: compactSources([currentMaturities, shortTermBorrowings])
  };
}

function debtRowUsesCombinedCurrentDebt(period: string, ctx: ResolveContext) {
  if (!ctx.template?.hasDebtInclCurrentPortionRow || ctx.template.hasCurrentDebtRow === true) return false;
  return Boolean(first(period, ctx.instant, TOTAL_DEBT_INCLUDING_CURRENT_CONCEPTS));
}

function resolveTotalDebt(period: string, ctx: ResolveContext): ResolvedValue {
  const aggregate = first(period, ctx.instant, TOTAL_DEBT_AGGREGATE_CONCEPTS);
  if (aggregate) return { value: aggregate.value, sources: [aggregate] };
  return (
    sum(period, ctx.instant, [
      "ShortTermBorrowings",
      "DebtCurrent",
      "LongTermDebtCurrent",
      "LongTermDebtNoncurrent",
      "LongTermDebtAndFinanceLeaseObligationsCurrent",
      "LongTermDebtAndFinanceLeaseObligationsNoncurrent",
      "LongTermDebtAndCapitalLeaseObligationsCurrent"
    ]) ?? { value: null, sources: [] }
  );
}

function resolveLongTermDebtInclCurrentPortion(period: string, ctx: ResolveContext): ResolvedValue {
  if (debtRowUsesCombinedCurrentDebt(period, ctx)) {
    const aggregate = first(period, ctx.instant, TOTAL_DEBT_INCLUDING_CURRENT_CONCEPTS);
    if (aggregate) {
      return {
        value: aggregate.value,
        sources: [aggregate],
        note: "Used total reported debt because the template debt row includes the current portion and no separate current debt row is present."
      };
    }
  }
  const noncurrent = first(period, ctx.instant, NONCURRENT_DEBT_CONCEPTS);
  if (noncurrent) return { value: noncurrent.value, sources: [noncurrent] };
  const combinedDebtAndLease = first(period, ctx.instant, COMBINED_NONCURRENT_DEBT_AND_LEASE_CONCEPTS);
  if (combinedDebtAndLease) {
    const leases = resolveNonCurrentLeaseLiabilities(period, ctx);
    if (leases.value !== null && leases.value > 0 && combinedDebtAndLease.value >= leases.value) {
      return {
        value: combinedDebtAndLease.value - leases.value,
        sources: compactSources([combinedDebtAndLease, leases]),
        note: "Derived non-current debt from an SEC combined debt-and-lease concept less separately reported non-current lease liabilities."
      };
    }
    return { value: combinedDebtAndLease.value, sources: [combinedDebtAndLease] };
  }
  const detailed = reportedCurrentDebt(period, ctx);
  if (detailed && ctx.template?.hasCurrentDebtRow !== true) {
    return {
      value: detailed.value,
      sources: detailed.sources,
      note: "Mapped current long-term debt to the debt row only because the template did not expose a separate current debt row."
    };
  }
  const aggregate = first(period, ctx.instant, ["LongTermDebt"]);
  return aggregate ? { value: aggregate.value, sources: [aggregate] } : { value: null, sources: [] };
}

function resolveNonCurrentLeaseLiabilities(period: string, ctx: ResolveContext): ResolvedValue {
  const direct = sum(period, ctx.instant, NONCURRENT_LEASE_LIABILITY_CONCEPTS);
  return direct ?? zeroResolved(NONCURRENT_LEASE_LIABILITY_CONCEPTS[0]);
}

function resolvePensionLiabilities(period: string, ctx: ResolveContext): ResolvedValue {
  const direct = sum(period, ctx.instant, PENSION_LIABILITY_CONCEPTS);
  return direct ?? zeroResolved(PENSION_LIABILITY_CONCEPTS[0]);
}

function resolveDeferredTaxLiability(period: string, ctx: ResolveContext): ResolvedValue {
  const direct = first(period, ctx.instant, C.deferredTaxLiability);
  if (!direct) {
    return {
      value: 0,
      sources: [zeroSource(C.deferredTaxLiability[0])],
      note: "No separate deferred tax liability was reported in the SEC balance sheet for this period."
    };
  }
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
  if (!currentInvestments.sources.some((source) => CASH_ROW_COMBINABLE_CURRENT_INVESTMENT_CONCEPTS.has(source.concept))) return false;
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
  const brokerDealerReceivables = sumWithNote(
    period,
    ctx.instant,
    BROKER_DEALER_RECEIVABLES,
    "Included broker-dealer receivables, customer receivables, and fees/interest/other receivables from the SEC filing."
  );
  if (brokerDealerReceivables.value !== null) return brokerDealerReceivables;
  return {
    value: 0,
    sources: [zeroSource(C.receivables[0])],
    note: "No separate SEC accounts receivable concept was reported for this period, so stale template receivables were cleared."
  };
}

function resolveInventory(period: string, ctx: ResolveContext): ResolvedValue {
  const direct = first(period, ctx.instant, C.inventory);
  return direct ? { value: direct.value, sources: [direct] } : { value: 0, sources: [zeroSource(C.inventory[0])] };
}

function resolvePrepaidAndOtherCurrentAssets(period: string, ctx: ResolveContext): ResolvedValue {
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
  const segregatedCash =
    first(period, ctx.instant, ["CashAndSecuritiesSegregatedUnderSecuritiesExchangeCommissionRegulation"]) ??
    first(period, ctx.instant, ["CashAndSecuritiesSegregatedUnderFederalAndOtherRegulations"]);
  const brokerDealerAssets = sum(period, ctx.instant, BROKER_DEALER_CURRENT_ASSETS);
  if (segregatedCash || brokerDealerAssets) {
    return {
      value: (segregatedCash?.value ?? 0) + (brokerDealerAssets?.value ?? 0),
      sources: compactSources([segregatedCash, brokerDealerAssets]),
      note:
        "Included SEC-regulation segregated cash plus financial instruments owned, investments and loans, securities borrowed, securities purchased under agreements to resell, and securities received as collateral because no current-assets residual bridge was available.",
      classification: "partial"
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
  const excludeCurrentDebt =
    ctx.template && (!currentDebtBelongsInAccruedLiabilities(ctx.template) || debtRowUsesCombinedCurrentDebt(period, ctx));
  const currentDebt = excludeCurrentDebt ? resolveCurrentDebt(period, ctx) : zeroResolved(C.currentDebt[0]);
  if (currentLiabilities && accountsPayable.value !== null && otherCurrent.value !== null && currentDebt.value !== null) {
    return {
      value: currentLiabilities.value - accountsPayable.value - otherCurrent.value - currentDebt.value,
      sources: compactSources([currentLiabilities, accountsPayable, otherCurrent, currentDebt]),
      note: excludeCurrentDebt
        ? "Derived from SEC current liabilities less separately modeled accounts payable, other current liabilities, and current debt."
        : "Derived from SEC current liabilities less separately modeled accounts payable and other current liabilities. Current debt remains in accrued liabilities when the template does not expose a separate current debt row."
    };
  }
  const direct = first(period, ctx.instant, C.accrued);
  if (direct) return { value: direct.value, sources: [direct] };
  const fallback = first(period, ctx.instant, ["OtherAccruedLiabilitiesCurrent"]);
  return fallback ? { value: fallback.value, sources: [fallback] } : { value: null, sources: [] };
}

function resolveIntangibleAssets(period: string, ctx: ResolveContext): ResolvedValue {
  const direct = first(period, ctx.instant, C.intangibles);
  if (direct) return { value: direct.value, sources: [direct] };
  if (first(period, ctx.instant, ["PropertyPlantAndEquipmentAndOperatingLeaseRightofUseAssetAfterAccumulatedDepreciationAndAmortization"])) {
    return { value: 0, sources: [zeroSource("IntangibleAssetsNet")], note: "No separate intangible assets line item was reported in the SEC balance sheet for this period." };
  }
  return {
    value: 0,
    sources: [zeroSource(C.intangibles[0])],
    note: "No separate intangible assets line item was reported in the SEC balance sheet for this period, so stale template values were cleared."
  };
}

function resolveGoodwill(period: string, ctx: ResolveContext): ResolvedValue {
  const direct = first(period, ctx.instant, C.goodwill);
  return direct
    ? { value: direct.value, sources: [direct] }
    : {
        value: 0,
        sources: [zeroSource(C.goodwill[0])],
        note: "No separate goodwill line item was reported in the SEC balance sheet for this period, so stale template values were cleared."
      };
}

function resolveOtherNonCurrentAssets(period: string, ctx: ResolveContext): ResolvedValue {
  const otherAssetsNoncurrent = first(period, ctx.instant, ["OtherAssetsNoncurrent"]);
  const operatingLeaseAssets = sum(period, ctx.instant, ["OperatingLeaseRightOfUseAsset", "OperatingLeaseRightOfUseAssetNet"]);

  const assets = first(period, ctx.instant, C.assets);
  const currentAssets = first(period, ctx.instant, C.currentAssets);
  const ppe = resolvePpe(period, ctx);
  const intangibles = resolveIntangibleAssets(period, ctx);
  const goodwill = resolveGoodwill(period, ctx);
  if (assets && currentAssets && currentAssets.value !== null && ppe.value !== null && intangibles.value !== null && goodwill.value !== null) {
    const residual = assets.value - currentAssets.value - (ppe.value ?? 0) - intangibles.value - goodwill.value;
    const operatingLeaseAssetValue = operatingLeaseAssets?.value ?? null;
    if (otherAssetsNoncurrent && statementMetricTies(residual, otherAssetsNoncurrent.value)) {
      return {
        value: otherAssetsNoncurrent.value,
        sources: [otherAssetsNoncurrent],
        note: "Mapped to directly reported other non-current assets."
      };
    }
    if (otherAssetsNoncurrent && operatingLeaseAssets && operatingLeaseAssetValue !== null && statementMetricTies(residual, otherAssetsNoncurrent.value + operatingLeaseAssetValue)) {
      return {
        value: otherAssetsNoncurrent.value + operatingLeaseAssetValue,
        sources: compactSources([otherAssetsNoncurrent, operatingLeaseAssets]),
        note: "Mapped to directly reported other non-current assets plus operating lease right-of-use assets so the non-current asset section ties to SEC total assets."
      };
    }
    return {
      value: residual,
      sources: compactSources([assets, currentAssets, ppe, intangibles, goodwill]),
      note: "Calculated from SEC total assets less current assets and separately modeled PP&E, intangible assets, and goodwill."
    };
  }
  if (otherAssetsNoncurrent) {
    return {
      value: otherAssetsNoncurrent.value,
      sources: [otherAssetsNoncurrent],
      note: "Mapped to directly reported other non-current assets."
    };
  }
  if (operatingLeaseAssets) {
    return {
      value: operatingLeaseAssets.value,
      sources: operatingLeaseAssets.sources,
      classification: "partial",
      note: "Mapped to operating lease right-of-use assets only. This is a partial non-current asset bucket when the model row is a broad other-assets category."
    };
  }
  const broadOtherAssets = sumWithNote(
    period,
    ctx.instant,
    ["OtherAssets", "AssetsOfDisposalGroupIncludingDiscontinuedOperation"],
    "Included other assets and assets of disposal groups / discontinued operations reported in the SEC filing."
  );
  if (broadOtherAssets.value !== null) {
    return {
      ...broadOtherAssets,
      classification: "partial",
      note:
        "Mapped to broad EDGAR other-asset concepts. These can be partial for financial-company templates, so remaining section differences may stay in the model's other-asset bucket."
    };
  }
  return broadOtherAssets;
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

function resolveTotalCurrentAssets(period: string, ctx: ResolveContext): ResolvedValue {
  const direct = first(period, ctx.instant, C.currentAssets);
  if (direct) return { value: direct.value, sources: [direct] };
  return resolveCurrentAssetsFromModeledRows(period, ctx);
}

function resolveTotalNonCurrentAssets(period: string, ctx: ResolveContext): ResolvedValue {
  const assets = first(period, ctx.instant, C.assets);
  const currentAssets = resolveTotalCurrentAssets(period, ctx);
  if (!assets || currentAssets.value === null) return { value: null, sources: [] };
  return {
    value: assets.value - currentAssets.value,
    sources: compactSources([assets, currentAssets]),
    note: "Calculated from SEC total assets less SEC current assets so modeled asset sections foot to total assets."
  };
}

function resolveOtherCurrentLiabilities(period: string, ctx: ResolveContext): ResolvedValue {
  const brokerDealerLiabilities = sumWithNote(period, ctx.instant, BROKER_DEALER_OTHER_CURRENT_LIABILITIES, "Included broker-dealer current liability concepts reported in the SEC filing.");
  if (brokerDealerLiabilities.value !== null) return brokerDealerLiabilities;
  const directOtherCurrent = resolveDirectOtherCurrentLiabilities(period, ctx);
  if (directOtherCurrent.value !== null) return directOtherCurrent;

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
  const currentDebtModeledSeparately = ctx.template ? !currentDebtBelongsInAccruedLiabilities(ctx.template) || debtRowUsesCombinedCurrentDebt(period, ctx) : true;
  return difference(
    period,
    ctx.instant,
    C.currentLiabilities,
    currentDebtModeledSeparately ? [C.ap, C.accrued, C.currentDebt] : [C.ap, C.accrued],
    currentDebtModeledSeparately
      ? "Included current liabilities less separately modeled accounts payable, accrued liabilities, and current debt."
      : "Included current liabilities less separately modeled accounts payable and accrued liabilities. Current debt remains in the current-liability residual because the template does not expose a separate current debt row."
  );
}

function resolveDirectOtherCurrentLiabilities(period: string, ctx: ResolveContext): ResolvedValue {
  const direct = first(period, ctx.instant, [
    "ContentLiabilitiesCurrent",
    "ContractWithCustomerLiabilityCurrent",
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
  const liabilities = resolveTotalLiabilities(period, ctx);
  const currentLiabilities = first(period, ctx.instant, C.currentLiabilities);
  const debt = resolveLongTermDebtInclCurrentPortion(period, ctx);
  const currentDebt = debtRowUsesCombinedCurrentDebt(period, ctx) ? resolveCurrentDebt(period, ctx) : zeroResolved(C.currentDebt[0]);
  const dtl = resolveDeferredTaxLiability(period, ctx);
  const leases = ctx.template?.hasNonCurrentLeaseLiabilityRow ? resolveNonCurrentLeaseLiabilities(period, ctx) : zeroResolved(NONCURRENT_LEASE_LIABILITY_CONCEPTS[0]);
  const pension = ctx.template?.hasPensionLiabilityRow ? resolvePensionLiabilities(period, ctx) : zeroResolved(PENSION_LIABILITY_CONCEPTS[0]);
  if (
    liabilities.value !== null &&
    currentLiabilities &&
    debt.value !== null &&
    currentDebt.value !== null &&
    dtl.value !== null &&
    leases.value !== null &&
    pension.value !== null
  ) {
    const currentLiabilityExclusion = currentLiabilities.value - currentDebt.value;
    return {
      value: liabilities.value - currentLiabilityExclusion - debt.value - dtl.value - leases.value - pension.value,
      sources: compactSources([liabilities, currentLiabilities, currentDebt, debt, dtl, leases, pension]),
      note: currentDebt.value
        ? "Derived from SEC total liabilities less non-debt current liabilities and separately modeled total debt, deferred tax, lease, and pension liabilities because the template debt row includes current debt."
        : "Derived from SEC total liabilities less current liabilities and separately modeled non-current debt, deferred tax, lease, and pension liabilities. Noncontrolling interests are equity, not a liability plug."
    };
  }

  const directConcepts = ["OtherLiabilitiesNoncurrent"];
  if (!ctx.template?.hasNonCurrentLeaseLiabilityRow) directConcepts.push(...NONCURRENT_LEASE_LIABILITY_CONCEPTS);
  if (!ctx.template?.hasPensionLiabilityRow) directConcepts.push(...PENSION_LIABILITY_CONCEPTS);
  const direct = sum(period, ctx.instant, directConcepts);
  if (direct) {
    return {
      value: direct.value,
      sources: direct.sources,
      note: "Mapped to a direct SEC other non-current liability concept when reported."
    };
  }

  const assets = first(period, ctx.instant, C.assets);
  const totalEquity = resolveTotalEquityIncludingNci(period, ctx);
  if (!assets || !currentLiabilities || totalEquity.value === null || debt.value === null || currentDebt.value === null || dtl.value === null || leases.value === null || pension.value === null) {
    return { value: null, sources: [], note: "Could not calculate other non-current liabilities because assets, liabilities, or equity were unavailable." };
  }
  const currentLiabilityExclusion = currentLiabilities.value - currentDebt.value;
  return {
    value: assets.value - totalEquity.value - currentLiabilityExclusion - debt.value - dtl.value - leases.value - pension.value,
    sources: compactSources([assets, totalEquity, currentLiabilities, currentDebt, debt, dtl, leases, pension]),
    note: currentDebt.value
      ? "Included total assets less total equity, non-debt current liabilities, total debt, deferred taxes, leases, and pensions because the template debt row includes current debt."
      : "Included total assets less total equity, current liabilities, non-current debt, deferred taxes, leases, and pensions. Noncontrolling interests are included only inside total equity when EDGAR reports them."
  };
}

function resolveModeledCurrentLiabilitiesSubtotal(period: string, ctx: ResolveContext): ResolvedValue {
  const currentLiabilities = first(period, ctx.instant, C.currentLiabilities);
  if (!currentLiabilities) return { value: null, sources: [] };
  const currentDebt = debtRowUsesCombinedCurrentDebt(period, ctx) ? resolveCurrentDebt(period, ctx) : zeroResolved(C.currentDebt[0]);
  if (currentDebt.value === null) return { value: null, sources: [] };
  return {
    value: currentLiabilities.value - currentDebt.value,
    sources: compactSources([currentLiabilities, currentDebt]),
    note: currentDebt.value
      ? "Calculated as SEC current liabilities less current debt because the model's debt row includes the current portion."
      : "Mapped to SEC current liabilities."
  };
}

function resolveModeledNonCurrentLiabilitiesSubtotal(period: string, ctx: ResolveContext): ResolvedValue {
  const liabilities = resolveTotalLiabilities(period, ctx);
  const modeledCurrentLiabilities = resolveModeledCurrentLiabilitiesSubtotal(period, ctx);
  if (liabilities.value === null || modeledCurrentLiabilities.value === null) return { value: null, sources: [] };
  return {
    value: liabilities.value - modeledCurrentLiabilities.value,
    sources: compactSources([liabilities, modeledCurrentLiabilities]),
    note: "Calculated from SEC total liabilities less the model's current-liability subtotal so modeled liability sections foot to total liabilities."
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
  const equity = resolveStockholdersEquity(period, ctx);
  const retained = first(period, ctx.instant, C.retained) ?? zeroSource("RetainedEarningsAccumulatedDeficit");
  const treasury = signed(firstWithPriorInstant(period, ctx.instant, C.treasury), -1) ?? zeroSource("TreasuryStockValue");
  const aoci = first(period, ctx.instant, C.aoci) ?? zeroSource("AccumulatedOtherComprehensiveIncomeLossNetOfTax");
  if (equity.value === null) return { value: null, sources: [], note: "Could not derive common stock and APIC because stockholders' equity was unavailable." };
  return {
    value: equity.value - retained.value - treasury.value - aoci.value,
    sources: compactSources([equity, retained, treasury, aoci]),
    note: "Included stockholders' equity less retained earnings, treasury stock, and AOCI."
  };
}

function resolveStockholdersEquity(period: string, ctx: ResolveContext): ResolvedValue {
  const direct = first(period, ctx.instant, ["StockholdersEquity"]);
  if (direct) return { value: direct.value, sources: [direct] };

  const totalIncludingNci = first(period, ctx.instant, ["StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest"]);
  if (!totalIncludingNci) {
    return { value: null, sources: [], note: "No EDGAR stockholders' equity fact was available." };
  }

  const nci = first(period, ctx.instant, C.nci) ?? zeroSource("NoncontrollingInterestInConsolidatedEntity");
  return {
    value: totalIncludingNci.value - nci.value,
    sources: compactSources([totalIncludingNci, nci]),
    note: "Derived parent stockholders' equity from total equity including noncontrolling interests less EDGAR noncontrolling interests."
  };
}

function resolveTotalEquityIncludingNci(period: string, ctx: ResolveContext): ResolvedValue {
  const direct = first(period, ctx.instant, ["StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest"]);
  if (direct) return { value: direct.value, sources: [direct] };

  const stockholdersEquity = first(period, ctx.instant, ["StockholdersEquity"]);
  if (!stockholdersEquity) return { value: null, sources: [], note: "No EDGAR total equity fact was available." };

  const nci = first(period, ctx.instant, C.nci);
  if (!nci) return { value: stockholdersEquity.value, sources: [stockholdersEquity] };

  return {
    value: stockholdersEquity.value + nci.value,
    sources: [stockholdersEquity, nci],
    note: "Derived total equity by adding EDGAR noncontrolling interests to parent stockholders' equity."
  };
}

function resolveTotalLiabilities(period: string, ctx: ResolveContext): ResolvedValue {
  const direct = first(period, ctx.instant, C.liabilities);
  if (direct) return { value: direct.value, sources: [direct] };

  const assets = first(period, ctx.instant, C.assets);
  const totalEquity = resolveTotalEquityIncludingNci(period, ctx);
  if (!assets || totalEquity.value === null) {
    return { value: null, sources: compactSources([assets, totalEquity]), note: "Could not derive total liabilities because EDGAR assets or total equity were unavailable." };
  }

  return {
    value: assets.value - totalEquity.value,
    sources: compactSources([assets, totalEquity]),
    note: "Derived total liabilities from EDGAR total assets less total equity."
  };
}

function resolveTreasuryAndPreferredStock(period: string, ctx: ResolveContext): ResolvedValue {
  const treasury = signed(firstWithPriorInstant(period, ctx.instant, C.treasury), -1);
  if (treasury) return treasury;
  return { value: 0, sources: [zeroSource("TreasuryStockValue")] };
}

function resolveRetainedEarnings(period: string, ctx: ResolveContext): ResolvedValue {
  const direct = first(period, ctx.instant, C.retained);
  return direct ? { value: direct.value, sources: [direct] } : { value: 0, sources: [zeroSource(C.retained[0])] };
}

function resolveAoci(period: string, ctx: ResolveContext): ResolvedValue {
  const direct = first(period, ctx.instant, C.aoci);
  return direct ? { value: direct.value, sources: [direct] } : { value: 0, sources: [zeroSource(C.aoci[0])] };
}

function resolveNoncontrollingInterests(period: string, ctx: ResolveContext): ResolvedValue {
  const direct = first(period, ctx.instant, C.nci);
  return direct ? { value: direct.value, sources: [direct] } : { value: 0, sources: [zeroSource(C.nci[0])] };
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

async function updateCoverCompanyMetadata(workbook: ExcelJS.Workbook, company: CompanyMatch, ctx: ResolveContext, warnings: string[]) {
  const cover = workbook.getWorksheet("Cover");
  if (!cover) return;
  setCoverValue(cover, "Company Name", company.title);
  setCoverValue(cover, "Ticker", company.ticker);

  const fiscalYearEndDate = latestFiscalYearEndDate(ctx);
  if (fiscalYearEndDate) {
    const fiscalYearCell = setCoverValue(cover, "Last Fiscal Year", dateFromIsoDate(fiscalYearEndDate));
    if (fiscalYearCell) fiscalYearCell.numFmt = "m/d/yy";
  }

  const quote = await fetchMarketQuote(company.ticker).catch(() => null);
  if (!quote) {
    if (findAnyCoverRow(cover, ["Current Price", "Current Share Price", "Current Stock Price"]) || findCoverRow(cover, "52-Week High") || findCoverRow(cover, "52-Week Low")) {
      warnings.push("Cover market-price fields were left unchanged because current quote data was unavailable.");
    }
    return;
  }

  const currentPriceCell = setCoverValueAny(cover, ["Current Price", "Current Share Price", "Current Stock Price"], roundCurrency(quote.currentPrice));
  const highCell = setCoverValue(cover, "52-Week High", roundCurrency(quote.fiftyTwoWeekHigh));
  const lowCell = setCoverValue(cover, "52-Week Low", roundCurrency(quote.fiftyTwoWeekLow));
  for (const cell of [currentPriceCell, highCell, lowCell]) {
    if (cell) cell.numFmt = "$0.00";
  }
  if (currentPriceCell || highCell || lowCell) {
    warnings.push("Cover market-price fields were sourced from Yahoo Finance chart data because those fields are not available in EDGAR.");
  }
}

function setCoverValue(sheet: ExcelJS.Worksheet, label: string, value: ExcelJS.CellValue) {
  if (value === null || value === undefined) return null;
  const row = findCoverRow(sheet, label);
  if (!row) return null;
  const cell = sheet.getCell(row, 6);
  cell.value = value;
  return cell;
}

function setCoverValueAny(sheet: ExcelJS.Worksheet, labels: string[], value: ExcelJS.CellValue) {
  if (value === null || value === undefined) return null;
  const row = findAnyCoverRow(sheet, labels);
  if (!row) return null;
  const cell = sheet.getCell(row, 6);
  cell.value = value;
  return cell;
}

function findAnyCoverRow(sheet: ExcelJS.Worksheet, labels: string[]) {
  for (const label of labels) {
    const row = findCoverRow(sheet, label);
    if (row) return row;
  }
  return null;
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

function latestFiscalYearEndDate(ctx: ResolveContext) {
  const candidates: string[] = [];
  for (const periodFacts of [...ctx.duration.values(), ...ctx.instant.values()]) {
    for (const source of periodFacts.values()) {
      if (source.end && source.fp === "FY" && isTenK(source.form)) candidates.push(source.end);
    }
  }
  return candidates.sort().at(-1) ?? null;
}

function dateFromIsoDate(value: string) {
  return new Date(`${value}T00:00:00Z`);
}

function roundCurrency(value: number | null) {
  if (value === null || !Number.isFinite(value)) return null;
  return Math.round(value * 100) / 100;
}

async function fetchMarketQuote(ticker: string): Promise<MarketQuote | null> {
  const yahooSymbol = encodeURIComponent(ticker.replace(/\./g, "-"));
  const response = await fetch(`https://query1.finance.yahoo.com/v8/finance/chart/${yahooSymbol}?range=1y&interval=1d&includePrePost=false`, {
    headers: {
      "User-Agent": SEC_HEADERS["User-Agent"],
      Accept: "application/json"
    },
    cache: "no-store"
  });
  if (!response.ok) return null;

  const payload = (await response.json()) as any;
  const result = payload?.chart?.result?.[0];
  if (!result || payload?.chart?.error) return null;

  const meta = result.meta ?? {};
  const quote = result.indicators?.quote?.[0] ?? {};
  const close = finiteNumbers(quote.close).at(-1) ?? null;
  const historyHigh = maxFinite(quote.high);
  const historyLow = minFinite(quote.low);
  return {
    currentPrice: finiteNumber(meta.regularMarketPrice) ?? close,
    fiftyTwoWeekHigh: finiteNumber(meta.fiftyTwoWeekHigh) ?? historyHigh,
    fiftyTwoWeekLow: finiteNumber(meta.fiftyTwoWeekLow) ?? historyLow
  };
}

function finiteNumber(value: unknown) {
  return typeof value === "number" && Number.isFinite(value) ? value : null;
}

function finiteNumbers(values: unknown) {
  return Array.isArray(values) ? values.filter((value): value is number => typeof value === "number" && Number.isFinite(value)) : [];
}

function maxFinite(values: unknown) {
  const numbers = finiteNumbers(values);
  return numbers.length ? Math.max(...numbers) : null;
}

function minFinite(values: unknown) {
  const numbers = finiteNumbers(values);
  return numbers.length ? Math.min(...numbers) : null;
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

  if (financialScore >= 4) {
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
    { key: "revenue", statement: "income", periodType: "duration", resolver: resolveTotalRevenue, rationale: "Total company revenue normalized from consolidated EDGAR revenue concepts, with reviewed component fallback when needed." },
    { key: "net_revenue", statement: "income", periodType: "duration", concepts: C.netRevenue, rationale: "Financial-company net revenue normalized from EDGAR revenues net of interest expense." },
    { key: "cogs", statement: "income", periodType: "duration", concepts: C.cogs, rationale: "Cost of revenue normalized from EDGAR cost concepts." },
    { key: "gross_profit", statement: "income", periodType: "duration", concepts: C.grossProfit, rationale: "Gross profit normalized from EDGAR gross profit." },
    { key: "sga", statement: "income", periodType: "duration", resolver: resolveSellingGeneralAdministrativeExpense, rationale: "SG&A normalized directly or from EDGAR operating-income bridge inputs." },
    { key: "rd", statement: "income", periodType: "duration", resolver: resolveResearchDevelopmentExpense, rationale: "R&D normalized from EDGAR research, development, technology, or technology/content operating expense concepts." },
    { key: "da", statement: "income", periodType: "duration", resolver: resolveIncomeStatementDepreciationAmortization, rationale: "Income-statement D&A normalized only when a standalone income-statement D&A expense line is reported." },
    { key: "ebit", statement: "income", periodType: "duration", concepts: C.operatingIncome, rationale: "EBIT / operating income normalized from EDGAR operating income or pre-tax operating profit concepts." },
    { key: "pretax_income", statement: "income", periodType: "duration", resolver: resolvePreTaxIncome, rationale: "Pre-tax income normalized from EDGAR income before taxes concepts, including equity-method income when the available pre-tax concept excludes it." },
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

function resolveFillRowForModelPeriod(
  fillRow: FillRow,
  period: string,
  ctx: ResolveContext,
  normalized: NormalizedHistoricalsPackage
): ResolvedValue {
  const lookupPeriod = fillRow.statement === "balance" && fillRow.kind === "instant" ? balanceSheetInstantLookupPeriod(period) : period;
  return resolveRowFromPackage(fillRow, lookupPeriod, normalized) ?? resolveRow(fillRow, lookupPeriod, ctx);
}

function normalizedMetricKeyForFillRow(fillRow: FillRow, profile: TemplateProfile): NormalizedMetricKey | null {
  const label = normalize(fillRow.label);
  const concepts = new Set(fillRow.concepts ?? []);
  const hasAny = (items: string[]) => items.some((concept) => concepts.has(concept));
  if (hasAny(C.netRevenue) || label === normalize("Net Revenue") || label === normalize("Revenue Net of Interest Expense")) return "net_revenue";
  if (hasAny(C.revenue) || /^(total)?revenues?$|^sales$|^totalsales$|^netsales$|^totalnetrevenue$/.test(label)) return profile.kind === "financial_company" && hasAny(C.netRevenue) ? "net_revenue" : "revenue";
  if (hasAny(C.cogs)) return "cogs";
  if (hasAny(C.grossProfit)) return "gross_profit";
  if (hasAny(C.rd) || /researchdevelopment|technologycontent|technologyinfrastructure/.test(label)) return "rd";
  if (hasAny(C.da) || /depreciationamortization/.test(label)) return "da";
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
    const bulkSupport = await loadSecBulkSupport(company.cik, SEC_HEADERS);
    const filingMetadata = await fetchFilingMetadata(company, bulkSupport);
    const liveFacts = await fetchCompanyFacts(company.cik).catch(() => null);
    const baseFacts = bulkSupport.companyFacts ?? liveFacts;
    if (!baseFacts) throw new Error("Could not load SEC company facts for that company from bulk files or live SEC APIs.");
    const ctx = buildFactContext(baseFacts, company, {
      filingMetadata,
      sourceLayer: bulkSupport.companyFacts ? "sec_bulk_companyfacts" : "sec_live_companyfacts"
    });
    if (bulkSupport.companyFacts && liveFacts) {
      mergeContexts(
        ctx,
        buildFactContext(liveFacts, company, {
          filingMetadata,
          sourceLayer: "sec_live_companyfacts"
        })
      );
    }

    const workbook = new ExcelJS.Workbook();
    workbook.creator = "Historicals Solver";
    await workbook.xlsx.load(Buffer.from(await file.arrayBuffer()) as unknown as ExcelJS.Buffer);
    requestAutomaticWorkbookCalculation(workbook);
    normalizeSharedFormulas(workbook);
    removeInvalidConditionalFormattingRules(workbook);
    removeExternalWorkbookDefinedNames(workbook);
    const coverWarnings: string[] = [];
    await updateCoverCompanyMetadata(workbook, company, ctx, coverWarnings);

    const sheet = primaryFinancialWorksheet(workbook);
    if (!sheet) return jsonError(`Could not find a worksheet with historical income statement or balance sheet input rows in this workbook.`, 400);
    const isStandardModelSheet = sheet.name === MODEL_SHEET;
    let columns = blueColumns(sheet);
    if (!columns.length) return jsonError(`Could not find blue historical input cells in the "${sheet.name}" worksheet.`, 400);

    let periodInfos = templatePeriodInfos(sheet, columns);
    let periods: string[];
    let balanceSheetPairs: Array<{ period: string; col: number }> = [];
    const detectedPairs = historicalPeriodColumnPairs(sheet, columns, ctx);
    if (detectedPairs.length) {
      balanceSheetPairs = detectedPairs
        .filter(({ period }) => isHistoricalReportedPeriod(period, false, ctx))
        .map(({ period, col }) => ({ period, col }));
      periods = detectedPairs.map((pair) => pair.period);
      columns = detectedPairs.map((pair) => pair.col);
      periodInfos = detectedPairs.map((pair) => ({ period: pair.period, isEstimate: false }));
    } else if (periodInfos.length === columns.length) {
      const supportedPairs = periodInfos
        .map((info, index) => ({ ...info, col: columns[index] }))
        .filter(({ period, isEstimate }) => isHistoricalReportedPeriod(period, isEstimate, ctx));
      balanceSheetPairs = supportedPairs.map(({ period, col }) => ({ period, col }));
      const pairs = preferredHistoricalPairs(supportedPairs);
      periods = pairs.map((pair) => pair.period);
      columns = pairs.map((pair) => pair.col);
      periodInfos = pairs.map((pair) => ({ period: pair.period, isEstimate: false }));
    } else {
      periods = choosePeriods(ctx, columns.length);
      columns = columns.slice(0, periods.length);
      periodInfos = periods.map((period) => ({ period, isEstimate: false }));
      balanceSheetPairs = periods.map((period, index) => ({ period, col: columns[index] }));
    }
    if (!periods.length) return jsonError("SEC company facts did not include usable quarterly periods for this company.", 422);
    const balanceSheetHeaderPairs = (bestPeriodHeaderRow(sheet)?.infos ?? [])
      .filter(({ period, isEstimate }) => isHistoricalReportedPeriod(period, isEstimate, ctx))
      .map(({ period, col }) => ({ period, col }));
    balanceSheetPairs.push(...balanceSheetHeaderPairs);
    if (!balanceSheetPairs.length) balanceSheetPairs = periods.map((period, index) => ({ period, col: columns[index] }));
    balanceSheetPairs = uniquePeriodColumnPairs(balanceSheetPairs);
    const balanceSheetPeriods = balanceSheetPairs.map((pair) => pair.period);
    const balanceSheetColumns = balanceSheetPairs.map((pair) => pair.col);
    const workbookSnapshot = snapshotWorkbook(workbook, unique([sheet.name, MODEL_SHEET, SEGMENT_SHEET]), Math.min(...columns));
    if (periods.some(isQuarterPeriod)) {
      const inlineCtx = await fetchInlineFactContext(company, periods, bulkSupport);
      mergeContexts(ctx, inlineCtx);
    }
    const segmentSheet = workbook.getWorksheet(SEGMENT_SHEET);
    const segmentRevenue = segmentSheet ? await fetchSegmentRevenueByPeriod(company, periods, bulkSupport) : [];
    const profile = detectTemplateProfile(workbook, sheet, ctx);
    const normalizedPackage = buildNormalizedHistoricalsPackage(company, profile, periods, ctx, segmentRevenue);
    const fillRows = discoverFillRows(sheet, columns, periodInfos);
    if (!fillRows.length) return jsonError("Could not match the Model tab's blue input rows to supported financial statement labels.", 422);
    ctx.template = buildLiabilityTemplateMappingContext(fillRows);

    const warnings: string[] = [];
    const auditRows: MappingAuditRow[] = [];
    const llmState = createLlmMappingState();
    let filledCells = 0;
    let commentsAdded = 0;
    warnings.push(...coverWarnings);
    warnings.push(...bulkSupport.warnings);
    if (bulkSupport.latestRefreshAt) warnings.push(`SEC bulk support latest archive timestamp: ${bulkSupport.latestRefreshAt}.`);
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
        if (isProjectedBalanceSheetCell(sheet, effectiveFillRow.row, col)) continue;
        const cell = sheet.getCell(effectiveFillRow.row, col);
        const writeDecision = historicalWriteDecision(effectiveFillRow, cell);
        if (!writeDecision.writable) {
          const formula = formulaForCell(cell);
          if (formula) {
            const resolved = resolveFillRowForModelPeriod(effectiveFillRow, period, ctx, normalizedPackage);
            if (resolved.value !== null && !Number.isNaN(resolved.value)) {
              const value = resolved.value / (effectiveFillRow.scale ?? 1);
              const validation = validateResolvedValueForWrite(company, effectiveFillRow, period, resolved);
              if (validation.status === "blocked") {
                unresolved += 1;
                const note = `SEC validation blocked formula refresh: ${validation.notes.join(" ")}`;
                addComment(cell, note);
                commentsAdded += 1;
                auditRows.push(blockedMappingAuditRow(sheet, cell, effectiveFillRow, period, value, resolved, validation, writeDecision));
                continue;
              }
              const overwriteFormula = shouldOverwriteHistoricalFormulaWithDirectValue(effectiveFillRow, period, cell, value, resolved);
              if (effectiveFillRow.statement === "balance" && effectiveFillRow.kind === "instant" && !overwriteFormula) {
                continue;
              }
              const refreshedFormula = overwriteFormula
                ? null
                : formula.includes("!")
                  ? formula
                  : formulaBridgeForTargetValue(cell, value) ?? formula;
              cell.value = refreshedFormula ? { formula: refreshedFormula, result: value } : value;
              filledCells += 1;
              const note = overwriteFormula
                ? effectiveFillRow.statement === "balance"
                  ? "Replaced a balance-sheet formula with the EDGAR period-end balance so instant values are not derived from prior quarters."
                  : "Replaced a stale model bridge formula with the EDGAR-supported value."
                : "Updated the preserved formula's cached result from EDGAR while keeping the model formula in place.";
              addComment(cell, appendValidationNotes(note, validation));
              commentsAdded += 1;
              auditRows.push({
                sheetName: sheet.name,
                cell: cell.address,
                modelRowLabel: effectiveFillRow.label,
                section: effectiveFillRow.modelContext?.sectionHeader,
                period,
                valueWritten: value,
                mappingType: "formula updated",
                conceptsUsed: resolved.sources.map((source) => `${source.concept}=${roundModelValue(source.value / (effectiveFillRow.scale ?? 1))}mm`).join("; "),
                secLabels: resolved.sources.map((source) => source.label).join("; "),
                sourceStatement: effectiveFillRow.statement,
                accession: unique(resolved.sources.map((source) => source.accn ?? "").filter(Boolean)).join("; "),
                sourceUrl: unique(resolved.sources.map((source) => source.sourceUrl ?? "").filter(Boolean)).join("; "),
                filingForm: unique(resolved.sources.map((source) => source.form ?? "").filter(Boolean)).join("; "),
                filedDate: unique(resolved.sources.map((source) => source.filed ?? "").filter(Boolean)).join("; "),
                startDate: unique(resolved.sources.map((source) => source.start ?? "").filter(Boolean)).join("; "),
                endDate: unique(resolved.sources.map((source) => source.end ?? "").filter(Boolean)).join("; "),
                cellWritable: true,
                formulaPreserved: Boolean(refreshedFormula),
                formulaStatus: refreshedFormula ? (refreshedFormula === formula ? "formula cached result updated" : "formula bridge updated") : "formula replaced with EDGAR-supported value",
                writeBlockedReason: "",
                signConvention: "formula result refreshed",
                confidence: validation.confidence,
                validationStatus: validationStatusText(validation),
                notes: appendValidationNotes(note, validation)
              });
              continue;
            }
          }
          if (shouldAuditSkippedWrite(writeDecision)) {
            auditRows.push(skippedMappingAuditRow(sheet, cell, effectiveFillRow, period, writeDecision));
          }
          continue;
        }

        let resolved = resolveFillRowForModelPeriod(effectiveFillRow, period, ctx, normalizedPackage);
        if (resolved.value === null && effectiveFillRow === fillRow && fillRow.classification === "unused") {
          const llmFillRow = await llmAssistedFillRow(fillRow, company, periods, ctx, llmState);
          if (llmFillRow) {
            effectiveFillRow = llmFillRow;
            rowNotes.add(llmFillRow.comment || "LLM-assisted EDGAR concept mapping was applied after deterministic row matching did not find a value.");
            resolved = resolveFillRowForModelPeriod(effectiveFillRow, period, ctx, normalizedPackage);
          }
        }

        if (resolved.value === null || Number.isNaN(resolved.value)) {
          unresolved += 1;
          continue;
        }

        const validation = validateResolvedValueForWrite(company, effectiveFillRow, period, resolved);
        if (validation.status === "blocked") {
          unresolved += 1;
          const value = resolved.value / (effectiveFillRow.scale ?? 1);
          const note = `SEC validation blocked write: ${validation.notes.join(" ")}`;
          addComment(cell, note);
          commentsAdded += 1;
          auditRows.push(blockedMappingAuditRow(sheet, cell, effectiveFillRow, period, value, resolved, validation, writeDecision));
          rowNotes.add(note);
          continue;
        }

        cell.value = resolved.value / (effectiveFillRow.scale ?? 1);
        filledCells += 1;

        const auditNote = auditNoteForResolvedValue(effectiveFillRow, resolved, period, ctx);
        if (auditNote) rowNotes.add(auditNote);
        const mappingConfidence = effectiveFillRow.comment?.startsWith("LLM-assisted") ? "medium" : "high";
        const confidence = lowerConfidence(mappingConfidence, validation.confidence);
        const notes = appendValidationNotes(auditNote, validation);
        const cellComment = mappingComment(effectiveFillRow, resolved, period, cell.value as number, confidence, notes);
        addComment(cell, cellComment);
        commentsAdded += 1;
        const auditRow = mappingAuditRow(sheet, cell, effectiveFillRow, period, cell.value as number, resolved, confidence, notes);
        auditRow.validationStatus = validationStatusText(validation);
        auditRows.push(auditRow);
      }

      if (unresolved && fillRow.classification === "partial") {
        rowNotes.add(fillRow.noFillComment || "Split / partial match: Needs review because EDGAR detail was insufficient for one or more periods.");
      } else if (unresolved) {
        rowNotes.add("Needs review: one or more historical periods were left unchanged because no matching EDGAR fact was found.");
      }

      if (rowNotes.size) {
        const sourceCell = labelCell(sheet, fillRow.row);
        if (canAddComment(sourceCell)) {
          addComment(sourceCell, rowAnnotationSummary(Array.from(rowNotes)));
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

      const revenueFormulaResult = reconcileIncomeStatementFormulaMetricToEdgar(
        sheet,
        periods,
        columns,
        ctx,
        auditRows,
        ["Revenue", "Revenues", "Total Revenue", "Total Revenues", "Sales", "Net Sales", "Net Revenue"],
        resolveTotalRevenue,
        "revenue"
      );
      filledCells += revenueFormulaResult.filledCells;
      commentsAdded += revenueFormulaResult.commentsAdded;
      warnings.push(...revenueFormulaResult.warnings);

      const incomeFormulaResult = reconcileIncomeStatementFormulaRowsToEdgar(sheet, periods, columns, ctx, auditRows);
      filledCells += incomeFormulaResult.filledCells;
      commentsAdded += incomeFormulaResult.commentsAdded;
      warnings.push(...incomeFormulaResult.warnings);

      const pretaxFormulaResult = reconcilePreTaxIncomeFormulaRowsToEdgar(sheet, periods, columns, ctx, auditRows);
      filledCells += pretaxFormulaResult.filledCells;
      commentsAdded += pretaxFormulaResult.commentsAdded;
      warnings.push(...pretaxFormulaResult.warnings);

      const netIncomeFormulaResult = reconcileNetIncomeFormulaRowsToEdgar(sheet, periods, columns, ctx, auditRows);
      filledCells += netIncomeFormulaResult.filledCells;
      commentsAdded += netIncomeFormulaResult.commentsAdded;
      warnings.push(...netIncomeFormulaResult.warnings);

      const balanceSheetTotalResult = reconcileBalanceSheetStatementTotalsToEdgar(sheet, balanceSheetPeriods, balanceSheetColumns, ctx, auditRows);
      filledCells += balanceSheetTotalResult.filledCells;
      commentsAdded += balanceSheetTotalResult.commentsAdded;
      warnings.push(...balanceSheetTotalResult.warnings);

      const balanceSheetCheckResult = reconcileBalanceSheetCheck(sheet, balanceSheetPeriods, balanceSheetColumns, auditRows);
      filledCells += balanceSheetCheckResult.filledCells;
      commentsAdded += balanceSheetCheckResult.commentsAdded;
      warnings.push(...balanceSheetCheckResult.warnings);
    }

    const formulaCacheColumns = uniqueNumbers([...columns, ...balanceSheetColumns]);
    restoreWorkbookLabels(workbook, workbookSnapshot);
    clearStaleFormulaErrorResults(workbook);
    refreshHistoricalFormulaCachedResults(workbook, formulaCacheColumns, unique([sheet.name, MODEL_SHEET, SEGMENT_SHEET]));
    ensureFormulaDisplayCaches(workbook, formulaCacheColumns, unique([sheet.name, MODEL_SHEET, SEGMENT_SHEET]));
    refreshFinalIncomeStatementKeyMetrics(sheet, periods, columns, ctx, auditRows);
    refreshFinalBalanceSheetKeyMetrics(sheet, balanceSheetPeriods, balanceSheetColumns, ctx, auditRows);
    ensureFormulaDisplayCaches(workbook, formulaCacheColumns, unique([sheet.name, MODEL_SHEET, SEGMENT_SHEET]));
    refreshFinalIncomeStatementKeyMetrics(sheet, periods, columns, ctx, auditRows);
    const partialBalanceSheetCheckResult = reconcilePartialBalanceSheetCheck(sheet, balanceSheetPeriods, balanceSheetColumns, ctx, auditRows);
    filledCells += partialBalanceSheetCheckResult.filledCells;
    commentsAdded += partialBalanceSheetCheckResult.commentsAdded;
    warnings.push(...partialBalanceSheetCheckResult.warnings);
    ensureFormulaDisplayCaches(workbook, formulaCacheColumns, unique([sheet.name, MODEL_SHEET, SEGMENT_SHEET]));
    requestAutomaticWorkbookCalculation(workbook);
    restoreProtectedCells(workbook, workbookSnapshot);
    refreshHistoricalFormulaCachedResults(workbook, formulaCacheColumns, unique([sheet.name, MODEL_SHEET, SEGMENT_SHEET]));
    refreshFinalBalanceSheetKeyMetrics(sheet, balanceSheetPeriods, balanceSheetColumns, ctx, auditRows);
    ensureFormulaDisplayCaches(workbook, formulaCacheColumns, unique([sheet.name, MODEL_SHEET, SEGMENT_SHEET]));
    if (isStandardModelSheet) {
      const finalIncomeConsistencyResult = reconcileFinalIncomeStatementConsistency(sheet, periods, columns, ctx, auditRows);
      filledCells += finalIncomeConsistencyResult.filledCells;
      commentsAdded += finalIncomeConsistencyResult.commentsAdded;
      warnings.push(...finalIncomeConsistencyResult.warnings);
      refreshFinalIncomeStatementKeyMetrics(sheet, periods, columns, ctx, auditRows);

      const finalBalanceSheetTotalResult = reconcileBalanceSheetStatementTotalsToEdgar(sheet, balanceSheetPeriods, balanceSheetColumns, ctx, auditRows);
      filledCells += finalBalanceSheetTotalResult.filledCells;
      commentsAdded += finalBalanceSheetTotalResult.commentsAdded;
      warnings.push(...finalBalanceSheetTotalResult.warnings);
      refreshHistoricalFormulaCachedResults(workbook, formulaCacheColumns, unique([sheet.name, MODEL_SHEET, SEGMENT_SHEET]));
      refreshFinalBalanceSheetKeyMetrics(sheet, balanceSheetPeriods, balanceSheetColumns, ctx, auditRows);
      const finalPartialBalanceSheetCheckResult = reconcilePartialBalanceSheetCheck(sheet, balanceSheetPeriods, balanceSheetColumns, ctx, auditRows);
      filledCells += finalPartialBalanceSheetCheckResult.filledCells;
      commentsAdded += finalPartialBalanceSheetCheckResult.commentsAdded;
      warnings.push(...finalPartialBalanceSheetCheckResult.warnings);
      const unsupportedTotalBalanceSheetCheckResult = reconcileBalanceSheetCheck(sheet, balanceSheetPeriods, balanceSheetColumns, auditRows, (period) => {
        const lookupPeriod = balanceSheetInstantLookupPeriod(period);
        return !first(lookupPeriod, ctx.instant, C.assets);
      });
      filledCells += unsupportedTotalBalanceSheetCheckResult.filledCells;
      commentsAdded += unsupportedTotalBalanceSheetCheckResult.commentsAdded;
      warnings.push(...unsupportedTotalBalanceSheetCheckResult.warnings);
      ensureFormulaDisplayCaches(workbook, formulaCacheColumns, unique([sheet.name, MODEL_SHEET, SEGMENT_SHEET]));
    }

    const validationErrors = validateWorkbookBeforeReturn(workbook, periods, columns, ctx, warnings, sheet.name, profile, balanceSheetPeriods, balanceSheetColumns);
    validationErrors.push(...validateWorkbookPreservation(workbook, workbookSnapshot));
    if (validationErrors.length) {
      return jsonError(`Validation failed: ${validationErrors.slice(0, 6).join(" | ")}`, 422);
    }

    addMappingAuditSheet(workbook, auditRows);

    const output = await writeWorkbookBufferWithRecalculation(workbook);
    const outputName = `${company.ticker}_historicals_filled.xlsx`;
    const summary = encodeURIComponent(
      JSON.stringify({
        companyName: company.title,
        ticker: company.ticker,
        templateProfile: profile.kind,
        templateProfileConfidence: profile.confidence,
        secBulkLatestRefreshAt: bulkSupport.latestRefreshAt ?? null,
        periods,
        filledCells,
        commentsAdded,
        warnings: unique(warnings).slice(0, 8)
      })
    );

    const responseBody = output.buffer.slice(output.byteOffset, output.byteOffset + output.byteLength) as ArrayBuffer;
    return new NextResponse(responseBody, {
      headers: {
        "content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "content-disposition": `attachment; filename="${outputName}"`,
        "x-output-filename": outputName,
        "x-fill-summary": summary
      }
    });
  } catch (error) {
    console.error(error);
    if (error instanceof SecArchiveRateLimitError) return jsonError(error.message, 503);
    return jsonError(error instanceof Error ? error.message : "Unexpected fill error.", 500);
  }
}

async function findCompany(query: string): Promise<CompanyMatch> {
  const response = await fetch("https://www.sec.gov/files/company_tickers.json", { headers: SEC_HEADERS, next: { revalidate: 86_400 } });
  if (!response.ok) {
    const fallback = fallbackCompanyMatch(query);
    if (fallback) return fallback;
    throw new Error("Could not load SEC ticker directory.");
  }
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

function fallbackCompanyMatch(query: string): CompanyMatch | null {
  const normalized = normalize(query);
  const fallbackCompanies: CompanyMatch[] = [
    { cik: "0000320193", ticker: "AAPL", title: "Apple Inc." },
    { cik: "0000927653", ticker: "MCK", title: "McKesson Corporation" },
    { cik: "0000063908", ticker: "MCD", title: "McDonald's Corporation" },
    { cik: "0001018724", ticker: "AMZN", title: "Amazon.com, Inc." },
    { cik: "0000004962", ticker: "AXP", title: "American Express Company" },
    { cik: "0001065280", ticker: "NFLX", title: "Netflix, Inc." },
    { cik: "0000310764", ticker: "SYK", title: "Stryker Corporation" }
  ];
  return (
    fallbackCompanies.find((company) => normalize(company.ticker) === normalized) ??
    fallbackCompanies.find((company) => normalize(company.title) === normalized) ??
    null
  );
}

async function fetchCompanyFacts(cik: string) {
  const response = await fetch(`https://data.sec.gov/api/xbrl/companyfacts/CIK${cik}.json`, { headers: SEC_HEADERS });
  if (!response.ok) throw new Error("Could not load SEC company facts for that company.");
  return response.json();
}

async function fetchFilingMetadata(company: CompanyMatch, bulkSupport?: SecBulkSupport) {
  const filings = await fetchFilingRefs(company, 120, bulkSupport);
  return new Map(filings.map((filing) => [normalizeAccession(filing.accessionNumber), filing]));
}

async function fetchSegmentRevenueByPeriod(company: CompanyMatch, periods: string[], bulkSupport?: SecBulkSupport) {
  const filings = await fetchFilingRefs(company, filingScanLimit(periods), bulkSupport);

  const annual = new Map<string, Map<string, SegmentMetrics>>();
  const quarterly = new Map<string, Map<string, SegmentMetrics>>();

  for (const filing of filings) {
    try {
      const accession = filing.accessionNumber.replace(/-/g, "");
      const cikNoZeros = String(Number(company.cik));
      const url = `https://www.sec.gov/Archives/edgar/data/${cikNoZeros}/${accession}/${filing.primaryDocument}`;
      const html = await fetchSecArchiveHtml(url, true);
      if (!html) continue;
      const parsed = parseInlineSegmentRevenue(html, filing.form);
      for (const [period, values] of parsed.entries()) {
        if (isTenK(filing.form)) annual.set(period, values);
        else quarterly.set(period, values);
      }
    } catch (error) {
      if (error instanceof SecArchiveRateLimitError) throw error;
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
    .map((label) => {
      const metrics = firstSegmentMetricsForLabel(quarterly, label);
      return {
        label,
        family: metrics?.family,
        aggregate: metrics?.aggregate,
        values: new Map(periods.map((period) => [period, quarterly.get(period)?.get(label)?.revenue ?? 0])),
        operatingIncome: new Map(periods.map((period) => [period, quarterly.get(period)?.get(label)?.operatingIncome ?? 0])),
        depreciationAmortization: new Map(periods.map((period) => [period, quarterly.get(period)?.get(label)?.depreciationAmortization ?? 0]))
      };
    });
}

function firstSegmentMetricsForLabel(periods: Map<string, Map<string, SegmentMetrics>>, label: string) {
  for (const values of periods.values()) {
    const metrics = values.get(label);
    if (metrics) return metrics;
  }
  return null;
}

async function fetchInlineFactContext(company: CompanyMatch, periods: string[], bulkSupport?: SecBulkSupport): Promise<ResolveContext> {
  const duration = new Map<string, Map<string, FactSource>>();
  const instant = new Map<string, Map<string, FactSource>>();
  const commentary = new Map<string, FilingCommentaryEvidence[]>();
  const filings = await fetchFilingRefs(company, filingScanLimit(periods), bulkSupport);
  const wanted = new Set(periods);

  for (const filing of filings) {
    try {
      const accession = filing.accessionNumber.replace(/-/g, "");
      const cikNoZeros = String(Number(company.cik));
      const url = `https://www.sec.gov/Archives/edgar/data/${cikNoZeros}/${accession}/${filing.primaryDocument}`;
      const html = await fetchSecArchiveHtml(url);
      if (!html) continue;
      mergeInlineFacts(html, filing, wanted, { duration, instant }, url);
      mergeNarrativeOperatingExpenseFacts(html, filing, wanted, { duration, instant }, url);
      mergeFilingCommentaryEvidence(html, filing, wanted, { duration, instant, commentary }, url);
    } catch {
      // Inline facts are a supplement to SEC companyfacts; keep filling with the facts already available.
    }
  }

  return { duration, instant, commentary };
}

function filingScanLimit(periods: string[]) {
  const years = periods.map(periodYear).filter(Number.isFinite);
  const span = years.length ? Math.max(...years) - Math.min(...years) + 1 : 4;
  return Math.min(28, Math.max(12, span * 4 + 8));
}

async function fetchSecArchiveHtml(url: string, failOnRateLimit = false) {
  const cached = secArchiveHtmlCache.get(url);
  if (cached !== undefined) return cached;
  if (Date.now() < secArchiveBlockedUntil) {
    if (failOnRateLimit) throw new SecArchiveRateLimitError();
    return "";
  }

  for (let attempt = 0; attempt < 3; attempt += 1) {
    await throttleSecArchiveFetch();
    const response = await fetch(url, { headers: { ...SEC_HEADERS, Accept: "text/html" } });
    if (response.ok) {
      const html = await response.text();
      secArchiveHtmlCache.set(url, html);
      return html;
    }
    if (response.status !== 429) break;
    if (attempt === 1) {
      secArchiveBlockedUntil = Date.now() + 5 * 60_000;
      if (failOnRateLimit) throw new SecArchiveRateLimitError();
      return "";
    }
    await sleep(1_000 * (attempt + 1));
  }

  return "";
}

async function throttleSecArchiveFetch() {
  const now = Date.now();
  const waitMs = Math.max(0, SEC_ARCHIVE_MIN_INTERVAL_MS - (now - lastSecArchiveFetchAt));
  if (waitMs > 0) await sleep(waitMs);
  lastSecArchiveFetchAt = Date.now();
}

function sleep(ms: number) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function fetchFilingRefs(company: CompanyMatch, limit: number, bulkSupport?: SecBulkSupport) {
  const filings: FilingRef[] = [];
  if (bulkSupport?.submissions) {
    collectFilingRefs(bulkSupport.submissions?.filings?.recent, filings, "sec_bulk_submissions");
    for (const file of bulkSupport.submissions?.filings?.files ?? []) {
      if (filings.length >= limit) break;
      const older = await readSecBulkSubmissionFile(file.name, SEC_HEADERS).catch(() => null);
      collectFilingRefs(older, filings, "sec_bulk_submissions");
    }
  }

  try {
    const response = await fetch(`https://data.sec.gov/submissions/CIK${company.cik}.json`, { headers: SEC_HEADERS });
    if (response.ok) {
      const submissions = await response.json();
      collectFilingRefs(submissions?.filings?.recent, filings, "sec_live_submissions");

      for (const file of submissions?.filings?.files ?? []) {
        if (filings.length >= limit * 2) break;
        try {
          const older = await fetch(`https://data.sec.gov/submissions/${file.name}`, { headers: SEC_HEADERS }).then((res) => (res.ok ? res.json() : null));
          collectFilingRefs(older, filings, "sec_live_submissions");
        } catch {
          // Older filing index fetches are best effort.
        }
      }
    }
  } catch {
    // Bulk submissions are the support layer when live submissions are temporarily unavailable.
  }

  return dedupeFilingRefs(filings).slice(0, limit);
}

function collectFilingRefs(source: any, filings: FilingRef[], sourceLayer: FilingRef["sourceLayer"]) {
  if (!source?.form) return;
  source.form.forEach((form: string, index: number) => {
    const filing = {
      form,
      filingDate: source.filingDate[index],
      accessionNumber: source.accessionNumber[index],
      primaryDocument: source.primaryDocument[index],
      reportDate: source.reportDate?.[index],
      sourceLayer
    };
    if ((isTenQ(filing.form) || isTenK(filing.form)) && filing.primaryDocument?.endsWith(".htm")) {
      filings.push(filing);
    }
  });
}

function dedupeFilingRefs(filings: FilingRef[]) {
  const byAccession = new Map<string, FilingRef>();
  for (const filing of filings) {
    const key = normalizeAccession(filing.accessionNumber);
    const existing = byAccession.get(key);
    if (!existing || filingRefPreference(filing) > filingRefPreference(existing)) byAccession.set(key, filing);
  }
  return Array.from(byAccession.values()).sort((a, b) => {
    const filingDateCompare = (b.filingDate ?? "").localeCompare(a.filingDate ?? "");
    if (filingDateCompare !== 0) return filingDateCompare;
    return normalizeAccession(b.accessionNumber).localeCompare(normalizeAccession(a.accessionNumber));
  });
}

function filingRefPreference(filing: FilingRef) {
  let score = filing.sourceLayer === "sec_live_submissions" ? 2 : 1;
  if (filing.form.endsWith("/A")) score += 1;
  if (filing.reportDate) score += 1;
  return score;
}

function mergeInlineFacts(html: string, filing: FilingRef, wanted: Set<string>, ctx: ResolveContext, sourceUrl = "") {
  const contexts = parseInlineContexts(html, filing.form);
  for (const match of html.matchAll(/<ix:nonFraction\b([^>]*)>([\s\S]*?)<\/ix:nonFraction>/g)) {
    const attrs = match[1];
    const unitRef = attrs.match(/\bunitRef="([^"]+)"/)?.[1] ?? "";
    if (!/usd|shares/i.test(unitRef)) continue;
    const name = attrs.match(/\bname="([^"]+)"/)?.[1] ?? "";
    const taxonomy = name.includes(":") ? name.split(":")[0] ?? "" : "";
    const concept = name.includes(":") ? name.split(":").pop() ?? name : name;
    const contextRef = attrs.match(/\bcontextRef="([^"]+)"/)?.[1];
    if (!concept || !contextRef) continue;
    if (["AssetsCurrent", "OtherAssets"].includes(concept)) continue;
    const context = contexts.get(contextRef);
    if (!context?.period || !wanted.has(context.period)) continue;
    if (context.instant && /^(?:us-gaap|dei|srt)$/i.test(taxonomy)) continue;
    if (context.instant && !INLINE_INSTANT_EXTENSION_CONCEPTS.has(concept)) continue;
    if (context.hasDimensions) continue;
    const value = ixNumber(match[2], attrs);
    if (value === null) continue;
    const source = {
      concept,
      label: concept,
      value,
      sourceUrl,
      unit: /shares/i.test(unitRef) && !/usd/i.test(unitRef) ? "shares" : "USD",
      sourceLayer: "sec_inline_xbrl" as const,
      form: filing.form,
      fp: `Q${periodQuarter(context.period)}`,
      filed: filing.filingDate,
      accn: filing.accessionNumber,
      start: context.start,
      end: context.end,
      periodKey: context.period,
      periodType: context.instant ? "instant" as const : "quarterly" as const,
      reportDate: filing.reportDate,
      isAmendment: filing.form.endsWith("/A")
    };
    setSource(context.instant ? ctx.instant : ctx.duration, context.period, concept, source);
  }
}

function mergeNarrativeOperatingExpenseFacts(html: string, filing: FilingRef, wanted: Set<string>, ctx: ResolveContext, sourceUrl = "") {
  const periodEnd = filingPeriodEndDate(filing, html);
  const period = periodEnd ? periodKeyFromDate(periodEnd) : null;
  if (!periodEnd || !period || !wanted.has(period) || isTenK(filing.form)) return;

  const text = htmlText(html);
  const lower = text.toLowerCase();
  const start =
    lower.indexOf("operating expenses:") >= 0
      ? lower.indexOf("operating expenses:")
      : lower.indexOf("operating expenses");
  const totalIndex = lower.indexOf("total operating expenses", Math.max(0, start));
  if (start < 0 || totalIndex < 0 || totalIndex <= start) return;

  const section = text.slice(start, Math.min(text.length, totalIndex + 4000));
  const rows = [
    { concept: "CostOfGoodsAndServicesSold", label: "Cost of sales", pattern: /cost\s+of\s+sales/i },
    { concept: "FulfillmentExpense", label: "Fulfillment", pattern: /fulfillment/i },
    { concept: "TechnologyAndContentExpense", label: "Technology and content", pattern: /technology\s+and\s+content/i },
    { concept: "TechnologyAndInfrastructureExpense", label: "Technology and infrastructure", pattern: /technology\s+and\s+infrastructure/i },
    { concept: "SalesAndMarketingExpense", label: "Sales and marketing", pattern: /sales\s+and\s+marketing/i },
    { concept: "MarketingExpense", label: "Marketing", pattern: /marketing/i },
    { concept: "GeneralAndAdministrativeExpense", label: "General and administrative", pattern: /general\s+and\s+administrative/i },
    { concept: "OtherOperatingIncomeExpenseNet", label: "Other operating income (expense), net", pattern: /other\s+operating\s+(?:income|expense)(?:\s*\([^)]+\))?(?:,\s*net)?/i },
    { concept: "CostsAndExpenses", label: "Total operating expenses", pattern: /total\s+operating\s+expenses/i }
  ];

  const matches = rows
    .map((row) => {
      const match = section.match(row.pattern);
      return match?.index === undefined ? null : { ...row, index: match.index };
    })
    .filter((row): row is (typeof rows)[number] & { index: number } => Boolean(row))
    .sort((a, b) => a.index - b.index);

  for (let index = 0; index < matches.length; index += 1) {
    const current = matches[index];
    const previous = matches[index - 1];
    if (current.concept === "MarketingExpense" && previous?.concept === "SalesAndMarketingExpense" && current.index - previous.index < 24) continue;
    const next = matches[index + 1];
    const slice = section.slice(current.index, next ? next.index : section.length);
    const numbers = financialNumbers(slice);
    if (!numbers.length) continue;
    const value = operatingExpenseTableCurrentQuarterValue(numbers, period);
    if (value === null) continue;
    const source: FactSource = {
      concept: current.concept,
      label: current.label,
      value,
      sourceUrl,
      unit: "USD",
      sourceLayer: "sec_inline_xbrl",
      form: filing.form,
      filed: filing.filingDate,
      accn: filing.accessionNumber,
      end: periodEnd,
      periodKey: period,
      periodType: "quarterly",
      reportDate: filing.reportDate,
      isAmendment: filing.form.endsWith("/A")
    };
    setSource(ctx.duration, period, current.concept, source);
  }
}

function mergeFilingCommentaryEvidence(html: string, filing: FilingRef, wanted: Set<string>, ctx: ResolveContext, sourceUrl = "") {
  const period = filingCommentaryPeriod(filing, html);
  if (!period || !wanted.has(period)) return;

  const text = htmlText(html);
  const evidence = extractFilingCommentaryEvidence(text, period, filing, sourceUrl);
  if (!evidence.length) return;
  const existing = ctx.commentary?.get(period) ?? [];
  ctx.commentary?.set(period, uniqueCommentaryEvidence([...existing, ...evidence]).slice(0, 18));
}

function filingCommentaryPeriod(filing: FilingRef, html: string) {
  const focus = parseInlineFiscalFocus(html);
  if (focus) return focus.fp === "FY" ? `4Q${String(focus.fy).slice(-2)}` : `${focus.fp.slice(1)}Q${String(focus.fy).slice(-2)}`;
  const periodEnd = filingPeriodEndDate(filing, html) ?? filing.reportDate;
  return periodEnd ? periodKeyFromDate(periodEnd) : null;
}

function extractFilingCommentaryEvidence(text: string, period: string, filing: FilingRef, sourceUrl = ""): FilingCommentaryEvidence[] {
  const topics = [
    { topic: "revenue", pattern: /\b(revenue|revenues|net sales|sales revenue|subscription revenue|product revenue|service revenue)\b/i },
    { topic: "segment revenue", pattern: /\b(segment|reportable segment|geographic|product category|external customers)\b/i },
    { topic: "cost of revenue", pattern: /\b(cost of revenue|cost of sales|costs of goods|costs of services|fulfillment)\b/i },
    { topic: "operating expenses", pattern: /\b(operating expenses|selling general and administrative|sales and marketing|research and development|technology and content)\b/i },
    { topic: "depreciation and amortization", pattern: /\b(depreciation|amortization)\b/i },
    { topic: "interest", pattern: /\b(interest income|interest expense|net interest|borrowings|debt)\b/i },
    { topic: "income tax", pattern: /\b(income tax|tax provision|effective tax rate)\b/i },
    { topic: "current assets", pattern: /\b(current assets|accounts receivable|inventory|prepaid|cash and cash equivalents)\b/i },
    { topic: "current liabilities", pattern: /\b(current liabilities|accounts payable|accrued liabilities|deferred revenue|customer deposits)\b/i },
    { topic: "debt and leases", pattern: /\b(long-term debt|short-term debt|lease liabilities|finance lease|operating lease)\b/i }
  ];

  const evidence: FilingCommentaryEvidence[] = [];
  for (const item of topics) {
    const match = text.match(item.pattern);
    if (!match || match.index === undefined) continue;
    const snippet = cleanCommentarySnippet(text.slice(Math.max(0, match.index - 260), Math.min(text.length, match.index + 620)));
    if (snippet.length < 80) continue;
    evidence.push({
      period,
      text: snippet,
      topics: [item.topic],
      sourceUrl,
      form: filing.form,
      filed: filing.filingDate,
      accn: filing.accessionNumber
    });
  }
  return evidence;
}

function cleanCommentarySnippet(text: string) {
  return text
    .replace(/\s+/g, " ")
    .replace(/\bTable of Contents\b/gi, "")
    .trim()
    .slice(0, 650);
}

function uniqueCommentaryEvidence(items: FilingCommentaryEvidence[]) {
  const seen = new Set<string>();
  return items.filter((item) => {
    const key = `${item.period}|${item.accn ?? ""}|${item.topics.join(",")}|${item.text.slice(0, 120)}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function filingPeriodEndDate(filing: FilingRef, html: string) {
  const ixPeriod = html.match(/name="dei:DocumentPeriodEndDate"[^>]*>([^<]+)</i)?.[1];
  if (ixPeriod) return decodeXml(ixPeriod.trim());
  const fileDate = filing.primaryDocument.match(/(\d{4})(\d{2})(\d{2})/)?.slice(1, 4);
  return fileDate ? `${fileDate[0]}-${fileDate[1]}-${fileDate[2]}` : null;
}

function htmlText(html: string) {
  return decodeXml(
    html
      .replace(/<script[\s\S]*?<\/script>/gi, " ")
      .replace(/<style[\s\S]*?<\/style>/gi, " ")
      .replace(/<\/(?:tr|p|div|table|h\d)>/gi, " ")
      .replace(/<br\s*\/?>/gi, " ")
      .replace(/<[^>]+>/g, " ")
  )
    .replace(/\u00a0/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function financialNumbers(input: string) {
  const matches = Array.from(input.matchAll(/\(?-?\$?\s*\d[\d,]*(?:\.\d+)?\)?/g)).map((match) => match[0]);
  return matches
    .map((raw) => {
      const clean = raw.replace(/\$/g, "").replace(/\s+/g, "").replace(/,/g, "");
      const negative = /^\(/.test(clean) || /^-/.test(clean);
      const parsed = Number(clean.replace(/[()]/g, "").replace(/^-/, ""));
      return Number.isFinite(parsed) ? (negative ? -parsed : parsed) * 1_000_000 : null;
    })
    .filter((value): value is number => value !== null);
}

function operatingExpenseTableCurrentQuarterValue(numbers: number[], period: string) {
  const quarter = periodQuarter(period);
  if (numbers.length >= 4 && quarter > 1) return numbers[1];
  if (numbers.length >= 2) return numbers[1];
  return numbers[0] ?? null;
}

function parseInlineContexts(html: string, form: string) {
  const fiscalFocus = supportedInlineFactFiscalFocus(parseInlineFiscalFocus(html));
  const contexts = new Map<string, InlineContext>();
  for (const match of html.matchAll(/<xbrli:context\b[^>]*id="([^"]+)"[\s\S]*?<\/xbrli:context>/g)) {
    const body = match[0];
    const instant = body.match(/<xbrli:instant>([^<]+)<\/xbrli:instant>/)?.[1];
    if (instant) {
      contexts.set(match[1], {
        period: fiscalFocus ? periodKeyFromFiscalFocus(instant, fiscalFocus) ?? periodKeyFromDate(instant) : periodKeyFromDate(instant),
        instant: true,
        end: instant,
        hasDimensions: hasInlineDimensions(body)
      });
      continue;
    }
    const start = body.match(/<xbrli:startDate>([^<]+)<\/xbrli:startDate>/)?.[1];
    const end = body.match(/<xbrli:endDate>([^<]+)<\/xbrli:endDate>/)?.[1];
    if (start && end) {
      if (!isSupportedInlineDuration(start, end)) continue;
      contexts.set(match[1], { period: contextPeriod(body, form, fiscalFocus) ?? periodKeyFromDate(end), instant: false, start, end, hasDimensions: hasInlineDimensions(body) });
    }
  }
  return contexts;
}

function supportedInlineFactFiscalFocus(focus: InlineFiscalFocus | null) {
  if (!focus) return null;
  const fiscalYearEndMonth = fiscalYearEndMonthFromFocus(focus);
  return fiscalYearEndMonth >= 1 && fiscalYearEndMonth <= 12 ? focus : null;
}

function fiscalYearEndMonthFromFocus(focus: InlineFiscalFocus) {
  const focusDate = new Date(`${focus.end}T00:00:00Z`);
  if (Number.isNaN(focusDate.getTime())) return 0;
  const periodOffset = focus.fp === "FY" ? 0 : (4 - Number(focus.fp.slice(1))) * 3;
  return ((focusDate.getUTCMonth() + periodOffset) % 12) + 1;
}

function isSupportedInlineDuration(start: string, end: string) {
  const startDate = new Date(`${start}T00:00:00Z`);
  const endDate = new Date(`${end}T00:00:00Z`);
  if (Number.isNaN(startDate.getTime()) || Number.isNaN(endDate.getTime())) return false;
  const days = Math.round((endDate.getTime() - startDate.getTime()) / 86_400_000) + 1;
  const isQuarter = days >= 80 && days <= 100;
  const isFiscalYear = start.endsWith("-01-01") && end.endsWith("-12-31");
  return isQuarter || isFiscalYear;
}

function hasInlineDimensions(contextBody: string) {
  return /<xbrldi:(explicitMember|typedMember)\b/i.test(contextBody);
}

function mergeContexts(target: ResolveContext, source: ResolveContext) {
  source.instant.forEach((facts, period) => facts.forEach((fact, concept) => setSource(target.instant, period, concept, fact)));
  source.duration.forEach((facts, period) => facts.forEach((fact, concept) => setSource(target.duration, period, concept, fact)));
  source.commentary?.forEach((items, period) => {
    const existing = target.commentary?.get(period) ?? [];
    const merged = uniqueCommentaryEvidence([...existing, ...items]);
    if (!target.commentary) target.commentary = new Map();
    target.commentary.set(period, merged);
  });
}

type SegmentMetricKey = "revenue" | "operatingIncome" | "depreciationAmortization";

type SegmentMetrics = {
  revenue?: number;
  operatingIncome?: number;
  depreciationAmortization?: number;
  family?: string;
  aggregate?: boolean;
};

type SegmentMetricMapKey = "values" | "operatingIncome" | "depreciationAmortization";

type InlineSegmentContext = {
  period: string | null;
  members: string[];
};

type InlineFiscalFocus = {
  fy: number;
  fp: "Q1" | "Q2" | "Q3" | "Q4" | "FY";
  end: string;
};

function parseInlineSegmentRevenue(html: string, form: string) {
  const fiscalFocus = parseInlineFiscalFocus(html);
  const contexts = new Map<string, InlineSegmentContext>();
  for (const match of html.matchAll(/<xbrli:context\b[^>]*id="([^"]+)"[\s\S]*?<\/xbrli:context>/g)) {
    const body = match[0];
    const period = contextPeriod(body, form, fiscalFocus);
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
    metrics.family = metrics.family ?? segmentFamilyFromMembers(context.members, label);
    metrics.aggregate = metrics.aggregate || isAggregateSegmentMember(context.members, label);
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

function segmentMetric(concept: string): SegmentMetricKey | null {
  const local = concept.split(":").pop() ?? concept;
  if (/^(RevenueFromContractWithCustomerExcludingAssessedTax|Revenues|SalesRevenueNet)$/i.test(local)) return "revenue";
  if (/(OperatingIncomeLoss|SegmentProfitLoss|IncomeLossFromContinuingOperationsBeforeIncomeTaxes)/i.test(local)) return "operatingIncome";
  if (/(DepreciationDepletionAndAmortization|DepreciationAndAmortization|DepreciationExpense)/i.test(local)) return "depreciationAmortization";
  return null;
}

function parseInlineFiscalFocus(html: string): InlineFiscalFocus | null {
  const fy = inlineDocumentFact(html, "DocumentFiscalYearFocus");
  const fp = inlineDocumentFact(html, "DocumentFiscalPeriodFocus");
  const end = normalizeInlineDate(inlineDocumentPeriodEndDate(html) ?? inlineDocumentFact(html, "DocumentPeriodEndDate") ?? "");
  const fiscalYear = fy ? Number(fy) : NaN;
  const fiscalPeriod = fp?.toUpperCase();
  if (!Number.isFinite(fiscalYear) || !end || !["Q1", "Q2", "Q3", "Q4", "FY"].includes(fiscalPeriod ?? "")) return null;
  return { fy: fiscalYear, fp: fiscalPeriod as InlineFiscalFocus["fp"], end };
}

function inlineDocumentFact(html: string, localName: string) {
  const pattern = new RegExp(`<ix:(?:nonNumeric|nonFraction)\\b[^>]*\\bname=["']dei:${localName}["'][^>]*>([\\s\\S]*?)<\\/ix:(?:nonNumeric|nonFraction)>`, "i");
  const match = html.match(pattern);
  return match ? decodeXml(match[1].replace(/<[^>]+>/g, "").trim()) : null;
}

function inlineDocumentPeriodEndDate(html: string) {
  const tag = html.match(/<ix:nonNumeric\b[^>]*\bname=["']dei:DocumentPeriodEndDate["'][^>]*>/i);
  if (!tag || tag.index === undefined) return null;
  const body = html.slice(tag.index + tag[0].length, tag.index + tag[0].length + 600);
  const firstClose = body.indexOf("</ix:nonNumeric>");
  if (firstClose < 0) return null;
  const firstRaw = inlineText(body.slice(0, firstClose));
  if (/\b\d{4}\b/.test(firstRaw) && normalizeInlineDate(firstRaw)) return firstRaw;
  const secondClose = body.indexOf("</ix:nonNumeric>", firstClose + "</ix:nonNumeric>".length);
  return inlineText(body.slice(0, secondClose >= 0 ? secondClose : firstClose));
}

function inlineText(value: string) {
  return decodeXml(value.replace(/<[^>]+>/g, "").replace(/\s+/g, " ").trim());
}

function normalizeInlineDate(value: string) {
  if (!value) return null;
  const trimmed = value.trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) return trimmed;
  const parsed = new Date(trimmed);
  return Number.isNaN(parsed.getTime()) ? null : parsed.toISOString().slice(0, 10);
}

function contextPeriod(contextXml: string, form: string, fiscalFocus?: InlineFiscalFocus | null) {
  const start = contextXml.match(/<xbrli:startDate>([^<]+)<\/xbrli:startDate>/)?.[1];
  const end = contextXml.match(/<xbrli:endDate>([^<]+)<\/xbrli:endDate>/)?.[1];
  if (!start || !end) return null;
  const endDate = new Date(`${end}T00:00:00Z`);
  const startDate = new Date(`${start}T00:00:00Z`);
  const days = (endDate.getTime() - startDate.getTime()) / 86_400_000;
  if (isTenK(form) && days < 330) return null;
  if (!isTenK(form) && days > 115) return null;
  const fiscalPeriod = fiscalFocus ? periodKeyFromFiscalFocus(end, fiscalFocus) : null;
  if (fiscalPeriod) return fiscalPeriod;
  const year = endDate.getUTCFullYear();
  const quarter = Math.floor(endDate.getUTCMonth() / 3) + 1;
  if (isTenK(form)) return `4Q${String(year).slice(-2)}`;
  return `${quarter}Q${String(year).slice(-2)}`;
}

function periodKeyFromFiscalFocus(end: string, focus: InlineFiscalFocus) {
  const endDate = new Date(`${end}T00:00:00Z`);
  const focusDate = new Date(`${focus.end}T00:00:00Z`);
  if (Number.isNaN(endDate.getTime()) || Number.isNaN(focusDate.getTime())) return null;
  const monthDelta = (endDate.getUTCFullYear() - focusDate.getUTCFullYear()) * 12 + endDate.getUTCMonth() - focusDate.getUTCMonth();
  const quarterOffset = Math.round(monthDelta / 3);
  let quarter = focus.fp === "FY" ? 4 : Number(focus.fp.slice(1));
  let fiscalYear = focus.fy;
  quarter += quarterOffset;
  while (quarter <= 0) {
    quarter += 4;
    fiscalYear -= 1;
  }
  while (quarter > 4) {
    quarter -= 4;
    fiscalYear += 1;
  }
  return `${quarter}Q${String(fiscalYear).slice(-2)}`;
}

function segmentLabelFromMembers(members: string[]) {
  const joined = members.join(" ");
  if (/U\.?S\.?MarketMember/i.test(joined)) return "U.S. Market";
  if (/InternationalOperatedMarketsMember/i.test(joined)) return "International Operated Markets";
  if (/InternationalDevelopmentalLicensedMarketsandCorporateMember/i.test(joined)) {
    return "International Developmental Licensed Markets and Corporate";
  }
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
    if (/OperatingSegmentsMember/i.test(member) && members.some((other) => other !== member && isReportableSegmentMember(other))) return false;
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

function segmentFamilyFromMembers(members: string[], label: string) {
  const joined = members.join(" ");
  const normalizedLabel = normalize(label);
  if (
    members.some((member) => member.includes("ProductOrServiceAxis")) ||
    /ServiceLine|OtherServiceLine|ProductAndService|ProductMember|ServiceMember|IPhone|IPad|Mac|Wearables/i.test(joined) ||
    /iphone|ipad|mac|wearables|service|product/.test(normalizedLabel)
  ) {
    return "product";
  }
  if (/BusinessSegment|OperatingSegment|StatementBusinessSegments|SegmentAxis|SegmentMember/i.test(joined)) return "reportable";
  if (/Geographic|Region|Country|country:/i.test(joined)) return "geographic";
  return "other";
}

function isAggregateSegmentMember(members: string[], label: string) {
  const normalizedLabel = normalize(label);
  if (["product", "products", "operatingsegments", "reportablesegments", "total"].includes(normalizedLabel)) return true;
  const joined = members.join(" ");
  return /OperatingSegmentsMember/i.test(joined) && members.length === 1;
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
  if (/^IPhone$/i.test(local)) return "iPhone";
  if (/^IPad$/i.test(local)) return "iPad";
  if (/^Service$/i.test(local)) return "Services";
  if (/^Wearables Homeand Accessories$/i.test(local)) return "Wearables, Home and Accessories";
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
    "Americas Segment",
    "Europe Segment",
    "Greater China Segment",
    "Japan Segment",
    "Rest Of Asia Pacific Segment",
    "Corporate Non Segment",
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

function buildFactContext(payload: any, company?: CompanyMatch, options: FactContextOptions = {}): ResolveContext {
  const duration = new Map<string, Map<string, FactSource>>();
  const instant = new Map<string, Map<string, FactSource>>();
  const cumulativeDuration = new Map<string, Map<string, FactSource>>();
  const annualDuration = new Map<string, Map<string, FactSource>>();
  const fiscalYearEndMonth = inferFiscalYearEndMonth(payload);
  const useFiscalDatePeriodKeys = fiscalYearEndMonth !== null && fiscalYearEndMonth !== 12;

  for (const [taxonomyName, taxonomy] of Object.entries<any>(payload?.facts ?? {})) {
    for (const [concept, detail] of Object.entries<any>(taxonomy)) {
      const label = detail.label || concept;
      const units = detail.units ?? {};
      const [unit, unitFacts] = preferredUnitFacts(units);
      for (const fact of unitFacts) {
        if (!isUsableFact(fact)) continue;
        const instantFact = isInstantFact(fact);
        const period =
          (useFiscalDatePeriodKeys ? periodKeyForFactEndDate(fact, fiscalYearEndMonth) : null) ??
          (fact.end ? periodKeyFromDate(fact.end) : periodKey(fact));
        if (!period) continue;
        const filing = fact.accn ? options.filingMetadata?.get(normalizeAccession(fact.accn)) : undefined;
        const source = {
          concept,
          label,
          value: fact.val,
          sourceUrl: sourceUrlForAccession(company?.cik ?? payload?.cik, fact.accn),
          cik: company?.cik ?? (payload?.cik ? String(payload.cik).padStart(10, "0") : undefined),
          unit,
          taxonomy: taxonomyName,
          sourceLayer: options.sourceLayer,
          form: fact.form,
          fp: fact.fp,
          filed: fact.filed,
          accn: fact.accn,
          start: fact.start,
          end: fact.end,
          periodKey: period,
          periodType: factPeriodType(fact, instantFact),
          reportDate: filing?.reportDate,
          isAmendment: Boolean(fact.form?.endsWith("/A") || filing?.form?.endsWith("/A"))
        };
        if (!instantFact && isAnnualDurationFact(fact)) {
          setSource(annualDuration, period, concept, source);
          const annualPeriod = annualPeriodKey(fact, useFiscalDatePeriodKeys ? fiscalYearEndMonth : null);
          if (annualPeriod) setSource(duration, annualPeriod, concept, { ...source, periodKey: annualPeriod });
        } else if (!instantFact && isYearToDateFact(fact)) {
          setSource(cumulativeDuration, period, concept, source);
        } else if (instantFact || isQuarterDurationFact(fact)) {
          setSource(instantFact ? instant : duration, period, concept, source);
          if (instantFact && isAnnualInstantFact(fact)) {
            const annualPeriod = annualPeriodKey(fact, useFiscalDatePeriodKeys ? fiscalYearEndMonth : null);
            if (annualPeriod) setSource(instant, annualPeriod, concept, { ...source, periodKey: annualPeriod });
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

function normalizeAccession(accession: string) {
  return accession.replace(/-/g, "");
}

function preferredUnitFacts(units: Record<string, SecFact[]>): [string, SecFact[]] {
  for (const unit of ["USD", "shares", "USD/shares", "pure"]) {
    if (Array.isArray(units[unit])) return [unit, units[unit]];
  }
  const firstUnit = Object.entries(units).find(([, facts]) => Array.isArray(facts));
  return firstUnit ?? ["", []];
}

function factPeriodType(fact: SecFact, instantFact: boolean): FactSource["periodType"] {
  if (instantFact) return "instant";
  if (isAnnualDurationFact(fact)) return "annual";
  if (isYearToDateFact(fact)) return "year_to_date";
  return "quarterly";
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
      cumulativeFacts.forEach((source, concept) => setSource(duration, period, concept, { ...source, periodType: "quarterly" }));
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
            derivedPriorPeriods: [`1Q${year}`],
            periodType: "quarterly"
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
            derivedPriorPeriods: [`1Q${year}`, `2Q${year}`],
            periodType: "quarterly"
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
        derivedPriorPeriods: [`1Q${year}`, `2Q${year}`, `3Q${year}`],
        periodType: "quarterly"
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
  return days > 115 && days < 330 && fact.fp !== "FY";
}

function isAnnualDurationFact(fact: SecFact) {
  if (!fact.start || !fact.end || fact.fp !== "FY" || !isTenK(fact.form)) return false;
  return factDurationDays(fact) >= 330;
}

function isAnnualInstantFact(fact: SecFact) {
  return Boolean(fact.end && fact.fp === "FY" && isTenK(fact.form));
}

function annualPeriodKey(fact: SecFact, fiscalYearEndMonth?: number | null) {
  const fiscalPeriod = periodKeyForFactEndDate(fact, fiscalYearEndMonth);
  const fiscalYear = fiscalPeriod ? periodYear(fiscalPeriod) : fact.fy ?? (fact.end ? new Date(`${fact.end}T00:00:00Z`).getUTCFullYear() : null);
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
  if (fact.end && !fact.start) return true;
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

function inferFiscalYearEndMonth(payload: any) {
  const counts = new Map<number, number>();
  const taxonomies = Object.values(payload?.facts ?? {}) as any[];
  for (const taxonomy of taxonomies) {
    for (const detail of Object.values<any>(taxonomy)) {
      const units = detail.units ?? {};
      const facts: SecFact[] = units.USD ?? units.shares ?? units["USD/shares"] ?? Object.values(units)[0] ?? [];
      for (const fact of facts) {
        if (!fact.end || fact.fp !== "FY" || !isTenK(fact.form)) continue;
        const date = new Date(`${fact.end}T00:00:00Z`);
        if (Number.isNaN(date.getTime())) continue;
        const month = date.getUTCMonth() + 1;
        counts.set(month, (counts.get(month) ?? 0) + 1);
      }
    }
  }
  return Array.from(counts.entries()).sort((a, b) => b[1] - a[1])[0]?.[0] ?? null;
}

function periodKeyForFactEndDate(fact: SecFact, fiscalYearEndMonth?: number | null) {
  if (!fact.end || !fiscalYearEndMonth) return null;
  if (!["Q1", "Q2", "Q3", "Q4", "FY"].includes(fact.fp ?? "")) return null;
  return fiscalPeriodKeyFromDate(fact.end, fiscalYearEndMonth);
}

function fiscalPeriodKeyFromDate(date: string, fiscalYearEndMonth: number) {
  const parsed = new Date(`${date}T00:00:00Z`);
  if (Number.isNaN(parsed.getTime())) return null;

  let best: { period: string; distance: number } | null = null;
  const endYear = parsed.getUTCFullYear();
  for (let fiscalYear = endYear - 1; fiscalYear <= endYear + 1; fiscalYear += 1) {
    for (let quarter = 1; quarter <= 4; quarter += 1) {
      const monthIndex = fiscalYearEndMonth - 1 - (4 - quarter) * 3;
      const quarterEnd = new Date(Date.UTC(fiscalYear, monthIndex + 1, 0));
      const distance = Math.abs(parsed.getTime() - quarterEnd.getTime());
      if (!best || distance < best.distance) {
        best = { period: `${quarter}Q${String(fiscalYear).slice(-2)}`, distance };
      }
    }
  }

  return best?.period ?? null;
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
  if (source.periodKey === period) score += 6;
  if (source.reportDate && source.end) {
    const reportDistance = Math.abs(new Date(`${source.reportDate}T00:00:00Z`).getTime() - new Date(`${source.end}T00:00:00Z`).getTime()) / 86_400_000;
    if (reportDistance <= 5) score += 4;
    else if (reportDistance > 20) score -= 20;
  }
  if (source.isAmendment) score += 2;
  if (source.sourceLayer === "sec_live_companyfacts") score += 1;
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

function balanceSheetInstantLookupPeriod(period: string) {
  return isAnnualPeriod(period) ? `4Q${periodYearSuffix(period)}` : period;
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
  return bestPeriodHeaderRow(sheet)?.columns ?? [];
}

function templatePeriodInfos(sheet: ExcelJS.Worksheet, columns: number[]) {
  const header = bestPeriodHeaderRow(sheet);
  if (!header) return [];
  const byColumn = new Map(header.infos.map((info) => [info.col, info]));
  return columns.map((col) => {
    const info = byColumn.get(col);
    return info ? { period: info.period, isEstimate: info.isEstimate } : { period: "", isEstimate: false };
  });
}

function historicalPeriodColumnPairs(sheet: ExcelJS.Worksheet, columns: number[], ctx: ResolveContext) {
  const header = bestPeriodHeaderRow(sheet);
  if (!header) return [];
  const columnSet = new Set(columns);
  const reportedInfos = header.infos.filter((info) => isHistoricalReportedPeriod(info.period, info.isEstimate, ctx));
  const activeQuarterYears = new Set(
    reportedInfos
      .filter((info) => columnSet.has(info.col) && isQuarterPeriod(info.period))
      .map((info) => periodYearSuffix(info.period))
  );
  const pairs = reportedInfos
    .filter((info) => columnSet.has(info.col) || (isAnnualPeriod(info.period) && activeQuarterYears.has(periodYearSuffix(info.period))))
    .map(({ col, period }) => ({ col, period }));
  return preferredHistoricalPairs(pairs);
}

function bestPeriodHeaderRow(sheet: ExcelJS.Worksheet) {
  let best:
    | { rowNumber: number; columns: number[]; infos: Array<{ col: number; period: string; isEstimate: boolean }>; validCount: number; quarterCount: number; score: number }
    | null = null;
  for (let rowNumber = 1; rowNumber <= Math.min(sheet.rowCount, 120); rowNumber += 1) {
    const infos: Array<{ col: number; period: string; isEstimate: boolean }> = [];
    for (let col = 4; col <= Math.min(sheet.columnCount, 160); col += 1) {
      const label = cellDisplay(sheet.getCell(rowNumber, col));
      const period = normalizePeriodLabel(label);
      if (!isSupportedPeriodKey(period)) continue;
      infos.push({ col, period, isEstimate: isEstimatePeriodLabel(label) });
    }
    const validCount = infos.length;
    const quarterCount = infos.filter((info) => isQuarterPeriod(info.period)).length;
    if (!validCount) continue;
    const score = validCount + quarterCount * 3;
    if (!best || score > best.score || (score === best.score && quarterCount > best.quarterCount)) {
      best = { rowNumber, columns: infos.map((info) => info.col), infos, validCount, quarterCount, score };
    }
  }
  return best;
}

const projectedPeriodColumnsBySheet = new WeakMap<ExcelJS.Worksheet, Set<number>>();
const balanceSheetRowsBySheet = new WeakMap<ExcelJS.Worksheet, Set<number>>();

function projectedPeriodColumns(sheet: ExcelJS.Worksheet) {
  const cached = projectedPeriodColumnsBySheet.get(sheet);
  if (cached) return cached;
  const columns = new Set(bestPeriodHeaderRow(sheet)?.infos.filter((info) => info.isEstimate).map((info) => info.col) ?? []);
  projectedPeriodColumnsBySheet.set(sheet, columns);
  return columns;
}

function balanceSheetRows(sheet: ExcelJS.Worksheet) {
  const cached = balanceSheetRowsBySheet.get(sheet);
  if (cached) return cached;
  const rows = new Set(balanceSheetSectionRows(sheet));
  balanceSheetRowsBySheet.set(sheet, rows);
  return rows;
}

function isProjectedBalanceSheetCell(sheet: ExcelJS.Worksheet, rowNumber: number, col: number) {
  return projectedPeriodColumns(sheet).has(col) && balanceSheetRows(sheet).has(rowNumber);
}

function isHistoricalReportedPeriod(period: string, isEstimate: boolean, ctx: ResolveContext) {
  if (!isSupportedPeriodKey(period) || isEstimate) return false;
  if (hasReportedFinancialStatementPeriod(period, ctx)) return true;
  return isAnnualPeriod(period) && hasReportedFinancialStatementPeriod(balanceSheetInstantLookupPeriod(period), ctx);
}

function hasReportedFinancialStatementPeriod(period: string, ctx: ResolveContext) {
  const durationFacts = ctx.duration.get(period);
  if (
    durationFacts &&
    [
      ...TOTAL_REVENUE_CONCEPTS,
      ...C.cogs,
      ...C.grossProfit,
      ...C.operatingIncome,
      ...PRETAX_INCOME_CONCEPTS,
      ...CONTINUING_NET_INCOME_CONCEPTS
    ].some((concept) => durationFacts.has(concept))
  ) {
    return true;
  }

  const instantFacts = ctx.instant.get(period);
  if (instantFacts && [...C.assets, ...C.liabilities, ...C.equity].some((concept) => instantFacts.has(concept))) {
    return true;
  }

  return false;
}

function preferredHistoricalPairs<T extends { period: string }>(pairs: T[]) {
  const quarterlyPairs = pairs.filter(({ period }) => isQuarterPeriod(period));
  const annualPairs = pairs.filter(({ period }) => isAnnualPeriod(period));
  return quarterlyPairs.length ? quarterlyPairs : annualPairs;
}

function isEstimatePeriodLabel(label: string) {
  const compact = label.trim().replace(/\s+/g, "").replace(/[’']/g, "");
  return /(?:^|[^a-z])(?:E|EST|ESTIMATE)$/i.test(compact) || /(?:\d{2}|\d{4})E$/i.test(compact);
}

function normalizePeriodLabel(label: string) {
  const compact = label.trim().replace(/\s+/g, "").replace(/[’']/g, "").replace(/(?:E|EST|ESTIMATE)$/i, "");
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

function shouldOverwriteHistoricalFormulaWithDirectValue(fillRow: FillRow, _period: string, cell: ExcelJS.Cell, _value: number, resolved: ResolvedValue) {
  if (fillRow.statement === "balance" && fillRow.kind === "instant") {
    return true;
  }
  if (fillRow.statement !== "income") return false;
  if (isIncomeStatementDepreciationAmortizationFillRow(fillRow)) return true;
  const formula = formulaForCell(cell);
  if (formula && !formulaHasCellReference(formula)) return true;
  const label = normalize(fillRow.label);
  if (!/pretaxadjustments|posttaxadjustments|discontinuedoperations|noncontrollinginterest|noncontrollingincome/.test(label)) return false;
  return resolved.sources.some((source) => /NotReported|Bridge/.test(source.concept));
}

function isIncomeStatementDepreciationAmortizationFillRow(fillRow: FillRow) {
  if (fillRow.statement !== "income") return false;
  const label = normalize(fillRow.label);
  return label === "da" || /^depreciation(?:and)?amortization/.test(label) || /^depreciationexpense$/.test(label);
}

function normalizeSharedFormulas(workbook: ExcelJS.Workbook) {
  workbook.eachSheet((sheet) => {
    sheet.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        if (isProjectedBalanceSheetCell(sheet, Number(cell.row), Number(cell.col))) return;
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

function setFormulaResult(cell: ExcelJS.Cell, result: FormulaDisplayResult) {
  const value = cell.value;
  if (!value || typeof value !== "object") return;
  if (!("formula" in value) && !("sharedFormula" in value)) return;
  cell.value = { ...value, result: typeof result === "number" ? persistedFormulaResult(result) : result };
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
  const fallbackRevenue = periods.map((period) => resolveTotalRevenue(period, ctx).value ?? 0);
  const rows = segmentMetricRows(sheet, "Total Company Revenue", "Revenue Mix", columns).slice(0, 6);
  const operatingIncomeRows = segmentMetricRows(sheet, "Total Company Operating Income", "Operating Income Check", columns).slice(0, 6);
  const depreciationRows = segmentMetricRows(sheet, "Total D&A", "D&A Check", columns).slice(0, 6);
  const selectedSegments = selectSegmentFamilyForTemplate(segments, periods, ctx, Math.max(rows.length, 1));

  const usableSegments = selectedSegments.length
    ? selectedSegments
    : [
        {
          label: company.title.replace(/,?\s+INC\.?$/i, ""),
          family: "company",
          values: new Map(periods.map((period, index) => [period, fallbackRevenue[index]])),
          operatingIncome: new Map(periods.map((period) => [period, 0])),
          depreciationAmortization: new Map(periods.map((period) => [period, 0]))
        }
      ];

  if (!segments.length) {
    warnings.push("No service-line revenue facts found in recent inline XBRL filings; Segment Analysis uses one total company revenue line.");
  }

  const revenueResult = fillSegmentMetricRows(sheet, periods, columns, rows, usableSegments, "values", "Revenue", auditRows);
  const operatingIncomeSegments = hasSegmentMetricData(usableSegments, "operatingIncome", periods)
    ? usableSegments
    : selectSegmentFamilyForMetric(segments, periods, ctx, Math.max(operatingIncomeRows.length, 1), "operatingIncome", C.operatingIncome);
  const depreciationSegments = usableSegments;
  const hasOperatingIncomeSegments = hasSegmentMetricData(operatingIncomeSegments, "operatingIncome", periods);
  const hasDepreciationSegments = hasSegmentMetricData(depreciationSegments, "depreciationAmortization", periods);
  const operatingIncomeResult = hasOperatingIncomeSegments
    ? fillSegmentMetricRows(sheet, periods, columns, operatingIncomeRows, operatingIncomeSegments, "operatingIncome", "Operating Income", auditRows, {
        preserveLinkedBaseLabels: operatingIncomeSegments !== usableSegments
      })
    : clearSegmentMetricRows(sheet, periods, columns, operatingIncomeRows, "Operating Income", auditRows);
  const depreciationResult = hasDepreciationSegments
    ? fillSegmentMetricRows(sheet, periods, columns, depreciationRows, depreciationSegments, "depreciationAmortization", "D&A", auditRows, {
        preserveLinkedBaseLabels: depreciationSegments !== usableSegments
      })
    : clearSegmentMetricRows(sheet, periods, columns, depreciationRows, "D&A", auditRows);
  filledCells += revenueResult.filledCells + operatingIncomeResult.filledCells + depreciationResult.filledCells;
  commentsAdded += revenueResult.commentsAdded + operatingIncomeResult.commentsAdded + depreciationResult.commentsAdded;
  const revenueReconciliation = reconcileSegmentMetricRowsToStatementTotal(sheet, periods, columns, rows, "Revenue", revenueConcepts, ctx, auditRows, resolveTotalRevenue);
  const operatingIncomeReconciliation = hasOperatingIncomeSegments
    ? reconcileSegmentMetricRowsToStatementTotal(sheet, periods, columns, operatingIncomeRows, "Operating Income", C.operatingIncome, ctx, auditRows)
    : { filledCells: 0, commentsAdded: 0, warnings: [] };
  filledCells += revenueReconciliation.filledCells + operatingIncomeReconciliation.filledCells;
  commentsAdded += revenueReconciliation.commentsAdded + operatingIncomeReconciliation.commentsAdded;
  warnings.push(...revenueReconciliation.warnings, ...operatingIncomeReconciliation.warnings);
  filledCells += restoreSegmentTotalFormula(sheet, periods, columns, "Total Company Revenue", "Revenue Mix", auditRows);
  filledCells += restoreSegmentTotalFormula(sheet, periods, columns, "Total Company Operating Income", "Operating Income Check", auditRows);
  filledCells += restoreSegmentTotalFormula(sheet, periods, columns, "Total D&A", "D&A Check", auditRows);
  filledCells += fillSegmentTotalRow(sheet, periods, columns, usableSegments, "values", "Total Company Revenue", auditRows, ctx, C.revenue, resolveTotalRevenue);
  if (hasOperatingIncomeSegments) {
    filledCells += fillSegmentTotalRow(sheet, periods, columns, operatingIncomeSegments, "operatingIncome", "Total Company Operating Income", auditRows, ctx, C.operatingIncome);
  }
  if (hasDepreciationSegments) {
    filledCells += fillSegmentTotalRow(sheet, periods, columns, depreciationSegments, "depreciationAmortization", "Total D&A", auditRows);
  }

  return { filledCells, commentsAdded, warnings };
}

function hasSegmentMetricData(
  segments: SegmentRevenue[],
  metric: SegmentMetricMapKey,
  periods: string[]
) {
  return segments.some((segment) => segmentHasMetricData(segment, metric, periods));
}

function clearSegmentMetricRows(
  sheet: ExcelJS.Worksheet,
  periods: string[],
  columns: number[],
  rows: number[],
  suffix: string,
  auditRows: MappingAuditRow[]
) {
  return rows.reduce(
    (total, rowNumber) => {
      const result = clearUnmatchedSegmentRow(sheet, rowNumber, periods, columns, segmentRowLabel(sheet, rowNumber, suffix), auditRows);
      return { filledCells: total.filledCells + result.filledCells, commentsAdded: total.commentsAdded + result.commentsAdded };
    },
    { filledCells: 0, commentsAdded: 0 }
  );
}

function selectSegmentFamilyForTemplate(
  segments: SegmentRevenue[],
  periods: string[],
  ctx: ResolveContext,
  maxRows: number
) {
  if (!segments.length) return [];
  const candidates = segmentFamilyCandidates(segments, maxRows);
  const scored = candidates
    .map((candidate) => ({ candidate, score: scoreSegmentFamilyCandidate(candidate.segments, periods, ctx, candidate.family) }))
    .filter(({ score }) => score.tieCount > 0 && score.badCount === 0)
    .sort((a, b) => {
      if (b.score.tieCount !== a.score.tieCount) return b.score.tieCount - a.score.tieCount;
      if (b.score.familyPreference !== a.score.familyPreference) return b.score.familyPreference - a.score.familyPreference;
      if (b.score.detailRows !== a.score.detailRows) return b.score.detailRows - a.score.detailRows;
      return a.score.totalError - b.score.totalError;
    });
  return scored[0]?.candidate.segments ?? candidates[0]?.segments ?? segments.filter((segment) => !segment.aggregate).slice(0, maxRows);
}

function selectSegmentFamilyForMetric(
  segments: SegmentRevenue[],
  periods: string[],
  ctx: ResolveContext,
  maxRows: number,
  metric: SegmentMetricMapKey,
  statementConcepts: string[]
) {
  const metricSegments = segments.filter((segment) => segmentHasMetricData(segment, metric, periods));
  if (!metricSegments.length) return [];
  const candidates = segmentFamilyCandidates(metricSegments, maxRows);
  const scored = candidates
    .map((candidate) => ({
      candidate,
      score: scoreSegmentMetricCandidate(candidate.segments, periods, ctx, candidate.family, metric, statementConcepts)
    }))
    .sort((a, b) => {
      if (b.score.tieCount !== a.score.tieCount) return b.score.tieCount - a.score.tieCount;
      if (b.score.dataPoints !== a.score.dataPoints) return b.score.dataPoints - a.score.dataPoints;
      if (b.score.familyPreference !== a.score.familyPreference) return b.score.familyPreference - a.score.familyPreference;
      if (b.score.detailRows !== a.score.detailRows) return b.score.detailRows - a.score.detailRows;
      return a.score.totalError - b.score.totalError;
    });
  return scored[0]?.candidate.segments ?? candidates[0]?.segments ?? metricSegments.filter((segment) => !segment.aggregate).slice(0, maxRows);
}

function segmentFamilyCandidates(segments: SegmentRevenue[], maxRows: number) {
  const byFamily = new Map<string, SegmentRevenue[]>();
  segments.forEach((segment) => {
    const family = segment.family ?? "other";
    const existing = byFamily.get(family) ?? [];
    existing.push(segment);
    byFamily.set(family, existing);
  });

  const candidates: Array<{ family: string; segments: SegmentRevenue[] }> = [];
  byFamily.forEach((familySegments, family) => {
    const sorted = familySegments.slice().sort((a, b) => segmentSort(a.label, b.label));

    const withoutAggregates = sorted.filter((segment) => !segment.aggregate);
    if (withoutAggregates.length >= 2 && withoutAggregates.length < sorted.length && withoutAggregates.length <= maxRows) {
      candidates.push({ family, segments: withoutAggregates });
    } else if (sorted.length <= maxRows) {
      candidates.push({ family, segments: sorted });
    }
  });

  if (segments.length <= maxRows) candidates.push({ family: "all", segments: segments.slice().sort((a, b) => segmentSort(a.label, b.label)) });
  return candidates;
}

function scoreSegmentMetricCandidate(
  segments: SegmentRevenue[],
  periods: string[],
  ctx: ResolveContext,
  family: string,
  metric: SegmentMetricMapKey,
  statementConcepts: string[]
) {
  let dataPoints = 0;
  let tieCount = 0;
  let totalError = 0;

  periods.forEach((period) => {
    const actual = segments.reduce((sum, segment) => sum + (segment[metric].get(period) ?? 0), 0);
    if (Math.abs(actual) <= 0.0001) return;
    dataPoints += 1;

    const expected = statementConcepts.length ? first(period, ctx.duration, statementConcepts)?.value : undefined;
    if (expected === undefined) return;
    const error = Math.abs(actual - expected);
    totalError += error;
    if (segmentStatementMetricTies(actual / 1_000_000, expected / 1_000_000)) tieCount += 1;
  });

  return {
    dataPoints,
    tieCount,
    totalError,
    detailRows: segments.filter((segment) => !segment.aggregate).length,
    familyPreference: segmentFamilyPreference(family)
  };
}

function scoreSegmentFamilyCandidate(segments: SegmentRevenue[], periods: string[], ctx: ResolveContext, family: string) {
  let tieCount = 0;
  let badCount = 0;
  let totalError = 0;

  periods.forEach((period) => {
    const expected = resolveTotalRevenue(period, ctx).value ?? undefined;
    if (expected === undefined) return;
    const actual = segments.reduce((sum, segment) => sum + (segment.values.get(period) ?? 0), 0);
    if (Math.abs(actual) <= 0.0001) return;
    const error = Math.abs(actual - expected);
    totalError += error;
    if (segmentRevenueTies(actual, expected)) tieCount += 1;
    else badCount += 1;
  });

  return {
    tieCount,
    badCount,
    totalError,
    detailRows: segments.filter((segment) => !segment.aggregate).length,
    familyPreference: segmentFamilyPreference(family)
  };
}

function segmentFamilyPreference(family: string) {
  if (family === "product") return 4;
  if (family === "reportable") return 3;
  if (family === "other") return 2;
  if (family === "geographic") return 1;
  return 0;
}

function fillSegmentTotalRow(
  sheet: ExcelJS.Worksheet,
  periods: string[],
  columns: number[],
  segments: SegmentRevenue[],
  metric: SegmentMetricMapKey,
  label: string,
  auditRows: MappingAuditRow[],
  ctx?: ResolveContext,
  statementConcepts: string[] = [],
  statementResolver?: (period: string, ctx: ResolveContext) => ResolvedValue
) {
  const rowNumber = findLabelRow(sheet, label);
  if (!rowNumber) return 0;
  let filledCells = 0;

  periods.forEach((period, periodIndex) => {
    const cell = sheet.getCell(rowNumber, columns[periodIndex]);
    if (!isHardcodedFinancialInput(cell)) return;
    const col = columns[periodIndex];
    const statementResolved = ctx && statementResolver ? statementResolver(period, ctx) : null;
    const statementSource =
      statementResolved?.value !== null && statementResolved?.value !== undefined
        ? resolvedAuditSource(period, `${label}Resolved`, `Resolved EDGAR ${label}`, statementResolved)
        : ctx && statementConcepts.length
          ? first(period, ctx.duration, statementConcepts)
          : null;
    const segmentTotal = segments.reduce((sum, segment) => sum + (segment[metric].get(period) ?? 0), 0) / 1_000_000;
    const statementValue = statementSource?.value !== undefined ? statementSource.value / 1_000_000 : null;
    const segmentTotalIsUsable = Math.abs(segmentTotal) > 0.0001 && (statementValue === null || segmentStatementMetricTies(segmentTotal, statementValue));
    const value =
      segmentTotalIsUsable
        ? segmentTotal
        : statementValue ?? segmentTotalFromModelRows(sheet, rowNumber, col, label) ?? segmentTotal;
    cell.value = value;
    filledCells += 1;
    auditRows.push({
      sheetName: sheet.name,
      cell: cell.address,
      modelRowLabel: label,
      period,
      valueWritten: value,
      mappingType: "calculated",
      conceptsUsed: segmentTotalIsUsable ? `Sum of EDGAR reportable segment ${metric} rows` : statementSource?.concept ?? `Sum of EDGAR reportable segment ${metric} rows`,
      sourceStatement: segmentTotalIsUsable ? "segment" : statementSource ? "income" : "segment",
      accession: statementSource?.accn ?? "",
      sourceUrl: "",
      cellWritable: true,
      formulaPreserved: false,
      writeBlockedReason: "",
      signConvention: "copied",
      confidence: "high",
      validationStatus: "not_run",
      notes: segmentTotalIsUsable
        ? `Rebuilt ${label} from the sum of EDGAR reportable segment ${metric} rows.`
        : statementSource
        ? `Rebuilt ${label} from consolidated EDGAR ${label.toLowerCase()} because the Segment Analysis total cell did not contain a formula.`
        : `Rebuilt ${label} because the Segment Analysis total cell did not contain a formula.`
    });
  });

  return filledCells;
}

function restoreSegmentTotalFormula(
  sheet: ExcelJS.Worksheet,
  periods: string[],
  columns: number[],
  totalLabel: string,
  endLabel: string,
  auditRows: MappingAuditRow[]
) {
  const totalRow = findLabelRow(sheet, totalLabel);
  const rows = segmentMetricRows(sheet, totalLabel, endLabel, columns);
  if (!totalRow || !rows.length) return 0;

  const firstRow = Math.min(...rows);
  const lastRow = Math.max(...rows);
  const evaluator = new FormulaEvaluator(sheet, { useCachedFormulaResults: false, skipCrossSheetFormulas: true });
  let filledCells = 0;

  periods.forEach((period, periodIndex) => {
    const col = columns[periodIndex];
    const cell = sheet.getCell(totalRow, col);
    if (!hasFormula(cell)) return;

    const segmentTotal = segmentMetricRowsTotal(sheet, rows, col, evaluator);
    const currentTotal = evaluator.evaluateCell(cell);
    if (currentTotal !== null && segmentModelRevenueTies(currentTotal, segmentTotal)) return;

    const formula = `SUM(${columnLetter(col)}${firstRow}:${columnLetter(col)}${lastRow})`;
    cell.value = { formula, result: segmentTotal };
    evaluator.clear();
    filledCells += 1;
    auditRows.push({
      sheetName: sheet.name,
      cell: cell.address,
      modelRowLabel: totalLabel,
      period,
      valueWritten: segmentTotal,
      mappingType: "calculated",
      conceptsUsed: `Sum of Segment Analysis ${totalLabel.toLowerCase()} rows`,
      sourceStatement: "segment",
      accession: "",
      sourceUrl: "",
      cellWritable: true,
      formulaPreserved: false,
      writeBlockedReason: "",
      signConvention: "copied",
      confidence: "high",
      validationStatus: "OK!",
      notes: `Rebuilt ${totalLabel} formula so every filled segment row is included in the Segment Analysis total.`
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
  auditRows: MappingAuditRow[],
  resolver?: (period: string, ctx: ResolveContext) => ResolvedValue
) {
  let filledCells = 0;
  let commentsAdded = 0;
  const warnings: string[] = [];
  if (!rows.length) return { filledCells, commentsAdded, warnings };

  periods.forEach((period, periodIndex) => {
    const resolved = resolver ? resolver(period, ctx) : null;
    const expectedSource =
      resolved?.value !== null && resolved?.value !== undefined
        ? resolvedAuditSource(period, `${suffix}Resolved`, `Resolved EDGAR ${suffix}`, resolved)
        : first(period, ctx.duration, concepts);
    if (!expectedSource) return;
    const col = columns[periodIndex];
    const expected = expectedSource.value / 1_000_000;
    const evaluator = new FormulaEvaluator(sheet, { useCachedFormulaResults: true, skipCrossSheetFormulas: true });
    const actual = segmentMetricRowsTotal(sheet, rows, col, evaluator);
    if (segmentStatementMetricTies(actual, expected)) return;

    if (isFourthQuarterPeriod(period) && Math.abs(actual) <= 0.0001 && Math.abs(expected) > 0.0001) {
      warnings.push(
        `Segment Analysis ${period}: ${suffix} detail did not tie to EDGAR because no reliable 4Q breakout was available; leaving the residual unallocated for review.`
      );
      auditRows.push({
        sheetName: sheet.name,
        cell: `${columnLetter(col)}${rows[0]}`,
        modelRowLabel: `${suffix} segment detail`,
        period,
        valueWritten: 0,
        mappingType: "skipped",
        conceptsUsed: concepts.join(", "),
        sourceStatement: "segment",
        accession: expectedSource.accn ?? "",
        sourceUrl: "",
        cellWritable: false,
        formulaPreserved: false,
        writeBlockedReason: "missing reliable 4Q segment breakout",
        signConvention: "not written",
        confidence: "low",
        validationStatus: "needs_review",
        notes: `Needs review: EDGAR consolidated ${suffix.toLowerCase()} was available for ${period}, but the parser could not validate detailed 4Q segment rows from annual less Q1-Q3 facts.`
      });
      return;
    }

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
  if (preferred) return preferred;

  const genericRow = writableRows.find((rowNumber) => isGenericSegmentPlaceholder(segmentBaseLabel(segmentRowLabel(sheet, rowNumber, suffix), suffix)));
  if (genericRow) {
    setSegmentMetricRowLabel(sheet, genericRow, suffix, "Other / Reconciliation");
    return genericRow;
  }

  const emptyWritableRow = writableRows.find((rowNumber) => {
    const value = numericCellValue(sheet.getCell(rowNumber, col));
    return value === null || Math.abs(value) <= 0.0001;
  });
  if (emptyWritableRow) {
    setSegmentMetricRowLabel(sheet, emptyWritableRow, suffix, "Other / Reconciliation");
    return emptyWritableRow;
  }

  return null;
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
  metric: SegmentMetricMapKey,
  suffix: string,
  auditRows: MappingAuditRow[],
  options: { preserveLinkedBaseLabels?: boolean } = {}
) {
  let filledCells = 0;
  let commentsAdded = 0;
  const usedSegments = new Set<number>();
  const assignedSegments = assignSegmentsToMetricRows(sheet, rows, segments, suffix, periods, usedSegments, options);

  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    const rowNumber = rows[rowIndex];
    const existingLabel = segmentBaseLabel(segmentRowLabel(sheet, rowNumber, suffix), suffix);
    const segmentIndex = assignedSegments.get(rowNumber) ?? null;
    if (segmentIndex === null) {
      const cleared = clearUnmatchedSegmentRow(sheet, rowNumber, periods, columns, existingLabel, auditRows);
      filledCells += cleared.filledCells;
      commentsAdded += cleared.commentsAdded;
      continue;
    }

    const segment = segments[segmentIndex];
    if (!segmentHasMetricData(segment, metric, periods)) continue;
    periods.forEach((period, periodIndex) => {
      const cell = sheet.getCell(rowNumber, columns[periodIndex]);
      const value = (segment[metric].get(period) ?? 0) / 1_000_000;
      if (!writeSegmentMetricCell(cell, value)) return;
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

function writeSegmentMetricCell(cell: ExcelJS.Cell, value: number) {
  if (isSegmentMetricInputCell(cell)) {
    cell.value = value;
    return true;
  }
  const bridgeFormula = formulaBridgeForTargetValue(cell, value);
  if (!bridgeFormula) return false;
  cell.value = { formula: bridgeFormula, result: value };
  return true;
}

function assignSegmentsToMetricRows(
  sheet: ExcelJS.Worksheet,
  rows: number[],
  segments: SegmentRevenue[],
  suffix: string,
  periods: string[],
  usedSegments: Set<number>,
  options: { preserveLinkedBaseLabels?: boolean } = {}
) {
  const assignments = new Map<number, number>();

  rows.forEach((rowNumber) => {
    const existingLabel = segmentBaseLabel(segmentRowLabel(sheet, rowNumber, suffix), suffix);
    if (isGenericSegmentPlaceholder(existingLabel)) return;
    const segmentIndex = segmentIndexForRow(existingLabel, segments, usedSegments);
    if (segmentIndex === null) return;
    assignments.set(rowNumber, segmentIndex);
    usedSegments.add(segmentIndex);
  });

  rows.forEach((rowNumber) => {
    if (assignments.has(rowNumber)) return;
    const segmentIndex = nextUnusedSegmentIndex(segments, usedSegments);
    if (segmentIndex === null) return;
    const segment = segments[segmentIndex];
    if (!segmentHasMetricData(segment, metricKeyForSegmentSuffix(suffix), periods)) return;
    setSegmentMetricRowLabel(sheet, rowNumber, suffix, segment.label, !options.preserveLinkedBaseLabels);
    assignments.set(rowNumber, segmentIndex);
    usedSegments.add(segmentIndex);
  });

  return assignments;
}

function metricKeyForSegmentSuffix(suffix: string): SegmentMetricMapKey {
  if (/operating income/i.test(suffix)) return "operatingIncome";
  if (/d&a|depreciation/i.test(suffix)) return "depreciationAmortization";
  return "values";
}

function setSegmentMetricRowLabel(sheet: ExcelJS.Worksheet, rowNumber: number, suffix: string, segmentLabel: string, updateReferencedBaseLabel = true) {
  const labelCellForRow = labelCell(sheet, rowNumber);
  const formula = cellFormula(labelCellForRow);
  const displayLabel = `${segmentLabel} ${suffix}`;
  const reference = formula?.match(/^=?\$?([A-Z]+)\$?(\d+)\s*(?:&|$)/i);
  if (reference && formula) {
    if (updateReferencedBaseLabel) {
      const baseCell = sheet.getCell(`${reference[1]}${reference[2]}`);
      baseCell.value = segmentLabel;
    }
    labelCellForRow.value = { formula, result: displayLabel };
    return;
  }
  labelCellForRow.value = displayLabel;
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
    const canClearCell = isSegmentMetricInputCell(cell) || Boolean(formulaBridgeForTargetValue(cell, 0)) || hasFormula(cell);
    if (!canClearCell) return;
    const existing = numericCellValue(cell);
    if (existing === 0 || existing === null) return;
    if (!writeSegmentMetricCell(cell, 0)) cell.value = 0;
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

function segmentHasMetricData(segment: SegmentRevenue, metric: SegmentMetricMapKey, periods: string[]) {
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
  metric: SegmentMetricMapKey,
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

function auditNoteForResolvedValue(fillRow: FillRow, resolved: ResolvedValue, period?: string, ctx?: ResolveContext) {
  const derived = derivedSource(resolved);
  const sourceLabels = unique(resolved.sources.map(sourceDisplayLabel).filter(Boolean));
  const classification = resolved.classification ?? fillRow.classification;
  const grouped = classification === "grouped" || sourceLabels.length > 1 || Boolean((resolved.note || fillRow.comment) && fillRow.classification === "grouped");
  const isDerived = Boolean(derived?.derivedTotalValue !== undefined);
  const hasAnalystNote = Boolean((resolved.note || fillRow.comment) && (grouped || isDerived));
  const commentary = period && ctx ? filingCommentaryJustification(fillRow, resolved, period, ctx) : "";
  if (!grouped && !hasAnalystNote && !isDerived && !commentary) return null;

  const parts: string[] = [];
  if (grouped && sourceLabels.length) parts.push(`Grouped from ${summarizeList(sourceLabels, 3)}.`);
  if (resolved.note || fillRow.comment) parts.push(shortAnnotationSentence(resolved.note || fillRow.comment || ""));
  if (isDerived && derived?.derivedPriorPeriods?.length) {
    parts.push(`Derived from ${derived.derivedTotalLabel || derived.label} less ${summarizeList(derived.derivedPriorPeriods, 4)}.`);
  }
  if (commentary) parts.push(commentary);
  return parts.filter(Boolean).join(" ");
}

function filingCommentaryJustification(fillRow: FillRow, resolved: ResolvedValue, period: string, ctx: ResolveContext) {
  if (!shouldUseFilingCommentary(fillRow, resolved)) return "";
  const evidence = rankedFilingCommentary(fillRow, resolved, period, ctx).slice(0, 1);
  if (!evidence.length) return "";
  return `Filing says: ${evidence.map(commentaryEvidenceSummary).join(" ")}`;
}

function shouldUseFilingCommentary(fillRow: FillRow, resolved: ResolvedValue) {
  const classification = resolved.classification ?? fillRow.classification;
  return classification === "grouped" || classification === "partial" || resolved.sources.length > 1 || Boolean(derivedSource(resolved)) || fillRow.comment?.startsWith("LLM-assisted");
}

function rankedFilingCommentary(fillRow: FillRow, resolved: ResolvedValue, period: string, ctx: ResolveContext) {
  const evidence = ctx.commentary?.get(period) ?? [];
  if (!evidence.length) return [];
  const tokens = significantTokens([
    fillRow.label,
    fillRow.modelContext?.sectionHeader,
    fillRow.modelContext?.previousLabel,
    fillRow.modelContext?.nextLabel,
    ...resolved.sources.flatMap((source) => [source.concept, source.label])
  ].join(" "));
  return evidence
    .map((item) => {
      const haystack = significantTokens(`${item.topics.join(" ")} ${item.text}`);
      let score = 0;
      tokens.forEach((token) => {
        if (haystack.has(token)) score += token.length >= 7 ? 3 : 1;
      });
      return { item, score };
    })
    .filter(({ score }) => score >= 2)
    .sort((a, b) => b.score - a.score)
    .map(({ item }) => item);
}

function commentaryEvidenceSummary(item: FilingCommentaryEvidence) {
  const source = [item.form, item.filed].filter(Boolean).join(" filed ");
  const text = shortAnnotationSentence(item.text, 140);
  return `${source ? `${source}: ` : ""}"${text}"`;
}

function validateResolvedValueForWrite(company: CompanyMatch, fillRow: FillRow, period: string, resolved: ResolvedValue): SecWriteValidation {
  const notes: string[] = [];
  let status: SecWriteValidation["status"] = "ok";
  let confidence: SecWriteValidation["confidence"] = resolved.classification === "partial" ? "medium" : "high";
  const sources = resolved.sources.filter((source) => source.sourceLayer !== "model" && !/NotReported/.test(source.concept));
  const resolvedIsDerived = Boolean(derivedSource(resolved));

  if (resolved.value === null || Number.isNaN(resolved.value)) {
    return { status: "blocked", confidence: "low", notes: ["No numeric SEC-backed value was available."] };
  }

  if (!sources.length) {
    return { status: "warning", confidence: "medium", notes: ["Value is a zero/derived model support value with no direct SEC fact."] };
  }

  const expectedUnits = expectedUnitsForFillRow(fillRow);
  for (const source of sources) {
    if (source.cik && source.cik !== company.cik) {
      status = "blocked";
      confidence = "low";
      notes.push(`Company CIK mismatch: source CIK ${source.cik} does not match ${company.cik}.`);
    }
    if (source.periodKey && source.periodKey !== period) {
      if (periodMismatchAllowed(fillRow, period, source)) {
        status = status === "blocked" ? status : "warning";
        confidence = lowerConfidence(confidence, "medium");
        notes.push(`Source period ${source.periodKey} supports ${period} through a beginning-balance or carry-forward rule.`);
      } else {
        status = "blocked";
        confidence = "low";
        notes.push(`Period mismatch: source period ${source.periodKey} does not match model period ${period}.`);
      }
    }
    if (source.periodType && !periodTypeMatchesFillRow(fillRow, source, resolvedIsDerived)) {
      status = "blocked";
      confidence = "low";
      notes.push(`Period type mismatch: ${source.concept} is ${source.periodType}, but ${fillRow.label} expects ${fillRow.kind}.`);
    }
    if (source.unit && expectedUnits.length && !expectedUnits.includes(source.unit)) {
      status = "blocked";
      confidence = "low";
      notes.push(`Unit mismatch: ${source.concept} is reported in ${source.unit}, expected ${expectedUnits.join(" or ")}.`);
    }
    if (source.isAmendment) {
      notes.push(`${source.form ?? "Filing"} is an amended filing; accession ${source.accn ?? "unknown"} was kept distinct in validation.`);
    }
  }

  if (isTotalRevenueFillRow(fillRow)) {
    const componentSources = sources.filter(isRevenueComponentSource);
    if (componentSources.length === 1 && sources.length === 1) {
      status = status === "blocked" ? status : "warning";
      confidence = lowerConfidence(confidence, "medium");
      notes.push("Only a single disaggregated revenue component was available; the item was not treated as high-confidence consolidated revenue.");
    } else if (componentSources.length > 1) {
      status = status === "blocked" ? status : "warning";
      confidence = lowerConfidence(confidence, "medium");
      notes.push("Consolidated revenue was derived from multiple product/service/subscription components because no total revenue concept was available.");
    }
  }

  if (resolved.classification === "grouped" || resolvedIsDerived) confidence = lowerConfidence(confidence, "medium");
  return { status, confidence, notes: unique(notes) };
}

function expectedUnitsForFillRow(fillRow: FillRow) {
  const label = normalize(fillRow.label);
  const concepts = fillRow.concepts ?? [];
  if (concepts.some((concept) => /Share|Shares|WeightedAverageNumberOfShares/i.test(concept)) || /shares/.test(label)) return ["shares"];
  if (concepts.some((concept) => /EarningsPerShare/i.test(concept))) return ["USD/shares"];
  return ["USD"];
}

function periodTypeMatchesFillRow(fillRow: FillRow, source: FactSource, resolvedIsDerived = false) {
  if (source.derivedTotalValue !== undefined) return true;
  if (resolvedIsDerived && source.periodType === "annual") return true;
  if (fillRow.kind === "instant") return source.periodType === "instant";
  return source.periodType === "quarterly" || (isAnnualPeriod(source.periodKey ?? "") && source.periodType === "annual");
}

function periodMismatchAllowed(fillRow: FillRow, period: string, source: FactSource) {
  if (!source.periodKey) return false;
  if (fillRow.statement === "balance" && fillRow.kind === "instant" && isAnnualPeriod(period) && source.periodKey === balanceSheetInstantLookupPeriod(period)) return true;
  if (/carried forward from/i.test(source.note ?? "")) return true;
  return /^beginning/i.test(fillRow.label.trim()) && comparePeriods(source.periodKey, period) < 0;
}

function isTotalRevenueFillRow(fillRow: FillRow) {
  const label = normalize(fillRow.label);
  return /^(total)?revenues?$|^sales$|^totalsales$|^netsales$|^totalnetrevenue$/.test(label) || fillRow.resolver === resolveTotalRevenue;
}

function isRevenueComponentSource(source: FactSource) {
  return REVENUE_COMPONENT_CONCEPTS.includes(source.concept);
}

function lowerConfidence(a: SecWriteValidation["confidence"], b: SecWriteValidation["confidence"]) {
  const rank = { high: 3, medium: 2, low: 1 };
  return rank[a] <= rank[b] ? a : b;
}

function validationStatusText(validation: SecWriteValidation) {
  if (validation.status === "ok") return "OK!";
  return `${validation.status}: ${validation.notes.join(" ")}`.slice(0, 500);
}

function appendValidationNotes(notes: string | null | undefined, validation: SecWriteValidation) {
  return [notes, validation.notes.length ? `SEC validation: ${validation.notes.join(" ")}` : ""].filter(Boolean).join(" ");
}

function summarizeList(items: string[], limit = 3) {
  const shown = items.map((item) => shortAnnotationSentence(item, 80)).slice(0, limit);
  const remaining = items.length - shown.length;
  return `${shown.join(", ")}${remaining > 0 ? `, +${remaining} more` : ""}`;
}

function shortAnnotationSentence(text: string, maxLength = 180) {
  const compact = text.replace(/\s+/g, " ").trim();
  if (compact.length <= maxLength) return compact;
  const boundary = compact.slice(0, maxLength + 1).search(/\s+\S*$/);
  const end = boundary > 80 ? boundary : maxLength;
  return `${compact.slice(0, end).trim()}...`;
}

function rowAnnotationSummary(notes: string[]) {
  const compactNotes = unique(notes.map(rowAnnotationNote).filter(Boolean));
  if (!compactNotes.length) return "";
  const shown = compactNotes.slice(0, 3);
  const remaining = compactNotes.length - shown.length;
  return [...shown, remaining > 0 ? `+${remaining} more item(s) in Mapping Audit.` : ""].filter(Boolean).join("\n");
}

function rowAnnotationNote(note: string) {
  const hasFilingSupport = /\bFiling says:/i.test(note);
  const withoutFilingExcerpt = note.replace(/\s*Filing says:.*$/i, "").trim();
  const sentences = withoutFilingExcerpt.match(/[^.!?]+[.!?]/g)?.map((item) => item.trim()) ?? [];
  const summary = sentences.length ? sentences.slice(0, 2).join(" ") : withoutFilingExcerpt;
  const suffix = hasFilingSupport ? " Filing support available in Mapping Audit." : "";
  return shortAnnotationSentence(`${summary}${suffix}`, 220);
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
  const ebitRow = findIncomeStatementMetricRow(sheet, ["EBIT", "Operating Income", "Operating Income (Loss)", "Income From Operations"]);
  const residualRow = findIncomeStatementResidualRowBefore(sheet, ebitRow, [
    "Other Operating Income (Expense)",
    "Other Operating Expenses",
    "Other Operating Expense",
    "Other Operating Income",
    "Other Operating Income Expense"
  ]);
  if (!ebitRow || !residualRow) return { filledCells, commentsAdded, warnings };

  const evaluator = new FormulaEvaluator(sheet, { skipCrossSheetFormulas: true });
  periods.forEach((period, index) => {
    const resolved = resolveModeledOperatingProfit(period, ctx);
    if (resolved.value === null) return;
    const source = resolvedAuditSource(period, "ModeledOperatingProfitResolved", "Resolved EDGAR operating profit", resolved);
    const col = columns[index];
    const ebitCell = sheet.getCell(ebitRow, col);
    if (!hasFormula(ebitCell)) return;
    const current = evaluator.evaluateCell(ebitCell);
    if (current === null) return;
    const expected = resolved.value / 1_000_000;
    if (incomeStatementFormulaTies(current, expected)) return;

    const residualCell = sheet.getCell(residualRow, col);
    const value = (numericCellValue(residualCell) ?? 0) + expected - current;
    if (!isIncomeStatementReconciliationResidualCellWritable(residualCell)) {
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
      conceptsUsed: resolved.sources.map((item) => item.concept).join(", "),
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

function resolveModeledOperatingProfit(period: string, ctx: ResolveContext): ResolvedValue {
  const operating = first(period, ctx.duration, ["OperatingIncomeLoss"]);
  if (operating) return { value: operating.value, sources: [operating], classification: "direct" };
  return resolvePreTaxIncome(period, ctx);
}

function reconcileIncomeStatementFormulaMetricToEdgar(
  sheet: ExcelJS.Worksheet,
  periods: string[],
  columns: number[],
  ctx: ResolveContext,
  auditRows: MappingAuditRow[],
  labels: string[],
  resolver: (period: string, ctx: ResolveContext) => ResolvedValue,
  metricName: string
) {
  let filledCells = 0;
  let commentsAdded = 0;
  const warnings: string[] = [];
  const rowNumber = findIncomeStatementMetricRow(sheet, labels);
  if (!rowNumber) return { filledCells, commentsAdded, warnings };

  periods.forEach((period, index) => {
    const resolved = resolver(period, ctx);
    if (resolved.value === null) return;
    const source = resolvedAuditSource(period, `${metricName}Resolved`, `Resolved EDGAR ${metricName}`, resolved);
    const cell = sheet.getCell(rowNumber, columns[index]);
    const formula = formulaForCell(cell);
    if (!formula) return;
    const value = resolved.value / 1_000_000;
    const current = numericCellValue(cell);
    if (current !== null && incomeStatementFormulaTies(current, value)) return;

    cell.value = { formula, result: value };
    filledCells += 1;
    const note = `Updated the preserved ${metricName} formula's cached result from EDGAR because the historical input cell links to another model sheet.`;
    addComment(cell, note);
    commentsAdded += 1;
    auditRows.push(statementTotalAuditRow(sheet, cell, rowLabel(sheet, rowNumber), period, value, source, "income", note, "formula result refreshed"));
  });

  return { filledCells, commentsAdded, warnings };
}

function reconcileFinalIncomeStatementConsistency(
  sheet: ExcelJS.Worksheet,
  periods: string[],
  columns: number[],
  ctx: ResolveContext,
  auditRows: MappingAuditRow[]
) {
  let filledCells = 0;
  let commentsAdded = 0;
  const warnings: string[] = [];
  const results = [
    reconcileIncomeStatementFormulaRowsToEdgar(sheet, periods, columns, ctx, auditRows),
    reconcilePreTaxIncomeFormulaRowsToEdgar(sheet, periods, columns, ctx, auditRows),
    reconcileNetIncomeFormulaRowsToEdgar(sheet, periods, columns, ctx, auditRows)
  ];

  for (const result of results) {
    filledCells += result.filledCells;
    commentsAdded += result.commentsAdded;
    warnings.push(...result.warnings);
  }

  return { filledCells, commentsAdded, warnings };
}

function refreshFinalIncomeStatementKeyMetrics(sheet: ExcelJS.Worksheet, periods: string[], columns: number[], ctx: ResolveContext, auditRows: MappingAuditRow[]) {
  refreshFormulaMetricResultsFromResolver(
    sheet,
    periods,
    columns,
    ctx,
    auditRows,
    ["Revenue", "Revenues", "Total Revenue", "Total Revenues", "Sales", "Net Sales", "Net Revenue"],
    resolveTotalRevenue,
    "revenue"
  );
  refreshFormulaMetricResultsFromResolver(
    sheet,
    periods,
    columns,
    ctx,
    auditRows,
    ["EBIT", "Operating Income", "Operating Income (Loss)", "Income From Operations"],
    resolveModeledOperatingProfit,
    "operating profit"
  );
  refreshFormulaMetricResultsFromEdgar(sheet, periods, columns, ctx, auditRows, ["Net Income (Loss)", "Net Income"], CONTINUING_NET_INCOME_CONCEPTS, "net income");
  refreshFormulaMetricResultsFromEdgar(sheet, periods, columns, ctx, auditRows, ["Net Income Available to Common Shareholders", "Net Income Available to Common Stockholders"], COMMON_SHAREHOLDER_INCOME_CONCEPTS, "net income available to common shareholders");

  const revenueRow = findIncomeStatementMetricRow(sheet, ["Revenue", "Revenues", "Total Revenue", "Total Revenues", "Sales", "Net Sales", "Net Revenue"]);
  const grossProfitRow = findIncomeStatementMetricRow(sheet, ["Gross Profit"]);
  const ebitRow = findIncomeStatementMetricRow(sheet, ["EBIT", "Operating Income", "Operating Income (Loss)", "Income From Operations"]);
  const pretaxRow = findIncomeStatementMetricRow(sheet, ["Pre-Tax Income (Loss)", "Pre-Tax Income", "Income Before Taxes", "Income Before Income Taxes"]);
  const netIncomeRow = findIncomeStatementMetricRow(sheet, ["Net Income (Loss)", "Net Income"]);
  const preTaxAdjustmentsRow = findIncomeStatementMetricRow(sheet, ["Pre-Tax Adjustments"]);
  const postTaxAdjustmentsRow = findIncomeStatementMetricRow(sheet, ["Post-Tax Adjustments", "Preferred Stock Dividend"]);
  const discontinuedOperationsRow = findIncomeStatementMetricRow(sheet, ["Discontinued Operations"]);
  const noncontrollingIncomeRow = findIncomeStatementMetricRow(sheet, ["Income (Loss) due to Non-Controlling Interest", "Income (Loss) due to Noncontrolling Interest"]);
  const adjustedNetIncomeRow = findIncomeStatementMetricRow(sheet, ["Adj. Net Income (Loss)", "Adjusted Net Income", "Adj. Net Income"]);
  if (revenueRow) refreshFormulaRowCachedResults(sheet, revenueRow, columns);
  if (grossProfitRow) refreshFormulaRowCachedResults(sheet, grossProfitRow, columns);
  if (ebitRow) refreshFormulaRowCachedResults(sheet, ebitRow, columns, true);
  if (pretaxRow) refreshFormulaRowCachedResults(sheet, pretaxRow, columns, true);
  if (netIncomeRow) refreshFormulaRowCachedResults(sheet, netIncomeRow, columns, true);
  if (preTaxAdjustmentsRow) refreshFormulaRowCachedResults(sheet, preTaxAdjustmentsRow, columns);
  if (postTaxAdjustmentsRow) refreshFormulaRowCachedResults(sheet, postTaxAdjustmentsRow, columns);
  if (discontinuedOperationsRow) refreshFormulaRowCachedResults(sheet, discontinuedOperationsRow, columns);
  if (noncontrollingIncomeRow) refreshFormulaRowCachedResults(sheet, noncontrollingIncomeRow, columns);
  if (adjustedNetIncomeRow) refreshFormulaRowCachedResults(sheet, adjustedNetIncomeRow, columns);
}

function refreshFormulaMetricResultsFromEdgar(
  sheet: ExcelJS.Worksheet,
  periods: string[],
  columns: number[],
  ctx: ResolveContext,
  auditRows: MappingAuditRow[],
  labels: string[],
  concepts: string[],
  metricName: string
) {
  const rowNumber = findIncomeStatementMetricRow(sheet, labels);
  if (!rowNumber) return;
  periods.forEach((period, index) => {
    const source = first(period, ctx.duration, concepts);
    if (!source) return;
    const cell = sheet.getCell(rowNumber, columns[index]);
    const formula = formulaForCell(cell);
    if (!formula) return;
    let value = source.value / 1_000_000;
    if (isFourthQuarterPeriod(period) && index >= 3 && isAnnualDurationSource(source) && source.derivedTotalValue === undefined) {
      const priorPeriods = periods.slice(index - 3, index);
      if (priorPeriods.length === 3 && priorPeriods.every((priorPeriod) => periodYear(priorPeriod) === periodYear(period))) {
        const priorTotal = columns
          .slice(index - 3, index)
          .reduce((total, priorCol) => total + (numericCellValue(sheet.getCell(rowNumber, priorCol)) ?? 0), 0);
        value -= priorTotal;
      }
    }
    if (numericCellValue(cell) !== null && incomeStatementFormulaTies(numericCellValue(cell)!, value)) return;
    cell.value = { formula, result: value };
    auditRows.push(statementTotalAuditRow(sheet, cell, rowLabel(sheet, rowNumber), period, value, source, "income", `Refreshed ${metricName} formula result from EDGAR after dependent formula caches were updated.`, "formula result refreshed"));
  });
}

function refreshFormulaMetricResultsFromResolver(
  sheet: ExcelJS.Worksheet,
  periods: string[],
  columns: number[],
  ctx: ResolveContext,
  auditRows: MappingAuditRow[],
  labels: string[],
  resolver: (period: string, ctx: ResolveContext) => ResolvedValue,
  metricName: string
) {
  const rowNumber = findIncomeStatementMetricRow(sheet, labels);
  if (!rowNumber) return;
  periods.forEach((period, index) => {
    const resolved = resolver(period, ctx);
    if (resolved.value === null) return;
    const source = resolvedAuditSource(period, `${metricName}Resolved`, `Resolved EDGAR ${metricName}`, resolved);
    const cell = sheet.getCell(rowNumber, columns[index]);
    const formula = formulaForCell(cell);
    if (!formula) return;
    const value = resolved.value / 1_000_000;
    if (numericCellValue(cell) !== null && incomeStatementFormulaTies(numericCellValue(cell)!, value)) return;
    cell.value = { formula, result: value };
    auditRows.push(statementTotalAuditRow(sheet, cell, rowLabel(sheet, rowNumber), period, value, source, "income", `Refreshed ${metricName} formula result from EDGAR after dependent formula caches were updated.`, "formula result refreshed"));
  });
}

function isAnnualDurationSource(source: FactSource) {
  return Boolean(source.start?.endsWith("-01-01") && source.end?.endsWith("-12-31"));
}

function refreshFormulaRowCachedResults(
  sheet: ExcelJS.Worksheet,
  rowNumber: number,
  columns: number[],
  skipPeriodColumns = false,
  shouldSkipCell: (rowNumber: number, col: number) => boolean = () => false
) {
  const evaluator = new FormulaEvaluator(sheet, { skipCrossSheetFormulas: true });
  const firstCol = Math.min(...columns);
  const lastCol = Math.min(sheet.columnCount, Math.max(...columns) + 1);
  const periodColumns = new Set(columns);
  for (let col = firstCol; col <= lastCol; col += 1) {
    if (skipPeriodColumns && periodColumns.has(col)) continue;
    if (shouldSkipCell(rowNumber, col)) continue;
    const cell = sheet.getCell(rowNumber, col);
    const formula = formulaForCell(cell);
    if (!formula || formula.includes("!")) continue;
    const result = evaluator.evaluateCell(cell);
    if (result !== null) cell.value = { formula, result };
  }
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

  const evaluator = new FormulaEvaluator(sheet, { skipCrossSheetFormulas: true });
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
    if (!isIncomeStatementReconciliationResidualCellWritable(residualCell)) {
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

function reconcilePreTaxIncomeFormulaRowsToEdgar(
  sheet: ExcelJS.Worksheet,
  periods: string[],
  columns: number[],
  ctx: ResolveContext,
  auditRows: MappingAuditRow[]
) {
  let filledCells = 0;
  let commentsAdded = 0;
  const warnings: string[] = [];
  const pretaxRow = findIncomeStatementMetricRow(sheet, ["Pre-Tax Income (Loss)", "Pre-Tax Income", "Income Before Taxes", "Income Before Income Taxes"]);
  const residualRow = findIncomeStatementResidualRowBefore(sheet, pretaxRow, [
    "Other Non-Operating Income (Expense)",
    "Other Nonoperating Income (Expense)",
    "Other Income (Expense)",
    "Other Expense (Income)"
  ]);
  if (!pretaxRow || !residualRow) return { filledCells, commentsAdded, warnings };

  const evaluator = new FormulaEvaluator(sheet, { skipCrossSheetFormulas: true });
  periods.forEach((period, index) => {
    const resolved = resolvePreTaxIncome(period, ctx);
    if (resolved.value === null) return;
    const source = resolvedAuditSource(period, "PreTaxIncomeResolved", "Resolved EDGAR pre-tax income", resolved);
    const col = columns[index];
    const pretaxCell = sheet.getCell(pretaxRow, col);
    const actual = evaluator.evaluateCell(pretaxCell);
    const expected = resolved.value / 1_000_000;
    if (actual === null || incomeStatementFormulaTies(actual, expected)) return;

    const residualCell = sheet.getCell(residualRow, col);
    const value = (numericCellValue(residualCell) ?? 0) + expected - actual;
    if (!isIncomeStatementReconciliationResidualCellWritable(residualCell)) {
      warnings.push(`Income Statement ${period}: pre-tax income does not tie to EDGAR and ${rowLabel(sheet, residualRow)} was not writable for reconciliation.`);
      return;
    }

    residualCell.value = value;
    evaluator.clear();
    filledCells += 1;
    const note = "Calculated as the residual needed for the model's pre-tax income formula to tie to EDGAR income before taxes.";
    addComment(residualCell, note);
    commentsAdded += 1;
    auditRows.push(statementTotalAuditRow(sheet, residualCell, rowLabel(sheet, residualRow), period, value, source, "income", note, "residual"));
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

function formulaHasCellReference(formula: string) {
  return /(?:'[^']+'|[A-Za-z0-9_ ]+!)?\$?[A-Z]{1,3}\$?\d+/i.test(formula);
}

function formatDividendFormulaNumber(value: number) {
  if (Number.isInteger(value)) return String(value);
  return value.toFixed(3).replace(/0+$/, "").replace(/\.$/, "");
}

function reconcileBalanceSheetCheck(
  sheet: ExcelJS.Worksheet,
  periods: string[],
  columns: number[],
  auditRows: MappingAuditRow[],
  shouldReconcilePeriod: (period: string) => boolean = () => true
) {
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
    if (!shouldReconcilePeriod(period)) return;
    if (isProjectedBalanceSheetCell(sheet, checkRow, col)) return;
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

function reconcilePartialBalanceSheetCheck(
  sheet: ExcelJS.Worksheet,
  periods: string[],
  columns: number[],
  ctx: ResolveContext,
  auditRows: MappingAuditRow[]
) {
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
    if (isProjectedBalanceSheetCell(sheet, checkRow, col)) return;
    const checkCell = sheet.getCell(checkRow, col);
    const check = evaluator.evaluateCell(checkCell);
    if (check === null || valuesTie(check, 0)) return;

    const residualRows = (check < 0 ? assetResidualRows : liabilityResidualRows).filter((rowNumber) =>
      balanceSheetResidualIsPartial(sheet, rowNumber, period, ctx)
    );
    const residualRow = findWritableBalanceSheetResidualRow(sheet, residualRows, col);
    if (!residualRow) {
      warnings.push(`Balance Sheet ${period}: check is ${roundModelValue(check)} and no partial other-bucket row was available for a final section-level reconciliation.`);
      return;
    }

    const residualCell = sheet.getCell(residualRow, col);
    const value = (numericCellValue(residualCell) ?? 0) + Math.abs(check);
    residualCell.value = value;
    evaluator.clear();
    filledCells += 1;
    const note =
      "Calculated as a final residual only because the EDGAR mapping for this other bucket is partial; exact SEC residual buckets are not overwritten by the balance-sheet check.";
    addComment(residualCell, note);
    commentsAdded += 1;
    auditRows.push({
      sheetName: sheet.name,
      cell: residualCell.address,
      modelRowLabel: rowLabel(sheet, residualRow),
      period,
      valueWritten: value,
      mappingType: "calculated",
      conceptsUsed: "Partial other-bucket Balance Sheet Check residual",
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

function balanceSheetResidualIsPartial(sheet: ExcelJS.Worksheet, rowNumber: number, period: string, ctx: ResolveContext) {
  const label = normalize(rowLabel(sheet, rowNumber));
  const resolver = balanceSheetResidualResolverForLabel(label);
  if (!resolver) return false;
  return resolver(balanceSheetInstantLookupPeriod(period), ctx).classification === "partial";
}

function balanceSheetResidualResolverForLabel(normalizedLabel: string) {
  if (/^(prepaidothercurrentassets|prepaidandothercurrentassets|othercurrentassets|prepaidsandothercurrentassets)$/.test(normalizedLabel)) {
    return resolvePrepaidAndOtherCurrentAssets;
  }
  if (/^(othernoncurrentassets|otherlongtermassets|otherltassets|otherassetsandloans)$/.test(normalizedLabel)) {
    return resolveOtherNonCurrentAssets;
  }
  return null;
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
      residualLabels: ["Common Stock & APIC", "Common Stock and APIC", "Common Stock and Additional Paid-In Capital", "Common Stock & Additional Paid-In Capital"]
    },
  ];

  for (const metric of metrics) {
    const totalRow = findRowInSection(sheet, "Balance Sheet", metric.labels, (label) =>
      /working capital|cash flow statement|cashflow statement|income statement|schedule|analysis|drivers/i.test(label)
    );
    if (!totalRow) continue;
    const residualRows = findBalanceSheetResidualRowsBefore(sheet, totalRow, metric.residualLabels);

    periods.forEach((period, index) => {
      const lookupPeriod = balanceSheetInstantLookupPeriod(period);
      const source = first(lookupPeriod, ctx.instant, metric.concepts);
      if (!source) return;
      const col = columns[index];
      if (isProjectedBalanceSheetCell(sheet, totalRow, col)) return;
      const totalCell = sheet.getCell(totalRow, col);
      const expected = source.value / 1_000_000;
      const actual = evaluator.evaluateCell(totalCell) ?? statementMetricCellValue(totalCell, evaluator, expected);
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

function isIncomeStatementReconciliationResidualCellWritable(cell: ExcelJS.Cell) {
  if (isReconciliationResidualCellWritable(cell)) return true;
  return isFormulaReplacementAllowedForReconciliation(rowLabel(cell.worksheet, Number(cell.row)));
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
    const decision = await requestLlmMappingDecision(company, fillRow, periods, candidates, modelChoice.model, ctx);
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
  model: string,
  ctx: ResolveContext
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
        { role: "user", content: JSON.stringify(llmMappingPayload(company, fillRow, periods, candidates, ctx)) }
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

function llmMappingPayload(company: CompanyMatch, fillRow: FillRow, periods: string[], candidates: LlmCandidateFact[], ctx: ResolveContext) {
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
    candidates,
    filingCommentary: llmFilingCommentaryEvidence(fillRow, periods, ctx)
  };
}

function llmFilingCommentaryEvidence(fillRow: FillRow, periods: string[], ctx: ResolveContext) {
  const byText = new Map<string, FilingCommentaryEvidence>();
  const emptyResolved: ResolvedValue = { value: null, sources: [] };
  periods.forEach((period) => {
    rankedFilingCommentary(fillRow, emptyResolved, period, ctx)
      .slice(0, 3)
      .forEach((item) => byText.set(`${item.accn ?? ""}|${item.text}`, item));
  });
  return Array.from(byText.values())
    .slice(0, 8)
    .map((item) => ({
      period: item.period,
      topics: item.topics,
      source: [item.form, item.filed, item.accn].filter(Boolean).join(" / "),
      text: item.text
    }));
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
  return { concept, label: concept, value: 0, sourceLayer: "model" };
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
  const conceptLabels = resolved.sources.length ? unique(resolved.sources.map(sourceDisplayLabel).filter(Boolean)) : [];
  const sourceText = [source?.form || "SEC filing", source?.filed ? `filed ${source.filed}` : ""].filter(Boolean).join(", ");
  const endText = source?.end && source.end !== period ? ` (ended ${source.end})` : "";
  const valueUnit = mappingValueUnit(fillRow, resolved);
  const noteText = notes ? shortAnnotationSentence(notes, 240) : resolved.note ? shortAnnotationSentence(resolved.note, 240) : "";
  return [
    `EDGAR: ${sourceText || "SEC filing"}`,
    `Period: ${period}${endText}`,
    `Value: ${roundModelValue(valueWritten)}${valueUnit}`,
    conceptLabels.length ? `From: ${summarizeList(conceptLabels, 2)}` : "",
    `Confidence: ${confidence}`,
    noteText ? `Note: ${noteText}` : ""
  ]
    .filter(Boolean)
    .join("\n");
}

function mappingValueUnit(fillRow: FillRow, resolved: ResolvedValue) {
  const units = unique(resolved.sources.map((source) => source.unit ?? "").filter(Boolean));
  if (units.length === 1 && units[0] === "shares") return " mm shares";
  if (units.length === 1 && units[0] === "USD/shares") return "";
  if ((fillRow.scale ?? 1) === 1_000_000) return " mm";
  return "";
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
    "EDGAR: inline XBRL segment table",
    `Period: ${period}`,
    `Value: ${roundModelValue(valueWritten)} mm`,
    `From: ${segment.label} ${suffix}`,
    "Confidence: high",
    "Note: Segment row matched to an EDGAR reportable segment."
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
    formulaStatus: hasFormula(cell) ? "formula updated" : "not a formula cell",
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

function blockedMappingAuditRow(
  sheet: ExcelJS.Worksheet,
  cell: ExcelJS.Cell,
  fillRow: FillRow,
  period: string,
  proposedValue: number,
  resolved: ResolvedValue,
  validation: SecWriteValidation,
  decision?: WriteDecision
): MappingAuditRow {
  return {
    sheetName: sheet.name,
    cell: cell.address,
    modelRowLabel: fillRow.label,
    section: fillRow.modelContext?.sectionHeader ?? "",
    period,
    valueWritten: proposedValue,
    mappingType: "skipped",
    conceptsUsed: resolved.sources.map((source) => `${source.concept}=${roundModelValue(source.value / (fillRow.scale ?? 1))}mm`).join("; "),
    secLabels: unique(resolved.sources.map((source) => source.label).filter(Boolean)).join("; "),
    sourceStatement: fillRow.statement,
    accession: unique(resolved.sources.map((source) => source.accn).filter(Boolean)).join("; "),
    sourceUrl: unique(resolved.sources.map((source) => source.sourceUrl).filter(Boolean)).join("; "),
    filingForm: unique(resolved.sources.map((source) => source.form).filter(Boolean)).join("; "),
    filedDate: unique(resolved.sources.map((source) => source.filed).filter(Boolean)).join("; "),
    startDate: unique(resolved.sources.map((source) => source.start).filter(Boolean)).join("; "),
    endDate: unique(resolved.sources.map((source) => source.end).filter(Boolean)).join("; "),
    cellWritable: decision?.writable ?? true,
    formulaPreserved: decision?.formulaPreserved ?? hasFormula(cell),
    formulaStatus: decision?.formulaPreserved ? "preserved" : hasFormula(cell) ? "formula not refreshed" : "not written",
    writeBlockedReason: validation.notes.join(" "),
    signConvention: fillRow.sign === -1 ? "would invert to match model sign convention" : "would copy",
    confidence: validation.confidence,
    validationStatus: validationStatusText(validation),
    notes: `Value was flagged for review instead of written. ${validation.notes.join(" ")}`
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
  profile: TemplateProfile = { kind: "generic", confidence: "low", rationale: [], sheetName: modelSheetName, hasSegmentAnalysis: Boolean(workbook.getWorksheet(SEGMENT_SHEET)) },
  balanceSheetPeriods = periods,
  balanceSheetColumns = columns
) {
  const errors: string[] = [];
  const segmentSheet = workbook.getWorksheet(SEGMENT_SHEET);
  if (segmentSheet) {
    const segmentEvaluator = new FormulaEvaluator(segmentSheet);
    const isFinancialCompanySegmentContext = profile.kind === "financial_company" || hasAnyConcept(ctx, "duration", C.netRevenue);
    errors.push(...validateSegmentGenericRows(segmentSheet, periods, columns));
    if (isFinancialCompanySegmentContext) {
      warnings.push("Segment Analysis tie-outs were treated as review warnings for a financial-company template because segment rows may represent revenue components rather than additive reportable segments.");
    } else {
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
  }
  if (modelSheetName !== MODEL_SHEET) return errors;

  const modelSheet = workbook.getWorksheet(modelSheetName) ?? workbook.getWorksheet(MODEL_SHEET);
  if (modelSheet) {
    const evaluator = new FormulaEvaluator(modelSheet, { useCachedFormulaResults: true });
    const isFinancialCompanyStatementContext = profile.kind === "financial_company" || hasAnyConcept(ctx, "duration", C.netRevenue);
    errors.push(...validateIncomeStatementKeyMetrics(modelSheet, periods, columns, ctx, evaluator, warnings, profile));
    if (!isFinancialCompanyStatementContext) {
      errors.push(...validateBalanceSheetStatementTotals(modelSheet, balanceSheetPeriods, balanceSheetColumns, ctx, evaluator, warnings));
    }
    errors.push(...validateBalanceSheetCheck(modelSheet, balanceSheetPeriods, balanceSheetColumns, evaluator, warnings));

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

type FormulaDisplayResult = number | string;

class FormulaEvaluator {
  private readonly cache = new Map<string, number | null>();
  private readonly displayCache = new Map<string, FormulaDisplayResult | null>();

  constructor(
    private readonly sheet: ExcelJS.Worksheet,
    private readonly options: { useCachedFormulaResults?: boolean; skipCrossSheetFormulas?: boolean } = {}
  ) {}

  clear() {
    this.cache.clear();
    this.displayCache.clear();
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
      if (formula && this.options.skipCrossSheetFormulas && formula.includes("!")) {
        result = "result" in value && typeof value.result === "number" ? value.result : null;
      } else {
        result = formula ? this.evaluateFormula(formula, cell, visited) : null;
      }
      if (result === null && "result" in value && typeof value.result === "number") result = value.result;
    } else if (value && typeof value === "object" && "result" in value && typeof value.result === "number") {
      result = value.result;
    }

    visited.delete(address);
    this.cache.set(address, result);
    return result;
  }

  evaluateDisplayCell(cell: ExcelJS.Cell, visited = new Set<string>()): FormulaDisplayResult | null {
    const address = `${cell.worksheet.name}!${cell.address}`;
    if (this.displayCache.has(address)) return this.displayCache.get(address) ?? null;

    if (visited.has(address)) return null;
    visited.add(address);

    const value = cell.value;
    let result: FormulaDisplayResult | null = null;
    if (typeof value === "number" || typeof value === "string") {
      result = value;
    } else if (value && typeof value === "object" && ("formula" in value || "sharedFormula" in value)) {
      const formula = formulaForCell(cell);
      result = formula ? this.evaluateDisplayFormula(formula, cell, visited) : null;
      if (result === null && "result" in value && (typeof value.result === "number" || typeof value.result === "string")) result = value.result;
    } else if (value && typeof value === "object" && "result" in value && (typeof value.result === "number" || typeof value.result === "string")) {
      result = value.result;
    }

    visited.delete(address);
    this.displayCache.set(address, result);
    return result;
  }

  private evaluateFormula(formula: string, cell: ExcelJS.Cell, visited: Set<string>): number | null {
    const expression = formula.replace(/^=/, "");
    const ifValue = this.evaluateIfExpression(expression, cell, visited);
    if (ifValue !== undefined) return ifValue;
    const functionValue = this.evaluateNumericFunctionExpression(expression, cell, visited);
    if (functionValue !== undefined) return functionValue;

    const withoutSums = this.replaceSumCalls(expression, cell, visited);
    if (withoutSums === null) return null;

    const withRefs = withoutSums.replace(/((?:'[^']+'|[A-Za-z0-9_ ]+)!)?\$?([A-Z]{1,3})\$?(\d+)/g, (reference, sheetPrefix: string, col: string, row: string) => {
      const target = referencedCell(cell, sheetPrefix, col, row);
      if (!target) return "0";
      const value = this.options.useCachedFormulaResults ? numericCellValue(target) : this.evaluateCell(target, visited);
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

  private evaluateNumericFunctionExpression(expression: string, cell: ExcelJS.Cell, visited: Set<string>) {
    const normalized = expression.trim();
    const functionName = normalized.match(/^([A-Z]+)\s*\(/i)?.[1]?.toUpperCase();
    if (!functionName || !["ABS", "ROUND", "MAX", "MIN", "AVERAGE"].includes(functionName)) return undefined;
    const body = functionBody(normalized, functionName);
    if (body === null) return undefined;
    const values = splitFormulaArgs(body).map((arg) => this.evaluateFormula(arg, cell, visited));
    if (values.some((value) => value === null)) return null;
    const numbers = values as number[];
    if (functionName === "ABS") return numbers.length === 1 ? Math.abs(numbers[0]) : null;
    if (functionName === "ROUND") {
      if (numbers.length !== 2) return null;
      const factor = 10 ** numbers[1];
      return Math.round(numbers[0] * factor) / factor;
    }
    if (functionName === "MAX") return numbers.length ? Math.max(...numbers) : null;
    if (functionName === "MIN") return numbers.length ? Math.min(...numbers) : null;
    if (functionName === "AVERAGE") return numbers.length ? numbers.reduce((total, value) => total + value, 0) / numbers.length : null;
    return null;
  }

  private evaluateDisplayFormula(formula: string, cell: ExcelJS.Cell, visited: Set<string>): FormulaDisplayResult | null {
    const expression = formula.replace(/^=/, "").trim();
    const stringLiteral = excelStringLiteral(expression);
    if (stringLiteral !== null) return stringLiteral;

    const ifValue = this.evaluateIfDisplayExpression(expression, cell, visited);
    if (ifValue !== undefined) return ifValue;

    return this.evaluateFormula(expression, cell, visited);
  }

  private evaluateIfDisplayExpression(expression: string, cell: ExcelJS.Cell, visited: Set<string>) {
    const body = functionBody(expression, "IF");
    if (body === null) return undefined;
    const [condition, whenTrue, whenFalse] = splitFormulaArgs(body);
    if (!condition || !whenTrue) return null;
    const conditionResult = this.evaluateCondition(condition, cell, visited);
    if (conditionResult === null) return null;
    return this.evaluateDisplayFormula(conditionResult ? whenTrue : whenFalse ?? "0", cell, visited);
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
            const targetCell = targetSheet.getCell(row, col);
            const value = this.options.useCachedFormulaResults ? numericCellValue(targetCell) : this.evaluateCell(targetCell, visited);
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

function excelStringLiteral(expression: string) {
  if (!/^"(?:""|[^"])*"$/.test(expression)) return null;
  return expression.slice(1, -1).replace(/""/g, '"');
}

function refreshHistoricalFormulaCachedResults(workbook: ExcelJS.Workbook, columns: number[], sheetNames = [MODEL_SHEET, SEGMENT_SHEET]) {
  const historicalColumns = uniqueNumbers(columns);
  for (const sheetName of sheetNames) {
    const sheet = workbook.getWorksheet(sheetName);
    if (!sheet) continue;
    const evaluator = new FormulaEvaluator(sheet, { useCachedFormulaResults: true, skipCrossSheetFormulas: true });
    for (let rowNumber = 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
      for (const col of historicalColumns) {
        if (isProjectedBalanceSheetCell(sheet, rowNumber, col)) continue;
        const cell = sheet.getCell(rowNumber, col);
        if (!hasFormula(cell)) continue;
        const formula = formulaForCell(cell);
        if (!formula || formula.includes("!")) continue;
        const result = evaluator.evaluateCell(cell);
        if (result !== null) setFormulaResult(cell, result);
      }
    }
  }
}

function refreshFinalBalanceSheetKeyMetrics(sheet: ExcelJS.Worksheet, periods: string[], columns: number[], ctx: ResolveContext, auditRows: MappingAuditRow[]) {
  refreshBalanceSheetInputFromResolver(
    sheet,
    periods,
    columns,
    ctx,
    auditRows,
    ["Cash & Cash Equivalents", "Cash and Cash Equivalents", "Cash and Equivalents", "Cash"],
    resolveCashAndCurrentInvestments
  );
  refreshBalanceSheetInputFromResolver(
    sheet,
    periods,
    columns,
    ctx,
    auditRows,
    ["Accounts Receivable", "Accounts Receivable, Net", "Trade Receivables", "Fees Receivable"],
    resolveAccountsReceivable
  );
  refreshBalanceSheetInputFromResolver(
    sheet,
    periods,
    columns,
    ctx,
    auditRows,
    ["Inventory"],
    resolveInventory
  );
  refreshBalanceSheetInputFromResolver(
    sheet,
    periods,
    columns,
    ctx,
    auditRows,
    ["PP&E, Net", "Property Plant and Equipment Net", "Property and Equipment, Net", "Property, Plant and Equipment, Net"],
    resolvePpe
  );
  refreshBalanceSheetInputFromResolver(
    sheet,
    periods,
    columns,
    ctx,
    auditRows,
    ["Intangible Assets, Net", "Intangibles, Net"],
    resolveIntangibleAssets
  );
  refreshBalanceSheetInputFromResolver(
    sheet,
    periods,
    columns,
    ctx,
    auditRows,
    ["Goodwill"],
    resolveGoodwill
  );
  refreshBalanceSheetInputFromResolver(
    sheet,
    periods,
    columns,
    ctx,
    auditRows,
    ["Accounts Payable"],
    resolveAccountsPayable
  );
  refreshBalanceSheetInputFromResolver(
    sheet,
    periods,
    columns,
    ctx,
    auditRows,
    ["Prepaid & Other Current Assets", "Prepaid and Other Current Assets", "Other Current Assets", "Prepaids and Other Current Assets"],
    resolvePrepaidAndOtherCurrentAssets
  );
  refreshBalanceSheetInputFromResolver(
    sheet,
    periods,
    columns,
    ctx,
    auditRows,
    ["Other Non-Current Assets", "Other Long-Term Assets", "Other LT Assets", "Other Assets and Loans"],
    resolveOtherNonCurrentAssets
  );
  refreshBalanceSheetInputFromResolver(
    sheet,
    periods,
    columns,
    ctx,
    auditRows,
    ["Accrued Liabilities"],
    resolveAccruedLiabilities
  );
  refreshBalanceSheetInputFromResolver(
    sheet,
    periods,
    columns,
    ctx,
    auditRows,
    ["Other Current Liabilities", "Other Current Liabs"],
    resolveOtherCurrentLiabilities
  );
  refreshBalanceSheetInputFromResolver(
    sheet,
    periods,
    columns,
    ctx,
    auditRows,
    ["LT Debt (Incl. Current Portion)", "Long Term Debt", "Long-Term Debt"],
    resolveLongTermDebtInclCurrentPortion
  );
  refreshBalanceSheetInputFromResolver(
    sheet,
    periods,
    columns,
    ctx,
    auditRows,
    ["Other Non-Current Liabilities", "Other Long-Term Liabilities", "Other LT Liabilities"],
    resolveOtherNonCurrentLiabilities
  );
  refreshBalanceSheetInputFromResolver(
    sheet,
    periods,
    columns,
    ctx,
    auditRows,
    ["Common Stock & APIC", "Common Stock and APIC", "Common Stock and Additional Paid-In Capital", "Common Stock & Additional Paid-In Capital"],
    resolveCommonStockAndApic
  );
  refreshBalanceSheetInputFromResolver(
    sheet,
    periods,
    columns,
    ctx,
    auditRows,
    ["Retained Earnings"],
    resolveRetainedEarnings
  );
  refreshBalanceSheetInputFromResolver(
    sheet,
    periods,
    columns,
    ctx,
    auditRows,
    ["Treasury Stock"],
    resolveTreasuryAndPreferredStock
  );
  refreshBalanceSheetInputFromResolver(
    sheet,
    periods,
    columns,
    ctx,
    auditRows,
    ["Accumulated Other Comprehensive Income (AOCI)", "Accumulated Other Comprehensive Income", "AOCI"],
    resolveAoci
  );
  refreshBalanceSheetInputFromResolver(
    sheet,
    periods,
    columns,
    ctx,
    auditRows,
    ["Noncontrolling Interests", "Non-controlling Interests"],
    resolveNoncontrollingInterests
  );

  const totalAssetsRow = findRowInSection(sheet, "Balance Sheet", ["Total Assets"], (label) =>
    /working capital|cash flow statement|cashflow statement|income statement|schedule|analysis|drivers/i.test(label)
  );
  const totalLiabilitiesRow = findRowInSection(sheet, "Balance Sheet", ["Total Liabilities"], (label) =>
    /working capital|cash flow statement|cashflow statement|income statement|schedule|analysis|drivers/i.test(label)
  );
  const totalEquityRow = findRowInSection(
    sheet,
    "Balance Sheet",
    ["Total Shareholder's Equity", "Total Shareholders' Equity", "Total Shareholders Equity", "Total Stockholders' Equity", "Total Stockholders Equity", "Total Equity"],
    (label) => /working capital|cash flow statement|cashflow statement|income statement|schedule|analysis|drivers/i.test(label)
  );
  const totalEquityIncludingNciRow = findRowInSection(sheet, "Balance Sheet", ["Total Equity"], (label) =>
    /working capital|cash flow statement|cashflow statement|income statement|schedule|analysis|drivers/i.test(label)
  );
  const totalLiabilitiesAndEquityRow = findRowInSection(
    sheet,
    "Balance Sheet",
    ["Total Liabilities & Shareholder's Equity", "Total Liabilities and Shareholder's Equity"],
    (label) => /working capital|cash flow statement|cashflow statement|income statement|schedule|analysis|drivers/i.test(label)
  );
  const balanceSheetCheckRow = findRowInSection(sheet, "Balance Sheet", ["Balance Sheet Check"], (label) =>
    /working capital|cash flow statement|cashflow statement|income statement|schedule|analysis|drivers/i.test(label)
  );

  const skipProjectedBalanceSheetCell = (rowNumber: number, col: number) => isProjectedBalanceSheetCell(sheet, rowNumber, col);
  for (const rowNumber of [totalAssetsRow, totalLiabilitiesRow, totalEquityRow, totalEquityIncludingNciRow, totalLiabilitiesAndEquityRow, balanceSheetCheckRow]) {
    if (rowNumber) refreshFormulaRowCachedResults(sheet, rowNumber, columns, false, skipProjectedBalanceSheetCell);
  }

  periods.forEach((period, index) => {
    const col = columns[index];
    const lookupPeriod = balanceSheetInstantLookupPeriod(period);
    const assets = first(lookupPeriod, ctx.instant, C.assets);
    const stockholdersEquity = resolveStockholdersEquity(lookupPeriod, ctx);
    const totalEquity = resolveTotalEquityIncludingNci(lookupPeriod, ctx);
    const liabilities = resolveTotalLiabilities(lookupPeriod, ctx);
    if (totalAssetsRow && assets) refreshBalanceSheetTotalFormulaResult(sheet, totalAssetsRow, col, period, assets.value / 1_000_000, assets, "total assets", auditRows);
    if (totalEquityRow && stockholdersEquity.value !== null) {
      const source = resolvedAuditSource(period, "StockholdersEquityResolved", "Resolved EDGAR stockholders' equity", stockholdersEquity);
      refreshBalanceSheetTotalFormulaResult(sheet, totalEquityRow, col, period, stockholdersEquity.value / 1_000_000, source, "shareholders' equity", auditRows);
    }
    if (totalEquityIncludingNciRow && totalEquityIncludingNciRow !== totalEquityRow && totalEquity.value !== null) {
      const source = resolvedAuditSource(period, "TotalEquityIncludingNciResolved", "Resolved EDGAR total equity including noncontrolling interests", totalEquity);
      refreshBalanceSheetTotalFormulaResult(sheet, totalEquityIncludingNciRow, col, period, totalEquity.value / 1_000_000, source, "total equity", auditRows);
    }
    if (totalLiabilitiesRow && liabilities.value !== null) {
      const source = resolvedAuditSource(period, "LiabilitiesResolved", "Resolved EDGAR total liabilities", liabilities);
      refreshBalanceSheetTotalFormulaResult(sheet, totalLiabilitiesRow, col, period, liabilities.value / 1_000_000, source, "total liabilities", auditRows);
    }
  });

  for (const rowNumber of [totalLiabilitiesAndEquityRow, balanceSheetCheckRow]) {
    if (rowNumber) refreshFormulaRowCachedResults(sheet, rowNumber, columns, false, skipProjectedBalanceSheetCell);
  }
  if (balanceSheetCheckRow) {
    const evaluator = new FormulaEvaluator(sheet, { skipCrossSheetFormulas: true });
    periods.forEach((period, index) => {
      const col = columns[index];
      if (skipProjectedBalanceSheetCell(balanceSheetCheckRow, col)) return;
      const cell = sheet.getCell(balanceSheetCheckRow, col);
      const formula = formulaForCell(cell);
      if (!formula) return;
      const result = evaluator.evaluateCell(cell);
      cell.value = { formula, result: persistedFormulaResult(result ?? 0) };
    });
  }
}

function resolvedAuditSource(period: string, concept: string, label: string, resolved: ResolvedValue): FactSource {
  if (resolved.sources.length === 1) return resolved.sources[0];
  return bridgeSource(period, concept, label, resolved.value ?? 0, [resolved]);
}

function refreshBalanceSheetInputFromResolver(
  sheet: ExcelJS.Worksheet,
  periods: string[],
  columns: number[],
  ctx: ResolveContext,
  auditRows: MappingAuditRow[],
  labels: string[],
  resolver: (period: string, ctx: ResolveContext) => ResolvedValue
) {
  const rowNumber = findRowInSection(sheet, "Balance Sheet", labels, (label) =>
    /working capital|cash flow statement|cashflow statement|income statement|schedule|analysis|drivers/i.test(label)
  );
  if (!rowNumber) return;

  periods.forEach((period, index) => {
    const col = columns[index];
    if (isProjectedBalanceSheetCell(sheet, rowNumber, col)) return;
    const cell = sheet.getCell(rowNumber, col);
    const formula = formulaForCell(cell);
    const resolved = resolver(balanceSheetInstantLookupPeriod(period), ctx);
    if (resolved.value === null || resolved.classification === "partial") return;
    const value = resolved.value / 1_000_000;
    if (numericCellValue(cell) !== null && incomeStatementFormulaTies(numericCellValue(cell)!, value) && !formula) return;
    cell.value = value;
    auditRows.push({
      sheetName: sheet.name,
      cell: cell.address,
      modelRowLabel: rowLabel(sheet, rowNumber),
      period,
      valueWritten: value,
      mappingType: "formula updated",
      conceptsUsed: resolved.sources.map((source) => `${source.concept}=${roundModelValue(source.value / 1_000_000)}mm`).join("; "),
      secLabels: resolved.sources.map((source) => source.label).join("; "),
      sourceStatement: "balance",
      accession: unique(resolved.sources.map((source) => source.accn ?? "").filter(Boolean)).join("; "),
      sourceUrl: unique(resolved.sources.map((source) => source.sourceUrl ?? "").filter(Boolean)).join("; "),
      filingForm: unique(resolved.sources.map((source) => source.form ?? "").filter(Boolean)).join("; "),
      filedDate: unique(resolved.sources.map((source) => source.filed ?? "").filter(Boolean)).join("; "),
      startDate: unique(resolved.sources.map((source) => source.start ?? "").filter(Boolean)).join("; "),
      endDate: unique(resolved.sources.map((source) => source.end ?? "").filter(Boolean)).join("; "),
      cellWritable: true,
      formulaPreserved: false,
      formulaStatus: formula ? "balance-sheet formula replaced with EDGAR period-end value" : "not a formula cell",
      writeBlockedReason: "",
      signConvention: "copied",
      confidence: "high",
      validationStatus: "OK!",
      notes: formula
        ? "Replaced the balance-sheet formula with the EDGAR period-end value so annual and 4Q instant balances are not derived from quarterly activity."
        : "Updated the balance-sheet input from EDGAR-sourced line items."
    });
  });
}

function formulaFromResolvedSources(formula: string, resolved: ResolvedValue, scale: number) {
  if (!isNumericSumFormula(formula)) return null;
  if (resolved.value === null) return null;
  const sources = resolved.sources.filter((source) => Number.isFinite(source.value));
  if (sources.length < 2) return null;
  const sourceSum = sources.reduce((total, source) => total + source.value, 0);
  if (!valuesTie(sourceSum, resolved.value)) return null;
  return `SUM(${sources.map((source) => formatDividendFormulaNumber(source.value / scale)).join(",")})`;
}

function refreshBalanceSheetTotalFormulaResult(
  sheet: ExcelJS.Worksheet,
  rowNumber: number,
  col: number,
  period: string,
  value: number,
  source: FactSource,
  metricName: string,
  auditRows: MappingAuditRow[]
) {
  const cell = sheet.getCell(rowNumber, col);
  if (isProjectedBalanceSheetCell(sheet, rowNumber, col)) return;
  const formula = formulaForCell(cell);
  if (!formula) return;
  if (numericCellValue(cell) !== null && incomeStatementFormulaTies(numericCellValue(cell)!, value)) return;
  cell.value = { formula, result: value };
  auditRows.push(statementTotalAuditRow(sheet, cell, rowLabel(sheet, rowNumber), period, value, source, "balance", `Refreshed ${metricName} formula result from EDGAR.`, "formula result refreshed"));
}

function ensureFormulaDisplayCaches(workbook: ExcelJS.Workbook, columns: number[], sheetNames = [MODEL_SHEET, SEGMENT_SHEET]) {
  const historicalColumns = uniqueNumbers(columns);
  for (const sheetName of sheetNames) {
    const sheet = workbook.getWorksheet(sheetName);
    if (!sheet) continue;
    const evaluator = new FormulaEvaluator(sheet, { useCachedFormulaResults: true, skipCrossSheetFormulas: true });

    for (let rowNumber = 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
      for (const col of historicalColumns) {
        if (isProjectedBalanceSheetCell(sheet, rowNumber, col)) continue;
        const cell = sheet.getCell(rowNumber, col);
        if (!hasFormula(cell)) continue;
        const formula = formulaForCell(cell);
        if (!formula || !isNumericDisplayFormula(formula)) continue;
        const result = evaluator.evaluateCell(cell);
        if (result !== null) setFormulaResult(cell, result);
        else if (!hasFormulaResult(cell)) setFormulaResult(cell, 0);
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

function isNumericSumFormula(formula: string) {
  return /^SUM\(\s*-?\d+(?:\.\d+)?(?:\s*,\s*-?\d+(?:\.\d+)?)+\s*\)$/i.test(formula.trim());
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
  const isFinancialCompanyStatementContext = profile.kind === "financial_company" || hasAnyConcept(ctx, "duration", C.netRevenue);
  if (isFinancialCompanyStatementContext && hasAnyConcept(ctx, "duration", C.netRevenue)) {
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
        C.revenue,
        resolveTotalRevenue
      )
    );
  }

  if (isFinancialCompanyStatementContext) {
    errors.push(
      ...validateIncomeStatementMetricAgainstEdgar(
        sheet,
        periods,
        columns,
        ctx,
        evaluator,
        warnings,
        "pre-tax income",
        ["Pre-Tax Income (Loss)", "Pre-Tax Income", "Income Before Taxes", "Income Before Income Taxes"],
        PRETAX_INCOME_CONCEPTS,
        resolvePreTaxIncome
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
    return errors;
  }

  errors.push(...validateIncomeStatementGrossProfitBridge(sheet, periods, columns, evaluator, warnings));
  errors.push(
    ...validateIncomeStatementMetricAgainstEdgar(
      sheet,
      periods,
      columns,
      ctx,
      evaluator,
      warnings,
      "EBIT",
      ["EBIT", "Operating Income", "Operating Income (Loss)", "Income From Operations"],
      C.operatingIncome,
      resolveModeledOperatingProfit
    )
  );
  errors.push(...validateIncomeStatementOperatingExpenseBridge(sheet, periods, columns, evaluator, warnings));
  errors.push(
    ...validateIncomeStatementMetricAgainstEdgar(
      sheet,
      periods,
      columns,
      ctx,
      evaluator,
      warnings,
      "pre-tax income",
      ["Pre-Tax Income (Loss)", "Pre-Tax Income", "Income Before Taxes", "Income Before Income Taxes"],
      PRETAX_INCOME_CONCEPTS,
      resolvePreTaxIncome
    )
  );
  errors.push(...validateIncomeStatementPreTaxBridge(sheet, periods, columns, evaluator, warnings));
  errors.push(...validateIncomeStatementTaxBridge(sheet, periods, columns, evaluator, warnings));
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
  errors.push(
    ...validateIncomeStatementMetricAgainstEdgar(
      sheet,
      periods,
      columns,
      ctx,
      evaluator,
      warnings,
      "net income available to common shareholders",
      ["Net Income Available to Common Shareholders", "Net Income Available to Common Stockholders"],
      COMMON_SHAREHOLDER_INCOME_CONCEPTS
    )
  );
  errors.push(...validateIncomeStatementEbitdaFormula(sheet, periods, columns, evaluator, warnings));
  errors.push(...validateIncomeStatementAdjustedNetIncomeFormula(sheet, periods, columns, evaluator, warnings));
  return errors;
}

function validateIncomeStatementGrossProfitBridge(
  sheet: ExcelJS.Worksheet,
  periods: string[],
  columns: number[],
  evaluator: FormulaEvaluator,
  warnings: string[]
) {
  const errors: string[] = [];
  const netRevenueRow = findIncomeStatementMetricRow(sheet, ["Net Revenue", "Revenue Net of Interest Expense", "Revenues Net of Interest Expense", "Total Net Revenue"]);
  const revenueRow = findIncomeStatementMetricRow(sheet, ["Revenue", "Revenues", "Total Revenue", "Total Revenues", "Sales", "Net Sales"]);
  const baseRow = netRevenueRow ?? revenueRow;
  const grossProfitRow = findIncomeStatementMetricRow(sheet, ["Gross Profit", "Gross Margin"]);
  if (!baseRow || !grossProfitRow || grossProfitRow <= baseRow) return errors;

  const costRows = incomeStatementCostOfRevenueRows(sheet, baseRow, grossProfitRow);
  if (!costRows.length) {
    warnings.unshift("Income Statement Analysis: gross profit bridge skipped because no cost of revenue / COGS rows were found between revenue and gross profit.");
    return errors;
  }

  periods.forEach((period, index) => {
    const col = columns[index];
    const base = evaluator.evaluateCell(sheet.getCell(baseRow, col));
    const grossProfit = evaluator.evaluateCell(sheet.getCell(grossProfitRow, col));
    const costValues = costRows.map((rowNumber) => evaluator.evaluateCell(sheet.getCell(rowNumber, col)));
    if (base === null || grossProfit === null || costValues.some((value) => value === null)) {
      warnings.unshift(`Income Statement ${period}: gross profit bridge could not be evaluated from revenue/net revenue, COGS rows, and gross profit.`);
      return;
    }

    const costs = (costValues as number[]).reduce((total, value) => total + value, 0);
    const expectedGrossProfit = base + costs;
    if (!statementMetricTies(expectedGrossProfit, grossProfit)) {
      const labels = costRows.map((rowNumber) => rowLabel(sheet, rowNumber)).filter(Boolean).join(", ");
      errors.push(
        `Income Statement ${period}: revenue/net revenue plus COGS/cost of revenue equals ${roundModelValue(expectedGrossProfit)}, but gross profit is ${roundModelValue(grossProfit)}. Cost rows checked: ${labels}.`
      );
    }
  });

  return errors;
}

function incomeStatementCostOfRevenueRows(sheet: ExcelJS.Worksheet, baseRow: number, grossProfitRow: number) {
  const rows: number[] = [];
  for (let rowNumber = baseRow + 1; rowNumber < grossProfitRow; rowNumber += 1) {
    const label = rowLabel(sheet, rowNumber);
    if (isCostOfRevenueBridgeRowLabel(label)) rows.push(rowNumber);
  }
  return rows;
}

function isCostOfRevenueBridgeRowLabel(label: string) {
  const normalized = normalize(label);
  if (!normalized) return false;
  if (/grossprofit|grossmargin|netrevenue|revenue|subtotal|total|operatingincome|ebit/.test(normalized)) return false;
  return /cogs|costofgoods|costofrevenue|costofsales|costofservices|costofproducts|costofproduct|costofgoodsandservices|cor/.test(normalized);
}

function validateIncomeStatementTaxBridge(
  sheet: ExcelJS.Worksheet,
  periods: string[],
  columns: number[],
  evaluator: FormulaEvaluator,
  warnings: string[]
) {
  const errors: string[] = [];
  const pretaxRow = findIncomeStatementMetricRow(sheet, ["Pre-Tax Income (Loss)", "Pre-Tax Income", "Income Before Taxes", "Income Before Income Taxes"]);
  const taxRow = findIncomeStatementMetricRow(sheet, ["Income Tax Benefit (Expense)", "Income Tax Expense", "Income Tax Provision (Expense)", "Income Tax"]);
  const netIncomeRow = findIncomeStatementMetricRow(sheet, ["Net Income (Loss)", "Net Income"]);
  if (!pretaxRow || !taxRow || !netIncomeRow || netIncomeRow <= taxRow) return errors;

  const postTaxRows = incomeStatementTaxToNetIncomeBridgeRows(sheet, taxRow, netIncomeRow);
  periods.forEach((period, index) => {
    const col = columns[index];
    const pretax = evaluator.evaluateCell(sheet.getCell(pretaxRow, col));
    const tax = evaluator.evaluateCell(sheet.getCell(taxRow, col));
    const netIncome = evaluator.evaluateCell(sheet.getCell(netIncomeRow, col));
    const postTaxValues = postTaxRows.map((rowNumber) => evaluator.evaluateCell(sheet.getCell(rowNumber, col)));
    if (pretax === null || tax === null || netIncome === null || postTaxValues.some((value) => value === null)) {
      warnings.unshift(`Income Statement ${period}: tax bridge could not be evaluated from pre-tax income, tax expense, post-tax rows, and net income.`);
      return;
    }

    const postTaxItems = (postTaxValues as number[]).reduce((total, value) => total + value, 0);
    const expectedNetIncome = pretax + tax + postTaxItems;
    if (!statementMetricTies(expectedNetIncome, netIncome)) {
      const labels = postTaxRows.map((rowNumber) => rowLabel(sheet, rowNumber)).filter(Boolean).join(", ") || "[none]";
      errors.push(
        `Income Statement ${period}: pre-tax income plus tax expense and post-tax rows equals ${roundModelValue(expectedNetIncome)}, but net income is ${roundModelValue(netIncome)}. Post-tax rows checked: ${labels}.`
      );
    }
  });

  return errors;
}

function incomeStatementTaxToNetIncomeBridgeRows(sheet: ExcelJS.Worksheet, taxRow: number, netIncomeRow: number) {
  const rows: number[] = [];
  for (let rowNumber = taxRow + 1; rowNumber < netIncomeRow; rowNumber += 1) {
    const label = rowLabel(sheet, rowNumber);
    if (isTaxToNetIncomeBridgeRowLabel(label)) rows.push(rowNumber);
  }
  return rows;
}

function isTaxToNetIncomeBridgeRowLabel(label: string) {
  const normalized = normalize(label);
  if (!normalized) return false;
  if (/netincome|earnings|subtotal|total/.test(normalized)) return false;
  return /discontinued|noncontrolling|minorityinterest|equitymethod|preferred|posttax|other/.test(normalized);
}

function validateIncomeStatementPreTaxBridge(
  sheet: ExcelJS.Worksheet,
  periods: string[],
  columns: number[],
  evaluator: FormulaEvaluator,
  warnings: string[]
) {
  const errors: string[] = [];
  const ebitRow = findIncomeStatementMetricRow(sheet, ["EBIT", "Operating Income", "Operating Income (Loss)", "Income From Operations"]);
  const pretaxRow = findIncomeStatementMetricRow(sheet, ["Pre-Tax Income (Loss)", "Pre-Tax Income", "Income Before Taxes", "Income Before Income Taxes"]);
  if (!ebitRow || !pretaxRow || pretaxRow <= ebitRow) return errors;

  const bridgeRows = incomeStatementPreTaxBridgeRows(sheet, ebitRow, pretaxRow);
  periods.forEach((period, index) => {
    const col = columns[index];
    const ebit = evaluator.evaluateCell(sheet.getCell(ebitRow, col));
    const pretax = evaluator.evaluateCell(sheet.getCell(pretaxRow, col));
    const bridgeValues = bridgeRows.map((rowNumber) => evaluator.evaluateCell(sheet.getCell(rowNumber, col)));
    if (ebit === null || pretax === null || bridgeValues.some((value) => value === null)) {
      warnings.unshift(`Income Statement ${period}: pre-tax bridge could not be evaluated from EBIT, below-operating rows, and pre-tax income.`);
      return;
    }

    const belowOperatingItems = (bridgeValues as number[]).reduce((total, value) => total + value, 0);
    const expectedPreTax = ebit + belowOperatingItems;
    if (!statementMetricTies(expectedPreTax, pretax)) {
      const labels = bridgeRows.map((rowNumber) => rowLabel(sheet, rowNumber)).filter(Boolean).join(", ") || "[none]";
      errors.push(
        `Income Statement ${period}: EBIT plus below-operating income/expense equals ${roundModelValue(expectedPreTax)}, but pre-tax income is ${roundModelValue(pretax)}. Below-operating rows checked: ${labels}.`
      );
    }
  });

  return errors;
}

function incomeStatementPreTaxBridgeRows(sheet: ExcelJS.Worksheet, ebitRow: number, pretaxRow: number) {
  const rows: number[] = [];
  for (let rowNumber = ebitRow + 1; rowNumber < pretaxRow; rowNumber += 1) {
    const label = rowLabel(sheet, rowNumber);
    if (isBelowOperatingPreTaxBridgeRowLabel(label)) rows.push(rowNumber);
  }
  return rows;
}

function isBelowOperatingPreTaxBridgeRowLabel(label: string) {
  const normalized = normalize(label);
  if (!normalized) return false;
  if (/othernonoperating|nonoperating|otherincome|otherexpense/.test(normalized)) return true;
  if (/pretax|incomebeforetax|tax|netincome|earnings|subtotal|total|ebit|ebitda|operatingincome/.test(normalized)) return false;
  return /interest|goodwillimpairment|impairment|gain|loss|equitymethod|unconsolidated|extinguishment|foreigncurrency|investment/.test(
    normalized
  );
}

function validateIncomeStatementOperatingExpenseBridge(
  sheet: ExcelJS.Worksheet,
  periods: string[],
  columns: number[],
  evaluator: FormulaEvaluator,
  warnings: string[]
) {
  const errors: string[] = [];
  const netRevenueRow = findIncomeStatementMetricRow(sheet, ["Net Revenue", "Revenue Net of Interest Expense", "Revenues Net of Interest Expense", "Total Net Revenue"]);
  const revenueRow = findIncomeStatementMetricRow(sheet, ["Revenue", "Revenues", "Total Revenue", "Total Revenues", "Sales", "Net Sales"]);
  const baseRow = netRevenueRow ?? revenueRow;
  const ebitRow = findIncomeStatementMetricRow(sheet, ["EBIT", "Operating Income", "Operating Income (Loss)", "Income From Operations"]);
  if (!baseRow || !ebitRow || ebitRow <= baseRow) return errors;

  const operatingExpenseRows = incomeStatementOperatingExpenseRows(sheet, baseRow, ebitRow);
  if (!operatingExpenseRows.length) {
    warnings.unshift("Income Statement Analysis: operating expense bridge skipped because no operating expense rows were found between revenue and EBIT.");
    return errors;
  }

  periods.forEach((period, index) => {
    const col = columns[index];
    const base = evaluator.evaluateCell(sheet.getCell(baseRow, col));
    const ebit = evaluator.evaluateCell(sheet.getCell(ebitRow, col));
    const expenseValues = operatingExpenseRows.map((rowNumber) => evaluator.evaluateCell(sheet.getCell(rowNumber, col)));
    if (base === null || ebit === null || expenseValues.some((value) => value === null)) {
      warnings.unshift(`Income Statement ${period}: operating expense bridge could not be evaluated from revenue/net revenue, operating expense rows, and EBIT.`);
      return;
    }

    const operatingExpenses = (expenseValues as number[]).reduce((total, value) => total + value, 0);
    const expectedEbit = base + operatingExpenses;
    if (!statementMetricTies(expectedEbit, ebit)) {
      const labels = operatingExpenseRows.map((rowNumber) => rowLabel(sheet, rowNumber)).filter(Boolean).join(", ");
      errors.push(
        `Income Statement ${period}: revenue/net revenue plus operating expenses equals ${roundModelValue(expectedEbit)}, but EBIT / operating income is ${roundModelValue(ebit)}. Operating expense rows checked: ${labels}.`
      );
    }
  });

  return errors;
}

function incomeStatementOperatingExpenseRows(sheet: ExcelJS.Worksheet, baseRow: number, ebitRow: number) {
  const totalRow = findOperatingExpenseTotalRowBetween(sheet, baseRow, ebitRow);
  if (totalRow) return [totalRow];

  const rows: number[] = [];
  for (let rowNumber = baseRow + 1; rowNumber < ebitRow; rowNumber += 1) {
    const label = rowLabel(sheet, rowNumber);
    if (isOperatingExpenseBridgeRowLabel(label)) rows.push(rowNumber);
  }
  return rows;
}

function findOperatingExpenseTotalRowBetween(sheet: ExcelJS.Worksheet, baseRow: number, ebitRow: number) {
  for (let rowNumber = baseRow + 1; rowNumber < ebitRow; rowNumber += 1) {
    const normalized = normalize(rowLabel(sheet, rowNumber));
    if (/^(total)?(operating|noninterest)(costs|expenses|costsandexpenses)$/.test(normalized)) return rowNumber;
    if (/^totaloperating(costs|expenses|costsandexpenses)$/.test(normalized)) return rowNumber;
    if (/^totalnoninterestexpenses?$/.test(normalized)) return rowNumber;
  }
  return null;
}

function isOperatingExpenseBridgeRowLabel(label: string) {
  const normalized = normalize(label);
  if (!normalized) return false;
  if (/otheroperating(income|expense)|otheroperatingincomeexpense/.test(normalized)) return true;
  if (/provision.*creditloss|creditloss.*provision|customerengagement|variablecustomerengagement|^vce$/.test(normalized)) return true;
  if (
    /grossprofit|grossmargin|netrevenue|revenue|ebit|ebitda|operatingincome|incomefromoperations|pretax|tax|netincome|earnings|subtotal|total/.test(
      normalized
    )
  ) {
    return false;
  }
  if (/interestincome|interestexpense|othernonoperating|nonoperating/.test(normalized)) return false;
  return /cost|cogs|cor|fulfillment|technology|content|research|development|salesmarketing|marketing|selling|generaladministrative|administrative|compensation|benefits|provision|creditloss|depreciation|amortization|restructuring|impairment|otheroperating|operatingexpense|noninterestexpense|expense/.test(
    normalized
  );
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
  concepts: string[],
  resolver?: (period: string, ctx: ResolveContext) => ResolvedValue
) {
  const errors: string[] = [];
  const rowNumber = findIncomeStatementMetricRow(sheet, labels);
  if (!rowNumber) return errors;

  periods.forEach((period, index) => {
    const resolved = resolver ? resolver(period, ctx) : null;
    const edgarSource = resolved?.value !== null && resolved?.value !== undefined
      ? resolvedAuditSource(period, `${metricName}Resolved`, `Resolved EDGAR ${metricName}`, resolved)
      : first(period, ctx.duration, concepts);
    const expectedRaw = resolved?.value !== null && resolved?.value !== undefined ? resolved.value : edgarSource?.value;
    if (!edgarSource || expectedRaw === null || expectedRaw === undefined) {
      warnings.unshift(`Income Statement ${period}: ${metricName} tie-out skipped because EDGAR ${concepts.join("/")} was unavailable for that period.`);
      return;
    }

    const cell = sheet.getCell(rowNumber, columns[index]);
    const expected = expectedRaw / 1_000_000;
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
  const ebitRow = findPriorAnyLabelRow(sheet, ebitdaRow, ["EBIT", "Operating Income", "Operating Income (Loss)", "Income From Operations"], 10);
  const daRow = findPriorAnyLabelRow(sheet, ebitdaRow, ["Depreciation & Amortization", "Depreciation and Amortization", "D&A"], 6);
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

function validateIncomeStatementAdjustedNetIncomeFormula(
  sheet: ExcelJS.Worksheet,
  periods: string[],
  columns: number[],
  evaluator: FormulaEvaluator,
  warnings: string[]
) {
  const errors: string[] = [];
  const adjustedNetIncomeRow = findIncomeStatementMetricRow(sheet, ["Adj. Net Income (Loss)", "Adjusted Net Income", "Adj. Net Income"]);
  if (!adjustedNetIncomeRow) return errors;

  const netIncomeRow = findPriorAnyLabelRow(
    sheet,
    adjustedNetIncomeRow,
    ["Net Income Available to Common Shareholders", "Net Income Available to Common Stockholders", "Net Income (Loss)", "Net Income"],
    10
  );
  const preTaxAdjustmentsRow = findPriorAnyLabelRow(sheet, adjustedNetIncomeRow, ["Pre-Tax Adjustments"], 10);
  const postTaxAdjustmentsRow = findPriorAnyLabelRow(sheet, adjustedNetIncomeRow, ["Post-Tax Adjustments", "Preferred Stock Dividend"], 10);

  if (!netIncomeRow || !preTaxAdjustmentsRow || !postTaxAdjustmentsRow) {
    warnings.unshift("Income Statement Analysis: adjusted net income tie-out skipped because net income, pre-tax adjustments, or post-tax adjustments row was unavailable.");
    return errors;
  }

  periods.forEach((period, index) => {
    const col = columns[index];
    const adjustedCell = sheet.getCell(adjustedNetIncomeRow, col);
    const adjusted = evaluator.evaluateCell(adjustedCell);
    const netIncome = evaluator.evaluateCell(sheet.getCell(netIncomeRow, col));
    const preTaxAdjustments = evaluator.evaluateCell(sheet.getCell(preTaxAdjustmentsRow, col));
    const postTaxAdjustments = evaluator.evaluateCell(sheet.getCell(postTaxAdjustmentsRow, col));

    if (adjusted === null || netIncome === null || preTaxAdjustments === null || postTaxAdjustments === null) {
      const message = `Income Statement ${columnLetter(col)}${adjustedNetIncomeRow} ${period}: could not evaluate adjusted net income bridge.`;
      if (hasFormula(adjustedCell)) warnings.unshift(message);
      else errors.push(message);
      return;
    }

    const expected = netIncome + preTaxAdjustments + postTaxAdjustments;
    if (!valuesTie(adjusted, expected)) {
      const message = `Income Statement ${columnLetter(col)}${adjustedNetIncomeRow} ${period}: adjusted net income ${roundModelValue(adjusted)} does not equal net income plus pre-tax and post-tax adjustments ${roundModelValue(expected)}.`;
      if (hasFormula(adjustedCell)) warnings.unshift(message);
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

function findPriorAnyLabelRow(sheet: ExcelJS.Worksheet, startRow: number, labels: string[], maxRows: number) {
  for (const label of labels) {
    const rowNumber = findPriorLabelRow(sheet, startRow, label, maxRows);
    if (rowNumber) return rowNumber;
  }
  return null;
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
    ...validateBalanceSheetMetricAgainstResolver(
      sheet,
      periods,
      columns,
      ctx,
      evaluator,
      warnings,
      "total current assets",
      ["Total Current Assets"],
      resolveTotalCurrentAssets
    )
  );
  errors.push(
    ...validateBalanceSheetMetricAgainstResolver(
      sheet,
      periods,
      columns,
      ctx,
      evaluator,
      warnings,
      "total non-current assets",
      ["Total Non-Current Assets", "Total Noncurrent Assets", "Total Long-Term Assets", "Total Long Term Assets"],
      resolveTotalNonCurrentAssets
    )
  );
  errors.push(
    ...validateBalanceSheetMetricAgainstResolver(
      sheet,
      periods,
      columns,
      ctx,
      evaluator,
      warnings,
      "modeled current liabilities",
      ["Total Current Liabilities"],
      resolveModeledCurrentLiabilitiesSubtotal
    )
  );
  errors.push(
    ...validateBalanceSheetMetricAgainstResolver(
      sheet,
      periods,
      columns,
      ctx,
      evaluator,
      warnings,
      "modeled non-current liabilities",
      ["Total Non-Current Liabilities", "Total Noncurrent Liabilities", "Total Long-Term Liabilities", "Total Long Term Liabilities"],
      resolveModeledNonCurrentLiabilitiesSubtotal
    )
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
  errors.push(...validateBalanceSheetOtherBucketMetrics(sheet, periods, columns, ctx, evaluator, warnings));
  errors.push(...validateBalanceSheetFourthQuarterAnnualTies(sheet, periods, columns, evaluator));
  errors.push(
    ...validateBalanceSheetSubtotalEqualsComponents(
      sheet,
      periods,
      columns,
      evaluator,
      "total current assets",
      ["Total Current Assets"],
      [
        ["Cash & Cash Equivalents", "Cash and Cash Equivalents", "Cash and Equivalents", "Cash"],
        ["Accounts Receivable", "Accounts Receivable, Net", "Trade Receivables", "Fees Receivable"],
        ["Inventory"],
        ["Prepaid & Other Current Assets", "Prepaid and Other Current Assets", "Other Current Assets", "Prepaids and Other Current Assets"]
      ]
    )
  );
  errors.push(
    ...validateBalanceSheetSubtotalEqualsComponents(
      sheet,
      periods,
      columns,
      evaluator,
      "total current liabilities",
      ["Total Current Liabilities", "Total Current Liabilities (Excl. Debt)", "Total Current Liabilities Excl. Debt"],
      [
        ["Accounts Payable", "Accounts Payable and Accrued Liabilities", "Accounts Payable & Accrued Liabilities", "Pharmacy Costs Payable"],
        ["Accrued Liabilities", "Accrued Expenses", "Accrued Expenses and Other", "Accrued Expenses and Other Current Liabilities"],
        ["Other Current Liabilities", "Other Current Liabs"]
      ]
    )
  );
  errors.push(
    ...validateBalanceSheetSubtotalEqualsComponents(
      sheet,
      periods,
      columns,
      evaluator,
      "total shareholders' equity",
      [
        "Total Shareholder's Equity",
        "Total Shareholders' Equity",
        "Total Shareholders Equity",
        "Total Stockholders' Equity",
        "Total Stockholders Equity"
      ],
      [
        ["Common Stock & APIC", "Common Stock and APIC", "Common Stock and Additional Paid-In Capital", "Common Stock & Additional Paid-In Capital"],
        ["Retained Earnings"],
        ["Treasury Stock"],
        ["Accumulated Other Comprehensive Income (AOCI)", "Accumulated Other Comprehensive Income", "AOCI"]
      ]
    )
  );
  return errors;
}

function validateBalanceSheetFourthQuarterAnnualTies(
  sheet: ExcelJS.Worksheet,
  periods: string[],
  columns: number[],
  evaluator: FormulaEvaluator
) {
  const errors: string[] = [];
  const sectionRows = balanceSheetSectionRows(sheet);
  if (!sectionRows.length) return errors;

  const byPeriod = new Map(periods.map((period, index) => [period, columns[index]]));
  const years = unique(periods.map(periodYearSuffix).filter((year) => year && !Number.isNaN(Number(year))));
  for (const year of years) {
    const fourthQuarterCol = byPeriod.get(`4Q${year}`);
    const annualCol = byPeriod.get(`FY${year}`);
    if (!fourthQuarterCol || !annualCol) continue;
    for (const rowNumber of sectionRows) {
      if (isProjectedBalanceSheetCell(sheet, rowNumber, fourthQuarterCol) || isProjectedBalanceSheetCell(sheet, rowNumber, annualCol)) continue;
      const label = rowLabel(sheet, rowNumber);
      if (!label || shouldSkipBalanceSheetAnnualTieRow(label)) continue;
      const fourthQuarter = evaluatedCellNumber(sheet.getCell(rowNumber, fourthQuarterCol), evaluator);
      const annual = evaluatedCellNumber(sheet.getCell(rowNumber, annualCol), evaluator);
      if (fourthQuarter === null || annual === null) continue;
      if (!statementMetricTies(fourthQuarter, annual)) {
        errors.push(
          `Balance Sheet ${periodYear(`FY${year}`)} ${label}: 4Q ${roundModelValue(fourthQuarter)} does not equal annual ${roundModelValue(annual)}.`
        );
      }
    }
  }
  return errors;
}

function validateBalanceSheetSubtotalEqualsComponents(
  sheet: ExcelJS.Worksheet,
  periods: string[],
  columns: number[],
  evaluator: FormulaEvaluator,
  metricName: string,
  totalLabels: string[],
  componentLabelGroups: string[][]
) {
  const errors: string[] = [];
  const totalRow = findBalanceSheetRow(sheet, totalLabels);
  if (!totalRow) return errors;
  const totalLabel = rowLabel(sheet, totalRow);
  const componentRows = componentLabelGroups
    .map((labels) => findBalanceSheetRow(sheet, labels))
    .filter((rowNumber): rowNumber is number => rowNumber !== null && rowNumber !== totalRow);
  if (componentRows.length < 2) return errors;

  periods.forEach((period, index) => {
    const col = columns[index];
    if (isProjectedBalanceSheetCell(sheet, totalRow, col)) return;
    const total = evaluatedCellNumber(sheet.getCell(totalRow, col), evaluator);
    const components = componentRows.map((rowNumber) => evaluatedCellNumber(sheet.getCell(rowNumber, col), evaluator));
    if (total === null || components.some((value) => value === null)) return;
    const expected = (components as number[]).reduce((sumValue, value) => sumValue + value, 0);
    if (!statementMetricTies(total, expected)) {
      errors.push(
        `Balance Sheet ${columnLetter(col)}${totalRow} ${period}: ${metricName} ${roundModelValue(total)} does not equal disclosed component sum ${roundModelValue(expected)} for ${totalLabel}.`
      );
    }
  });

  return errors;
}

function balanceSheetSectionRows(sheet: ExcelJS.Worksheet) {
  const sectionStart = findSectionHeaderRow(sheet, "Balance Sheet");
  if (!sectionStart) return [];
  const rows: number[] = [];
  for (let rowNumber = sectionStart + 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const label = rowLabel(sheet, rowNumber);
    if (!label) continue;
    if (isBalanceSheetBoundary(label)) break;
    rows.push(rowNumber);
  }
  return rows;
}

function findBalanceSheetRow(sheet: ExcelJS.Worksheet, labels: string[]) {
  return findRowInSection(sheet, "Balance Sheet", labels, isBalanceSheetBoundary);
}

function isBalanceSheetBoundary(label: string) {
  return /working capital|cash flow statement|cashflow statement|income statement|schedule|analysis|drivers/i.test(label);
}

function shouldSkipBalanceSheetAnnualTieRow(label: string) {
  return /balance sheet check/i.test(label);
}

function evaluatedCellNumber(cell: ExcelJS.Cell, evaluator: FormulaEvaluator) {
  const value = evaluator.evaluateCell(cell);
  if (value !== null) return value;
  return numericCellValue(cell);
}

function validateBalanceSheetOtherBucketMetrics(
  sheet: ExcelJS.Worksheet,
  periods: string[],
  columns: number[],
  ctx: ResolveContext,
  evaluator: FormulaEvaluator,
  warnings: string[]
) {
  const errors: string[] = [];
  const checks: Array<{
    metricName: string;
    labels: string[];
    resolver: (period: string, ctx: ResolveContext) => ResolvedValue;
  }> = [
    {
      metricName: "prepaid and other current assets",
      labels: ["Prepaid & Other Current Assets", "Prepaid and Other Current Assets", "Other Current Assets", "Prepaids and Other Current Assets"],
      resolver: resolvePrepaidAndOtherCurrentAssets
    },
    {
      metricName: "other non-current assets",
      labels: ["Other Non-Current Assets", "Other Long-Term Assets", "Other LT Assets", "Other Assets and Loans"],
      resolver: resolveOtherNonCurrentAssets
    },
    {
      metricName: "accrued liabilities",
      labels: ["Accrued Liabilities", "Accrued Expenses", "Accrued Expenses and Other", "Accrued Expenses and Other Current Liabilities"],
      resolver: resolveAccruedLiabilities
    },
    {
      metricName: "other current liabilities",
      labels: ["Other Current Liabilities", "Other Current Liabs"],
      resolver: resolveOtherCurrentLiabilities
    },
    {
      metricName: "other non-current liabilities",
      labels: ["Other Non-Current Liabilities", "Other Long-Term Liabilities", "Other LT Liabilities"],
      resolver: resolveOtherNonCurrentLiabilities
    }
  ];

  for (const check of checks) {
    errors.push(...validateBalanceSheetMetricAgainstResolver(sheet, periods, columns, ctx, evaluator, warnings, check.metricName, check.labels, check.resolver));
  }
  return errors;
}

function validateBalanceSheetMetricAgainstResolver(
  sheet: ExcelJS.Worksheet,
  periods: string[],
  columns: number[],
  ctx: ResolveContext,
  evaluator: FormulaEvaluator,
  warnings: string[],
  metricName: string,
  labels: string[],
  resolver: (period: string, ctx: ResolveContext) => ResolvedValue
) {
  const errors: string[] = [];
  const rowNumber = findRowInSection(sheet, "Balance Sheet", labels, (label) =>
    /working capital|cash flow statement|cashflow statement|income statement|schedule|analysis|drivers/i.test(label)
  );
  if (!rowNumber) return errors;

  periods.forEach((period, index) => {
    const cell = sheet.getCell(rowNumber, columns[index]);
    if (isProjectedBalanceSheetCell(sheet, rowNumber, columns[index])) return;
    const resolved = resolver(balanceSheetInstantLookupPeriod(period), ctx);
    if (resolved.value === null) {
      warnings.unshift(`Balance Sheet ${period}: ${metricName} tie-out skipped because the EDGAR residual bucket could not be resolved.`);
      return;
    }

    const expected = resolved.value / 1_000_000;
    const modelValue = statementMetricCellValue(cell, evaluator, expected);
    if (modelValue === null) {
      const message = `Balance Sheet ${cell.address} ${period}: could not evaluate model ${metricName}.`;
      errors.push(message);
      return;
    }

    if (!otherBucketMetricTies(modelValue, expected)) {
      const concepts = resolved.sources.map((source) => source.concept).filter(Boolean).join("/");
      const message = `Balance Sheet ${cell.address} ${period}: ${metricName} ${roundModelValue(modelValue)} does not match EDGAR residual bucket ${roundModelValue(expected)}${concepts ? ` from ${concepts}` : ""}.`;
      if (isPartialOtherBucketResolver(metricName, resolved)) {
        warnings.unshift(`${message} The available EDGAR concept is a partial component, so the broader model other-bucket residual was preserved for review.`);
        return;
      }
      errors.push(message);
    }
  });

  return errors;
}

function otherBucketMetricTies(actual: number, expected: number) {
  return Math.abs(actual - expected) <= Math.max(5, Math.abs(expected) * 0.00075);
}

function isPartialOtherBucketResolver(metricName: string, resolved: ResolvedValue) {
  if (resolved.classification === "partial") return true;
  const concepts = new Set(resolved.sources.map((source) => source.concept));
  if (metricName === "prepaid and other current assets" && BROKER_DEALER_CURRENT_ASSETS.some((concept) => concepts.has(concept))) return true;
  return false;
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
    const cell = sheet.getCell(rowNumber, columns[index]);
    if (isProjectedBalanceSheetCell(sheet, rowNumber, columns[index])) return;
    const edgarValue = first(balanceSheetInstantLookupPeriod(period), ctx.instant, concepts);
    if (!edgarValue) {
      warnings.unshift(`Balance Sheet ${period}: ${metricName} tie-out skipped because EDGAR ${concepts.join("/")} was unavailable for that period.`);
      return;
    }

    const expected = edgarValue.value / 1_000_000;
    const modelValue = statementMetricCellValue(cell, evaluator, expected);
    if (modelValue === null) {
      const message = `Balance Sheet ${cell.address} ${period}: could not evaluate model ${metricName}.`;
      errors.push(message);
      return;
    }

    if (!statementMetricTies(modelValue, expected)) {
      const message = `Balance Sheet ${cell.address} ${period}: ${metricName} ${roundModelValue(modelValue)} does not match EDGAR ${roundModelValue(expected)}.`;
      errors.push(message);
    }
  });

  return errors;
}

function validateBalanceSheetCheck(sheet: ExcelJS.Worksheet, periods: string[], columns: number[], evaluator: FormulaEvaluator, warnings: string[] = []) {
  const errors: string[] = [];
  const checkRow = findRowInSection(sheet, "Balance Sheet", ["Balance Sheet Check"], (label) =>
    /working capital|cash flow statement|cashflow statement|income statement|schedule|analysis|drivers/i.test(label)
  );

  if (checkRow) {
    periods.forEach((period, index) => {
      const cell = sheet.getCell(checkRow, columns[index]);
      if (isProjectedBalanceSheetCell(sheet, checkRow, columns[index])) return;
      const check = evaluator.evaluateCell(cell);
      if (check === null) {
        errors.push(`Balance Sheet ${cell.address} ${period}: could not evaluate the model's balance sheet check row.`);
      } else if (!valuesTie(check, 0)) {
        errors.push(`Balance Sheet ${cell.address} ${period}: check is ${roundModelValue(check)}, not OK.`);
      }
    });
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
    if (isProjectedBalanceSheetCell(sheet, assetsRow, col) || isProjectedBalanceSheetCell(sheet, liabilitiesAndEquityRow, col)) return;
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
  const protectedCells = new Map<string, ProtectedCellSnapshot>();
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
    for (const rowNumber of balanceSheetRows(sheet)) {
      for (let col = 1; col <= Math.min(formulaSnapshotEndCol, sheet.columnCount); col += 1) {
        const cell = sheet.getCell(rowNumber, col);
        if (!projectedPeriodColumns(sheet).has(col)) continue;
        protectedCells.set(snapshotAddress(sheet, rowNumber, col), {
          value: cloneCellValue(cell.value),
          note: cloneCellNote(cell.note),
          fingerprint: protectedCellFingerprint(cell)
        });
      }
    }
  }
  return { labels, formulas, protectedCells };
}

function validateWorkbookPreservation(workbook: ExcelJS.Workbook, snapshot: WorkbookSnapshot) {
  const errors: string[] = [];
  for (const [address, expected] of snapshot.labels.entries()) {
    const cell = cellFromSnapshotAddress(workbook, address);
    if (!cell) continue;
    const actual = cellDisplay(cell);
    if (actual !== expected) {
      if (isFormulaErrorResult(expected) && cellFormula(cell)) continue;
      if (isSegmentAnalysisLabelRewriteCell(cell)) continue;
      errors.push(`${address}: row label changed from "${expected}" to "${actual}".`);
    }
  }
  for (const [address, expected] of snapshot.protectedCells.entries()) {
    const cell = cellFromSnapshotAddress(workbook, address);
    if (!cell) continue;
    const actual = protectedCellFingerprint(cell);
    if (actual !== expected.fingerprint) {
      errors.push(`${address}: projected Balance Sheet cell changed after being protected.`);
    }
  }
  for (const [address, expected] of snapshot.formulas.entries()) {
    if (snapshot.protectedCells.has(address)) continue;
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

function restoreProtectedCells(workbook: ExcelJS.Workbook, snapshot: WorkbookSnapshot) {
  for (const [address, expected] of snapshot.protectedCells.entries()) {
    const cell = cellFromSnapshotAddress(workbook, address);
    if (!cell) continue;
    if (protectedCellFingerprint(cell) === expected.fingerprint) continue;
    cell.value = cloneCellValue(expected.value);
    cell.note = cloneCellNote(expected.note);
  }
}

function protectedCellFingerprint(cell: ExcelJS.Cell) {
  return JSON.stringify({
    value: serializableCellValue(cell.value),
    note: cell.note ?? null
  });
}

function serializableCellValue(value: ExcelJS.CellValue): unknown {
  if (value instanceof Date) return { date: value.toISOString() };
  if (!value || typeof value !== "object") return value ?? null;
  if ("formula" in value) {
    const formulaValue = value as { formula: string; shareType?: unknown; ref?: unknown };
    return { formula: formulaValue.formula, shareType: formulaValue.shareType, ref: formulaValue.ref };
  }
  if ("sharedFormula" in value) {
    return { sharedFormula: value.sharedFormula };
  }
  return value;
}

function cloneCellValue(value: ExcelJS.CellValue): ExcelJS.CellValue {
  if (value instanceof Date) return new Date(value.getTime());
  if (!value || typeof value !== "object") return value;
  return JSON.parse(JSON.stringify(value)) as ExcelJS.CellValue;
}

function cloneCellNote(note: ExcelJS.Cell["note"]): ExcelJS.Cell["note"] {
  if (!note || typeof note === "string") return note;
  return JSON.parse(JSON.stringify(note)) as ExcelJS.Cell["note"];
}

function isAllowedFormulaPreservationUpdate(cell: ExcelJS.Cell, expected: string, actual: string | null) {
  if (!actual && balanceSheetRows(cell.worksheet).has(Number(cell.row)) && !isProjectedBalanceSheetCell(cell.worksheet, Number(cell.row), Number(cell.col))) {
    return true;
  }
  if (!actual && isIncomeStatementConstantFormulaReplacement(cell, expected)) return true;
  if (!actual && isIncomeStatementDepreciationAmortizationCell(cell)) return true;
  if (!actual && isFormulaReplacementAllowedForReconciliation(rowLabel(cell.worksheet, Number(cell.row)))) return true;
  if (!actual && isSegmentAnalysisMetricRewriteCell(cell)) return true;
  if (actual && isSegmentAnalysisTotalRewriteCell(cell)) return true;
  if (actual && isBridgeFormula(expected) && isBridgeFormula(actual)) return true;
  if (actual && isNumericSumFormula(expected) && isNumericSumFormula(actual)) return true;
  if (actual && !formulaHasCellReference(expected) && isNumericConstantFormula(actual)) return true;
  return false;
}

function isIncomeStatementConstantFormulaReplacement(cell: ExcelJS.Cell, expected: string) {
  if (formulaHasCellReference(expected)) return false;
  const rowNumber = Number(cell.row);
  const sectionStart = findSectionHeaderRow(cell.worksheet, "Income Statement");
  const sectionEnd = findSectionHeaderRow(cell.worksheet, "Income Statement Analysis");
  return Boolean(sectionStart && rowNumber > sectionStart && (!sectionEnd || rowNumber < sectionEnd));
}

function isIncomeStatementDepreciationAmortizationCell(cell: ExcelJS.Cell) {
  const rowNumber = Number(cell.row);
  return findIncomeStatementMetricRow(cell.worksheet, ["Depreciation & Amortization", "Depreciation and Amortization", "D&A"]) === rowNumber;
}

function isSegmentAnalysisLabelRewriteCell(cell: ExcelJS.Cell) {
  const sheet = cell.worksheet;
  if (sheet.name !== SEGMENT_SHEET || Number(cell.col) > 5) return false;
  const rowNumber = Number(cell.row);
  const rows = new Set([
    ...segmentMetricSectionRows(sheet, "Total Company Revenue", "Revenue Mix"),
    ...segmentMixLabelRows(sheet),
    ...segmentMetricSectionRows(sheet, "Total Company Operating Income", "Operating Income Check"),
    ...segmentMetricSectionRows(sheet, "Total D&A", "D&A Check")
  ]);
  return rows.has(rowNumber);
}

function isSegmentAnalysisMetricRewriteCell(cell: ExcelJS.Cell) {
  const sheet = cell.worksheet;
  if (sheet.name !== SEGMENT_SHEET || Number(cell.col) < 6) return false;
  const rowNumber = Number(cell.row);
  const rows = new Set([
    ...segmentMetricSectionRows(sheet, "Total Company Revenue", "Revenue Mix"),
    ...segmentMetricSectionRows(sheet, "Total Company Operating Income", "Operating Income Check"),
    ...segmentMetricSectionRows(sheet, "Total D&A", "D&A Check")
  ]);
  return rows.has(rowNumber);
}

function isSegmentAnalysisTotalRewriteCell(cell: ExcelJS.Cell) {
  const sheet = cell.worksheet;
  if (sheet.name !== SEGMENT_SHEET || Number(cell.col) < 6) return false;
  const rowNumber = Number(cell.row);
  return ["Total Company Revenue", "Total Company Operating Income", "Total D&A"].some((label) => findLabelRow(sheet, label) === rowNumber);
}

function segmentMetricSectionRows(sheet: ExcelJS.Worksheet, startLabel: string, endLabel: string) {
  const startRow = findLabelRow(sheet, startLabel);
  const endRow = findLabelRow(sheet, endLabel);
  if (!startRow || !endRow || endRow <= startRow) return [];
  const rows: number[] = [];
  for (let rowNumber = startRow + 1; rowNumber < endRow; rowNumber += 1) {
    if (rowLabel(sheet, rowNumber)) rows.push(rowNumber);
  }
  return rows;
}

function isFormulaReplacementAllowedForReconciliation(label: string) {
  const normalized = normalize(label);
  if (["pretaxadjustments", "posttaxadjustments", "discontinuedoperations", "incomelossduetononcontrollinginterest"].includes(normalized)) {
    return true;
  }
  return /other\s+operating\s+income|other\s+operating\s+expense|other\s+non\s*-?\s*operating\s+income|other\s+non\s*-?\s*operating\s+expense|other\s+income\s+\(expense\)|other\s+expense\s+\(income\)|income\s*tax|pre\s*-\s*tax\s+adjustments|post\s*-\s*tax\s+adjustments|discontinued\s+operations|non\s*-\s*controlling\s+interest|prepaid\s+(?:&|and)\s+other\s+current\s+assets|other\s+current\s+assets|other\s+non-current\s+assets|other\s+long-term\s+assets|other\s+lt\s+assets|other\s+current\s+liabilities|other\s+non-current\s+liabilities|accounts\s+payable\s+(?:&|and)\s+accrued\s+liabilities|accrued\s+liabilities/i.test(label);
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
    if (isSegmentAnalysisLabelRewriteCell(cell)) continue;
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
      if (isProjectedBalanceSheetCell(sheet, fillRow.row, col)) return [];
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
      if (isProjectedBalanceSheetCell(sheet, fillRow.row, col)) return;
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
    const edgarRevenue = resolveTotalRevenue(period, ctx);
    const expectedRevenue = edgarRevenue.value !== null ? edgarRevenue.value / 1_000_000 : null;
    if (sheetTotal !== null && !segmentModelRevenueTies(sheetTotal, segmentTotal)) {
      const message = `${sheet.name}!${columnLetter(col)}${totalRow} ${period}: total company revenue formula evaluates to ${roundModelValue(sheetTotal)}, but segment rows sum to ${roundModelValue(segmentTotal)}.`;
      errors.push(message);
    }
    if (expectedRevenue === null) return;
    if (!segmentModelRevenueTies(segmentTotal, expectedRevenue)) {
      const message = `${sheet.name}!${columnLetter(col)}${totalRow} ${period}: segment revenue rows sum to ${roundModelValue(segmentTotal)}, but EDGAR revenue is ${roundModelValue(expectedRevenue)}.`;
      if (isMissingFourthQuarterSegmentDetail(period, segmentTotal, expectedRevenue)) warnings.unshift(`${message} Reliable 4Q segment detail was unavailable, so this is reported as a review warning rather than a blocking error.`);
      else errors.push(message);
    }
    if (sheetTotal !== null && !segmentModelRevenueTies(sheetTotal, expectedRevenue)) {
      const message = `${sheet.name}!${columnLetter(col)}${totalRow} ${period}: total company revenue evaluates to ${roundModelValue(sheetTotal)}, but EDGAR revenue is ${roundModelValue(expectedRevenue)}.`;
      if (isMissingFourthQuarterSegmentDetail(period, sheetTotal, expectedRevenue)) warnings.unshift(`${message} Reliable 4Q segment detail was unavailable, so this is reported as a review warning rather than a blocking error.`);
      else errors.push(message);
    }
  });

  return errors;
}

function isMissingFourthQuarterSegmentDetail(period: string, actual: number, expected: number) {
  return isFourthQuarterPeriod(period) && Math.abs(actual) <= 0.0001 && Math.abs(expected) > 0.0001;
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
  if (metricName !== "Revenue" && !segmentMetricRowsHaveValues(sheet, rows, columns, evaluator)) return errors;

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
      if (metricName === "D&A" && Math.abs(expected) <= 0.0001) {
        warnings.unshift(`${message} The model has no standalone income-statement D&A total for this period, so segment/cash-flow D&A detail is reported as a review warning rather than forced into EBIT.`);
      } else if (metricName === "D&A" && Math.abs(segmentTotal) <= 0.0001) warnings.unshift(`${message} Segment-level D&A detail appears unavailable, so this is reported as a review warning rather than a blocking error.`);
      else if (metricName === "D&A") warnings.unshift(`${message} Segment-level D&A can differ from the linked model D&A line, so this is reported as a review warning.`);
      else if (isMissingFourthQuarterSegmentDetail(period, segmentTotal, expected)) warnings.unshift(`${message} Reliable 4Q segment detail was unavailable, so this is reported as a review warning rather than a blocking error.`);
      else if (rows.some((rowNumber) => hasFormula(sheet.getCell(rowNumber, col)))) warnings.unshift(message);
      else errors.push(message);
    }
    if (sheetTotal !== null && !segmentStatementMetricTies(sheetTotal, expected)) {
      const message = `${sheet.name}!${columnLetter(col)}${totalRow} ${period}: ${metricName} total evaluates to ${roundModelValue(sheetTotal)}, but ${sourceName} is ${roundModelValue(expected)}.`;
      if (isMissingFourthQuarterSegmentDetail(period, sheetTotal, expected)) warnings.unshift(`${message} Reliable 4Q segment detail was unavailable, so this is reported as a review warning rather than a blocking error.`);
      else if (hasFormula(totalCell)) warnings.unshift(message);
      else errors.push(message);
    }
  });

  return errors;
}

function segmentMetricRowsHaveValues(sheet: ExcelJS.Worksheet, rows: number[], columns: number[], evaluator: FormulaEvaluator) {
  return rows.some((rowNumber) =>
    columns.some((col) => Math.abs(evaluator.evaluateCell(sheet.getCell(rowNumber, col)) ?? numericCellValue(sheet.getCell(rowNumber, col)) ?? 0) > 0.0001)
  );
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

function uniqueNumbers(items: number[]) {
  return Array.from(new Set(items)).sort((a, b) => a - b);
}

function uniquePeriodColumnPairs(pairs: Array<{ period: string; col: number }>) {
  const seen = new Set<string>();
  const result: Array<{ period: string; col: number }> = [];
  for (const pair of pairs) {
    const key = `${pair.period}:${pair.col}`;
    if (seen.has(key)) continue;
    seen.add(key);
    result.push(pair);
  }
  return result.sort((a, b) => a.col - b.col);
}

function requestAutomaticWorkbookCalculation(workbook: ExcelJS.Workbook) {
  const calcProperties = workbook.calcProperties as ExcelJS.Workbook["calcProperties"] & {
    forceFullCalc?: boolean;
    calcMode?: string;
    calcOnSave?: boolean;
  };
  calcProperties.calcMode = "auto";
  const optionalCalcProperties = calcProperties as Partial<typeof calcProperties>;
  delete optionalCalcProperties.fullCalcOnLoad;
  delete optionalCalcProperties.forceFullCalc;
  delete optionalCalcProperties.calcOnSave;
}

async function writeWorkbookBufferWithRecalculation(workbook: ExcelJS.Workbook): Promise<Buffer<ArrayBuffer>> {
  refreshWorkbookFormulaDisplayCaches(workbook);
  validateWorkbookFormulaCellsReadyForDisplay(workbook);
  const output = Buffer.from((await workbook.xlsx.writeBuffer()) as ArrayBuffer);
  return enforceXlsxAutomaticCalculation(output);
}

async function enforceXlsxAutomaticCalculation(output: Buffer<ArrayBuffer>): Promise<Buffer<ArrayBuffer>> {
  const zip = await JSZip.loadAsync(output);
  const workbookXmlFile = zip.file("xl/workbook.xml");
  if (!workbookXmlFile) return output;

  const workbookXml = await workbookXmlFile.async("string");
  const calcPr = '<calcPr calcMode="auto"/>';
  const updatedWorkbookXml = /<calcPr\b[^>]*(?:\/>|>[\s\S]*?<\/calcPr>)/.test(workbookXml)
    ? workbookXml.replace(/<calcPr\b[^>]*(?:\/>|>[\s\S]*?<\/calcPr>)/, calcPr)
    : insertWorkbookCalcPr(workbookXml, calcPr);
  zip.file("xl/workbook.xml", updatedWorkbookXml);

  zip.remove("xl/calcChain.xml");
  await removeCalcChainRelationship(zip);
  await removeCalcChainContentType(zip);
  await validateXlsxFormulaCellsReadyForDisplay(zip);

  return Buffer.from(await zip.generateAsync({ type: "arraybuffer", compression: "DEFLATE" }));
}

function insertWorkbookCalcPr(workbookXml: string, calcPr: string) {
  if (/<extLst\b/i.test(workbookXml)) return workbookXml.replace(/<extLst\b/i, `${calcPr}<extLst`);
  return workbookXml.replace(/<\/workbook>\s*$/i, `${calcPr}</workbook>`);
}

function refreshWorkbookFormulaDisplayCaches(workbook: ExcelJS.Workbook) {
  for (const sheet of workbook.worksheets) {
    const evaluator = new FormulaEvaluator(sheet, { useCachedFormulaResults: false });
    sheet.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        if (!hasFormula(cell)) return;
        const formulaBefore = formulaForCell(cell);
        if (!formulaBefore) return;
        if (hasFormulaResult(cell) && !formulaCanReturnText(formulaBefore)) return;
        const result = evaluator.evaluateDisplayCell(cell);
        if (result === null) return;
        setFormulaResult(cell, result);
        const formulaAfter = formulaForCell(cell);
        if (formulaAfter !== formulaBefore) {
          throw new Error(`Formula display-cache refresh changed ${sheet.name}!${cell.address}.`);
        }
      });
    });
  }
}

function formulaCanReturnText(formula: string) {
  return /"(?:""|[^"])*"/.test(formula);
}

function validateWorkbookFormulaCellsReadyForDisplay(workbook: ExcelJS.Workbook) {
  const textFormulas: string[] = [];
  workbook.eachSheet((sheet) => {
    sheet.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        if (typeof cell.value === "string" && isPlainTextFormula(cell.value)) {
          textFormulas.push(`${sheet.name}!${cell.address}`);
        }
      });
    });
  });
  if (textFormulas.length) {
    throw new Error(`Workbook contains formulas stored as text: ${textFormulas.slice(0, 12).join(", ")}`);
  }
}

async function markWorksheetFormulasForRecalculation(zip: JSZip) {
  const sharedStrings = await sharedStringValues(zip);
  const worksheetPaths = Object.keys(zip.files).filter((path) => /^xl\/worksheets\/sheet\d+\.xml$/i.test(path));
  for (const path of worksheetPaths) {
    const file = zip.file(path);
    if (!file) continue;
    const xml = await file.async("string");
    const updated = xml.replace(/<c\b[^>]*>[\s\S]*?<\/c>/g, (cellXml) => {
      if (/<f\b/.test(cellXml)) return markFormulaCellForRecalculation(cellXml);
      return convertPlainTextFormulaCell(cellXml, sharedStrings);
    });
    zip.file(path, updated);
  }
}

function markFormulaCellForRecalculation(cellXml: string) {
  return cellXml.replace(/<f\b([^>]*)>/, (match, attrs: string) => {
    if (/\bca=/.test(attrs)) return match;
    if (/\/\s*$/.test(attrs)) return `<f${attrs.replace(/\/\s*$/, "")} ca="1"/>`;
    return `<f${attrs} ca="1">`;
  });
}

function convertPlainTextFormulaCell(cellXml: string, sharedStrings: Map<number, string>) {
  const stringValue = formulaStringValue(cellXml, sharedStrings);
  if (!stringValue || !isPlainTextFormula(stringValue)) return cellXml;
  const formula = escapeXml(stringValue.replace(/^=/, ""));
  const attrs = (cellXml.match(/^<c\b([^>]*)>/)?.[1] ?? "").replace(/\s+t=(["'])(?:s|str|inlineStr)\1/g, "");
  return `<c${attrs}><f ca="1">${formula}</f></c>`;
}

async function sharedStringValues(zip: JSZip) {
  const values = new Map<number, string>();
  const sharedStringsFile = zip.file("xl/sharedStrings.xml");
  if (!sharedStringsFile) return values;

  const xml = await sharedStringsFile.async("string");
  let index = 0;
  for (const match of xml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)) {
    const text = Array.from(match[1].matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g))
      .map((item) => unescapeXml(item[1]))
      .join("");
    values.set(index, text);
    index += 1;
  }
  return values;
}

function formulaStringValue(cellXml: string, sharedStrings: Map<number, string>) {
  if (/\bt=(["'])s\1/.test(cellXml)) {
    const sharedStringIndex = Number(cellXml.match(/<v>(\d+)<\/v>/)?.[1]);
    return Number.isFinite(sharedStringIndex) ? sharedStrings.get(sharedStringIndex) ?? null : null;
  }
  const value = cellXml.match(/<v>([\s\S]*?)<\/v>/)?.[1];
  if (value !== undefined) return unescapeXml(value);
  const inlineText = cellXml.match(/<is>\s*<t[^>]*>([\s\S]*?)<\/t>\s*<\/is>/)?.[1];
  return inlineText === undefined ? null : unescapeXml(inlineText);
}

async function validateXlsxFormulaCellsReadyForDisplay(zip: JSZip) {
  const workbookXml = await zip.file("xl/workbook.xml")?.async("string");
  if (!workbookXml || !/<calcPr\b[^>]*calcMode="auto"/.test(workbookXml)) {
    throw new Error("Workbook calculation properties are not set to automatic calculation.");
  }

  const problems: string[] = [];
  const sharedStrings = await sharedStringValues(zip);
  const worksheetPaths = Object.keys(zip.files).filter((path) => /^xl\/worksheets\/sheet\d+\.xml$/i.test(path));
  for (const path of worksheetPaths) {
    const file = zip.file(path);
    if (!file) continue;
    const xml = await file.async("string");
    for (const cellXml of xml.match(/<c\b[^>]*>[\s\S]*?<\/c>/g) ?? []) {
      if (/<f\b/.test(cellXml)) continue;
      const textValue = formulaStringValue(cellXml, sharedStrings);
      if (textValue && isPlainTextFormula(textValue)) {
        const address = cellXml.match(/\br="([^"]+)"/)?.[1] ?? "?";
        problems.push(`${path}!${address}: formula is stored as text`);
      }
    }
  }

  if (problems.length) {
    throw new Error(`Workbook has formula cells that may not display until manual edit: ${problems.slice(0, 12).join(", ")}`);
  }
}

function isPlainTextFormula(value: string) {
  return /^=\s*(?:[A-Z_@]|\d|[+\-.]|\()/i.test(value.trim());
}

function escapeXml(value: string) {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function unescapeXml(value: string) {
  return value.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");
}

async function removeCalcChainRelationship(zip: JSZip) {
  const relsFile = zip.file("xl/_rels/workbook.xml.rels");
  if (!relsFile) return;
  const relsXml = await relsFile.async("string");
  const updatedRelsXml = relsXml.replace(
    /\s*<Relationship\b[^>]*Type=(["'])http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/calcChain\1[^>]*\/>/gi,
    ""
  );
  zip.file("xl/_rels/workbook.xml.rels", updatedRelsXml);
}

async function removeCalcChainContentType(zip: JSZip) {
  const contentTypesFile = zip.file("[Content_Types].xml");
  if (!contentTypesFile) return;
  const contentTypesXml = await contentTypesFile.async("string");
  const updatedContentTypesXml = contentTypesXml.replace(/\s*<Override\b[^>]*PartName=(["'])\/xl\/calcChain\.xml\1[^>]*\/>/gi, "");
  zip.file("[Content_Types].xml", updatedContentTypesXml);
}

function jsonError(error: string, status: number) {
  return NextResponse.json({ error }, { status });
}
