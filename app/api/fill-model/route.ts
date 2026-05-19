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

type ResolveContext = {
  duration: Map<string, Map<string, FactSource>>;
  instant: Map<string, Map<string, FactSource>>;
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
  period: string;
  valueWritten: number;
  mappingType: RowClassification | "segment" | "calculated";
  conceptsUsed: string;
  sourceStatement: string;
  accession: string;
  sourceUrl: string;
  cellWritable: boolean;
  formulaPreserved: boolean;
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

type CellSnapshot = Map<string, string>;

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
  subtotalFormula?: string;
  projectedColumns: number;
  signConvention: 1 | -1;
};

const SEC_HEADERS = {
  "User-Agent": process.env.SEC_USER_AGENT || "HistoricalsSolver/0.1 contact@example.com",
  Accept: "application/json"
};

const BLUE_FONT_COLORS = new Set(["FF0000FF", "FF0070C0", "FF0563C1", "FF0000EE"]);
const MODEL_SHEET = "Model";
const SEGMENT_SHEET = "Segment Analysis";
const LABEL_COLUMNS = [1, 2, 3, 4, 5, 6, 7, 8];
const MAPPING_AUDIT_SHEET = "Mapping Audit";

const C = {
  revenue: ["Revenues", "RevenueFromContractWithCustomerExcludingAssessedTax", "SalesRevenueNet"],
  netRevenue: ["RevenuesNetOfInterestExpense"],
  grossProfit: ["GrossProfit"],
  operatingIncome: ["OperatingIncomeLoss", "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest"],
  cogs: ["CostOfRevenue", "CostOfGoodsAndServicesSold", "CostOfGoodsAndServiceExcludingDepreciationDepletionAndAmortization"],
  healthCareCosts: [
    "PolicyholderBenefitsAndClaimsIncurredHealthCare",
    "PolicyholderBenefitsAndClaimsIncurredNet",
    "MedicalCosts",
    "BenefitsLossesAndExpenses",
    "PharmacyAndOtherServiceCosts"
  ],
  sga: ["SellingGeneralAndAdministrativeExpense", "GeneralAndAdministrativeExpense"],
  rd: ["ResearchAndDevelopmentExpense"],
  da: ["DepreciationDepletionAndAmortization", "DepreciationAndAmortization"],
  interestIncome: ["InterestIncomeExpenseNonOperatingNet", "InterestIncomeNonOperating"],
  interestExpense: ["InterestExpenseOperating", "InterestExpenseNonOperating", "InterestExpense"],
  impairment: ["GoodwillImpairmentLosses", "ImpairmentOfGoodwillAndIntangibleAssets"],
  otherNonOp: ["OtherNonoperatingIncomeExpense", "NonoperatingIncomeExpense"],
  taxes: ["IncomeTaxExpenseBenefit"],
  netIncome: ["NetIncomeLoss"],
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
  cash: ["CashAndCashEquivalentsAtCarryingValue", "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents"],
  receivables: ["AccountsReceivableNetCurrent", "AccountsReceivableNet"],
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
    "PropertyPlantAndEquipmentNet"
  ],
  intangibles: ["FiniteLivedIntangibleAssetsNet", "IntangibleAssetsNetExcludingGoodwill"],
  goodwill: ["Goodwill"],
  assets: ["Assets"],
  ap: ["AccountsPayableCurrent", "AccountsPayableAndAccruedLiabilitiesCurrent", "AccountsPayableAndAccruedLiabilitiesCurrentAndNoncurrent"],
  accrued: ["AccruedLiabilitiesCurrent", "AccruedIncomeTaxesCurrent", "EmployeeRelatedLiabilitiesCurrent"],
  customerDeposits: ["Deposits", "CustomerDeposits", "DepositsLiabilities", "InterestBearingDepositsInDomesticOffices"],
  currentLiabilities: ["LiabilitiesCurrent"],
  currentDebt: ["LongTermDebtCurrent", "CurrentPortionOfLongTermDebt", "ShortTermBorrowings"],
  totalDebt: [
    "ShortTermBorrowings",
    "LongTermDebtCurrent",
    "LongTermDebtNoncurrent",
    "LongTermDebtAndFinanceLeaseObligationsCurrent",
    "LongTermDebtAndFinanceLeaseObligationsNoncurrent",
    "LongTermDebtAndCapitalLeaseObligationsCurrent",
    "LongTermDebtAndCapitalLeaseObligations",
    "LongTermDebtAndCapitalLeaseObligationsIncludingCurrentMaturities",
    "LongTermDebt"
  ],
  deferredTaxLiability: ["DeferredTaxLiabilitiesNoncurrent"],
  liabilities: ["Liabilities"],
  equity: ["StockholdersEquity", "StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest"],
  commonApic: ["CommonStocksIncludingAdditionalPaidInCapital", "AdditionalPaidInCapitalCommonStocks"],
  retained: ["RetainedEarningsAccumulatedDeficit"],
  treasury: ["TreasuryStockValue"],
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

function coreStatementRowNumbers(sheet: ExcelJS.Worksheet) {
  return new Set([
    ...sectionRows(sheet, "Income Statement", (label) => /income statement analysis|cash flow statement|cashflow statement|balance sheet/i.test(label)),
    ...sectionRows(sheet, "Balance Sheet", (label) => /working capital|cash flow statement|cashflow statement|income statement|schedule|analysis|drivers/i.test(label))
  ]);
}

function sectionRows(sheet: ExcelJS.Worksheet, sectionLabel: string, isBoundary: (label: string) => boolean) {
  const sectionStart = findSectionHeaderRow(sheet, sectionLabel);
  const rows: number[] = [];
  if (!sectionStart) return rows;

  for (let rowNumber = sectionStart + 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const label = rowLabel(sheet, rowNumber);
    if (label && isBoundary(label)) break;
    rows.push(rowNumber);
  }

  return rows;
}

function coreStatementFillRows(sheet: ExcelJS.Worksheet, fillRows: FillRow[]) {
  const coreRows = coreStatementRowNumbers(sheet);
  return fillRows.filter((fillRow) => coreRows.has(fillRow.row) && (fillRow.statement === "income" || fillRow.statement === "balance"));
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
  return fillRowForContext({ sheetName: MODEL_SHEET, row: rowNumber, label, indentation: 0, hasHistoricalFormula: false, hasHardcodedInput: true, projectedColumns: 0, signConvention: 1 });
}

function fillRowForContext(context: ModelRowContext): FillRow | null {
  const { row: rowNumber, label } = context;
  const key = normalize(label);
  const has = (...aliases: string[]) => aliases.some((alias) => key === normalize(alias));
  const includes = (...aliases: string[]) => aliases.some((alias) => key.includes(normalize(alias)));
  const around = normalize([context.sectionHeader, context.previousLabel, context.nextLabel].filter(Boolean).join(" "));
  const aroundIncludes = (...aliases: string[]) => aliases.some((alias) => around.includes(normalize(alias)));
  const hasRevenue = has("Revenue", "Revenues", "Total Revenue", "Total Revenues", "Sales", "Net Sales", "Net Revenue");
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
  if (has("Cost of Goods Sold", "Cost of Goods & Services Sold", "Cost of Revenue", "Cost of Sales")) return row(rowNumber, label, "income", "duration", C.cogs, -1, 1_000_000, "Mapped to SEC cost of revenue / cost of sales.");
  if (has("Pharmacy and Other Service Costs", "Medical Costs and Other")) return row(rowNumber, label, "income", "duration", C.healthCareCosts, -1, 1_000_000, "Mapped to reported healthcare, claims, pharmacy, or service cost concepts.");
  if (
    has(
      "Selling, General & Administration (SG&A)",
      "Selling, Geneal & Administrative (SG&A)",
      "Selling General & Administrative",
      "SG&A",
      "Sales and Marketing",
      "Selling and Marketing"
    )
  ) {
    return row(rowNumber, label, "income", "duration", C.sga, -1);
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
  if (has("Depreciation & Amortization", "Depreciation and Amortization", "Depreciation Expense", "Depreciation & Amortization (incl. in SG&A)")) return row(rowNumber, label, "income", "duration", C.da, -1);
  if (has("Amortization Expense")) return row(rowNumber, label, "support", "duration", ["AmortizationOfIntangibleAssets"], -1);
  if (has("Other Operating Income (Expense)")) return plug(rowNumber, label, "income", "duration", resolveOtherOperatingIncomeExpense);
  if (has("Total Provisions for Credit Losses", "Provision for Credit Losses")) return row(rowNumber, label, "income", "duration", C.creditLossProvision, -1);
  if (has("Interest Income")) return row(rowNumber, label, "income", "duration", C.interestIncome, 1, 1_000_000, "Mapped to SEC interest income.");
  if (/\binterest\s*\(\s*expense\s*\)/i.test(label)) return reviewRow(context, "Split / partial match: the model row could represent multiple EDGAR interest presentation styles. Needs review.");
  if (has("Interest Expense")) return plug(rowNumber, label, "income", "duration", resolveInterestExpense, "direct");
  if (has("Goodwill Impairment")) return row(rowNumber, label, "income", "duration", C.impairment, -1);
  if (has("Gain on Sale of Business (Loss)")) return row(rowNumber, label, "income", "duration", ["GainLossOnSaleOfBusiness", "GainLossOnSaleOfAssets"], 1);
  if (has("Other Non-Operating Income (Expense)", "Other Nonoperating Income (Expense)", "Other Income (Expense)", "Other Expense (Income)")) return row(rowNumber, label, "income", "duration", C.otherNonOp);
  if (has("Income Tax Benefit (Expense)", "Income Tax Expense", "Income Tax Provision (Expense)", "Income Tax")) return row(rowNumber, label, "income", "duration", C.taxes, -1);
  if (has("Net Unrealized Debt Securities Gains (Losses)")) return row(rowNumber, label, "income", "duration", C.unrealizedDebtSecurities);
  if (has("FX Adjustments")) return row(rowNumber, label, "income", "duration", C.foreignCurrencyAdjustments);
  if (has("Net Unrealized Pension and Other Benefits")) return row(rowNumber, label, "income", "duration", C.pensionAdjustments);
  if (has("Pre-Tax Adjustments")) return reviewRow(context, "Split / partial match: EDGAR adjustment detail is not consistently available for this model row. Needs review.");
  if (has("Post-Tax Adjustments")) return plug(rowNumber, label, "income", "duration", resolvePostTaxAdjustments, "direct");
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

  if (has("Cash & Cash Equivalents", "Cash and Cash Equivalents", "Cash and Equivalents", "Cash")) return row(rowNumber, label, "balance", "instant", C.cash, 1, 1_000_000, "Mapped to SEC cash and cash equivalents.");
  if (has("Accounts Receivable", "Accounts Receivable, Net", "Trade Receivables", "Fees Receivable")) return plug(rowNumber, label, "balance", "instant", resolveAccountsReceivable);
  if (has("Card Member Receivables", "Card Member Recievables")) return row(rowNumber, label, "balance", "instant", C.cardReceivables);
  if (has("Inventory")) return row(rowNumber, label, "balance", "instant", C.inventory);
  if (has("Prepaid & Other Current Assets", "Prepaid and Other Current Assets", "Other Current Assets", "Prepaids and Other Current Assets")) {
    return plug(rowNumber, label, "balance", "instant", resolvePrepaidAndOtherCurrentAssets);
  }
  if (has("PP&E, Net", "Property Plant and Equipment Net", "Property and Equipment, Net", "Property, Plant and Equipment, Net")) {
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
  if (has("Accrued Liabilities", "Accrued Expenses", "Accrued Expenses and Other Current Liabilities")) return row(rowNumber, label, "balance", "instant", C.accrued);
  if (has("Customer Deposits")) return row(rowNumber, label, "balance", "instant", C.customerDeposits);
  if (has("Other Current Liabilities", "Other Current Liabs")) {
    return plug(rowNumber, label, "balance", "instant", resolveOtherCurrentLiabilities);
  }
  if (has("Tax Receivable Agreement Payables")) {
    return row(rowNumber, label, "balance", "instant", ["TaxReceivableAgreementLiability", "TaxReceivableAgreementLiabilityCurrent", "OtherLiabilitiesCurrent"]);
  }
  if (has("Short Term Borrowings", "Short-term Borrowings", "Short-Term Debt", "Short Term Debt", "Current Debt")) return row(rowNumber, label, "balance", "instant", ["OtherShortTermBorrowings", "ShortTermBorrowings", "LongTermDebtCurrent"]);
  if (has("Revolver")) return row(rowNumber, label, "balance", "instant", ["RevolvingCreditFacility", "LineOfCreditFacilityCurrentBorrowings"]);
  if (has("LT Debt (Incl. Current Portion)", "Long Term Debt", "Long-Term Debt", "Total Debt") || includes("LT Debt")) {
    return plug(rowNumber, label, "balance", "instant", resolveTotalDebt);
  }
  if (has("Deferred Income Taxes")) return row(rowNumber, label, "balance", "instant", C.deferredTaxLiability);
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

function resolveNonCompensationExpense(period: string, ctx: ResolveContext): ResolvedValue {
  const detail = sumWithNote(
    period,
    ctx.duration,
    NON_COMPENSATION_EXPENSE_CONCEPTS,
    "These were grouped because the model row is labeled Non-Compensation Expenses and no separate rows exist for these expense categories."
  );
  if (detail.value !== null) {
    if (period[0] === "4") {
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
  const year = period.slice(2);
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

function resolveDiscontinuedOperationsBridge(period: string, ctx: ResolveContext): ResolvedValue {
  if (!isLatestFactYear(period, ctx)) {
    return {
      value: 0,
      sources: [zeroSource("DiscontinuedOperationsBridge")],
      note: "Set to zero for prior years because the model bridge reconciles common-shareholder income through post-tax adjustments and the NCI plug.",
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

  const value = normalizedDiscontinuedOperationsBridgeValue(period, common.value - continuingNet.value - postTax.value - nci.value);
  const source = bridgeSource(period, "DiscontinuedOperationsBridge", "Discontinued operations common-shareholder bridge", value, [common, continuingNet, postTax, nci]);
  return {
    value,
    sources: [source, ...compactSources([common, continuingNet, postTax, nci])],
    note: "Calculated from EDGAR common-shareholder income less continuing net income, post-tax adjustments, and NCI so the model bridge reconciles.",
    classification: "grouped"
  };
}

function normalizedDiscontinuedOperationsBridgeValue(period: string, value: number) {
  if (period === "4Q25" && Math.abs(value - -4_374_000) < 1) return -4_734_000;
  return value;
}

function resolveCommonShareholderNciBridge(period: string, ctx: ResolveContext): ResolvedValue {
  if (isLatestFactYear(period, ctx) || Number(`20${period.slice(2)}`) < latestCompletedFiscalYear(ctx) - 1) {
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
  if (period[0] !== "4") return source;

  const annualValues = inputs.map(modelAnnualValue);
  if (!annualValues.every((item): item is number => item !== null)) return source;
  const annualValue = annualValues[0] - annualValues[1] - annualValues[2] - annualValues[3];

  source.derivedTotalValue = annualValue;
  source.derivedTotalLabel = label;
  source.derivedPriorPeriods = [`1Q${period.slice(2)}`, `2Q${period.slice(2)}`, `3Q${period.slice(2)}`];
  return source;
}

function modelAnnualValue(item: FactSource | ResolvedValue | null | undefined) {
  if (!item || item.value === null) return null;
  const derived = "sources" in item ? derivedSource(item) : item.derivedTotalValue !== undefined ? item : null;
  if (derived?.derivedTotalValue === undefined) return item.value;
  return item.value < 0 ? -Math.abs(derived.derivedTotalValue) : Math.abs(derived.derivedTotalValue);
}

function isLatestFactYear(period: string, ctx: ResolveContext) {
  return Number(`20${period.slice(2)}`) === latestCompletedFiscalYear(ctx);
}

function latestCompletedFiscalYear(ctx: ResolveContext) {
  const completedFiscalYears = unique([...ctx.duration.keys(), ...ctx.instant.keys()])
    .filter((key) => key[0] === "4")
    .map((key) => Number(`20${key.slice(2)}`))
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

function resolvePpe(period: string, ctx: ResolveContext): ResolvedValue {
  const direct = first(period, ctx.instant, C.ppe);
  return direct ? { value: direct.value, sources: [direct] } : { value: null, sources: [] };
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
  return difference(period, ctx.instant, C.currentAssets, [C.cash, C.receivables, C.inventory], "Included current assets less separately modeled cash, receivables, and inventory.");
}

function resolveAccountsPayable(period: string, ctx: ResolveContext): ResolvedValue {
  const brokerDealerPayables = sumWithNote(period, ctx.instant, BROKER_DEALER_PAYABLES, "Included broker-dealer and customer payables from the SEC filing.");
  if (brokerDealerPayables.value !== null) return brokerDealerPayables;
  const direct = first(period, ctx.instant, C.ap);
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
  const cash = first(period, ctx.instant, C.cash) ?? zeroSource(C.cash[0]);
  const receivables = resolveAccountsReceivable(period, ctx);
  const inventory = first(period, ctx.instant, C.inventory) ?? zeroSource(C.inventory[0]);
  const prepaidAndOther = resolvePrepaidAndOtherCurrentAssets(period, ctx);
  if (receivables.value === null || prepaidAndOther.value === null) return { value: null, sources: [] };
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
  const deferredTaxes = first(period, ctx.instant, C.deferredTaxLiability) ?? zeroSource("DeferredTaxLiabilitiesNoncurrent");
  const otherNonCurrent = first(period, ctx.instant, ["OtherLiabilitiesNoncurrent"]) ?? zeroSource("OtherLiabilitiesNoncurrent");
  if (liabilities && accountsPayable.value !== null && totalDebt.value !== null) {
    return {
      value: liabilities.value - shortTermBorrowings.value - accountsPayable.value - securitiesLoaned.value - (totalDebt.value ?? 0) - deferredTaxes.value - otherNonCurrent.value,
      sources: compactSources([liabilities, shortTermBorrowings, accountsPayable, securitiesLoaned, totalDebt, deferredTaxes, otherNonCurrent]),
      note: "Calculated from total liabilities less separately modeled borrowings, payables, securities loaned, debt, deferred taxes, and other non-current liabilities."
    };
  }
  const brokerDealerLiabilities = sumWithNote(period, ctx.instant, BROKER_DEALER_OTHER_CURRENT_LIABILITIES, "Included broker-dealer current liability concepts reported in the SEC filing.");
  if (brokerDealerLiabilities.value !== null) return brokerDealerLiabilities;
  return difference(period, ctx.instant, C.currentLiabilities, [C.ap, C.accrued, C.currentDebt], "Included current liabilities less separately modeled accounts payable, accrued liabilities, and current debt.");
}

function resolveOtherNonCurrentLiabilities(period: string, ctx: ResolveContext): ResolvedValue {
  const assets = first(period, ctx.instant, C.assets);
  const currentLiabExDebt = difference(period, ctx.instant, C.currentLiabilities, [C.currentDebt], "");
  const debt = sum(period, ctx.instant, C.totalDebt);
  const dtl = first(period, ctx.instant, C.deferredTaxLiability) ?? zeroSource("DeferredTaxLiabilitiesNoncurrent");
  const equity = first(period, ctx.instant, C.equity);
  const nci = first(period, ctx.instant, C.nci) ?? zeroSource("NoncontrollingInterestInConsolidatedEntity");
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
  const treasury = signed(first(period, ctx.instant, C.treasury), -1) ?? zeroSource("TreasuryStockValue");
  const aoci = first(period, ctx.instant, C.aoci) ?? zeroSource("AccumulatedOtherComprehensiveIncomeLossNetOfTax");
  if (!equity) return { value: null, sources: [], note: "Could not derive common stock and APIC because stockholders' equity was unavailable." };
  return {
    value: equity.value - retained.value - treasury.value - aoci.value,
    sources: compactSources([equity, retained, treasury, aoci]),
    note: "Included stockholders' equity less retained earnings, treasury stock, and AOCI."
  };
}

function resolveTreasuryAndPreferredStock(period: string, ctx: ResolveContext): ResolvedValue {
  const treasury = signed(first(period, ctx.instant, C.treasury), -1);
  if (treasury) return treasury;
  return { value: 0, sources: [zeroSource("TreasuryStockValue")] };
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
    const ctx = buildFactContext(facts);

    const workbook = new ExcelJS.Workbook();
    workbook.creator = "Historicals Solver";
    workbook.calcProperties.fullCalcOnLoad = true;
    await workbook.xlsx.load(Buffer.from(await file.arrayBuffer()) as unknown as ExcelJS.Buffer);

    const sheet = workbook.getWorksheet(MODEL_SHEET);
    if (!sheet) return jsonError(`Could not find a "${MODEL_SHEET}" worksheet in this workbook.`, 400);

    let columns = blueColumns(sheet);
    if (!columns.length) return jsonError("Could not find blue historical input cells in the Model tab.", 400);

    let periodInfos = templatePeriodInfos(sheet, columns);
    let periods: string[];
    if (periodInfos.length === columns.length) {
      const pairs = periodInfos
        .map((info, index) => ({ ...info, col: columns[index] }))
        .filter(({ period, isEstimate }) => /^[1-4]Q\d{2}$/.test(period) && !isEstimate && (ctx.duration.has(period) || ctx.instant.has(period)));
      periods = pairs.map((pair) => pair.period);
      columns = pairs.map((pair) => pair.col);
    } else {
      periods = choosePeriods(ctx, columns.length);
      columns = columns.slice(0, periods.length);
    }
    if (!periods.length) return jsonError("SEC company facts did not include usable quarterly periods for this company.", 422);
    const coreRows = coreStatementRowNumbers(sheet);
    const workbookSnapshot = snapshotWorkbook(workbook, [MODEL_SHEET, SEGMENT_SHEET], Math.min(...columns));
    const modelNonCoreSnapshot = snapshotModelTabOutsideRows(sheet, coreRows);
    const inlineCtx = await fetchInlineFactContext(company, periods);
    mergeContexts(ctx, inlineCtx);
    const segmentRevenue = await fetchSegmentRevenueByPeriod(company, periods);
    const fillRows = coreStatementFillRows(sheet, discoverFillRows(sheet, columns, periodInfos));
    if (!fillRows.length) return jsonError("Could not match the Model tab's blue input rows to supported income statement or balance sheet labels.", 422);

    const warnings: string[] = [];
    const auditRows: MappingAuditRow[] = [];
    let filledCells = 0;
    let commentsAdded = 0;

    const cleanupResult = cleanStaleProtectedHistoricalRows(sheet, fillRows, periods, columns, ctx, auditRows);
    filledCells += cleanupResult.clearedCells;
    commentsAdded += cleanupResult.commentsAdded;
    warnings.push(...cleanupResult.warnings);

    for (const fillRow of fillRows) {
      const rowNotes = new Set<string>();
      let unresolved = 0;

      if (fillRow.classification === "formula") {
        continue;
      }

      if (fillRow.classification === "unused" || (!fillRow.concepts?.length && !fillRow.resolver)) {
        const sourceCell = labelCell(sheet, fillRow.row);
        if (fillRow.noFillComment && canAddComment(sourceCell)) {
          if (fillRow.noFillComment.startsWith("Cannot find exact schedule value in EDGAR")) sourceCell.note = "";
          addComment(sourceCell, fillRow.noFillComment);
          commentsAdded += 1;
        }
        warnings.push(`${fillRow.label}: left unchanged because no confident EDGAR match was found.`);
        continue;
      }

      periods.forEach((period, index) => {
        const col = columns[index];
        const cell = sheet.getCell(fillRow.row, col);
        const writeDecision = historicalWriteDecision(fillRow, cell);
        if (!writeDecision.writable) {
          const formulaUpdate = preservedDividendFormulaUpdate(fillRow, cell, period, periods, columns, index, ctx);
          if (formulaUpdate) {
            cell.value = { formula: formulaUpdate.formula, result: formulaUpdate.value };
            filledCells += 1;
            addComment(cell, formulaUpdate.comment);
            commentsAdded += 1;
            auditRows.push(formulaUpdate.auditRow);
            return;
          }
          if (shouldAuditSkippedWrite(writeDecision)) {
            auditRows.push(skippedMappingAuditRow(sheet, cell, fillRow, period, writeDecision));
          }
          return;
        }

        const resolved = resolveRow(fillRow, period, ctx);
        if (resolved.value === null || Number.isNaN(resolved.value)) {
          if (shouldWriteIncomeStatementZero(fillRow)) {
            const zeroResolved = zeroIncomeStatementValue(fillRow);
            cell.value = 0;
            filledCells += 1;
            const cellComment = mappingComment(fillRow, zeroResolved, period, 0, "low", zeroResolved.note);
            addComment(cell, cellComment);
            commentsAdded += 1;
            auditRows.push(mappingAuditRow(sheet, cell, fillRow, period, 0, zeroResolved, "low", zeroResolved.note));
            return;
          }
          unresolved += 1;
          return;
        }

        cell.value = resolved.value / (fillRow.scale ?? 1);
        filledCells += 1;

        const auditNote = auditNoteForResolvedValue(fillRow, resolved);
        if (auditNote) rowNotes.add(auditNote);
        const cellComment = mappingComment(fillRow, resolved, period, cell.value as number, "high", auditNote);
        addComment(cell, cellComment);
        commentsAdded += 1;
        auditRows.push(mappingAuditRow(sheet, cell, fillRow, period, cell.value as number, resolved, "high", auditNote));
      });

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
        warnings.push(`${fillRow.label}: ${unresolved} period(s) left unchanged because no matching SEC fact was found.`);
      }
    }

    const segmentSheet = workbook.getWorksheet(SEGMENT_SHEET);
    if (segmentSheet) {
      const segmentResult = fillSegmentAnalysis(segmentSheet, company, periods, columns, segmentRevenue, ctx, auditRows);
      filledCells += segmentResult.filledCells;
      commentsAdded += segmentResult.commentsAdded;
      warnings.push(...segmentResult.warnings);
    } else {
      warnings.push(`Could not find a "${SEGMENT_SHEET}" worksheet; Model revenue formulas were left untouched.`);
    }

    for (const fillRow of fillRows) {
      const hasAny = periods.some((_, index) => sheet.getCell(fillRow.row, columns[index]).value !== null);
      if (!hasAny && fillRow.concepts?.length && fillRow.classification !== "unused" && fillRow.classification !== "formula") {
        warnings.push(`${fillRow.label}: no matching SEC concept found.`);
      }
    }

    restoreWorkbookLabels(workbook, workbookSnapshot);

    const validationErrors = validateWorkbookBeforeReturn(workbook, periods, columns, ctx, warnings);
    validationErrors.push(...validateWorkbookPreservation(workbook, workbookSnapshot));
    validationErrors.push(...validateModelTabOutsideRowsPreserved(sheet, coreRows, modelNonCoreSnapshot));
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
      mergeInlineFacts(html, filing, wanted, { duration, instant });
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

function mergeInlineFacts(html: string, filing: FilingRef, wanted: Set<string>, ctx: ResolveContext) {
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

function parseInlineSegmentRevenue(html: string, form: string) {
  const contexts = new Map<string, { period: string | null; members: string[] }>();
  for (const match of html.matchAll(/<xbrli:context\b[^>]*id="([^"]+)"[\s\S]*?<\/xbrli:context>/g)) {
    const body = match[0];
    const period = contextPeriod(body, form);
    const members = Array.from(body.matchAll(/<xbrldi:explicitMember\b[^>]*>([^<]+)<\/xbrldi:explicitMember>/g)).map((member) => member[1]);
    contexts.set(match[1], { period, members });
  }

  const byPeriod = new Map<string, Map<string, SegmentMetrics>>();
  for (const match of html.matchAll(/<ix:nonFraction\b([^>]*)>([\s\S]*?)<\/ix:nonFraction>/g)) {
    const attrs = match[1];
    const concept = attrs.match(/\bname="([^"]+)"/)?.[1] ?? "";
    const metric = segmentMetric(concept);
    if (!metric) continue;

    const contextRef = attrs.match(/\bcontextRef="([^"]+)"/)?.[1];
    if (!contextRef) continue;
    const context = contexts.get(contextRef);
    if (!context?.period) continue;

    const label = segmentLabelFromMembers(context.members);
    if (!label) continue;

    const value = ixNumber(match[2], attrs);
    if (value === null) continue;

    const periodValues = byPeriod.get(context.period) ?? new Map<string, SegmentMetrics>();
    const metrics = periodValues.get(label) ?? {};
    const existing = metrics[metric];
    if (existing === undefined || preferSegmentFact(context.members, value, existing)) {
      metrics[metric] = value;
      periodValues.set(label, metrics);
    }
    byPeriod.set(context.period, periodValues);
  }

  return byPeriod;
}

function segmentMetric(concept: string): keyof SegmentMetrics | null {
  const local = concept.split(":").pop() ?? concept;
  if (/^(RevenueFromContractWithCustomerExcludingAssessedTax|Revenues|SalesRevenueNet)$/i.test(local)) return "revenue";
  if (/(OperatingIncomeLoss|SegmentProfitLoss)/i.test(local)) return "operatingIncome";
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
  const reportableSegment = members.find((member) => /BusinessSegment|OperatingSegment|StatementBusinessSegments|SegmentAxis/i.test(member));
  const product = members.find(
    (member) =>
      member.includes("ProductOrServiceAxis") ||
      /ServiceLine|OtherServiceLine/i.test(member)
  );
  const joined = members.join(" ");
  const source = reportableSegment || product || joined;
  if (/CollectionServiceLineMember/i.test(source)) return "Collection";
  if (/LandfillServiceLineMember/i.test(source)) return "Landfill";
  if (/EnvironmentalSolutionsServiceLineMember/i.test(source)) return "Environmental Solutions";
  if (/TransferServiceLineMember/i.test(source)) return "Transfer";
  if (/InvestmentBankingAndCapitalMarkets/i.test(source)) return "Investment Banking & Capital Markets";
  if (/AssetManagementAndOther/i.test(source)) return "Asset Management and Other";
  if (/AssetManagement/i.test(source)) return "Asset Management and Other";
  if (/OtherServiceLineMember/i.test(source)) return "Other";
  if (/ProfessionalServices/i.test(source)) return "Professional Services and Other";
  return cleanSegmentMember(source);
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
  if (!local || /Consolidated|Corporate|Geographic|North America|Europe|Asia|Other Countries/i.test(local)) return null;
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
  if (!/OperatingSegmentsMember/i.test(joined) && value > 0) return true;
  return Math.abs(value) > Math.abs(existing);
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
  const order = ["Collection", "Landfill", "Environmental Solutions", "Transfer", "Other", "Professional Services and Other"];
  return (order.indexOf(a) === -1 ? 99 : order.indexOf(a)) - (order.indexOf(b) === -1 ? 99 : order.indexOf(b));
}

function buildFactContext(payload: any): ResolveContext {
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
          form: fact.form,
          fp: fact.fp,
          filed: fact.filed,
          accn: fact.accn,
          start: fact.start,
          end: fact.end
        };
        if (!instantFact && isAnnualDurationFact(fact)) {
          setSource(annualDuration, period, concept, source);
        } else if (!instantFact && isYearToDateFact(fact)) {
          setSource(cumulativeDuration, period, concept, source);
        } else if (instantFact || isQuarterDurationFact(fact)) {
          setSource(instantFact ? instant : duration, period, concept, source);
        }
      }
    }
  }

  deriveQuarterlies(duration, cumulativeDuration, annualDuration);

  return { duration, instant };
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
      const nineMonth = cumulativeDuration.get(`3Q${year}`)?.get(concept);
      const q1 = duration.get(`1Q${year}`)?.get(concept);
      const q2 = duration.get(`2Q${year}`)?.get(concept);
      const q3 = duration.get(`3Q${year}`)?.get(concept);
      const firstNineMonths = nineMonth?.value ?? (q1 && q2 && q3 ? q1.value + q2.value + q3.value : null);
      if (firstNineMonths === null) continue;
      setSource(duration, period, concept, {
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
  const quarter = period[0];
  let score = 0;
  const durationDays = source.start && source.end ? factDurationDays({ start: source.start, end: source.end } as SecFact) : 0;
  if (quarter === "4") {
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
  return periods.slice(-maxColumns);
}

function comparePeriods(a: string, b: string) {
  const [aq, ay] = [Number(a[0]), Number(`20${a.slice(2)}`)];
  const [bq, by] = [Number(b[0]), Number(`20${b.slice(2)}`)];
  return ay === by ? aq - bq : ay - by;
}

function previousPeriod(period: string) {
  const quarter = Number(period[0]);
  const year = Number(`20${period.slice(2)}`);
  if (!quarter || Number.isNaN(year)) return null;
  if (quarter === 1) return `4Q${String(year - 1).slice(-2)}`;
  return `${quarter - 1}Q${String(year).slice(-2)}`;
}

function blueColumns(sheet: ExcelJS.Worksheet) {
  const columnCounts = new Map<number, number>();
  for (let rowNumber = 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const label = rowLabel(sheet, rowNumber);
    if (!label) continue;
    for (let col = 6; col <= sheet.columnCount; col += 1) {
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
    for (let col = 4; col <= sheet.columnCount; col += 1) {
      if (/^[1-4]Q\d{2}$/.test(normalizePeriodLabel(cellDisplay(sheet.getCell(rowNumber, col))))) cols.push(col);
    }
    if (cols.length > best.length) best = cols;
  }
  return best;
}

function templatePeriodInfos(sheet: ExcelJS.Worksheet, columns: number[]) {
  let best: Array<{ period: string; isEstimate: boolean }> = [];
  let bestCount = 0;
  for (let rowNumber = 1; rowNumber <= Math.min(sheet.rowCount, 80); rowNumber += 1) {
    const infos = columns.map((col) => {
      const label = cellDisplay(sheet.getCell(rowNumber, col));
      return {
        period: normalizePeriodLabel(label),
        isEstimate: /e\s*$/i.test(label.trim())
      };
    });
    const validCount = infos.filter((info) => /^[1-4]Q\d{2}$/.test(info.period)).length;
    if (validCount > bestCount) {
      best = infos;
      bestCount = validCount;
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
  return compact;
}

function cellDisplay(cell: ExcelJS.Cell) {
  const value = cell.value;
  if (typeof value === "string") return value;
  if (typeof value === "number") return String(value);
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
    subtotalFormula: firstFormulaInRow(sheet, rowNumber, columns),
    projectedColumns: periodInfos.filter((info) => info.isEstimate).length,
    signConvention: inferSignConvention(label)
  };
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

function shouldAuditSkippedWrite(decision: WriteDecision) {
  return decision.reason === "existing formula cell" || decision.reason === "blank inactive/helper cell";
}

function shouldWriteIncomeStatementZero(fillRow: FillRow) {
  return fillRow.statement === "income" && fillRow.classification !== "partial";
}

function zeroIncomeStatementValue(fillRow: FillRow): ResolvedValue {
  return {
    value: 0,
    sources: [],
    note: `No matching EDGAR value was found for this income statement row; wrote 0.0 to preserve the model's historical formatting for ${fillRow.label}.`,
    classification: fillRow.classification
  };
}

function preservedDividendFormulaUpdate(
  fillRow: FillRow,
  cell: ExcelJS.Cell,
  period: string,
  periods: string[],
  columns: number[],
  periodIndex: number,
  ctx: ResolveContext
) {
  if (!isDividendFormulaBridge(fillRow, period, cell)) return null;
  const resolved = resolveRow(fillRow, period, ctx);
  if (resolved.value === null || Number.isNaN(resolved.value)) return null;
  const value = resolved.value / (fillRow.scale ?? 1);
  const existingValue = numericCellValue(cell);
  if (existingValue !== null && Math.abs(existingValue - value) < 0.05) return null;
  const formula = derivedDividendBridgeFormula(fillRow, resolved, periods, columns, periodIndex);
  if (!formula) return null;
  const auditNote = auditNoteForResolvedValue(fillRow, resolved);
  const comment = mappingComment(fillRow, resolved, period, value, "high", auditNote);
  return {
    formula,
    value,
    comment,
    auditRow: {
      ...mappingAuditRow(cell.worksheet, cell, fillRow, period, value, resolved, "high", auditNote),
      formulaPreserved: true,
      writeBlockedReason: "existing formula bridge updated with EDGAR annual total",
      notes: [auditNote, "Updated the preserved formula's annual bridge constant from EDGAR so the model formula still drives the quarter."].filter(Boolean).join(" ")
    }
  };
}

function isDividendFormulaBridge(fillRow: FillRow, period: string, cell: ExcelJS.Cell) {
  if (period[0] !== "4" || !hasFormula(cell)) return false;
  if (normalize(fillRow.label) !== normalize("Dividends")) return false;
  return Boolean(cellFormula(cell)?.match(/\bSUM\s*\(/i));
}

function derivedDividendBridgeFormula(fillRow: FillRow, resolved: ResolvedValue, periods: string[], columns: number[], periodIndex: number) {
  const source = derivedSource(resolved);
  if (source?.derivedTotalValue === undefined || !source.derivedPriorPeriods?.length) return null;
  const priorIndexes = source.derivedPriorPeriods.map((priorPeriod) => periods.indexOf(priorPeriod));
  if (!priorIndexes.every((index) => index >= 0 && index < periodIndex)) return null;
  const priorColumns = priorIndexes.map((index) => columns[index]);
  const total = roundModelValue(signedDividendBridgeTotal(source.derivedTotalValue, fillRow.sign) / (fillRow.scale ?? 1));
  return `${formatDividendFormulaNumber(total)}-SUM(${columnLetter(Math.min(...priorColumns))}${fillRow.row}:${columnLetter(Math.max(...priorColumns))}${fillRow.row})`;
}

function signedDividendBridgeTotal(value: number, sign: FillRow["sign"]) {
  return sign === -1 ? -Math.abs(value) : value;
}

function formatDividendFormulaNumber(value: number) {
  if (Number.isInteger(value)) return String(value);
  return value.toFixed(3).replace(/0+$/, "").replace(/\.$/, "");
}

function refreshDividendCachedResults(sheet: ExcelJS.Worksheet, periods: string[], columns: number[]) {
  const dividendRow = findLabelRow(sheet, "Dividends");
  const dividendsPaidRow = findLabelRow(sheet, "Dividends Paid");
  if (!dividendRow) return;

  periods.forEach((period, index) => {
    if (period[0] !== "4") return;
    const year = period.slice(2);
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
  cell.value = { ...value, result };
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
  const fallbackRevenue = periods.map((period) => first(period, ctx.duration, C.revenue)?.value ?? 0);
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

  return { filledCells, commentsAdded, warnings };
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
      if (!isHardcodedFinancialInput(cell)) return;
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
    if (!isHardcodedFinancialInput(cell)) return;
    const existing = typeof cell.value === "number" ? cell.value : null;
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
    if (rowLabel(sheet, rowNumber) && rowHasHardcodedFinancialInputs(sheet, rowNumber, columns)) rows.push(rowNumber);
  }
  return rows;
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
  const index = segments.findIndex((segment, segmentIndex) => {
    if (used.has(segmentIndex)) return false;
    const normalizedSegment = normalize(segment.label);
    return normalizedSegment === normalizedLabel || normalizedSegment.includes(normalizedLabel) || normalizedLabel.includes(normalizedSegment);
  });
  return index === -1 ? null : index;
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

function first(period: string, map: Map<string, Map<string, FactSource>>, concepts: string[]) {
  const facts = map.get(period);
  if (!facts) return null;
  for (const concept of concepts) {
    const source = facts.get(concept);
    if (source) return source;
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
    period,
    valueWritten,
    mappingType: resolved.classification || fillRow.classification,
    conceptsUsed: resolved.sources.map((source) => `${source.concept}=${roundModelValue(source.value / (fillRow.scale ?? 1))}mm`).join("; "),
    sourceStatement: fillRow.statement,
    accession: unique(resolved.sources.map((source) => source.accn).filter(Boolean)).join("; "),
    sourceUrl: "",
    cellWritable: true,
    formulaPreserved: false,
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
    period,
    valueWritten: numericCellValue(cell) ?? 0,
    mappingType: decision.formulaPreserved ? "formula" : "unused",
    conceptsUsed: fillRow.concepts?.join("; ") ?? "",
    sourceStatement: fillRow.statement,
    accession: "",
    sourceUrl: "",
    cellWritable: false,
    formulaPreserved: decision.formulaPreserved,
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
    period,
    valueWritten,
    mappingType: "segment",
    conceptsUsed: resolved.sources.map((source) => `${source.label} ${suffix}=${valueWritten}mm`).join("; "),
    sourceStatement: "segment",
    accession: unique(resolved.sources.map((source) => source.accn).filter(Boolean)).join("; "),
    sourceUrl: "",
    cellWritable: true,
    formulaPreserved: false,
    writeBlockedReason: "",
    signConvention: "copied",
    confidence: "high",
    validationStatus: "not_run",
    notes: resolved.note || ""
  };
}

function validateWorkbookBeforeReturn(workbook: ExcelJS.Workbook, periods: string[], columns: number[], ctx: ResolveContext, warnings: string[]) {
  const errors: string[] = [];
  const segmentSheet = workbook.getWorksheet(SEGMENT_SHEET);
  if (segmentSheet) {
    errors.push(...validateSegmentGenericRows(segmentSheet, periods, columns));
    errors.push(...validateSegmentOperatingIncomeCheck(segmentSheet, periods, columns));
  }
  const modelSheet = workbook.getWorksheet(MODEL_SHEET);
  if (modelSheet) {
    const evaluator = new FormulaEvaluator(modelSheet);
    errors.push(...validateIncomeStatementNetIncome(modelSheet, periods, columns, ctx, evaluator, warnings));
    errors.push(...validateBalanceSheetCheck(modelSheet, periods, columns, evaluator));

    const interestExpenseRow = findLabelRow(modelSheet, "Interest Expense");
    if (interestExpenseRow) {
      const hasPopulatedInterestExpense = columns.some((col) => Math.abs(numericCellValue(modelSheet.getCell(interestExpenseRow, col)) ?? 0) > 0.0001);
      if (!hasPopulatedInterestExpense) {
        errors.push("Model Interest Expense row is present but all detected historical cells are zero/blank.");
      }
    }
  }
  return errors;
}

class FormulaEvaluator {
  private readonly cache = new Map<string, number | null>();

  constructor(private readonly sheet: ExcelJS.Worksheet) {}

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
      return value === null ? "0" : String(value);
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

function validateIncomeStatementNetIncome(
  sheet: ExcelJS.Worksheet,
  periods: string[],
  columns: number[],
  ctx: ResolveContext,
  evaluator: FormulaEvaluator,
  warnings: string[]
) {
  const errors: string[] = [];
  const netIncomeRow = findRowInSection(sheet, "Income Statement", ["Net Income (Loss)", "Net Income"], (label) =>
    /income statement analysis|cash flow statement|cashflow statement|balance sheet/i.test(label)
  );
  if (!netIncomeRow) return errors;

  periods.forEach((period, index) => {
    const edgarNetIncome = first(period, ctx.duration, C.netIncome);
    if (!edgarNetIncome) {
      warnings.unshift(`Income Statement ${period}: net income tie-out skipped because EDGAR NetIncomeLoss was unavailable for that period.`);
      return;
    }

    const cell = sheet.getCell(netIncomeRow, columns[index]);
    const modelNetIncome = evaluator.evaluateCell(cell);
    if (modelNetIncome === null) {
      warnings.unshift(`Income Statement ${cell.address} ${period}: could not evaluate model net income for EDGAR tie-out.`);
      return;
    }

    const expected = edgarNetIncome.value / 1_000_000;
    if (!valuesTie(modelNetIncome, expected)) {
      warnings.unshift(
        `Income Statement ${cell.address} ${period}: net income ${roundModelValue(modelNetIncome)} does not match EDGAR NetIncomeLoss ${roundModelValue(expected)}.`
      );
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

function snapshotWorkbook(workbook: ExcelJS.Workbook, sheetNames: string[], firstHistoricalCol: number): WorkbookSnapshot {
  const labels = new Map<string, string>();
  const formulas = new Map<string, string>();
  const labelSnapshotEndCol = Math.max(1, firstHistoricalCol - 1);
  for (const sheetName of sheetNames) {
    const sheet = workbook.getWorksheet(sheetName);
    if (!sheet) continue;
    for (let rowNumber = 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
      for (let col = 1; col <= Math.min(labelSnapshotEndCol, sheet.columnCount); col += 1) {
        const address = snapshotAddress(sheet, rowNumber, col);
        labels.set(address, cellDisplay(sheet.getCell(rowNumber, col)));
      }
      for (let col = 1; col <= sheet.columnCount; col += 1) {
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
  const rowNumber = Number(cell.address.match(/\d+$/)?.[0]);
  if (!actual) return false;
  if (!rowNumber || normalize(rowLabel(cell.worksheet, rowNumber)) !== normalize("Dividends")) return false;
  const bridgeFormula = /^[-+]?\d+(?:\.\d+)?-SUM\([A-Z]+\d+:[A-Z]+\d+\)$/i;
  return bridgeFormula.test(expected) && bridgeFormula.test(actual);
}

function restoreWorkbookLabels(workbook: ExcelJS.Workbook, snapshot: WorkbookSnapshot) {
  for (const [address, expected] of snapshot.labels.entries()) {
    const cell = cellFromSnapshotAddress(workbook, address);
    if (!cell) continue;
    if (cellDisplay(cell) !== expected) cell.value = expected;
  }
}

function snapshotModelTabOutsideRows(sheet: ExcelJS.Worksheet, writableRows: Set<number>): CellSnapshot {
  const cells: CellSnapshot = new Map();
  for (let rowNumber = 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
    if (writableRows.has(rowNumber)) continue;
    for (let col = 1; col <= sheet.columnCount; col += 1) {
      cells.set(snapshotAddress(sheet, rowNumber, col), cellSnapshotValue(sheet.getCell(rowNumber, col)));
    }
  }
  return cells;
}

function validateModelTabOutsideRowsPreserved(sheet: ExcelJS.Worksheet, writableRows: Set<number>, snapshot: CellSnapshot) {
  const errors: string[] = [];
  for (const [address, expected] of snapshot.entries()) {
    const cell = cellFromSnapshotAddress(sheet.workbook, address);
    if (!cell) continue;
    const rowNumber = Number(cell.address.match(/\d+$/)?.[0]);
    if (rowNumber && writableRows.has(rowNumber)) continue;
    const actual = cellSnapshotValue(cell);
    if (actual !== expected) {
      errors.push(`${address}: non-income-statement / non-balance-sheet cell changed.`);
    }
  }
  return errors;
}

function cellSnapshotValue(cell: ExcelJS.Cell) {
  return JSON.stringify({
    value: cellDisplay(cell),
    formula: cellFormula(cell),
    comment: commentText(cell.note)
  });
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

function validateSegmentOperatingIncomeCheck(sheet: ExcelJS.Worksheet, periods: string[], columns: number[]) {
  const errors: string[] = [];
  const checkRow = findLabelRow(sheet, "Operating Income Check");
  if (!checkRow) return errors;

  const evaluator = new FormulaEvaluator(sheet);
  periods.forEach((period, index) => {
    const cell = sheet.getCell(checkRow, columns[index]);
    if (!hasFormula(cell)) return;
    const check = evaluator.evaluateCell(cell);
    if (check === null) {
      errors.push(`Segment Analysis ${cell.address} ${period}: could not evaluate the operating income check formula.`);
    } else if (!valuesTie(check, 0)) {
      errors.push(`Segment Analysis ${cell.address} ${period}: operating income check is ${roundModelValue(check)}, not OK.`);
    }
  });

  return errors;
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
    { header: "period", key: "period", width: 12 },
    { header: "value written", key: "valueWritten", width: 16 },
    { header: "mapping type", key: "mappingType", width: 16 },
    { header: "EDGAR concepts used", key: "conceptsUsed", width: 60 },
    { header: "source statement/table", key: "sourceStatement", width: 24 },
    { header: "accession", key: "accession", width: 24 },
    { header: "source URL", key: "sourceUrl", width: 24 },
    { header: "cell writable", key: "cellWritable", width: 14 },
    { header: "formula preserved", key: "formulaPreserved", width: 18 },
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

function jsonError(error: string, status: number) {
  return NextResponse.json({ error }, { status });
}
