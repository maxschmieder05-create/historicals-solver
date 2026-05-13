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
  statement: "income" | "balance" | "support";
  kind: "duration" | "instant";
  scale?: number;
  sign?: 1 | -1;
  concepts?: string[];
  resolver?: (period: string, ctx: ResolveContext) => ResolvedValue;
  comment?: string;
};

type ResolvedValue = {
  value: number | null;
  sources: FactSource[];
  note?: string;
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

const SEC_HEADERS = {
  "User-Agent": process.env.SEC_USER_AGENT || "HistoricalsSolver/0.1 contact@example.com",
  Accept: "application/json"
};

const BLUE_FONT_COLORS = new Set(["FF0000FF", "FF0070C0", "FF0563C1", "FF0000EE"]);
const MODEL_SHEET = "Model";
const SEGMENT_SHEET = "Segment Analysis";
const LABEL_COLUMNS = [1, 2, 3, 4, 5];

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
  "NetIncomeLossAttributableToNonredeemableNoncontrollingInterest",
  "NetIncomeLossAttributableToRedeemableNoncontrollingInterest",
  "NetIncomeLossAttributableToNoncontrollingInterest"
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
  comment?: string
): FillRow {
  return { row: rowNumber, label, statement, kind, concepts, sign, scale, comment };
}

function plug(
  rowNumber: number,
  label: string,
  statement: FillRow["statement"],
  kind: FillRow["kind"],
  resolver: FillRow["resolver"]
): FillRow {
  return { row: rowNumber, label, statement, kind, resolver, scale: 1_000_000 };
}

function discoverFillRows(sheet: ExcelJS.Worksheet, columns: number[]) {
  const rows: FillRow[] = [];
  const seen = new Set<number>();

  for (let rowNumber = 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
    if (!columns.some((col) => isHardcodedFinancialInput(sheet.getCell(rowNumber, col)))) continue;
    const label = rowLabel(sheet, rowNumber);
    if (!label || seen.has(rowNumber)) continue;
    const fillRow = fillRowForLabel(rowNumber, label);
    if (!fillRow) continue;
    rows.push(fillRow);
    seen.add(rowNumber);
  }

  return rows;
}

function fillRowForLabel(rowNumber: number, label: string): FillRow | null {
  const key = normalize(label);
  const has = (...aliases: string[]) => aliases.some((alias) => key === normalize(alias));
  const includes = (...aliases: string[]) => aliases.some((alias) => key.includes(normalize(alias)));
  const hasRevenue = has("Revenue", "Revenues", "Total Revenue", "Total Revenues", "Sales", "Net Sales", "Net Revenue");
  const hasNetRevenue = has("Net Revenue", "Revenue Net of Interest Expense", "Revenues Net of Interest Expense");

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
    return plug(rowNumber, label, "income", "duration", resolveCompensationExpense);
  }
  if (has("Non-Compensation Expenses")) return plug(rowNumber, label, "income", "duration", resolveNonCompensationExpense);
  if (has("Depreciation & Amortization", "Depreciation and Amortization", "Depreciation Expense", "Depreciation & Amortization (incl. in SG&A)")) return row(rowNumber, label, "income", "duration", C.da, -1);
  if (has("Amortization Expense")) return row(rowNumber, label, "support", "duration", ["AmortizationOfIntangibleAssets"], -1);
  if (has("Other Operating Income (Expense)")) return plug(rowNumber, label, "income", "duration", resolveOtherOperatingIncomeExpense);
  if (has("Total Provisions for Credit Losses", "Provision for Credit Losses")) return row(rowNumber, label, "income", "duration", C.creditLossProvision, -1);
  if (has("Interest Income")) return row(rowNumber, label, "income", "duration", C.interestIncome, 1, 1_000_000, "Mapped to SEC interest income.");
  if (has("Interest (Expense)")) return row(rowNumber, label, "income", "duration", []);
  if (has("Interest Expense")) return plug(rowNumber, label, "income", "duration", resolveInterestExpense);
  if (has("Goodwill Impairment")) return row(rowNumber, label, "income", "duration", C.impairment, -1);
  if (has("Gain on Sale of Business (Loss)")) return row(rowNumber, label, "income", "duration", ["GainLossOnSaleOfBusiness", "GainLossOnSaleOfAssets"], 1);
  if (has("Other Non-Operating Income (Expense)", "Other Nonoperating Income (Expense)", "Other Income (Expense)", "Other Expense (Income)")) return row(rowNumber, label, "income", "duration", C.otherNonOp);
  if (has("Income Tax Benefit (Expense)", "Income Tax Expense", "Income Tax Provision (Expense)", "Income Tax")) return row(rowNumber, label, "income", "duration", C.taxes, -1);
  if (has("Net Unrealized Debt Securities Gains (Losses)")) return row(rowNumber, label, "income", "duration", C.unrealizedDebtSecurities);
  if (has("FX Adjustments")) return row(rowNumber, label, "income", "duration", C.foreignCurrencyAdjustments);
  if (has("Net Unrealized Pension and Other Benefits")) return row(rowNumber, label, "income", "duration", C.pensionAdjustments);
  if (has("Pre-Tax Adjustments")) return row(rowNumber, label, "income", "duration", []);
  if (has("Post-Tax Adjustments")) return row(rowNumber, label, "income", "duration", ["PreferredStockDividendsIncomeStatementImpact", "ConvertiblePreferredDividendsNetOfTax"], -1);
  if (has("Discontinued Operations")) return row(rowNumber, label, "income", "duration", ["IncomeLossFromDiscontinuedOperationsNetOfTax"]);
  if (has("Income (Loss) due to Non-Controlling Interest", "Income Loss Due To Non Controlling Interest")) {
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
  if (has("Capital Expenditures", "Capex")) return row(rowNumber, label, "support", "duration", C.capex, -1);
  if (has("Purchases of Intangibles")) return row(rowNumber, label, "support", "duration", ["PaymentsToAcquireIntangibleAssets"]);
  if (has("Purchases of Investments")) return row(rowNumber, label, "support", "duration", PURCHASES_OF_INVESTMENTS_CONCEPTS);
  if (has("Acquisition / (Divestment) of Businesses", "Proceeds From/(Acquisitions of) Businesses", "Proceeds From (Acquisitions of) Businesses")) {
    return row(rowNumber, label, "support", "duration", ACQUISITION_CONCEPTS);
  }
  if (has("Issuance/(Repayment) of Revolver")) return plug(rowNumber, label, "support", "duration", resolveRevolverIssuanceRepayment);
  if (has("Issuance of Debt")) return row(rowNumber, label, "support", "duration", C.debtIssuance);
  if (has("(Repayment of Debt)", "Repayment of Debt")) return row(rowNumber, label, "support", "duration", C.debtRepayment, -1);
  if (has("Issuance of Equity")) return row(rowNumber, label, "support", "duration", C.equityIssuance);
  if (has("Shares Repurchased ($ Amount)", "Share Repurchases ($ Amount)", "(Repurchase) of Equity", "Repurchase of Equity")) return row(rowNumber, label, "support", "duration", C.repurchases, -1);
  if (has("Stock-Based Comp Expense", "Stock-Based Compensation")) return row(rowNumber, label, "support", "duration", C.sbc);
  if (has("Dividends", "Dividends Issued")) return row(rowNumber, label, "support", "duration", C.dividends, -1);
  if (has("Change in Noncontrolling Interests")) return row(rowNumber, label, "support", "duration", C.noncontrollingInterestChange);
  if (has("Effect of FX Rate Changes on Cash")) return row(rowNumber, label, "support", "duration", C.fxCashEffect);
  if (has("Beginning Cash Adjustments", "Ending Cash Adjustments")) return row(rowNumber, label, "support", "duration", []);
  if (has("Beginning Cash Balance")) return plug(rowNumber, label, "support", "instant", resolveBeginningCashBalance);
  if (has("Ending Cash Balance")) return row(rowNumber, label, "support", "instant", C.cash);
  if (has("Weighted Average Basic Shares", "Basic Shares")) return row(rowNumber, label, "support", "duration", C.basicShares, 1, 1_000_000);
  if (has("Weighted Average Dilutive Shares", "Weighted Average Diluted Shares", "Diluted Shares")) return row(rowNumber, label, "support", "duration", C.dilutedShares, 1, 1_000_000);

  return null;
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
    "Included floor brokerage and clearing fees, underwriting costs, technology and communications, occupancy and equipment rental, business development, and professional services from the SEC non-interest expense table."
  );
  if (detail.value !== null) return signed(detail, -1) ?? detail;

  const noninterestExpense = first(period, ctx.duration, ["NoninterestExpense"]);
  const compensation = resolveCompensationExpense(period, ctx);
  const depreciation = first(period, ctx.duration, C.da) ?? zeroSource(C.da[0]);
  const otherOperating = resolveOtherOperatingIncomeExpense(period, ctx);
  if (noninterestExpense && compensation.value !== null && otherOperating.value !== null) {
    return {
      value: -Math.abs(noninterestExpense.value - Math.abs(compensation.value) - depreciation.value - Math.abs(otherOperating.value)),
      sources: compactSources([noninterestExpense, compensation, depreciation, otherOperating]),
      note: "Calculated from SEC non-interest expense less compensation, depreciation and amortization, and other operating expense detail."
    };
  }
  return { value: null, sources: [], note: "No non-compensation expense detail was available in SEC facts." };
}

function resolveOtherOperatingIncomeExpense(period: string, ctx: ResolveContext): ResolvedValue {
  const detail = sumWithNote(
    period,
    ctx.duration,
    OTHER_OPERATING_EXPENSE_CONCEPTS,
    "Included cost of sales and other expenses from the SEC non-interest expense table."
  );
  if (detail.value !== null) return signed(detail, -1) ?? detail;
  const direct = first(period, ctx.duration, ["OtherOperatingIncomeExpenseNet"]);
  return direct ? { value: direct.value, sources: [direct] } : { value: null, sources: [] };
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
    const inlineCtx = await fetchInlineFactContext(company, periods);
    mergeContexts(ctx, inlineCtx);
    const segmentRevenue = await fetchSegmentRevenueByPeriod(company, periods);
    const fillRows = discoverFillRows(sheet, columns);
    if (!fillRows.length) return jsonError("Could not match the Model tab's blue input rows to supported financial statement labels.", 422);

    const warnings: string[] = [];
    let filledCells = 0;
    let commentsAdded = 0;

    for (const fillRow of fillRows) {
      const rowNotes = new Set<string>();
      let unresolved = 0;
      periods.forEach((period, index) => {
        const col = columns[index];
        const cell = sheet.getCell(fillRow.row, col);
        if (!isHardcodedFinancialInput(cell)) return;

        const resolved = resolveRow(fillRow, period, ctx);
        if (resolved.value === null || Number.isNaN(resolved.value)) {
          unresolved += 1;
          return;
        }

        cell.value = resolved.value / (fillRow.scale ?? 1);
        filledCells += 1;

        const auditNote = auditNoteForResolvedValue(fillRow, resolved);
        if (auditNote) rowNotes.add(auditNote);
      });

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
      const segmentResult = fillSegmentAnalysis(segmentSheet, company, periods, columns, segmentRevenue, ctx);
      filledCells += segmentResult.filledCells;
      commentsAdded += segmentResult.commentsAdded;
      warnings.push(...segmentResult.warnings);
    } else {
      warnings.push(`Could not find a "${SEGMENT_SHEET}" worksheet; Model revenue formulas were left untouched.`);
    }

    for (const fillRow of fillRows) {
      const hasAny = periods.some((_, index) => sheet.getCell(fillRow.row, columns[index]).value !== null);
      if (!hasAny && fillRow.concepts?.length) {
        warnings.push(`${fillRow.label}: no matching SEC concept found.`);
      }
    }

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
      contexts.set(match[1], { period: periodKeyFromDate(instant), instant: true, end: instant });
      continue;
    }
    const start = body.match(/<xbrli:startDate>([^<]+)<\/xbrli:startDate>/)?.[1];
    const end = body.match(/<xbrli:endDate>([^<]+)<\/xbrli:endDate>/)?.[1];
    if (start && end) {
      contexts.set(match[1], { period: periodKeyFromDate(end), instant: false, start, end });
    }
  }
  return contexts;
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
  const product = members.find(
    (member) =>
      member.includes("ProductOrServiceAxis") ||
      /ServiceLine|OtherServiceLine|BusinessSegment|OperatingSegment|StatementBusinessSegments|SegmentAxis/i.test(member)
  );
  const joined = members.join(" ");
  const source = product || joined;
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
  if (existing && existing.value !== 0 && source.value === 0) {
    map.set(period, periodFacts);
    return;
  }
  if (existing && existing.value === 0 && source.value !== 0) {
    periodFacts.set(concept, source);
    map.set(period, periodFacts);
    return;
  }
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
  for (const col of LABEL_COLUMNS) {
    const text = cellDisplay(sheet.getCell(rowNumber, col)).trim();
    if (text) return text;
  }
  return "";
}

function derivedQuarterFormula(fillRow: FillRow, resolved: ResolvedValue, periods: string[], columns: number[], periodIndex: number) {
  const source = derivedSource(resolved);
  if (source?.derivedTotalValue === undefined || !source.derivedPriorPeriods?.length) return null;

  const priorIndexes = source.derivedPriorPeriods.map((priorPeriod) => periods.indexOf(priorPeriod));
  if (!priorIndexes.every((index) => index >= 0 && index < periodIndex)) return null;

  const priorColumns = priorIndexes.map((index) => columns[index]);
  const total = roundModelValue(signedDerivedTotal(source.derivedTotalValue, fillRow.sign) / (fillRow.scale ?? 1));
  return `${formatFormulaNumber(total)}-SUM(${columnLetter(Math.min(...priorColumns))}${fillRow.row}:${columnLetter(Math.max(...priorColumns))}${fillRow.row})`;
}

function derivedSource(resolved: ResolvedValue) {
  return resolved.sources.find((source) => source.derivedTotalValue !== undefined && source.derivedPriorPeriods?.length);
}

function signedDerivedTotal(value: number, sign: FillRow["sign"]) {
  return sign === -1 ? -Math.abs(value) : value;
}

function formatFormulaNumber(value: number) {
  return Number.isInteger(value) ? String(value) : value.toFixed(1);
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
  ctx: ResolveContext
) {
  const warnings: string[] = [];
  let filledCells = 0;
  const commentsAdded = 0;
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

  const revenueResult = fillSegmentMetricRows(sheet, periods, columns, rows, usableSegments, "values", "Revenue");
  const operatingIncomeResult = fillSegmentMetricRows(sheet, periods, columns, operatingIncomeRows, usableSegments, "operatingIncome", "Operating Income");
  const depreciationResult = fillSegmentMetricRows(sheet, periods, columns, depreciationRows, usableSegments, "depreciationAmortization", "D&A");
  filledCells += revenueResult.filledCells + operatingIncomeResult.filledCells + depreciationResult.filledCells;

  return { filledCells, commentsAdded, warnings };
}

function fillSegmentMetricRows(
  sheet: ExcelJS.Worksheet,
  periods: string[],
  columns: number[],
  rows: number[],
  segments: SegmentRevenue[],
  metric: "values" | "operatingIncome" | "depreciationAmortization",
  suffix: string
) {
  let filledCells = 0;
  const usedSegments = new Set<number>();

  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    const rowNumber = rows[rowIndex];
    const existingLabel = segmentBaseLabel(cellDisplay(sheet.getCell(rowNumber, 3)), suffix);
    if (isGenericSegmentPlaceholder(existingLabel)) continue;
    const segmentIndex = segmentIndexForRow(existingLabel, segments, usedSegments) ?? nextUnusedSegmentIndex(segments, usedSegments);
    if (segmentIndex === null) continue;

    const segment = segments[segmentIndex];
    if (!segmentHasMetricData(segment, metric, periods)) continue;
    usedSegments.add(segmentIndex);

    periods.forEach((period, periodIndex) => {
      const cell = sheet.getCell(rowNumber, columns[periodIndex]);
      if (!isHardcodedFinancialInput(cell)) return;
      cell.value = (segment[metric].get(period) ?? 0) / 1_000_000;
      filledCells += 1;
    });
  }

  return { filledCells };
}

function segmentHasMetricData(segment: SegmentRevenue, metric: "values" | "operatingIncome" | "depreciationAmortization", periods: string[]) {
  return periods.some((period) => segment[metric].has(period) && segment[metric].get(period) !== 0);
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
    .replace(new RegExp(`\\s*${suffix.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}\\s*$`, "i"), "")
    .replace(/\s*Revenue\s*$/i, "")
    .replace(/\s*Operating Income\s*$/i, "")
    .replace(/\s*D&A\s*$/i, "")
    .trim();
}

function isGenericSegmentPlaceholder(label: string) {
  return !label || /^segment\s*\d+$/i.test(label);
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

function findLabelRow(sheet: ExcelJS.Worksheet, label: string) {
  const wanted = normalize(label);
  for (let rowNumber = 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
    if (normalize(rowLabel(sheet, rowNumber)) === wanted) return rowNumber;
  }
  return null;
}

function labelCell(sheet: ExcelJS.Worksheet, rowNumber: number) {
  for (const col of LABEL_COLUMNS) {
    const cell = sheet.getCell(rowNumber, col);
    if (cellDisplay(cell).trim()) return cell;
  }
  return sheet.getCell(rowNumber, 3);
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

function canAddComment(cell: ExcelJS.Cell) {
  return !hasFormula(cell);
}

function hasFormula(cell: ExcelJS.Cell) {
  const value = cell.value;
  return Boolean(value && typeof value === "object" && ("formula" in value || "sharedFormula" in value));
}

function auditNoteForResolvedValue(fillRow: FillRow, resolved: ResolvedValue) {
  const derived = derivedSource(resolved);
  const sourceLabels = unique(resolved.sources.map((source) => source.note || source.label || source.concept).filter(Boolean));
  const grouped = sourceLabels.length > 1;
  const isDerived = Boolean(derived?.derivedTotalValue !== undefined);
  const hasAnalystNote = Boolean((resolved.note || fillRow.comment) && (grouped || isDerived));
  if (!grouped && !hasAnalystNote && !isDerived) return null;

  const parts: string[] = [];
  if (resolved.note || fillRow.comment) parts.push(resolved.note || fillRow.comment || "");
  if (grouped) parts.push(`Included EDGAR line items: ${sourceLabels.join(", ")}.`);
  if (isDerived && derived?.derivedPriorPeriods?.length) {
    parts.push(`Quarterly value was derived from ${derived.derivedTotalLabel || derived.label} less ${derived.derivedPriorPeriods.join(", ")}.`);
  }
  return parts.filter(Boolean).join(" ");
}

function resolveRow(fillRow: FillRow, period: string, ctx: ResolveContext): ResolvedValue {
  if (fillRow.resolver) return fillRow.resolver(period, ctx);
  if (!fillRow.concepts?.length) return { value: 0, sources: [zeroSource(fillRow.label)] };
  const source = first(period, fillRow.kind === "instant" ? ctx.instant : ctx.duration, fillRow.concepts);
  if (!source) return { value: null, sources: [] };
  return signed(source, fillRow.sign ?? 1) ?? { value: null, sources: [] };
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
  return result ? { ...result, note } : { value: null, sources: [], note };
}

function sumResolved(items: Array<ResolvedValue | null | undefined>, note: string): ResolvedValue {
  const valid = items.filter((item): item is ResolvedValue => Boolean(item && item.value !== null));
  if (!valid.length) return { value: null, sources: [], note };
  return {
    value: valid.reduce((total, item) => total + (item.value ?? 0), 0),
    sources: compactSources(valid),
    note
  };
}

function difference(period: string, map: Map<string, Map<string, FactSource>>, totalConcepts: string[], lessConceptGroups: string[][], note: string): ResolvedValue {
  const total = first(period, map, totalConcepts);
  if (!total) return { value: null, sources: [], note };
  const less = lessConceptGroups.map((concepts) => first(period, map, concepts) ?? zeroSource(concepts[0]));
  return {
    value: total.value - less.reduce((acc, source) => acc + source.value, 0),
    sources: [total, ...less],
    note
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

function addComment(cell: ExcelJS.Cell, text: string) {
  const existing = typeof cell.note === "string" ? `${cell.note}\n\n` : "";
  cell.note = `${existing}Historicals Solver: ${text}`;
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
