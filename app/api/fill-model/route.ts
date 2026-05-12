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

type SegmentRevenue = {
  label: string;
  values: Map<string, number>;
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

const BLUE = "FF0000FF";
const MODEL_SHEET = "Model";
const SEGMENT_SHEET = "Segment Analysis";

const C = {
  revenue: ["RevenueFromContractWithCustomerExcludingAssessedTax", "Revenues", "SalesRevenueNet"],
  cogs: ["CostOfRevenue", "CostOfGoodsAndServicesSold", "CostOfGoodsAndServiceExcludingDepreciationDepletionAndAmortization"],
  sga: ["SellingGeneralAndAdministrativeExpense", "GeneralAndAdministrativeExpense"],
  rd: ["ResearchAndDevelopmentExpense"],
  da: ["DepreciationDepletionAndAmortization", "DepreciationAndAmortization"],
  interestIncome: ["InterestIncomeExpenseNonOperatingNet", "InterestIncomeNonOperating"],
  interestExpense: ["InterestExpenseNonOperating", "InterestExpense"],
  impairment: ["GoodwillImpairmentLosses", "ImpairmentOfGoodwillAndIntangibleAssets"],
  otherNonOp: ["OtherNonoperatingIncomeExpense", "NonoperatingIncomeExpense"],
  taxes: ["IncomeTaxExpenseBenefit"],
  netIncome: ["NetIncomeLoss"],
  sbc: ["ShareBasedCompensation"],
  capex: ["PaymentsToAcquirePropertyPlantAndEquipment"],
  dividends: ["PaymentsOfDividends", "PaymentsOfDividendsCommonStock"],
  repurchases: ["PaymentsForRepurchaseOfCommonStock"],
  basicShares: ["WeightedAverageNumberOfSharesOutstandingBasic"],
  dilutedShares: ["WeightedAverageNumberOfDilutedSharesOutstanding"],
  cash: ["CashAndCashEquivalentsAtCarryingValue", "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents"],
  receivables: ["AccountsReceivableNetCurrent"],
  inventory: ["InventoryNet"],
  currentAssets: ["AssetsCurrent"],
  ppe: ["PropertyPlantAndEquipmentAndFinanceLeaseRightOfUseAssetAfterAccumulatedDepreciationAndAmortization", "PropertyPlantAndEquipmentNet"],
  intangibles: ["FiniteLivedIntangibleAssetsNet", "IntangibleAssetsNetExcludingGoodwill"],
  goodwill: ["Goodwill"],
  assets: ["Assets"],
  ap: ["AccountsPayableCurrent", "AccountsPayableAndAccruedLiabilitiesCurrent"],
  accrued: ["AccruedLiabilitiesCurrent", "AccruedIncomeTaxesCurrent", "EmployeeRelatedLiabilitiesCurrent"],
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
    if (!columns.some((col) => isHardcodedBlueInput(sheet.getCell(rowNumber, col)))) continue;
    const label = cellDisplay(sheet.getCell(rowNumber, 3)).trim();
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

  if (has("Cost of Goods Sold", "Cost of Goods & Services Sold", "Cost of Revenue")) return row(rowNumber, label, "income", "duration", C.cogs, -1);
  if (has("Selling, General & Administration (SG&A)", "Selling General & Administrative", "SG&A")) return row(rowNumber, label, "income", "duration", C.sga, -1);
  if (has("Research & Development (R&D)", "Research and Development")) return row(rowNumber, label, "income", "duration", C.rd, -1);
  if (has("Compensation and Benefits", "Compensation, Commissions, and Benefits")) {
    return plug(rowNumber, label, "income", "duration", resolveCompensationExpense);
  }
  if (has("Non-Compensation Expenses")) return row(rowNumber, label, "income", "duration", C.sga, -1);
  if (has("Depreciation & Amortization", "Depreciation and Amortization", "Depreciation Expense")) return row(rowNumber, label, "income", "duration", C.da, -1);
  if (has("Amortization Expense")) return row(rowNumber, label, "support", "duration", ["AmortizationOfIntangibleAssets"], -1);
  if (has("Other Operating Income (Expense)")) return row(rowNumber, label, "income", "duration", ["OtherOperatingIncomeExpenseNet"]);
  if (has("Interest Income")) return row(rowNumber, label, "income", "duration", C.interestIncome);
  if (has("Interest (Expense)", "Interest Expense")) return row(rowNumber, label, "income", "duration", C.interestExpense, -1);
  if (has("Goodwill Impairment")) return row(rowNumber, label, "income", "duration", C.impairment, -1);
  if (has("Other Non-Operating Income (Expense)", "Other Nonoperating Income (Expense)")) return row(rowNumber, label, "income", "duration", C.otherNonOp);
  if (has("Income Tax Benefit (Expense)", "Income Tax Expense")) return row(rowNumber, label, "income", "duration", C.taxes, -1);
  if (has("Pre-Tax Adjustments", "Post-Tax Adjustments")) return row(rowNumber, label, "income", "duration", []);
  if (has("Discontinued Operations")) return row(rowNumber, label, "income", "duration", ["IncomeLossFromDiscontinuedOperationsNetOfTax"]);
  if (has("Income (Loss) due to Non-Controlling Interest", "Income Loss Due To Non Controlling Interest")) return row(rowNumber, label, "income", "duration", ["NetIncomeLossAttributableToNoncontrollingInterest"], -1);

  if (has("Cash & Cash Equivalents", "Cash and Cash Equivalents")) return row(rowNumber, label, "balance", "instant", C.cash);
  if (has("Accounts Receivable", "Fees Receivable")) return row(rowNumber, label, "balance", "instant", C.receivables);
  if (has("Inventory")) return row(rowNumber, label, "balance", "instant", C.inventory);
  if (has("Prepaid & Other Current Assets", "Prepaid and Other Current Assets")) {
    return plug(rowNumber, label, "balance", "instant", (period, ctx) =>
      difference(period, ctx.instant, C.currentAssets, [C.cash, C.receivables, C.inventory], "Included current assets less separately modeled cash, receivables, and inventory.")
    );
  }
  if (has("PP&E, Net", "Property Plant and Equipment Net")) return row(rowNumber, label, "balance", "instant", C.ppe);
  if (has("Intangible Assets, Net")) return row(rowNumber, label, "balance", "instant", C.intangibles);
  if (has("Goodwill")) return row(rowNumber, label, "balance", "instant", C.goodwill);
  if (has("Investments and Assets of Consolidated VIEs", "Investments", "Investments and Assets of Consolidated Variable Interest Entities")) {
    return row(rowNumber, label, "balance", "instant", INVESTMENT_ASSET_CONCEPTS);
  }
  if (has("Other Non-Current Assets")) {
    return plug(rowNumber, label, "balance", "instant", (period, ctx) =>
      difference(period, ctx.instant, C.assets, [C.currentAssets, C.ppe, C.intangibles, C.goodwill, INVESTMENT_ASSET_CONCEPTS], "Included total assets less separately modeled current assets, PP&E, intangibles, goodwill, and investments.")
    );
  }
  if (has("Accounts Payable")) return row(rowNumber, label, "balance", "instant", C.ap);
  if (has("Accrued Liabilities")) return row(rowNumber, label, "balance", "instant", C.accrued);
  if (has("Other Current Liabilities")) {
    return plug(rowNumber, label, "balance", "instant", (period, ctx) =>
      difference(period, ctx.instant, C.currentLiabilities, [C.ap, C.accrued, C.currentDebt], "Included current liabilities less separately modeled accounts payable, accrued liabilities, and current debt.")
    );
  }
  if (has("Tax Receivable Agreement Payables")) {
    return row(rowNumber, label, "balance", "instant", ["TaxReceivableAgreementLiability", "TaxReceivableAgreementLiabilityCurrent", "OtherLiabilitiesCurrent"]);
  }
  if (has("Revolver", "Short Term Borrowings")) return row(rowNumber, label, "balance", "instant", ["ShortTermBorrowings"]);
  if (includes("LT Debt", "Long Term Debt")) {
    return plug(rowNumber, label, "balance", "instant", resolveTotalDebt);
  }
  if (has("Deferred Income Taxes")) return row(rowNumber, label, "balance", "instant", ["DeferredTaxAssetsNet", "DeferredTaxAssetsLiabilitiesNet", ...C.deferredTaxLiability]);
  if (has("Other Non-Current Liabilities")) return plug(rowNumber, label, "balance", "instant", resolveOtherNonCurrentLiabilities);
  if (has("Common Stock & APIC", "Common Stock and APIC")) return plug(rowNumber, label, "balance", "instant", resolveCommonStockAndApic);
  if (has("Retained Earnings")) return row(rowNumber, label, "balance", "instant", C.retained);
  if (has("Treasury Stock")) return plug(rowNumber, label, "balance", "instant", (period, ctx) => signed(first(period, ctx.instant, C.treasury), -1) ?? { value: 0, sources: [zeroSource("TreasuryStockValue")] });
  if (has("Accumulated Other Comprehensive Income (AOCI)", "AOCI")) return row(rowNumber, label, "balance", "instant", C.aoci);
  if (has("Noncontrolling Interests", "Non-Controlling Interests")) return row(rowNumber, label, "balance", "instant", C.nci);

  if (has("Capital Expenditures", "Capex")) return row(rowNumber, label, "support", "duration", C.capex);
  if (has("Purchases of Intangibles")) return row(rowNumber, label, "support", "duration", ["PaymentsToAcquireIntangibleAssets"]);
  if (has("Purchases of Investments")) return row(rowNumber, label, "support", "duration", PURCHASES_OF_INVESTMENTS_CONCEPTS);
  if (has("Acquisition / (Divestment) of Businesses", "Proceeds From/(Acquisitions of) Businesses", "Proceeds From (Acquisitions of) Businesses")) {
    return row(rowNumber, label, "support", "duration", ACQUISITION_CONCEPTS);
  }
  if (has("Shares Repurchased ($ Amount)", "Share Repurchases ($ Amount)")) return row(rowNumber, label, "support", "duration", C.repurchases);
  if (has("Stock-Based Comp Expense", "Stock-Based Compensation")) return row(rowNumber, label, "support", "duration", C.sbc);
  if (has("Dividends")) return row(rowNumber, label, "support", "duration", C.dividends, -1);
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
    const segmentRevenue = await fetchSegmentRevenueByPeriod(company, periods);
    const fillRows = discoverFillRows(sheet, columns);
    if (!fillRows.length) return jsonError("Could not match the Model tab's blue input rows to supported financial statement labels.", 422);

    const warnings: string[] = [];
    let filledCells = 0;
    let commentsAdded = 0;

    for (const fillRow of fillRows) {
      let unresolved = 0;
      periods.forEach((period, index) => {
        const col = columns[index];
        const cell = sheet.getCell(fillRow.row, col);
        if (!isHardcodedBlueInput(cell)) return;

        const resolved = resolveRow(fillRow, period, ctx);
        if (resolved.value === null || Number.isNaN(resolved.value)) {
          unresolved += 1;
          return;
        }

        const result = roundModelValue(resolved.value / (fillRow.scale ?? 1));
        cell.value = result;
        filledCells += 1;
      });

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
  const response = await fetch(`https://data.sec.gov/submissions/CIK${company.cik}.json`, { headers: SEC_HEADERS });
  if (!response.ok) return [];
  const submissions = await response.json();
  const recent = submissions?.filings?.recent;
  if (!recent) return [];

  const filings: FilingRef[] = recent.form
    .map((form: string, index: number) => ({
      form,
      filingDate: recent.filingDate[index],
      accessionNumber: recent.accessionNumber[index],
      primaryDocument: recent.primaryDocument[index]
    }))
    .filter((filing: FilingRef) => (isTenQ(filing.form) || isTenK(filing.form)) && filing.primaryDocument?.endsWith(".htm"))
    .slice(0, 28);

  const annual = new Map<string, Map<string, number>>();
  const quarterly = new Map<string, Map<string, number>>();

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
      values: new Map(periods.map((period) => [period, quarterly.get(period)?.get(label) ?? 0]))
    }));
}

function parseInlineSegmentRevenue(html: string, form: string) {
  const contexts = new Map<string, { period: string | null; members: string[] }>();
  for (const match of html.matchAll(/<xbrli:context\b[^>]*id="([^"]+)"[\s\S]*?<\/xbrli:context>/g)) {
    const body = match[0];
    const period = contextPeriod(body, form);
    const members = Array.from(body.matchAll(/<xbrldi:explicitMember\b[^>]*>([^<]+)<\/xbrldi:explicitMember>/g)).map((member) => member[1]);
    contexts.set(match[1], { period, members });
  }

  const byPeriod = new Map<string, Map<string, number>>();
  for (const match of html.matchAll(/<ix:nonFraction\b([^>]*)>([\s\S]*?)<\/ix:nonFraction>/g)) {
    const attrs = match[1];
    const concept = attrs.match(/\bname="([^"]+)"/)?.[1] ?? "";
    if (!concept.endsWith(":RevenueFromContractWithCustomerExcludingAssessedTax") && !concept.endsWith(":Revenues")) continue;

    const contextRef = attrs.match(/\bcontextRef="([^"]+)"/)?.[1];
    if (!contextRef) continue;
    const context = contexts.get(contextRef);
    if (!context?.period) continue;

    const label = segmentLabelFromMembers(context.members);
    if (!label) continue;

    const value = ixNumber(match[2], attrs);
    if (value === null) continue;

    const periodValues = byPeriod.get(context.period) ?? new Map<string, number>();
    const existing = periodValues.get(label);
    if (existing === undefined || preferSegmentFact(context.members, value, existing)) {
      periodValues.set(label, value);
    }
    byPeriod.set(context.period, periodValues);
  }

  return byPeriod;
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
  const product = members.find((member) => member.includes("ProductOrServiceAxis") || /ServiceLine|OtherServiceLine/i.test(member));
  const joined = members.join(" ");
  const source = product || joined;
  if (/CollectionServiceLineMember/i.test(source)) return "Collection";
  if (/LandfillServiceLineMember/i.test(source)) return "Landfill";
  if (/EnvironmentalSolutionsServiceLineMember/i.test(source)) return "Environmental Solutions";
  if (/TransferServiceLineMember/i.test(source)) return "Transfer";
  if (/OtherServiceLineMember/i.test(source)) return "Other";
  if (/ProfessionalServices/i.test(source)) return "Professional Services and Other";
  return null;
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

function deriveSegmentFourthQuarters(quarterly: Map<string, Map<string, number>>, annual: Map<string, Map<string, number>>) {
  for (const [period, annualValues] of annual.entries()) {
    const year = period.slice(2);
    const q1 = quarterly.get(`1Q${year}`);
    const q2 = quarterly.get(`2Q${year}`);
    const q3 = quarterly.get(`3Q${year}`);
    if (!q1 || !q2 || !q3) continue;
    const q4 = quarterly.get(period) ?? new Map<string, number>();
    annualValues.forEach((value, label) => {
      q4.set(label, value - (q1.get(label) ?? 0) - (q2.get(label) ?? 0) - (q3.get(label) ?? 0));
    });
    quarterly.set(period, q4);
  }
}

function segmentSort(a: string, b: string) {
  const order = ["Collection", "Landfill", "Environmental Solutions", "Transfer", "Other", "Professional Services and Other"];
  return (order.indexOf(a) === -1 ? 99 : order.indexOf(a)) - (order.indexOf(b) === -1 ? 99 : order.indexOf(b));
}

function buildFactContext(payload: any): ResolveContext {
  const usGaap = payload?.facts?.["us-gaap"] ?? {};
  const duration = new Map<string, Map<string, FactSource>>();
  const instant = new Map<string, Map<string, FactSource>>();
  const cumulativeDuration = new Map<string, Map<string, FactSource>>();
  const annualDuration = new Map<string, Map<string, FactSource>>();

  for (const [concept, detail] of Object.entries<any>(usGaap)) {
    const label = detail.label || concept;
    const units = detail.units ?? {};
    const unitFacts: SecFact[] = units.USD ?? units.shares ?? units["USD/shares"] ?? Object.values(units)[0] ?? [];
    for (const fact of unitFacts) {
      if (!isUsableFact(fact)) continue;
      const instantFact = isInstantFact(fact);
      const period = periodKey(fact);
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

function preferSource(period: string, next: FactSource, current: FactSource) {
  const nextScore = sourceScore(period, next);
  const currentScore = sourceScore(period, current);
  if (nextScore !== currentScore) return nextScore > currentScore;
  return (next.filed ?? "") > (current.filed ?? "");
}

function sourceScore(period: string, source: FactSource) {
  const quarter = period[0];
  let score = 0;
  if (quarter === "4") {
    if (isTenK(source.form)) score += 20;
    if (source.fp === "FY" || source.fp === "Q4") score += 10;
  } else {
    if (isTenQ(source.form)) score += 20;
    if (source.fp === `Q${quarter}`) score += 10;
  }
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

function blueColumns(sheet: ExcelJS.Worksheet) {
  const columnCounts = new Map<number, number>();
  for (let rowNumber = 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const label = cellDisplay(sheet.getCell(rowNumber, 3));
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

  return Array.from(columnCounts.keys()).sort((a, b) => a - b);
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
  let commentsAdded = 0;
  const fallbackRevenue = periods.map((period) => first(period, ctx.duration, C.revenue)?.value ?? 0);
  const rows = segmentRevenueRows(sheet).slice(0, 6);

  const usableSegments = segments.length
    ? segments
    : [
        {
          label: company.title.replace(/,?\s+INC\.?$/i, ""),
          values: new Map(periods.map((period, index) => [period, fallbackRevenue[index]]))
        }
      ];

  if (!segments.length) {
    warnings.push("No service-line revenue facts found in recent inline XBRL filings; Segment Analysis uses one total company revenue line.");
  }

  for (let index = 0; index < rows.length; index += 1) {
    const segment = usableSegments[index];
    periods.forEach((period, periodIndex) => {
      const col = columns[periodIndex];
      const cell = sheet.getCell(rows[index], col);
      if (!isHardcodedBlueInput(cell)) return;
      cell.value = roundModelValue((segment?.values.get(period) ?? 0) / 1_000_000);
      filledCells += 1;
    });
  }

  return { filledCells, commentsAdded, warnings };
}

function segmentRevenueRows(sheet: ExcelJS.Worksheet) {
  const totalRevenueRow = findLabelRow(sheet, "Total Company Revenue");
  const revenueMixRow = findLabelRow(sheet, "Revenue Mix");
  if (!totalRevenueRow || !revenueMixRow || revenueMixRow <= totalRevenueRow) return [8, 9, 10, 11, 12, 13];
  const rows: number[] = [];
  for (let rowNumber = totalRevenueRow + 1; rowNumber < revenueMixRow; rowNumber += 1) {
    if (isBlue(sheet.getCell(rowNumber, 3)) || rowHasBlueInputs(sheet, rowNumber)) rows.push(rowNumber);
  }
  return rows.length ? rows : [8, 9, 10, 11, 12, 13];
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

function findLabelRow(sheet: ExcelJS.Worksheet, label: string) {
  const wanted = normalize(label);
  for (let rowNumber = 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
    if (normalize(cellDisplay(sheet.getCell(rowNumber, 3))) === wanted) return rowNumber;
  }
  return null;
}

function rowHasBlueInputs(sheet: ExcelJS.Worksheet, rowNumber: number) {
  for (let col = 6; col <= sheet.columnCount; col += 1) {
    if (isHardcodedBlueInput(sheet.getCell(rowNumber, col))) return true;
  }
  return false;
}

function isBlue(cell: ExcelJS.Cell) {
  return cell.font?.color?.argb === BLUE;
}

function isHardcodedBlueInput(cell: ExcelJS.Cell) {
  return isBlue(cell) && !hasFormula(cell);
}

function hasFormula(cell: ExcelJS.Cell) {
  const value = cell.value;
  return Boolean(value && typeof value === "object" && ("formula" in value || "sharedFormula" in value));
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
