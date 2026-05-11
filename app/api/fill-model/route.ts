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
const PERIOD_HEADER_ROWS = [25, 63, 73, 118, 160, 190, 204, 216, 234, 275, 286];

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

const ROWS: FillRow[] = [
  row(29, "Cost of Goods Sold", "income", "duration", C.cogs, -1),
  row(32, "Selling, General & Administration (SG&A)", "income", "duration", C.sga, -1),
  row(33, "Research & Development (R&D)", "income", "duration", C.rd, -1),
  row(34, "Depreciation & Amortization", "income", "duration", C.da, -1),
  row(35, "Other Operating Income (Expense)", "income", "duration", ["OtherOperatingIncomeExpenseNet"]),
  row(38, "Interest Income", "income", "duration", C.interestIncome),
  row(39, "Interest (Expense)", "income", "duration", C.interestExpense, -1),
  row(40, "Goodwill Impairment", "income", "duration", C.impairment, -1),
  row(41, "Other Non-Operating Income (Expense)", "income", "duration", C.otherNonOp),
  row(44, "Income Tax Benefit (Expense)", "income", "duration", C.taxes, -1),
  row(47, "Pre-Tax Adjustments", "income", "duration", []),
  row(48, "Post-Tax Adjustments", "income", "duration", []),
  row(49, "Discontinued Operations", "income", "duration", ["IncomeLossFromDiscontinuedOperationsNetOfTax"]),
  row(50, "Income (Loss) due to Non-Controlling Interest", "income", "duration", ["NetIncomeLossAttributableToNoncontrollingInterest"], -1),
  row(120, "Cash & Cash Equivalents", "balance", "instant", C.cash),
  row(121, "Accounts Receivable", "balance", "instant", C.receivables),
  row(122, "Inventory", "balance", "instant", C.inventory),
  plug(123, "Prepaid & Other Current Assets", "balance", "instant", (period, ctx) =>
    difference(period, ctx.instant, C.currentAssets, [C.cash, C.receivables, C.inventory], "Plug: AssetsCurrent less cash, receivables, and inventory.")
  ),
  row(126, "PP&E, Net", "balance", "instant", C.ppe),
  row(127, "Intangible Assets, Net", "balance", "instant", C.intangibles),
  row(128, "Goodwill", "balance", "instant", C.goodwill),
  plug(129, "Other Non-Current Assets", "balance", "instant", (period, ctx) =>
    difference(period, ctx.instant, C.assets, [C.currentAssets, C.ppe, C.intangibles, C.goodwill], "Plug: Total assets less current assets, PP&E, intangibles, and goodwill.")
  ),
  row(134, "Accounts Payable", "balance", "instant", C.ap),
  row(135, "Accrued Liabilities", "balance", "instant", C.accrued),
  plug(136, "Other Current Liabilities", "balance", "instant", (period, ctx) =>
    difference(period, ctx.instant, C.currentLiabilities, [C.ap, C.accrued, C.currentDebt], "Plug: Current liabilities less AP, accrued liabilities, and current debt.")
  ),
  row(139, "Revolver", "balance", "instant", ["ShortTermBorrowings"]),
  {
    row: 140,
    label: "LT Debt (Incl. Current Portion)",
    statement: "balance",
    kind: "instant",
    scale: 1_000_000,
    comment: "Includes current portion of long-term debt and short-term borrowings when reported.",
    resolver: (period, ctx) => {
      const aggregate = first(period, ctx.instant, [
        "LongTermDebtAndCapitalLeaseObligationsIncludingCurrentMaturities",
        "LongTermDebtAndFinanceLeaseObligations",
        "LongTermDebtAndCapitalLeaseObligations",
        "LongTermDebt"
      ]);
      if (aggregate) return { value: aggregate.value, sources: [aggregate] };
      return sum(period, ctx.instant, [
        "ShortTermBorrowings",
        "LongTermDebtCurrent",
        "LongTermDebtNoncurrent",
        "LongTermDebtAndFinanceLeaseObligationsCurrent",
        "LongTermDebtAndFinanceLeaseObligationsNoncurrent",
        "LongTermDebtAndCapitalLeaseObligationsCurrent"
      ]) ?? { value: null, sources: [] };
    }
  },
  row(141, "Deferred Income Taxes", "balance", "instant", C.deferredTaxLiability),
  plug(142, "Other Non-Current Liabilities", "balance", "instant", (period, ctx) => {
    const assets = first(period, ctx.instant, C.assets);
    const currentLiabExDebt = difference(period, ctx.instant, C.currentLiabilities, [C.currentDebt], "");
    const debt = sum(period, ctx.instant, C.totalDebt);
    const dtl = first(period, ctx.instant, C.deferredTaxLiability) ?? zeroSource("DeferredTaxLiabilitiesNoncurrent");
    const equity = first(period, ctx.instant, C.equity);
    const nci = first(period, ctx.instant, C.nci) ?? zeroSource("NoncontrollingInterestInConsolidatedEntity");
    if (!assets || !currentLiabExDebt.value || !equity) return { value: null, sources: [], note: "Could not calculate balance-sheet plug because assets, liabilities, or equity were unavailable." };
    return {
      value: assets.value - currentLiabExDebt.value - (debt?.value ?? 0) - dtl.value - equity.value - nci.value,
      sources: compactSources([assets, ...currentLiabExDebt.sources, debt, dtl, equity, nci]),
      note: "Balance plug: total assets less current liabilities, debt, deferred taxes, shareholder equity, and noncontrolling interests."
    };
  }),
  plug(147, "Common Stock & APIC", "balance", "instant", (period, ctx) => {
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
      note: "Plug: stockholders' equity less retained earnings, treasury stock, and AOCI."
    };
  }),
  row(148, "Retained Earnings", "balance", "instant", C.retained),
  plug(149, "Treasury Stock", "balance", "instant", (period, ctx) => signed(first(period, ctx.instant, C.treasury), -1) ?? { value: 0, sources: [zeroSource("TreasuryStockValue")] }),
  row(150, "Accumulated Other Comprehensive Income (AOCI)", "balance", "instant", C.aoci),
  row(153, "Noncontrolling Interests", "balance", "instant", C.nci),
  row(193, "Capital Expenditures", "support", "duration", C.capex),
  row(194, "Depreciation Expense", "support", "duration", C.da, -1),
  row(207, "Purchases of Intangibles", "support", "duration", ["PaymentsToAcquireIntangibleAssets"]),
  row(208, "Amortization Expense", "support", "duration", ["AmortizationOfIntangibleAssets"], -1),
  row(241, "Shares Repurchased ($ Amount)", "support", "duration", C.repurchases),
  row(246, "Stock-Based Comp Expense", "support", "duration", C.sbc),
  row(254, "Dividends", "support", "duration", C.dividends, -1),
  row(282, "Weighted Average Basic Shares", "support", "duration", C.basicShares, 1, 1_000_000),
  row(284, "Weighted Average Dilutive Shares", "support", "duration", C.dilutedShares, 1, 1_000_000)
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
    normalizeSharedFormulas(workbook);

    const sheet = workbook.getWorksheet(MODEL_SHEET);
    if (!sheet) return jsonError(`Could not find a "${MODEL_SHEET}" worksheet in this workbook.`, 400);

    let columns = blueColumns(sheet);
    if (!columns.length) return jsonError("Could not find blue historical input cells in the Model tab.", 400);

    let periods = templatePeriods(sheet, columns);
    if (periods.length === columns.length) {
      const pairs = periods
        .map((period, index) => ({ period, col: columns[index] }))
        .filter(({ period }) => ctx.duration.has(period) || ctx.instant.has(period));
      periods = pairs.map((pair) => pair.period);
      columns = pairs.map((pair) => pair.col);
    } else {
      periods = choosePeriods(ctx, columns.length);
      columns = columns.slice(0, periods.length);
    }
    if (!periods.length) return jsonError("SEC company facts did not include usable quarterly periods for this company.", 422);
    const segmentRevenue = await fetchSegmentRevenueByPeriod(company, periods);

    const warnings: string[] = [];
    let filledCells = 0;
    let commentsAdded = 0;

    for (const headerRow of PERIOD_HEADER_ROWS) {
      periods.forEach((period, index) => {
        const cell = sheet.getCell(headerRow, columns[index]);
        cell.value = period;
      });
    }

    for (const fillRow of ROWS) {
      const rowNotes = new Set<string>();
      periods.forEach((period, index) => {
        const col = columns[index];
        const cell = sheet.getCell(fillRow.row, col);

        const resolved = resolveRow(fillRow, period, ctx);
        if (resolved.value === null || Number.isNaN(resolved.value)) {
          cell.value = 0;
          return;
        }

        cell.value = roundModelValue(resolved.value / (fillRow.scale ?? 1));
        filledCells += 1;

        const sourceLabels = resolved.sources.map((source) => source.label || source.concept);
        if (resolved.note) rowNotes.add(resolved.note);
        if (sourceLabels.length > 1 || fillRow.comment) {
          rowNotes.add(fillRow.comment || `Included SEC concepts: ${unique(sourceLabels).join(", ")}.`);
        }
      });

      if (rowNotes.size) {
        addComment(sheet.getCell(fillRow.row, 3), Array.from(rowNotes).join(" "));
        commentsAdded += 1;
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

    commentsAdded += balanceSheetWithModelPlugs(sheet, periods, columns);
    commentsAdded += addWorkbookComment(sheet, company, periods);
    addBalanceCheckComments(sheet, periods, columns);

    for (const fillRow of ROWS) {
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
    .filter((filing: FilingRef) => ["10-Q", "10-K"].includes(filing.form) && filing.primaryDocument?.endsWith(".htm"))
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
        if (filing.form === "10-K") annual.set(period, values);
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
  if (form !== "10-K" && days > 115) return null;
  const year = endDate.getUTCFullYear();
  const quarter = Math.floor(endDate.getUTCMonth() / 3) + 1;
  if (form === "10-K") return `4Q${String(year).slice(-2)}`;
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
      const period = periodKey(fact, instantFact);
      if (!period) continue;
      const source = { concept, label, value: fact.val };
      if (!instantFact && fact.fp === "FY") {
        setSource(annualDuration, period, concept, source);
      } else if (!instantFact && isYearToDateFact(fact)) {
        setSource(cumulativeDuration, period, concept, source);
      } else {
        setSource(instantFact ? instant : duration, period, concept, source);
      }
    }
  }

  deriveQuarterlies(duration, cumulativeDuration, annualDuration);

  return { duration, instant };
}

function setSource(map: Map<string, Map<string, FactSource>>, period: string, concept: string, source: FactSource) {
  const periodFacts = map.get(period) ?? new Map<string, FactSource>();
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
          setSource(duration, period, concept, { ...source, label: `${source.label} (derived Q2)`, value: source.value - q1.value });
        }
      });
    } else if (quarter === 3) {
      cumulativeFacts.forEach((source, concept) => {
        const q2Cumulative = cumulativeDuration.get(`2Q${year}`)?.get(concept);
        if (q2Cumulative && !duration.get(period)?.get(concept)) {
          setSource(duration, period, concept, { ...source, label: `${source.label} (derived Q3)`, value: source.value - q2Cumulative.value });
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
        value: annual.value - firstNineMonths
      });
    }
  }
}

function isYearToDateFact(fact: SecFact) {
  if (!fact.start || !fact.end) return false;
  const days = (new Date(`${fact.end}T00:00:00Z`).getTime() - new Date(`${fact.start}T00:00:00Z`).getTime()) / 86_400_000;
  return days > 115 && fact.fp !== "FY";
}

function isUsableFact(fact: SecFact) {
  return typeof fact.val === "number" && Boolean(fact.end) && ["10-K", "10-Q", "20-F", "40-F"].includes(fact.form ?? "");
}

function isInstantFact(fact: SecFact) {
  return Boolean(fact.frame?.endsWith("I")) || !["Q1", "Q2", "Q3", "Q4", "FY"].includes(fact.fp ?? "");
}

function periodKey(fact: SecFact, instantFact: boolean) {
  const fy = fact.fy;
  if (!fy) return null;
  const yy = String(fy).slice(-2);
  if (fact.fp === "Q1") return `1Q${yy}`;
  if (fact.fp === "Q2") return `2Q${yy}`;
  if (fact.fp === "Q3") return `3Q${yy}`;
  if (fact.fp === "Q4") return `4Q${yy}`;
  if (fact.fp === "FY") return instantFact ? `4Q${yy}` : `4Q${yy}`;
  return null;
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
  const anchorRows = [29, 120, 134, 140];
  for (const rowNumber of anchorRows) {
    const cols: number[] = [];
    for (let col = 6; col <= sheet.columnCount; col += 1) {
      if (isBlue(sheet.getCell(rowNumber, col))) cols.push(col);
    }
    if (cols.length) return cols;
  }
  return [];
}

function templatePeriods(sheet: ExcelJS.Worksheet, columns: number[]) {
  const periods = columns.map((col) => cellDisplay(sheet.getCell(25, col)).replace(/e$/i, ""));
  return periods.every((period) => /^[1-4]Q\d{2}$/.test(period)) ? periods : [];
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

function normalizeSharedFormulas(workbook: ExcelJS.Workbook) {
  workbook.worksheets.forEach((sheet) => {
    sheet.eachRow((row) => {
      row.eachCell((cell) => {
        const formula = cell.formula;
        if (!formula) return;
        cell.value = { formula, result: cell.result };
      });
    });
  });
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
  const rows = [8, 9, 10, 11, 12, 13];
  const labelRows = [16, 17, 18, 19, 20, 21];

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
    const label = segment?.label ?? `Segment ${index + 1}`;
    sheet.getCell(labelRows[index], 3).value = label;

    periods.forEach((period, periodIndex) => {
      const col = columns[periodIndex];
      const cell = sheet.getCell(rows[index], col);
      cell.value = roundModelValue((segment?.values.get(period) ?? 0) / 1_000_000);
      filledCells += 1;
    });
  }

  addComment(
    sheet.getCell("C16"),
    segments.length
      ? `Revenue segment labels and historical revenue values were populated from service-line dimensional facts in recent SEC 10-Q/10-K inline XBRL filings.`
      : `Revenue is using a one-line fallback because service-line dimensional revenue was not available in recent SEC filings.`
  );
  commentsAdded += 1;

  return { filledCells, commentsAdded, warnings };
}

function isBlue(cell: ExcelJS.Cell) {
  return cell.font?.color?.argb === BLUE;
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

function addComment(cell: ExcelJS.Cell, text: string) {
  const existing = typeof cell.note === "string" ? `${cell.note}\n\n` : "";
  cell.note = `${existing}Historicals Solver: ${text}`;
}

function addWorkbookComment(sheet: ExcelJS.Worksheet, company: CompanyMatch, periods: string[]) {
  addComment(sheet.getCell("C25"), `Filled from SEC company facts for ${company.title} (${company.ticker}) across ${periods.join(", ")}. Dollar values are in millions except per-share data.`);
  return 1;
}

function balanceSheetWithModelPlugs(sheet: ExcelJS.Worksheet, periods: string[], columns: number[]) {
  periods.forEach((_, index) => {
    const col = columns[index];
    const totalAssets = sumRows(sheet, col, [120, 121, 122, 123, 126, 127, 128, 129]);
    const currentLiabilitiesExDebt = sumRows(sheet, col, [134, 135, 136]);
    const debtAndTaxes = sumRows(sheet, col, [139, 140, 141]);
    const equity = sumRows(sheet, col, [147, 148, 149, 150, 153]);
    sheet.getCell(142, col).value = roundModelValue(totalAssets - currentLiabilitiesExDebt - debtAndTaxes - equity);
  });
  addComment(sheet.getCell("C142"), "Final balance plug recalculated from the model rows so Total Assets equals Total Liabilities & Shareholder's Equity after SEC mappings.");
  return 1;
}

function sumRows(sheet: ExcelJS.Worksheet, col: number, rows: number[]) {
  return rows.reduce((total, rowNumber) => total + numericCell(sheet.getCell(rowNumber, col)), 0);
}

function numericCell(cell: ExcelJS.Cell) {
  const value = cell.value;
  if (typeof value === "number") return value;
  if (value && typeof value === "object" && "result" in value && typeof value.result === "number") return value.result;
  return 0;
}

function addBalanceCheckComments(sheet: ExcelJS.Worksheet, periods: string[], columns: number[]) {
  periods.forEach((period, index) => {
    const cell = sheet.getCell(158, columns[index]);
    addComment(cell, `Balance sheet formulas are preserved for ${period}; historical input plugs reconcile assets to liabilities and equity.`);
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
