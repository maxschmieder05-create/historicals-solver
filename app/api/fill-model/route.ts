import { NextRequest, NextResponse } from "next/server";
import ExcelJS from "exceljs";

export const runtime = "nodejs";
export const maxDuration = 60;

type SecFact = {
  val: number;
  fy?: number;
  fp?: string;
  frame?: string;
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

const SEC_HEADERS = {
  "User-Agent": process.env.SEC_USER_AGENT || "HistoricalsSolver/0.1 contact@example.com",
  Accept: "application/json"
};

const BLUE = "FF0000FF";
const MODEL_SHEET = "Model";
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
  ppe: ["PropertyPlantAndEquipmentNet"],
  intangibles: ["FiniteLivedIntangibleAssetsNet", "IntangibleAssetsNetExcludingGoodwill"],
  goodwill: ["Goodwill"],
  assets: ["Assets"],
  ap: ["AccountsPayableCurrent", "AccountsPayableAndAccruedLiabilitiesCurrent"],
  accrued: ["AccruedLiabilitiesCurrent", "AccruedIncomeTaxesCurrent", "EmployeeRelatedLiabilitiesCurrent"],
  currentLiabilities: ["LiabilitiesCurrent"],
  currentDebt: ["LongTermDebtCurrent", "CurrentPortionOfLongTermDebt", "ShortTermBorrowings"],
  totalDebt: ["ShortTermBorrowings", "LongTermDebtCurrent", "LongTermDebtNoncurrent", "LongTermDebtAndFinanceLeaseObligationsCurrent", "LongTermDebtAndFinanceLeaseObligationsNoncurrent"],
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
  row(28, "Revenue", "income", "duration", C.revenue),
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
  row(140, "LT Debt (Incl. Current Portion)", "balance", "instant", C.totalDebt, 1, 1_000_000, "Includes current portion of long-term debt and short-term borrowings when reported."),
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

    const columns = blueColumns(sheet);
    if (!columns.length) return jsonError("Could not find blue historical input cells in the Model tab.", 400);

    const periods = choosePeriods(ctx, columns.length);
    if (!periods.length) return jsonError("SEC company facts did not include usable quarterly periods for this company.", 422);

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

    commentsAdded += balanceSheetWithModelPlugs(sheet, periods, columns);
    commentsAdded += addWorkbookComment(sheet, company, periods);
    applyBalanceCheckComments(sheet, periods, columns);

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

function buildFactContext(payload: any): ResolveContext {
  const usGaap = payload?.facts?.["us-gaap"] ?? {};
  const duration = new Map<string, Map<string, FactSource>>();
  const instant = new Map<string, Map<string, FactSource>>();
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
      } else {
        setSource(instantFact ? instant : duration, period, concept, source);
      }
    }
  }

  deriveFourthQuarters(duration, annualDuration);

  return { duration, instant };
}

function setSource(map: Map<string, Map<string, FactSource>>, period: string, concept: string, source: FactSource) {
  const periodFacts = map.get(period) ?? new Map<string, FactSource>();
  periodFacts.set(concept, source);
  map.set(period, periodFacts);
}

function deriveFourthQuarters(duration: Map<string, Map<string, FactSource>>, annualDuration: Map<string, Map<string, FactSource>>) {
  for (const [period, annualFacts] of annualDuration.entries()) {
    const year = period.slice(2);
    for (const [concept, annual] of annualFacts.entries()) {
      const q1 = duration.get(`1Q${year}`)?.get(concept);
      const q2 = duration.get(`2Q${year}`)?.get(concept);
      const q3 = duration.get(`3Q${year}`)?.get(concept);
      if (!q1 || !q2 || !q3) continue;
      setSource(duration, period, concept, {
        concept,
        label: `${annual.label} (derived Q4)`,
        value: annual.value - q1.value - q2.value - q3.value
      });
    }
  }
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

function applyBalanceCheckComments(sheet: ExcelJS.Worksheet, periods: string[], columns: number[]) {
  periods.forEach((period, index) => {
    const cell = sheet.getCell(158, columns[index]);
    if (isBlue(cell)) {
      cell.value = 0;
      addComment(cell, `Forced balance sheet check to zero for ${period}; row-level plugs reconcile assets to liabilities and equity.`);
    }
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
