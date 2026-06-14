import { normalizeAccession } from "./sec-accession";

export type FinancialStatementName = "income_statement" | "balance_sheet" | "cash_flow" | "segment_analysis";

export type FinancialSourceTableType = "primary_statement" | "footnote" | "segment_table" | "roll_forward" | "cash_flow_reconciliation";

export type FinancialStatementSection =
  | "current assets"
  | "non-current assets"
  | "current liabilities"
  | "non-current liabilities"
  | "equity"
  | "revenue"
  | "operating expenses"
  | "below operating income"
  | "tax"
  | "net income"
  | "unknown";

export type FinancialLineItemClassificationRequest = {
  company: {
    name: string;
    ticker: string;
  };
  filing: {
    accession: string;
    form?: string;
    filingDate?: string;
  };
  fiscalPeriod: string;
  statement: FinancialStatementName;
  sourceTableType: FinancialSourceTableType;
  reportedLineItemLabel: string;
  cleanLabel: string;
  xbrlTag?: string;
  amount?: number | null;
  unit?: string;
  periodType: "instant" | "duration";
  section: FinancialStatementSection;
  nearbyRows: string[];
  parentSubtotal?: {
    label?: string;
    concept?: string;
  };
  isSubtotal: boolean;
  priorPeriodSourceLabels?: string[];
  currentPeriodSourceLines?: string[];
  availableModelRows: string[];
  modelRowDefinitions: Record<string, string>;
  deterministicCandidate?: string;
  uncertaintyReason: string;
  validationError?: string;
  alreadyMappedRows?: string[];
};

export type FinancialLineItemClassification = {
  source_line_item: string;
  recommended_action: "map" | "remap" | "set_zero" | "merge_into_other" | "split_across_rows" | "keep_existing" | "exclude";
  recommended_model_row: string;
  recommended_model_row_mappings: Array<{
    source_line_item: string;
    model_row: string;
    amount: number | null;
    reason: string;
  }>;
  explicit_zero_rows: Array<{
    model_row: string;
    reason: string;
  }>;
  classification_type: string;
  is_current: boolean | null;
  is_debt: boolean;
  is_operating: boolean | null;
  is_tax_related: boolean;
  is_deferred_revenue_or_contract_liability: boolean;
  is_deferred_tax: boolean;
  is_subtotal: boolean;
  should_exclude_from_other_bucket: boolean;
  confidence: "high" | "medium" | "low";
  reason: string;
  requires_validation: boolean;
  requires_revalidation: boolean;
  llm_used: boolean;
  mapping_passed_validation: boolean;
  warning?: string;
};

export type FinancialLineItemClassificationStore = Map<string, FinancialLineItemClassification>;

type LlmClassificationOptions = {
  enabled: boolean;
  apiKey: string;
  endpoint: string;
  model: string;
  siteUrl: string;
  appTitle: string;
  timeoutMs?: number;
  fetchImpl?: typeof fetch;
};

type ClassifierOptions = {
  llm?: LlmClassificationOptions;
};

export const MODEL_ROW_DEFINITIONS: Record<string, string> = {
  "Cash & Cash Equivalents":
    "Cash, cash equivalents, marketable securities, and short-term investments when no dedicated short-term/current investments row exists.",
  "Accounts Receivable": "Trade accounts receivable, accounts receivable, receivables net of allowances.",
  Inventory:
    "Inventories and inventory-like operating assets, including supplies, parts, merchandise inventory, raw materials, WIP, finished goods, aircraft fuel/spare parts/supplies.",
  "Prepaid & Other Current Assets":
    "Current assets not better mapped elsewhere, such as prepaid expenses, other current assets, tax receivables, contract assets if no better row exists.",
  "PP&E, Net": "Property, plant and equipment, operating property and equipment, net PP&E.",
  "Intangible Assets, Net": "Intangible assets, acquired intangibles, customer relationships, developed technology, intangible assets net.",
  Goodwill: "Goodwill.",
  "Other Non-Current Assets":
    "Non-current assets without better dedicated rows, such as operating lease ROU assets, long-term investments, long-term receivables, deferred contract costs, deferred tax assets if no better row exists, other assets.",
  "Accounts Payable": "Accounts payable, trade payables.",
  "Accrued Liabilities":
    "Accrued operating liabilities such as accrued expenses, accrued compensation, payroll, benefits, accrued operating costs, accrued taxes if no better row exists.",
  "Other Current Liabilities":
    "Current non-debt liabilities not better mapped elsewhere, such as deferred revenue, deferred income, contract liabilities, customer advances, current lease liabilities, taxes payable, other current liabilities.",
  "Total Current Liabilities (Excl. Debt)":
    "Current liabilities excluding current debt, short-term borrowings, revolver, and current maturities/current portion included in debt.",
  Revolver:
    "True short-term borrowing facilities/instruments, including revolver borrowings, revolving credit facility, short-term borrowings, commercial paper, current borrowings, notes payable current, short-term debt that is clearly a short-term borrowing facility.",
  "LT Debt (Incl. Current Portion)":
    "Long-term debt instruments including current portion/current maturities, senior notes, convertible senior notes, long-term debt net of current portion plus current portion, long-term debt and finance lease obligations.",
  "Deferred Income Taxes":
    "True deferred tax liabilities / deferred income taxes / deferred tax liabilities, non-current / deferred tax assets and liabilities net.",
  "Other Non-Current Liabilities":
    "Non-current liabilities without better dedicated rows, including long-term deferred revenue, long-term contract liabilities, operating lease liabilities non-current, pension and postretirement obligations, asset retirement obligations, other long-term liabilities.",
  "Common Stock & APIC": "Common stock, additional paid-in capital, capital in excess of par.",
  "Retained Earnings": "Retained earnings or retained deficit.",
  "Treasury Stock": "Treasury stock or contra-equity items such as employee benefit trust / ESOP trust if no better row exists.",
  AOCI: "Accumulated other comprehensive income/loss.",
  "Noncontrolling Interests": "Noncontrolling interests / minority interest.",
  Revenue: "Revenue, net revenue, net sales, operating revenue, total revenue.",
  "COGS / Cost of Goods Sold":
    "Direct cost of revenue / cost of sales / cost of operations, including industry-specific direct operating costs if presented as direct operating expenses and not better mapped elsewhere.",
  "SG&A": "Selling, general and administrative, marketing, sales and marketing, G&A, corporate overhead.",
  "R&D": "Research and development, technology development, product development, engineering expense.",
  "D&A": "Only standalone income-statement depreciation and amortization lines. Do not use cash-flow-only D&A.",
  "Other Operating Income / Expense":
    "Operating items between gross profit and operating income not better mapped elsewhere: restructuring, impairment, special charges, acquired IPR&D if not grouped into R&D, accretion, gain/loss on divestitures, other operating expense/income.",
  "Interest Income": "Standalone interest income.",
  "Interest Expense": "Standalone interest expense, interest expense net, interest and debt expense.",
  "Goodwill Impairment":
    "Only if goodwill impairment is explicitly reported as a primary income statement line or clearly part of the income statement bridge.",
  "Other Non-Operating Income / Expense":
    "Below-EBIT items not better mapped elsewhere, such as other income/expense net, equity method income/loss, investment gains/losses, FX gains/losses, gains/losses on equity investments.",
  "Income Tax Benefit / Expense": "Tax expense or benefit based on filing label and sign convention."
};

const AMBIGUOUS_LINE_ITEM_TERMS = [
  "deferred",
  "debt",
  "notes",
  "other",
  "accrued",
  "special",
  "impairment",
  "restructuring",
  "investment",
  "supplies",
  "contract",
  "financing",
  "lease",
  "tax",
  "income"
];

const MODEL_ROW_ALIASES: Record<string, string[]> = {
  "Cash & Cash Equivalents": ["cash", "cash and cash equivalents", "cash equivalents", "cash & short-term investments", "cash and short-term investments"],
  Inventory: ["inventory", "inventories"],
  "Prepaid & Other Current Assets": ["prepaid and other current assets", "prepaid & other current assets", "other current assets"],
  Revolver: ["revolver", "short-term debt", "short term debt", "short-term borrowings", "short term borrowings", "current borrowings"],
  "LT Debt (Incl. Current Portion)": ["lt debt incl current portion", "long-term debt", "long term debt", "borrowings", "senior notes", "total debt"],
  "Deferred Income Taxes": ["deferred income taxes", "deferred tax liabilities", "deferred taxes"],
  "Other Current Liabilities": ["other current liabilities", "other current liabs"],
  "Other Non-Current Liabilities": ["other non-current liabilities", "other long-term liabilities", "other lt liabilities"],
  "Common Stock & APIC": ["common stock & apic", "common stock and apic", "common stock and additional paid-in capital"],
  "Treasury Stock": ["treasury stock", "treasury & preferred stock"],
  "SG&A": ["sga", "sg&a", "selling general administrative", "sales and marketing", "general and administrative"],
  "R&D": ["research and development", "r&d", "research & development"],
  "D&A": ["depreciation and amortization", "depreciation & amortization", "d&a"],
  "Other Operating Income / Expense": ["other operating income expense", "other operating income", "other operating expense"],
  "Other Non-Operating Income / Expense": ["other non-operating income expense", "other income expense", "other expense income"]
};

export function modelRowDefinitionsForRows(availableRows: string[]) {
  const output: Record<string, string> = {};
  for (const [row, definition] of Object.entries(MODEL_ROW_DEFINITIONS)) {
    if (modelRowAvailable(row, availableRows) || rowIsCoreClassificationTarget(row)) output[row] = definition;
  }
  return output;
}

export function classificationSourceKeys(input: {
  period?: string;
  accession?: string;
  xbrlTag?: string;
  label?: string;
  amount?: number | null;
}) {
  const period = normalizeKey(input.period ?? "");
  const accession = normalizeAccession(input.accession ?? "");
  const concept = normalizeKey(input.xbrlTag ?? "");
  const label = normalizeKey(input.label ?? "");
  const amount = typeof input.amount === "number" && Number.isFinite(input.amount) ? String(Math.round(input.amount)) : "";
  const keys = [
    ["period", period, accession, concept, label, amount],
    ["period-concept-label", period, accession, concept, label],
    ["period-concept", period, accession, concept],
    ["accession-concept-label", accession, concept, label],
    ["concept-label", concept, label]
  ]
    .map((parts) => parts.filter(Boolean).join("|"))
    .filter(Boolean);
  return Array.from(new Set(keys));
}

export function lineItemNeedsClassification(request: FinancialLineItemClassificationRequest) {
  if (request.isSubtotal) return false;
  if (request.sourceTableType !== "primary_statement" && request.sourceTableType !== "cash_flow_reconciliation") return false;
  const text = requestSearchText(request);
  if (AMBIGUOUS_LINE_ITEM_TERMS.some((term) => text.includes(term))) return true;
  if (/short[-\s]?term|current maturit|current portion|senior notes?|convertible|contract liabilit|deferred revenue|deferred income|spare parts?|supplies|in[-\s]?process research|special items?|other/.test(text)) return true;
  if (request.deterministicCandidate && /other|accrued|revolver|deferred|d&a|depreciation|amortization/i.test(request.deterministicCandidate)) return true;
  return Boolean(request.uncertaintyReason);
}

export async function classifyFinancialLineItem(
  request: FinancialLineItemClassificationRequest,
  options: ClassifierOptions = {}
): Promise<FinancialLineItemClassification> {
  const deterministic = deterministicFinancialLineItemClassification(request);
  const fallback = deterministic ?? conservativeFallbackClassification(request);
  const deterministicIsValidated =
    deterministic?.confidence === "high" &&
    classificationPassesValidation(request, {
      ...deterministic,
      recommended_model_row: normalizeModelRow(deterministic.recommended_model_row) || deterministic.recommended_model_row
    });
  const shouldCallLlm = Boolean(
    options.llm?.enabled &&
      options.llm.apiKey &&
      options.llm.model &&
      lineItemNeedsClassification(request) &&
      !deterministicIsValidated
  );

  if (!shouldCallLlm) return finalizeClassification(request, fallback);

  try {
    const llm = options.llm!;
    const llmClassification = await requestLlmClassification(request, llm);
    return finalizeClassification(request, {
      ...llmClassification,
      source_line_item: llmClassification.source_line_item || request.cleanLabel || request.reportedLineItemLabel,
      recommended_model_row:
        normalizeModelRow(llmClassification.recommended_model_row) ||
        normalizeModelRow(fallback.recommended_model_row) ||
        fallback.recommended_model_row,
      confidence: llmClassification.confidence ?? "low",
      requires_validation: true,
      llm_used: true
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : "unknown classifier LLM error";
    return finalizeClassification(request, {
      ...fallback,
      confidence: lowerClassificationConfidence(fallback.confidence, "medium"),
      llm_used: false,
      warning: `LLM classification skipped (${message}).`
    });
  }
}

export function modelRowsMatch(a: string, b: string) {
  const left = normalizeKey(a);
  const right = normalizeKey(b);
  if (!left || !right) return false;
  if (left === right) return true;
  return equivalentModelRows(a).some((candidate) => normalizeKey(candidate) === right) || equivalentModelRows(b).some((candidate) => normalizeKey(candidate) === left);
}

export function modelRowAvailable(row: string, availableRows: string[]) {
  return availableRows.some((available) => modelRowsMatch(row, available));
}

function deterministicFinancialLineItemClassification(
  request: FinancialLineItemClassificationRequest
): FinancialLineItemClassification | null {
  const text = requestSearchText(request);
  const section = request.section;
  const current = section.includes("current") ? section.startsWith("current") : null;
  const base = baseClassification(request);
  const preferred = (...rows: string[]) => rows.find((row) => modelRowAvailable(row, request.availableModelRows)) ?? rows[0];
  const hasCurrentInvestmentsRow = request.availableModelRows.some((row) => /short[-\s]?term investments?|current investments?|marketable securities/i.test(row));

  if (/\bdeferred\b/.test(text) && /\btax(?:es)?\b/.test(text) && !/assets?/.test(section)) {
    return {
      ...base,
      recommended_model_row: preferred("Deferred Income Taxes"),
      classification_type: "deferred tax liability",
      is_current: current,
      is_tax_related: true,
      is_deferred_tax: true,
      should_exclude_from_other_bucket: true,
      confidence: "high",
      reason: "Deferred tax liabilities are true deferred income taxes and belong in the deferred tax row when available."
    };
  }

  if (/\bdeferred\b/.test(text) && /\btax(?:es)?\b/.test(text) && /assets?/.test(section)) {
    return {
      ...base,
      recommended_model_row: preferred("Other Non-Current Assets"),
      classification_type: "deferred tax asset",
      is_current: current,
      is_tax_related: true,
      is_deferred_tax: true,
      should_exclude_from_other_bucket: false,
      confidence: "high",
      reason: "Deferred tax assets are asset balances and belong in Other Non-Current Assets when the template has no dedicated deferred tax asset row."
    };
  }

  if (/\bdeferred (?:income|revenue)\b|\bunearned revenue\b|\bcontract liabilit|\bcustomer advances?\b/.test(text)) {
    return {
      ...base,
      recommended_model_row: section === "non-current liabilities" ? preferred("Other Non-Current Liabilities") : preferred("Other Current Liabilities"),
      classification_type: "deferred revenue or contract liability",
      is_current: current,
      is_tax_related: false,
      is_deferred_revenue_or_contract_liability: true,
      should_exclude_from_other_bucket: false,
      confidence: "high",
      reason: "Deferred income/revenue is a contract liability, not deferred income taxes."
    };
  }

  if (/\bconvertible\b.*\bnotes?\b|\bsenior notes?\b|\bcurrent maturit|\bcurrent portion\b.*\blong[-\s]?term debt\b|\blong[-\s]?term debt\b.*\bcurrent\b/.test(text)) {
    return {
      ...base,
      recommended_model_row: preferred("LT Debt (Incl. Current Portion)"),
      classification_type: "debt capital structure instrument",
      is_current: current,
      is_debt: true,
      is_operating: false,
      should_exclude_from_other_bucket: true,
      confidence: "high",
      reason:
        "Long-term debt instruments and their current maturities belong with LT Debt including current portion rather than Revolver or accrued liabilities."
    };
  }

  if (/\b(short[-\s]?term borrowings?|commercial paper|revolver|revolving credit|line of credit|current borrowings?)\b/.test(text)) {
    return {
      ...base,
      recommended_model_row: preferred("Revolver"),
      classification_type: "short-term borrowing facility",
      is_current: true,
      is_debt: true,
      is_operating: false,
      should_exclude_from_other_bucket: true,
      confidence: "high",
      reason: "The label describes a true short-term borrowing facility rather than current maturities of long-term debt."
    };
  }

  if (/\bshort[-\s]?term debt\b/.test(text) && !/\bconvertible|senior notes?|current maturit|current portion|long[-\s]?term/.test(text)) {
    return {
      ...base,
      recommended_model_row: preferred("Revolver"),
      classification_type: "short-term debt borrowing",
      is_current: true,
      is_debt: true,
      is_operating: false,
      should_exclude_from_other_bucket: true,
      confidence: "medium",
      reason: "Short-term debt generally maps to the short-term borrowing/Revolver row unless the filing identifies it as current long-term debt."
    };
  }

  if (/\baircraft fuel\b|\bspare parts?\b|\bparts and supplies\b|\bmerchandise inventory\b|\braw materials?\b|\bwork in process\b|\bfinished goods?\b|\binventor(?:y|ies)\b/.test(text)) {
    return {
      ...base,
      recommended_model_row: preferred("Inventory"),
      classification_type: "inventory-like operating current asset",
      is_current: true,
      is_operating: true,
      should_exclude_from_other_bucket: true,
      confidence: "high",
      reason: "Supplies, parts, fuel, merchandise, raw materials, WIP, and finished goods are inventory-like operating assets."
    };
  }

  if (/\bshort[-\s]?term investments?\b|\bmarketable securities\b|\bavailable[-\s]?for[-\s]?sale securities\b/.test(text) && section === "current assets") {
    return {
      ...base,
      recommended_model_row: hasCurrentInvestmentsRow
        ? request.availableModelRows.find((row) => /short[-\s]?term investments?|current investments?|marketable securities/i.test(row)) ?? "Short-Term Investments"
        : preferred("Cash & Cash Equivalents"),
      classification_type: "cash and marketable securities current investment",
      is_current: true,
      is_operating: false,
      should_exclude_from_other_bucket: true,
      confidence: "high",
      reason: hasCurrentInvestmentsRow
        ? "Current investments map to the dedicated current investments row when the template provides one."
        : "Marketable securities, available-for-sale securities, and short-term investments are included with Cash & Cash Equivalents when no dedicated current investments row exists."
    };
  }

  if (/\bassets?\b.*\bheld for sale\b|\bdisposal group\b.*\bassets?\b|\bassets?\b.*\bdiscontinued operation/.test(text) && section === "current assets") {
    return {
      ...base,
      recommended_model_row: preferred("Prepaid & Other Current Assets"),
      classification_type: "current assets held for sale",
      is_current: true,
      is_operating: null,
      should_exclude_from_other_bucket: false,
      confidence: "high",
      reason: "Current assets held for sale are current assets without a dedicated template row and should be grouped into Prepaid & Other Current Assets."
    };
  }

  if (/\bacquired\b.*\bin[-\s]?process\b.*\bresearch\b.*\bdevelopment\b|\bin[-\s]?process\b.*\bresearch\b.*\bdevelopment\b|\bipr&d\b|\biprd\b/.test(text)) {
    const row = modelRowAvailable("Other Operating Income / Expense", request.availableModelRows)
      ? "Other Operating Income / Expense"
      : preferred("R&D");
    return {
      ...base,
      recommended_model_row: row,
      classification_type: "acquired in-process R&D",
      is_current: null,
      is_operating: true,
      should_exclude_from_other_bucket: row !== "Other Operating Income / Expense",
      confidence: "high",
      reason: "Acquired in-process R&D is not depreciation or amortization; classify it as R&D or another operating item based on template convention."
    };
  }

  if (/\bdepreciation\b|\bamortization\b|\bdepletion\b|\bd&a\b/.test(text)) {
    const standaloneIncomeStatementDa = request.statement === "income_statement" && request.section === "operating expenses" && !/\bcash flows?|operating activities|reconciliation|supplemental/.test(text);
    return {
      ...base,
      recommended_model_row: standaloneIncomeStatementDa ? preferred("D&A") : "Unmapped / Cash-flow-only D&A",
      classification_type: standaloneIncomeStatementDa ? "standalone income-statement depreciation and amortization" : "cash-flow-only D&A disclosure",
      is_current: null,
      is_operating: standaloneIncomeStatementDa,
      should_exclude_from_other_bucket: true,
      confidence: standaloneIncomeStatementDa ? "high" : "medium",
      reason: standaloneIncomeStatementDa
        ? "The line is a standalone primary income-statement D&A expense."
        : "Cash-flow-only D&A should not be inserted into income-statement D&A."
    };
  }

  if (/\bspecial items?\b|\brestructuring\b|\bimpairment\b|\bspecial charges?\b|\bintegration costs?\b|\blitigation\b|\bsettlement\b|\baccretion\b/.test(text)) {
    const operating = request.section === "operating expenses";
    return {
      ...base,
      recommended_model_row: operating ? preferred("Other Operating Income / Expense") : preferred("Other Non-Operating Income / Expense"),
      classification_type: operating ? "special operating charge" : "below-operating special item",
      is_current: null,
      is_operating: operating,
      should_exclude_from_other_bucket: true,
      confidence: "high",
      reason: operating
        ? "The item is presented above operating income and belongs in other operating income/expense."
        : "The item is presented below operating income and belongs in other non-operating income/expense."
    };
  }

  if (/\binterest\b.*\bother\b|\bother\b.*\binterest\b/.test(text) && request.section === "below operating income") {
    return {
      ...base,
      recommended_model_row: preferred("Other Non-Operating Income / Expense"),
      classification_type: "combined below-operating line",
      is_current: null,
      is_operating: false,
      should_exclude_from_other_bucket: true,
      confidence: "high",
      reason: "Combined interest-and-other lines below EBIT remain non-operating unless reliable primary-statement detail splits interest."
    };
  }

  if (/\btreasury stock\b|\bcontra[-\s]?equity\b|\besop\b|\bemployee benefit trust\b/.test(text)) {
    return {
      ...base,
      recommended_model_row: preferred("Treasury Stock"),
      classification_type: "contra-equity",
      is_current: null,
      is_operating: false,
      should_exclude_from_other_bucket: true,
      confidence: "high",
      reason: "Treasury stock and employee trust shares are contra-equity items."
    };
  }

  if (/\bcommon stock\b|\badditional paid[-\s]?in capital\b|\bpaid[-\s]?in capital\b|\bcapital in excess\b/.test(text)) {
    return {
      ...base,
      recommended_model_row: preferred("Common Stock & APIC"),
      classification_type: "contributed equity",
      is_current: null,
      is_operating: false,
      should_exclude_from_other_bucket: true,
      confidence: "high",
      reason: "Common stock and additional paid-in capital belong in Common Stock & APIC."
    };
  }

  return null;
}

function conservativeFallbackClassification(request: FinancialLineItemClassificationRequest): FinancialLineItemClassification {
  const base = baseClassification(request);
  const candidate = request.deterministicCandidate && modelRowAvailable(request.deterministicCandidate, request.availableModelRows)
    ? request.deterministicCandidate
    : safestFallbackRowForSection(request);
  return {
    ...base,
    recommended_model_row: candidate,
    classification_type: "ambiguous financial statement line item",
    confidence: "low",
    should_exclude_from_other_bucket: /other/i.test(candidate) ? false : true,
    reason: request.uncertaintyReason || "No exact deterministic accounting classification was available; validation should keep this at review confidence."
  };
}

function safestFallbackRowForSection(request: FinancialLineItemClassificationRequest) {
  if (request.section === "current assets") return availableOrDefault(request, "Prepaid & Other Current Assets");
  if (request.section === "non-current assets") return availableOrDefault(request, "Other Non-Current Assets");
  if (request.section === "current liabilities") return availableOrDefault(request, "Other Current Liabilities");
  if (request.section === "non-current liabilities") return availableOrDefault(request, "Other Non-Current Liabilities");
  if (request.section === "operating expenses") return availableOrDefault(request, "Other Operating Income / Expense");
  if (request.section === "below operating income") return availableOrDefault(request, "Other Non-Operating Income / Expense");
  if (request.section === "tax") return availableOrDefault(request, "Income Tax Benefit / Expense");
  if (request.section === "revenue") return availableOrDefault(request, "Revenue");
  return request.deterministicCandidate || "Unmapped / Needs Review";
}

function availableOrDefault(request: FinancialLineItemClassificationRequest, row: string) {
  return request.availableModelRows.find((available) => modelRowsMatch(available, row)) ?? row;
}

function baseClassification(request: FinancialLineItemClassificationRequest): FinancialLineItemClassification {
  const text = requestSearchText(request);
  const isDeferredTax = /\bdeferred\b/.test(text) && /\btax(?:es)?\b/.test(text);
  const isDeferredRevenue = /\bdeferred (?:income|revenue)\b|\bunearned revenue\b|\bcontract liabilit|\bcustomer advances?\b/.test(text);
  return {
    source_line_item: request.cleanLabel || request.reportedLineItemLabel,
    recommended_action: "map",
    recommended_model_row: request.deterministicCandidate || "Unmapped / Needs Review",
    recommended_model_row_mappings: [],
    explicit_zero_rows: [],
    classification_type: "unclassified",
    is_current: request.section.includes("current") ? request.section.startsWith("current") : null,
    is_debt: /\bdebt\b|\bnotes?\b|\bborrowings?\b|\bcommercial paper\b|\brevolver\b|\bcredit facility\b/.test(text),
    is_operating: request.section === "operating expenses" ? true : request.section === "below operating income" ? false : null,
    is_tax_related: /\btax(?:es)?\b/.test(text),
    is_deferred_revenue_or_contract_liability: isDeferredRevenue,
    is_deferred_tax: isDeferredTax,
    is_subtotal: request.isSubtotal,
    should_exclude_from_other_bucket: false,
    confidence: "medium",
    reason: "",
    requires_validation: true,
    requires_revalidation: true,
    llm_used: false,
    mapping_passed_validation: false
  };
}

function finalizeClassification(
  request: FinancialLineItemClassificationRequest,
  classification: FinancialLineItemClassification
): FinancialLineItemClassification {
  const recommended = normalizeModelRow(classification.recommended_model_row) || classification.recommended_model_row;
  const mappingPassedValidation = classification.mapping_passed_validation || classificationPassesValidation(request, { ...classification, recommended_model_row: recommended });
  return {
    ...classification,
    recommended_action: classification.recommended_action || "map",
    recommended_model_row: recommended,
    recommended_model_row_mappings: Array.isArray(classification.recommended_model_row_mappings)
      ? classification.recommended_model_row_mappings
      : [],
    explicit_zero_rows: Array.isArray(classification.explicit_zero_rows) ? classification.explicit_zero_rows : [],
    source_line_item: classification.source_line_item || request.cleanLabel || request.reportedLineItemLabel,
    is_subtotal: request.isSubtotal || classification.is_subtotal,
    requires_validation: true,
    requires_revalidation: true,
    mapping_passed_validation: mappingPassedValidation,
    reason: shortReason(classification.reason || request.uncertaintyReason || "Accounting classification requires validation.")
  };
}

function classificationPassesValidation(request: FinancialLineItemClassificationRequest, classification: FinancialLineItemClassification) {
  const text = requestSearchText(request);
  const row = classification.recommended_model_row;
  if (request.isSubtotal || /unmapped|needs review/i.test(row)) return false;
  if (modelRowsMatch(row, "D&A") && request.statement !== "income_statement") return false;
  if (modelRowsMatch(row, "D&A") && /cash flows?|operating activities|reconciliation|supplemental/.test(text)) return false;
  if (modelRowsMatch(row, "Revolver") && /\bcurrent maturit|\bcurrent portion\b.*\blong[-\s]?term debt|convertible|senior notes?/.test(text)) return false;
  if (modelRowsMatch(row, "Deferred Income Taxes") && !classification.is_deferred_tax) return false;
  if ((modelRowsMatch(row, "Accrued Liabilities") || modelRowsMatch(row, "Other Current Liabilities")) && classification.is_debt) return false;
  if (modelRowsMatch(row, "Prepaid & Other Current Assets") && /inventor|spare parts?|aircraft fuel|supplies|short[-\s]?term investments?/.test(text)) return false;
  if (request.section === "current assets" && !modelRowsMatch(row, "Cash & Cash Equivalents") && !modelRowsMatch(row, "Accounts Receivable") && !modelRowsMatch(row, "Inventory") && !modelRowsMatch(row, "Prepaid & Other Current Assets")) return false;
  if (request.section === "current liabilities" && !modelRowsMatch(row, "Accounts Payable") && !modelRowsMatch(row, "Accrued Liabilities") && !modelRowsMatch(row, "Other Current Liabilities") && !modelRowsMatch(row, "Revolver") && !modelRowsMatch(row, "LT Debt (Incl. Current Portion)")) return false;
  return true;
}

async function requestLlmClassification(
  request: FinancialLineItemClassificationRequest,
  options: LlmClassificationOptions
): Promise<FinancialLineItemClassification> {
  const system = [
    "You are a structured accounting classifier for SEC EDGAR financial statement line items.",
    "Classify the reported source line item into the best model template row using accounting meaning, statement section, XBRL tag semantics, parent subtotal, and template row definitions.",
    "Do not use keyword matching alone. Reported statement location and accounting meaning take precedence over mathematical tie-outs.",
    "Use Other buckets only when no dedicated row exists. Do not use residual plugging.",
    "Current maturities/current portion of long-term debt and convertible senior notes belong with LT Debt including current portion, not Revolver.",
    "Deferred income/revenue is a contract liability, not deferred income taxes. Deferred tax liabilities are Deferred Income Taxes.",
    "Cash-flow-only D&A must not be inserted into income-statement D&A.",
    "When the correct repair is no reported line item, return recommended_action set_zero with explicit_zero_rows populated.",
    "Use recommended_action remap for validation failures caused by a source line belonging in a different row; use exclude for subtotal/component double-counting.",
    "Return strict JSON only."
  ].join(" ");
  const fetchImpl = options.fetchImpl ?? fetch;
  const controller = new AbortController();
  const timeoutMs = options.timeoutMs ?? 15_000;
  let timeout: ReturnType<typeof setTimeout> | undefined;
  const response = await Promise.race([
    fetchImpl(options.endpoint, {
      method: "POST",
      signal: controller.signal,
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${options.apiKey}`,
        "HTTP-Referer": options.siteUrl,
        "X-Title": options.appTitle
      },
      body: JSON.stringify({
        model: options.model,
        messages: [
          { role: "system", content: system },
          { role: "user", content: JSON.stringify(request) }
        ],
        temperature: 0,
        max_tokens: 700,
        provider: { require_parameters: true },
        response_format: {
          type: "json_schema",
          json_schema: financialLineItemClassificationJsonSchema()
        }
      })
    }),
    new Promise<Response>((_, reject) => {
      timeout = setTimeout(() => {
        controller.abort();
        reject(new Error(`classifier LLM timed out after ${timeoutMs}ms`));
      }, timeoutMs);
    })
  ]).finally(() => {
    if (timeout) clearTimeout(timeout);
  });
  const body = await response.json().catch(() => null);
  if (!response.ok) {
    const message = body?.error?.message || `${response.status} ${response.statusText}`;
    throw new Error(message);
  }
  const text = responseOutputText(body);
  if (!text) throw new Error("classifier response did not include text output");
  return JSON.parse(text) as FinancialLineItemClassification;
}

function financialLineItemClassificationJsonSchema() {
  return {
    name: "financial_line_item_classification",
    strict: true,
    schema: {
      type: "object",
      additionalProperties: false,
      required: [
        "source_line_item",
        "recommended_action",
        "recommended_model_row",
        "recommended_model_row_mappings",
        "explicit_zero_rows",
        "classification_type",
        "is_current",
        "is_debt",
        "is_operating",
        "is_tax_related",
        "is_deferred_revenue_or_contract_liability",
        "is_deferred_tax",
        "is_subtotal",
        "should_exclude_from_other_bucket",
        "confidence",
        "reason",
        "requires_validation",
        "requires_revalidation"
      ],
      properties: {
        source_line_item: { type: "string" },
        recommended_action: { type: "string", enum: ["map", "remap", "set_zero", "merge_into_other", "split_across_rows", "keep_existing", "exclude"] },
        recommended_model_row: { type: "string" },
        recommended_model_row_mappings: {
          type: "array",
          items: {
            type: "object",
            additionalProperties: false,
            required: ["source_line_item", "model_row", "amount", "reason"],
            properties: {
              source_line_item: { type: "string" },
              model_row: { type: "string" },
              amount: { anyOf: [{ type: "number" }, { type: "null" }] },
              reason: { type: "string" }
            }
          }
        },
        explicit_zero_rows: {
          type: "array",
          items: {
            type: "object",
            additionalProperties: false,
            required: ["model_row", "reason"],
            properties: {
              model_row: { type: "string" },
              reason: { type: "string" }
            }
          }
        },
        classification_type: { type: "string" },
        is_current: { anyOf: [{ type: "boolean" }, { type: "null" }] },
        is_debt: { type: "boolean" },
        is_operating: { anyOf: [{ type: "boolean" }, { type: "null" }] },
        is_tax_related: { type: "boolean" },
        is_deferred_revenue_or_contract_liability: { type: "boolean" },
        is_deferred_tax: { type: "boolean" },
        is_subtotal: { type: "boolean" },
        should_exclude_from_other_bucket: { type: "boolean" },
        confidence: { type: "string", enum: ["high", "medium", "low"] },
        reason: { type: "string" },
        requires_validation: { type: "boolean" },
        requires_revalidation: { type: "boolean" }
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
    for (const chunk of item.content ?? []) {
      if (chunk?.type === "output_text" && typeof chunk.text === "string") chunks.push(chunk.text);
    }
  }
  return chunks.join("").trim();
}

function equivalentModelRows(row: string) {
  const normalized = normalizeKey(row);
  const matches = Object.entries(MODEL_ROW_ALIASES)
    .filter(([canonical, aliases]) => normalizeKey(canonical) === normalized || aliases.some((alias) => normalizeKey(alias) === normalized))
    .map(([canonical]) => canonical);
  return matches.length ? matches : [row];
}

function normalizeModelRow(row: string) {
  return equivalentModelRows(row)[0] ?? row;
}

function rowIsCoreClassificationTarget(row: string) {
  return [
    "Cash & Cash Equivalents",
    "Inventory",
    "Revolver",
    "LT Debt (Incl. Current Portion)",
    "Deferred Income Taxes",
    "Other Current Liabilities",
    "Other Non-Current Liabilities",
    "R&D",
    "D&A",
    "Other Operating Income / Expense",
    "Other Non-Operating Income / Expense"
  ].includes(row);
}

function requestSearchText(request: FinancialLineItemClassificationRequest) {
  return [
    request.reportedLineItemLabel,
    request.cleanLabel,
    request.xbrlTag ?? "",
    request.parentSubtotal?.label ?? "",
    request.parentSubtotal?.concept ?? "",
    request.section
  ]
    .join(" ")
    .replace(/([a-z])([A-Z])/g, "$1 $2")
    .toLowerCase();
}

function normalizeKey(input: string) {
  return input.toLowerCase().replace(/[^a-z0-9]/g, "");
}

function lowerClassificationConfidence(a: FinancialLineItemClassification["confidence"], b: FinancialLineItemClassification["confidence"]) {
  const rank = { high: 3, medium: 2, low: 1 };
  return rank[a] <= rank[b] ? a : b;
}

function shortReason(reason: string) {
  const compact = reason.replace(/\s+/g, " ").trim();
  return compact.length <= 320 ? compact : `${compact.slice(0, 317).trim()}...`;
}
