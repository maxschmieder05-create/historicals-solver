import { normalizeAccession } from "./sec-accession";
import {
  balanceSheetRowDefinitionForLabel,
  balanceSheetRowsEquivalent,
  balanceSheetSectionCompatible,
  balanceSheetSourceLooksLikeDebtCarryingValueAdjustment
} from "./balance-sheet-row-resolver";
import { currentNonCurrentSignalFromText } from "./current-non-current";
import {
  AccountingLlmResult,
  AccountingLlmStatus,
  AccountingLlmTelemetry,
  AccountingLlmValidationResult,
  requestAccountingJson
} from "./llm-accounting-controller";

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
  sourceRowKey?: string;
  rowOrder?: number;
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
  llm_status?: AccountingLlmStatus;
  mapping_passed_validation: boolean;
  warning?: string;
};

export type FinancialLineItemClassificationStore = Map<string, FinancialLineItemClassification>;

type LlmClassificationOptions = {
  enabled: boolean;
  apiKey: string;
  endpoint: string;
  model: string;
  fallbackModels?: string[];
  siteUrl: string;
  appTitle: string;
  timeoutMs?: number;
  fetchImpl?: typeof fetch;
};

type ClassifierOptions = {
  llm?: LlmClassificationOptions;
  statementAnalystPass?: {
    enabled: boolean;
    materialityThreshold?: number;
    coverage?: "material_and_ambiguous" | "all_primary_rows";
  };
};

type PreparedLineItemClassification = {
  request: FinancialLineItemClassificationRequest;
  rowKey: string;
  deterministic: FinancialLineItemClassification | null;
  fallback: FinancialLineItemClassification;
  initialClassification: FinancialLineItemClassification;
  deterministicIsValidated: boolean;
  materialAnalystTarget: boolean;
  fullStatementTarget: boolean;
  needsClassification: boolean;
  needsLlm: boolean;
};

type StatementLlmClassificationDecision = {
  source_row_key: string;
  recommended_action: FinancialLineItemClassification["recommended_action"];
  recommended_model_row: string;
  confidence: FinancialLineItemClassification["confidence"];
  reason: string;
};

type StatementLlmClassificationResponse = {
  classifications: StatementLlmClassificationDecision[];
};

export type FinancialStatementLineItemClassificationResult = {
  classifications: Array<{
    request: FinancialLineItemClassificationRequest;
    classification: FinancialLineItemClassification;
  }>;
  warnings: string[];
  llmCalls: number;
  llmAttempts: number;
  llmSuccessfulCompletions: number;
  llmTelemetry: AccountingLlmTelemetry[];
  targetCount: number;
  llmReviewedCount: number;
  acceptedDecisionCount: number;
  unreviewedTargetKeys: string[];
};

export const MODEL_ROW_DEFINITIONS: Record<string, string> = {
  "Cash & Cash Equivalents":
    "Cash and cash equivalents. Include current marketable securities or short-term investments when no dedicated current-investments row exists, because they are cash-like current investment balances.",
  "Accounts Receivable": "Trade accounts receivable, accounts receivable, receivables net of allowances.",
  Inventory:
    "Inventories and inventory-like operating assets, including supplies, parts, merchandise inventory, raw materials, WIP, finished goods, aircraft fuel/spare parts/supplies.",
  "Prepaid & Other Current Assets":
    "Current assets not better mapped elsewhere, such as prepaid expenses, other current assets, tax receivables, and contract assets. Current marketable securities and short-term investments belong here only if the template has no cash row and no dedicated current-investments row.",
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
    "Long-term debt instruments including current portion/current maturities, senior notes, convertible senior notes, long-term debt net of current portion plus current portion, long-term debt and finance lease obligations. Debt discounts and issuance costs are carrying-value adjustments to this row unless a reported net debt balance already includes them.",
  "Deferred Income Taxes":
    "True deferred tax liabilities / deferred income taxes / deferred tax liabilities, non-current / deferred tax assets and liabilities net.",
  "Other Non-Current Liabilities":
    "Non-current liabilities without better dedicated rows, including long-term deferred revenue, long-term contract liabilities, operating lease liabilities non-current, pension and postretirement obligations, asset retirement obligations, other long-term liabilities. If the template has no mezzanine-equity row, include redeemable noncontrolling interests here so total liabilities and equity reconcile.",
  "Mezzanine Equity":
    "Redeemable noncontrolling interests and other mezzanine equity presented outside permanent stockholders' equity.",
  "Common Stock & APIC": "Common stock, additional paid-in capital, capital in excess of par.",
  "Retained Earnings": "Retained earnings or retained deficit.",
  "Treasury Stock": "Treasury stock or contra-equity items such as employee benefit trust / ESOP trust if no better row exists.",
  AOCI: "Accumulated other comprehensive income/loss.",
  "Noncontrolling Interests": "Noncontrolling interests / minority interest.",
  Revenue: "Revenue, net revenue, net sales, operating revenue, total revenue.",
  "COGS / Cost of Goods Sold":
    "Direct cost of revenue / cost of sales / cost of operations, including industry-specific direct operating costs if presented as direct operating expenses and not better mapped elsewhere.",
  "SG&A": "Selling, general and administrative, advertising, marketing, promotional, sales expense, G&A, corporate overhead.",
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
  "securities",
  "supplies",
  "contract",
  "financing",
  "lease",
  "tax",
  "income",
  "advertising",
  "marketing",
  "promotional",
  "promotion",
  "selling",
  "administrative",
  "redeemable",
  "mezzanine",
  "noncontrolling"
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
  AOCI: ["accumulated other comprehensive income", "accumulated other comprehensive income (aoci)", "accumulated other comprehensive loss"],
  "Mezzanine Equity": ["mezzanine equity", "redeemable noncontrolling interests", "redeemable nci"],
  "SG&A": [
    "sga",
    "sg&a",
    "selling general administrative",
    "selling general administration sga",
    "selling general and administration sga",
    "selling general administrative sga",
    "selling general and administrative sga",
    "sales and marketing",
    "selling and marketing",
    "marketing expense",
    "advertising expense",
    "promotional expense",
    "promotion expense",
    "sales expense",
    "selling expense",
    "general and administrative",
    "administrative expense"
  ],
  "R&D": ["research and development", "r&d", "research & development", "research development rd", "research and development rd"],
  "COGS / Cost of Goods Sold": ["cost of goods sold", "cost of goods & services sold", "cost of revenue", "cost of sales", "cogs"],
  "D&A": ["depreciation and amortization", "depreciation & amortization", "d&a"],
  "Other Operating Income / Expense": ["other operating income expense", "other operating income", "other operating expense"],
  "Other Non-Operating Income / Expense": ["other non-operating income expense", "other income expense", "other expense income"]
};

const GENERAL_ACCOUNTING_ROUTING_INSTRUCTIONS = [
  "This is a general side-by-side mapping task, not a keyword lookup and not a rule limited to the examples.",
  "For every target row, compare the SEC source row against all available model rows and definitions, then choose the row whose accounting substance best fits.",
  "Use statement placement, current/non-current section, parent subtotal, nearby rows, XBRL concept semantics, and prior-period labels to infer meaning when filing labels and model labels differ.",
  "Prefer a specific model row when the template exposes one; otherwise group into the appropriate Other bucket with a reusable accounting reason.",
  "Exclude subtotals, totals, component detail, and duplicate rows when mapping them would double-count a model row.",
  "Preserve EDGAR tie-outs: major model totals should reconcile to the SEC filing through assigned components, formulas, or explicit exclusion reasons."
];

const DEFAULT_MATERIAL_STATEMENT_ROW_THRESHOLD = 500_000;

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
  if (/short[-\s]?term|current investments?|marketable securities|available[-\s]?for[-\s]?sale securities|current maturit|current portion|senior notes?|convertible|contract liabilit|deferred revenue|deferred income|spare parts?|supplies|in[-\s]?process research|special items?|other/.test(text)) return true;
  if (
    request.statement === "income_statement" &&
    /\badvertising\b|\bmarketing\b|\bpromotion(?:al)?\b|\bsales and marketing\b|\bselling and marketing\b|\bsales expense\b|\bselling expense\b|\bgeneral and administrative\b|\badministrative expense\b|\bcorporate overhead\b/.test(text)
  ) return true;
  if (request.deterministicCandidate && /other|accrued|revolver|deferred|d&a|depreciation|amortization/i.test(request.deterministicCandidate)) return true;
  return Boolean(request.uncertaintyReason);
}

export function materialStatementLineItemNeedsAnalystPass(
  request: FinancialLineItemClassificationRequest,
  materialityThreshold = DEFAULT_MATERIAL_STATEMENT_ROW_THRESHOLD
) {
  if (request.sourceTableType !== "primary_statement") return false;
  if (request.statement !== "income_statement" && request.statement !== "balance_sheet") return false;
  if (request.unit && !/usd/i.test(request.unit)) return false;
  if (typeof request.amount !== "number" || !Number.isFinite(request.amount)) return false;
  const threshold = Number.isFinite(materialityThreshold) && materialityThreshold >= 0
    ? materialityThreshold
    : DEFAULT_MATERIAL_STATEMENT_ROW_THRESHOLD;
  return Math.abs(request.amount) >= threshold;
}

function primaryBalanceSheetLineItemNeedsLlmReview(request: FinancialLineItemClassificationRequest) {
  return request.statement === "balance_sheet" && request.sourceTableType === "primary_statement" && lineItemNeedsClassification(request);
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
      (!deterministicIsValidated || primaryBalanceSheetLineItemNeedsLlmReview(request))
  );

  if (!shouldCallLlm) return finalizeClassification(request, fallback);

  const llm = options.llm!;
  const result = await requestLlmClassification(request, llm);
  if (result.value) {
    const llmClassification = result.value;
    return finalizeClassification(request, {
      ...llmClassification,
      source_line_item: llmClassification.source_line_item || request.cleanLabel || request.reportedLineItemLabel,
      recommended_model_row:
        normalizeModelRow(llmClassification.recommended_model_row) ||
        normalizeModelRow(fallback.recommended_model_row) ||
        fallback.recommended_model_row,
      confidence: llmClassification.confidence ?? "low",
      requires_validation: true,
      llm_used: result.telemetry.affectedOutput,
      llm_status: result.status
    });
  }
  return failedLlmClassification(request, fallback, result.status, result.error || result.telemetry.errorMessage || "unknown classifier LLM error");
}

export async function classifyFinancialStatementLineItems(
  requests: FinancialLineItemClassificationRequest[],
  options: ClassifierOptions = {}
): Promise<FinancialStatementLineItemClassificationResult> {
  const prepared = requests.map((request, index) => prepareLineItemClassification(request, index, options));
  const classificationTargets = prepared.filter((item) => item.needsClassification);
  const classifications = classificationTargets.map((item) => ({ request: item.request, classification: item.initialClassification }));
  const targets = prepared.filter((item) => item.needsLlm);

  if (!targets.length) {
    return {
      classifications,
      warnings: [],
      llmCalls: 0,
      llmAttempts: 0,
      llmSuccessfulCompletions: 0,
      llmTelemetry: [],
      targetCount: classificationTargets.length,
      llmReviewedCount: 0,
      acceptedDecisionCount: 0,
      unreviewedTargetKeys: classificationTargets.map((item) => item.rowKey)
    };
  }

  const llm = options.llm!;
  const result = await requestStatementLlmClassification(prepared, targets, llm);
  const telemetry = result.attemptTelemetry ?? [result.telemetry];
  if (result.value) {
    const response = result.value;
    const byRowKey = new Map(response.classifications.map((item) => [item.source_row_key, item]));
    const merged = classificationTargets.map((item) => {
      const llmDecision = item.needsLlm ? byRowKey.get(item.rowKey) : null;
      if (!llmDecision) return { request: item.request, classification: item.initialClassification };
      return {
        request: item.request,
        classification: finalizeClassification(item.request, {
          ...statementDecisionToClassification(item, llmDecision),
          llm_used: result.telemetry.affectedOutput,
          llm_status: result.status
        })
      };
    });
    const missingTargets = targets.filter((item) => !byRowKey.has(item.rowKey));
    const warnings = missingTargets.map(
      (item) =>
        `${item.request.cleanLabel || item.request.reportedLineItemLabel}: statement-level LLM did not return a classification; deterministic fallback was used.`
    );
    const reviewed = merged.filter(({ classification }) => classification.llm_used);
    const accepted = reviewed.filter(({ classification }) => statementClassificationAccepted(classification));
    const acceptedRequests = new Set(accepted.map(({ request }) => request));
    const unreviewedTargetKeys = classificationTargets
      .filter((item) => !acceptedRequests.has(item.request))
      .map((item) => item.rowKey);
    return {
      classifications: merged,
      warnings,
      llmCalls: telemetry.filter((item) => item.completed).length,
      llmAttempts: telemetry.filter((item) => item.attempted).length,
      llmSuccessfulCompletions: telemetry.filter((item) => item.completed).length,
      llmTelemetry: telemetry,
      targetCount: classificationTargets.length,
      llmReviewedCount: reviewed.length,
      acceptedDecisionCount: accepted.length,
      unreviewedTargetKeys
    };
  }
  if (statementClassificationFailureShouldSplit(result, requests.length)) {
    const midpoint = Math.ceil(requests.length / 2);
    const left = await classifyFinancialStatementLineItems(requests.slice(0, midpoint), options);
    const right = await classifyFinancialStatementLineItems(requests.slice(midpoint), options);
    return {
      classifications: [...left.classifications, ...right.classifications],
      warnings: [...left.warnings, ...right.warnings],
      llmCalls: left.llmCalls + right.llmCalls,
      llmAttempts: telemetry.filter((item) => item.attempted).length + left.llmAttempts + right.llmAttempts,
      llmSuccessfulCompletions: left.llmSuccessfulCompletions + right.llmSuccessfulCompletions,
      llmTelemetry: [...telemetry, ...left.llmTelemetry, ...right.llmTelemetry],
      targetCount: left.targetCount + right.targetCount,
      llmReviewedCount: left.llmReviewedCount + right.llmReviewedCount,
      acceptedDecisionCount: left.acceptedDecisionCount + right.acceptedDecisionCount,
      unreviewedTargetKeys: [...left.unreviewedTargetKeys, ...right.unreviewedTargetKeys]
    };
  }
  const message = result.error || result.telemetry.errorMessage || "unknown classifier LLM error";
  const failed = classificationTargets.map((item) => ({
    request: item.request,
    classification: item.needsLlm
      ? failedLlmClassification(item.request, item.fallback, result.status, message)
      : item.initialClassification
  }));
  const warnings = targets.map(
    (item) => `${item.request.cleanLabel || item.request.reportedLineItemLabel}: statement-level LLM classification ${result.status} (${message}).`
  );
  return {
    classifications: failed,
    warnings,
    llmCalls: 0,
    llmAttempts: telemetry.filter((item) => item.attempted).length,
    llmSuccessfulCompletions: 0,
    llmTelemetry: telemetry,
    targetCount: classificationTargets.length,
    llmReviewedCount: 0,
    acceptedDecisionCount: 0,
    unreviewedTargetKeys: classificationTargets.map((item) => item.rowKey)
  };
}

function statementClassificationFailureShouldSplit(
  result: AccountingLlmResult<StatementLlmClassificationResponse>,
  requestCount: number
) {
  const error = result.error || result.telemetry.errorMessage || "";
  if (
    requestCount > 1 &&
    /recommended_model_row must be a string|recommended_action is invalid|must include source_row_key|unexpected source_row_key|duplicate source_row_key|every statement classification must be a json object|failed deterministic accounting validation|low confidence without a validated fallback/i.test(
      error
    )
  ) {
    return true;
  }
  if (requestCount <= 4) return false;
  if (/timed out|did not include text output|provider returned error|capacity|temporarily unavailable|overloaded/i.test(error)) return true;
  if (!result.rawText) return false;
  return /json|parse|expected ['",}\]]|incomplete|unterminated|classifications array|omitted target/i.test(error);
}

function statementClassificationAccepted(classification: FinancialLineItemClassification) {
  if (!classification.llm_used || classification.confidence === "low" || classification.warning) return false;
  if (
    classification.recommended_action === "exclude" ||
    classification.recommended_action === "set_zero" ||
    classification.recommended_action === "keep_existing"
  ) {
    return true;
  }
  return classification.mapping_passed_validation;
}

function prepareLineItemClassification(
  request: FinancialLineItemClassificationRequest,
  index: number,
  options: ClassifierOptions
): PreparedLineItemClassification {
  const deterministic = deterministicFinancialLineItemClassification(request);
  const fallback = deterministic ?? conservativeFallbackClassification(request);
  const deterministicIsValidated =
    deterministic?.confidence === "high" &&
    classificationPassesValidation(request, {
      ...deterministic,
      recommended_model_row: normalizeModelRow(deterministic.recommended_model_row) || deterministic.recommended_model_row
    });
  const completeMappingCoverage = options.statementAnalystPass?.coverage === "all_primary_rows";
  const fullStatementTarget = Boolean(
    options.statementAnalystPass?.enabled &&
      completeMappingCoverage &&
      fullStatementLineItemNeedsAnalystPass(request)
  );
  const materialAnalystTarget = Boolean(
    options.statementAnalystPass?.enabled &&
      !completeMappingCoverage &&
      materialStatementLineItemNeedsAnalystPass(request, options.statementAnalystPass.materialityThreshold)
  );
  const needsClassification = completeMappingCoverage
    ? fullStatementTarget
    : materialAnalystTarget || lineItemNeedsClassification(request);
  const needsLlm = Boolean(
    options.llm?.enabled &&
      options.llm.apiKey &&
      options.llm.model &&
      needsClassification &&
      (fullStatementTarget || materialAnalystTarget || !deterministicIsValidated || primaryBalanceSheetLineItemNeedsLlmReview(request))
  );

  return {
    request,
    rowKey: sourceRowKeyForRequest(request, index),
    deterministic,
    fallback,
    initialClassification: finalizeClassification(request, fallback),
    deterministicIsValidated,
    materialAnalystTarget,
    fullStatementTarget,
    needsClassification,
    needsLlm
  };
}

export function fullStatementLineItemNeedsAnalystPass(request: FinancialLineItemClassificationRequest) {
  if (
    request.sourceTableType !== "primary_statement" ||
    (request.statement !== "income_statement" && request.statement !== "balance_sheet") ||
    request.isSubtotal ||
    typeof request.amount !== "number" ||
    !Number.isFinite(request.amount) ||
    (Boolean(request.unit) && !/usd/i.test(request.unit ?? ""))
  ) {
    return false;
  }
  const text = `${request.cleanLabel || request.reportedLineItemLabel} ${request.xbrlTag ?? ""}`.toLowerCase();
  const tagCompact = (request.xbrlTag ?? "").toLowerCase().replace(/[^a-z0-9]/g, "");
  if (
    /\bper (?:common )?share\b|\bearnings per share\b|\beps\b|\bweighted average (?:number of )?shares?\b|\bshares? (?:outstanding|used|weighted)\b|\bnumber of shares?\b/.test(
      text
    )
  ) {
    return false;
  }
  if (
    request.statement === "income_statement" &&
    (/\bgross profit\b|\bgross margin\b|\boperating income\b|\boperating loss\b|\bincome (?:loss )?from operations\b|\bincome before (?:income )?tax|\bpretax (?:income|loss)\b|\bnet income\b|\bnet loss\b|\bprofit loss\b/.test(
      text
    ) ||
      /^(?:grossprofit|operatingincomeloss|incomelossfromcontinuingoperationsbeforeincometaxes|incomelossfromcontinuingoperations|profitloss|netincomeloss)$/.test(
        tagCompact
      ))
  ) {
    return false;
  }
  if (
    request.statement === "balance_sheet" &&
    /^(?:assetscurrent|assets|liabilitiescurrent|liabilities|stockholdersequity|stockholdersequityincludingportionattributabletononcontrollinginterest|liabilitiesandstockholdersequity)$/.test(
      tagCompact
    )
  ) {
    return false;
  }
  return true;
}

export function modelRowsMatch(a: string, b: string) {
  const left = normalizeKey(a);
  const right = normalizeKey(b);
  if (!left || !right) return false;
  if (left === right) return true;
  if (balanceSheetRowsEquivalent(a, b)) return true;
  return equivalentModelRows(a).some((candidate) => normalizeKey(candidate) === right) || equivalentModelRows(b).some((candidate) => normalizeKey(candidate) === left);
}

export function modelRowAvailable(row: string, availableRows: string[]) {
  return availableRows.some((available) => modelRowsMatch(row, available));
}

export function classificationModelRowAssignmentForPrimaryStatement(
  classification: FinancialLineItemClassification | null | undefined,
  availableModelRows: string[]
) {
  if (!classification) return null;
  if (classification.confidence === "low") return null;
  if (classification.recommended_action === "exclude" || classification.recommended_action === "set_zero") return null;
  if (!classification.mapping_passed_validation) return null;

  const normalizedRow = normalizeModelRow(classification.recommended_model_row) || classification.recommended_model_row;
  if (!normalizedRow || /unmapped|needs review/i.test(normalizedRow)) return null;
  const modelRow = availableModelRows.find((available) => modelRowsMatch(available, normalizedRow));
  if (!modelRow) return null;
  const grouped =
    classification.recommended_action === "merge_into_other" ||
    classification.recommended_action === "split_across_rows" ||
    isReusableOtherBucketModelRow(modelRow);
  return {
    modelRow,
    grouped,
    llmUsed: classification.llm_used,
    reason: `${classification.llm_used ? "LLM" : "Validated"} line-item classification: ${shortReason(classification.reason)}`
  };
}

function isReusableOtherBucketModelRow(modelRow: string) {
  const balanceSheetDefinition = balanceSheetRowDefinitionForLabel(modelRow);
  if (balanceSheetDefinition?.kind === "catch_all") return true;
  return (
    modelRowsMatch(modelRow, "Prepaid & Other Current Assets") ||
    modelRowsMatch(modelRow, "Other Non-Current Assets") ||
    modelRowsMatch(modelRow, "Other Current Liabilities") ||
    modelRowsMatch(modelRow, "Other Non-Current Liabilities") ||
    modelRowsMatch(modelRow, "Common Stock & APIC")
  );
}

function deterministicFinancialLineItemClassification(
  request: FinancialLineItemClassificationRequest
): FinancialLineItemClassification | null {
  const text = requestSearchText(request);
  const ownText = requestOwnSearchText(request);
  const section = request.section;
  const current = section.includes("current") ? section.startsWith("current") : null;
  const base = baseClassification(request);
  const preferred = (...rows: string[]) => rows.find((row) => modelRowAvailable(row, request.availableModelRows)) ?? rows[0];
  const rowGroupsCashAndCurrentInvestments = (row: string) =>
    /cash.*(short[-\s]?term investments?|current investments?|marketable securities)|(short[-\s]?term investments?|current investments?|marketable securities).*cash/i.test(row);
  const cashAndCurrentInvestmentsRow = request.availableModelRows.find((row) =>
    rowGroupsCashAndCurrentInvestments(row)
  );
  const cashRow = request.availableModelRows.find((row) => modelRowsMatch(row, "Cash & Cash Equivalents"));
  const currentInvestmentsRow = request.availableModelRows.find((row) =>
    !rowGroupsCashAndCurrentInvestments(row) && /short[-\s]?term investments?|current investments?|marketable securities|investment securities/i.test(row)
  );
  const currentNonCurrentSignal = currentNonCurrentSignalFromText(text);
  const explicitNonCurrent = currentNonCurrentSignal === "non-current";
  const effectiveCurrent = current === null && explicitNonCurrent ? false : current;
  const isDeferredTaxLine = /\bdeferred\b/.test(text) && /\btax(?:es)?\b/.test(text);
  const textLooksDeferredTaxAsset = /\bassets?\b/.test(text);
  const textLooksDeferredTaxLiability = /\bliabilit/.test(text);
  const redeemableNoncontrollingInterest =
    /\bredeemable\b.*\bnon[-\s]?controlling interests?\b|\bnon[-\s]?controlling interests?\b.*\bredeemable\b|\bredeemable nci\b/.test(text) ||
    /RedeemableNoncontrollingInterest/i.test(request.xbrlTag ?? "");
  const selfInsuranceReserve = requestLooksLikeSelfInsuranceReserve(request, text);

  if (request.statement === "balance_sheet" && redeemableNoncontrollingInterest) {
    const row = preferred("Mezzanine Equity", "Other Non-Current Liabilities");
    return {
      ...base,
      recommended_model_row: row,
      classification_type: "redeemable noncontrolling interest mezzanine equity",
      is_current: false,
      is_debt: false,
      is_operating: false,
      should_exclude_from_other_bucket: !modelRowsMatch(row, "Other Non-Current Liabilities"),
      confidence: "high",
      reason: modelRowsMatch(row, "Mezzanine Equity")
        ? "Redeemable noncontrolling interests are mezzanine equity and map to the dedicated mezzanine row when the template provides one."
        : "Redeemable noncontrolling interests are mezzanine equity outside permanent equity; because the template has no mezzanine row, group them into Other Non-Current Liabilities so the liabilities-and-equity side reconciles."
    };
  }

  if (request.statement === "balance_sheet" && selfInsuranceReserve) {
    const nonCurrent = explicitNonCurrent || section === "non-current liabilities" || /Noncurrent/i.test(request.xbrlTag ?? "");
    const row = nonCurrent
      ? preferred("Other Non-Current Liabilities")
      : preferred("Accrued Liabilities", "Other Current Liabilities");
    return {
      ...base,
      recommended_model_row: row,
      classification_type: nonCurrent ? "non-current self-insurance operating reserve" : "current self-insurance operating reserve",
      is_current: !nonCurrent,
      is_debt: false,
      is_operating: true,
      should_exclude_from_other_bucket: !modelRowsMatch(row, "Other Non-Current Liabilities") && !modelRowsMatch(row, "Other Current Liabilities"),
      confidence: "high",
      reason: nonCurrent
        ? "Non-current self-insurance reserves are operating liabilities and belong in Other Non-Current Liabilities, not debt."
        : "The current portion of self-insurance reserves is an accrued operating liability, not current maturities of debt."
    };
  }

  if (
    request.statement === "balance_sheet" &&
    balanceSheetSourceLooksLikeDebtCarryingValueAdjustment({
      label: request.cleanLabel || request.reportedLineItemLabel,
      tag: request.xbrlTag
    })
  ) {
    return {
      ...base,
      recommended_model_row: preferred("LT Debt (Incl. Current Portion)"),
      classification_type: "debt carrying-value adjustment",
      is_current: false,
      is_debt: true,
      is_operating: false,
      should_exclude_from_other_bucket: true,
      confidence: "high",
      reason:
        "Unamortized debt discounts and debt issuance costs adjust the carrying value of long-term debt; they belong with LT Debt unless the statement already reports a net debt carrying amount that includes the adjustment."
    };
  }

  if (
    request.statement === "balance_sheet" &&
    explicitNonCurrent &&
    (/\brestricted cash\b/.test(text) || /RestrictedCash.*Noncurrent/i.test(request.xbrlTag ?? ""))
  ) {
    return {
      ...base,
      recommended_model_row: preferred("Other Non-Current Assets"),
      classification_type: "non-current restricted cash or investment",
      is_current: false,
      is_debt: false,
      is_operating: false,
      should_exclude_from_other_bucket: false,
      confidence: "high",
      reason: "Non-current restricted cash and investments are long-term assets and must not be included in Cash & Cash Equivalents."
    };
  }

  if (isDeferredTaxLine && (/assets?/.test(section) || (textLooksDeferredTaxAsset && !/liabilit/.test(section)))) {
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

  if (isDeferredTaxLine && (textLooksDeferredTaxLiability || /liabilit/.test(section) || !textLooksDeferredTaxAsset)) {
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

  if (/\bdeferred (?:income|revenue)\b|\bunearned revenue\b|\bcontract liabilit|\bcustomer advances?\b/.test(text)) {
    const row = section === "non-current liabilities" || explicitNonCurrent ? preferred("Other Non-Current Liabilities") : preferred("Other Current Liabilities");
    return {
      ...base,
      recommended_model_row: row,
      classification_type: "deferred revenue or contract liability",
      is_current: modelRowsMatch(row, "Other Non-Current Liabilities") ? false : effectiveCurrent,
      is_tax_related: false,
      is_deferred_revenue_or_contract_liability: true,
      should_exclude_from_other_bucket: false,
      confidence: "high",
      reason: "Deferred income/revenue is a contract liability, not deferred income taxes."
    };
  }

  if (
    request.statement === "balance_sheet" &&
    effectiveCurrent !== false &&
    !explicitNonCurrent &&
    /\baccrued\b|\bcompensation\b|\bpayroll\b|\bsalar(?:y|ies)\b|\bwages payable\b|\bbenefits payable\b|\binterest payable\b|\brebates?\b|\breturns?\b|\bpromotions?\b|\bdiscounts payable\b/.test(text) &&
    !/\baccounts? payable\b|\btrade payables?\b|\bdeferred (?:income|revenue)\b|\bunearned revenue\b|\bcontract liabilit|\bcustomer advances?\b|\bdebt\b|\bborrowings?\b|\bnotes?\b/.test(text)
  ) {
    return {
      ...base,
      recommended_model_row: preferred("Accrued Liabilities"),
      classification_type: "accrued operating current liability",
      is_current: true,
      is_operating: true,
      should_exclude_from_other_bucket: true,
      confidence: "high",
      reason: "Accrued compensation, payroll, interest payable, taxes, rebates, returns, promotions, and similar current accruals belong in accrued liabilities, not the other-current-liabilities bucket."
    };
  }

  if (!explicitNonCurrent && /\bshort[-\s]?term\b.*\bincome taxes?\b|\bincome taxes?\b.*\bcurrent\b|\baccrued income taxes current\b|\bincome taxes payable\b|\btaxes payable\b/.test(text)) {
    return {
      ...base,
      recommended_model_row: preferred("Accrued Liabilities"),
      classification_type: "current income tax payable",
      is_current: true,
      is_tax_related: true,
      should_exclude_from_other_bucket: true,
      confidence: "high",
      reason: "Short-term income taxes payable are current accrued tax liabilities."
    };
  }

  if (explicitNonCurrent && /\bincome taxes?\b|\btaxes payable\b/.test(text) && !isDeferredTaxLine) {
    return {
      ...base,
      recommended_model_row: preferred("Other Non-Current Liabilities"),
      classification_type: "non-current income tax payable",
      is_current: false,
      is_tax_related: true,
      should_exclude_from_other_bucket: false,
      confidence: "high",
      reason: "Long-term income taxes payable are non-current tax liabilities, not current accrued liabilities."
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

  if (/\baircraft fuel\b|\bspare parts?\b|\bparts and supplies\b|\bmerchandise inventory\b|\braw materials?\b|\bwork[-\s]?in[-\s]?process\b|\bfinished goods?\b|\binventor(?:y|ies)\b/.test(text)) {
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
    const recommendedRow =
      currentInvestmentsRow ??
      cashAndCurrentInvestmentsRow ??
      cashRow ??
      preferred("Prepaid & Other Current Assets");
    const mapsToDedicatedRow = Boolean(currentInvestmentsRow || cashAndCurrentInvestmentsRow || cashRow);
    return {
      ...base,
      recommended_model_row: recommendedRow,
      classification_type: currentInvestmentsRow ? "current investment" : "cash-like current investment",
      is_current: true,
      is_operating: false,
      should_exclude_from_other_bucket: mapsToDedicatedRow,
      confidence: "high",
      reason: currentInvestmentsRow
        ? "Current investments map to the dedicated current investments row when the template provides one."
        : cashAndCurrentInvestmentsRow
          ? "Current investments map to the template's explicit cash-and-current-investments row."
          : cashRow
            ? "Current marketable securities and short-term investments group with cash when the template has no dedicated current-investments row."
            : "Current investments group into the current-assets residual row when the template has no cash row and no dedicated current-investments row."
    };
  }

  if (/\binvestment securities\b|\bdebt and equity securities\b|\bavailable[-\s]?for[-\s]?sale securities\b|\bmarketable securities\b/.test(text) && section === "non-current assets") {
    return {
      ...base,
      recommended_model_row: preferred("Other Non-Current Assets"),
      classification_type: "investment securities asset",
      is_current: false,
      is_operating: false,
      should_exclude_from_other_bucket: false,
      confidence: "high",
      reason: "Investment securities are asset investments, not shareholder equity, even when the SEC concept contains the word equity."
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

  if (/\bland\b.*\bbuilding|\bbuildings?\b.*\bimprovements?\b|\bmachinery\b.*\bequipment\b|\bfurniture\b.*\bfixtures?\b|\bconstruction in progress\b|\bleasehold improvements?\b|\bproperty\b.*\bplant\b.*\bequipment\b|\bproperty and equipment\b|\bpp&e\b/.test(text)) {
    return {
      ...base,
      recommended_model_row: preferred("PP&E, Net"),
      classification_type: "property plant and equipment component",
      is_current: false,
      is_operating: true,
      should_exclude_from_other_bucket: true,
      confidence: "high",
      reason: "Land, buildings, improvements, machinery, equipment, and construction-in-progress are PP&E component rows."
    };
  }

  const excludesAcquiredInProcessCost =
    /\bexclud(?:e|es|ing)\b.*\bacquired\b.*\bin[-\s]?process\b.*\bcost\b/.test(ownText) ||
    /researchanddevelopmentexpenseexcludingacquiredinprocesscost/i.test(request.xbrlTag ?? "");
  if (!excludesAcquiredInProcessCost && /\bacquired\b.*\bin[-\s]?process\b.*\bresearch\b.*\bdevelopment\b|\bin[-\s]?process\b.*\bresearch\b.*\bdevelopment\b|\bipr&d\b|\biprd\b/.test(ownText)) {
    const row = modelRowAvailable("Other Operating Income / Expense", request.availableModelRows)
      ? "Other Operating Income / Expense"
      : preferred("R&D");
    return {
      ...base,
      recommended_model_row: row,
      classification_type: "in-process R&D charge",
      is_current: null,
      is_operating: true,
      should_exclude_from_other_bucket: row !== "Other Operating Income / Expense",
      confidence: "high",
      reason: "In-process R&D charges and impairments are not depreciation or amortization; classify them as R&D or another operating item based on template convention."
    };
  }

  if (
    request.statement === "income_statement" &&
    request.section === "operating expenses" &&
    /\bresearch\b|\br&d\b|\bproduct development\b|\bengineering expense\b|\btechnology development\b|\btechnology and content\b/.test(ownText)
  ) {
    return {
      ...base,
      recommended_model_row: preferred("R&D"),
      classification_type: "research and development operating expense",
      is_current: null,
      is_operating: true,
      should_exclude_from_other_bucket: true,
      confidence: "high",
      reason: "Research, product development, engineering, and technology development expenses belong in R&D when reported as operating expenses."
    };
  }

  if (
    request.statement === "income_statement" &&
    request.section === "operating expenses" &&
    /\bspecial items?\b|\brestructuring\b|\bimpairment\b|\bspecial charges?\b|\bintegration costs?\b|\blitigation\b|\bsettlement\b|\baccretion\b/.test(ownText)
  ) {
    return {
      ...base,
      recommended_model_row: preferred("Other Operating Income / Expense"),
      classification_type: "special operating charge",
      is_current: null,
      is_operating: true,
      should_exclude_from_other_bucket: true,
      confidence: "high",
      reason: "The item is presented above operating income and belongs in other operating income/expense."
    };
  }

  if (
    request.statement === "income_statement" &&
    request.section === "operating expenses" &&
    /\bcost\b.*\b(?:sales|revenue|goods|products?|services?|operations?)\b|\b(?:sales|revenue|goods|products?|services?|operations?)\b.*\bcost\b|\bmerchandise costs?\b|\bfulfillment\b.*\b(?:costs?|expense)\b/.test(ownText)
  ) {
    return {
      ...base,
      recommended_model_row: preferred("COGS / Cost of Goods Sold"),
      classification_type: "direct operating cost or cost of revenue",
      is_current: null,
      is_operating: true,
      should_exclude_from_other_bucket: true,
      confidence: "high",
      reason: "Primary income-statement cost of revenue/sales/goods/services lines belong in COGS / Cost of Goods Sold."
    };
  }

  if (
    request.statement === "income_statement" &&
    request.section === "operating expenses" &&
    /\badvertising\b|\bmarketing\b|\bpromotion(?:al)?\b|\bsales and marketing\b|\bselling and marketing\b|\bsales expense\b|\bselling expense\b|\bgeneral and administrative\b|\badministrative expense\b|\bcorporate overhead\b|\bsg&a\b|\bselling\b.*\bgeneral\b.*\badministrative\b/.test(ownText)
  ) {
    return {
      ...base,
      recommended_model_row: preferred("SG&A"),
      classification_type: "selling general and administrative operating expense",
      is_current: null,
      is_operating: true,
      should_exclude_from_other_bucket: true,
      confidence: "high",
      reason: "Advertising, marketing, selling, and administrative expenses are SG&A operating expenses when the template has no more specific row."
    };
  }

  if (/\bdepreciation\b|\bamortization\b|\bdepletion\b|\bd&a\b/.test(ownText)) {
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

  if (/\bspecial items?\b|\brestructuring\b|\bimpairment\b|\bspecial charges?\b|\bintegration costs?\b|\blitigation\b|\bsettlement\b|\baccretion\b/.test(ownText)) {
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

function failedLlmClassification(
  request: FinancialLineItemClassificationRequest,
  fallback: FinancialLineItemClassification,
  status: AccountingLlmStatus,
  message: string
): FinancialLineItemClassification {
  return finalizeClassification(request, {
    ...fallback,
    confidence: "low",
    llm_used: false,
    llm_status: status,
    mapping_passed_validation: false,
    warning: `LLM classification ${status} (${message}); human review is required before this ambiguous SEC line item can be treated as LLM-reviewed.`
  });
}

export function classificationPassesValidation(request: FinancialLineItemClassificationRequest, classification: FinancialLineItemClassification) {
  const text = requestSearchText(request);
  const row = classification.recommended_model_row;
  if (request.isSubtotal || /unmapped|needs review/i.test(row)) return false;
  if (modelRowsMatch(row, "D&A") && request.statement !== "income_statement") return false;
  if (modelRowsMatch(row, "D&A") && /cash flows?|operating activities|reconciliation|supplemental/.test(text)) return false;
  if (modelRowsMatch(row, "Revolver") && /\bcurrent maturit|\bcurrent portion\b.*\blong[-\s]?term debt|convertible|senior notes?/.test(text)) return false;
  if (modelRowsMatch(row, "Deferred Income Taxes") && !classification.is_deferred_tax) return false;
  if ((modelRowsMatch(row, "Accrued Liabilities") || modelRowsMatch(row, "Other Current Liabilities")) && classification.is_debt) return false;
  if (
    requestLooksLikeSelfInsuranceReserve(request, text) &&
    (modelRowsMatch(row, "Revolver") || modelRowsMatch(row, "LT Debt (Incl. Current Portion)"))
  ) return false;
  if (requestLooksLikeSelfInsuranceReserve(request, text)) {
    const nonCurrent = currentNonCurrentSignalFromText(text) === "non-current" || /Noncurrent/i.test(request.xbrlTag ?? "");
    return nonCurrent
      ? modelRowsMatch(row, "Other Non-Current Liabilities")
      : modelRowsMatch(row, "Accrued Liabilities") || modelRowsMatch(row, "Other Current Liabilities");
  }
  if (
    request.statement === "balance_sheet" &&
    (/\bredeemable\b.*\bnon[-\s]?controlling interests?\b|\bnon[-\s]?controlling interests?\b.*\bredeemable\b|\bredeemable nci\b/.test(text) ||
      /RedeemableNoncontrollingInterest/i.test(request.xbrlTag ?? "")) &&
    (modelRowsMatch(row, "Mezzanine Equity") || modelRowsMatch(row, "Other Non-Current Liabilities"))
  ) {
    return true;
  }
  if (modelRowsMatch(row, "Prepaid & Other Current Assets") && /inventor|spare parts?|aircraft fuel|supplies/.test(text)) return false;
  if (modelRowsMatch(row, "Prepaid & Other Current Assets") && /short[-\s]?term investments?|marketable securities|available[-\s]?for[-\s]?sale securities/.test(text)) {
    const hasCurrentInvestmentsRow = request.availableModelRows.some((availableRow) =>
      /short[-\s]?term investments?|current investments?|marketable securities|investment securities/i.test(availableRow) &&
      !/cash.*(short[-\s]?term investments?|current investments?|marketable securities)|(short[-\s]?term investments?|current investments?|marketable securities).*cash/i.test(availableRow)
    );
    const hasCashAndCurrentInvestmentsRow = request.availableModelRows.some((availableRow) =>
      /cash.*(short[-\s]?term investments?|current investments?|marketable securities)|(short[-\s]?term investments?|current investments?|marketable securities).*cash/i.test(availableRow)
    );
    const hasCashRow = request.availableModelRows.some((availableRow) => modelRowsMatch(availableRow, "Cash & Cash Equivalents"));
    if (hasCurrentInvestmentsRow || hasCashAndCurrentInvestmentsRow || hasCashRow) return false;
  }
  if (modelRowsMatch(row, "Cash & Cash Equivalents") && /short[-\s]?term investments?|marketable securities|available[-\s]?for[-\s]?sale securities/.test(text)) {
    const hasCurrentInvestmentsRow = request.availableModelRows.some((availableRow) =>
      /short[-\s]?term investments?|current investments?|marketable securities|investment securities/i.test(availableRow) &&
      !/cash.*(short[-\s]?term investments?|current investments?|marketable securities)|(short[-\s]?term investments?|current investments?|marketable securities).*cash/i.test(availableRow)
    );
    if (hasCurrentInvestmentsRow) return false;
  }
  if (modelRowsMatch(row, "Common Stock & APIC") && /\binvestment securities\b|\bdebt and equity securities\b|\bavailable[-\s]?for[-\s]?sale securities\b|\bmarketable securities\b/.test(text)) return false;
  if (request.section === "current assets" && !modelRowsMatch(row, "Cash & Cash Equivalents") && !modelRowsMatch(row, "Short-Term Investments") && !modelRowsMatch(row, "Accounts Receivable") && !modelRowsMatch(row, "Inventory") && !modelRowsMatch(row, "Prepaid & Other Current Assets")) return false;
  if (request.section === "non-current assets" && !modelRowsMatch(row, "PP&E, Net") && !modelRowsMatch(row, "Intangible Assets, Net") && !modelRowsMatch(row, "Goodwill") && !modelRowsMatch(row, "Other Non-Current Assets")) return false;
  if (request.section === "current liabilities" && !modelRowsMatch(row, "Accounts Payable") && !modelRowsMatch(row, "Accrued Liabilities") && !modelRowsMatch(row, "Other Current Liabilities") && !modelRowsMatch(row, "Revolver") && !modelRowsMatch(row, "LT Debt (Incl. Current Portion)")) return false;
  if (request.section === "non-current liabilities" && !modelRowsMatch(row, "LT Debt (Incl. Current Portion)") && !modelRowsMatch(row, "Deferred Income Taxes") && !modelRowsMatch(row, "Other Non-Current Liabilities")) return false;
  if (request.section === "equity" && !modelRowsMatch(row, "Common Stock & APIC") && !modelRowsMatch(row, "Retained Earnings") && !modelRowsMatch(row, "Treasury Stock") && !modelRowsMatch(row, "AOCI") && !modelRowsMatch(row, "Noncontrolling Interests")) return false;
  if (modelRowsMatch(row, "Other Current Liabilities") && /\baccrued\b.*\b(rebates?|returns?|promotions?|compensation|payroll|tax(?:es)?)\b/.test(text)) return false;
  if (
    request.statement === "balance_sheet" &&
    !balanceSheetSectionCompatible(row, request.section, {
      label: request.cleanLabel || request.reportedLineItemLabel,
      tag: request.xbrlTag
    })
  ) {
    return false;
  }
  if (
    request.statement === "income_statement" &&
    request.section === "operating expenses" &&
    modelRowsMatch(row, "Other Non-Operating Income / Expense")
  ) {
    return false;
  }
  if (
    request.statement === "income_statement" &&
    request.section === "below operating income" &&
    modelRowsMatch(row, "Other Operating Income / Expense")
  ) {
    return false;
  }
  return true;
}

async function requestLlmClassification(
  request: FinancialLineItemClassificationRequest,
  options: LlmClassificationOptions
): Promise<AccountingLlmResult<FinancialLineItemClassification>> {
  const system = [
    "You are a structured accounting classifier for SEC EDGAR financial statement line items.",
    "Classify the reported source line item into the best model template row using accounting meaning, statement section, XBRL tag semantics, parent subtotal, and template row definitions.",
    "Think like a human reviewer with the SEC statement and model open side by side: map the source row to the model row whose accounting definition fits, even when labels do not match one-for-one.",
    ...GENERAL_ACCOUNTING_ROUTING_INSTRUCTIONS,
    "Do not use keyword matching alone. Reported statement location and accounting meaning take precedence over mathematical tie-outs.",
    "Use Other buckets only when no dedicated row exists. Do not use residual plugging.",
    "Examples are non-exhaustive: current investments may belong in a dedicated investments row, a cash/current-investments row, a plain cash row when no dedicated investments row exists, or the current-assets residual only when no cash/current-investment row exists; current maturities may belong with LT debt; advertising may belong with SG&A.",
    "Current maturities/current portion of long-term debt and convertible senior notes belong with LT Debt including current portion, not Revolver.",
    "Debt discounts, premiums, and debt issuance costs are debt carrying-value adjustments. Use the ordered statement context to determine whether they adjust LT Debt or should be excluded because a reported net debt balance already includes them.",
    "Deferred income/revenue is a contract liability, not deferred income taxes. Deferred tax liabilities are Deferred Income Taxes.",
    "Cash-flow-only D&A must not be inserted into income-statement D&A.",
    "When the correct repair is no reported line item, return recommended_action set_zero with explicit_zero_rows populated.",
    "Use recommended_action remap for validation failures caused by a source line belonging in a different row; use exclude for subtotal/component double-counting.",
    "Return strict JSON only."
  ].join(" ");
  return requestAccountingJson<FinancialLineItemClassification>({
    purpose: "line_item_classification",
    apiKey: options.apiKey,
    endpoint: options.endpoint,
    model: options.model,
    fallbackModels: options.fallbackModels,
    siteUrl: options.siteUrl,
    appTitle: options.appTitle,
    timeoutMs: options.timeoutMs ?? 15_000,
    maxTokens: 700,
    reasoningEffort: "low",
    fetchImpl: options.fetchImpl,
    jsonSchema: financialLineItemClassificationJsonSchema(),
    messages: [
      { role: "system", content: system },
      { role: "user", content: JSON.stringify(request) }
    ],
    validate: (value) => validateFinancialLineItemClassification(value, request),
    repair: {
      enabled: true,
      instruction:
        "Repair the classification. It must choose an available model row, pass current/non-current and statement-section validation, and return only JSON matching the schema."
    }
  });
}

async function requestStatementLlmClassification(
  prepared: PreparedLineItemClassification[],
  targets: PreparedLineItemClassification[],
  options: LlmClassificationOptions
): Promise<AccountingLlmResult<StatementLlmClassificationResponse>> {
  const system = [
    "You are a structured accounting classifier for SEC EDGAR financial statement line items.",
    "Do not classify one row in isolation. Review the entire provided statement or statement mapping batch, including sibling labels, parent subtotals, current/non-current sections, XBRL tag semantics, deterministic classifications, and model row definitions.",
    "Think like a human reviewer with the SEC statement and model open side by side: the filing labels and model labels will not always match, so assign by accounting substance and template definitions.",
    "Use deterministic candidates and fallback rows as evidence, validation hints, and guardrails only. They are not the primary mapper for target rows.",
    ...GENERAL_ACCOUNTING_ROUTING_INSTRUCTIONS,
    "Only return classifications for targetSourceRowKeys. Use each source_row_key exactly as provided.",
    "The model template order may differ from the filing statement order. Assign each source line to the model row that best fits the accounting meaning, even if that model row appears much earlier or later in the template.",
    "Use Other buckets only when no dedicated model row exists. Do not use residual plugging.",
    "Examples are non-exhaustive: current investments may belong in a dedicated investments row, a cash/current-investments row, a plain cash row when no dedicated investments row exists, or the current-assets residual only when no cash/current-investment row exists; current maturities may belong with LT debt; advertising may belong with SG&A.",
    "Current marketable securities, available-for-sale securities, and short-term investments belong in a dedicated current-investments row when present. Otherwise group them with Cash & Cash Equivalents if the template has a cash row; use the current-assets residual only when there is no cash or current-investment row.",
    "Current maturities/current portion of long-term debt and convertible senior notes belong with LT Debt including current portion, not Revolver. Short-term borrowings, commercial paper, notes payable current, and revolving facilities may belong in Revolver/current borrowings.",
    "Debt discounts, premiums, and debt issuance costs are debt carrying-value adjustments. Compare gross debt, the adjustment, and reported net debt in statementContexts; exclude supporting detail when net debt already includes it, otherwise map the signed adjustment with LT Debt.",
    "Deferred income/revenue is a contract liability, not deferred income taxes. Deferred tax liabilities are Deferred Income Taxes.",
    "Cash-flow-only D&A must not be inserted into income-statement D&A.",
    "When the correct repair is no reported line item, return recommended_action set_zero with explicit_zero_rows populated.",
    "Use recommended_action remap for validation failures caused by a source line belonging in a different row; use exclude for subtotal/component double-counting.",
    "Return one compact decision per target with only source_row_key, recommended_action, recommended_model_row, confidence, and reason. Semantic flags are derived and validated by the accounting controller.",
    "Return strict JSON only."
  ].join(" ");
  return requestAccountingJson<StatementLlmClassificationResponse>({
    purpose: "statement_line_item_classification",
    apiKey: options.apiKey,
    endpoint: options.endpoint,
    model: options.model,
    fallbackModels: options.fallbackModels?.filter((model) => !/^openrouter\/free$/i.test(model.trim())),
    siteUrl: options.siteUrl,
    appTitle: options.appTitle,
    timeoutMs: options.timeoutMs ?? 15_000,
    maxTokens: Math.max(3_000, Math.min(12_000, 1_200 + targets.length * 400)),
    maxAttemptsPerModel: /(?:deepseek-v4-flash|^openrouter\/free|:free$)/i.test(options.model.trim()) ? 2 : 1,
    reasoningEffort: "low",
    fetchImpl: options.fetchImpl,
    jsonSchema: financialStatementLineItemClassificationJsonSchema(),
    messages: [
      { role: "system", content: system },
      { role: "user", content: JSON.stringify(statementLlmClassificationPayload(prepared, targets)) }
    ],
    validate: (value) => validateStatementLineItemClassification(value, targets),
    repair: {
      enabled: false,
      instruction:
        "Repair the statement classification batch. Return one compact classification for every targetSourceRowKey, use source_row_key exactly, choose available model rows, and return only JSON matching the schema."
    }
  });
}

function statementLlmClassificationPayload(prepared: PreparedLineItemClassification[], targets: PreparedLineItemClassification[]) {
  const first = prepared[0]?.request;
  return {
    company: first?.company,
    filing: first?.filing,
    fiscalPeriod: first?.fiscalPeriod,
    statement: first?.statement,
    sourceTableType: first?.sourceTableType,
    targetSourceRowKeys: targets.map((item) => item.rowKey),
    availableModelRows: first?.availableModelRows ?? [],
    modelRowDefinitions: first?.modelRowDefinitions ?? {},
    alreadyMappedRows: first?.alreadyMappedRows ?? [],
    statementContexts: statementLlmContexts(prepared),
    classificationGoal:
      "Classify only the target rows after reviewing the supplied SEC statement context as if the SEC statement and model template were open side by side. Batches can contain distinct semantic rows from multiple EDGAR filings; use each row's filing, fiscal-period, section, and nearby-row context. A row can map to any available model row whose accounting definition fits; filing order and model order do not need to match. Deterministic candidates are evidence and validation guardrails, not the primary mapper. The examples in the system prompt are non-exhaustive; apply the same accounting-substance reasoning to any random SEC line item.",
    routingPrinciples: GENERAL_ACCOUNTING_ROUTING_INSTRUCTIONS,
    statementRows: prepared
      .slice()
      .sort((a, b) => (a.request.rowOrder ?? 0) - (b.request.rowOrder ?? 0))
      .map((item) => ({
        sourceRowKey: item.rowKey,
        filing: item.request.filing,
        fiscalPeriod: item.request.fiscalPeriod,
        statement: item.request.statement,
        rowOrder: item.request.rowOrder ?? null,
        target: targets.some((target) => target.rowKey === item.rowKey),
        targetReason: item.fullStatementTarget
          ? "complete_primary_statement_coverage"
          : item.materialAnalystTarget
            ? "material_pre_fill_analyst_pass"
            : item.needsLlm
              ? "ambiguous_or_validation_required"
              : "",
        reportedLineItemLabel: item.request.reportedLineItemLabel,
        cleanLabel: item.request.cleanLabel,
        xbrlTag: item.request.xbrlTag ?? "",
        amount: item.request.amount ?? null,
        unit: item.request.unit ?? "",
        periodType: item.request.periodType,
        section: item.request.section,
        parentSubtotal: item.request.parentSubtotal ?? null,
        isSubtotal: item.request.isSubtotal,
        nearbyRows: item.request.nearbyRows,
        priorPeriodSourceLabels: item.request.priorPeriodSourceLabels ?? [],
        deterministicCandidate: item.request.deterministicCandidate ?? "",
        deterministicClassification: item.deterministicIsValidated
          ? {
              recommendedModelRow: item.initialClassification.recommended_model_row,
              confidence: item.initialClassification.confidence,
              reason: item.initialClassification.reason
            }
          : null,
        fallbackModelRow: item.fallback.recommended_model_row,
        uncertaintyReason: item.request.uncertaintyReason,
        validationError: item.request.validationError ?? ""
      }))
  };
}

function statementLlmContexts(prepared: PreparedLineItemClassification[]) {
  const contexts = new Map<
    string,
    {
      filing: FinancialLineItemClassificationRequest["filing"];
      fiscalPeriod: string;
      statement: FinancialStatementName;
      sourceTableType: FinancialSourceTableType;
      orderedSourceLines: string[];
    }
  >();
  for (const item of prepared) {
    const request = item.request;
    const key = [normalizeAccession(request.filing.accession), request.fiscalPeriod, request.statement].join("|");
    if (contexts.has(key)) continue;
    contexts.set(key, {
      filing: request.filing,
      fiscalPeriod: request.fiscalPeriod,
      statement: request.statement,
      sourceTableType: request.sourceTableType,
      orderedSourceLines: request.currentPeriodSourceLines ?? []
    });
  }
  return Array.from(contexts.values());
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

function financialStatementLineItemClassificationJsonSchema() {
  return {
    name: "financial_statement_line_item_classification",
    strict: true,
    schema: {
      type: "object",
      additionalProperties: false,
      required: ["classifications"],
      properties: {
        classifications: {
          type: "array",
          items: {
            type: "object",
            additionalProperties: false,
            required: ["source_row_key", "recommended_action", "recommended_model_row", "confidence", "reason"],
            properties: {
              source_row_key: { type: "string" },
              recommended_action: {
                type: "string",
                enum: ["map", "remap", "set_zero", "merge_into_other", "split_across_rows", "keep_existing", "exclude"]
              },
              recommended_model_row: { type: "string" },
              confidence: { type: "string", enum: ["high", "medium", "low"] },
              reason: { type: "string" }
            }
          }
        }
      }
    }
  };
}

function statementDecisionToClassification(
  target: PreparedLineItemClassification,
  decision: StatementLlmClassificationDecision
): FinancialLineItemClassification {
  const requestedRow = normalizeModelRow(decision.recommended_model_row) || decision.recommended_model_row;
  const fallback = target.fallback;
  const mapsToDebt = modelRowsMatch(requestedRow, "Revolver") || modelRowsMatch(requestedRow, "LT Debt (Incl. Current Portion)");
  const mapsToNonDebtLiability =
    modelRowsMatch(requestedRow, "Accounts Payable") ||
    modelRowsMatch(requestedRow, "Accrued Liabilities") ||
    modelRowsMatch(requestedRow, "Other Current Liabilities") ||
    modelRowsMatch(requestedRow, "Other Non-Current Liabilities");
  const current =
    modelRowsMatch(requestedRow, "Other Non-Current Liabilities") || modelRowsMatch(requestedRow, "Deferred Income Taxes")
      ? false
      : modelRowsMatch(requestedRow, "Accounts Payable") ||
          modelRowsMatch(requestedRow, "Accrued Liabilities") ||
          modelRowsMatch(requestedRow, "Other Current Liabilities") ||
          modelRowsMatch(requestedRow, "Revolver")
        ? true
        : fallback.is_current;
  const actionIsNonMapping =
    decision.recommended_action === "exclude" ||
    decision.recommended_action === "set_zero" ||
    decision.recommended_action === "keep_existing";
  const classification: FinancialLineItemClassification = {
    ...fallback,
    source_line_item: target.request.cleanLabel || target.request.reportedLineItemLabel,
    recommended_action: decision.recommended_action,
    recommended_model_row: actionIsNonMapping ? requestedRow || fallback.recommended_model_row : requestedRow,
    recommended_model_row_mappings: [],
    explicit_zero_rows:
      decision.recommended_action === "set_zero" && requestedRow
        ? [{ model_row: requestedRow, reason: decision.reason }]
        : [],
    classification_type: actionIsNonMapping ? `LLM ${decision.recommended_action}` : `LLM assignment to ${requestedRow}`,
    is_current: current,
    is_debt: mapsToDebt ? true : mapsToNonDebtLiability ? false : fallback.is_debt,
    is_operating:
      modelRowsMatch(requestedRow, "Other Non-Operating Income / Expense") ||
      modelRowsMatch(requestedRow, "Interest Income") ||
      modelRowsMatch(requestedRow, "Interest Expense")
        ? false
        : target.request.section === "operating expenses"
          ? true
          : fallback.is_operating,
    is_deferred_tax: modelRowsMatch(requestedRow, "Deferred Income Taxes") ? true : fallback.is_deferred_tax,
    should_exclude_from_other_bucket: !isReusableOtherBucketModelRow(requestedRow),
    confidence: decision.confidence,
    reason: decision.reason,
    requires_validation: true,
    requires_revalidation: true,
    llm_used: false,
    mapping_passed_validation: false
  };
  classification.mapping_passed_validation =
    classificationPassesValidation(target.request, classification) ||
    statementDecisionPassesStructuralGuardrails(target.request, classification);
  return classification;
}

function statementDecisionPassesStructuralGuardrails(
  request: FinancialLineItemClassificationRequest,
  classification: FinancialLineItemClassification
) {
  if (["exclude", "set_zero", "keep_existing"].includes(classification.recommended_action)) return true;
  const row = classification.recommended_model_row;
  if (!row || /unmapped|needs review/i.test(row) || !modelRowAvailable(row, request.availableModelRows)) return false;
  const text = requestSearchText(request);
  if (request.statement === "balance_sheet") {
    if (!balanceSheetRowDefinitionForLabel(row)) return false;
    if (
      !balanceSheetSectionCompatible(row, request.section, {
        label: request.cleanLabel || request.reportedLineItemLabel,
        tag: request.xbrlTag
      })
    ) {
      return false;
    }
    if (modelRowsMatch(row, "Revolver") && /\bcurrent maturit|\bcurrent portion\b.*\blong[-\s]?term debt|convertible|senior notes?/.test(text)) {
      return false;
    }
    if (
      requestLooksLikeSelfInsuranceReserve(request, text) &&
      (modelRowsMatch(row, "Revolver") || modelRowsMatch(row, "LT Debt (Incl. Current Portion)"))
    ) {
      return false;
    }
    if (modelRowsMatch(row, "Deferred Income Taxes") && !/\bdeferred\b.*\btax|\btax\b.*\bdeferred/.test(text)) return false;
    return true;
  }
  if (request.statement !== "income_statement") return false;
  if (request.section === "revenue") return modelRowsMatch(row, "Revenue");
  if (request.section === "operating expenses") {
    return ["COGS / Cost of Goods Sold", "SG&A", "R&D", "D&A", "Other Operating Income / Expense"].some((candidate) =>
      modelRowsMatch(row, candidate)
    );
  }
  if (request.section === "below operating income") {
    return ["Interest Income", "Interest Expense", "Goodwill Impairment", "Other Non-Operating Income / Expense"].some((candidate) =>
      modelRowsMatch(row, candidate)
    );
  }
  if (request.section === "tax") return modelRowsMatch(row, "Income Tax Benefit / Expense");
  return false;
}

function validateFinancialLineItemClassification(
  value: unknown,
  request: FinancialLineItemClassificationRequest
): AccountingLlmValidationResult<FinancialLineItemClassification> {
  const shape = validateFinancialLineItemClassificationShape(value);
  if (!shape.ok) return shape;
  const normalized = {
    ...shape.value,
    recommended_model_row: normalizeModelRow(shape.value.recommended_model_row) || shape.value.recommended_model_row
  };
  const actionIsValidNonMapping =
    normalized.recommended_action === "exclude" ||
    normalized.recommended_action === "set_zero" ||
    normalized.recommended_action === "keep_existing";
  const validated =
    actionIsValidNonMapping ||
    classificationPassesValidation(request, {
      ...normalized,
      mapping_passed_validation: false
    });
  if (!validated) {
    return {
      ok: false,
      needsHumanReview: true,
      error: `${request.cleanLabel || request.reportedLineItemLabel}: LLM decision failed deterministic accounting validation for ${normalized.recommended_model_row}.`
    };
  }
  return {
    ok: true,
    value: normalized,
    validated: true,
    affectedOutput: true
  };
}

function validateStatementLineItemClassification(
  value: unknown,
  targets: PreparedLineItemClassification[]
): AccountingLlmValidationResult<StatementLlmClassificationResponse> {
  const rawClassifications = statementClassificationArray(value, targets.length);
  if (!rawClassifications) {
    return {
      ok: false,
      needsHumanReview: true,
      error: "Statement classifier response must contain a classifications, decisions, mappings, or items array."
    };
  }
  const targetByKey = new Map(targets.map((target) => [target.rowKey, target]));
  const seen = new Set<string>();
  const classifications: StatementLlmClassificationDecision[] = [];
  for (let index = 0; index < rawClassifications.length; index += 1) {
    const item = rawClassifications[index];
    if (!isRecord(item)) {
      return { ok: false, needsHumanReview: true, error: "Every statement classification must be a JSON object." };
    }
    const sourceRowKey = statementDecisionSourceRowKey(item, targets, rawClassifications.length, index);
    if (!sourceRowKey) {
      return { ok: false, needsHumanReview: true, error: "Every statement classification must include source_row_key." };
    }
    const target = targetByKey.get(sourceRowKey);
    if (!target) {
      return { ok: false, needsHumanReview: true, error: `LLM returned unexpected source_row_key ${sourceRowKey}.` };
    }
    if (seen.has(sourceRowKey)) {
      return { ok: false, needsHumanReview: true, error: `LLM returned duplicate source_row_key ${sourceRowKey}.` };
    }
    const recommendedAction = normalizeStatementDecisionAction(
      firstString(item, ["recommended_action", "recommendedAction", "action", "operation"])
    );
    const confidence = normalizeStatementDecisionConfidence(firstString(item, ["confidence", "certainty"]));
    if (!recommendedAction) return { ok: false, needsHumanReview: true, error: `${sourceRowKey}: recommended_action is invalid.` };
    const rawRecommendedRow = firstString(item, [
      "recommended_model_row",
      "recommendedModelRow",
      "model_row",
      "modelRow",
      "target_row",
      "targetRow"
    ]);
    const rawReason = firstString(item, ["reason", "rationale", "explanation"]);
    if (rawRecommendedRow === null && !["exclude", "keep_existing"].includes(recommendedAction)) {
      return { ok: false, needsHumanReview: true, error: `${sourceRowKey}: recommended_model_row must be a string.` };
    }
    let normalized: StatementLlmClassificationDecision = {
      source_row_key: sourceRowKey,
      recommended_action: recommendedAction,
      recommended_model_row: normalizeModelRow(rawRecommendedRow ?? "") || rawRecommendedRow || target.fallback.recommended_model_row,
      confidence,
      reason: rawReason || "LLM whole-statement accounting classification."
    };
    const expanded = statementDecisionToClassification(target, normalized);
    const actionIsValidNonMapping =
      normalized.recommended_action === "exclude" ||
      normalized.recommended_action === "set_zero" ||
      normalized.recommended_action === "keep_existing";
    const validated = actionIsValidNonMapping || expanded.mapping_passed_validation;
    if (!validated || normalized.confidence === "low") {
      if (target.deterministicIsValidated) {
        normalized = {
          source_row_key: sourceRowKey,
          recommended_action: target.initialClassification.recommended_action,
          recommended_model_row: target.initialClassification.recommended_model_row,
          confidence: "medium",
          reason: !validated
            ? `LLM recommendation ${normalized.recommended_model_row || normalized.recommended_action} was rejected by accounting validation; retained the independently validated mapping ${target.initialClassification.recommended_model_row}.`
            : `LLM returned low confidence; retained the independently validated mapping ${target.initialClassification.recommended_model_row}.`
        };
      } else {
        return {
          ok: false,
          needsHumanReview: true,
          error: !validated
            ? `${target.request.cleanLabel || target.request.reportedLineItemLabel}: LLM decision failed deterministic accounting validation for ${normalized.recommended_model_row}.`
            : `${target.request.cleanLabel || target.request.reportedLineItemLabel}: LLM returned low confidence without a validated fallback.`
        };
      }
    }
    classifications.push(normalized);
    seen.add(sourceRowKey);
  }
  const missing = targets.filter((target) => !seen.has(target.rowKey));
  if (missing.length) {
    return {
      ok: false,
      needsHumanReview: true,
      error: `Statement classifier omitted target row(s): ${missing.map((target) => target.rowKey).join(", ")}.`
    };
  }
  return {
    ok: true,
    value: { classifications },
    validated: true,
    affectedOutput: classifications.length > 0
  };
}

function statementClassificationArray(value: unknown, targetCount: number): unknown[] | null {
  if (Array.isArray(value)) return value;
  if (!isRecord(value)) return null;
  for (const key of ["classifications", "decisions", "mappings", "items", "results"]) {
    if (Array.isArray(value[key])) return value[key];
  }
  if (targetCount === 1 && statementDecisionLooksLikeSingleItem(value)) return [value];
  return null;
}

function statementDecisionLooksLikeSingleItem(value: Record<string, any>) {
  return ["recommended_action", "recommendedAction", "action", "recommended_model_row", "recommendedModelRow", "modelRow"].some(
    (key) => key in value
  );
}

function statementDecisionSourceRowKey(
  item: Record<string, any>,
  targets: PreparedLineItemClassification[],
  decisionCount: number,
  index: number
) {
  const explicit = firstString(item, ["source_row_key", "sourceRowKey", "row_key", "rowKey"]);
  if (explicit) return explicit;
  const label = firstString(item, ["source_line_item", "sourceLineItem", "reportedLineItemLabel", "label"]);
  if (label) {
    const normalizedLabel = normalizeKey(label);
    const matches = targets.filter(
      (target) =>
        normalizeKey(target.request.cleanLabel || target.request.reportedLineItemLabel) === normalizedLabel ||
        normalizeKey(target.request.reportedLineItemLabel) === normalizedLabel
    );
    if (matches.length === 1) return matches[0].rowKey;
  }
  return decisionCount === targets.length ? targets[index]?.rowKey ?? null : null;
}

function normalizeStatementDecisionAction(value: string | null): FinancialLineItemClassification["recommended_action"] | null {
  if (!value) return null;
  const normalized = value.trim().toLowerCase().replace(/[\s-]+/g, "_");
  const aliases: Record<string, FinancialLineItemClassification["recommended_action"]> = {
    assign: "map",
    assigned: "map",
    direct: "map",
    mapped: "map",
    mapping: "map",
    reassign: "remap",
    reclassified: "remap",
    group: "merge_into_other",
    grouped: "merge_into_other",
    merge: "merge_into_other",
    group_into_other: "merge_into_other",
    omit: "exclude",
    omitted: "exclude",
    excluded: "exclude",
    skip: "exclude",
    zero: "set_zero",
    keep: "keep_existing"
  };
  if (/^(?:map|assign|direct)(?:_|\b)/.test(normalized)) return "map";
  if (/^(?:remap|reassign|reclassif)/.test(normalized)) return "remap";
  if (/^(?:merge|group)/.test(normalized)) return "merge_into_other";
  if (/^(?:exclude|omit|skip)/.test(normalized)) return "exclude";
  if (/^(?:zero|set_zero)/.test(normalized)) return "set_zero";
  if (/^(?:keep|retain)/.test(normalized)) return "keep_existing";
  return (
    stringEnum(normalized, ["map", "remap", "set_zero", "merge_into_other", "split_across_rows", "keep_existing", "exclude"]) ||
    aliases[normalized] ||
    null
  );
}

function normalizeStatementDecisionConfidence(value: string | null): FinancialLineItemClassification["confidence"] {
  const normalized = value?.trim().toLowerCase() ?? "";
  if (/^(high|certain|strong)$/.test(normalized)) return "high";
  if (/^(low|uncertain|weak)$/.test(normalized)) return "low";
  return "medium";
}

function firstString(value: Record<string, any>, keys: string[]) {
  for (const key of keys) {
    if (typeof value[key] === "string" && value[key].trim()) return value[key].trim();
  }
  return null;
}

function validateFinancialLineItemClassificationShape(value: unknown): AccountingLlmValidationResult<FinancialLineItemClassification> {
  if (!isRecord(value)) return { ok: false, needsHumanReview: true, error: "Classification must be a JSON object." };
  const recommendedAction = stringEnum(value.recommended_action, [
    "map",
    "remap",
    "set_zero",
    "merge_into_other",
    "split_across_rows",
    "keep_existing",
    "exclude"
  ]);
  const confidence = stringEnum(value.confidence, ["high", "medium", "low"]);
  const requiredStrings = ["source_line_item", "recommended_model_row", "classification_type", "reason"];
  const missingString = requiredStrings.find((key) => typeof value[key] !== "string");
  if (missingString) return { ok: false, needsHumanReview: true, error: `Classification field ${missingString} must be a string.` };
  if (!recommendedAction) return { ok: false, needsHumanReview: true, error: "Classification recommended_action is invalid." };
  if (!confidence) return { ok: false, needsHumanReview: true, error: "Classification confidence is invalid." };
  if (!Array.isArray(value.recommended_model_row_mappings)) {
    return { ok: false, needsHumanReview: true, error: "recommended_model_row_mappings must be an array." };
  }
  if (!Array.isArray(value.explicit_zero_rows)) {
    return { ok: false, needsHumanReview: true, error: "explicit_zero_rows must be an array." };
  }
  const booleanKeys = [
    "is_debt",
    "is_tax_related",
    "is_deferred_revenue_or_contract_liability",
    "is_deferred_tax",
    "is_subtotal",
    "should_exclude_from_other_bucket",
    "requires_validation",
    "requires_revalidation"
  ];
  const badBoolean = booleanKeys.find((key) => typeof value[key] !== "boolean");
  if (badBoolean) return { ok: false, needsHumanReview: true, error: `Classification field ${badBoolean} must be boolean.` };
  if (!booleanOrNull(value.is_current)) return { ok: false, needsHumanReview: true, error: "is_current must be boolean or null." };
  if (!booleanOrNull(value.is_operating)) return { ok: false, needsHumanReview: true, error: "is_operating must be boolean or null." };
  return {
    ok: true,
    validated: false,
    affectedOutput: false,
    value: {
      source_line_item: value.source_line_item,
      recommended_action: recommendedAction,
      recommended_model_row: value.recommended_model_row,
      recommended_model_row_mappings: value.recommended_model_row_mappings.map((item) => ({
        source_line_item: isRecord(item) && typeof item.source_line_item === "string" ? item.source_line_item : "",
        model_row: isRecord(item) && typeof item.model_row === "string" ? item.model_row : "",
        amount: isRecord(item) && typeof item.amount === "number" ? item.amount : null,
        reason: isRecord(item) && typeof item.reason === "string" ? item.reason : ""
      })),
      explicit_zero_rows: value.explicit_zero_rows.map((item) => ({
        model_row: isRecord(item) && typeof item.model_row === "string" ? item.model_row : "",
        reason: isRecord(item) && typeof item.reason === "string" ? item.reason : ""
      })),
      classification_type: value.classification_type,
      is_current: value.is_current,
      is_debt: value.is_debt,
      is_operating: value.is_operating,
      is_tax_related: value.is_tax_related,
      is_deferred_revenue_or_contract_liability: value.is_deferred_revenue_or_contract_liability,
      is_deferred_tax: value.is_deferred_tax,
      is_subtotal: value.is_subtotal,
      should_exclude_from_other_bucket: value.should_exclude_from_other_bucket,
      confidence,
      reason: value.reason,
      requires_validation: value.requires_validation,
      requires_revalidation: value.requires_revalidation,
      llm_used: false,
      mapping_passed_validation: false
    }
  };
}

function isRecord(value: unknown): value is Record<string, any> {
  return Boolean(value && typeof value === "object" && !Array.isArray(value));
}

function stringEnum<T extends string>(value: unknown, allowed: T[]): T | null {
  return typeof value === "string" && (allowed as string[]).includes(value) ? (value as T) : null;
}

function booleanOrNull(value: unknown) {
  return typeof value === "boolean" || value === null;
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

function sourceRowKeyForRequest(request: FinancialLineItemClassificationRequest, index: number) {
  if (request.sourceRowKey) return request.sourceRowKey;
  const key = classificationSourceKeys({
    period: request.fiscalPeriod,
    accession: request.filing.accession,
    xbrlTag: request.xbrlTag,
    label: request.cleanLabel || request.reportedLineItemLabel,
    amount: request.amount
  })[0];
  return key || `${request.statement}:${request.fiscalPeriod}:${request.rowOrder ?? index}:${normalizeKey(request.cleanLabel || request.reportedLineItemLabel)}`;
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
    "Revenue",
    "COGS / Cost of Goods Sold",
    "SG&A",
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

function requestOwnSearchText(request: FinancialLineItemClassificationRequest) {
  return [
    request.reportedLineItemLabel,
    request.cleanLabel,
    request.xbrlTag ?? "",
    request.section
  ]
    .join(" ")
    .replace(/([a-z])([A-Z])/g, "$1 $2")
    .toLowerCase();
}

function requestLooksLikeSelfInsuranceReserve(request: FinancialLineItemClassificationRequest, text = requestSearchText(request)) {
  return /\bself[-\s]?insurance reserves?\b/.test(text) || /SelfInsuranceReserve/i.test(request.xbrlTag ?? "");
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
