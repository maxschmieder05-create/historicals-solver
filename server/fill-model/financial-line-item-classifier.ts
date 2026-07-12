import { normalizeAccession } from "./sec-accession";
import {
  balanceSheetRowDefinitionForLabel,
  balanceSheetRowsEquivalent,
  balanceSheetSectionCompatible,
  classifyBalanceSheetSourceSection,
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
  deadlineAt?: number;
  maxAttempts?: number;
  fetchImpl?: typeof fetch;
  signal?: AbortSignal;
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
  decision_origin: "validated_llm" | "deterministic_fallback";
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
    ["period-source", period, accession, concept, label, amount],
    ["period-concept-label", period, concept, label, amount],
    ["accession-concept-label", accession, concept, label, amount]
  ]
    .map((parts) => parts.filter(Boolean).join("|"))
    .filter(Boolean);
  return Array.from(new Set(keys));
}

export function lineItemNeedsClassification(request: FinancialLineItemClassificationRequest) {
  if (request.isSubtotal) return false;
  if (request.sourceTableType !== "primary_statement" && request.sourceTableType !== "cash_flow_reconciliation") return false;
  if (aggregateOperatingExpenseSourceHasReportedComponents(request)) return false;
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
  if (aggregateOperatingExpenseSourceHasReportedComponents(request)) return false;
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

const DEFAULT_STATEMENT_LLM_ATTEMPT_BUDGET = 1;

type StatementLlmAttemptBudget = {
  remaining: number;
};

export async function classifyFinancialStatementLineItems(
  requests: FinancialLineItemClassificationRequest[],
  options: ClassifierOptions = {}
): Promise<FinancialStatementLineItemClassificationResult> {
  const configuredAttemptBudget = options.llm?.maxAttempts;
  const attemptBudget = Number.isFinite(configuredAttemptBudget)
    ? Math.max(0, Math.floor(configuredAttemptBudget!))
    : DEFAULT_STATEMENT_LLM_ATTEMPT_BUDGET;
  return classifyFinancialStatementLineItemsWithinBudget(requests, options, { remaining: attemptBudget });
}

async function classifyFinancialStatementLineItemsWithinBudget(
  requests: FinancialLineItemClassificationRequest[],
  options: ClassifierOptions,
  attemptBudget: StatementLlmAttemptBudget
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
  if (attemptBudget.remaining <= 0) {
    return statementLlmBudgetExhaustedResult(classificationTargets, targets);
  }
  const result = await requestStatementLlmClassification(prepared, targets, llm, attemptBudget.remaining);
  const telemetry = result.attemptTelemetry ?? [result.telemetry];
  const currentAttempts = telemetry.filter((item) => item.attempted).length;
  attemptBudget.remaining = Math.max(0, attemptBudget.remaining - currentAttempts);
  if (result.value) {
    const response = result.value;
    const byRowKey = new Map(response.classifications.map((item) => [item.source_row_key, item]));
    let merged = classificationTargets.map((item) => {
      const llmDecision = item.needsLlm ? byRowKey.get(item.rowKey) : null;
      if (!llmDecision) return { request: item.request, classification: item.initialClassification };
      return {
        request: item.request,
        classification: finalizeClassification(item.request, {
          ...statementDecisionToClassification(item, llmDecision),
          llm_used: statementDecisionWasAcceptedFromLlm(llmDecision) && result.telemetry.affectedOutput,
          llm_status: result.status
        })
      };
    });
    const missingTargets = targets.filter((item) => {
      const decision = byRowKey.get(item.rowKey);
      return !decision || !statementDecisionWasAcceptedFromLlm(decision);
    });
    let retryResult: FinancialStatementLineItemClassificationResult | null = null;
    if (missingTargets.length && attemptBudget.remaining > 0 && !llm.signal?.aborted) {
      retryResult = await classifyFinancialStatementLineItemsWithinBudget(
        missingTargets.map((item) => item.request),
        options,
        attemptBudget
      );
      const retryByKey = new Map(
        retryResult.classifications.map((item, index) => [sourceRowKeyForRequest(item.request, index), item.classification])
      );
      merged = merged.map((item) => {
        const key = sourceRowKeyForRequest(item.request, 0);
        const retried = retryByKey.get(key);
        return retried ? { request: item.request, classification: retried } : item;
      });
    }
    const warnings = retryResult
      ? retryResult.warnings
      : missingTargets.map(
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
      llmCalls: telemetry.filter((item) => item.completed).length + (retryResult?.llmCalls ?? 0),
      llmAttempts: telemetry.filter((item) => item.attempted).length + (retryResult?.llmAttempts ?? 0),
      llmSuccessfulCompletions: telemetry.filter((item) => item.completed).length + (retryResult?.llmSuccessfulCompletions ?? 0),
      llmTelemetry: [...telemetry, ...(retryResult?.llmTelemetry ?? [])],
      targetCount: classificationTargets.length,
      llmReviewedCount: reviewed.length,
      acceptedDecisionCount: accepted.length,
      unreviewedTargetKeys
    };
  }
  if (attemptBudget.remaining > 0 && !llm.signal?.aborted && statementClassificationFailureShouldSplit(result, requests.length)) {
    const midpoint = Math.ceil(requests.length / 2);
    const left = await classifyFinancialStatementLineItemsWithinBudget(requests.slice(0, midpoint), options, attemptBudget);
    const right = await classifyFinancialStatementLineItemsWithinBudget(requests.slice(midpoint), options, attemptBudget);
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

function statementLlmBudgetExhaustedResult(
  classificationTargets: PreparedLineItemClassification[],
  targets: PreparedLineItemClassification[]
): FinancialStatementLineItemClassificationResult {
  const message = "statement-level LLM request-attempt budget was exhausted";
  return {
    classifications: classificationTargets.map((item) => ({
      request: item.request,
      classification: item.needsLlm
        ? failedLlmClassification(item.request, item.fallback, "attempted_failed", message)
        : item.initialClassification
    })),
    warnings: targets.map(
      (item) => `${item.request.cleanLabel || item.request.reportedLineItemLabel}: ${message}; deterministic fallback was used.`
    ),
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
  return classification.mapping_passed_validation;
}

function statementDecisionWasAcceptedFromLlm(decision: StatementLlmClassificationDecision) {
  return decision.decision_origin === "validated_llm";
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
  if (aggregateOperatingExpenseSourceHasReportedComponents(request)) return false;
  const text = `${request.cleanLabel || request.reportedLineItemLabel} ${request.xbrlTag ?? ""}`.toLowerCase();
  const tagCompact = (request.xbrlTag ?? "").toLowerCase().replace(/[^a-z0-9]/g, "");
  if (
    /\bper (?:common )?share\b|\bearnings per share\b|\beps\b|\bweighted average (?:number of )?shares?\b|\bshares? (?:outstanding|used|weighted)\b|\bnumber of shares?\b/.test(
      text
    ) ||
    /earningspershare|incomelosspercommonshare|weightedaveragenumberof.*shares|commonstocksharesoutstanding|preferredstocksharesoutstanding|dividendspershare/.test(
      tagCompact
    )
  ) {
    return false;
  }
  if (
    request.statement === "income_statement" &&
    (/\bgross profit\b|\bgross margin\b|\boperating income\b|\boperating loss\b|\bincome (?:loss )?from operations\b|\bincome before (?:provision for )?(?:income )?tax|\bpretax (?:income|loss)\b|\bnet income\b|\bnet loss\b|\bprofit loss\b/.test(
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

function aggregateOperatingExpenseSourceHasReportedComponents(request: FinancialLineItemClassificationRequest) {
  if (request.statement !== "income_statement" || request.section !== "operating expenses") return false;
  const concept = normalizeKey((request.xbrlTag ?? "").split(":").pop() ?? "");
  const label = normalizeKey(request.cleanLabel || request.reportedLineItemLabel);
  const aggregatePattern = /^(?:total)?(?:operating(?:costs?and)?expenses?(?:andother)?|costsandexpenses(?:andother)?)$/;
  if (!aggregatePattern.test(concept) && !aggregatePattern.test(label)) return false;

  const componentPattern =
    /costof.*(?:revenue|sales|goods|products?|services?|sold)|researchanddevelopment|productdevelopment|sellingandmarketing|generalandadministrative|sellinggeneralandadministrative|depreciation|amortization|restructuring|impairment|specialcharges?|acquiredinprocess|technologyandcontent|fulfillment|fuelandpurchasedpower|purchasedpower|otheroperating(?:income|expense)/;
  return (request.currentPeriodSourceLines ?? []).some((sourceLine) => {
    const normalized = normalizeKey(sourceLine);
    return Boolean(normalized && normalized !== label && !aggregatePattern.test(normalized) && componentPattern.test(normalized));
  });
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

  if (request.statement === "income_statement" && request.section === "below operating income") {
    const semantics = trustedSourceSemantics(request);
    const target = semantics.netInterestExpense
      ? preferred("Interest Expense")
      : semantics.netInterestIncome
        ? preferred("Interest Income")
        : semantics.combinedNetInterest
          ? preferred("Other Non-Operating Income / Expense")
          : null;
    if (target) {
      return {
        ...base,
        recommended_model_row: target,
        classification_type: semantics.netInterestExpense
          ? "net non-operating interest expense"
          : semantics.netInterestIncome
            ? "net non-operating interest income"
            : "combined net interest income expense",
        is_current: null,
        is_operating: false,
        should_exclude_from_other_bucket: !modelRowsMatch(target, "Other Non-Operating Income / Expense"),
        confidence: "high",
        reason: semantics.netInterestExpense
          ? "A primary-statement interest expense explicitly net of interest income remains an Interest Expense line."
          : semantics.netInterestIncome
            ? "A primary-statement interest income line explicitly net of interest expense remains an Interest Income line."
            : "An undirected combined interest income/expense net line remains in Other Non-Operating Income / Expense rather than being forced into one interest direction."
      };
    }
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
    (/\bcost\b.*\b(?:sales|revenue|goods|products?|services?|operations?)\b|\b(?:sales|revenue|goods|products?|services?|operations?)\b.*\bcost\b|\bmerchandise costs?\b|\bfulfillment\b.*\b(?:costs?|expense)\b/.test(ownText) ||
      /\b(?:company[-\s]?(?:owned(?:\s+and\s+operated)?|operated)|franchised|other) restaurants?\b.*\b(?:costs?|expenses?)\b|\brestaurants?\b.*\b(?:occupancy|operating) expenses?\b/.test(text))
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
  const mappingPassedValidation = classificationPassesValidation(request, { ...classification, recommended_model_row: recommended });
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

type TrustedSourceSemantics = ReturnType<typeof trustedSourceSemantics>;

const INCOME_STATEMENT_CLASSIFICATION_ROWS = [
  "Revenue",
  "COGS / Cost of Goods Sold",
  "SG&A",
  "R&D",
  "D&A",
  "Other Operating Income / Expense",
  "Interest Income",
  "Interest Expense",
  "Goodwill Impairment",
  "Other Non-Operating Income / Expense",
  "Income Tax Benefit / Expense"
] as const;

export function classificationPassesValidation(request: FinancialLineItemClassificationRequest, classification: FinancialLineItemClassification) {
  const action = classification.recommended_action ?? "map";
  if (action === "exclude") {
    return request.isSubtotal || analystSupportedSubtotalExclusion(classification);
  }
  if (action === "set_zero" || action === "keep_existing" || action === "split_across_rows") return false;
  if (action !== "map" && action !== "remap" && action !== "merge_into_other") return false;
  if (request.isSubtotal) return false;
  if (aggregateOperatingExpenseSourceHasReportedComponents(request)) return false;
  if (request.sourceTableType !== "primary_statement") return false;

  const requestedRow = normalizeModelRow(classification.recommended_model_row ?? "") || classification.recommended_model_row;
  if (!requestedRow || /unmapped|needs review/i.test(requestedRow)) return false;
  const availableRow = request.availableModelRows.find((row) => modelRowsMatch(row, requestedRow));
  if (!availableRow) return false;

  const semantics = trustedSourceSemantics(request);
  if (request.statement === "balance_sheet") {
    if (request.periodType !== "instant") return false;
    return balanceSheetClassificationPassesTrustedValidation(request, requestedRow, availableRow, action, semantics);
  }
  if (request.statement === "income_statement") {
    if (request.periodType !== "duration") return false;
    return incomeStatementClassificationPassesTrustedValidation(request, requestedRow, action, semantics);
  }
  return false;
}

function analystSupportedSubtotalExclusion(classification: FinancialLineItemClassification) {
  if (!classification.is_subtotal || classification.confidence === "low") return false;
  return /subtotal|duplicate|double[-\s]?count|components? (?:are|is|mapped|reported|presented) separately|already (?:mapped|included|captured)/i.test(
    classification.reason || ""
  );
}

function trustedSourceSemantics(request: FinancialLineItemClassificationRequest) {
  const ownText = [request.reportedLineItemLabel, request.cleanLabel, request.xbrlTag ?? ""]
    .join(" ")
    .replace(/([a-z])([A-Z])/g, "$1 $2")
    .toLowerCase();
  const parentText = [request.parentSubtotal?.label ?? "", request.parentSubtotal?.concept ?? ""]
    .join(" ")
    .replace(/([a-z])([A-Z])/g, "$1 $2")
    .toLowerCase();
  const compact = normalizeKey(ownText);
  const balanceSheetSection = classifyBalanceSheetSourceSection(request.section, {
    label: request.cleanLabel || request.reportedLineItemLabel,
    tag: request.xbrlTag
  });
  const deferredTax = /\bdeferred\b.*\btax(?:es)?\b|\btax(?:es)?\b.*\bdeferred\b/.test(ownText) || /deferred(?:income)?tax/.test(compact);
  const deferredRevenue =
    /\bdeferred (?:income|revenue)\b|\bunearned revenue\b|\bcontract liabilit|\bcustomer advances?\b/.test(ownText) ||
    /contractwithcustomerliability|deferredrevenue|deferredincome|unearnedrevenue|customeradvances/.test(compact);
  const selfInsurance = /\bself[-\s]?insurance reserves?\b/.test(ownText) || /selfinsurancereserve/.test(compact);
  const pension = /\bpension\b|\bpostretirement\b|\bdefined benefit\b/.test(ownText) || /pension|postretirement|definedbenefit/.test(compact);
  const financeLease = /\bfinance lease\b|\bcapital lease\b|\blease financing\b/.test(ownText) || /financelease|capitallease|leasefinancing/.test(compact);
  const investmentSecurities =
    /\bshort[-\s]?term investments?\b|\bcurrent investments?\b|\bmarketable securities\b|\binvestment securities\b|\bdebt and equity securities\b|\bequity securities\b|\bavailable[-\s]?for[-\s]?sale securities\b/.test(
      ownText
    ) || /shortterminvest|marketablesecurit|investmentsecurit|equitysecurit|availableforsalesecurit/.test(compact);
  const currentLongTermDebt =
    /\bcurrent maturit|\bcurrent portion\b.*\blong[-\s]?term debt|\blong[-\s]?term debt\b.*\bcurrent\b|\bconvertible\b.*\bnotes?\b|\bsenior notes?\b/.test(
      ownText
    ) || /longtermdebtcurrent|currentportionoflongtermdebt|currentmaturitiesoflongtermdebt/.test(compact);
  const shortTermBorrowing =
    /\bshort[-\s]?term borrowings?\b|\bcommercial paper\b|\brevolver\b|\brevolving credit\b|\bline of credit\b|\bcurrent borrowings?\b|\bnotes? payable\b.*\bcurrent\b/.test(
      ownText
    ) || /shorttermborrowing|commercialpaper|revolvingcreditfacilitycurrent|currentborrowings|notespayablecurrent/.test(compact);
  const debtAdjustment = balanceSheetSourceLooksLikeDebtCarryingValueAdjustment({
    label: request.cleanLabel || request.reportedLineItemLabel,
    tag: request.xbrlTag
  });
  const debt =
    !selfInsurance &&
    !pension &&
    !deferredRevenue &&
    !investmentSecurities &&
    (currentLongTermDebt || shortTermBorrowing || debtAdjustment || /\bdebt\b|\bborrowings?\b|\bnotes?\b|\bterm loans?\b/.test(ownText));
  const acquiredInProcessResearchDevelopment =
    !/\bexclud(?:e|es|ing)\b.*\bacquired\b.*\bin[-\s]?process\b/.test(ownText) &&
    (/\bacquired\b.*\bin[-\s]?process\b.*\bresearch\b.*\bdevelopment\b|\bin[-\s]?process\b.*\bresearch\b.*\bdevelopment\b|\bipr&d\b|\biprd\b/.test(
      ownText
    ) || /acquiredinprocessresearchanddevelopment/.test(compact));
  const restaurantDirectCost =
    /\b(?:company[-\s]?(?:owned(?:\s+and\s+operated)?|operated)|franchised|other) restaurants?\b.*\b(?:costs?|expenses?)\b|\brestaurants?\b.*\b(?:occupancy|operating) expenses?\b/.test(
      `${ownText} ${parentText}`
    ) || /companyoperatedrestaurantexpense|franchisedrestaurantexpense|restaurantoperatingexpense/.test(compact);
  const directCost =
    /\bcost\b.*\b(?:sales|revenue|goods|products?|services?|operations?)\b|\b(?:sales|revenue|goods|products?|services?|operations?)\b.*\bcost\b|\bmerchandise costs?\b|\bfulfillment\b.*\b(?:costs?|expense)\b|\bfuel and purchased power\b|\bpurchased power\b.*\b(?:cost|expense)\b|\bproduction costs?\b/.test(
      ownText
    ) || /costofgoods|costofrevenue|costofsales|fuelandpurchasedpower/.test(compact) || restaurantDirectCost;
  const depreciationAmortizationMention =
    /\bdepreciation\b|\bamortization\b|\bdepletion\b|\bd&a\b/.test(ownText) || /depreciationandamortization/.test(compact);
  const depreciationAmortizationExcluded =
    /\b(?:exclud(?:e|es|ed|ing)|exclusive of|before)\b(?:\s+\w+){0,3}\s+\b(?:depreciation|amortization|depletion|d&a)\b/.test(ownText) ||
    /(?:excluding|exclusiveof|before)(?:depreciation|amortization|depletion)/.test(compact);
  const rawInterestIncome = /\binterest income\b|\binterest and dividend income\b/.test(ownText) || /interestincome/.test(compact);
  const rawInterestExpense = /\binterest expense\b|\binterest and debt expense\b|\bdebt expense\b/.test(ownText) || /interestexpense/.test(compact);
  const netInterestExpense =
    /\binterest expense\b.*\bnet of\b.*\binterest income\b/.test(ownText) ||
    /^(?:[^ ]*\s+)*interest expense,? net\b/.test(ownText) ||
    /interestexpensenonoperatingnet/.test(compact);
  const netInterestIncome =
    /\binterest income\b.*\bnet of\b.*\binterest expense\b/.test(ownText) ||
    /interestincomenonoperatingnet/.test(compact);
  const combinedNetInterest =
    /\binterest income\s*\(expense\),?\s*net\b|\binterest income\s*\/\s*expense\b|\bnet interest income\s*\(expense\)/.test(ownText) ||
    /interestincomeexpensenonoperatingnet/.test(compact) ||
    (rawInterestIncome && rawInterestExpense && !netInterestExpense && !netInterestIncome);

  return {
    ownText,
    compact,
    balanceSheetSection,
    deferredTax,
    deferredRevenue,
    selfInsurance,
    pension,
    financeLease,
    lease: /\blease liabilit|\blease obligations?\b/.test(ownText) || /leaseliabilit|leaseobligation/.test(compact),
    assetRetirementObligation: /\basset retirement obligations?\b/.test(ownText) || /assetretirementobligation/.test(compact),
    restrictedCash: /\brestricted cash\b/.test(ownText) || /restrictedcash/.test(compact),
    cashAggregate: /cashcashequivalentsrestrictedcash/.test(compact),
    cash:
      /\bcash(?: and| &)? cash equivalents?\b|\bcash and equivalents?\b|\bcash and due from banks?\b|\binterest[-\s]?bearing deposits? in banks?\b/.test(
        ownText
      ) || /cashandcashequivalents|cashandduefrombanks|interestbearingdepositsinbanks/.test(compact),
    investmentSecurities,
    receivable: /\baccounts? receivables?\b|\btrade receivables?\b/.test(ownText) || /accountsreceivable|tradeaccountsreceivable|receivablesnetcurrent/.test(compact),
    inventory:
      /\binventor(?:y|ies)\b|\bspare parts?\b|\bparts and supplies\b|\baircraft fuel\b|\braw materials?\b|\bwork[-\s]?in[-\s]?process\b|\bfinished goods?\b/.test(
        ownText
      ) || /inventory|partsandsupplies/.test(compact),
    propertyPlantEquipment:
      /\bproperty\b.*\bplant\b.*\bequipment\b|\bproperty and equipment\b|\bpp&e\b|\butility plant\b|\breal estate investment propert/.test(ownText) ||
      /propertyplantandequipment|propertyandequipmentnet|utilityplant|realestateinvestmentproperty/.test(compact),
    intangible: /\bintangible assets?\b|\btrademarks?\b|\bcustomer relationships?\b/.test(ownText) || /intangibleassets|trademarks/.test(compact),
    goodwill: /\bgoodwill\b/.test(ownText) || /goodwill/.test(compact),
    accountsPayable: /\baccounts? payable\b|\btrade payables?\b|\bvendor payables?\b|\bpharmacy costs? payable\b/.test(ownText) || /accountspayable|tradepayable/.test(compact),
    accrued:
      /\baccrued\b|\bcompensation\b|\bpayroll\b|\bwages payable\b|\bbenefits payable\b|\binterest payable\b|\brebates?\b|\breturns?\b|\bpromotions?\b/.test(
        ownText
      ) || /accruedliabilit|employeerelatedliabilit|interestpayable/.test(compact),
    currentTaxPayable: /\bincome taxes? payable\b|\btaxes payable\b|\baccrued income taxes\b/.test(ownText) || /incometaxespayable|accruedincometaxes/.test(compact),
    currentLongTermDebt,
    shortTermBorrowing,
    debtAdjustment,
    debt,
    commonCapital:
      !investmentSecurities &&
      (/\bcommon stock\b|\badditional paid[-\s]?in capital\b|\bpaid[-\s]?in capital\b|\bcapital in excess\b/.test(ownText) ||
        /commonstock|additionalpaidincapital/.test(compact)),
    retainedEarnings: /\bretained earnings\b|\bretained deficit\b|\baccumulated deficit\b/.test(ownText) || /retainedearningsaccumulateddeficit/.test(compact),
    treasuryStock: /\btreasury stock\b|\bcontra[-\s]?equity\b|\besop\b|\bemployee benefit trust\b/.test(ownText) || /treasurystock/.test(compact),
    accumulatedOtherComprehensiveIncome:
      /\baccumulated other comprehensive (?:income|loss)\b|\baoci\b/.test(ownText) || /accumulatedothercomprehensiveincomeloss/.test(compact),
    noncontrollingInterest: /\bnon[-\s]?controlling interests?\b|\bminority interest\b/.test(ownText) || /noncontrollinginterest|minorityinterest/.test(compact),
    redeemableNoncontrollingInterest:
      /\bredeemable\b.*\bnon[-\s]?controlling interests?\b|\bnon[-\s]?controlling interests?\b.*\bredeemable\b|\bredeemable nci\b/.test(ownText) ||
      /redeemablenoncontrollinginterest/.test(compact),
    acquiredInProcessResearchDevelopment,
    researchDevelopment:
      acquiredInProcessResearchDevelopment ||
      /\bresearch\b|\br&d\b|\bproduct development\b|\bengineering expense\b|\btechnology development\b|\btechnology and content\b/.test(ownText) ||
      /researchanddevelopment|productdevelopment|technologydevelopment/.test(compact),
    sellingGeneralAdministrative:
      /\badvertising\b|\bmarketing\b|\bpromotion(?:al)?\b|\bsales and marketing\b|\bselling and marketing\b|\bsales expense\b|\bselling expense\b|\bgeneral and administrative\b|\badministrative expense\b|\bcorporate overhead\b|\bsg&a\b/.test(
        ownText
      ) || /sellinggeneralandadministrative|advertisingexpense|salesandmarketing/.test(compact),
    depreciationAmortization: depreciationAmortizationMention && !depreciationAmortizationExcluded && !directCost,
    directCost,
    goodwillImpairment: /\bgoodwill\b.*\bimpairment\b|\bimpairment\b.*\bgoodwill\b/.test(ownText) || /goodwillimpairment/.test(compact),
    specialOperating:
      /\bspecial items?\b|\brestructuring\b|\bimpairment\b|\bspecial charges?\b|\bintegration costs?\b|\blitigation\b|\bsettlement\b|\baccretion\b/.test(
        ownText
      ),
    combinedInterestOther: /\binterest\b.*\bother\b|\bother\b.*\binterest\b/.test(ownText),
    combinedNetInterest,
    netInterestExpense,
    netInterestIncome,
    interestIncome: rawInterestIncome && !netInterestExpense && !combinedNetInterest,
    interestExpense: rawInterestExpense && !netInterestIncome && !combinedNetInterest,
    incomeTaxExpense: /\bincome tax(?:es)?\b|\bprovision for (?:income )?tax(?:es)?\b/.test(ownText) || /incometaxexpensebenefit|provisionforincometax/.test(compact)
  };
}

function balanceSheetClassificationPassesTrustedValidation(
  request: FinancialLineItemClassificationRequest,
  row: string,
  availableRow: string,
  action: FinancialLineItemClassification["recommended_action"],
  semantics: TrustedSourceSemantics
) {
  const definition = balanceSheetRowDefinitionForLabel(row) ?? balanceSheetRowDefinitionForLabel(availableRow);
  if (!definition || definition.family === "totals" || definition.kind === "subtotal" || definition.kind === "total") return false;
  const analystSupportedGenericCurrentDebt =
    semantics.balanceSheetSection === "current liabilities" &&
    semantics.debt &&
    !semantics.shortTermBorrowing &&
    (modelRowsMatch(row, "LT Debt (Incl. Current Portion)") || modelRowsMatch(row, "Total Debt"));
  if (
    !balanceSheetSectionCompatible(row, semantics.balanceSheetSection, {
      label: request.cleanLabel || request.reportedLineItemLabel,
      tag: request.xbrlTag
    }) &&
    !analystSupportedGenericCurrentDebt
  ) {
    return false;
  }

  const routeTarget = trustedBalanceSheetRouteTarget(request, semantics);
  if (!balanceSheetTargetAllowedForSection(row, definition.family, semantics, routeTarget)) return false;
  if (routeTarget && !modelRowsMatch(row, routeTarget)) return false;
  if (!balanceSheetTargetHasTrustedSemantics(request, row, definition.kind, semantics)) return false;
  if (action === "merge_into_other" && definition.kind !== "catch_all") return false;
  return true;
}

function balanceSheetTargetAllowedForSection(
  row: string,
  family: string,
  semantics: TrustedSourceSemantics,
  routeTarget: string | null
) {
  const section = semantics.balanceSheetSection;
  if (section === "current assets") return family === "current_assets";
  if (section === "non-current assets") return family === "non_current_assets";
  if (section === "current liabilities") {
    if (family === "current_liabilities") return true;
    return (
      semantics.debt &&
      !semantics.shortTermBorrowing &&
      (modelRowsMatch(row, "LT Debt (Incl. Current Portion)") || modelRowsMatch(row, "Total Debt"))
    );
  }
  if (section === "non-current liabilities") return family === "non_current_liabilities";
  if (section === "equity") {
    if (family === "equity") return true;
    return semantics.redeemableNoncontrollingInterest && modelRowsMatch(row, "Other Non-Current Liabilities");
  }
  return Boolean(routeTarget && modelRowsMatch(row, routeTarget));
}

function trustedBalanceSheetRouteTarget(request: FinancialLineItemClassificationRequest, semantics: TrustedSourceSemantics) {
  const preferred = (...rows: string[]) => rows.find((row) => modelRowAvailable(row, request.availableModelRows)) ?? null;
  const section = semantics.balanceSheetSection;
  if (semantics.redeemableNoncontrollingInterest) return preferred("Mezzanine Equity", "Other Non-Current Liabilities");
  if (semantics.selfInsurance) {
    return section === "non-current liabilities"
      ? preferred("Other Non-Current Liabilities")
      : preferred("Accrued Liabilities", "Other Current Liabilities");
  }
  if (semantics.debtAdjustment || semantics.currentLongTermDebt || semantics.financeLease) {
    return preferred("LT Debt (Incl. Current Portion)", "Current Portion of Long-Term Debt", "Total Debt");
  }
  if (semantics.shortTermBorrowing) return preferred("Revolver");
  if (semantics.deferredTax) {
    if (section === "current assets") return preferred("Prepaid & Other Current Assets");
    if (section === "non-current assets") {
      const dedicatedAssetRow = request.availableModelRows.find((row) => /deferred.*tax.*asset/i.test(row));
      return dedicatedAssetRow ?? preferred("Other Non-Current Assets");
    }
    if (section === "current liabilities") return preferred("Other Current Liabilities");
    return preferred("Deferred Income Taxes", "Other Non-Current Liabilities");
  }
  if (semantics.deferredRevenue) {
    return section === "non-current liabilities" ? preferred("Other Non-Current Liabilities") : preferred("Other Current Liabilities");
  }
  if (semantics.pension) {
    return section === "current liabilities"
      ? preferred("Other Current Liabilities")
      : preferred("Pension Liabilities", "Other Non-Current Liabilities");
  }
  if (semantics.lease) {
    return section === "non-current liabilities"
      ? preferred("Lease Liabilities", "Other Non-Current Liabilities")
      : preferred("Other Current Liabilities");
  }
  if (semantics.assetRetirementObligation) return preferred("Other Non-Current Liabilities");
  if (semantics.investmentSecurities) {
    if (section === "current assets") {
      const dedicatedInvestmentRow = request.availableModelRows.find(
        (row) =>
          /short[-\s]?term investments?|current investments?|marketable securities|investment securities/i.test(row) &&
          !/cash.*(?:investments?|securities)|(?:investments?|securities).*cash/i.test(row)
      );
      if (dedicatedInvestmentRow) return dedicatedInvestmentRow;
      const combinedCashRow = request.availableModelRows.find((row) => /cash.*(?:investments?|securities)|(?:investments?|securities).*cash/i.test(row));
      return combinedCashRow ?? preferred("Cash & Cash Equivalents", "Prepaid & Other Current Assets");
    }
    return preferred("Other Non-Current Assets");
  }
  if (semantics.restrictedCash && !semantics.cashAggregate) {
    return section === "non-current assets" ? preferred("Other Non-Current Assets") : preferred("Prepaid & Other Current Assets");
  }
  if (semantics.cash) return preferred("Cash & Cash Equivalents");
  if (semantics.receivable) return preferred("Accounts Receivable", "Prepaid & Other Current Assets");
  if (semantics.inventory) return preferred("Inventory", "Prepaid & Other Current Assets");
  if (semantics.propertyPlantEquipment) return preferred("PP&E, Net", "Other Non-Current Assets");
  if (semantics.intangible) return preferred("Intangible Assets, Net", "Other Non-Current Assets");
  if (semantics.goodwill && !semantics.goodwillImpairment) return preferred("Goodwill", "Other Non-Current Assets");
  if (semantics.accountsPayable) return preferred("Accounts Payable", "Accrued Liabilities", "Other Current Liabilities");
  if ((semantics.accrued || semantics.currentTaxPayable) && section === "current liabilities") {
    return preferred("Accrued Liabilities", "Other Current Liabilities");
  }
  if (semantics.debt) return preferred("LT Debt (Incl. Current Portion)", "Total Debt");
  if (semantics.treasuryStock) return preferred("Treasury Stock");
  if (semantics.retainedEarnings) return preferred("Retained Earnings");
  if (semantics.accumulatedOtherComprehensiveIncome) return preferred("AOCI");
  if (semantics.noncontrollingInterest) return preferred("Noncontrolling Interests");
  if (semantics.commonCapital) return preferred("Common Stock & APIC");
  return null;
}

function balanceSheetTargetHasTrustedSemantics(
  request: FinancialLineItemClassificationRequest,
  row: string,
  kind: string,
  semantics: TrustedSourceSemantics
) {
  if (modelRowsMatch(row, "Cash & Cash Equivalents")) return semantics.cash || semantics.investmentSecurities;
  if (modelRowsMatch(row, "Short-Term Investments")) return semantics.investmentSecurities;
  if (modelRowsMatch(row, "Accounts Receivable")) return semantics.receivable;
  if (modelRowsMatch(row, "Inventory")) return semantics.inventory;
  if (modelRowsMatch(row, "PP&E, Net")) return semantics.propertyPlantEquipment;
  if (modelRowsMatch(row, "Intangible Assets, Net")) return semantics.intangible;
  if (modelRowsMatch(row, "Goodwill")) return semantics.goodwill && !semantics.goodwillImpairment;
  if (modelRowsMatch(row, "Accounts Payable")) return semantics.accountsPayable;
  if (modelRowsMatch(row, "Accrued Liabilities")) return semantics.accrued || semantics.currentTaxPayable || semantics.selfInsurance;
  if (modelRowsMatch(row, "Revolver")) return semantics.shortTermBorrowing && !semantics.currentLongTermDebt;
  if (modelRowsMatch(row, "LT Debt (Incl. Current Portion)") || modelRowsMatch(row, "Total Debt")) {
    return semantics.debt || semantics.currentLongTermDebt || semantics.debtAdjustment || semantics.financeLease;
  }
  if (modelRowsMatch(row, "Current Portion of Long-Term Debt")) return semantics.currentLongTermDebt;
  if (modelRowsMatch(row, "Deferred Income Taxes")) return semantics.deferredTax && /liabilit/.test(semantics.balanceSheetSection);
  if (modelRowsMatch(row, "Lease Liabilities")) return semantics.lease;
  if (modelRowsMatch(row, "Pension Liabilities")) return semantics.pension;
  if (modelRowsMatch(row, "Common Stock & APIC")) return semantics.commonCapital;
  if (modelRowsMatch(row, "Retained Earnings")) return semantics.retainedEarnings;
  if (modelRowsMatch(row, "Treasury Stock")) return semantics.treasuryStock;
  if (modelRowsMatch(row, "AOCI")) return semantics.accumulatedOtherComprehensiveIncome;
  if (modelRowsMatch(row, "Noncontrolling Interests")) return semantics.noncontrollingInterest && !semantics.redeemableNoncontrollingInterest;
  if (modelRowsMatch(row, "Mezzanine Equity")) return semantics.redeemableNoncontrollingInterest;
  if (kind === "catch_all") return true;
  return sourceMatchesBalanceSheetRowDefinition(request, row);
}

function sourceMatchesBalanceSheetRowDefinition(request: FinancialLineItemClassificationRequest, row: string) {
  const definition = balanceSheetRowDefinitionForLabel(row);
  if (!definition) return false;
  const sourceTag = normalizeKey((request.xbrlTag ?? "").split(":").pop() ?? "");
  const sourceLabel = normalizeKey(request.cleanLabel || request.reportedLineItemLabel);
  return [...definition.tags, ...definition.aliases, ...(definition.sourceAliases ?? []), definition.canonical].some((candidate) => {
    const key = normalizeKey(candidate);
    return Boolean(key && (key === sourceTag || key === sourceLabel));
  });
}

function incomeStatementClassificationPassesTrustedValidation(
  request: FinancialLineItemClassificationRequest,
  row: string,
  action: FinancialLineItemClassification["recommended_action"],
  semantics: TrustedSourceSemantics
) {
  const canonicalRow = INCOME_STATEMENT_CLASSIFICATION_ROWS.find((candidate) => modelRowsMatch(row, candidate));
  if (!canonicalRow) return false;
  const allowedRows = incomeStatementRowsAllowedForSection(request.section);
  if (!allowedRows.some((candidate) => modelRowsMatch(canonicalRow, candidate))) return false;
  const routeTarget = trustedIncomeStatementRouteTarget(request, semantics);
  if (routeTarget && !modelRowsMatch(canonicalRow, routeTarget)) return false;
  if (!incomeStatementTargetHasTrustedSemantics(canonicalRow, request, semantics)) return false;
  if (
    action === "merge_into_other" &&
    !modelRowsMatch(canonicalRow, "Other Operating Income / Expense") &&
    !modelRowsMatch(canonicalRow, "Other Non-Operating Income / Expense")
  ) {
    return false;
  }
  return true;
}

function incomeStatementRowsAllowedForSection(section: FinancialStatementSection): readonly string[] {
  if (section === "revenue") return ["Revenue"];
  if (section === "operating expenses") {
    return ["COGS / Cost of Goods Sold", "SG&A", "R&D", "D&A", "Goodwill Impairment", "Other Operating Income / Expense"];
  }
  if (section === "below operating income") {
    return ["Interest Income", "Interest Expense", "Goodwill Impairment", "Other Non-Operating Income / Expense"];
  }
  if (section === "tax") return ["Income Tax Benefit / Expense"];
  return [];
}

function trustedIncomeStatementRouteTarget(request: FinancialLineItemClassificationRequest, semantics: TrustedSourceSemantics) {
  const preferred = (...rows: string[]) => rows.find((row) => modelRowAvailable(row, request.availableModelRows)) ?? null;
  if (request.section === "revenue") return preferred("Revenue");
  if (request.section === "tax" && semantics.incomeTaxExpense) return preferred("Income Tax Benefit / Expense");
  if (semantics.acquiredInProcessResearchDevelopment) return preferred("Other Operating Income / Expense", "R&D");
  if (semantics.researchDevelopment) return preferred("R&D");
  if (semantics.sellingGeneralAdministrative) return preferred("SG&A");
  if (semantics.directCost) return preferred("COGS / Cost of Goods Sold");
  if (semantics.depreciationAmortization) return preferred("D&A");
  if (semantics.goodwillImpairment) {
    return preferred(
      "Goodwill Impairment",
      request.section === "operating expenses" ? "Other Operating Income / Expense" : "Other Non-Operating Income / Expense"
    );
  }
  if (semantics.specialOperating) {
    return request.section === "operating expenses"
      ? preferred("Other Operating Income / Expense")
      : preferred("Other Non-Operating Income / Expense");
  }
  if (semantics.netInterestExpense) return preferred("Interest Expense");
  if (semantics.netInterestIncome) return preferred("Interest Income");
  if (semantics.combinedInterestOther || semantics.combinedNetInterest) return preferred("Other Non-Operating Income / Expense");
  if (semantics.interestIncome) return preferred("Interest Income");
  if (semantics.interestExpense) return preferred("Interest Expense");
  return null;
}

function incomeStatementTargetHasTrustedSemantics(
  row: string,
  request: FinancialLineItemClassificationRequest,
  semantics: TrustedSourceSemantics
) {
  if (modelRowsMatch(row, "Revenue")) return request.section === "revenue";
  if (modelRowsMatch(row, "COGS / Cost of Goods Sold")) return semantics.directCost;
  if (modelRowsMatch(row, "SG&A")) return semantics.sellingGeneralAdministrative;
  if (modelRowsMatch(row, "R&D")) return semantics.researchDevelopment;
  if (modelRowsMatch(row, "D&A")) return semantics.depreciationAmortization;
  if (modelRowsMatch(row, "Goodwill Impairment")) return semantics.goodwillImpairment;
  if (modelRowsMatch(row, "Interest Income")) return (semantics.interestIncome || semantics.netInterestIncome) && !semantics.combinedInterestOther;
  if (modelRowsMatch(row, "Interest Expense")) return (semantics.interestExpense || semantics.netInterestExpense) && !semantics.combinedInterestOther;
  if (modelRowsMatch(row, "Income Tax Benefit / Expense")) return request.section === "tax" && semantics.incomeTaxExpense;
  if (modelRowsMatch(row, "Other Operating Income / Expense")) return request.section === "operating expenses";
  if (modelRowsMatch(row, "Other Non-Operating Income / Expense")) return request.section === "below operating income";
  return false;
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
    deadlineAt: options.deadlineAt,
    signal: options.signal,
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
  options: LlmClassificationOptions,
  maxTotalAttempts: number
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
    deadlineAt: options.deadlineAt,
    signal: options.signal,
    maxTokens: Math.max(3_000, Math.min(12_000, 1_200 + targets.length * 400)),
    maxAttemptsPerModel: /(?:deepseek-v4-flash|^openrouter\/free|:free$)/i.test(options.model.trim()) ? 2 : 1,
    maxTotalAttempts,
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
    is_subtotal:
      decision.recommended_action === "exclude"
        ? /subtotal|duplicate|double[-\s]?count|components? (?:are|is|mapped|reported|presented) separately|already (?:mapped|included|captured)/i.test(
            decision.reason
          )
        : fallback.is_subtotal,
    should_exclude_from_other_bucket: !isReusableOtherBucketModelRow(requestedRow),
    confidence: decision.confidence,
    reason: decision.reason,
    requires_validation: true,
    requires_revalidation: true,
    llm_used: false,
    mapping_passed_validation: false
  };
  classification.mapping_passed_validation = classificationPassesValidation(target.request, classification);
  return classification;
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
  const validated = classificationPassesValidation(request, {
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
  const rejected: string[] = [];
  for (const item of rawClassifications) {
    if (!isRecord(item)) {
      return { ok: false, needsHumanReview: true, error: "Every statement classification must be a JSON object." };
    }
    const sourceRowKey = statementDecisionSourceRowKey(item);
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
    if (!recommendedAction) {
      seen.add(sourceRowKey);
      if (target.deterministicIsValidated) {
        classifications.push(statementDeterministicFallbackDecision(target, "LLM recommended_action was invalid"));
      } else {
        rejected.push(`${sourceRowKey}: recommended_action is invalid.`);
      }
      continue;
    }
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
      seen.add(sourceRowKey);
      if (target.deterministicIsValidated) {
        classifications.push(statementDeterministicFallbackDecision(target, "LLM omitted the required recommended_model_row"));
      } else {
        rejected.push(`${sourceRowKey}: recommended_model_row must be a string.`);
      }
      continue;
    }
    let normalized: StatementLlmClassificationDecision = {
      source_row_key: sourceRowKey,
      recommended_action: recommendedAction,
      recommended_model_row: normalizeModelRow(rawRecommendedRow ?? "") || rawRecommendedRow || target.fallback.recommended_model_row,
      confidence,
      reason: rawReason || "LLM whole-statement accounting classification.",
      decision_origin: "validated_llm"
    };
    const expanded = statementDecisionToClassification(target, normalized);
    const validated = expanded.mapping_passed_validation;
    if (!validated || normalized.confidence === "low") {
      if (target.deterministicIsValidated) {
        normalized = statementDeterministicFallbackDecision(
          target,
          !validated
            ? `LLM recommendation ${normalized.recommended_model_row || normalized.recommended_action} was rejected by accounting validation`
            : "LLM returned low confidence"
        );
      } else {
        rejected.push(
          !validated
            ? `${target.request.cleanLabel || target.request.reportedLineItemLabel}: LLM decision failed deterministic accounting validation for ${normalized.recommended_model_row}.`
            : `${target.request.cleanLabel || target.request.reportedLineItemLabel}: LLM returned low confidence without a validated fallback.`
        );
        seen.add(sourceRowKey);
        continue;
      }
    }
    classifications.push(normalized);
    seen.add(sourceRowKey);
  }
  const missing = targets.filter((target) => !seen.has(target.rowKey));
  if (!classifications.length) {
    return {
      ok: false,
      needsHumanReview: true,
      error:
        rejected[0] ??
        `Statement classifier omitted target row(s): ${missing.map((target) => target.rowKey).join(", ")}.`
    };
  }
  return {
    ok: true,
    value: { classifications },
    validated: true,
    affectedOutput: classifications.some(statementDecisionWasAcceptedFromLlm)
  };
}

function statementDeterministicFallbackDecision(
  target: PreparedLineItemClassification,
  rejectionReason: string
): StatementLlmClassificationDecision {
  return {
    source_row_key: target.rowKey,
    recommended_action: target.initialClassification.recommended_action,
    recommended_model_row: target.initialClassification.recommended_model_row,
    confidence: "medium",
    reason: `${rejectionReason}; retained the independently validated mapping ${target.initialClassification.recommended_model_row}.`,
    decision_origin: "deterministic_fallback"
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

function statementDecisionSourceRowKey(item: Record<string, any>) {
  return firstString(item, ["source_row_key"]);
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
