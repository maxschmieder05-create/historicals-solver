export type LlmWorkbenchFact = {
  period: string;
  sourcePeriod?: string;
  periodType: "duration" | "instant";
  concept: string;
  label: string;
  value: number;
  valueMillions: number;
  unit?: string;
  taxonomy?: string;
  sourceLayer?: string;
  form?: string;
  filed?: string;
  accession?: string;
  sourceUrl?: string;
  statement?: string;
  section?: string;
};

export type LlmWorkbenchStatementRow = {
  period: string;
  statementName: string;
  sourceTableType: string;
  rowLabel: string;
  xbrlConcept?: string;
  value: number | string | null;
  valueMillions?: number;
  unit?: string;
  accession: string;
  form?: string;
  filingDate?: string;
  rowOrder?: number;
  section?: string;
  currentNonCurrentSection?: string;
  dimensionMembers?: string[];
  sourceUrl?: string;
};

export type LlmWorkbenchCell = {
  sheetName: string;
  cell: string;
  row?: number;
  column?: number;
  period: string;
  modelRowLabel: string;
  section?: string;
  value: number | string | null;
  formula?: string;
  isFormula: boolean;
  mappingType?: string;
  validationStatus?: string;
  sourceLedgerStatus?: string;
  sourceLineItemLabel?: string;
  sourceXbrlTag?: string;
  classificationReason?: string;
};

export type LlmWorkbenchAuditRow = {
  period: string;
  cell: string;
  modelRowLabel: string;
  valueWritten: number;
  mappingType: string;
  validationStatus: string;
  conceptsUsed: string;
  secLabels?: string;
  sourceStatement?: string;
  confidence?: string;
  notes?: string;
};

export type LlmWorkbenchSourceLedgerRow = {
  period: string;
  cell: string;
  modelRow: number;
  modelRowLabel?: string;
  value: number | string | null;
  sourceLineItemLabel: string;
  sourceXbrlTag: string;
  mappingStatus: string;
  sourceStatement: string;
  sourceTableType: string;
  classificationReason: string;
};

export type LlmWorkbenchBalanceAssignment = {
  period: string;
  sourceLineItemLabel: string;
  sourceXbrlTag: string;
  amount: number;
  amountMillions: number;
  assignedModelRow: string;
  assignmentStatus: string;
  classificationReason: string;
  validationStatus: string;
  side: string;
  sourceSection: string;
};

export type LlmWorkbenchIncomeAssignment = {
  period: string;
  sourceLineItemLabel: string;
  sourceXbrlTag: string;
  sourceAmount: number;
  sourceAmountMillions: number;
  modelAmount: number;
  modelAmountMillions: number;
  assignedModelRow: string;
  assignmentStatus: string;
  classificationReason: string;
  validationStatus: string;
  sourceSection: string;
};

export type LlmWorkbookToolboxInput = {
  company: {
    ticker: string;
    name: string;
    cik: string;
  };
  periods: string[];
  facts: LlmWorkbenchFact[];
  statementRows: LlmWorkbenchStatementRow[];
  workbookCells: LlmWorkbenchCell[];
  auditRows: LlmWorkbenchAuditRow[];
  sourceLedgerRows: LlmWorkbenchSourceLedgerRow[];
  balanceSheetAssignments: LlmWorkbenchBalanceAssignment[];
  incomeStatementAssignments?: LlmWorkbenchIncomeAssignment[];
  validationFailures: string[];
  warnings: string[];
  limits?: Partial<LlmWorkbookToolboxLimits>;
};

export type LlmWorkbookToolboxLimits = {
  maxFacts: number;
  maxStatementRows: number;
  maxWorkbookCells: number;
  maxAuditRows: number;
  maxSourceLedgerRows: number;
  maxBalanceSheetAssignments: number;
  maxIncomeStatementAssignments: number;
  maxWarnings: number;
};

export type LlmWorkbookToolbox = {
  protocol: "historicals-solver.llm-workbench.v1";
  company: LlmWorkbookToolboxInput["company"];
  periods: string[];
  verificationGate: {
    status: "passed" | "failed";
    failClosed: boolean;
    blockingFailures: string[];
    warnings: string[];
    rule: string;
  };
  toolOutputs: Array<{
    name: string;
    description: string;
    output: unknown;
  }>;
  omittedCounts: {
    facts: number;
    statementRows: number;
    workbookCells: number;
    auditRows: number;
    sourceLedgerRows: number;
    balanceSheetAssignments: number;
    incomeStatementAssignments: number;
    warnings: number;
  };
};

const DEFAULT_LIMITS: LlmWorkbookToolboxLimits = {
  maxFacts: 500,
  maxStatementRows: 650,
  maxWorkbookCells: 700,
  maxAuditRows: 500,
  maxSourceLedgerRows: 600,
  maxBalanceSheetAssignments: 500,
  maxIncomeStatementAssignments: 500,
  maxWarnings: 80
};

export function buildLlmWorkbookToolbox(input: LlmWorkbookToolboxInput): LlmWorkbookToolbox {
  const limits = { ...DEFAULT_LIMITS, ...(input.limits ?? {}) };
  const facts = rankFacts(input.facts).slice(0, limits.maxFacts);
  const statementRows = rankStatementRows(input.statementRows).slice(0, limits.maxStatementRows);
  const workbookCells = rankWorkbookCells(input.workbookCells).slice(0, limits.maxWorkbookCells);
  const auditRows = rankAuditRows(input.auditRows).slice(0, limits.maxAuditRows);
  const sourceLedgerRows = rankSourceLedgerRows(input.sourceLedgerRows).slice(0, limits.maxSourceLedgerRows);
  const balanceSheetAssignments = rankBalanceAssignments(input.balanceSheetAssignments).slice(0, limits.maxBalanceSheetAssignments);
  const incomeStatementAssignments = rankIncomeAssignments(input.incomeStatementAssignments ?? []).slice(0, limits.maxIncomeStatementAssignments);
  const warnings = input.warnings.slice(0, limits.maxWarnings);
  const blockingFailures = uniqueStrings(input.validationFailures).filter(Boolean);

  return {
    protocol: "historicals-solver.llm-workbench.v1",
    company: input.company,
    periods: input.periods,
    verificationGate: {
      status: blockingFailures.length ? "failed" : "passed",
      failClosed: true,
      blockingFailures,
      warnings,
      rule:
        "Do not approve or return a workbook when blockingFailures is non-empty, when source-ledger status is stale_or_unsupported for a nonblank value, or when a material SEC primary-statement line is unmapped without an explicit reusable accounting reason."
    },
    toolOutputs: [
      {
        name: "sec.get_company_facts",
        description: "EDGAR companyfacts values normalized to model periods. Values are SEC-sourced or EDGAR-derived only.",
        output: { facts }
      },
      {
        name: "sec.get_financial_statement_rows",
        description: "Parsed SEC filing statement rows, including primary statements and segment tables when available.",
        output: { rows: statementRows }
      },
      {
        name: "workbook.inspect_output_cells",
        description: "Final workbook cells after fill logic, including formulas, values, period labels, mapping status, and source trace hints.",
        output: { cells: workbookCells }
      },
      {
        name: "workbook.trace_source_ledger",
        description: "Source-backed ledger for output cells. Nonblank hardcoded values must have explicit current-company SEC support, validated derivation support, explicit-zero support, or preserved formula support.",
        output: { rows: sourceLedgerRows }
      },
      {
        name: "workbook.get_mapping_audit",
        description: "Mapping and write decisions, including concepts used, validation status, confidence, and classification notes.",
        output: { rows: auditRows }
      },
      {
        name: "workbook.get_balance_sheet_assignment_ledger",
        description: "Primary balance-sheet source line items and how each was assigned, grouped, excluded, or flagged.",
        output: { rows: balanceSheetAssignments }
      },
      {
        name: "workbook.get_income_statement_assignment_ledger",
        description: "Primary income-statement source line items and how each was assigned, grouped, excluded, or flagged.",
        output: { rows: incomeStatementAssignments }
      },
      {
        name: "workbook.validate_return",
        description: "Deterministic fail-closed validation gate for returned workbooks.",
        output: {
          status: blockingFailures.length ? "failed" : "passed",
          blockingFailures,
          warnings
        }
      }
    ],
    omittedCounts: {
      facts: Math.max(0, input.facts.length - facts.length),
      statementRows: Math.max(0, input.statementRows.length - statementRows.length),
      workbookCells: Math.max(0, input.workbookCells.length - workbookCells.length),
      auditRows: Math.max(0, input.auditRows.length - auditRows.length),
      sourceLedgerRows: Math.max(0, input.sourceLedgerRows.length - sourceLedgerRows.length),
      balanceSheetAssignments: Math.max(0, input.balanceSheetAssignments.length - balanceSheetAssignments.length),
      incomeStatementAssignments: Math.max(0, (input.incomeStatementAssignments ?? []).length - incomeStatementAssignments.length),
      warnings: Math.max(0, input.warnings.length - warnings.length)
    }
  };
}

export function llmWorkbookToolboxSystemInstruction() {
  return [
    "You have access to materialized MCP-style tool outputs under payload.llmWorkbookToolbox.",
    "Treat each tool output as authoritative for its domain: SEC tools for EDGAR data, workbook tools for output cells and traces, and workbook.validate_return for deterministic return validation.",
    "Use these tool outputs to verify workbook behavior, grouping logic, source provenance, formula preservation, and line-item assignment. Never invent missing SEC facts or fill gaps with estimates.",
    "A workbook cannot be approved if workbook.validate_return reports failed, if nonblank workbook cells lack source-ledger support, or if a material primary SEC line item is unmapped without an explicit reusable accounting reason."
  ].join(" ");
}

function rankFacts(facts: LlmWorkbenchFact[]) {
  return facts
    .slice()
    .sort((a, b) => periodSort(a.period, b.period) || materiality(b.value) - materiality(a.value) || a.concept.localeCompare(b.concept));
}

function rankStatementRows(rows: LlmWorkbenchStatementRow[]) {
  return rows.slice().sort((a, b) => {
    const sourceRank = sourceTableRank(a.sourceTableType) - sourceTableRank(b.sourceTableType);
    return sourceRank || periodSort(a.period, b.period) || (a.rowOrder ?? 0) - (b.rowOrder ?? 0) || materialityValue(b.value) - materialityValue(a.value);
  });
}

function rankIncomeAssignments(rows: LlmWorkbenchIncomeAssignment[]) {
  return rows.slice().sort((a, b) => {
    const statusRank = assignmentStatusRank(a.assignmentStatus) - assignmentStatusRank(b.assignmentStatus);
    return statusRank || periodSort(a.period, b.period) || Math.abs(b.modelAmount) - Math.abs(a.modelAmount);
  });
}

function rankWorkbookCells(cells: LlmWorkbenchCell[]) {
  return cells.slice().sort((a, b) => {
    const statusRank = cellStatusRank(a) - cellStatusRank(b);
    return statusRank || periodSort(a.period, b.period) || a.sheetName.localeCompare(b.sheetName) || (a.row ?? 0) - (b.row ?? 0) || (a.column ?? 0) - (b.column ?? 0);
  });
}

function rankAuditRows(rows: LlmWorkbenchAuditRow[]) {
  return rows.slice().sort((a, b) => {
    const statusRank = auditStatusRank(a) - auditStatusRank(b);
    return statusRank || periodSort(a.period, b.period) || a.cell.localeCompare(b.cell);
  });
}

function rankSourceLedgerRows(rows: LlmWorkbenchSourceLedgerRow[]) {
  return rows.slice().sort((a, b) => {
    const statusRank = sourceLedgerStatusRank(a.mappingStatus) - sourceLedgerStatusRank(b.mappingStatus);
    return statusRank || periodSort(a.period, b.period) || a.cell.localeCompare(b.cell);
  });
}

function rankBalanceAssignments(rows: LlmWorkbenchBalanceAssignment[]) {
  return rows.slice().sort((a, b) => {
    const statusRank = assignmentStatusRank(a.assignmentStatus) - assignmentStatusRank(b.assignmentStatus);
    return statusRank || periodSort(a.period, b.period) || materiality(b.amount) - materiality(a.amount);
  });
}

function sourceTableRank(value: string) {
  if (value === "primary_statement") return 0;
  if (value === "segment_table") return 1;
  if (value === "cash_flow_reconciliation") return 2;
  if (value === "roll_forward") return 3;
  return 4;
}

function cellStatusRank(cell: LlmWorkbenchCell) {
  const text = `${cell.validationStatus ?? ""} ${cell.sourceLedgerStatus ?? ""} ${cell.mappingType ?? ""}`;
  if (/failed|blocked|stale_or_unsupported|unsupported|mismatch/i.test(text)) return 0;
  if (/warning|needs_review|skipped|cleared/i.test(text)) return 1;
  if (cell.isFormula) return 2;
  return 3;
}

function auditStatusRank(row: LlmWorkbenchAuditRow) {
  const text = `${row.validationStatus} ${row.mappingType} ${row.notes ?? ""}`;
  if (/failed|blocked|mismatch|stale_or_unsupported|unsupported/i.test(text)) return 0;
  if (/warning|needs_review|skipped|cleared/i.test(text)) return 1;
  return 2;
}

function sourceLedgerStatusRank(status: string) {
  if (status === "stale_or_unsupported") return 0;
  if (status === "explicit_zero_no_source_disclosed") return 1;
  if (status === "validated_current_company_derived_value") return 2;
  if (status === "formula_preserved") return 3;
  if (status === "explicit_current_sec_source") return 4;
  return 5;
}

function assignmentStatusRank(status: string) {
  if (/missing|unassigned|not assigned|unsupported/i.test(status)) return 0;
  if (/excluded/i.test(status)) return 1;
  if (/grouped/i.test(status)) return 2;
  return 3;
}

function periodSort(a: string, b: string) {
  return periodComparable(a).localeCompare(periodComparable(b));
}

function periodComparable(period: string) {
  const match = period.match(/^(?:(FY)|([1-4])Q)(\d{2}|\d{4})$/i);
  if (!match) return period;
  const year = match[3].length === 2 ? `20${match[3]}` : match[3];
  const quarter = match[1] ? "4.5" : match[2];
  return `${year}-${quarter}`;
}

function materiality(value: number) {
  return Math.abs(value);
}

function materialityValue(value: number | string | null) {
  return typeof value === "number" && Number.isFinite(value) ? Math.abs(value) : 0;
}

function uniqueStrings(values: string[]) {
  return Array.from(new Set(values.map((value) => value.trim()).filter(Boolean)));
}
