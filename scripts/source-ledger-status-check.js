const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const ts = require("typescript");

const repoRoot = path.resolve(__dirname, "..");

require.extensions[".ts"] = function compileTypeScriptModule(module, filename) {
  const source = fs.readFileSync(filename, "utf8");
  const compiled = ts.transpileModule(source, {
    compilerOptions: {
      module: ts.ModuleKind.CommonJS,
      target: ts.ScriptTarget.ES2020,
      esModuleInterop: true
    }
  }).outputText;
  module._compile(compiled, filename);
};

const { sourceLedgerStatusForAuditRow } = require(path.join(repoRoot, "server", "fill-model", "fill-model-service.ts"));

function auditRow(overrides = {}) {
  return {
    sheetName: "Model",
    cell: "F42",
    modelRowLabel: "Other Current Liabilities",
    period: "1Q23",
    valueWritten: 0,
    mappingType: "direct",
    conceptsUsed: "NoCurrentSecSource=0mm",
    sourceStatement: "balance",
    accession: "",
    sourceUrl: "",
    cellWritable: true,
    formulaPreserved: false,
    formulaStatus: "reported-period value explicitly sourced as zero",
    writeBlockedReason: "",
    signConvention: "explicit zero",
    confidence: "high",
    validationStatus: "OK!",
    notes: "Explicitly set to zero because the current SEC filing has no source for this row after prior filings reported a balance.",
    ...overrides
  };
}

assert.equal(
  sourceLedgerStatusForAuditRow(
    auditRow({
      formulaStatus: "unsupported reported-period value explicitly zeroed"
    })
  ),
  "explicit_zero_no_source_disclosed",
  "explicit no-source zeros should not be downgraded to stale just because older audit wording said unsupported"
);

assert.equal(
  sourceLedgerStatusForAuditRow(
    auditRow({
      modelRowLabel: "Deferred Income Taxes",
      conceptsUsed: "NoCurrentSecSource:DeferredIncomeTaxLiabilitiesNet=0mm",
      secLabels: "No current SEC source disclosed",
      validationStatus: "warning: Value is a zero/derived model support value with no direct SEC fact.",
      notes: "No separate deferred tax liability was reported in the SEC balance sheet for this period."
    })
  ),
  "explicit_zero_no_source_disclosed",
  "balance-sheet resolver zeros from model-layer no-source support should count as explicit sourced zeros"
);

assert.equal(
  sourceLedgerStatusForAuditRow(
    auditRow({
      modelRowLabel: "Deferred Income Taxes",
      conceptsUsed: "DeferredIncomeTaxLiabilitiesNet=0mm",
      formulaStatus: "not a formula cell",
      validationStatus: "warning: Value is a zero/derived model support value with no direct SEC fact.",
      notes: "No separate deferred tax liability was reported in the SEC balance sheet for this period."
    })
  ),
  "explicit_zero_no_source_disclosed",
  "older model-zero audit rows should also be recognized when the no-separate-row explanation is present"
);

assert.equal(
  sourceLedgerStatusForAuditRow(
    auditRow({
      mappingType: "cleared",
      conceptsUsed: "",
      formulaStatus: "unsupported hardcoded historical value cleared to zero",
      validationStatus: "cleared",
      notes: "",
      writeBlockedReason: "Prior hardcoded value 5.9 had no explicit SEC source for this filing period."
    })
  ),
  "stale_or_unsupported",
  "generic stale hardcode clears still need source support and must not masquerade as explicit zeros"
);

assert.equal(
  sourceLedgerStatusForAuditRow(
    auditRow({
      valueWritten: 12.3,
      mappingType: "residual",
      conceptsUsed: "LiabilitiesCurrent=100mm; AccountsPayableCurrent=20mm; AccruedLiabilitiesCurrent=67.7mm",
      formulaStatus: "historical input refreshed from primary SEC balance-sheet assignment",
      notes: "Residual calculated from SEC current liabilities less SEC-sourced components."
    })
  ),
  "validated_current_company_derived_value"
);

assert.equal(
  sourceLedgerStatusForAuditRow(
    auditRow({
      valueWritten: 12.3,
      conceptsUsed: "OtherLiabilitiesCurrent=12.3mm",
      accession: "0000000000-23-000001",
      formulaStatus: "not a formula cell",
      notes: "Mapped directly to SEC-sourced balance-sheet line item support."
    })
  ),
  "explicit_current_sec_source"
);

console.log("Source-ledger status regression passed.");
