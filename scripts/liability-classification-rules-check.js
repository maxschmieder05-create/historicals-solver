const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const Module = require("node:module");
const ts = require("typescript");

const repoRoot = path.resolve(__dirname, "..");
const sourcePath = path.join(repoRoot, "server", "fill-model", "liability-classification.ts");

function loadTypeScriptModule(file) {
  const source = fs.readFileSync(file, "utf8");
  const compiled = ts.transpileModule(source, {
    compilerOptions: {
      module: ts.ModuleKind.CommonJS,
      target: ts.ScriptTarget.ES2020,
      esModuleInterop: true
    }
  }).outputText;
  const mod = new Module(file, module);
  mod.filename = file;
  mod.paths = Module._nodeModulePaths(path.dirname(file));
  mod._compile(compiled, file);
  return mod.exports;
}

const {
  buildLiabilityTemplateMappingContext,
  currentDebtBelongsInAccruedLiabilities,
  otherNonCurrentLiabilityResidualExclusions
} = loadTypeScriptModule(sourcePath);

function accruedLiabilityResidual({ currentLiabilities, accountsPayable, otherCurrentLiabilities, currentDebt }, context) {
  return (
    currentLiabilities -
    accountsPayable -
    otherCurrentLiabilities -
    (currentDebtBelongsInAccruedLiabilities(context) ? 0 : currentDebt)
  );
}

function otherNonCurrentLiabilityResidual(values, context) {
  const exclusions = otherNonCurrentLiabilityResidualExclusions(context);
  return (
    values.totalLiabilities -
    (exclusions.alwaysExcludeCurrentLiabilities ? values.currentLiabilities : 0) -
    (exclusions.alwaysExcludeNonCurrentDebt ? values.nonCurrentDebt : 0) -
    (exclusions.excludeDeferredTaxLiabilities ? values.deferredTaxLiabilities : 0) -
    (exclusions.excludeLeaseLiabilities ? values.nonCurrentLeaseLiabilities : 0) -
    (exclusions.excludePensionLiabilities ? values.pensionLiabilities : 0)
  );
}

const baseRows = [
  { statement: "balance", label: "Accounts Payable" },
  { statement: "balance", label: "Accrued Expenses and Other" },
  { statement: "balance", label: "LT Debt (Incl. Current Portion)" },
  { statement: "balance", label: "Other Non-Current Liabilities" }
];

const embeddedDebtContext = buildLiabilityTemplateMappingContext(baseRows);
assert.equal(embeddedDebtContext.hasCurrentDebtRow, false);
assert.equal(embeddedDebtContext.hasCurrentLiabilitiesExcludingDebtRow, false);
assert.equal(embeddedDebtContext.hasDebtInclCurrentPortionRow, true);
assert.equal(currentDebtBelongsInAccruedLiabilities(embeddedDebtContext), true);
assert.equal(
  accruedLiabilityResidual(
    { currentLiabilities: 1000, accountsPayable: 250, otherCurrentLiabilities: 125, currentDebt: 75 },
    embeddedDebtContext
  ),
  625
);

const currentLiabilitiesExDebtContext = buildLiabilityTemplateMappingContext([
  ...baseRows,
  { statement: "balance", label: "Other Current Liabilities" },
  { statement: "balance", label: "Total Current Liabilities (Excl. Debt)" }
]);
assert.equal(currentLiabilitiesExDebtContext.hasCurrentLiabilitiesExcludingDebtRow, true);
assert.equal(currentLiabilitiesExDebtContext.hasOtherCurrentLiabilityRow, true);
assert.equal(currentDebtBelongsInAccruedLiabilities(currentLiabilitiesExDebtContext), false);
assert.equal(
  accruedLiabilityResidual(
    { currentLiabilities: 1000, accountsPayable: 250, otherCurrentLiabilities: 125, currentDebt: 75 },
    currentLiabilitiesExDebtContext
  ),
  550
);

const currentLiabilitiesLessDebtContext = buildLiabilityTemplateMappingContext([
  ...baseRows,
  { statement: "balance", label: "Total Current Liabilities Less Debt" }
]);
assert.equal(currentLiabilitiesLessDebtContext.hasCurrentLiabilitiesExcludingDebtRow, true);
assert.equal(currentDebtBelongsInAccruedLiabilities(currentLiabilitiesLessDebtContext), false);

const currentLiabilitiesExDebtWithoutRecognizedDebtRowContext = buildLiabilityTemplateMappingContext([
  { statement: "balance", label: "Accounts Payable" },
  { statement: "balance", label: "Accrued Expenses and Other" },
  { statement: "balance", label: "Total Current Liabilities (Excl. Debt)" },
  { statement: "balance", label: "Other Non-Current Liabilities" }
]);
assert.equal(currentLiabilitiesExDebtWithoutRecognizedDebtRowContext.hasCurrentLiabilitiesExcludingDebtRow, true);
assert.equal(currentDebtBelongsInAccruedLiabilities(currentLiabilitiesExDebtWithoutRecognizedDebtRowContext), false);
assert.equal(
  accruedLiabilityResidual(
    { currentLiabilities: 1000, accountsPayable: 250, otherCurrentLiabilities: 125, currentDebt: 75 },
    currentLiabilitiesExDebtWithoutRecognizedDebtRowContext
  ),
  550
);

const noDebtLineContext = buildLiabilityTemplateMappingContext([
  { statement: "balance", label: "Accounts Payable" },
  { statement: "balance", label: "Accrued Expenses and Other" },
  { statement: "balance", label: "Other Non-Current Liabilities" }
]);
assert.equal(noDebtLineContext.hasCurrentDebtRow, false);
assert.equal(noDebtLineContext.hasDebtInclCurrentPortionRow, false);
assert.equal(currentDebtBelongsInAccruedLiabilities(noDebtLineContext), true);
assert.equal(
  accruedLiabilityResidual(
    { currentLiabilities: 1000, accountsPayable: 250, otherCurrentLiabilities: 125, currentDebt: 75 },
    noDebtLineContext
  ),
  625
);

const explicitCurrentDebtContext = buildLiabilityTemplateMappingContext([
  ...baseRows,
  { statement: "balance", label: "Current Portion of Long-Term Debt" }
]);
assert.equal(explicitCurrentDebtContext.hasCurrentDebtRow, true);
assert.equal(currentDebtBelongsInAccruedLiabilities(explicitCurrentDebtContext), false);
assert.equal(
  accruedLiabilityResidual(
    { currentLiabilities: 1000, accountsPayable: 250, otherCurrentLiabilities: 125, currentDebt: 75 },
    explicitCurrentDebtContext
  ),
  550
);

const notesPayableCurrentContext = buildLiabilityTemplateMappingContext([
  ...baseRows,
  { statement: "balance", label: "Notes Payable, Current" }
]);
assert.equal(notesPayableCurrentContext.hasShortTermBorrowingsRow, true);
assert.equal(notesPayableCurrentContext.hasCurrentDebtMaturitiesRow, false);

const explicitLeaseAndPensionContext = buildLiabilityTemplateMappingContext([
  ...baseRows,
  { statement: "balance", label: "Deferred Income Taxes" },
  { statement: "balance", label: "Operating Lease Liabilities" },
  { statement: "balance", label: "Pension and Other Postretirement Liabilities" }
]);
assert.equal(explicitLeaseAndPensionContext.hasDeferredTaxLiabilityRow, true);
assert.deepEqual(otherNonCurrentLiabilityResidualExclusions(explicitLeaseAndPensionContext), {
  alwaysExcludeCurrentLiabilities: true,
  alwaysExcludeNonCurrentDebt: true,
  excludeDeferredTaxLiabilities: true,
  excludeLeaseLiabilities: true,
  excludePensionLiabilities: true
});
assert.equal(
  otherNonCurrentLiabilityResidual(
    {
      totalLiabilities: 5000,
      currentLiabilities: 1200,
      nonCurrentDebt: 1400,
      deferredTaxLiabilities: 150,
      nonCurrentLeaseLiabilities: 300,
      pensionLiabilities: 90
    },
    explicitLeaseAndPensionContext
  ),
  1860
);

const groupedLeaseContext = buildLiabilityTemplateMappingContext(baseRows);
assert.equal(groupedLeaseContext.hasDeferredTaxLiabilityRow, false);
assert.deepEqual(otherNonCurrentLiabilityResidualExclusions(groupedLeaseContext), {
  alwaysExcludeCurrentLiabilities: true,
  alwaysExcludeNonCurrentDebt: true,
  excludeDeferredTaxLiabilities: false,
  excludeLeaseLiabilities: false,
  excludePensionLiabilities: false
});
assert.equal(
  otherNonCurrentLiabilityResidual(
    {
      totalLiabilities: 5000,
      currentLiabilities: 1200,
      nonCurrentDebt: 1400,
      deferredTaxLiabilities: 150,
      nonCurrentLeaseLiabilities: 300,
      pensionLiabilities: 90
    },
    groupedLeaseContext
  ),
  2400
);

const conceptDrivenContext = buildLiabilityTemplateMappingContext([
  { statement: "income", label: "Current Debt", concepts: ["ShortTermBorrowings"] },
  { statement: "balance", label: "Debt Due Within One Year", concepts: ["ShortTermBorrowings"] },
  { statement: "balance", label: "Current Finance Lease Liabilities", concepts: ["FinanceLeaseLiabilityCurrent"] },
  { statement: "balance", label: "Income Tax Deferrals", concepts: ["DeferredIncomeTaxLiabilitiesNet"] },
  { statement: "balance", label: "Lease Liability Detail", concepts: ["OperatingLeaseLiabilityNoncurrent"] }
]);
assert.equal(conceptDrivenContext.hasCurrentDebtRow, true);
assert.equal(conceptDrivenContext.hasDeferredTaxLiabilityRow, true);
assert.equal(conceptDrivenContext.hasNonCurrentLeaseLiabilityRow, true);

console.log("Generic liability classification rules passed.");
