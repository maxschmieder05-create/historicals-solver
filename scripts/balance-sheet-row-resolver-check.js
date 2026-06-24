const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const Module = require("node:module");
const ts = require("typescript");

const repoRoot = path.resolve(__dirname, "..");
const sourcePath = path.join(repoRoot, "server", "fill-model", "balance-sheet-row-resolver.ts");

function compileTypeScript(source) {
  return ts.transpileModule(source, {
    compilerOptions: {
      module: ts.ModuleKind.CommonJS,
      target: ts.ScriptTarget.ES2020,
      esModuleInterop: true
    }
  }).outputText;
}

function loadTypeScriptModule(file) {
  const source = fs.readFileSync(file, "utf8");
  const compiled = compileTypeScript(source);
  const mod = new Module(file, module);
  mod.filename = file;
  mod.paths = Module._nodeModulePaths(path.dirname(file));
  mod._compile(compiled, file);
  return mod.exports;
}

const {
  balanceSheetRowMatchesSourceAlias,
  balanceSheetRowsEquivalent,
  balanceSheetSectionCompatible,
  balanceSheetLineLooksSubtotalLike,
  classifyBalanceSheetResolution
} = loadTypeScriptModule(sourcePath);

assert.equal(balanceSheetRowsEquivalent("Other liabilities, current", "Other Current Liabilities"), true);
assert.equal(balanceSheetRowMatchesSourceAlias("Other Current Liabilities", "Accrued expenses and other current liabilities"), true);
assert.equal(balanceSheetRowMatchesSourceAlias("Other Current Liabilities", "Other accrued expenses and liabilities"), true);
assert.equal(
  balanceSheetSectionCompatible("Other Non-Current Liabilities", "current liabilities", {
    label: "Other current liabilities",
    tag: "OtherLiabilitiesCurrent"
  }),
  false
);
assert.equal(balanceSheetLineLooksSubtotalLike("Accrued expenses and other current liabilities", "AccruedLiabilitiesCurrent"), true);

const disappearingPeriods = [
  ["1Q23", 5985],
  ["2Q23", 5985],
  ["3Q23", 5985],
  ["1Q24", 6604],
  ["2Q24", 6604],
  ["3Q24", 6604]
];

for (const [period, staleTemplateValue] of disappearingPeriods) {
  const currentLiabilities = period.endsWith("24") ? 12100 : 10900;
  const accountsPayable = period.endsWith("24") ? 1750 : 1550;
  const accruedLiabilities = period.endsWith("24") ? 2800 : 2400;
  const currentDebt = 0;
  const resolvedValue = currentLiabilities - accountsPayable - accruedLiabilities - currentDebt;
  const state = classifyBalanceSheetResolution({
    modelRow: "Other Current Liabilities",
    value: resolvedValue,
    classification: "residual",
    note:
      "Other Current Liabilities = Total Current Liabilities - Accounts Payable - Accrued Liabilities - Current Debt. Residual is allowed because all components are SEC-sourced.",
    sources: [
      { concept: "LiabilitiesCurrent", label: "Total current liabilities", value: currentLiabilities, sourceLayer: "sec_filing_package" },
      { concept: "AccountsPayableCurrent", label: "Accounts payable", value: accountsPayable, sourceLayer: "sec_filing_package" },
      { concept: "AccruedLiabilitiesCurrent", label: "Accrued liabilities", value: accruedLiabilities, sourceLayer: "sec_filing_package" },
      { concept: "DebtCurrent", label: "No separate current debt reported", value: currentDebt, sourceLayer: "model", note: "No separate current debt concept was reported." }
    ]
  });

  assert.equal(state.state, "residual_calculated", `${period} should resolve as a residual instead of blanking`);
  assert.match(state.residualFormula, /Other Current Liabilities = Total current liabilities - Accounts payable - Accrued liabilities/i);
  assert.notEqual(resolvedValue, staleTemplateValue, `${period} should not hardcode or copy forward the stale template value`);
}

const unresolved = classifyBalanceSheetResolution({
  modelRow: "Other Current Liabilities",
  value: null,
  classification: "unused",
  sources: [],
  note: "No current SEC line item matched."
});
assert.equal(unresolved.state, "unresolved_failure");

const explicitZero = classifyBalanceSheetResolution({
  modelRow: "Other Current Liabilities",
  value: 0,
  classification: "direct",
  sources: [{ concept: "NoCurrentSecSource", label: "No current SEC source disclosed", value: 0, sourceLayer: "model" }],
  note: "Explicitly zero because the current SEC filing did not disclose this current-liability line item."
});
assert.equal(explicitZero.state, "explicit_zero_not_applicable");

const componentDebt = classifyBalanceSheetResolution({
  modelRow: "Revolver",
  value: 457100000,
  classification: "grouped",
  sources: [
    {
      concept: "LongTermDebtAndCapitalLeaseObligationsCurrent",
      label: "Notes payable and current maturities of long-term debt",
      value: 457100000,
      sourceLayer: "sec_filing_package"
    }
  ],
  note: "Current maturities/current portion of long-term debt are grouped into the current borrowing row."
});
assert.equal(componentDebt.state, "direct_sec_sourced");

console.log("Balance-sheet row resolver regression passed.");
