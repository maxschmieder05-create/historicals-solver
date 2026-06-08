const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const Module = require("node:module");
const ts = require("typescript");

const repoRoot = path.resolve(__dirname, "..");
const sourcePath = path.join(repoRoot, "app", "api", "fill-model", "audit-notes.ts");

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

const { lineItemMappingSentence, sourceLineItemLabels, normalizedLineItemComment } = loadTypeScriptModule(sourcePath);

assert.equal(
  lineItemMappingSentence("Accounts Receivable", {
    classification: "direct",
    sources: [{ concept: "AccountsReceivableNetCurrent", label: "Accounts Receivable", value: 125 }]
  }),
  "Accounts Receivable maps directly to Accounts Receivable."
);

assert.equal(
  lineItemMappingSentence("Prepaid & Other Current Assets", {
    classification: "grouped",
    note: "Mapped to directly reported prepaid expense / other current asset concepts.",
    sources: [
      { concept: "PrepaidExpenseCurrent", label: "Prepaid Expenses", value: 40 },
      { concept: "OtherAssetsCurrent", label: "Other Current Assets", value: 60 }
    ]
  }),
  "Prepaid & Other Current Assets includes Prepaid Expenses and Other Current Assets."
);

const residualPrepaid = {
  classification: "grouped",
  note: "Included current assets less separately modeled cash, receivables, and inventory.",
  sources: [
    { concept: "AssetsCurrent", label: "Assets, Current", value: 1000 },
    { concept: "CashAndCashEquivalentsAtCarryingValue", label: "Cash and Cash Equivalents", value: 250 },
    { concept: "AccountsReceivableNetCurrent", label: "Accounts Receivable", value: 200 },
    { concept: "InventoryNet", label: "Inventory", value: 150 }
  ]
};

assert.deepEqual(sourceLineItemLabels(residualPrepaid), []);
assert.equal(lineItemMappingSentence("Prepaid & Other Current Assets", residualPrepaid), "");

assert.equal(
  lineItemMappingSentence("Prepaid & Other Current Assets", {
    classification: "grouped",
    includedLineItems: [
      "Assets, Current excluding Cash and Cash Equivalents, Accounts Receivable, and Inventory",
      "Prepaid Expenses",
      "Other Current Assets",
      "Total Current Assets"
    ],
    sources: residualPrepaid.sources
  }),
  "Prepaid & Other Current Assets includes Prepaid Expenses and Other Current Assets."
);

assert.equal(
  lineItemMappingSentence("Other Current Liabilities", {
    classification: "grouped",
    note: "Mapped to directly reported current liability concepts that do not have dedicated model rows.",
    sources: [
      { concept: "DeferredRevenueCurrent", label: "Deferred Revenue", value: 10 },
      { concept: "OperatingLeaseLiabilityCurrent", label: "Operating Lease Liabilities", value: 20 },
      { concept: "TaxesPayableCurrent", label: "Taxes Payable", value: 30 },
      { concept: "OtherLiabilitiesCurrent", label: "Other Current Liabilities", value: 40 }
    ]
  }),
  "Other Current Liabilities includes Deferred Revenue, Operating Lease Liabilities, Taxes Payable, and Other Current Liabilities."
);

assert.equal(
  normalizedLineItemComment("Prepaid & Other Current Assets includes Assets, Current, Cash and Cash Equivalents, Accounts Receivable, and Inventory."),
  ""
);

assert.equal(normalizedLineItemComment("EDGAR: AssetsCurrent accession 0000000000"), "");

console.log("Audit note format rules passed.");
