const fs = require("node:fs/promises");
const os = require("node:os");
const path = require("node:path");
const vm = require("node:vm");
const { createRequire } = require("node:module");
const ts = require("typescript");
const ExcelJS = require("exceljs");

const repoRoot = path.resolve(__dirname, "..");
const sourcePath = path.join(repoRoot, "app", "api", "fill-model", "gold-model-library.ts");

function loadTypescriptModule(filePath) {
  const source = require("node:fs").readFileSync(filePath, "utf8");
  const compiled = ts.transpileModule(source, {
    compilerOptions: {
      esModuleInterop: true,
      module: ts.ModuleKind.CommonJS,
      target: ts.ScriptTarget.ES2020
    },
    fileName: filePath
  });
  const localRequire = createRequire(filePath);
  const module = { exports: {} };
  const sandbox = {
    module,
    exports: module.exports,
    require: localRequire,
    __dirname: path.dirname(filePath),
    __filename: filePath,
    console,
    process,
    Buffer,
    setTimeout,
    clearTimeout
  };
  vm.runInNewContext(compiled.outputText, sandbox, { filename: filePath });
  return module.exports;
}

async function writeModelWorkbook(filePath, options = {}) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = options.creator || "Verified User";
  const model = workbook.addWorksheet("Model");
  model.getCell("C3").value = "Income Statement";
  model.getCell("F4").value = "1Q23";
  model.getCell("G4").value = "2Q23";
  model.getCell("H4").value = "3Q23";
  model.getCell("I4").value = "FY23";
  model.getCell("C5").value = options.financial ? "Net Revenues" : "Revenue";
  model.getCell("C6").value = options.financial ? "Investment Banking" : "Cost of Revenue";
  model.getCell("C7").value = options.financial ? "Capital Markets" : "Gross Profit";
  model.getCell("C8").value = options.financial ? "Asset Management" : "Inventory";
  model.getCell("C9").value = "Net Income";
  model.getCell("C15").value = "Balance Sheet";
  model.getCell("C16").value = options.financial ? "Financial Instruments Owned" : "Accounts Receivable";
  model.getCell("C17").value = "Total Assets";
  model.getCell("C18").value = "Total Liabilities";
  model.getCell("C19").value = "Total Equity";
  workbook.addWorksheet("Segment Analysis").getCell("A1").value = "Segment Analysis";
  if (options.generated) workbook.addWorksheet("Mapping Audit").getCell("A1").value = "Generated";
  await workbook.xlsx.writeFile(filePath);
}

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

async function main() {
  const library = loadTypescriptModule(sourcePath);
  const tmp = await fs.mkdtemp(path.join(os.tmpdir(), "gold-model-library-"));
  const jefPath = path.join(tmp, "Jefferies Financial Group Inc. (JEF)_Valuation Workbook (10-Mar-2026) (1)(1).xlsx");
  const filledPath = path.join(tmp, "Jefferies Financial Group Inc. (JEF)_historicals_filled.xlsx");
  const verifiedFilledPath = path.join(tmp, "Verified Filled Co. (VFC)_historicals_filled.xlsx");
  const unverifiedPath = path.join(tmp, "Example Manufacturing Inc. (EXM)_Valuation Workbook (10-Mar-2026).xlsx");
  const generatedPath = path.join(tmp, "Generated Metadata Co. (GMC)_Valuation Workbook (10-Mar-2026).xlsx");
  const manifestPath = path.join(tmp, "gold_models.json");

  await writeModelWorkbook(jefPath, { financial: true });
  await writeModelWorkbook(filledPath, { financial: true });
  await writeModelWorkbook(verifiedFilledPath, { financial: false });
  await writeModelWorkbook(unverifiedPath, { financial: false });
  await writeModelWorkbook(generatedPath, { financial: false, generated: true, creator: "Historicals Solver" });
  await fs.writeFile(
    manifestPath,
    JSON.stringify(
      [
        {
          ticker: "JEF",
          company_name: "Jefferies Financial Group Inc.",
          CIK: "0001084580",
          model_type: "financial_services_broker_dealer",
          file_path: jefPath,
          verified_by_user: true,
          notes: "Completed user-verified financial-services model."
        },
        {
          ticker: "VFC",
          company_name: "Verified Filled Co.",
          CIK: "0000000001",
          model_type: "standard_operating_company",
          file_path: verifiedFilledPath,
          verified_by_user: true,
          notes: "User explicitly verified this completed filled workbook as a gold model."
        }
      ],
      null,
      2
    )
  );

  const scan = await library.scanGoldModelLibrary({ libraryPath: tmp, manifestPath });
  const jef = scan.models.find((model) => model.ticker === "JEF");
  assert(jef, "JEF verified model was not identified as a candidate.");
  assert(jef.usableAsGold, "JEF manifest-verified workbook should be usable as gold.");
  assert(jef.modelType === "financial_services_broker_dealer", `JEF model type should be financial_services_broker_dealer, got ${jef.modelType}.`);
  assert(jef.companyName === "Jefferies Financial Group Inc.", "Manifest company name should override filename guessing.");
  assert(jef.historicalPeriodsCovered.includes("1Q23") && jef.historicalPeriodsCovered.includes("FY23"), "Historical periods were not detected.");

  const filled = scan.excluded.find((model) => model.filePath === filledPath);
  assert(filled, "Workbook with filled in filename should be excluded.");
  assert(
    filled.exclusionReasons.some((reason) => /filled/i.test(reason)),
    "Filled workbook exclusion should cite the filled filename rule."
  );

  const verifiedFilled = scan.models.find((model) => model.filePath === verifiedFilledPath);
  assert(verifiedFilled, "Manifest-verified workbook with filled in filename should be allowed as a gold model.");
  assert(verifiedFilled.usableAsGold, "Manifest-verified filled workbook should be usable as gold.");
  assert(verifiedFilled.verifiedByUser, "Manifest-verified filled workbook should retain verified status.");

  const generated = scan.excluded.find((model) => model.filePath === generatedPath);
  assert(generated, "Historicals Solver generated workbook should be excluded.");
  assert(
    generated.exclusionReasons.some((reason) => /generated metadata|Historicals Solver/i.test(reason)),
    "Generated workbook exclusion should cite generated metadata."
  );

  const unverified = scan.models.find((model) => model.filePath === unverifiedPath);
  assert(unverified, "Unverified structurally valid workbook should remain a candidate.");
  assert(!unverified.verifiedByUser, "Unverified workbook should not be marked verified by filename alone.");
  assert(!unverified.usableAsGold, "Unverified workbook should not be usable as gold.");

  const company = { modelType: "financial_services_broker_dealer", confidence: "high", source: "sec_filing_signals", rationale: [] };
  const standardTemplate = { modelType: "standard_operating_company", confidence: "high", source: "workbook", rationale: [] };
  const compatibility = library.checkModelTemplateCompatibility(company, standardTemplate);
  assert(!compatibility.compatible, "Financial-services company should not be compatible with standard operating template.");
  assert(
    compatibility.message === "This company appears to require a financial-services model template. The uploaded standard operating-company template is not compatible.",
    `Unexpected compatibility message: ${compatibility.message}`
  );

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(jefPath);
  const workbookType = library.classifyWorkbookModelType(workbook);
  assert(
    workbookType.modelType === "financial_services_broker_dealer",
    `JEF-style workbook labels should classify as financial_services_broker_dealer, got ${workbookType.modelType}.`
  );

  const standardWithInterestIncome = library.classifyCompanyModelTypeFromSecSignals({
    companyName: "Alphabet Inc.",
    ticker: "GOOG",
    cik: "0001652044",
    conceptNames: [
      "RevenueFromContractWithCustomerExcludingAssessedTax",
      "CostOfRevenue",
      "GrossProfit",
      "InventoryNet",
      "AccountsReceivableNetCurrent",
      "InterestIncomeExpenseNonOperatingNet",
      "BrokerDealerRelatedReceivables",
      "SecuritiesBorrowed",
      "SecuritiesLoaned",
      "InsuranceRecoveries",
      "PremiumsPaid",
      "ClaimsExpense",
      "CustomerDeposits",
      "PropertyAndEquipmentNet",
      "DebtUnderwritingFees",
      "Forfeitures",
      "LoansReceivable",
      "NetRevenues"
    ],
    labels: ["Revenue", "Cost of revenues", "Gross profit", "Inventories", "Accounts receivable", "Interest income", "Insurance", "Premiums", "Claims", "Bank deposits", "Property", "Underwriting", "Forfeitures", "Loans receivable", "Net revenues"]
  });
  assert(
    standardWithInterestIncome.modelType === "standard_operating_company",
    `Standard companies with ordinary interest income should not route as financial services, got ${standardWithInterestIncome.modelType}.`
  );

  const bankSignals = library.classifyCompanyModelTypeFromSecSignals({
    companyName: "Example Bancorp",
    ticker: "BNK",
    cik: "0000000002",
    conceptNames: ["Deposits", "LoansReceivable", "NetInterestIncome", "ProvisionForCreditLosses"],
    labels: ["Deposits", "Loans receivable", "Net interest income", "Provision for credit losses"]
  });
  assert(bankSignals.modelType === "bank", `Bank signals should classify as bank, got ${bankSignals.modelType}.`);

  const jefCompanySignals = library.classifyCompanyModelTypeFromSecSignals({
    companyName: "Jefferies Financial Group Inc.",
    ticker: "JEF",
    cik: "0001084580",
    sic: "6211",
    sicDescription: "Security brokers, dealers, and flotation companies",
    conceptNames: ["InvestmentBankingRevenue", "PrincipalTransactionsRevenue", "FinancialInstrumentsOwned", "RevenuesNetOfInterestExpense"],
    labels: ["Investment banking", "Capital markets", "Principal transactions", "Financial instruments owned", "Net revenues"]
  });
  assert(
    jefCompanySignals.modelType === "financial_services_broker_dealer",
    `Strong JEF-like signals should classify as financial_services_broker_dealer, got ${jefCompanySignals.modelType}.`
  );

  console.log("Gold model library guard passed.");
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
