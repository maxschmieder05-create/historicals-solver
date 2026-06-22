const fs = require("node:fs/promises");
const fsSync = require("node:fs");
const path = require("node:path");
const os = require("node:os");
const Module = require("node:module");
const ts = require("typescript");
const ExcelJS = require("exceljs");
const { postWorkbook } = require("./fill-workbook-api");

const repoRoot = path.resolve(__dirname, "..");
const apiUrl = process.env.FILL_API_URL || "http://localhost:3000/api/fill-model";
const outputDir = path.resolve(process.env.REGRESSION_OUTPUT_DIR || path.join(repoRoot, "tmp", "regression-basket"));
const resultFile = path.resolve(process.env.REGRESSION_RESULTS_FILE || path.join(outputDir, "results.json"));
const baselineOut = process.env.REGRESSION_BASELINE_OUT ? path.resolve(process.env.REGRESSION_BASELINE_OUT) : "";
const compareBaseline = process.env.REGRESSION_COMPARE_BASELINE ? path.resolve(process.env.REGRESSION_COMPARE_BASELINE) : "";
const allowMissingInputs = process.env.REGRESSION_ALLOW_MISSING_INPUTS === "1";
const routingOnly = process.env.REGRESSION_ROUTING_ONLY === "1";
const quiet = process.env.REGRESSION_QUIET === "1";
const secHeaders = {
  "User-Agent": process.env.SEC_USER_AGENT || "HistoricalsSolver regression harness contact@example.com"
};
const secFetchTimeoutMs = Number(process.env.SEC_FETCH_TIMEOUT_MS || 60_000);

const defaultStandardTemplate = path.join(os.homedir(), "Desktop", "Owl Fund Integrated Model Template (03-Sep-2025)_v25 (3).xlsx");
const defaultFinancialTemplate = path.join(os.homedir(), "Desktop", "Jefferies Financial Group Inc. (JEF)_Valuation Workbook (10-Mar-2026) (1).xlsx");

const basket = [
  { ticker: "AAPL", cik: "0000320193", companyName: "Apple Inc.", companyType: "standard_operating_company" },
  { ticker: "GOOG", cik: "0001652044", companyName: "Alphabet Inc.", companyType: "standard_operating_company" },
  { ticker: "NVDA", cik: "0001045810", companyName: "NVIDIA Corporation", companyType: "standard_operating_company" },
  { ticker: "KO", cik: "0000021344", companyName: "The Coca-Cola Company", companyType: "standard_operating_company" },
  { ticker: "PANW", cik: "0001327567", companyName: "Palo Alto Networks, Inc.", companyType: "standard_operating_company" },
  { ticker: "LLY", cik: "0000059478", companyName: "Eli Lilly and Company", companyType: "standard_operating_company" },
  { ticker: "INTC", cik: "0000050863", companyName: "Intel Corporation", companyType: "standard_operating_company" },
  { ticker: "MCD", cik: "0000063908", companyName: "McDonald's Corporation", companyType: "standard_operating_company" },
  { ticker: "PEP", cik: "0000077476", companyName: "PepsiCo, Inc.", companyType: "standard_operating_company" },
  { ticker: "CPB", cik: "0000016732", companyName: "The Campbell's Company", companyType: "standard_operating_company" },
  { ticker: "IBM", cik: "0000051143", companyName: "International Business Machines Corporation", companyType: "standard_operating_company" },
  { ticker: "RSG", cik: "0001060391", companyName: "Republic Services, Inc.", companyType: "standard_operating_company" },
  {
    ticker: "DE",
    cik: "0000315189",
    companyName: "Deere & Company",
    companyType: "standard_operating_company",
    profile: "financing_heavy_operating_balance_sheet"
  },
  {
    ticker: "JEF",
    cik: "0001084580",
    companyName: "Jefferies Financial Group Inc.",
    companyType: "financial_services_broker_dealer",
    sic: "6211",
    sicDescription: "Security brokers, dealers, and flotation companies",
    requireFinancialTemplate: true
  }
];

const incomeChecks = [
  {
    key: "revenue",
    label: "Revenue",
    aliases: ["Revenue", "Revenues", "Total Revenue", "Total Revenues", "Net Sales", "Net Revenue", "Net Revenues"],
    concepts: ["RevenueFromContractWithCustomerExcludingAssessedTax", "RevenueFromContractWithCustomerIncludingAssessedTax", "SalesRevenueNet", "Revenues"],
    sign: "reported",
    required: true
  },
  {
    key: "cogs",
    label: "COGS / Cost of Revenue",
    aliases: ["Cost of Revenue", "Cost of Goods Sold", "COGS", "COGS / Cost of Revenue", "Cost of Sales"],
    concepts: [
      "CostOfRevenue",
      "CostOfGoodsAndServicesSold",
      "CostOfGoodsSold",
      "CostOfProductsSold",
      "CostOfServicesRevenue",
      "CostOfRevenueExcludingDepreciationDepletionAndAmortization",
      "CostOfGoodsSoldExcludingDepreciationDepletionAndAmortization"
    ],
    sign: "expense",
    required: true
  },
  {
    key: "gross_profit",
    label: "Gross Profit",
    aliases: ["Gross Profit", "Gross Margin"],
    concepts: ["GrossProfit"],
    sign: "reported",
    required: true
  },
  {
    key: "sga",
    label: "SG&A",
    aliases: ["Selling, General & Administration (SG&A)", "SG&A", "Selling General and Administrative", "Selling, General and Administrative"],
    concepts: ["SellingGeneralAndAdministrativeExpense", "SellingGeneralAndAdministrativeExpenseExcludingDepreciationDepletionAndAmortization", "GeneralAndAdministrativeExpense"],
    sign: "expense",
    required: true
  },
  {
    key: "rd",
    label: "R&D",
    aliases: ["Research & Development (R&D)", "Research and Development", "R&D"],
    concepts: ["ResearchAndDevelopmentExpense", "ResearchAndDevelopmentExpenseExcludingAcquiredInProcessCost"],
    sign: "expense",
    optionalIfNoSource: true
  },
  {
    key: "da_income_statement",
    label: "D&A",
    aliases: ["D&A", "Depreciation & Amortization", "Depreciation and Amortization"],
    concepts: ["DepreciationDepletionAndAmortization", "DepreciationAndAmortization"],
    sign: "expense",
    incomeStatementOnly: true,
    optionalIfNoSource: true
  },
  {
    key: "other_operating",
    label: "Other Operating Income / Expense",
    aliases: ["Other Operating Income / Expense", "Other Operating Income (Expense)", "Other Operating Expense", "Other Operating Income"],
    concepts: ["OtherOperatingIncomeExpenseNet", "OtherOperatingIncome", "OtherOperatingExpense", "OtherOperatingCostsAndExpenses"],
    sign: "reported",
    optionalIfNoSource: true
  },
  {
    key: "operating_income",
    label: "EBIT / Operating Income",
    aliases: ["EBIT", "Operating Income", "Operating Income (Loss)", "Income from Operations"],
    concepts: ["OperatingIncomeLoss", "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest"],
    sign: "reported",
    required: true
  },
  {
    key: "interest_income",
    label: "Interest Income",
    aliases: ["Interest Income", "Interest Income (Expense)", "Interest and Other Income"],
    concepts: ["InterestIncomeExpenseNonOperatingNet", "InterestIncomeNonOperating", "InvestmentIncomeInterest"],
    sign: "reported",
    optionalIfNoSource: true
  },
  {
    key: "interest_expense",
    label: "Interest Expense",
    aliases: ["Interest Expense", "Interest (Expense)", "Interest Expense, Net"],
    concepts: ["InterestExpenseNonOperating", "InterestExpense", "InterestAndDebtExpense"],
    sign: "expense",
    optionalIfNoSource: true
  },
  {
    key: "other_non_operating",
    label: "Other Non-Operating Income / Expense",
    aliases: ["Other Non-Operating Income / Expense", "Other Non-Operating Income (Expense)", "Other Income (Expense)", "Other Income / Expense"],
    concepts: ["OtherNonoperatingIncomeExpense", "NonoperatingIncomeExpense", "OtherIncomeExpenseNet", "OtherIncome", "OtherExpense"],
    sign: "reported",
    optionalIfNoSource: true
  },
  {
    key: "pretax_income",
    label: "Pre-Tax Income",
    aliases: ["Pre-Tax Income (Loss)", "Pretax Income", "Income Before Taxes", "Income Before Income Taxes"],
    concepts: [
      "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
      "IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments",
      "IncomeLossFromContinuingOperationsBeforeIncomeTaxes"
    ],
    sign: "reported",
    required: true
  },
  {
    key: "tax_expense",
    label: "Tax Expense / Benefit",
    aliases: ["Income Tax Benefit (Expense)", "Income Tax Expense", "Tax Expense", "Provision for Income Taxes"],
    concepts: ["IncomeTaxExpenseBenefit"],
    sign: "expense",
    required: true
  },
  {
    key: "net_income",
    label: "Net Income",
    aliases: ["Net Income (Loss)", "Net Income", "Profit (Loss)", "Profit Loss"],
    concepts: ["NetIncomeLoss", "ProfitLoss"],
    sign: "reported",
    required: true
  },
  {
    key: "parent_net_income",
    label: "Parent-attributable Net Income",
    aliases: ["Net Income Attributable to Parent", "Net Income Attributable to Common Shareholders", "Net Income Available to Common Shareholders"],
    concepts: ["NetIncomeLossAvailableToCommonStockholdersBasic", "NetIncomeLossAttributableToParent", "ProfitLossAttributableToOwnersOfParent"],
    sign: "reported",
    optionalIfNoSource: true
  }
];

const balanceSheetChecks = [
  {
    key: "current_assets",
    label: "Total Current Assets",
    aliases: ["Total Current Assets", "Current Assets"],
    concepts: ["AssetsCurrent"],
    optionalIfNoSource: true
  },
  {
    key: "total_assets",
    label: "Total Assets",
    aliases: ["Total Assets"],
    concepts: ["Assets"],
    required: true
  },
  {
    key: "current_liabilities",
    label: "Total Current Liabilities",
    aliases: ["Total Current Liabilities", "Current Liabilities"],
    concepts: ["LiabilitiesCurrent"],
    optionalIfNoSource: true
  },
  {
    key: "total_liabilities",
    label: "Total Liabilities",
    aliases: ["Total Liabilities"],
    concepts: ["Liabilities"],
    required: true
  },
  {
    key: "total_equity",
    label: "Total Equity",
    aliases: ["Total Shareholder's Equity", "Total Shareholders' Equity", "Total Stockholders' Equity", "Total Equity", "Shareholders' Equity", "Stockholders' Equity"],
    concepts: ["StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest", "StockholdersEquity", "PartnersCapital"],
    required: true
  },
  {
    key: "liabilities_and_equity",
    label: "Total Liabilities & Equity",
    aliases: ["Total Liabilities & Shareholder's Equity", "Total Liabilities and Shareholder's Equity", "Total Liabilities & Stockholders' Equity", "Total Liabilities and Stockholders' Equity", "Total Liabilities & Equity"],
    concepts: ["LiabilitiesAndStockholdersEquity", "LiabilitiesAndPartnersCapital", "Assets"],
    required: true
  }
];

function compileTypeScript(source) {
  return ts.transpileModule(source, {
    compilerOptions: {
      esModuleInterop: true,
      module: ts.ModuleKind.CommonJS,
      target: ts.ScriptTarget.ES2020
    }
  }).outputText;
}

function loadTypeScriptModule(file) {
  const source = fsSync.readFileSync(file, "utf8");
  const mod = new Module(file, module);
  mod.filename = file;
  mod.paths = Module._nodeModulePaths(path.dirname(file));
  mod._compile(compileTypeScript(source), file);
  return mod.exports;
}

async function main() {
  await fs.mkdir(outputDir, { recursive: true });
  const library = loadTypeScriptModule(path.join(repoRoot, "server", "fill-model", "gold-model-library.ts"));
  logProgress("Scanning verified gold-model library");
  const goldScan = await scanGoldLibrary(library);
  logProgress(`Gold-model scan found ${goldScan.models.length} usable candidate(s), ${goldScan.excluded.length} excluded candidate(s)`);
  const selectedCases = selectedBasketCases(await readCaseOverrides());
  const results = [];

  for (const testCase of selectedCases) {
    const result = await runCase(testCase, { library, goldScan });
    results.push(result);
    const icon = result.status === "passed" ? "PASS" : result.status === "skipped" ? "SKIP" : "FAIL";
    console.log(`${icon} ${testCase.ticker}: ${result.summary}`);
    for (const detail of result.errors.slice(0, 8)) console.log(`  - ${detail}`);
  }

  const payload = {
    generatedAt: new Date().toISOString(),
    apiUrl,
    routingOnly,
    resultFile,
    cases: results
  };
  await fs.writeFile(resultFile, `${JSON.stringify(payload, null, 2)}\n`);
  if (baselineOut) await fs.writeFile(baselineOut, `${JSON.stringify(payload, null, 2)}\n`);
  if (compareBaseline) await compareAgainstBaseline(payload, compareBaseline);

  const failed = results.filter((result) => result.status === "failed");
  const skipped = results.filter((result) => result.status === "skipped");
  console.log(`Regression basket complete: ${results.length - failed.length - skipped.length} passed, ${failed.length} failed, ${skipped.length} skipped.`);
  console.log(`Results written to ${resultFile}`);
  if (failed.length) process.exit(1);
  if (skipped.length && !allowMissingInputs) process.exit(1);
}

async function scanGoldLibrary(library) {
  const manifestPath = configuredManifestPath();
  const libraryPath = configuredLibraryPath(manifestPath);
  const verifiedFolders = splitPathList(process.env.GOLD_MODEL_VERIFIED_PATHS || process.env.GOLD_MODEL_VERIFIED_PATH || "");
  const scan = await library.scanGoldModelLibrary({ libraryPath, manifestPath, verifiedFolders });
  for (const warning of scan.warnings) console.warn(`Gold library warning: ${warning}`);
  return scan;
}

function configuredManifestPath() {
  const explicit = process.env.GOLD_MODEL_MANIFEST_PATH;
  if (explicit) return path.resolve(explicit);
  for (const candidate of [path.join(repoRoot, "config", "gold_models.json"), path.join(repoRoot, "config", "gold_models.example.json")]) {
    if (fsSync.existsSync(candidate)) return candidate;
  }
  return null;
}

function configuredLibraryPath(manifestPath) {
  const explicit = process.env.GOLD_MODEL_LIBRARY_PATH || process.env.GOLD_MODELS_PATH;
  if (explicit) return path.resolve(explicit);
  if (manifestPath) return path.dirname(manifestPath);
  return null;
}

function splitPathList(value) {
  return String(value || "")
    .split(path.delimiter)
    .flatMap((item) => item.split(","))
    .map((item) => item.trim())
    .filter(Boolean)
    .map((item) => path.resolve(item));
}

async function readCaseOverrides() {
  const configPath = process.env.REGRESSION_BASKET_CONFIG
    ? path.resolve(process.env.REGRESSION_BASKET_CONFIG)
    : path.join(repoRoot, "config", "regression_basket.local.json");
  if (!fsSync.existsSync(configPath)) return {};
  const parsed = JSON.parse(await fs.readFile(configPath, "utf8"));
  const cases = Array.isArray(parsed?.cases) ? parsed.cases : [];
  return Object.fromEntries(cases.map((item) => [String(item.ticker || "").toUpperCase(), item]));
}

function selectedBasketCases(overrides) {
  const enabled = new Set(
    (process.env.REGRESSION_BASKET_CASES || basket.map((item) => item.ticker).join(","))
      .split(",")
      .map((item) => item.trim().toUpperCase())
      .filter(Boolean)
  );
  return basket
    .filter((testCase) => enabled.has(testCase.ticker))
    .map((testCase) => ({ ...testCase, ...(overrides[testCase.ticker] || {}) }));
}

async function runCase(testCase, context) {
  const startedAt = Date.now();
  const result = {
    ticker: testCase.ticker,
    cik: normalizeCik(testCase.cik),
    companyName: testCase.companyName,
    companyType: testCase.companyType,
    status: "failed",
    summary: "",
    outputWorkbook: null,
    inputWorkbook: null,
    goldModel: null,
    checks: {},
    warnings: [],
    errors: [],
    durationMs: 0
  };

  try {
    logProgress(`${testCase.ticker}: loading SEC companyfacts`);
    const facts = await fetchCompanyFacts(testCase.cik);
    logProgress(`${testCase.ticker}: resolving gold model and input workbook`);
    const goldModel = context.library.findVerifiedGoldModelForCompany(context.goldScan, {
      ticker: testCase.ticker,
      cik: testCase.cik,
      title: testCase.companyName
    });
    if (goldModel) {
      result.goldModel = {
        filePath: goldModel.filePath,
        fileName: goldModel.fileName,
        modelType: goldModel.modelType
      };
    }

    const inputWorkbook = resolveInputWorkbook(testCase, goldModel);
    result.inputWorkbook = inputWorkbook;
    if (!inputWorkbook || !fsSync.existsSync(inputWorkbook)) {
      const message = `${testCase.ticker}: no input template workbook found. Set REGRESSION_INPUT_${testCase.ticker}, REGRESSION_STANDARD_TEMPLATE, REGRESSION_FINANCIAL_TEMPLATE, or add a verified gold model manifest entry.`;
      if (allowMissingInputs) {
        result.status = "skipped";
        result.summary = message;
        result.warnings.push(message);
        return result;
      }
      throw new Error(message);
    }

    logProgress(`${testCase.ticker}: validating model-type routing`);
    await validateRouting(testCase, facts, inputWorkbook, goldModel, context.library, result);
    if (result.errors.length) throw new Error(result.errors[0]);
    if (routingOnly) {
      result.status = "passed";
      result.summary = "routing and gold manifest checks passed";
      return result;
    }

    const outputWorkbook = path.join(outputDir, `${testCase.ticker}_regression_output.xlsx`);
    logProgress(`${testCase.ticker}: generating workbook through ${apiUrl}`);
    await postWorkbook({ apiUrl, ticker: testCase.ticker, inputWorkbook, outputWorkbook });
    result.outputWorkbook = outputWorkbook;

    logProgress(`${testCase.ticker}: validating returned workbook ${outputWorkbook}`);
    await validateReturnedWorkbook(testCase, {
      facts,
      inputWorkbook,
      outputWorkbook,
      goldModel,
      library: context.library,
      result
    });

    if (result.errors.length) {
      result.status = "failed";
      result.summary = `${result.errors.length} validation issue(s)`;
    } else {
      result.status = "passed";
      result.summary = `${result.checks.totalChecks || 0} validation checks passed`;
    }
    return result;
  } catch (error) {
    result.status = result.status === "skipped" ? "skipped" : "failed";
    result.errors.push(error instanceof Error ? error.message : String(error));
    result.summary = result.errors[0] || "failed";
    return result;
  } finally {
    result.durationMs = Date.now() - startedAt;
  }
}

function resolveInputWorkbook(testCase, goldModel) {
  const ticker = testCase.ticker.toUpperCase();
  const specific = process.env[`REGRESSION_INPUT_${ticker}`] || testCase.inputWorkbook;
  if (specific) return path.resolve(specific);
  if (goldModel?.filePath && fsSync.existsSync(goldModel.filePath)) return goldModel.filePath;
  if (testCase.requireFinancialTemplate || testCase.companyType === "financial_services_broker_dealer") {
    const financial = process.env.REGRESSION_FINANCIAL_TEMPLATE || defaultFinancialTemplate;
    return fsSync.existsSync(financial) ? financial : "";
  }
  const standard = process.env.REGRESSION_STANDARD_TEMPLATE || defaultStandardTemplate;
  return fsSync.existsSync(standard) ? standard : "";
}

async function validateRouting(testCase, facts, inputWorkbook, goldModel, library, result) {
  const signals = companySignals(testCase, facts);
  const companyType = goldModel ? library.classifyModelTypeFromGoldReference(goldModel) : library.classifyCompanyModelTypeFromSecSignals(signals);
  const templateWorkbook = new ExcelJS.Workbook();
  await templateWorkbook.xlsx.readFile(inputWorkbook);
  const templateType = library.classifyWorkbookModelType(templateWorkbook);
  const compatibility = library.checkModelTemplateCompatibility(companyType, templateType);
  result.checks.modelRouting = { companyType, templateType, compatible: compatibility.compatible, message: compatibility.message };

  if (testCase.ticker === "JEF") {
    if (companyType.modelType !== "financial_services_broker_dealer") {
      result.errors.push(`JEF should classify as financial_services_broker_dealer, got ${companyType.modelType}.`);
    }
    const standardCompatibility = library.checkModelTemplateCompatibility(companyType, {
      modelType: "standard_operating_company",
      confidence: "high",
      source: "template_profile",
      rationale: ["Regression harness standard-template incompatibility probe."]
    });
    if (standardCompatibility.compatible) {
      result.errors.push("JEF compatibility probe allowed a standard operating-company template.");
    }
    if (!/financial-services model template/i.test(standardCompatibility.message || "")) {
      result.errors.push(`JEF compatibility message was not clean/financial-services-specific: ${standardCompatibility.message || "[blank]"}`);
    }
  }

  if (!compatibility.compatible) {
    result.errors.push(compatibility.message || `${testCase.ticker}: input workbook is not compatible with company model type.`);
  }
}

function companySignals(testCase, facts) {
  const conceptNames = [];
  const labels = [];
  for (const taxonomy of Object.keys(facts.facts || {})) {
    for (const [concept, payload] of Object.entries(facts.facts[taxonomy] || {})) {
      conceptNames.push(concept);
      if (payload?.label) labels.push(payload.label);
      if (payload?.description) labels.push(payload.description);
    }
  }
  return {
    companyName: facts.entityName || testCase.companyName,
    ticker: testCase.ticker,
    cik: testCase.cik,
    sic: facts.sic || testCase.sic,
    sicDescription: facts.sicDescription || testCase.sicDescription,
    conceptNames,
    labels
  };
}

async function validateReturnedWorkbook(testCase, options) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(options.outputWorkbook);
  const input = new ExcelJS.Workbook();
  await input.xlsx.readFile(options.inputWorkbook);
  const gold = options.goldModel ? await readOptionalWorkbook(options.goldModel.filePath) : null;
  const result = options.result;
  const errors = result.errors;
  const warnings = result.warnings;
  const model = workbook.getWorksheet("Model");
  if (!model) {
    errors.push("Returned workbook is missing the Model sheet.");
    return;
  }

  const periods = detectPeriodColumns(model).filter((item) => !item.isEstimate);
  if (!periods.length) errors.push("Returned workbook has no detected historical period columns.");

  if (testCase.companyType === "financial_services_broker_dealer") {
    validateFinancialServicesWorkbook(testCase, workbook, result);
  } else {
    await validateIncomeStatementAnchors(testCase, model, periods, options.facts, result);
    await validateBalanceSheetAnchors(testCase, model, periods, options.facts, result);
  }
  validateBalanceSheetCheck(model, periods, result);
  validateSegmentAnalysis(workbook, model, periods, result);
  validateFormulaSafety(workbook, input, result);
  validateSourceLedger(workbook, model, periods, testCase, result);
  validateClassificationLedgers(workbook, testCase, result);
  if (gold && options.goldModel) compareGoldStructure(workbook, gold, options.goldModel, result);
  else warnings.push(`${testCase.ticker}: no verified gold model was available; EDGAR, formula, ledger, and routing checks still ran.`);
}

function validateFinancialServicesWorkbook(testCase, workbook, result) {
  const model = workbook.getWorksheet("Model");
  if (!model) return;
  const text = workbookText(workbook);
  if (!/net revenues|investment banking|capital markets|financial instruments owned/i.test(text)) {
    result.errors.push(`${testCase.ticker}: financial-services workbook does not expose expected broker-dealer labels.`);
  }
  const standardRows = ["Gross Profit", "Inventory", "Cost of Goods Sold"].filter((label) => findRow(model, [label]));
  if (standardRows.length >= 2) {
    result.errors.push(`${testCase.ticker}: financial-services output appears to have been forced into standard operating-company rows (${standardRows.join(", ")}).`);
  }
  increment(result, "financialServicesChecks");
}

async function validateIncomeStatementAnchors(testCase, sheet, periods, facts, result) {
  const errors = result.errors;
  const warnings = result.warnings;
  for (const check of incomeChecks) {
    const row = findRowInSection(sheet, check.aliases, "income");
    if (!row) {
      if (check.required) errors.push(`${testCase.ticker}: missing required income-statement row for ${check.label}.`);
      else warnings.push(`${testCase.ticker}: optional income-statement row not found for ${check.label}.`);
      continue;
    }
    if (check.incomeStatementOnly && !rowHasIncomeStatementLedgerSupport(sheet, row)) {
      warnings.push(`${testCase.ticker}: skipped ${check.label} EDGAR comparison because the output did not show standalone income-statement source support.`);
      continue;
    }
    for (const period of periods) {
      const source = expectedDurationValue(facts, period.period, check.concepts);
      if (!source) {
        if (check.required && !check.optionalIfNoSource) warnings.push(`${testCase.ticker} ${period.period}: no EDGAR source fact found for ${check.label}.`);
        continue;
      }
      const expected = applyModelSign(source.value / 1_000_000, check.sign);
      const actual = numericCellValue(sheet.getCell(row, period.col));
      if (!valuesMatch(actual, expected)) {
        errors.push(`${testCase.ticker} ${period.period} ${check.label} ${sheet.getCell(row, period.col).address}: expected ${round(expected)} from EDGAR ${source.concept}, got ${actual ?? "[blank]"}.`);
      } else {
        increment(result, "incomeStatementEdgarChecks");
      }
    }
  }
}

async function validateBalanceSheetAnchors(testCase, sheet, periods, facts, result) {
  const errors = result.errors;
  const warnings = result.warnings;
  for (const check of balanceSheetChecks) {
    const row = findRowInSection(sheet, check.aliases, "balance");
    if (!row) {
      if (check.required) errors.push(`${testCase.ticker}: missing required balance-sheet row for ${check.label}.`);
      else warnings.push(`${testCase.ticker}: optional balance-sheet row not found for ${check.label}.`);
      continue;
    }
    for (const period of periods) {
      const source = expectedInstantValue(facts, period.period, check.concepts);
      if (!source) {
        if (check.required && !check.optionalIfNoSource) warnings.push(`${testCase.ticker} ${period.period}: no EDGAR instant source found for ${check.label}.`);
        continue;
      }
      const expected = source.value / 1_000_000;
      const actual = numericCellValue(sheet.getCell(row, period.col));
      if (!valuesMatch(actual, expected)) {
        errors.push(`${testCase.ticker} ${period.period} ${check.label} ${sheet.getCell(row, period.col).address}: expected ${round(expected)} from EDGAR ${source.concept}, got ${actual ?? "[blank]"}.`);
      } else {
        increment(result, "balanceSheetEdgarChecks");
      }
    }
  }
}

function validateBalanceSheetCheck(sheet, periods, result) {
  const row = findRowInSection(sheet, ["Balance Sheet Check", "BS Check", "Check"], "balance");
  if (!row) {
    result.errors.push("Balance Sheet Check row was not found.");
    return;
  }
  for (const period of periods) {
    const cell = sheet.getCell(row, period.col);
    const value = cellValue(cell);
    const numeric = numericCellValue(cell);
    if (typeof numeric === "number" ? Math.abs(numeric) > 0.5 : !/^ok$/i.test(String(value || "").trim())) {
      result.errors.push(`${period.period} Balance Sheet Check ${cell.address}: expected OK/zero, got ${value ?? "[blank]"}.`);
    } else {
      increment(result, "balanceSheetCheckCells");
    }
  }
}

function validateSegmentAnalysis(workbook, model, periods, result) {
  const segment = workbook.getWorksheet("Segment Analysis");
  if (!segment) {
    result.warnings.push("Segment Analysis sheet not found; segment checks skipped.");
    return;
  }
  const totalCompanyRevenueRow = findRow(segment, ["Total Company Revenue", "Total Revenue", "Reported Revenue"]);
  const modelRevenueRow = findRowInSection(model, ["Revenue", "Revenues", "Total Revenue", "Net Sales", "Net Revenue"], "income");
  if (!totalCompanyRevenueRow || !modelRevenueRow) {
    result.warnings.push("Segment Analysis total revenue or Model revenue row was not found; segment tie check skipped.");
    return;
  }
  for (const period of periods) {
    const segmentCol = findPeriodColumn(segment, period.period) || period.col;
    const segmentRevenue = numericCellValue(segment.getCell(totalCompanyRevenueRow, segmentCol));
    const modelRevenue = numericCellValue(model.getCell(modelRevenueRow, period.col));
    if (!valuesMatch(segmentRevenue, modelRevenue)) {
      result.errors.push(`${period.period} Segment Analysis Total Company Revenue should equal Income Statement Revenue: expected ${modelRevenue ?? "[blank]"}, got ${segmentRevenue ?? "[blank]"}.`);
    } else {
      increment(result, "segmentRevenueTieChecks");
    }

    const detailRows = segmentRevenueDetailRows(segment, totalCompanyRevenueRow);
    if (!detailRows.length) {
      result.warnings.push(`${period.period}: no safe detailed segment breakout found; relying on Reported Revenue fallback.`);
      continue;
    }
    const detailSum = detailRows.reduce((sum, row) => sum + (numericCellValue(segment.getCell(row, segmentCol)) || 0), 0);
    if (!valuesMatch(detailSum, segmentRevenue)) {
      result.errors.push(`${period.period} Segment Analysis detail rows plus Other/Reconciliation should equal Total Company Revenue: expected ${segmentRevenue ?? "[blank]"}, got ${round(detailSum)}.`);
    } else {
      increment(result, "segmentDetailTieChecks");
    }
  }
}

function segmentRevenueDetailRows(sheet, totalRow) {
  const rows = [];
  const start = Math.max(1, totalRow - 20);
  for (let row = start; row < totalRow; row += 1) {
    const label = rowLabel(sheet, row);
    const normalized = normalize(label);
    if (!label) continue;
    if (/segmentanalysis|revenue|totalcompanyrevenue|reportedrevenue|check/.test(normalized)) continue;
    if (rowHasAnyNumber(sheet, row)) rows.push(row);
  }
  return rows;
}

function validateFormulaSafety(workbook, input, result) {
  for (const sheetName of ["Model", "Segment Analysis"]) {
    const sheet = workbook.getWorksheet(sheetName);
    if (!sheet) continue;
    forEachUsedCell(sheet, (cell, row, col) => {
      const value = cellValue(cell);
      const formula = cellFormula(cell);
      if (String(value ?? "").trim() === "#DIV/0!") {
        result.errors.push(`${sheetName}!${cell.address}: formula cache contains #DIV/0!.`);
      }
      if (/0\.00000001/.test(String(formula ?? value ?? ""))) {
        result.errors.push(`${sheetName}!${cell.address}: formula/check cell contains hardcoded 0.00000001.`);
      }
      if (/check|ok/i.test(rowLabel(sheet, row)) && !formula && input.getWorksheet(sheetName)) {
        const inputFormula = cellFormula(input.getWorksheet(sheetName).getCell(row, col));
        if (inputFormula) result.errors.push(`${sheetName}!${cell.address}: check/OK formula was overwritten.`);
      }
    });

    const inputSheet = input.getWorksheet(sheetName);
    if (inputSheet) validateForecastFormulaProtection(sheetName, sheet, inputSheet, result);
  }
}

function validateForecastFormulaProtection(sheetName, outputSheet, inputSheet, result) {
  const forecastCols = detectPeriodColumns(inputSheet).filter((period) => period.isEstimate).map((period) => period.col);
  for (const col of forecastCols) {
    for (let row = 1; row <= inputSheet.rowCount; row += 1) {
      const before = inputSheet.getCell(row, col);
      const beforeFormula = cellFormula(before);
      if (!beforeFormula || isNumericConstantFormula(beforeFormula)) continue;
      const afterFormula = cellFormula(outputSheet.getCell(row, col));
      if (!afterFormula) {
        result.errors.push(`${sheetName}!${outputSheet.getCell(row, col).address}: forecast/projected formula area was overwritten.`);
      } else {
        increment(result, "forecastFormulaProtectionChecks");
      }
    }
  }
}

function validateSourceLedger(workbook, model, periods, testCase, result) {
  const ledger = workbook.getWorksheet("Source Ledger");
  if (!ledger) {
    result.errors.push("Source Ledger sheet was not found in returned workbook.");
    return;
  }
  const rows = ledgerObjects(ledger);
  const byCellPeriod = new Map();
  for (const row of rows) {
    const period = normalizePeriodLabel(row["fiscal period"]);
    const key = `${row["workbook sheet"] || "Model"}!${row.cell}!${period}`;
    if (!byCellPeriod.has(key)) byCellPeriod.set(key, []);
    byCellPeriod.get(key).push(row);
    validateLedgerRow(row, testCase, result);
  }

  const hardcodedRows = candidateHardcodedHistoricalRows(model);
  for (const period of periods) {
    for (const row of hardcodedRows) {
      const cell = model.getCell(row, period.col);
      if (cellFormula(cell) || !isNumericCell(cell)) continue;
      const label = rowLabel(model, row);
      if (/check|margin|growth|ratio|percent|%|multiple/i.test(label)) continue;
      const entries = byCellPeriod.get(`Model!${cell.address}!${period.period}`) || [];
      if (!entries.length) {
        result.errors.push(`Model!${cell.address} ${period.period}: hardcoded historical input "${label}" has no Source Ledger support.`);
      } else if (!entries.some(isSupportedLedgerEntry)) {
        result.errors.push(`Model!${cell.address} ${period.period}: Source Ledger does not show current-company SEC, validated derived, explicit zero, or formula support.`);
      } else {
        increment(result, "sourceLedgerHardcodeChecks");
      }
    }
  }
}

function validateLedgerRow(row, testCase, result) {
  const status = String(row["mapping status"] || "");
  const value = Number(row.value);
  const ticker = String(row.ticker || "").toUpperCase();
  const cik = normalizeCik(row.CIK || row.cik || "");
  const normalizedAccessions = normalizeAccessionList(row["accession normalized"] || row["accession raw"] || row["accession number"] || "");
  if (ticker && ticker !== testCase.ticker) result.errors.push(`Source Ledger ${row.cell || "[unknown]"}: ticker ${ticker} does not match ${testCase.ticker}.`);
  if (cik && cik !== normalizeCik(testCase.cik)) result.errors.push(`Source Ledger ${row.cell || "[unknown]"}: CIK ${cik} does not match ${normalizeCik(testCase.cik)}.`);
  if (/stale_or_unsupported/i.test(status) && Number.isFinite(value) && Math.abs(value) > 0.0001) {
    result.errors.push(`Source Ledger ${row.cell || "[unknown]"}: nonzero hardcoded value has stale_or_unsupported status.`);
  }
  if (/explicit_zero_no_source_disclosed/i.test(status) && Number.isFinite(value) && Math.abs(value) > 0.0001) {
    result.errors.push(`Source Ledger ${row.cell || "[unknown]"}: explicit zero status used for nonzero value ${value}.`);
  }
  if (/explicit_current_sec_source/i.test(status)) {
    if (!normalizedAccessions.length) result.errors.push(`Source Ledger ${row.cell || "[unknown]"}: current SEC source entry has no normalized accession.`);
    for (const accession of normalizedAccessions) {
      const accessionCikValue = accessionCik(accession);
      if (accessionCikValue && accessionCikValue !== normalizeCik(testCase.cik)) {
        result.errors.push(`Source Ledger ${row.cell || "[unknown]"}: accession ${accession} belongs to ${accessionCikValue}, not ${normalizeCik(testCase.cik)}.`);
      }
    }
    if (!row["source XBRL tag"] && !row["source line item label"]) {
      result.errors.push(`Source Ledger ${row.cell || "[unknown]"}: current SEC source entry lacks source tag/label support.`);
    }
  }
  if (/validated_current_company_derived_value/i.test(status) && !row["classification reason"] && !row["source XBRL tag"]) {
    result.errors.push(`Source Ledger ${row.cell || "[unknown]"}: derived value lacks derivation support metadata.`);
  }
}

function isSupportedLedgerEntry(row) {
  return /explicit_current_sec_source|validated_current_company_derived_value|explicit_zero_no_source_disclosed|formula_preserved/i.test(String(row["mapping status"] || ""));
}

function validateClassificationLedgers(workbook, testCase, result) {
  const assignment = workbook.getWorksheet("Balance Sheet Assignment Ledger");
  if (!assignment) {
    result.warnings.push("Balance Sheet Assignment Ledger sheet was not found; primary balance-sheet assignment checks skipped.");
    return;
  }
  const rows = ledgerObjects(assignment);
  const assignedBySource = new Map();
  for (const row of rows) {
    const sourceLabel = String(row["source line item label"] || "");
    const modelRow = String(row["assigned model row"] || "");
    const status = String(row["assignment status"] || "");
    const sourceStatement = String(row["source statement"] || "");
    const sourceSection = String(row["source section"] || "");
    const key = [row["fiscal period"], row["source filing accession"], row["source row key"] || normalize(sourceLabel)].join("|");

    if (/assigned/i.test(status)) {
      if (!assignedBySource.has(key)) assignedBySource.set(key, []);
      assignedBySource.get(key).push(row);
    }

    if (/accounts and notes receivable/i.test(sourceLabel) && /asset/i.test(sourceSection) && /debt|revolver|lt debt/i.test(modelRow)) {
      result.errors.push(`${testCase.ticker}: Accounts and notes receivable in asset section mapped to debt row "${modelRow}".`);
    }
    if (/current maturit(?:y|ies).*long.?term debt/i.test(sourceLabel) && /revolver/i.test(modelRow)) {
      result.errors.push(`${testCase.ticker}: current maturities of long-term debt mapped to Revolver.`);
    }
    if (/(convertible senior notes|current maturit(?:y|ies).*long.?term debt)/i.test(sourceLabel) && /assigned/i.test(status) && !/lt debt/i.test(modelRow)) {
      result.errors.push(`${testCase.ticker}: ${sourceLabel} should map to LT Debt (Incl. Current Portion), got "${modelRow}".`);
    }
    if (/(short.?term borrowings|commercial paper|revolver borrowings)/i.test(sourceLabel) && /assigned/i.test(status) && !/revolver|short.?term borrow/i.test(modelRow)) {
      result.errors.push(`${testCase.ticker}: true short-term borrowing "${sourceLabel}" did not map to Revolver/short-term borrowings row.`);
    }
    if (/(deferred revenue|deferred income(?! tax))/i.test(sourceLabel) && /deferred income taxes/i.test(modelRow)) {
      result.errors.push(`${testCase.ticker}: deferred revenue/deferred income mapped to Deferred Income Taxes.`);
    }
    if (/deferred tax liabilit/i.test(sourceLabel) && /assigned/i.test(status) && !/deferred income taxes/i.test(modelRow)) {
      result.errors.push(`${testCase.ticker}: deferred tax liabilities should map to Deferred Income Taxes, got "${modelRow}".`);
    }
    if (/cash.?flow/i.test(sourceStatement) && /depreciation|amortization/i.test(sourceLabel) && /^d&a$/i.test(modelRow.trim())) {
      result.errors.push(`${testCase.ticker}: cash-flow-only D&A source populated income-statement D&A.`);
    }
    if (/primary/i.test(sourceStatement) && !/(assigned|excluded|formula|grouped|unsupported|subtotal)/i.test(status)) {
      result.errors.push(`${testCase.ticker}: primary balance-sheet source "${sourceLabel}" is not assigned exactly once or explicitly excluded (status "${status}").`);
    }
  }

  for (const [key, entries] of assignedBySource.entries()) {
    if (entries.length > 1) {
      result.errors.push(`${testCase.ticker}: primary balance-sheet source assigned more than once (${key}).`);
    }
  }

  validateInventoryDoubleCount(rows, testCase, result);
  increment(result, "classificationLedgerChecks");
}

function validateInventoryDoubleCount(rows, testCase, result) {
  const assigned = rows.filter((row) => /assigned/i.test(String(row["assignment status"] || "")) && /inventory/i.test(String(row["assigned model row"] || "")));
  const byPeriod = new Map();
  for (const row of assigned) {
    const key = `${row["fiscal period"]}|${row["source filing accession"]}`;
    if (!byPeriod.has(key)) byPeriod.set(key, []);
    byPeriod.get(key).push(row);
  }
  for (const [key, periodRows] of byPeriod.entries()) {
    const hasSubtotal = periodRows.some((row) => /^total inventories?$|^inventories?$/i.test(String(row["source line item label"] || "").trim()));
    const hasComponents = periodRows.some((row) => !/^total inventories?$|^inventories?$/i.test(String(row["source line item label"] || "").trim()));
    if (hasSubtotal && hasComponents) {
      result.errors.push(`${testCase.ticker}: inventory subtotal and components are both assigned for ${key}.`);
    }
  }
}

function compareGoldStructure(workbook, gold, goldModel, result) {
  const ignored = new Set(["Mapping Audit", "Source Ledger", "Filing Period Map", "Balance Sheet Assignment Ledger", "LLM Mapping Review"]);
  for (const goldSheet of gold.worksheets) {
    if (ignored.has(goldSheet.name)) continue;
    if (!workbook.getWorksheet(goldSheet.name)) {
      result.errors.push(`Verified gold model sheet "${goldSheet.name}" is missing from generated output.`);
    }
  }

  for (const sheetName of ["Model", "Segment Analysis"]) {
    const actualSheet = workbook.getWorksheet(sheetName);
    const goldSheet = gold.getWorksheet(sheetName);
    if (!actualSheet || !goldSheet) continue;
    compareRowLayout(sheetName, goldSheet, actualSheet, result);
    compareGoldCheckAndForecastFormulas(sheetName, goldSheet, actualSheet, result);
  }

  const model = workbook.getWorksheet("Model");
  const goldModelSheet = gold.getWorksheet("Model");
  if (model && goldModelSheet) {
    const actualPeriods = detectPeriodColumns(model).map((period) => period.period).join(",");
    const goldPeriods = detectPeriodColumns(goldModelSheet).map((period) => period.period).join(",");
    if (actualPeriods !== goldPeriods) {
      result.errors.push(`Historical period layout differs from verified gold model: expected ${goldPeriods}, got ${actualPeriods}.`);
    }
  }
  increment(result, "goldStructureChecks");
}

function compareRowLayout(sheetName, expected, actual, result) {
  const expectedLabels = rowLayoutSignature(expected);
  const actualLabels = rowLayoutSignature(actual);
  const max = Math.max(expectedLabels.length, actualLabels.length);
  const mismatches = [];
  for (let index = 0; index < max; index += 1) {
    if ((expectedLabels[index]?.label || "") !== (actualLabels[index]?.label || "")) {
      mismatches.push({ expected: expectedLabels[index], actual: actualLabels[index] });
      if (mismatches.length >= 8) break;
    }
  }
  if (mismatches.length) {
    result.errors.push(`${sheetName}: row layout differs from verified gold model. First mismatch expected "${mismatches[0].expected?.label || "[blank]"}", got "${mismatches[0].actual?.label || "[blank]"}".`);
  }
}

function compareGoldCheckAndForecastFormulas(sheetName, expected, actual, result) {
  const forecastCols = detectPeriodColumns(expected).filter((period) => period.isEstimate).map((period) => period.col);
  for (let row = 1; row <= expected.rowCount; row += 1) {
    const label = rowLabel(expected, row);
    const cols = new Set(forecastCols);
    if (/check|ok/i.test(label)) {
      for (let col = 1; col <= Math.min(expected.columnCount, actual.columnCount); col += 1) cols.add(col);
    }
    for (const col of cols) {
      const expectedFormula = cellFormula(expected.getCell(row, col));
      if (!expectedFormula || isNumericConstantFormula(expectedFormula)) continue;
      const actualFormula = cellFormula(actual.getCell(row, col));
      if (!actualFormula) {
        result.errors.push(`${sheetName}!${actual.getCell(row, col).address}: verified gold formula/check behavior was not preserved.`);
      }
    }
  }
}

function rowLayoutSignature(sheet) {
  const rows = [];
  for (let row = 1; row <= sheet.rowCount; row += 1) {
    const label = rowLabel(sheet, row);
    if (label) rows.push({ row, label: normalize(label) });
  }
  return rows;
}

async function readOptionalWorkbook(filePath) {
  if (!filePath || !fsSync.existsSync(filePath)) return null;
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  return workbook;
}

async function fetchCompanyFacts(cik) {
  return fetchJsonWithTimeout(
    `https://data.sec.gov/api/xbrl/companyfacts/CIK${normalizeCik(cik)}.json`,
    `SEC companyfacts for CIK ${normalizeCik(cik)}`
  );
}

async function fetchJsonWithTimeout(url, description) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), secFetchTimeoutMs);
  try {
    const response = await fetch(url, { headers: secHeaders, signal: controller.signal });
    if (!response.ok) throw new Error(`Could not load ${description}: ${response.status} ${response.statusText}`);
    return await response.json();
  } catch (error) {
    if (error?.name === "AbortError") throw new Error(`Timed out loading ${description} after ${Math.round(secFetchTimeoutMs / 1000)}s.`);
    throw error;
  } finally {
    clearTimeout(timer);
  }
}

function expectedDurationValue(facts, period, concepts) {
  const parsed = parsePeriod(period);
  if (!parsed) return null;
  for (const concept of concepts) {
    const factList = usdFacts(facts, concept).filter((fact) => factMatchesFiscalPeriod(fact, parsed));
    if (!factList.length) continue;
    if (parsed.kind === "fy") {
      const fact = bestFact(factList.filter((fact) => fact.fp === "FY" && durationDays(fact) >= 250));
      if (fact) return { concept, value: fact.val, fact };
      continue;
    }
    const direct = bestFact(factList.filter((fact) => fact.fp === `Q${parsed.quarter}` && durationDays(fact) > 0 && durationDays(fact) <= 120));
    if (direct) return { concept, value: direct.val, fact: direct };

    const ytd = bestFact(factList.filter((fact) => fact.fp === `Q${parsed.quarter}` && durationDays(fact) > 120));
    if (ytd && parsed.quarter === 1) return { concept, value: ytd.val, fact: ytd };
    if (ytd && parsed.quarter > 1) {
      const prior = expectedDurationValue(facts, `${parsed.quarter - 1}Q${String(parsed.year).slice(-2)}`, [concept]);
      const priorYtd = bestFact(factList.filter((fact) => fact.fp === `Q${parsed.quarter - 1}` && durationDays(fact) > 120));
      if (priorYtd) return { concept, value: ytd.val - priorYtd.val, fact: ytd, derived: "quarter_from_ytd" };
      if (prior) return { concept, value: ytd.val - prior.value, fact: ytd, derived: "quarter_from_ytd_less_prior_quarter" };
    }
    if (parsed.quarter === 4) {
      const annual = expectedDurationValue(facts, `FY${String(parsed.year).slice(-2)}`, [concept]);
      const q1 = expectedDurationValue(facts, `1Q${String(parsed.year).slice(-2)}`, [concept]);
      const q2 = expectedDurationValue(facts, `2Q${String(parsed.year).slice(-2)}`, [concept]);
      const q3 = expectedDurationValue(facts, `3Q${String(parsed.year).slice(-2)}`, [concept]);
      if (annual && q1 && q2 && q3) return { concept, value: annual.value - q1.value - q2.value - q3.value, fact: annual.fact, derived: "fourth_quarter_from_annual_less_quarters" };
    }
  }
  return null;
}

function expectedInstantValue(facts, period, concepts) {
  const parsed = parsePeriod(period);
  if (!parsed) return null;
  for (const concept of concepts) {
    const candidates = usdFacts(facts, concept).filter((fact) => {
      if (fact.start) return false;
      if (Number(fact.fy) !== parsed.year) return false;
      if (parsed.kind === "fy") return fact.fp === "FY" || fact.fp === "Q4";
      if (parsed.quarter === 4) return fact.fp === "FY" || fact.fp === "Q4";
      return fact.fp === `Q${parsed.quarter}`;
    });
    const fact = bestFact(candidates);
    if (fact) return { concept, value: fact.val, fact };
  }
  return null;
}

function usdFacts(facts, concept) {
  const out = [];
  for (const taxonomy of Object.keys(facts.facts || {})) {
    const payload = facts.facts[taxonomy]?.[concept];
    const units = payload?.units?.USD;
    if (Array.isArray(units)) out.push(...units.filter((fact) => typeof fact.val === "number"));
  }
  return out;
}

function factMatchesFiscalPeriod(fact, parsed) {
  if (Number(fact.fy) !== parsed.year) return false;
  if (parsed.kind === "fy") return fact.fp === "FY";
  if (parsed.quarter === 4) return fact.fp === "FY" || fact.fp === "Q4";
  return fact.fp === `Q${parsed.quarter}`;
}

function bestFact(candidates) {
  return candidates
    .slice()
    .filter((fact) => fact.form === "10-Q" || fact.form === "10-K" || fact.form === "20-F")
    .sort((a, b) => {
      const formScore = scoreForm(b.form) - scoreForm(a.form);
      if (formScore) return formScore;
      return String(b.filed || "").localeCompare(String(a.filed || "")) || String(b.end || "").localeCompare(String(a.end || ""));
    })[0] || null;
}

function scoreForm(form) {
  if (form === "10-K" || form === "20-F") return 3;
  if (form === "10-Q") return 2;
  return 1;
}

function durationDays(fact) {
  if (!fact.start || !fact.end) return 0;
  return Math.round((Date.parse(fact.end) - Date.parse(fact.start)) / 86_400_000);
}

function applyModelSign(value, sign) {
  if (sign === "expense") return value === 0 ? 0 : -Math.abs(value);
  return value;
}

function detectPeriodColumns(sheet) {
  const headerRow = bestPeriodHeaderRow(sheet);
  if (!headerRow) return [];
  const periods = [];
  for (let col = 1; col <= Math.min(sheet.columnCount, 200); col += 1) {
    const raw = String(cellValue(sheet.getCell(headerRow, col)) ?? "").trim();
    const period = normalizePeriodLabel(raw);
    if (!period) continue;
    periods.push({ period, col, raw, isEstimate: isEstimatePeriodLabel(raw) });
  }
  return periods;
}

function bestPeriodHeaderRow(sheet) {
  let best = null;
  for (let row = 1; row <= Math.min(sheet.rowCount, 140); row += 1) {
    let score = 0;
    for (let col = 1; col <= Math.min(sheet.columnCount, 200); col += 1) {
      const raw = String(cellValue(sheet.getCell(row, col)) ?? "").trim();
      const period = normalizePeriodLabel(raw);
      if (/^[1-4]Q\d{2}$/.test(period)) score += 4;
      else if (/^FY\d{2}$/.test(period)) score += 2;
    }
    if (score && (!best || score > best.score)) best = { row, score };
  }
  return best?.row || null;
}

function findPeriodColumn(sheet, period) {
  return detectPeriodColumns(sheet).find((item) => item.period === period)?.col || null;
}

function normalizePeriodLabel(value) {
  const raw = String(value ?? "").trim();
  const compact = raw.replace(/\s+/g, "").replace(/[’']/g, "").replace(/(?:A|E|EST|ESTIMATE|ACTUAL)$/i, "");
  const quarter = compact.match(/^([1-4])Q(\d{2}|\d{4})$/i);
  if (quarter) return `${quarter[1]}Q${quarter[2].slice(-2)}`.toUpperCase();
  const fiscal = compact.match(/^(?:FY)?(\d{4}|\d{2})$/i);
  if (fiscal) return `FY${fiscal[1].slice(-2)}`.toUpperCase();
  return "";
}

function isEstimatePeriodLabel(value) {
  return /(?:E|EST|ESTIMATE)$/i.test(String(value ?? "").trim().replace(/\s+/g, ""));
}

function parsePeriod(period) {
  const quarter = String(period).match(/^([1-4])Q(\d{2})$/i);
  if (quarter) return { kind: "quarter", quarter: Number(quarter[1]), year: 2000 + Number(quarter[2]) };
  const fiscal = String(period).match(/^FY(\d{2})$/i);
  if (fiscal) return { kind: "fy", year: 2000 + Number(fiscal[1]) };
  return null;
}

function findRowInSection(sheet, aliases, section) {
  const wanted = new Set(aliases.map(normalize));
  let activeSection = "";
  for (let row = 1; row <= sheet.rowCount; row += 1) {
    const label = rowLabel(sheet, row);
    const sectionHit = sectionName(label);
    if (sectionHit) activeSection = sectionHit;
    if (section && activeSection && activeSection !== section) continue;
    if (wanted.has(normalize(label))) return row;
  }
  return null;
}

function findRow(sheet, aliases) {
  const wanted = new Set(aliases.map(normalize));
  for (let row = 1; row <= sheet.rowCount; row += 1) {
    if (wanted.has(normalize(rowLabel(sheet, row)))) return row;
  }
  return null;
}

function sectionName(label) {
  const normalized = normalize(label);
  if (normalized === "incomestatement" || normalized === "statementofoperations") return "income";
  if (normalized === "balancesheet" || normalized === "statementoffinancialposition") return "balance";
  if (/cashflowstatement/.test(normalized)) return "cash_flow";
  if (/segmentanalysis/.test(normalized)) return "segment";
  return "";
}

function rowLabel(sheet, rowNumber) {
  for (let col = 1; col <= Math.min(sheet.columnCount || 20, 8); col += 1) {
    const text = String(cellValue(sheet.getCell(rowNumber, col)) ?? "").trim();
    if (text && !/^x$/i.test(text)) return text;
  }
  return "";
}

function workbookText(workbook) {
  const chunks = [];
  for (const sheet of workbook.worksheets) {
    chunks.push(sheet.name);
    for (let row = 1; row <= Math.min(sheet.rowCount, 200); row += 1) {
      const label = rowLabel(sheet, row);
      if (label) chunks.push(label);
    }
  }
  return chunks.join(" ");
}

function rowHasAnyNumber(sheet, row) {
  for (let col = 1; col <= Math.min(sheet.columnCount, 200); col += 1) {
    if (typeof numericCellValue(sheet.getCell(row, col)) === "number") return true;
  }
  return false;
}

function rowHasIncomeStatementLedgerSupport(sheet, row) {
  const workbook = sheet.workbook;
  const ledger = workbook.getWorksheet("Source Ledger");
  if (!ledger) return false;
  const rowNumber = String(row);
  return ledgerObjects(ledger).some((entry) => {
    return String(entry["model row"] || "") === rowNumber && /income/i.test(String(entry["source statement"] || ""));
  });
}

function candidateHardcodedHistoricalRows(sheet) {
  const rows = [];
  let activeSection = "";
  for (let row = 1; row <= sheet.rowCount; row += 1) {
    const label = rowLabel(sheet, row);
    const sectionHit = sectionName(label);
    if (sectionHit) activeSection = sectionHit;
    if ((activeSection === "income" || activeSection === "balance") && label && !sectionHit) rows.push(row);
  }
  return rows;
}

function ledgerObjects(sheet) {
  const headers = {};
  sheet.getRow(1).eachCell((cell, col) => {
    const key = String(cellValue(cell) || "").trim();
    if (key) headers[col] = key;
  });
  const rows = [];
  for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const row = {};
    let hasValue = false;
    for (const [col, key] of Object.entries(headers)) {
      const value = cellValue(sheet.getRow(rowNumber).getCell(Number(col)));
      if (value !== null && value !== undefined && value !== "") hasValue = true;
      row[key] = value;
    }
    if (hasValue) rows.push(row);
  }
  return rows;
}

function forEachUsedCell(sheet, callback) {
  for (let row = 1; row <= sheet.rowCount; row += 1) {
    for (let col = 1; col <= sheet.columnCount; col += 1) {
      const cell = sheet.getCell(row, col);
      if (cell.value !== null && cell.value !== undefined) callback(cell, row, col);
    }
  }
}

function cellValue(cell) {
  const value = cell.value;
  if (typeof value === "number" || typeof value === "string") return value;
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (value && typeof value === "object") {
    if (typeof value.result === "number" || typeof value.result === "string") return value.result;
    if (typeof value.text === "string") return value.text;
    if (typeof value.formula === "string") return `=${value.formula}`;
    if (Array.isArray(value.richText)) return value.richText.map((part) => part.text).join("");
  }
  return value ?? null;
}

function numericCellValue(cell) {
  const value = cellValue(cell);
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value === "string" && value.trim() !== "" && Number.isFinite(Number(value))) return Number(value);
  return null;
}

function isNumericCell(cell) {
  return typeof numericCellValue(cell) === "number";
}

function cellFormula(cell) {
  const value = cell.value;
  if (value && typeof value === "object") return value.formula || value.sharedFormula || "";
  return "";
}

function isNumericConstantFormula(formula) {
  return /^[+-]?\d+(?:\.\d+)?$/.test(String(formula || "").replace(/^=/, "").trim());
}

function valuesMatch(actual, expected, tolerance = 0.75) {
  return typeof actual === "number" && Number.isFinite(actual) && Math.abs(actual - expected) <= tolerance;
}

function normalize(value) {
  return String(value ?? "").toLowerCase().replace(/[^a-z0-9]/g, "");
}

function normalizeCik(value) {
  const digits = String(value ?? "").replace(/\D/g, "");
  return digits ? digits.padStart(10, "0") : "";
}

function normalizeAccessionList(value) {
  return String(value || "")
    .split(/[;,|\s]+/)
    .map((item) => item.trim())
    .filter(Boolean)
    .map((item) => {
      const digits = item.replace(/\D/g, "");
      if (digits.length === 18) return `${digits.slice(0, 10)}-${digits.slice(10, 12)}-${digits.slice(12)}`;
      return item;
    });
}

function accessionCik(accession) {
  const digits = String(accession || "").replace(/\D/g, "");
  return digits.length >= 10 ? digits.slice(0, 10) : "";
}

function round(value) {
  return Math.round(value * 10) / 10;
}

function increment(result, key) {
  result.checks[key] = (result.checks[key] || 0) + 1;
  result.checks.totalChecks = (result.checks.totalChecks || 0) + 1;
}

function logProgress(message) {
  if (!quiet) console.log(`[regression] ${message}`);
}

async function compareAgainstBaseline(currentPayload, baselinePath) {
  const baseline = JSON.parse(await fs.readFile(baselinePath, "utf8"));
  const currentByTicker = new Map(currentPayload.cases.map((item) => [item.ticker, item]));
  const regressions = [];
  for (const oldCase of baseline.cases || []) {
    if (oldCase.status !== "passed") continue;
    const current = currentByTicker.get(oldCase.ticker);
    if (!current || current.status !== "passed") {
      regressions.push(`${oldCase.ticker} previously passed and now ${current?.status || "did not run"}.`);
    }
  }
  if (regressions.length) {
    throw new Error(`Regression baseline comparison failed:\n${regressions.join("\n")}`);
  }
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : error);
  process.exit(1);
});
