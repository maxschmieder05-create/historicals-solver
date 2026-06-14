import fs from "node:fs/promises";
import path from "node:path";
import ExcelJS from "exceljs";

export type GoldModelType =
  | "standard_operating_company"
  | "financial_services_broker_dealer"
  | "bank"
  | "insurance"
  | "reit_real_estate"
  | "utility"
  | "unknown";

export type ModelTypeClassificationSource = "manifest" | "verified_gold_model" | "workbook" | "sec_filing_signals" | "template_profile" | "unknown";

export type ModelTypeClassification = {
  modelType: GoldModelType;
  confidence: "high" | "medium" | "low";
  source: ModelTypeClassificationSource;
  rationale: string[];
};

export type GoldModelManifestEntry = {
  companyName?: string | null;
  ticker?: string | null;
  cik?: string | null;
  modelType?: GoldModelType;
  filePath: string;
  templateVersion?: string | null;
  verifiedByUser?: boolean;
  notes?: string | null;
};

export type GoldModelMetadata = {
  filePath: string;
  fileName: string;
  companyName: string | null;
  ticker: string | null;
  cik: string | null;
  modelType: GoldModelType;
  modelTypeConfidence: "high" | "medium" | "low";
  modelTypeSource: ModelTypeClassificationSource;
  templateVersion: string | null;
  sheetsPresent: string[];
  historicalPeriodsCovered: string[];
  lastModifiedDate: string | null;
  verifiedByUser: boolean;
  usableAsGold: boolean;
  notes: string | null;
  exclusionReasons: string[];
};

export type GoldModelLibraryScan = {
  libraryPath: string | null;
  manifestPath: string | null;
  models: GoldModelMetadata[];
  excluded: GoldModelMetadata[];
  warnings: string[];
};

export type CompanyModelTypeSignals = {
  companyName: string;
  ticker?: string | null;
  cik?: string | null;
  sic?: string | number | null;
  sicDescription?: string | null;
  conceptNames?: string[];
  labels?: string[];
};

const MANIFEST_FILENAMES = ["gold_models.json", "gold_models.csv"];
const GENERATED_METADATA_SHEETS = ["Mapping Audit", "Source Ledger", "Filing Period Map", "Balance Sheet Assignment Ledger"];
const DEFAULT_SCAN_DEPTH = 3;

const MODEL_TYPE_LABELS: Record<GoldModelType, string> = {
  standard_operating_company: "standard operating-company",
  financial_services_broker_dealer: "financial-services / broker-dealer",
  bank: "bank",
  insurance: "insurance",
  reit_real_estate: "REIT / real estate",
  utility: "utility",
  unknown: "unknown"
};

type ManifestRow = Record<string, unknown>;

type ManifestRecord = GoldModelManifestEntry & {
  raw: ManifestRow;
  resolvedFilePath: string;
};

type WorkbookStructure = {
  matches: boolean;
  score: number;
  primarySheetName: string | null;
  reasons: string[];
};

type TextSignals = {
  text: string;
  normalized: string;
};

export function modelTypeDisplayName(modelType: GoldModelType) {
  return MODEL_TYPE_LABELS[modelType];
}

export function normalizeGoldModelType(value: unknown): GoldModelType {
  const normalized = normalize(String(value ?? ""));
  if (!normalized) return "unknown";
  if (/standard|operatingcompany|industrial|corporate/.test(normalized)) return "standard_operating_company";
  if (/financialservices|brokerdealer|brokerage|assetmanager|assetmanagement|capitalmarkets|investmentbank/.test(normalized)) {
    return "financial_services_broker_dealer";
  }
  if (/bank|banking|depository|creditcard|lender|lending/.test(normalized)) return "bank";
  if (/insurance|insurer|underwriting/.test(normalized)) return "insurance";
  if (/reit|realestate|property|realty/.test(normalized)) return "reit_real_estate";
  if (/utility|utilities|regulatedutility|electric|gasutility|waterutility/.test(normalized)) return "utility";
  return "unknown";
}

export async function scanConfiguredGoldModelLibrary(env: NodeJS.ProcessEnv = process.env): Promise<GoldModelLibraryScan> {
  const libraryPath = resolveOptionalPath(env.GOLD_MODEL_LIBRARY_PATH || env.GOLD_MODELS_PATH || "");
  const manifestPath = await resolveManifestPath(libraryPath, env.GOLD_MODEL_MANIFEST_PATH || "");
  const verifiedFolders = parsePathList(env.GOLD_MODEL_VERIFIED_PATHS || env.GOLD_MODEL_VERIFIED_PATH || "");
  if (!libraryPath && !manifestPath) {
    return { libraryPath: null, manifestPath: null, models: [], excluded: [], warnings: [] };
  }
  return scanGoldModelLibrary({ libraryPath, manifestPath, verifiedFolders });
}

export async function scanGoldModelLibrary(options: {
  libraryPath?: string | null;
  manifestPath?: string | null;
  verifiedFolders?: string[];
  maxDepth?: number;
}): Promise<GoldModelLibraryScan> {
  const libraryPath = resolveOptionalPath(options.libraryPath || "");
  const manifestPath = resolveOptionalPath(options.manifestPath || "");
  const verifiedFolders = (options.verifiedFolders ?? []).map((item) => path.resolve(item));
  const warnings: string[] = [];
  const manifestRecords = manifestPath ? await readGoldModelManifest(manifestPath, warnings) : [];
  const manifestByPath = new Map(manifestRecords.map((entry) => [normalizePathKey(entry.resolvedFilePath), entry]));

  const filePaths = new Set<string>();
  if (libraryPath) {
    try {
      for (const filePath of await listWorkbookFiles(libraryPath, options.maxDepth ?? DEFAULT_SCAN_DEPTH)) {
        filePaths.add(filePath);
      }
    } catch (error) {
      warnings.push(`Could not scan gold-model library folder "${libraryPath}": ${errorMessage(error)}.`);
    }
  }
  for (const entry of manifestRecords) {
    filePaths.add(entry.resolvedFilePath);
  }

  const models: GoldModelMetadata[] = [];
  const excluded: GoldModelMetadata[] = [];
  for (const filePath of Array.from(filePaths).sort((a, b) => a.localeCompare(b))) {
    const manifestEntry = manifestByPath.get(normalizePathKey(filePath)) ?? null;
    const metadata = await inspectGoldModelCandidate(filePath, manifestEntry, verifiedFolders);
    if (metadata.exclusionReasons.length) excluded.push(metadata);
    else models.push(metadata);
  }

  return {
    libraryPath,
    manifestPath,
    models,
    excluded,
    warnings
  };
}

export function findVerifiedGoldModelForCompany(scan: GoldModelLibraryScan, company: { cik?: string | null; ticker?: string | null; title?: string | null }) {
  const cik = normalizeCik(company.cik ?? "");
  const ticker = normalize(company.ticker ?? "");
  const title = normalize(company.title ?? "");
  const matches = scan.models
    .filter((model) => model.usableAsGold && model.verifiedByUser)
    .map((model) => {
      let score = 0;
      if (cik && normalizeCik(model.cik ?? "") === cik) score += 100;
      if (ticker && normalize(model.ticker ?? "") === ticker) score += 80;
      if (title && normalize(model.companyName ?? "") === title) score += 50;
      if (title && normalize(model.companyName ?? "").includes(title)) score += 20;
      return { model, score };
    })
    .filter(({ score }) => score > 0)
    .sort((a, b) => b.score - a.score);
  return matches[0]?.model ?? null;
}

export function classifyCompanyModelTypeFromSecSignals(signals: CompanyModelTypeSignals): ModelTypeClassification {
  const text = [
    signals.companyName,
    signals.ticker,
    signals.cik,
    signals.sic,
    signals.sicDescription,
    ...(signals.conceptNames ?? []),
    ...(signals.labels ?? [])
  ]
    .filter((item) => item !== null && item !== undefined && String(item).trim())
    .join(" ");
  return classifyTextModelType(text, "sec_filing_signals", "Company SEC filing signals");
}

export function classifyWorkbookModelType(workbook: ExcelJS.Workbook, templateProfileKind?: string | null): ModelTypeClassification {
  const text = workbookTextSignals(workbook);
  const classification = classifyTextModelType(text.text, "workbook", "Workbook labels and sheet names");
  if (classification.modelType !== "unknown") return classification;
  const profile = normalize(templateProfileKind ?? "");
  if (profile === "financialcompany") {
    return {
      modelType: "financial_services_broker_dealer",
      confidence: "medium",
      source: "template_profile",
      rationale: ["Template profile detection selected the financial-company workbook profile."]
    };
  }
  if (profile === "owlstandard") {
    return {
      modelType: "standard_operating_company",
      confidence: "medium",
      source: "template_profile",
      rationale: ["Template profile detection selected the standard operating-company workbook profile."]
    };
  }
  return classification;
}

export function classifyTemplateModelTypeFromProfile(templateProfileKind: string): ModelTypeClassification {
  const normalized = normalize(templateProfileKind);
  if (normalized === "financialcompany") {
    return {
      modelType: "financial_services_broker_dealer",
      confidence: "medium",
      source: "template_profile",
      rationale: ["Template profile is financial_company."]
    };
  }
  if (normalized === "owlstandard") {
    return {
      modelType: "standard_operating_company",
      confidence: "medium",
      source: "template_profile",
      rationale: ["Template profile is owl_standard."]
    };
  }
  return { modelType: "unknown", confidence: "low", source: "template_profile", rationale: ["Template profile is generic or unknown."] };
}

export function classifyModelTypeFromGoldReference(model: GoldModelMetadata): ModelTypeClassification {
  return {
    modelType: model.modelType,
    confidence: model.modelType === "unknown" ? "low" : "high",
    source: "verified_gold_model",
    rationale: [`Matched verified gold model reference "${model.fileName}".`]
  };
}

export function checkModelTemplateCompatibility(company: ModelTypeClassification, template: ModelTypeClassification) {
  if (company.modelType === "unknown" || template.modelType === "unknown") {
    return { compatible: true, message: null as string | null };
  }
  if (company.modelType === "standard_operating_company" && template.modelType !== "standard_operating_company") {
    return {
      compatible: false,
      message: `This company appears to require a standard operating-company model template. The uploaded ${modelTypeDisplayName(template.modelType)} template is not compatible.`
    };
  }
  if (company.modelType !== "standard_operating_company" && template.modelType === "standard_operating_company") {
    if (isFinancialCompanyModelType(company.modelType)) {
      return {
        compatible: false,
        message: "This company appears to require a financial-services model template. The uploaded standard operating-company template is not compatible."
      };
    }
    return {
      compatible: false,
      message: `This company appears to require a ${modelTypeDisplayName(company.modelType)} model template. The uploaded standard operating-company template is not compatible.`
    };
  }
  if (company.modelType !== "standard_operating_company" && template.modelType !== "standard_operating_company") {
    if (company.modelType === template.modelType) return { compatible: true, message: null as string | null };
    if (isFinancialCompanyModelType(company.modelType) && isFinancialCompanyModelType(template.modelType)) {
      return { compatible: true, message: null as string | null };
    }
    return {
      compatible: false,
      message: `This company appears to require a ${modelTypeDisplayName(company.modelType)} model template. The uploaded ${modelTypeDisplayName(template.modelType)} template is not compatible.`
    };
  }
  return { compatible: true, message: null as string | null };
}

async function inspectGoldModelCandidate(filePath: string, manifestEntry: ManifestRecord | null, verifiedFolders: string[]): Promise<GoldModelMetadata> {
  const fileName = path.basename(filePath);
  const guessed = guessCompletedModelIdentity(fileName);
  const base: GoldModelMetadata = {
    filePath,
    fileName,
    companyName: manifestEntry?.companyName ?? guessed.companyName,
    ticker: manifestEntry?.ticker ?? guessed.ticker,
    cik: manifestEntry?.cik ?? null,
    modelType: manifestEntry?.modelType ?? "unknown",
    modelTypeConfidence: manifestEntry?.modelType && manifestEntry.modelType !== "unknown" ? "high" : "low",
    modelTypeSource: manifestEntry?.modelType && manifestEntry.modelType !== "unknown" ? "manifest" : "unknown",
    templateVersion: manifestEntry?.templateVersion ?? null,
    sheetsPresent: [],
    historicalPeriodsCovered: [],
    lastModifiedDate: null,
    verifiedByUser: manifestEntry?.verifiedByUser ?? isInsideAnyFolder(filePath, verifiedFolders),
    usableAsGold: false,
    notes: manifestEntry?.notes ?? null,
    exclusionReasons: initialFileExclusionReasons(fileName)
  };

  let stat;
  try {
    stat = await fs.stat(filePath);
  } catch {
    base.exclusionReasons.push("File listed for gold-model scan does not exist.");
    return base;
  }
  base.lastModifiedDate = stat.mtime.toISOString();

  if (base.exclusionReasons.length) return base;

  const workbook = new ExcelJS.Workbook();
  try {
    await workbook.xlsx.readFile(filePath);
  } catch (error) {
    base.exclusionReasons.push(`Workbook could not be read as a valid Excel file: ${errorMessage(error)}.`);
    return base;
  }

  base.sheetsPresent = workbook.worksheets.map((sheet) => sheet.name);
  base.historicalPeriodsCovered = detectHistoricalPeriods(workbook);
  base.templateVersion = base.templateVersion ?? detectTemplateVersion(workbook);

  const generatedReasons = generatedWorkbookReasons(workbook);
  base.exclusionReasons.push(...generatedReasons);

  const structure = workbookStructure(workbook);
  if (!structure.matches) {
    base.exclusionReasons.push(`Workbook does not match expected valuation-model structure: ${structure.reasons.join(" ") || `score ${structure.score}`}.`);
  }

  if (!manifestEntry?.modelType || manifestEntry.modelType === "unknown") {
    const classified = classifyWorkbookModelType(workbook);
    base.modelType = classified.modelType;
    base.modelTypeConfidence = classified.confidence;
    base.modelTypeSource = classified.source;
    base.notes = [base.notes, classified.rationale.join(" ")].filter(Boolean).join(" ") || null;
  }

  if (!base.verifiedByUser) {
    base.notes = [
      base.notes,
      "Candidate is not marked verified by a manifest entry or configured verified folder, so it will not be used as a gold model."
    ]
      .filter(Boolean)
      .join(" ");
  }

  base.usableAsGold = base.exclusionReasons.length === 0 && base.verifiedByUser && base.modelType !== "unknown";
  return base;
}

async function resolveManifestPath(libraryPath: string | null, explicitManifestPath: string) {
  const explicit = resolveOptionalPath(explicitManifestPath);
  if (explicit) return explicit;
  if (!libraryPath) return null;
  for (const fileName of MANIFEST_FILENAMES) {
    const candidate = path.join(libraryPath, fileName);
    if (await fileExists(candidate)) return candidate;
  }
  return null;
}

async function readGoldModelManifest(manifestPath: string, warnings: string[]): Promise<ManifestRecord[]> {
  try {
    const text = await fs.readFile(manifestPath, "utf8");
    const rows = manifestPath.toLowerCase().endsWith(".csv") ? parseManifestCsv(text) : parseManifestJson(text);
    return rows.map((row) => manifestEntryFromRow(row, manifestPath)).filter((entry): entry is ManifestRecord => Boolean(entry));
  } catch (error) {
    warnings.push(`Could not read gold-model manifest "${manifestPath}": ${errorMessage(error)}.`);
    return [];
  }
}

function parseManifestJson(text: string): ManifestRow[] {
  const parsed = JSON.parse(text);
  if (Array.isArray(parsed)) return parsed;
  if (Array.isArray(parsed?.models)) return parsed.models;
  if (parsed && typeof parsed === "object") return [parsed];
  return [];
}

function parseManifestCsv(text: string): ManifestRow[] {
  const rows = parseCsvRows(text).filter((row) => row.some((cell) => cell.trim()));
  if (!rows.length) return [];
  const headers = rows[0].map((cell) => cell.trim());
  return rows.slice(1).map((row) => Object.fromEntries(headers.map((header, index) => [header, row[index] ?? ""])));
}

function parseCsvRows(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let cell = "";
  let quoted = false;
  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    const next = text[index + 1];
    if (quoted) {
      if (char === "\"" && next === "\"") {
        cell += "\"";
        index += 1;
      } else if (char === "\"") {
        quoted = false;
      } else {
        cell += char;
      }
      continue;
    }
    if (char === "\"") {
      quoted = true;
      continue;
    }
    if (char === ",") {
      row.push(cell);
      cell = "";
      continue;
    }
    if (char === "\n") {
      row.push(cell);
      rows.push(row);
      row = [];
      cell = "";
      continue;
    }
    if (char !== "\r") cell += char;
  }
  row.push(cell);
  rows.push(row);
  return rows;
}

function manifestEntryFromRow(row: ManifestRow, manifestPath: string): ManifestRecord | null {
  const filePath = stringField(row, "file_path", "filePath", "path");
  if (!filePath) return null;
  const resolvedFilePath = path.isAbsolute(filePath) ? filePath : path.resolve(path.dirname(manifestPath), filePath);
  return {
    raw: row,
    resolvedFilePath,
    filePath,
    companyName: stringField(row, "company_name", "companyName") || null,
    ticker: stringField(row, "ticker") || null,
    cik: normalizeCik(stringField(row, "CIK", "cik") || "") || null,
    modelType: normalizeGoldModelType(stringField(row, "model_type", "modelType")),
    templateVersion: stringField(row, "template_version", "templateVersion") || null,
    verifiedByUser: booleanField(row, "verified_by_user", "verifiedByUser"),
    notes: stringField(row, "notes") || null
  };
}

async function listWorkbookFiles(root: string, maxDepth: number): Promise<string[]> {
  const result: string[] = [];
  async function visit(dir: string, depth: number) {
    const entries = await fs.readdir(dir, { withFileTypes: true });
    for (const entry of entries) {
      const fullPath = path.join(dir, entry.name);
      if (entry.isDirectory()) {
        if (depth >= maxDepth || shouldSkipDirectory(entry.name)) continue;
        await visit(fullPath, depth + 1);
        continue;
      }
      if (entry.isFile() && isWorkbookFileName(entry.name)) result.push(fullPath);
    }
  }
  await visit(root, 0);
  return result;
}

function shouldSkipDirectory(name: string) {
  return name.startsWith(".") || /^(node_modules|tmp|dist|out|build)$/i.test(name);
}

function isWorkbookFileName(fileName: string) {
  return /\.(xlsx|xlsm)$/i.test(fileName);
}

function initialFileExclusionReasons(fileName: string) {
  const reasons: string[] = [];
  if (!isWorkbookFileName(fileName)) reasons.push("File is not an .xlsx or .xlsm workbook.");
  if (/^~\$/.test(fileName)) reasons.push("Temporary Excel lock file.");
  if (/filled/i.test(fileName)) reasons.push("Filename contains \"filled\", which marks a system-generated output.");
  return reasons;
}

function generatedWorkbookReasons(workbook: ExcelJS.Workbook) {
  const reasons: string[] = [];
  const sheetNames = new Set(workbook.worksheets.map((sheet) => normalize(sheet.name)));
  const generatedSheets = GENERATED_METADATA_SHEETS.filter((sheet) => sheetNames.has(normalize(sheet)));
  if (generatedSheets.length) reasons.push(`Workbook contains generated metadata sheet(s): ${generatedSheets.join(", ")}.`);
  const props = [workbook.creator, workbook.lastModifiedBy, workbook.company, workbook.subject, workbook.title]
    .filter((item): item is string => typeof item === "string" && Boolean(item.trim()))
    .join(" ");
  if (/historicalssolver|historicals solver/i.test(props)) reasons.push("Workbook metadata indicates it was produced by Historicals Solver.");
  return reasons;
}

function workbookStructure(workbook: ExcelJS.Workbook): WorkbookStructure {
  const scored = workbook.worksheets
    .map((sheet) => ({ sheetName: sheet.name, ...scoreFinancialSheet(sheet) }))
    .sort((a, b) => b.score - a.score);
  const best = scored[0] ?? { score: 0, reasons: [], sheetName: null };
  return {
    matches: best.score >= 45,
    score: best.score,
    primarySheetName: best.sheetName,
    reasons: best.reasons
  };
}

function scoreFinancialSheet(sheet: ExcelJS.Worksheet) {
  const labels = sheetLabels(sheet, 500, 10);
  const text = labels.join(" ");
  const normalized = normalize(text);
  let score = 0;
  const reasons: string[] = [];
  if (/model|financial|historical|valuation/i.test(sheet.name)) {
    score += 15;
    reasons.push(`"${sheet.name}" looks like a model sheet.`);
  }
  if (normalized.includes("incomestatement")) {
    score += 15;
    reasons.push("Income Statement section found.");
  }
  if (normalized.includes("balancesheet")) {
    score += 15;
    reasons.push("Balance Sheet section found.");
  }
  if (/cashflowstatement|shareholdersequity|debtandinterest|ppe|depreciation|segmentanalysis/.test(normalized)) {
    score += 10;
    reasons.push("Supplemental valuation-model sections found.");
  }
  if (/revenue|netrevenues|grossprofit|operatingincome|netincome|assets|liabilities|equity/i.test(text)) {
    score += 15;
    reasons.push("Core financial statement line items found.");
  }
  const periods = detectPeriodsFromText(text);
  if (periods.length >= 3) {
    score += 10;
    reasons.push("Historical period headers found.");
  }
  return { score, reasons };
}

function guessCompletedModelIdentity(fileName: string) {
  const stem = path.parse(fileName).name;
  const match = stem.match(/^(.+?)\s+\(([A-Z0-9.\-]{1,12})\)_Valuation Workbook(?:\s+\(([^)]*)\))?(?:\s+\(\d+\))*$/i);
  if (!match) return { companyName: null, ticker: null, workbookDate: null };
  return {
    companyName: match[1].trim(),
    ticker: match[2].trim().toUpperCase(),
    workbookDate: match[3]?.trim() ?? null
  };
}

function detectTemplateVersion(workbook: ExcelJS.Workbook) {
  for (const sheet of workbook.worksheets) {
    for (let row = 1; row <= Math.min(sheet.rowCount, 120); row += 1) {
      for (let col = 1; col <= Math.min(sheet.columnCount, 16); col += 1) {
        const text = cellDisplay(sheet.getCell(row, col)).trim();
        if (!/^(template\s*)?version$/i.test(text) && !/^template version$/i.test(text)) continue;
        for (const candidateCol of [col + 1, col + 2]) {
          const value = cellDisplay(sheet.getCell(row, candidateCol)).trim();
          if (value) return value;
        }
      }
    }
  }
  return null;
}

function detectHistoricalPeriods(workbook: ExcelJS.Workbook) {
  const periods = new Set<string>();
  for (const sheet of workbook.worksheets) {
    for (let row = 1; row <= Math.min(sheet.rowCount, 60); row += 1) {
      for (let col = 1; col <= Math.min(sheet.columnCount, 80); col += 1) {
        for (const period of detectPeriodsFromText(cellDisplay(sheet.getCell(row, col)))) periods.add(period);
      }
    }
  }
  return Array.from(periods).sort(comparePeriodLabels);
}

function detectPeriodsFromText(text: string) {
  const periods = new Set<string>();
  const pattern = /\b(?:Q([1-4])\s*['’]?\s*(\d{2}|\d{4})|([1-4])Q\s*['’]?\s*(\d{2}|\d{4})|FY\s*['’]?\s*(\d{2}|\d{4}))\b/gi;
  for (const match of text.matchAll(pattern)) {
    if (match[1] && match[2]) periods.add(`${match[1]}Q${match[2].slice(-2)}`);
    if (match[3] && match[4]) periods.add(`${match[3]}Q${match[4].slice(-2)}`);
    if (match[5]) periods.add(`FY${match[5].slice(-2)}`);
  }
  return Array.from(periods);
}

function comparePeriodLabels(a: string, b: string) {
  const yearA = Number(a.slice(-2));
  const yearB = Number(b.slice(-2));
  if (yearA !== yearB) return yearA - yearB;
  return periodOrder(a) - periodOrder(b);
}

function periodOrder(period: string) {
  if (period.startsWith("FY")) return 5;
  return Number(period[0]) || 0;
}

function classifyTextModelType(text: string, source: ModelTypeClassificationSource, context: string): ModelTypeClassification {
  const signals = textSignals(text);
  const scores: Record<Exclude<GoldModelType, "unknown">, { score: number; reasons: string[] }> = {
    standard_operating_company: { score: 0, reasons: [] },
    financial_services_broker_dealer: { score: 0, reasons: [] },
    bank: { score: 0, reasons: [] },
    insurance: { score: 0, reasons: [] },
    reit_real_estate: { score: 0, reasons: [] },
    utility: { score: 0, reasons: [] }
  };
  addScore(scores.financial_services_broker_dealer, signals, 6, "broker-dealer / capital-markets terms", [
    "brokerdealer",
    "investmentbanking",
    "capitalmarkets",
    "assetmanagement",
    "principaltransactions",
    "tradingrevenue",
    "brokerage",
    "securitiesborrowed",
    "securitiesloaned",
    "financialinstrumentsowned",
    "netrevenues",
    "revenuesnetofinterestexpense"
  ]);
  addScore(scores.bank, signals, 5, "banking terms", [
    "bank",
    "bancorp",
    "deposits",
    "loansreceivable",
    "loansheldforinvestment",
    "netinterestincome",
    "interestincome",
    "provisionforcreditlosses",
    "allowanceforcreditlosses"
  ]);
  addScore(scores.insurance, signals, 6, "insurance terms", [
    "insurance",
    "premiums",
    "policyholder",
    "claims",
    "lossadjustment",
    "underwriting",
    "reinsurance",
    "benefitsclaims"
  ]);
  addScore(scores.reit_real_estate, signals, 6, "REIT / real estate terms", [
    "reit",
    "realestate",
    "rentalrevenues",
    "realestateinvestments",
    "fundsfromoperations",
    "sameproperty",
    "netoperatingincome",
    "tenant"
  ]);
  addScore(scores.utility, signals, 6, "regulated utility terms", [
    "utility",
    "electricutility",
    "gasutility",
    "regulated",
    "ratebase",
    "kilowatthours",
    "megawatthours",
    "customersales"
  ]);
  addScore(scores.standard_operating_company, signals, 3, "standard operating-company terms", [
    "costofrevenue",
    "costofgoodssold",
    "grossprofit",
    "inventory",
    "accountsreceivable",
    "workingcapital",
    "propertyplantequipment",
    "ppe",
    "net sales",
    "productrevenue"
  ]);

  const ranked = Object.entries(scores)
    .map(([modelType, item]) => ({ modelType: modelType as Exclude<GoldModelType, "unknown">, ...item }))
    .sort((a, b) => b.score - a.score);
  const best = ranked[0];
  if (!best || best.score <= 0) {
    return { modelType: "unknown", confidence: "low", source, rationale: [`${context} did not contain enough model-type signals.`] };
  }
  const second = ranked[1];
  const confidence = best.score >= 12 || best.score - (second?.score ?? 0) >= 5 ? "high" : best.score >= 5 ? "medium" : "low";
  if (confidence === "low" && best.modelType !== "standard_operating_company") {
    return { modelType: "unknown", confidence: "low", source, rationale: [`${context} had weak ${modelTypeDisplayName(best.modelType)} signals only.`] };
  }
  return {
    modelType: best.modelType,
    confidence,
    source,
    rationale: [`${context} classified as ${modelTypeDisplayName(best.modelType)} because it includes ${best.reasons.join(", ")}.`]
  };
}

function addScore(target: { score: number; reasons: string[] }, signals: TextSignals, points: number, reason: string, patterns: string[]) {
  let hits = 0;
  for (const pattern of patterns) {
    const normalizedPattern = normalize(pattern);
    if (signals.normalized.includes(normalizedPattern) || signals.text.toLowerCase().includes(pattern.toLowerCase())) hits += 1;
  }
  if (!hits) return;
  target.score += points + Math.min(hits, 6);
  target.reasons.push(reason);
}

function workbookTextSignals(workbook: ExcelJS.Workbook) {
  const chunks: string[] = [];
  chunks.push(...workbook.worksheets.map((sheet) => sheet.name));
  for (const sheet of workbook.worksheets) {
    chunks.push(...sheetLabels(sheet, 500, 12));
  }
  return textSignals(chunks.join(" "));
}

function textSignals(text: string): TextSignals {
  return { text, normalized: normalize(text) };
}

function sheetLabels(sheet: ExcelJS.Worksheet, maxRows: number, maxColumns: number) {
  const labels: string[] = [];
  for (let row = 1; row <= Math.min(sheet.rowCount, maxRows); row += 1) {
    for (let col = 1; col <= Math.min(sheet.columnCount, maxColumns); col += 1) {
      const text = cellDisplay(sheet.getCell(row, col)).trim();
      if (text && /[A-Za-z]{3,}/.test(text)) labels.push(text);
    }
  }
  return labels;
}

function isFinancialCompanyModelType(modelType: GoldModelType) {
  return modelType === "financial_services_broker_dealer" || modelType === "bank" || modelType === "insurance";
}

function cellDisplay(cell: ExcelJS.Cell) {
  const value = cell.value;
  if (typeof value === "string") return value;
  if (typeof value === "number") return String(value);
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (value && typeof value === "object" && "result" in value && value.result !== undefined && value.result !== null) {
    return String(value.result);
  }
  if (value && typeof value === "object" && "richText" in value && Array.isArray(value.richText)) {
    return value.richText.map((part) => part.text).join("");
  }
  return "";
}

function stringField(row: ManifestRow, ...keys: string[]) {
  for (const key of keys) {
    const value = row[key] ?? row[key.toLowerCase()] ?? row[toSnakeCase(key)];
    if (value !== null && value !== undefined && String(value).trim()) return String(value).trim();
  }
  return "";
}

function booleanField(row: ManifestRow, ...keys: string[]) {
  const value = stringField(row, ...keys);
  if (!value) return false;
  return /^(true|1|yes|y)$/i.test(value);
}

function toSnakeCase(value: string) {
  return value.replace(/[A-Z]/g, (match) => `_${match.toLowerCase()}`).replace(/^_/, "");
}

function parsePathList(value: string) {
  return value
    .split(path.delimiter)
    .flatMap((item) => item.split(","))
    .map((item) => item.trim())
    .filter(Boolean)
    .map((item) => path.resolve(item));
}

function resolveOptionalPath(value: string) {
  const trimmed = value.trim();
  return trimmed ? path.resolve(trimmed) : null;
}

async function fileExists(filePath: string) {
  try {
    await fs.access(filePath);
    return true;
  } catch {
    return false;
  }
}

function isInsideAnyFolder(filePath: string, folders: string[]) {
  const resolved = path.resolve(filePath);
  return folders.some((folder) => {
    const relative = path.relative(folder, resolved);
    return relative === "" || (!relative.startsWith("..") && !path.isAbsolute(relative));
  });
}

function normalizePathKey(filePath: string) {
  return path.resolve(filePath);
}

function normalizeCik(value: string) {
  const digits = value.replace(/\D/g, "");
  return digits ? digits.padStart(10, "0") : "";
}

function normalize(input: string) {
  return input.toLowerCase().replace(/[^a-z0-9]/g, "");
}

function errorMessage(error: unknown) {
  return error instanceof Error ? error.message : String(error);
}
