const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const Module = require("node:module");
const ts = require("typescript");
const JSZip = require("jszip");

const repoRoot = path.resolve(__dirname, "..");

function registerTypeScriptRequire() {
  if (require.extensions[".ts"]) return;
  require.extensions[".ts"] = (mod, file) => {
    mod._compile(
      ts.transpileModule(fs.readFileSync(file, "utf8"), {
        compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020, esModuleInterop: true }
      }).outputText,
      file
    );
  };
}

function loadService() {
  registerTypeScriptRequire();
  const file = path.join(repoRoot, "server", "fill-model", "fill-model-service.ts");
  const mod = new Module(file, module);
  mod.filename = file;
  mod.paths = Module._nodeModulePaths(path.dirname(file));
  mod._compile(
    ts.transpileModule(fs.readFileSync(file, "utf8"), {
      compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020, esModuleInterop: true }
    }).outputText,
    file
  );
  return mod.exports;
}

async function packageBuffer(files) {
  const zip = new JSZip();
  for (const [name, value] of Object.entries(files)) zip.file(name, value);
  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}

(async () => {
  const { __fillModelServiceTestHooks } = loadService();
  const validate = __fillModelServiceTestHooks.validateWorkbookPackageSafety;
  const safe = await packageBuffer({
    "[Content_Types].xml": '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>',
    "xl/workbook.xml": "<workbook/>"
  });
  await validate(safe, "test-debug.log");

  const renamedMacro = await packageBuffer({
    "[Content_Types].xml": '<Types><Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>',
    "xl/workbook.xml": "<workbook/>",
    "xl/vbaProject.bin": Buffer.from([1, 2, 3])
  });
  await assert.rejects(
    () => validate(renamedMacro, "test-debug.log"),
    /contains VBA, ActiveX, an embedded OLE object, or a digital signature/i,
    "macro content must be rejected by package contents even when the filename says .xlsx"
  );

  const extendedConditionalFormatting = await packageBuffer({
    "[Content_Types].xml": '<ct:Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types"><ct:Override PartName="/unusual/worksheet-part.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></ct:Types>',
    "xl/workbook.xml": "<workbook/>",
    "unusual/worksheet-part.xml": '<worksheet xmlns:futureCf="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><extLst><futureCf:cfRule type="expression"/></extLst></worksheet>'
  });
  await assert.rejects(
    () => validate(extendedConditionalFormatting, "test-debug.log"),
    /extended Excel conditional-formatting rules/i,
    "extended conditional formatting that ExcelJS cannot safely serialize must fail before workbook processing"
  );

  const prefixedAlternateParts = await packageBuffer({
    "[Content_Types].xml": `<?xml version="1.0"?><ct:Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types">
      <ct:Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <ct:Default Extension="xml" ContentType="application/xml"/>
      <ct:Override PartName="/custom/book.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
      <ct:Override PartName="/data/odd-sheet.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
      <ct:Override PartName="/strings/text-pool.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
      <ct:Override PartName="/calcs/stale.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>
    </ct:Types>`,
    "_rels/.rels": `<?xml version="1.0"?><pkg:Relationships xmlns:pkg="http://schemas.openxmlformats.org/package/2006/relationships">
      <pkg:Relationship Id="office" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="custom/book.xml"/>
    </pkg:Relationships>`,
    "custom/book.xml": `<?xml version="1.0"?><ss:workbook xmlns:ss="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:link="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <ss:sheets><ss:sheet name="Model" sheetId="1" link:id="sheetRel"/></ss:sheets><ss:extLst/>
    </ss:workbook>`,
    "custom/_rels/book.xml.rels": `<?xml version="1.0"?><pkg:Relationships xmlns:pkg="http://schemas.openxmlformats.org/package/2006/relationships">
      <pkg:Relationship Id="sheetRel" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="../data/odd-sheet.xml"/>
      <pkg:Relationship Id="stringsRel" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="../strings/text-pool.xml"/>
      <pkg:Relationship Id="calcRel" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="../calcs/stale.xml"/>
    </pkg:Relationships>`,
    "data/odd-sheet.xml": `<?xml version="1.0"?><ss:worksheet xmlns:ss="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><ss:sheetData><ss:row r="1">
      <ss:c r="A1" t="s"><ss:v>0</ss:v></ss:c><ss:c r="B1"><ss:f>1+1</ss:f><ss:v>2</ss:v></ss:c>
    </ss:row></ss:sheetData></ss:worksheet>`,
    "strings/text-pool.xml": `<?xml version="1.0"?><ss:sst xmlns:ss="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><ss:si><ss:t>=SUM(1,2)</ss:t></ss:si></ss:sst>`,
    "calcs/stale.xml": `<?xml version="1.0"?><ss:calcChain xmlns:ss="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>`
  });
  const recalculatedAlternateParts = await __fillModelServiceTestHooks.enforceXlsxAutomaticCalculation(
    prefixedAlternateParts,
    ["Model"]
  );
  const recalculatedZip = await JSZip.loadAsync(recalculatedAlternateParts);
  const recalculatedWorkbookXml = await recalculatedZip.file("custom/book.xml").async("string");
  const recalculatedSheetXml = await recalculatedZip.file("data/odd-sheet.xml").async("string");
  const recalculatedWorkbookRels = await recalculatedZip.file("custom/_rels/book.xml.rels").async("string");
  const recalculatedContentTypes = await recalculatedZip.file("[Content_Types].xml").async("string");
  assert.match(recalculatedWorkbookXml, /<ss:calcPr\b[^>]*calcMode="auto"[^>]*forceFullCalc="1"/i);
  assert.match(recalculatedSheetXml, /<ss:c\b[^>]*r="A1"[^>]*><ss:f\b[^>]*ca="1">SUM\(1,2\)<\/ss:f><\/ss:c>/i);
  assert.match(recalculatedSheetXml, /<ss:c\b[^>]*r="B1"[^>]*>[\s\S]*?<ss:f\b[^>]*ca="1">1\+1<\/ss:f>/i);
  assert.equal(recalculatedZip.file("calcs/stale.xml"), null, "relationship-targeted calcChain must be removed regardless of part name");
  assert.doesNotMatch(recalculatedWorkbookRels, /\/calcChain["']/i);
  assert.doesNotMatch(recalculatedContentTypes, /calcChain/i);

  const evercoreTemplate = path.join(repoRoot, "github", "templates", "fig", "Evercore, Inc. (EVR)_Valuation Workbook (08-May-2026).xlsx");
  if (fs.existsSync(evercoreTemplate)) {
    await assert.rejects(
      () => validate(fs.readFileSync(evercoreTemplate), "test-debug.log"),
      /extended Excel conditional-formatting rules/i,
      "the shipped EVR template must fail cleanly instead of crashing during ExcelJS serialization"
    );
  }

  const kkrTemplate = path.join(repoRoot, "github", "templates", "fig", "KKR & Co., Inc. (KKR)_Valuation Workbook (10-Mar-2026).xlsx");
  if (fs.existsSync(kkrTemplate)) {
    await assert.rejects(
      () => validate(fs.readFileSync(kkrTemplate), "test-debug.log"),
      /VBA, ActiveX, an embedded OLE object, or a digital signature/i,
      "the shipped KKR workbook must reject its ActiveX control before any lossy round trip"
    );
  }

  const previousMaxEntries = process.env.FILL_MODEL_MAX_XLSX_ENTRIES;
  process.env.FILL_MODEL_MAX_XLSX_ENTRIES = "2";
  const tooManyEntries = await packageBuffer({
    "[Content_Types].xml": "<Types/>",
    "xl/workbook.xml": "<workbook/>",
    "xl/worksheets/sheet1.xml": "<worksheet/>"
  });
  await assert.rejects(() => validate(tooManyEntries, "test-debug.log"), /too many files/i);
  if (previousMaxEntries === undefined) delete process.env.FILL_MODEL_MAX_XLSX_ENTRIES;
  else process.env.FILL_MODEL_MAX_XLSX_ENTRIES = previousMaxEntries;

  const previousMaxEntryBytes = process.env.FILL_MODEL_MAX_XLSX_ENTRY_BYTES;
  process.env.FILL_MODEL_MAX_XLSX_ENTRY_BYTES = "128";
  const oversizedUncompressedEntry = await packageBuffer({
    "[Content_Types].xml": "<Types/>",
    "xl/workbook.xml": "<workbook/>",
    "xl/worksheets/sheet1.xml": `<worksheet>${"A".repeat(1_024)}</worksheet>`
  });
  await assert.rejects(
    () => validate(oversizedUncompressedEntry, "test-debug.log"),
    /entry .* exceeds the uncompressed-size limit/i,
    "a highly compressed workbook part must be rejected by uncompressed size before ExcelJS loads it"
  );
  if (previousMaxEntryBytes === undefined) delete process.env.FILL_MODEL_MAX_XLSX_ENTRY_BYTES;
  else process.env.FILL_MODEL_MAX_XLSX_ENTRY_BYTES = previousMaxEntryBytes;

  assert.equal(
    __fillModelServiceTestHooks.sourceLooksLikeReportedCashAggregate({
      concept: "CashAndCashEquivalentsAtCarryingValue",
      label: "Cash and cash equivalents",
      value: 100
    }),
    true,
    "the standard cash-and-equivalents total must be treated as an aggregate when component cash rows are also present"
  );

  console.log("Workbook package safety and cash aggregate guards passed.");
})().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
