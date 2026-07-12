#!/usr/bin/env node

const assert = require("node:assert/strict");
const fs = require("node:fs/promises");
const Module = require("node:module");
const path = require("node:path");
const ExcelJS = require("exceljs");
const JSZip = require("jszip");
const ts = require("typescript");

const ROOT = path.resolve(__dirname, "..");
const MODULE_PATH = path.join(ROOT, "server/fill-model/ooxml-preservation.ts");
const CASES = [
  {
    name: "general integrated model",
    file: path.join(
      ROOT,
      "github/templates/general/Owl Fund Integrated Model Template (03-Sep-2025)_v25 (3).xlsx",
    ),
    requirePrinterSettings: false,
    requireExactStructureRoundTrip: true,
    mergeOrderNormalizationSheets: [
      "Cover",
      "Model Output",
      "Comparables",
      "Valuation",
      "Multiples Data",
      "Benchmark",
    ],
  },
  {
    name: "financial-services valuation model",
    file: path.join(
      ROOT,
      "github/templates/fig/Jefferies Financial Group Inc. (JEF)_Valuation Workbook (10-Mar-2026) (1).xlsx",
    ),
    requirePrinterSettings: true,
  },
];

function xmlAttribute(attrs, name) {
  const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return attrs.match(new RegExp(`(?:^|\\s)${escaped}\\s*=\\s*(["'])(.*?)\\1`, "i"))?.[2];
}

function replaceXmlAttribute(element, name, value) {
  const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return element.replace(
    new RegExp(`(\\b${escaped}\\s*=\\s*)(["'])(.*?)\\2`, "i"),
    (_match, prefix, quote) => `${prefix}${quote}${value}${quote}`,
  );
}

function normalizePartPath(partName) {
  return path.posix.normalize(partName.replace(/^\/+/, ""));
}

function resolveTarget(ownerPart, target) {
  if (target.startsWith("/")) return normalizePartPath(target);
  return normalizePartPath(path.posix.join(path.posix.dirname(ownerPart), target));
}

function relativeTarget(ownerPart, targetPart) {
  return path.posix.relative(path.posix.dirname(ownerPart), targetPart);
}

function prefixRelationshipsXml(xml) {
  return xml
    .replace(/<Relationships\b/, '<pkg:Relationships xmlns:pkg="http://schemas.openxmlformats.org/package/2006/relationships"')
    .replace(/\s+xmlns="http:\/\/schemas\.openxmlformats\.org\/package\/2006\/relationships"/, "")
    .replace(/<Relationship\b/g, "<pkg:Relationship")
    .replace(/<\/Relationships>/g, "</pkg:Relationships>");
}

function prefixContentTypesXml(xml) {
  return xml
    .replace(/<Types\b/, '<ct:Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types"')
    .replace(/\s+xmlns="http:\/\/schemas\.openxmlformats\.org\/package\/2006\/content-types"/, "")
    .replace(/<(Default|Override)\b/g, "<ct:$1")
    .replace(/<\/Types>/g, "</ct:Types>");
}

async function exerciseAlternatePartNamesAndPrefixes(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const oldWorkbookPart = "xl/workbook.xml";
  const newWorkbookPart = "custom/package/book-main.xml";
  const oldWorkbookRelsPart = "xl/_rels/workbook.xml.rels";
  const newWorkbookRelsPart = "custom/package/_rels/book-main.xml.rels";

  let workbookXml = await zip.file(oldWorkbookPart).async("string");
  workbookXml = workbookXml
    .replace(/\bxmlns:r=/g, "xmlns:officeRel=")
    .replace(/\br:id=/g, "officeRel:id=")
    .replace(/\bxmlns:mc=/g, "xmlns:compat=")
    .replace(/\bmc:/g, "compat:");
  zip.file(newWorkbookPart, workbookXml);
  zip.remove(oldWorkbookPart);

  let workbookRelsXml = await zip.file(oldWorkbookRelsPart).async("string");
  workbookRelsXml = workbookRelsXml.replace(/<(?:[A-Za-z_][\w.-]*:)?Relationship\b[^>]*\/>/g, (element) => {
    const attrs = element.slice(element.indexOf(" "));
    if (/^external$/i.test(xmlAttribute(attrs, "TargetMode") ?? "")) return element;
    const target = xmlAttribute(attrs, "Target");
    if (!target) return element;
    return replaceXmlAttribute(
      element,
      "Target",
      relativeTarget(newWorkbookPart, resolveTarget(oldWorkbookPart, target)),
    );
  });
  zip.file(newWorkbookRelsPart, prefixRelationshipsXml(workbookRelsXml));
  zip.remove(oldWorkbookRelsPart);

  const rootRelsFile = zip.file("_rels/.rels");
  let rootRelsXml = await rootRelsFile.async("string");
  rootRelsXml = rootRelsXml.replace(/<(?:[A-Za-z_][\w.-]*:)?Relationship\b[^>]*\/>/g, (element) => {
    const attrs = element.slice(element.indexOf(" "));
    const type = xmlAttribute(attrs, "Type") ?? "";
    return /\/officeDocument$/i.test(type)
      ? replaceXmlAttribute(element, "Target", newWorkbookPart)
      : element;
  });
  zip.file("_rels/.rels", prefixRelationshipsXml(rootRelsXml));

  const contentTypesFile = zip.file("[Content_Types].xml");
  let contentTypesXml = await contentTypesFile.async("string");
  contentTypesXml = contentTypesXml.replace(/<(?:[A-Za-z_][\w.-]*:)?Override\b[^>]*\/>/g, (element) =>
    normalizePartPath(xmlAttribute(element, "PartName") ?? "") === oldWorkbookPart
      ? replaceXmlAttribute(element, "PartName", `/${newWorkbookPart}`)
      : element,
  );
  zip.file("[Content_Types].xml", prefixContentTypesXml(contentTypesXml));

  for (const partName of Object.keys(zip.files).filter((name) => /^xl\/worksheets\/[^/]+\.xml$/i.test(name))) {
    const worksheetFile = zip.file(partName);
    const worksheetXml = await worksheetFile.async("string");
    zip.file(
      partName,
      worksheetXml
        .replace(/\bxmlns:r=/g, "xmlns:officeRel=")
        .replace(/\br:id=/g, "officeRel:id=")
        .replace(/\bxmlns:mc=/g, "xmlns:compat=")
        .replace(/\bmc:/g, "compat:"),
    );
  }

  return zip.generateAsync({ type: "nodebuffer" });
}

function loadTypeScriptModule(filename) {
  const source = require("node:fs").readFileSync(filename, "utf8");
  const compiled = ts.transpileModule(source, {
    fileName: filename,
    reportDiagnostics: true,
    compilerOptions: {
      target: ts.ScriptTarget.ES2020,
      module: ts.ModuleKind.CommonJS,
      moduleResolution: ts.ModuleResolutionKind.NodeJs,
      esModuleInterop: true,
      strict: true,
    },
  });
  const errors = (compiled.diagnostics || []).filter(
    (diagnostic) => diagnostic.category === ts.DiagnosticCategory.Error,
  );
  if (errors.length > 0) {
    throw new Error(
      errors
        .map((diagnostic) => ts.flattenDiagnosticMessageText(diagnostic.messageText, "\n"))
        .join("\n"),
    );
  }
  const loaded = new Module(filename, module);
  loaded.filename = filename;
  loaded.paths = Module._nodeModulePaths(path.dirname(filename));
  loaded._compile(compiled.outputText, filename);
  return loaded.exports;
}

function totalDrawingGraphs(manifest) {
  return Object.values(manifest.sheets).reduce(
    (total, sheet) => total + sheet.drawingGraphHashes.length,
    0,
  );
}

function totalHyperlinks(manifest) {
  return Object.values(manifest.sheets).reduce(
    (total, sheet) => total + sheet.hyperlinkCount,
    0,
  );
}

function totalPrinterSettings(manifest) {
  return Object.values(manifest.sheets).reduce(
    (total, sheet) => total + sheet.printerSettingsGraphHashes.length,
    0,
  );
}

function totalCommentMetadataGraphs(manifest) {
  return Object.values(manifest.sheets).reduce(
    (total, sheet) => total + sheet.commentMetadataGraphHashes.length,
    0,
  );
}

function sortFingerprintValue(value) {
  if (Array.isArray(value)) return value.map(sortFingerprintValue);
  if (value instanceof Date) return value.toISOString();
  if (!value || typeof value !== "object") return value ?? null;
  return Object.fromEntries(
    Object.entries(value)
      .filter(([, item]) => typeof item !== "function" && item !== undefined)
      .sort(([left], [right]) => left.localeCompare(right))
      .map(([key, item]) => [key, sortFingerprintValue(item)]),
  );
}

function stableFingerprint(value) {
  return JSON.stringify(sortFingerprintValue(value));
}

function externalDefinedNameRanges(workbook) {
  return (workbook.definedNames.model ?? []).flatMap((item) =>
    (item.ranges ?? [])
      .filter(
        (range) =>
          /\[[^\]]+\.(?:xlsx?|xlsm|xlsb)\][^!]*!/i.test(range) ||
          /(?:^|[=+\-*/,(])\s*'?\[[^\]]+\][^'!]*'?!/i.test(range),
      )
      .map((range) => `${item.name} -> ${range}`),
  );
}

function worksheetStructureFingerprint(sheet) {
  const columns = Array.from({ length: sheet.columnCount }, (_unused, index) => {
    const column = sheet.getColumn(index + 1);
    return {
      number: index + 1,
      width: column.width ?? null,
      hidden: Boolean(column.hidden),
      outlineLevel: column.outlineLevel ?? 0,
      collapsed: Boolean(column.collapsed),
      style: column.style ?? {},
    };
  });
  const rows = Array.from({ length: sheet.rowCount }, (_unused, index) => {
    const row = sheet.getRow(index + 1);
    return {
      number: index + 1,
      height: row.height ?? null,
      hidden: Boolean(row.hidden),
      outlineLevel: row.outlineLevel ?? 0,
      collapsed: Boolean(row.collapsed),
      style: row.style ?? {},
    };
  });
  return stableFingerprint({
    name: sheet.name,
    state: sheet.state,
    rowCount: sheet.rowCount,
    columnCount: sheet.columnCount,
    properties: sheet.properties,
    pageSetup: sheet.pageSetup,
    views: sheet.views,
    autoFilter: sheet.autoFilter ?? null,
    merges: sheet.model?.merges ?? [],
    dataValidations: sheet.dataValidations?.model ?? {},
    columns,
    rows,
  });
}

async function runCase(api, testCase) {
  const excelJsOriginal = await fs.readFile(testCase.file);
  let original = excelJsOriginal;
  if (testCase.exerciseAlternatePartNamesAndPrefixes) {
    original = await exerciseAlternatePartNamesAndPrefixes(original);
  }
  const originalManifest = await api.captureOoxmlFeatureManifest(original);
  assert.ok(totalDrawingGraphs(originalManifest) > 0, `${testCase.name}: fixture needs drawings`);
  assert.ok(totalHyperlinks(originalManifest) > 0, `${testCase.name}: fixture needs hyperlinks`);
  assert.ok(originalManifest.definedNameCount > 0, `${testCase.name}: fixture needs defined names`);
  assert.ok(
    originalManifest.externalReferenceCount > 0,
    `${testCase.name}: fixture needs external references`,
  );
  assert.ok(originalManifest.workbookExtLstHash, `${testCase.name}: fixture needs workbook extLst`);
  assert.ok(
    originalManifest.customPropertiesGraphHashes.length > 0,
    `${testCase.name}: fixture needs custom properties`,
  );
  if (testCase.requirePrinterSettings) {
    assert.ok(totalPrinterSettings(originalManifest) > 0, `${testCase.name}: fixture needs printer settings`);
  }
  if (testCase.name === "general integrated model") {
    assert.ok(totalCommentMetadataGraphs(originalManifest) > 0, `${testCase.name}: fixture needs original comment metadata`);
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(excelJsOriginal);
  assert.ok(
    externalDefinedNameRanges(workbook).length > 0,
    `${testCase.name}: fixture needs legacy external-workbook defined names`,
  );
  const originalStructureFingerprints = new Map(
    workbook.worksheets.map((sheet) => [sheet.name, worksheetStructureFingerprint(sheet)]),
  );
  const cover = workbook.getWorksheet("Cover");
  assert.ok(cover, `${testCase.name}: missing Cover sheet`);
  const commentCell = cover.getCell("Z200");
  const commentSentinel = `OOXML preservation check: ${testCase.name}`;
  commentCell.note = commentSentinel;
  if (!workbook.getWorksheet("Source Ledger")) {
    const audit = workbook.addWorksheet("Source Ledger");
    audit.addRow(["workbook sheet", "model row", "model row label", "model column"]);
  }
  const generated = Buffer.from(await workbook.xlsx.writeBuffer());
  if (testCase.mergeOrderNormalizationSheets) {
    const generatedReopened = new ExcelJS.Workbook();
    await generatedReopened.xlsx.load(generated);
    const mergeOrderChangedSheets = workbook.worksheets
      .filter((sheet) => {
        const generatedSheet = generatedReopened.getWorksheet(sheet.name);
        return stableFingerprint(sheet.model?.merges ?? []) !== stableFingerprint(generatedSheet?.model?.merges ?? []);
      })
      .map((sheet) => sheet.name);
    assert.deepEqual(
      mergeOrderChangedSheets,
      testCase.mergeOrderNormalizationSheets,
      `${testCase.name}: fixture no longer exercises the ExcelJS merged-range order rewrite`,
    );
  }
  const generatedManifest = await api.captureOoxmlFeatureManifest(generated);
  assert.notDeepEqual(
    generatedManifest,
    originalManifest,
    `${testCase.name}: ExcelJS fixture unexpectedly preserved every opaque OOXML feature`,
  );

  const restored = await api.restoreOoxmlPreservedFeatures(original, generated);
  await api.validateOoxmlPackage(restored);
  const restoredManifest = await api.captureOoxmlFeatureManifest(restored);
  api.validateOoxmlFeatureInvariants(originalManifest, restoredManifest);

  const restoredZip = await JSZip.loadAsync(restored);
  assert.equal(restoredZip.file("xl/calcChain.xml"), null, `${testCase.name}: calcChain survived`);
  assert.ok(
    Object.keys(restoredZip.files).some((partName) => partName.includes(".codex-preserved-")),
    `${testCase.name}: collision-safe part allocation was not exercised`,
  );

  const reopened = new ExcelJS.Workbook();
  await reopened.xlsx.load(restored);
  assert.deepEqual(
    externalDefinedNameRanges(reopened),
    [],
    `${testCase.name}: OOXML restoration reintroduced external-workbook defined names`,
  );
  const restoredComment = reopened.getWorksheet("Cover").getCell("Z200").note;
  assert.ok(
    JSON.stringify(restoredComment).includes(commentSentinel),
    `${testCase.name}: generated comment was overwritten by original package data`,
  );
  if (testCase.requireExactStructureRoundTrip) {
    for (const [sheetName, originalFingerprint] of originalStructureFingerprints) {
      const restoredSheet = reopened.getWorksheet(sheetName);
      assert.ok(restoredSheet, `${testCase.name}: restored workbook lost ${sheetName}`);
      assert.equal(
        worksheetStructureFingerprint(restoredSheet),
        originalFingerprint,
        `${testCase.name}: ${sheetName} structure changed after ExcelJS write and OOXML restore`,
      );
    }
  }

  console.log(
    [
      `PASS ${testCase.name}`,
      `sheets=${originalManifest.sheetNames.length}`,
      `drawings=${totalDrawingGraphs(originalManifest)}`,
      `hyperlinks=${totalHyperlinks(originalManifest)}`,
      `printerSettings=${totalPrinterSettings(originalManifest)}`,
      `commentMetadata=${totalCommentMetadataGraphs(originalManifest)}`,
      `definedNames=${originalManifest.definedNameCount}`,
      `externalReferences=${originalManifest.externalReferenceCount}`,
    ].join(" | "),
  );
  return restored;
}

async function expectInvalidPackage(api, validBuffer, mutate, expectedError) {
  const zip = await JSZip.loadAsync(validBuffer);
  await mutate(zip);
  const invalidBuffer = await zip.generateAsync({ type: "nodebuffer" });
  await assert.rejects(() => api.validateOoxmlPackage(invalidBuffer), expectedError);
}

async function runValidatorNegativeChecks(api, restored) {
  await expectInvalidPackage(
    api,
    restored,
    async (zip) => {
      const file = zip.file("xl/workbook.xml");
      const xml = await file.async("string");
      zip.file(
        "xl/workbook.xml",
        xml.replace(
          /(<\/(?:[A-Za-z_][\w.-]*:)?definedNames\s*>)/i,
          (_match, closingTag) =>
            `<definedName name="LegacyNamedRange">'[Legacy.xlsx]Model'!$A$1</definedName>${closingTag}`,
        ),
      );
    },
    /external-workbook defined name/i,
  );
  await expectInvalidPackage(
    api,
    restored,
    async (zip) => {
      const file = zip.file("_rels/.rels");
      const xml = await file.async("string");
      const relationship = xml.match(/<Relationship\b[^>]*\/>/)[0];
      zip.file("_rels/.rels", xml.replace("</Relationships>", `${relationship}</Relationships>`));
    },
    /Duplicate relationship id/,
  );
  await expectInvalidPackage(
    api,
    restored,
    async (zip) => {
      const file = zip.file("_rels/.rels");
      const xml = await file.async("string");
      zip.file("_rels/.rels", xml.replace(/Target="[^"]+"/, 'Target="missing-part.xml"'));
    },
    /targets missing part/,
  );
  await expectInvalidPackage(
    api,
    restored,
    async (zip) => {
      const file = zip.file("[Content_Types].xml");
      const xml = await file.async("string");
      const defaultType = xml.match(/<Default\b[^>]*\/>/)[0];
      zip.file("[Content_Types].xml", xml.replace("</Types>", `${defaultType}</Types>`));
    },
    /Duplicate content-type default/,
  );
  await expectInvalidPackage(
    api,
    restored,
    async (zip) => {
      const file = zip.file("[Content_Types].xml");
      const xml = await file.async("string");
      zip.file(
        "[Content_Types].xml",
        xml.replace(/<Default\b[^>]*Extension="rels"[^>]*\/>/i, ""),
      );
    },
    /has no declared content type: _rels\/\.rels/,
  );
}

async function runAlternatePartAndPrefixCheck(api) {
  const fixture = await fs.readFile(CASES[0].file);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(fixture);
  const generated = Buffer.from(await workbook.xlsx.writeBuffer());
  const alternateOriginal = await exerciseAlternatePartNamesAndPrefixes(fixture);
  const expected = await api.captureOoxmlFeatureManifest(alternateOriginal);
  const restored = await api.restoreOoxmlPreservedFeatures(alternateOriginal, generated);
  await api.validateOoxmlPackage(restored);
  api.validateOoxmlFeatureInvariants(expected, await api.captureOoxmlFeatureManifest(restored));
  const restoredZip = await JSZip.loadAsync(restored);
  assert.ok(restoredZip.file("xl/workbook.xml"), "generated workbook part should remain authoritative");
  assert.equal(restoredZip.file("custom/package/book-main.xml"), null, "source workbook part must not be copied as an orphan");
  console.log("PASS alternate OOXML part names and namespace prefixes");
}

async function main() {
  const api = loadTypeScriptModule(MODULE_PATH);
  let validatorFixture;
  for (const testCase of CASES) {
    const restored = await runCase(api, testCase);
    validatorFixture ||= restored;
  }
  await runAlternatePartAndPrefixCheck(api);
  await runValidatorNegativeChecks(api, validatorFixture);
  assert.throws(
    () =>
      api.validateOoxmlFeatureInvariants(
        { sheetNames: ["Original"] },
        { sheetNames: ["Changed"] },
      ),
    /preservation invariant mismatch/,
  );
  console.log("OOXML preservation checks passed.");
}

main().catch((error) => {
  console.error(error instanceof Error ? error.stack : error);
  process.exitCode = 1;
});
