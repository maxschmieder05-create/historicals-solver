import { createHash } from "node:crypto";
import path from "node:path";
import JSZip from "jszip";

const CONTENT_TYPES_PART = "[Content_Types].xml";
const RELATIONSHIPS_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships";
const CONTENT_TYPES_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/content-types";
const MARKUP_COMPATIBILITY_NAMESPACE = "http://schemas.openxmlformats.org/markup-compatibility/2006";

type OoxmlInput = Buffer | Uint8Array | ArrayBuffer;

type OoxmlRelationship = {
  id: string;
  type: string;
  target: string;
  targetMode?: string;
};

type WorksheetPart = {
  name: string;
  partPath: string;
};

export type OoxmlSheetFeatureManifest = {
  drawingGraphHashes: string[];
  hyperlinkCount: number;
  hyperlinksHash?: string;
  mergeCellCount: number;
  mergeCellsHash?: string;
  pageSetupHash?: string;
  printerSettingsGraphHashes: string[];
  commentMetadataGraphHashes: string[];
};

export type OoxmlFeatureManifest = {
  sheetNames: string[];
  sheets: Record<string, OoxmlSheetFeatureManifest>;
  definedNameCount: number;
  definedNamesHash?: string;
  externalReferenceCount: number;
  externalReferencesHash?: string;
  workbookExtLstHash?: string;
  customPropertiesGraphHashes: string[];
  auxiliaryWorkbookRelationshipHashes: string[];
};

type CopyContext = {
  source: JSZip;
  output: JSZip;
  sourceWorkbookPart: string;
  outputWorkbookPart: string;
  sourceContentTypes: ContentTypesEditor;
  outputContentTypes: ContentTypesEditor;
  copiedPartPaths: Map<string, string>;
};

/**
 * Restore package-level workbook features that ExcelJS cannot round-trip.
 *
 * The generated workbook remains authoritative for cells, formulas, styles,
 * comments, and VML comments. The original package is authoritative only for
 * drawings/charts, hyperlinks, merged-range structure, printer settings, workbook defined names,
 * external references, workbook extensions/metadata, and custom properties.
 * calcChain is deliberately removed so Excel recalculates the edited model.
 */
export async function restoreOoxmlPreservedFeatures(
  originalInput: OoxmlInput,
  generatedInput: OoxmlInput,
): Promise<Buffer<ArrayBuffer>> {
  const source = await JSZip.loadAsync(originalInput);
  const output = await JSZip.loadAsync(generatedInput);
  const sourceWorkbookPart = await workbookPartPath(source);
  const outputWorkbookPart = await workbookPartPath(output);
  const expectedManifest = await captureOoxmlFeatureManifestFromZip(source);
  const sourceContentTypes = await ContentTypesEditor.fromZip(source);
  const outputContentTypes = await ContentTypesEditor.fromZip(output);
  const context: CopyContext = {
    source,
    output,
    sourceWorkbookPart,
    outputWorkbookPart,
    sourceContentTypes,
    outputContentTypes,
    copiedPartPaths: new Map<string, string>(),
  };

  await restoreWorksheetFeatures(context);
  await restoreWorkbookFeatures(context);
  await restoreRootCustomProperties(context);
  await removeCalculationChain(output, outputContentTypes, outputWorkbookPart);
  output.file(CONTENT_TYPES_PART, outputContentTypes.toXml());

  await validateOoxmlPackageZip(output);
  const actualManifest = await captureOoxmlFeatureManifestFromZip(output);
  validateOoxmlFeatureInvariants(expectedManifest, actualManifest);

  const restored = await output.generateAsync({
    type: "nodebuffer",
    // The caller deliberately serializes large model/audit workbooks with
    // STORE. Re-DEFLATEing every generated XML part here can add minutes while
    // providing no preservation benefit, so keep the same fast-write policy.
    compression: "STORE",
  });
  await validateOoxmlPackage(restored);
  return restored as Buffer<ArrayBuffer>;
}

export async function captureOoxmlFeatureManifest(input: OoxmlInput): Promise<OoxmlFeatureManifest> {
  return captureOoxmlFeatureManifestFromZip(await JSZip.loadAsync(input));
}

export function validateOoxmlFeatureInvariants(
  expected: OoxmlFeatureManifest,
  actual: OoxmlFeatureManifest,
): void {
  if (!expected.sheets || !actual.sheets) {
    const expectedJson = JSON.stringify(expected);
    const actualJson = JSON.stringify(actual);
    if (expectedJson !== actualJson) {
      throw new Error(`OOXML preservation invariant mismatch.\nExpected: ${expectedJson}\nActual: ${actualJson}`);
    }
    return;
  }
  const expectedSheetNames = new Set(expected.sheetNames);
  const actualOriginalSheetOrder = actual.sheetNames.filter((sheetName) => expectedSheetNames.has(sheetName));
  const projectedActual: OoxmlFeatureManifest = {
    ...actual,
    sheetNames: actualOriginalSheetOrder,
    sheets: Object.fromEntries(
      expected.sheetNames.map((sheetName) => [sheetName, actual.sheets[sheetName]]),
    ),
  };
  const expectedJson = JSON.stringify(expected);
  const actualJson = JSON.stringify(projectedActual);
  if (expectedJson !== actualJson) {
    throw new Error(`OOXML preservation invariant mismatch.\nExpected: ${expectedJson}\nActual: ${actualJson}`);
  }
}

export async function validateOoxmlPackage(input: OoxmlInput): Promise<void> {
  await validateOoxmlPackageZip(await JSZip.loadAsync(input));
}

async function restoreWorksheetFeatures(context: CopyContext): Promise<void> {
  const sourceSheets = await worksheetPartsByName(context.source);
  const outputSheets = await worksheetPartsByName(context.output);

  for (const sourceSheet of sourceSheets.values()) {
    const outputSheet = outputSheets.get(sourceSheet.name);
    if (!outputSheet) {
      throw new Error(`Generated workbook is missing worksheet ${JSON.stringify(sourceSheet.name)}.`);
    }

    const sourceXml = await requiredXml(context.source, sourceSheet.partPath);
    let outputXml = await requiredXml(context.output, outputSheet.partPath);
    const sourceRelationships = await RelationshipsEditor.fromZip(context.source, sourceSheet.partPath);
    const outputRelationships = await RelationshipsEditor.fromZip(context.output, outputSheet.partPath);
    const sourceToOutputRelationshipIds = new Map<string, string>();

    await restoreWorksheetCommentMetadata(
      context,
      sourceSheet.partPath,
      outputSheet.partPath,
      sourceRelationships,
      outputRelationships,
    );

    // ExcelJS preserves the set of merged ranges but rewrites their XML order.
    // Its in-memory worksheet model retains the source order, so a write/reload
    // otherwise looks like a protected template-structure mutation even though
    // the ranges are semantically identical. The solver never creates or removes
    // merges on source worksheets, making the original block authoritative.
    const sourceMergeCells = findElementBlock(sourceXml, "mergeCells");
    outputXml = ensureFragmentNamespaces(sourceXml, outputXml, sourceMergeCells);
    outputXml = replaceOrInsertElementBlock(
      outputXml,
      "mergeCells",
      sourceMergeCells,
      [
        "phoneticPr",
        "conditionalFormatting",
        "dataValidations",
        "hyperlinks",
        "printOptions",
        "pageMargins",
        "pageSetup",
        "headerFooter",
        "rowBreaks",
        "colBreaks",
        "customProperties",
        "cellWatches",
        "ignoredErrors",
        "smartTags",
        "drawing",
        "legacyDrawing",
        "legacyDrawingHF",
        "picture",
        "oleObjects",
        "controls",
        "webPublishItems",
        "tableParts",
        "extLst",
      ],
    );

    const sourceDrawing = findElementBlock(sourceXml, "drawing");
    const generatedDrawing = findElementBlock(outputXml, "drawing");
    outputRelationships.removeIds(relationshipIdsFromFragment(generatedDrawing, outputXml));
    outputRelationships.removeWhere((relationship) => relationshipTypeIs(relationship, "drawing"));
    const restoredDrawing = sourceDrawing
      ? await remapFragmentRelationships(
          sourceDrawing,
          sourceSheet.partPath,
          outputSheet.partPath,
          sourceRelationships,
          outputRelationships,
          sourceToOutputRelationshipIds,
          context,
          sourceXml,
        )
      : undefined;
    outputXml = ensureFragmentNamespaces(sourceXml, outputXml, restoredDrawing);
    outputXml = replaceOrInsertElementBlock(
      outputXml,
      "drawing",
      restoredDrawing,
      ["legacyDrawing", "legacyDrawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst"],
    );

    const sourceHyperlinks = findElementBlock(sourceXml, "hyperlinks");
    const generatedHyperlinks = findElementBlock(outputXml, "hyperlinks");
    outputRelationships.removeIds(relationshipIdsFromFragment(generatedHyperlinks, outputXml));
    outputRelationships.removeWhere((relationship) => relationshipTypeIs(relationship, "hyperlink"));
    const restoredHyperlinks = sourceHyperlinks
      ? await remapFragmentRelationships(
          sourceHyperlinks,
          sourceSheet.partPath,
          outputSheet.partPath,
          sourceRelationships,
          outputRelationships,
          sourceToOutputRelationshipIds,
          context,
          sourceXml,
        )
      : undefined;
    outputXml = ensureFragmentNamespaces(sourceXml, outputXml, restoredHyperlinks);
    outputXml = replaceOrInsertElementBlock(
      outputXml,
      "hyperlinks",
      restoredHyperlinks,
      [
        "printOptions",
        "pageMargins",
        "pageSetup",
        "headerFooter",
        "rowBreaks",
        "colBreaks",
        "customProperties",
        "cellWatches",
        "ignoredErrors",
        "smartTags",
        "drawing",
        "legacyDrawing",
        "legacyDrawingHF",
        "picture",
        "oleObjects",
        "controls",
        "webPublishItems",
        "tableParts",
        "extLst",
      ],
    );

    const sourcePageSetup = findElementBlock(sourceXml, "pageSetup");
    const generatedPageSetup = findElementBlock(outputXml, "pageSetup");
    outputRelationships.removeIds(relationshipIdsFromFragment(generatedPageSetup, outputXml));
    outputRelationships.removeWhere((relationship) => relationshipTypeIs(relationship, "printerSettings"));
    const restoredPageSetup = sourcePageSetup
      ? await remapFragmentRelationships(
          sourcePageSetup,
          sourceSheet.partPath,
          outputSheet.partPath,
          sourceRelationships,
          outputRelationships,
          sourceToOutputRelationshipIds,
          context,
          sourceXml,
        )
      : undefined;
    outputXml = ensureFragmentNamespaces(sourceXml, outputXml, restoredPageSetup);
    outputXml = replaceOrInsertElementBlock(
      outputXml,
      "pageSetup",
      restoredPageSetup,
      [
        "headerFooter",
        "rowBreaks",
        "colBreaks",
        "customProperties",
        "cellWatches",
        "ignoredErrors",
        "smartTags",
        "drawing",
        "legacyDrawing",
        "legacyDrawingHF",
        "picture",
        "oleObjects",
        "controls",
        "webPublishItems",
        "tableParts",
        "extLst",
      ],
    );

    context.output.file(outputSheet.partPath, outputXml);
    saveRelationships(context.output, outputSheet.partPath, outputRelationships);
  }
}

async function restoreWorksheetCommentMetadata(
  context: CopyContext,
  sourceSheetPart: string,
  outputSheetPart: string,
  sourceSheetRelationships: RelationshipsEditor,
  outputSheetRelationships: RelationshipsEditor,
): Promise<void> {
  const sourceCommentsRelationship = sourceSheetRelationships.relationships.find((relationship) => relationshipTypeIs(relationship, "comments"));
  const outputCommentsRelationship = outputSheetRelationships.relationships.find((relationship) => relationshipTypeIs(relationship, "comments"));
  if (!sourceCommentsRelationship || !outputCommentsRelationship) return;
  const sourceCommentsPart = resolveRelationshipTarget(sourceSheetPart, sourceCommentsRelationship.target);
  const outputCommentsPart = resolveRelationshipTarget(outputSheetPart, outputCommentsRelationship.target);
  const sourceRelationships = await RelationshipsEditor.fromZip(context.source, sourceCommentsPart);
  const outputRelationships = await RelationshipsEditor.fromZip(context.output, outputCommentsPart);
  outputRelationships.removeWhere((relationship) => relationshipTypeIs(relationship, "workbookmetadata"));
  const relationshipIdMap = new Map<string, string>();
  for (const relationship of sourceRelationships.relationships) {
    if (!relationshipTypeIs(relationship, "workbookmetadata")) continue;
    await copyRelationshipToOwner(
      relationship,
      sourceCommentsPart,
      outputCommentsPart,
      outputRelationships,
      relationshipIdMap,
      context,
    );
  }
  saveRelationships(context.output, outputCommentsPart, outputRelationships);
}

async function restoreWorkbookFeatures(context: CopyContext): Promise<void> {
  const sourceXml = await requiredXml(context.source, context.sourceWorkbookPart);
  let outputXml = await requiredXml(context.output, context.outputWorkbookPart);
  const sourceRelationships = await RelationshipsEditor.fromZip(context.source, context.sourceWorkbookPart);
  const outputRelationships = await RelationshipsEditor.fromZip(context.output, context.outputWorkbookPart);
  const sourceToOutputRelationshipIds = new Map<string, string>();

  const generatedExternalReferences = findElementBlock(outputXml, "externalReferences");
  const generatedExtLst = findElementBlock(outputXml, "extLst");
  outputRelationships.removeIds([
    ...relationshipIdsFromFragment(generatedExternalReferences, outputXml),
    ...relationshipIdsFromFragment(generatedExtLst, outputXml),
  ]);
  outputRelationships.removeWhere(
    (relationship) =>
      relationshipTypeIs(relationship, "externalLink") ||
      isAuxiliaryWorkbookRelationship(relationship),
  );

  const sourceExternalReferences = findElementBlock(sourceXml, "externalReferences");
  const restoredExternalReferences = sourceExternalReferences
    ? await remapFragmentRelationships(
        sourceExternalReferences,
        context.sourceWorkbookPart,
        context.outputWorkbookPart,
        sourceRelationships,
        outputRelationships,
        sourceToOutputRelationshipIds,
        context,
        sourceXml,
      )
    : undefined;
  outputXml = ensureFragmentNamespaces(sourceXml, outputXml, restoredExternalReferences);
  outputXml = replaceOrInsertElementBlock(
    outputXml,
    "externalReferences",
    restoredExternalReferences,
    ["definedNames", "calcPr", "oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst"],
  );

  const sourceDefinedNames = sanitizeExternalWorkbookDefinedNames(findElementBlock(sourceXml, "definedNames"));
  outputXml = ensureFragmentNamespaces(sourceXml, outputXml, sourceDefinedNames);
  outputXml = replaceOrInsertElementBlock(
    outputXml,
    "definedNames",
    sourceDefinedNames,
    ["calcPr", "oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst"],
  );

  const sourceExtLst = findElementBlock(sourceXml, "extLst");
  const restoredExtLst = sourceExtLst
    ? await remapFragmentRelationships(
        sourceExtLst,
        context.sourceWorkbookPart,
        context.outputWorkbookPart,
        sourceRelationships,
        outputRelationships,
        sourceToOutputRelationshipIds,
        context,
        sourceXml,
      )
    : undefined;
  outputXml = ensureFragmentNamespaces(sourceXml, outputXml, restoredExtLst);
  outputXml = replaceOrInsertElementBlock(outputXml, "extLst", restoredExtLst, []);

  for (const relationship of sourceRelationships.relationships) {
    if (!isAuxiliaryWorkbookRelationship(relationship)) continue;
    if (sourceToOutputRelationshipIds.has(relationship.id)) continue;
    await copyRelationshipToOwner(
      relationship,
      context.sourceWorkbookPart,
      context.outputWorkbookPart,
      outputRelationships,
      sourceToOutputRelationshipIds,
      context,
    );
  }

  context.output.file(context.outputWorkbookPart, outputXml);
  saveRelationships(context.output, context.outputWorkbookPart, outputRelationships);
}

async function restoreRootCustomProperties(context: CopyContext): Promise<void> {
  const sourceRelationships = await RelationshipsEditor.fromZip(context.source, "");
  const outputRelationships = await RelationshipsEditor.fromZip(context.output, "");
  outputRelationships.removeWhere((relationship) => relationshipTypeIs(relationship, "custom-properties"));
  const relationshipIdMap = new Map<string, string>();

  for (const relationship of sourceRelationships.relationships) {
    if (!relationshipTypeIs(relationship, "custom-properties")) continue;
    await copyRelationshipToOwner(
      relationship,
      "",
      "",
      outputRelationships,
      relationshipIdMap,
      context,
    );
  }
  saveRelationships(context.output, "", outputRelationships);
}

async function remapFragmentRelationships(
  sourceFragment: string,
  sourceOwnerPart: string,
  outputOwnerPart: string,
  sourceRelationships: RelationshipsEditor,
  outputRelationships: RelationshipsEditor,
  sourceToOutputRelationshipIds: Map<string, string>,
  context: CopyContext,
  sourceDocumentXml: string,
): Promise<string> {
  for (const sourceRelationshipId of relationshipIdsFromFragment(sourceFragment, sourceDocumentXml)) {
    let outputRelationshipId = sourceToOutputRelationshipIds.get(sourceRelationshipId);
    if (!outputRelationshipId) {
      const sourceRelationship = sourceRelationships.find(sourceRelationshipId);
      if (!sourceRelationship) {
        throw new Error(
          `${sourceOwnerPart || "package"} references missing relationship ${sourceRelationshipId}.`,
        );
      }
      outputRelationshipId = await copyRelationshipToOwner(
        sourceRelationship,
        sourceOwnerPart,
        outputOwnerPart,
        outputRelationships,
        sourceToOutputRelationshipIds,
        context,
      );
    }
  }
  return replaceRelationshipIds(sourceFragment, sourceToOutputRelationshipIds, sourceDocumentXml);
}

async function copyRelationshipToOwner(
  sourceRelationship: OoxmlRelationship,
  sourceOwnerPart: string,
  outputOwnerPart: string,
  outputRelationships: RelationshipsEditor,
  sourceToOutputRelationshipIds: Map<string, string>,
  context: CopyContext,
): Promise<string> {
  if (relationshipTypeIs(sourceRelationship, "calcChain")) {
    throw new Error("calcChain relationships are intentionally not restorable.");
  }
  const existing = sourceToOutputRelationshipIds.get(sourceRelationship.id);
  if (existing) return existing;

  const outputRelationshipId = outputRelationships.allocateId();
  let target = sourceRelationship.target;
  if (!isExternalRelationship(sourceRelationship)) {
    const sourceTargetPart = resolveRelationshipTarget(sourceOwnerPart, sourceRelationship.target);
    const outputTargetPart = await copyPartGraph(sourceTargetPart, context);
    target = relativeRelationshipTarget(outputOwnerPart, outputTargetPart);
  }
  outputRelationships.add({
    id: outputRelationshipId,
    type: sourceRelationship.type,
    target,
    targetMode: sourceRelationship.targetMode,
  });
  sourceToOutputRelationshipIds.set(sourceRelationship.id, outputRelationshipId);
  return outputRelationshipId;
}

async function copyPartGraph(sourcePartPath: string, context: CopyContext): Promise<string> {
  const normalizedSourcePartPath = normalizePartPath(sourcePartPath);
  const previouslyCopied = context.copiedPartPaths.get(normalizedSourcePartPath);
  if (previouslyCopied) return previouslyCopied;

  const sourcePart = context.source.file(normalizedSourcePartPath);
  if (!sourcePart) {
    throw new Error(`Original OOXML relationship target is missing: ${normalizedSourcePartPath}.`);
  }
  const outputPartPath = allocatePartPath(context.output, normalizedSourcePartPath);
  context.copiedPartPaths.set(normalizedSourcePartPath, outputPartPath);

  const contentType = context.sourceContentTypes.contentTypeForPart(normalizedSourcePartPath);
  if (!contentType) {
    throw new Error(`Original OOXML part has no content type: ${normalizedSourcePartPath}.`);
  }
  context.outputContentTypes.setOverride(outputPartPath, contentType);
  context.output.file(outputPartPath, await sourcePart.async("uint8array"));

  const sourceRelationships = await RelationshipsEditor.fromZip(context.source, normalizedSourcePartPath);
  const outputRelationships = new RelationshipsEditor();
  for (const relationship of sourceRelationships.relationships) {
    if (relationshipTypeIs(relationship, "calcChain")) continue;
    if (isExternalRelationship(relationship)) {
      outputRelationships.add({ ...relationship });
      continue;
    }
    const sourceTargetPart = resolveRelationshipTarget(normalizedSourcePartPath, relationship.target);
    const outputTargetPart = await copyPartGraph(sourceTargetPart, context);
    outputRelationships.add({
      ...relationship,
      target: relativeRelationshipTarget(outputPartPath, outputTargetPart),
    });
  }
  saveRelationships(context.output, outputPartPath, outputRelationships);
  return outputPartPath;
}

async function removeCalculationChain(
  zip: JSZip,
  contentTypes: ContentTypesEditor,
  workbookPart: string,
): Promise<void> {
  const relationships = await RelationshipsEditor.fromZip(zip, workbookPart);
  for (const relationship of relationships.relationships) {
    if (!relationshipTypeIs(relationship, "calcChain")) continue;
    if (!isExternalRelationship(relationship)) {
      zip.remove(resolveRelationshipTarget(workbookPart, relationship.target));
    }
  }
  relationships.removeWhere((relationship) => relationshipTypeIs(relationship, "calcChain"));
  saveRelationships(zip, workbookPart, relationships);
  contentTypes.removeWhere(
    (partName, contentType) =>
      partName.toLowerCase().includes("calcchain") || contentType.toLowerCase().includes("calcchain"),
  );
}

async function captureOoxmlFeatureManifestFromZip(zip: JSZip): Promise<OoxmlFeatureManifest> {
  const workbookPart = await workbookPartPath(zip);
  const workbookXml = await requiredXml(zip, workbookPart);
  const workbookRelationships = await RelationshipsEditor.fromZip(zip, workbookPart);
  const sheets = await worksheetPartsByName(zip, workbookPart);
  const graphMemo = new Map<string, string>();
  const sheetManifest: Record<string, OoxmlSheetFeatureManifest> = {};

  for (const sheet of sheets.values()) {
    const sheetXml = await requiredXml(zip, sheet.partPath);
    const sheetRelationships = await RelationshipsEditor.fromZip(zip, sheet.partPath);
    const drawing = findElementBlock(sheetXml, "drawing");
    const hyperlinks = findElementBlock(sheetXml, "hyperlinks");
    const mergeCells = findElementBlock(sheetXml, "mergeCells");
    const pageSetup = findElementBlock(sheetXml, "pageSetup");

    const drawingGraphHashes: string[] = [];
    for (const relationshipId of relationshipIdsFromFragment(drawing, sheetXml)) {
      const relationship = requiredRelationship(sheetRelationships, relationshipId, sheet.partPath);
      if (!relationshipTypeIs(relationship, "drawing")) continue;
      drawingGraphHashes.push(
        await relationshipGraphHash(zip, sheet.partPath, relationship, graphMemo, new Set<string>()),
      );
    }

    const printerSettingsGraphHashes: string[] = [];
    for (const relationshipId of relationshipIdsFromFragment(pageSetup, sheetXml)) {
      const relationship = requiredRelationship(sheetRelationships, relationshipId, sheet.partPath);
      if (!relationshipTypeIs(relationship, "printerSettings")) continue;
      printerSettingsGraphHashes.push(
        await relationshipGraphHash(zip, sheet.partPath, relationship, graphMemo, new Set<string>()),
      );
    }

    const commentMetadataGraphHashes: string[] = [];
    for (const commentsRelationship of sheetRelationships.relationships.filter((relationship) => relationshipTypeIs(relationship, "comments"))) {
      const commentsPart = resolveRelationshipTarget(sheet.partPath, commentsRelationship.target);
      const commentsRelationships = await RelationshipsEditor.fromZip(zip, commentsPart);
      for (const metadataRelationship of commentsRelationships.relationships.filter((relationship) => relationshipTypeIs(relationship, "workbookmetadata"))) {
        commentMetadataGraphHashes.push(
          await relationshipGraphHash(zip, commentsPart, metadataRelationship, graphMemo, new Set<string>()),
        );
      }
    }

    sheetManifest[sheet.name] = {
      drawingGraphHashes: drawingGraphHashes.sort(),
      hyperlinkCount: countElements(hyperlinks, "hyperlink"),
      hyperlinksHash: hyperlinks
        ? await canonicalFragmentHash(zip, sheet.partPath, hyperlinks, sheetRelationships, graphMemo, sheetXml)
        : undefined,
      mergeCellCount: countElements(mergeCells, "mergeCell"),
      mergeCellsHash: mergeCells ? sha256(mergeCells) : undefined,
      pageSetupHash: pageSetup
        ? await canonicalFragmentHash(zip, sheet.partPath, pageSetup, sheetRelationships, graphMemo, sheetXml)
        : undefined,
      printerSettingsGraphHashes: printerSettingsGraphHashes.sort(),
      commentMetadataGraphHashes: commentMetadataGraphHashes.sort(),
    };
  }

  const definedNames = sanitizeExternalWorkbookDefinedNames(findElementBlock(workbookXml, "definedNames"));
  const externalReferences = findElementBlock(workbookXml, "externalReferences");
  const workbookExtLst = findElementBlock(workbookXml, "extLst");
  const rootRelationships = await RelationshipsEditor.fromZip(zip, "");
  const customPropertiesGraphHashes: string[] = [];
  for (const relationship of rootRelationships.relationships) {
    if (!relationshipTypeIs(relationship, "custom-properties")) continue;
    customPropertiesGraphHashes.push(
      await relationshipGraphHash(zip, "", relationship, graphMemo, new Set<string>()),
    );
  }

  const auxiliaryWorkbookRelationshipHashes: string[] = [];
  for (const relationship of workbookRelationships.relationships) {
    if (!isAuxiliaryWorkbookRelationship(relationship)) continue;
    auxiliaryWorkbookRelationshipHashes.push(
      await relationshipGraphHash(zip, workbookPart, relationship, graphMemo, new Set<string>()),
    );
  }

  return {
    sheetNames: [...sheets.keys()],
    sheets: sheetManifest,
    definedNameCount: countElements(definedNames, "definedName"),
    definedNamesHash: definedNames ? sha256(definedNames) : undefined,
    externalReferenceCount: countElements(externalReferences, "externalReference"),
    externalReferencesHash: externalReferences
      ? await canonicalFragmentHash(
          zip,
          workbookPart,
          externalReferences,
          workbookRelationships,
          graphMemo,
          workbookXml,
        )
      : undefined,
    workbookExtLstHash: workbookExtLst
      ? await canonicalFragmentHash(zip, workbookPart, workbookExtLst, workbookRelationships, graphMemo, workbookXml)
      : undefined,
    customPropertiesGraphHashes: customPropertiesGraphHashes.sort(),
    auxiliaryWorkbookRelationshipHashes: auxiliaryWorkbookRelationshipHashes.sort(),
  };
}

async function canonicalFragmentHash(
  zip: JSZip,
  ownerPart: string,
  fragment: string,
  relationships: RelationshipsEditor,
  graphMemo: Map<string, string>,
  ownerDocumentXml: string,
): Promise<string> {
  let canonical = fragment;
  for (const relationshipId of relationshipIdsFromFragment(fragment, ownerDocumentXml)) {
    const relationship = requiredRelationship(relationships, relationshipId, ownerPart);
    const relationshipHash = await relationshipGraphHash(
      zip,
      ownerPart,
      relationship,
      graphMemo,
      new Set<string>(),
    );
    canonical = replaceRelationshipId(
      canonical,
      relationshipId,
      `REL-${relationshipHash}`,
      ownerDocumentXml,
    );
  }
  return sha256(canonical);
}

async function relationshipGraphHash(
  zip: JSZip,
  ownerPart: string,
  relationship: OoxmlRelationship,
  memo: Map<string, string>,
  visiting: Set<string>,
): Promise<string> {
  if (isExternalRelationship(relationship)) {
    return sha256(
      JSON.stringify({
        type: relationship.type,
        targetMode: relationship.targetMode ?? "External",
        target: relationship.target,
      }),
    );
  }
  const targetPart = resolveRelationshipTarget(ownerPart, relationship.target);
  const targetHash = await partGraphHash(zip, targetPart, memo, visiting);
  return sha256(JSON.stringify({ type: relationship.type, targetHash }));
}

async function partGraphHash(
  zip: JSZip,
  partPath: string,
  memo: Map<string, string>,
  visiting: Set<string>,
): Promise<string> {
  const normalizedPartPath = normalizePartPath(partPath);
  const memoized = memo.get(normalizedPartPath);
  if (memoized) return memoized;
  if (visiting.has(normalizedPartPath)) {
    throw new Error(`Cyclic OOXML part relationship graph at ${normalizedPartPath}.`);
  }
  visiting.add(normalizedPartPath);
  const part = zip.file(normalizedPartPath);
  if (!part) throw new Error(`OOXML relationship target is missing: ${normalizedPartPath}.`);
  const payloadHash = sha256(await part.async("uint8array"));
  const relationships = await RelationshipsEditor.fromZip(zip, normalizedPartPath);
  const relationshipHashes: Array<{ id: string; hash: string }> = [];
  for (const relationship of relationships.relationships) {
    relationshipHashes.push({
      id: relationship.id,
      hash: await relationshipGraphHash(zip, normalizedPartPath, relationship, memo, visiting),
    });
  }
  relationshipHashes.sort((left, right) => left.id.localeCompare(right.id));
  visiting.delete(normalizedPartPath);
  const result = sha256(JSON.stringify({ payloadHash, relationshipHashes }));
  memo.set(normalizedPartPath, result);
  return result;
}

async function workbookPartPath(zip: JSZip): Promise<string> {
  const rootRelationships = await RelationshipsEditor.fromZip(zip, "");
  const officeDocumentRelationships = rootRelationships.relationships.filter((relationship) =>
    relationshipTypeIs(relationship, "officeDocument"),
  );
  if (officeDocumentRelationships.length !== 1) {
    throw new Error(
      `OOXML package must declare exactly one officeDocument relationship; found ${officeDocumentRelationships.length}.`,
    );
  }
  const relationship = officeDocumentRelationships[0];
  if (isExternalRelationship(relationship)) {
    throw new Error("OOXML officeDocument relationship cannot be external.");
  }
  const partPath = resolveRelationshipTarget("", relationship.target);
  if (!zip.file(partPath)) {
    throw new Error(`OOXML officeDocument relationship targets missing part ${partPath}.`);
  }
  return partPath;
}

async function validateOoxmlPackageZip(zip: JSZip): Promise<void> {
  const contentTypesXml = await requiredXml(zip, CONTENT_TYPES_PART);
  const contentTypes = ContentTypesEditor.fromXml(contentTypesXml);
  validateNoDuplicateContentTypes(contentTypesXml);

  for (const [partPath, entry] of Object.entries(zip.files)) {
    if (entry.dir || partPath === CONTENT_TYPES_PART) continue;
    if (!contentTypes.contentTypeForPart(partPath)) {
      throw new Error(`OOXML part has no declared content type: ${partPath}.`);
    }
  }

  for (const [relationshipPartPath, entry] of Object.entries(zip.files)) {
    if (entry.dir || !relationshipPartPath.endsWith(".rels")) continue;
    const relationshipsXml = await requiredXml(zip, relationshipPartPath);
    const relationships = RelationshipsEditor.fromXml(relationshipsXml);
    const ids = new Set<string>();
    for (const relationship of relationships.relationships) {
      if (ids.has(relationship.id)) {
        throw new Error(`Duplicate relationship id ${relationship.id} in ${relationshipPartPath}.`);
      }
      ids.add(relationship.id);
      if (isExternalRelationship(relationship)) continue;
      const ownerPart = ownerPartForRelationshipsPath(relationshipPartPath);
      const targetPart = resolveRelationshipTarget(ownerPart, relationship.target);
      if (!zip.file(targetPart)) {
        throw new Error(
          `Relationship ${relationship.id} in ${relationshipPartPath} targets missing part ${targetPart}.`,
        );
      }
    }
  }

  const workbookPart = await workbookPartPath(zip);
  const workbookXml = await requiredXml(zip, workbookPart);
  const unsafeDefinedNames = externalWorkbookDefinedNameElements(findElementBlock(workbookXml, "definedNames"));
  if (unsafeDefinedNames.length) {
    throw new Error(
      `Restored OOXML package still contains ${unsafeDefinedNames.length} external-workbook defined name(s).`,
    );
  }
  const workbookRelationships = await RelationshipsEditor.fromZip(zip, workbookPart);
  if (workbookRelationships.relationships.some((relationship) => relationshipTypeIs(relationship, "calcChain"))) {
    throw new Error("Restored OOXML package still contains a calcChain relationship.");
  }
  const calcChainParts = contentTypes.partNamesWhere((contentType) =>
    contentType.toLowerCase().includes("calcchain"),
  );
  if (calcChainParts.some((partPath) => zip.file(partPath))) {
    throw new Error(`Restored OOXML package still contains a calcChain part: ${calcChainParts[0]}.`);
  }
}

async function worksheetPartsByName(
  zip: JSZip,
  workbookPart?: string,
): Promise<Map<string, WorksheetPart>> {
  const resolvedWorkbookPart = workbookPart ?? (await workbookPartPath(zip));
  const workbookXml = await requiredXml(zip, resolvedWorkbookPart);
  const workbookRelationships = await RelationshipsEditor.fromZip(zip, resolvedWorkbookPart);
  const result = new Map<string, WorksheetPart>();
  for (const match of workbookXml.matchAll(elementStartTagPattern("sheet", true))) {
    const attributes = match[1];
    const name = xmlAttribute(attributes, "name");
    const relationshipId = officeDocumentRelationshipAttribute(attributes, "id", workbookXml);
    if (!name || !relationshipId) continue;
    const relationship = workbookRelationships.find(relationshipId);
    if (!relationship || !relationshipTypeIs(relationship, "worksheet")) continue;
    const partPath = resolveRelationshipTarget(resolvedWorkbookPart, relationship.target);
    if (!zip.file(partPath)) {
      throw new Error(`Worksheet ${JSON.stringify(name)} targets missing part ${partPath}.`);
    }
    if (result.has(name)) throw new Error(`Workbook contains duplicate worksheet name ${JSON.stringify(name)}.`);
    result.set(name, { name, partPath });
  }
  return result;
}

class RelationshipsEditor {
  relationships: OoxmlRelationship[];

  constructor(relationships: OoxmlRelationship[] = []) {
    this.relationships = relationships;
  }

  static fromXml(xml: string): RelationshipsEditor {
    const relationships: OoxmlRelationship[] = [];
    for (const match of xml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?Relationship\b([^>]*?)\/?\s*>/g)) {
      const attributes = match[1];
      const id = xmlAttribute(attributes, "Id");
      const type = xmlAttribute(attributes, "Type");
      const target = xmlAttribute(attributes, "Target");
      if (!id || !type || !target) {
        throw new Error("Malformed OOXML Relationship element.");
      }
      relationships.push({
        id,
        type,
        target,
        targetMode: xmlAttribute(attributes, "TargetMode"),
      });
    }
    return new RelationshipsEditor(relationships);
  }

  static async fromZip(zip: JSZip, ownerPart: string): Promise<RelationshipsEditor> {
    const file = zip.file(relationshipsPathForPart(ownerPart));
    if (!file) return new RelationshipsEditor();
    return RelationshipsEditor.fromXml(await file.async("string"));
  }

  find(id: string): OoxmlRelationship | undefined {
    return this.relationships.find((relationship) => relationship.id === id);
  }

  add(relationship: OoxmlRelationship): void {
    if (this.find(relationship.id)) {
      throw new Error(`Duplicate relationship id ${relationship.id}.`);
    }
    this.relationships.push(relationship);
  }

  removeIds(ids: Iterable<string>): void {
    const idSet = new Set(ids);
    this.relationships = this.relationships.filter((relationship) => !idSet.has(relationship.id));
  }

  removeWhere(predicate: (relationship: OoxmlRelationship) => boolean): void {
    this.relationships = this.relationships.filter((relationship) => !predicate(relationship));
  }

  allocateId(): string {
    const ids = new Set(this.relationships.map((relationship) => relationship.id));
    let numericId = 1;
    while (ids.has(`rId${numericId}`)) numericId += 1;
    return `rId${numericId}`;
  }

  toXml(): string {
    const relationships = this.relationships
      .map((relationship) => {
        const targetMode = relationship.targetMode
          ? ` TargetMode="${encodeXmlAttribute(relationship.targetMode)}"`
          : "";
        return `<Relationship Id="${encodeXmlAttribute(relationship.id)}" Type="${encodeXmlAttribute(relationship.type)}" Target="${encodeXmlAttribute(relationship.target)}"${targetMode}/>`;
      })
      .join("");
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${RELATIONSHIPS_NAMESPACE}">${relationships}</Relationships>`;
  }
}

class ContentTypesEditor {
  private defaults: Map<string, string>;
  private overrides: Map<string, string>;

  constructor(defaults = new Map<string, string>(), overrides = new Map<string, string>()) {
    this.defaults = defaults;
    this.overrides = overrides;
  }

  static fromXml(xml: string): ContentTypesEditor {
    const defaults = new Map<string, string>();
    const overrides = new Map<string, string>();
    for (const match of xml.matchAll(elementStartTagPattern("Default", true))) {
      const extension = xmlAttribute(match[1], "Extension")?.toLowerCase();
      const contentType = xmlAttribute(match[1], "ContentType");
      if (extension && contentType) defaults.set(extension, contentType);
    }
    for (const match of xml.matchAll(elementStartTagPattern("Override", true))) {
      const partName = xmlAttribute(match[1], "PartName");
      const contentType = xmlAttribute(match[1], "ContentType");
      if (partName && contentType) overrides.set(normalizeContentTypePartName(partName), contentType);
    }
    return new ContentTypesEditor(defaults, overrides);
  }

  static async fromZip(zip: JSZip): Promise<ContentTypesEditor> {
    return ContentTypesEditor.fromXml(await requiredXml(zip, CONTENT_TYPES_PART));
  }

  contentTypeForPart(partPath: string): string | undefined {
    const normalized = normalizePartPath(partPath);
    const override = this.overrides.get(normalized);
    if (override) return override;
    const basename = path.posix.basename(normalized);
    const extension =
      basename === ".rels" ? "rels" : path.posix.extname(normalized).slice(1).toLowerCase();
    return extension ? this.defaults.get(extension) : undefined;
  }

  setOverride(partPath: string, contentType: string): void {
    this.overrides.set(normalizePartPath(partPath), contentType);
  }

  removeWhere(predicate: (partName: string, contentType: string) => boolean): void {
    for (const [partName, contentType] of this.overrides) {
      if (predicate(partName, contentType)) this.overrides.delete(partName);
    }
  }

  partNamesWhere(predicate: (contentType: string, partName: string) => boolean): string[] {
    return [...this.overrides.entries()]
      .filter(([partName, contentType]) => predicate(contentType, partName))
      .map(([partName]) => partName);
  }

  toXml(): string {
    const defaults = [...this.defaults.entries()]
      .map(
        ([extension, contentType]) =>
          `<Default Extension="${encodeXmlAttribute(extension)}" ContentType="${encodeXmlAttribute(contentType)}"/>`,
      )
      .join("");
    const overrides = [...this.overrides.entries()]
      .map(
        ([partName, contentType]) =>
          `<Override PartName="/${encodeXmlAttribute(partName)}" ContentType="${encodeXmlAttribute(contentType)}"/>`,
      )
      .join("");
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="${CONTENT_TYPES_NAMESPACE}">${defaults}${overrides}</Types>`;
  }
}

function validateNoDuplicateContentTypes(xml: string): void {
  const defaults = new Set<string>();
  for (const match of xml.matchAll(elementStartTagPattern("Default", true))) {
    const extension = xmlAttribute(match[1], "Extension")?.toLowerCase();
    if (!extension) continue;
    if (defaults.has(extension)) throw new Error(`Duplicate content-type default for .${extension}.`);
    defaults.add(extension);
  }
  const overrides = new Set<string>();
  for (const match of xml.matchAll(elementStartTagPattern("Override", true))) {
    const partName = xmlAttribute(match[1], "PartName");
    if (!partName) continue;
    const normalized = normalizeContentTypePartName(partName);
    if (overrides.has(normalized)) throw new Error(`Duplicate content-type override for /${normalized}.`);
    overrides.add(normalized);
  }
}

function requiredRelationship(
  relationships: RelationshipsEditor,
  relationshipId: string,
  ownerPart: string,
): OoxmlRelationship {
  const relationship = relationships.find(relationshipId);
  if (!relationship) {
    throw new Error(`${ownerPart || "package"} references missing relationship ${relationshipId}.`);
  }
  return relationship;
}

function relationshipTypeIs(relationship: OoxmlRelationship, localType: string): boolean {
  return relationship.type.toLowerCase().endsWith(`/${localType.toLowerCase()}`);
}

function isAuxiliaryWorkbookRelationship(relationship: OoxmlRelationship): boolean {
  return (
    relationshipTypeIs(relationship, "sheetMetadata") ||
    relationshipTypeIs(relationship, "workbookmetadata")
  );
}

function isExternalRelationship(relationship: OoxmlRelationship): boolean {
  return relationship.targetMode?.toLowerCase() === "external";
}

function relationshipsPathForPart(partPath: string): string {
  if (!partPath) return "_rels/.rels";
  const directory = path.posix.dirname(partPath);
  const filename = path.posix.basename(partPath);
  return `${directory}/_rels/${filename}.rels`;
}

function ownerPartForRelationshipsPath(relationshipsPath: string): string {
  if (relationshipsPath === "_rels/.rels") return "";
  const match = relationshipsPath.match(/^(.*)\/_rels\/([^/]+)\.rels$/);
  if (!match) throw new Error(`Invalid OOXML relationships part path: ${relationshipsPath}.`);
  return `${match[1]}/${match[2]}`;
}

function saveRelationships(zip: JSZip, ownerPart: string, relationships: RelationshipsEditor): void {
  const relationshipPartPath = relationshipsPathForPart(ownerPart);
  if (relationships.relationships.length === 0) {
    zip.remove(relationshipPartPath);
    return;
  }
  zip.file(relationshipPartPath, relationships.toXml());
}

function resolveRelationshipTarget(ownerPart: string, target: string): string {
  const targetWithoutFragment = target.split("#", 1)[0].split("?", 1)[0];
  let decodedTarget = targetWithoutFragment;
  try {
    decodedTarget = decodeURI(targetWithoutFragment);
  } catch {
    // Keep the original target; validation below will still fail if no such part exists.
  }
  if (decodedTarget.startsWith("/")) return normalizePartPath(decodedTarget);
  const ownerDirectory = ownerPart ? path.posix.dirname(ownerPart) : ".";
  return normalizePartPath(path.posix.join(ownerDirectory, decodedTarget));
}

function relativeRelationshipTarget(ownerPart: string, targetPart: string): string {
  const ownerDirectory = ownerPart ? path.posix.dirname(ownerPart) : ".";
  const relative = path.posix.relative(ownerDirectory, normalizePartPath(targetPart));
  return encodeURI(relative);
}

function normalizePartPath(partPath: string): string {
  const normalized = path.posix.normalize(partPath.replace(/^\/+/, ""));
  if (!normalized || normalized === "." || normalized === ".." || normalized.startsWith("../")) {
    throw new Error(`Invalid OOXML part path: ${partPath}.`);
  }
  return normalized;
}

function normalizeContentTypePartName(partName: string): string {
  return normalizePartPath(partName.replace(/^\/+/, ""));
}

function allocatePartPath(zip: JSZip, preferredPartPath: string): string {
  const normalizedPreferredPath = normalizePartPath(preferredPartPath);
  if (!zip.files[normalizedPreferredPath]) return normalizedPreferredPath;
  const extension = path.posix.extname(normalizedPreferredPath);
  const stem = extension ? normalizedPreferredPath.slice(0, -extension.length) : normalizedPreferredPath;
  let index = 1;
  while (true) {
    const candidate = `${stem}.codex-preserved-${index}${extension}`;
    if (!zip.files[candidate]) return candidate;
    index += 1;
  }
}

function relationshipIdsFromFragment(fragment?: string, ownerDocumentXml = fragment): string[] {
  if (!fragment) return [];
  const ids = new Set<string>();
  for (const match of fragment.matchAll(officeDocumentRelationshipAttributePattern(ownerDocumentXml ?? fragment, true))) {
    ids.add(decodeXml(match[2]));
  }
  return [...ids];
}

function replaceRelationshipId(
  fragment: string,
  oldId: string,
  newId: string,
  ownerDocumentXml = fragment,
): string {
  return fragment.replace(
    officeDocumentRelationshipAttributePattern(ownerDocumentXml, true),
    (match, prefix, encodedId) =>
      decodeXml(encodedId) === oldId
        ? `${prefix}${encodeXmlAttribute(newId)}${match.slice(-1)}`
        : match,
  );
}

function replaceRelationshipIds(
  fragment: string,
  replacements: Map<string, string>,
  ownerDocumentXml = fragment,
): string {
  return fragment.replace(officeDocumentRelationshipAttributePattern(ownerDocumentXml, true), (match, prefix, encodedId) => {
    const replacement = replacements.get(decodeXml(encodedId));
    return replacement ? `${prefix}${encodeXmlAttribute(replacement)}${match.slice(-1)}` : match;
  });
}

function officeDocumentRelationshipAttribute(
  attributes: string,
  localName: string,
  ownerDocumentXml: string,
): string | undefined {
  for (const prefix of officeDocumentRelationshipPrefixes(ownerDocumentXml)) {
    const value = xmlAttribute(attributes, `${prefix}:${localName}`);
    if (value !== undefined) return value;
  }
  return undefined;
}

function officeDocumentRelationshipAttributePattern(ownerDocumentXml: string, global = false): RegExp {
  const prefixes = officeDocumentRelationshipPrefixes(ownerDocumentXml);
  if (prefixes.length === 0) return /$a/;
  const prefixAlternation = prefixes.map(escapeRegExp).join("|");
  return new RegExp(
    `(\\b(?:${prefixAlternation}):id\\s*=\\s*["'])(.*?)["']`,
    global ? "gi" : "i",
  );
}

function officeDocumentRelationshipPrefixes(xml: string): string[] {
  return namespacePrefixesWhere(xml, (namespace) =>
    /\/officeDocument\/(?:2006\/)?relationships\/?$/i.test(namespace),
  );
}

function namespacePrefixesWhere(xml: string, predicate: (namespace: string) => boolean): string[] {
  const prefixes = new Set<string>();
  for (const match of xml.matchAll(/\bxmlns:([A-Za-z_][\w.-]*)\s*=\s*(["'])(.*?)\2/g)) {
    if (predicate(decodeXml(match[3]))) prefixes.add(match[1]);
  }
  return [...prefixes];
}

function elementStartTagPattern(localName: string, global = false): RegExp {
  const escapedName = escapeRegExp(localName);
  const prefix = "(?:[A-Za-z_][\\w.-]*:)?";
  return new RegExp(`<${prefix}${escapedName}\\b([^>]*)\\/?\\s*>`, global ? "gi" : "i");
}

function findElementBlock(xml: string, localName: string): string | undefined {
  return xml.match(elementBlockPattern(localName))?.[0];
}

function sanitizeExternalWorkbookDefinedNames(fragment?: string): string | undefined {
  if (!fragment) return undefined;
  const unsafe = new Set(externalWorkbookDefinedNameElements(fragment));
  const sanitized = fragment.replace(elementBlockPattern("definedName", true), (element) =>
    unsafe.has(element) ? "" : element,
  );
  return countElements(sanitized, "definedName") ? sanitized : undefined;
}

function externalWorkbookDefinedNameElements(fragment?: string): string[] {
  if (!fragment) return [];
  return Array.from(fragment.matchAll(elementBlockPattern("definedName", true)))
    .map((match) => match[0])
    .filter((element) => definedNameReferencesExternalWorkbook(element));
}

function definedNameReferencesExternalWorkbook(element: string): boolean {
  const openEnd = element.indexOf(">");
  const closeStart = element.lastIndexOf("</");
  if (openEnd < 0 || closeStart <= openEnd) return false;
  const expression = decodeXml(element.slice(openEnd + 1, closeStart)).trim();
  return (
    /\[[^\]]+\.(?:xlsx?|xlsm|xlsb)\][^!]*!/i.test(expression) ||
    /(?:^|[=+\-*/,(])\s*'?\[[^\]]+\][^'!]*'?!/i.test(expression)
  );
}

function elementBlockPattern(localName: string, global = false): RegExp {
  const escapedName = escapeRegExp(localName);
  const prefix = "(?:[A-Za-z_][\\w.-]*:)?";
  return new RegExp(
    `<${prefix}${escapedName}\\b[^>]*(?:\\/>|>[\\s\\S]*?<\\/${prefix}${escapedName}\\s*>)`,
    global ? "gi" : "i",
  );
}

function replaceOrInsertElementBlock(
  xml: string,
  localName: string,
  replacement: string | undefined,
  followingElementNames: string[],
): string {
  const existingPattern = elementBlockPattern(localName, true);
  if (!replacement) return xml.replace(existingPattern, "");
  if (elementBlockPattern(localName).test(xml)) return xml.replace(existingPattern, replacement);

  for (const followingElementName of followingElementNames) {
    const followingPattern = new RegExp(
      `<(?:[A-Za-z_][\\w.-]*:)?${escapeRegExp(followingElementName)}\\b`,
      "i",
    );
    const match = followingPattern.exec(xml);
    if (match?.index !== undefined) {
      return `${xml.slice(0, match.index)}${replacement}${xml.slice(match.index)}`;
    }
  }
  const closingTag = /<\/(?:[A-Za-z_][\w.-]*:)?(?:worksheet|workbook)\s*>/i;
  const closingMatch = closingTag.exec(xml);
  if (!closingMatch?.index && closingMatch?.index !== 0) {
    throw new Error(`Cannot insert ${localName}; document has no worksheet/workbook closing tag.`);
  }
  return `${xml.slice(0, closingMatch.index)}${replacement}${xml.slice(closingMatch.index)}`;
}

function ensureFragmentNamespaces(sourceXml: string, outputXml: string, fragment?: string): string {
  if (!fragment) return outputXml;
  const prefixes = new Set<string>();
  for (const tagMatch of fragment.matchAll(/<[^>]+>/g)) {
    const tag = tagMatch[0];
    const elementPrefix = tag.match(/^<\/?\s*([A-Za-z_][\w.-]*):[A-Za-z_][\w.-]*/)?.[1];
    if (elementPrefix && elementPrefix !== "xml" && elementPrefix !== "xmlns") {
      prefixes.add(elementPrefix);
    }
    for (const attributeMatch of tag.matchAll(/\s([A-Za-z_][\w.-]*):[A-Za-z_][\w.-]*\s*=/g)) {
      if (attributeMatch[1] !== "xml" && attributeMatch[1] !== "xmlns") {
        prefixes.add(attributeMatch[1]);
      }
    }
  }
  if (prefixes.size === 0) return outputXml;

  const rootPattern = /<(?:[A-Za-z_][\w.-]*:)?(?:worksheet|workbook)\b[^>]*>/i;
  const sourceRoot = sourceXml.match(rootPattern)?.[0];
  const outputRoot = outputXml.match(rootPattern)?.[0];
  if (!sourceRoot || !outputRoot) return outputXml;
  let updatedRoot = outputRoot;
  const sourceMarkupCompatibilityPrefix = namespacePrefixesWhere(
    sourceRoot,
    (namespace) => namespace === MARKUP_COMPATIBILITY_NAMESPACE,
  )[0];
  let outputMarkupCompatibilityPrefix = namespacePrefixesWhere(
    outputRoot,
    (namespace) => namespace === MARKUP_COMPATIBILITY_NAMESPACE,
  )[0];
  const ignorablePrefixes = new Set(
    (
      outputMarkupCompatibilityPrefix
        ? (xmlAttribute(outputRoot, `${outputMarkupCompatibilityPrefix}:Ignorable`) ?? "")
        : ""
    )
      .split(/\s+/)
      .filter(Boolean),
  );
  const sourceIgnorablePrefixes = new Set(
    (
      sourceMarkupCompatibilityPrefix
        ? (xmlAttribute(sourceRoot, `${sourceMarkupCompatibilityPrefix}:Ignorable`) ?? "")
        : ""
    )
      .split(/\s+/)
      .filter(Boolean),
  );

  for (const prefix of prefixes) {
    if (new RegExp(`\\bxmlns:${escapeRegExp(prefix)}\\s*=`, "i").test(fragment)) continue;
    if (!new RegExp(`\\bxmlns:${escapeRegExp(prefix)}\\s*=`, "i").test(updatedRoot)) {
      const namespace = xmlAttribute(sourceRoot, `xmlns:${prefix}`);
      if (!namespace) {
        throw new Error(`Cannot restore XML fragment because namespace ${prefix} is undeclared.`);
      }
      updatedRoot = updatedRoot.replace(/>$/, ` xmlns:${prefix}="${encodeXmlAttribute(namespace)}">`);
    }
    if (sourceIgnorablePrefixes.has(prefix)) ignorablePrefixes.add(prefix);
  }

  if (ignorablePrefixes.size > 0) {
    const value = [...ignorablePrefixes].join(" ");
    if (!outputMarkupCompatibilityPrefix && sourceMarkupCompatibilityPrefix) {
      updatedRoot = updatedRoot.replace(
        />$/,
        ` xmlns:${sourceMarkupCompatibilityPrefix}="${encodeXmlAttribute(MARKUP_COMPATIBILITY_NAMESPACE)}">`,
      );
      outputMarkupCompatibilityPrefix = sourceMarkupCompatibilityPrefix;
    }
    if (!outputMarkupCompatibilityPrefix) {
      throw new Error("Cannot restore ignorable XML namespaces without a markup-compatibility namespace.");
    }
    const ignorableAttributeName = `${outputMarkupCompatibilityPrefix}:Ignorable`;
    if (new RegExp(`\\b${escapeRegExp(ignorableAttributeName)}\\s*=`).test(updatedRoot)) {
      updatedRoot = updatedRoot.replace(
        new RegExp(`\\b${escapeRegExp(ignorableAttributeName)}\\s*=\\s*(["']).*?\\1`),
        `${ignorableAttributeName}="${encodeXmlAttribute(value)}"`,
      );
    } else {
      updatedRoot = updatedRoot.replace(
        />$/,
        ` ${ignorableAttributeName}="${encodeXmlAttribute(value)}">`,
      );
    }
  }
  return outputXml.replace(outputRoot, updatedRoot);
}

function xmlAttribute(attributes: string, name: string): string | undefined {
  const match = attributes.match(
    new RegExp(`(?:^|\\s)${escapeRegExp(name)}\\s*=\\s*(["'])(.*?)\\1`, "i"),
  );
  return match ? decodeXml(match[2]) : undefined;
}

function encodeXmlAttribute(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function decodeXml(value: string): string {
  return value.replace(/&(?:#x[0-9a-f]+|#\d+|amp|lt|gt|quot|apos);/gi, (entity) => {
    const lower = entity.toLowerCase();
    if (lower === "&amp;") return "&";
    if (lower === "&lt;") return "<";
    if (lower === "&gt;") return ">";
    if (lower === "&quot;") return '"';
    if (lower === "&apos;") return "'";
    if (lower.startsWith("&#x")) return String.fromCodePoint(Number.parseInt(lower.slice(3, -1), 16));
    return String.fromCodePoint(Number.parseInt(lower.slice(2, -1), 10));
  });
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function countElements(fragment: string | undefined, localName: string): number {
  if (!fragment) return 0;
  const prefix = "(?:[A-Za-z_][\\w.-]*:)?";
  return [...fragment.matchAll(new RegExp(`<${prefix}${escapeRegExp(localName)}\\b`, "gi"))].length;
}

function sha256(value: string | Uint8Array): string {
  return createHash("sha256").update(value).digest("hex");
}

async function requiredXml(zip: JSZip, partPath: string): Promise<string> {
  const file = zip.file(partPath);
  if (!file) throw new Error(`Required OOXML part is missing: ${partPath}.`);
  return file.async("string");
}
