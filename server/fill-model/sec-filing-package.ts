import { normalizeAccession } from "./sec-accession";

export type SecFilingPackageRequest = {
  cik: string;
  accessionNumber: string;
  form?: string;
  filingDate?: string;
  reportDate?: string;
  primaryDocument?: string;
};

export type SecFilingPackageArtifactKind =
  | "filing_index"
  | "primary_html"
  | "instance"
  | "calculation"
  | "presentation"
  | "label"
  | "definition";

export type SecFilingPackageArtifact = {
  kind: SecFilingPackageArtifactKind;
  name: string;
  url: string;
  type?: string;
  size?: string;
};

export type SecFilingContextDimension = {
  dimension: string;
  member: string;
  typedValue?: string;
};

export type SecFilingStatementSourceTableType =
  | "primary_statement"
  | "footnote"
  | "segment_table"
  | "roll_forward"
  | "support_table"
  | "unknown";

export type SecFilingStatementPeriod = {
  contextRef?: string;
  start?: string;
  end?: string;
  instant?: string;
  periodType?: "instant" | "duration";
};

export type SecFilingParentSubtotal = {
  concept?: string;
  label?: string;
  relationship: "calculation" | "presentation";
  weight?: number;
  roleUri?: string;
};

export type SecFilingDefinitionRelationship = {
  fromConcept: string;
  toConcept: string;
  roleUri?: string;
  arcrole?: string;
};

export type SecFilingStatementRow = {
  statementName: string;
  sourceTableType: SecFilingStatementSourceTableType;
  rowLabel: string;
  xbrlConcept?: string;
  taxonomy?: string;
  value: number | string | null;
  unit?: string;
  period: SecFilingStatementPeriod;
  consolidated: boolean;
  dimensions: SecFilingContextDimension[];
  currentNonCurrentSection?: "current" | "non_current";
  rowOrder: number;
  parentSubtotal?: SecFilingParentSubtotal;
  accession: string;
  reportingPeriod?: string;
  sourceUrl?: string;
};

export type SecFilingStatementStructure = {
  statementName: string;
  sourceTableType: SecFilingStatementSourceTableType;
  roleUri?: string;
  sourceUrl?: string;
  accession: string;
  reportingPeriod?: string;
  form?: string;
  filingDate?: string;
  rows: SecFilingStatementRow[];
};

export type SecFilingPackage = {
  cik: string;
  cikNoLeadingZeroes: string;
  accessionNumber: string;
  accessionKey: string;
  form?: string;
  filingDate?: string;
  reportingPeriod?: string;
  baseUrl: string;
  indexUrl: string;
  primaryDocument?: string;
  primaryDocumentUrl?: string;
  artifacts: Partial<Record<SecFilingPackageArtifactKind, SecFilingPackageArtifact>>;
  definitionRelationships: SecFilingDefinitionRelationship[];
  statements: SecFilingStatementStructure[];
  warnings: string[];
};

export type SecFilingPackageSupport = {
  packages: SecFilingPackage[];
  statements: SecFilingStatementStructure[];
  warnings: string[];
};

type SecArchiveIndex = {
  directory?: {
    item?: SecArchiveIndexItem[];
  };
};

type SecArchiveIndexItem = {
  name: string;
  type?: string;
  size?: string;
  "last-modified"?: string;
};

type ParsedConceptRef = {
  taxonomy?: string;
  concept: string;
  raw: string;
};

type ParsedContext = {
  id: string;
  start?: string;
  end?: string;
  instant?: string;
  periodType?: "instant" | "duration";
  dimensions: SecFilingContextDimension[];
};

type ParsedFact = {
  concept: string;
  taxonomy?: string;
  contextRef: string;
  unitRef?: string;
  unit?: string;
  value: number | string | null;
  decimals?: string;
  rawValue: string;
  context?: ParsedContext;
};

type ParsedInstance = {
  contexts: Map<string, ParsedContext>;
  units: Map<string, string>;
  facts: ParsedFact[];
  factsByConcept: Map<string, ParsedFact[]>;
};

type ConceptLabel = {
  text: string;
  role?: string;
};

type LabelMap = Map<string, ConceptLabel[]>;

type PresentationNode = {
  concept: string;
  taxonomy?: string;
  order: number;
  preferredLabelRole?: string;
  parentConcept?: string;
  children: PresentationNode[];
};

type PresentationStatement = {
  statementName: string;
  roleUri?: string;
  sourceTableType: SecFilingStatementSourceTableType;
  nodes: PresentationNode[];
};

type CalculationParent = {
  parentConcept: string;
  weight?: number;
  roleUri?: string;
};

type LinkbaseLocator = {
  label: string;
  concept: string;
  taxonomy?: string;
};

type LinkbaseArc = {
  from: string;
  to: string;
  order: number;
  preferredLabelRole?: string;
  weight?: number;
  arcrole?: string;
};

const SEC_DEFAULT_USER_AGENT = process.env.SEC_USER_AGENT || "HistoricalsSolver/0.1 contact@example.com";
const SEC_ARCHIVE_ROOT = "https://www.sec.gov/Archives/edgar/data";
const SEC_ARCHIVE_MIN_INTERVAL_MS = Number(process.env.SEC_ARCHIVE_MIN_INTERVAL_MS || 150);
const SEC_FILING_PACKAGE_MAX_FILINGS = Number(process.env.SEC_FILING_PACKAGE_MAX_FILINGS || 24);
const SEC_FILING_PACKAGE_MAX_ROWS_PER_STATEMENT = Number(process.env.SEC_FILING_PACKAGE_MAX_ROWS_PER_STATEMENT || 2000);

const responseTextCache = new Map<string, Promise<string | null>>();
const responseJsonCache = new Map<string, Promise<any | null>>();
const packageCache = new Map<string, Promise<SecFilingPackage | null>>();
let lastSecArchiveFetchAt = 0;
let secArchiveFetchQueue = Promise.resolve();

export async function fetchSecFilingPackageSupport(
  filings: SecFilingPackageRequest[],
  headers: Record<string, string>
): Promise<SecFilingPackageSupport> {
  const uniqueFilings = uniqueFilingsByAccession(filings).slice(0, SEC_FILING_PACKAGE_MAX_FILINGS);
  const packages: SecFilingPackage[] = [];
  const warnings: string[] = [];
  if (filings.length > SEC_FILING_PACKAGE_MAX_FILINGS) {
    warnings.push(`SEC filing-package parsing was limited to ${SEC_FILING_PACKAGE_MAX_FILINGS} selected filing(s) for this request.`);
  }

  for (const filing of uniqueFilings) {
    const secPackage = await fetchSecFilingPackage(filing, headers);
    if (secPackage) packages.push(secPackage);
    else warnings.push(`SEC filing package could not be loaded for accession ${filing.accessionNumber}.`);
  }

  const statements = packages.flatMap((item) => item.statements);
  warnings.push(...packages.flatMap((item) => item.warnings));
  return { packages, statements, warnings: unique(warnings) };
}

export async function fetchSecXbrlFrame(
  taxonomy: string,
  tag: string,
  unit: string,
  period: string,
  headers: Record<string, string>
) {
  const pathParts = [taxonomy, tag, unit, period].map(encodeURIComponent);
  const url = `https://data.sec.gov/api/xbrl/frames/${pathParts.join("/")}.json`;
  return fetchSecJson(url, headers);
}

async function fetchSecFilingPackage(filing: SecFilingPackageRequest, headers: Record<string, string>) {
  const key = `${filing.cik}:${normalizeAccession(filing.accessionNumber)}`;
  let cached = packageCache.get(key);
  if (!cached) {
    cached = fetchSecFilingPackageUncached(filing, headers);
    packageCache.set(key, cached);
  }
  return cached;
}

async function fetchSecFilingPackageUncached(
  filing: SecFilingPackageRequest,
  headers: Record<string, string>
): Promise<SecFilingPackage | null> {
  const cikNoLeadingZeroes = String(Number(filing.cik));
  const accessionKey = normalizeAccession(filing.accessionNumber);
  const baseUrl = `${SEC_ARCHIVE_ROOT}/${cikNoLeadingZeroes}/${accessionKey}`;
  const indexUrl = `${baseUrl}/index.json`;
  const warnings: string[] = [];
  const index = (await fetchSecJson(indexUrl, headers)) as SecArchiveIndex | null;
  const indexItems = index?.directory?.item ?? [];
  if (!indexItems.length) return null;

  const artifacts = discoverFilingArtifacts(indexItems, filing, baseUrl);
  const [primaryHtml, instanceXml, presentationXml, calculationXml, labelXml, definitionXml] = await Promise.all([
    artifacts.primary_html ? fetchSecText(artifacts.primary_html.url, headers, "text/html") : Promise.resolve(null),
    artifacts.instance ? fetchSecText(artifacts.instance.url, headers, "application/xml") : Promise.resolve(null),
    artifacts.presentation ? fetchSecText(artifacts.presentation.url, headers, "application/xml") : Promise.resolve(null),
    artifacts.calculation ? fetchSecText(artifacts.calculation.url, headers, "application/xml") : Promise.resolve(null),
    artifacts.label ? fetchSecText(artifacts.label.url, headers, "application/xml") : Promise.resolve(null),
    artifacts.definition ? fetchSecText(artifacts.definition.url, headers, "application/xml") : Promise.resolve(null)
  ]);

  if (!primaryHtml) warnings.push(`Primary filing HTML was not available for accession ${filing.accessionNumber}.`);
  if (!instanceXml) warnings.push(`XBRL instance XML was not discovered for accession ${filing.accessionNumber}.`);
  if (!presentationXml) warnings.push(`Presentation linkbase was not discovered for accession ${filing.accessionNumber}.`);
  if (!calculationXml) warnings.push(`Calculation linkbase was not discovered for accession ${filing.accessionNumber}.`);
  if (!labelXml) warnings.push(`Label linkbase was not discovered for accession ${filing.accessionNumber}.`);

  const instance = parseInstanceXml(instanceXml ?? primaryHtml ?? "");
  const labels = parseLabelLinkbase(labelXml ?? "");
  const calculations = parseCalculationLinkbase(calculationXml ?? "");
  const definitions = parseDefinitionLinkbase(definitionXml ?? "");
  const presentation = parsePresentationLinkbase(presentationXml ?? "", labels);
  const metadata = {
    cik: filing.cik,
    accessionNumber: filing.accessionNumber,
    accessionKey,
    form: filing.form,
    filingDate: filing.filingDate,
    reportingPeriod: filing.reportDate,
    primaryDocumentUrl: artifacts.primary_html?.url,
    definitionRelationships: definitions
  };

  const primaryStatements = primaryHtml
    ? parsePrimaryHtmlStatements(primaryHtml, instance, labels, calculations, metadata)
    : [];
  const presentationStatements = buildPresentationStatementStructures(presentation, instance, labels, calculations, metadata);
  const statements = [...primaryStatements, ...presentationStatements].filter((statement) => statement.rows.length);

  return {
    cik: filing.cik,
    cikNoLeadingZeroes,
    accessionNumber: filing.accessionNumber,
    accessionKey,
    form: filing.form,
    filingDate: filing.filingDate,
    reportingPeriod: filing.reportDate,
    baseUrl,
    indexUrl,
    primaryDocument: artifacts.primary_html?.name,
    primaryDocumentUrl: artifacts.primary_html?.url,
    artifacts,
    definitionRelationships: definitions,
    statements,
    warnings: unique(warnings)
  };
}

function discoverFilingArtifacts(
  items: SecArchiveIndexItem[],
  filing: SecFilingPackageRequest,
  baseUrl: string
): Partial<Record<SecFilingPackageArtifactKind, SecFilingPackageArtifact>> {
  const artifact = (kind: SecFilingPackageArtifactKind, item: SecArchiveIndexItem): SecFilingPackageArtifact => ({
    kind,
    name: item.name,
    url: `${baseUrl}/${encodeURIComponent(item.name)}`,
    type: item.type,
    size: item.size
  });
  const byName = new Map(items.map((item) => [item.name, item]));
  const artifacts: Partial<Record<SecFilingPackageArtifactKind, SecFilingPackageArtifact>> = {
    filing_index: {
      kind: "filing_index",
      name: "index.json",
      url: `${baseUrl}/index.json`
    }
  };

  const primary =
    (filing.primaryDocument ? byName.get(filing.primaryDocument) : undefined) ??
    items.find((item) => isHtmlFile(item.name) && !isGeneratedFilingSupportFile(item.name));
  if (primary) artifacts.primary_html = artifact("primary_html", primary);

  const xmlItems = items.filter((item) => /\.xml$/i.test(item.name));
  const pickXml = (kind: SecFilingPackageArtifactKind, pattern: RegExp) => {
    const item = xmlItems.find((candidate) => pattern.test(candidate.name));
    if (item) artifacts[kind] = artifact(kind, item);
  };
  pickXml("calculation", /(?:^|[_-])cal(?:[_-]?\d*)?\.xml$/i);
  pickXml("presentation", /(?:^|[_-])pre(?:[_-]?\d*)?\.xml$/i);
  pickXml("label", /(?:^|[_-])lab(?:[_-]?\d*)?\.xml$/i);
  pickXml("definition", /(?:^|[_-])def(?:[_-]?\d*)?\.xml$/i);

  const supportNames = new Set(
    Object.values(artifacts)
      .map((item) => item?.name)
      .filter((name): name is string => Boolean(name))
  );
  const instance =
    xmlItems.find((item) => /_htm\.xml$/i.test(item.name) && !supportNames.has(item.name)) ??
    xmlItems.find((item) => !supportNames.has(item.name) && !isGeneratedFilingSupportFile(item.name) && !/(?:^|[_-])(cal|pre|lab|def)(?:[_-]?\d*)?\.xml$/i.test(item.name));
  if (instance) artifacts.instance = artifact("instance", instance);

  return artifacts;
}

function parseInstanceXml(xml: string): ParsedInstance {
  const contexts = parseInstanceContexts(xml);
  const units = parseInstanceUnits(xml);
  const facts: ParsedFact[] = [];
  const seen = new Set<string>();
  const factPattern = /<([A-Za-z_][\w.-]*):([A-Za-z_][\w.-]*)\b([^>]*)\bcontextRef=["']([^"']+)["']([^>]*)>([\s\S]*?)<\/\1:\2>/g;

  for (const match of xml.matchAll(factPattern)) {
    const prefix = match[1];
    const localName = match[2];
    if (/^(xbrli|xbrldi|link|xlink|ix|dei)$/i.test(prefix) && !/^dei$/i.test(prefix)) continue;
    const attrs = `${match[3]} ${match[5]}`;
    const contextRef = match[4];
    const unitRef = attr(attrs, "unitRef");
    const rawValue = decodeXml(stripTags(match[6]).trim());
    if (!rawValue) continue;
    const value = parseXmlFactValue(rawValue);
    const context = contexts.get(contextRef);
    const unit = unitRef ? units.get(unitRef) ?? unitRef : undefined;
    const key = `${localName}|${contextRef}|${unit ?? ""}|${rawValue}`;
    if (seen.has(key)) continue;
    seen.add(key);
    facts.push({
      concept: localName,
      taxonomy: prefix,
      contextRef,
      unitRef: unitRef ?? undefined,
      unit,
      value,
      decimals: attr(attrs, "decimals") ?? undefined,
      rawValue,
      context
    });
  }

  const factsByConcept = new Map<string, ParsedFact[]>();
  for (const fact of facts) {
    const conceptFacts = factsByConcept.get(fact.concept) ?? [];
    conceptFacts.push(fact);
    factsByConcept.set(fact.concept, conceptFacts);
  }

  return { contexts, units, facts, factsByConcept };
}

function parseInstanceContexts(xml: string) {
  const contexts = new Map<string, ParsedContext>();
  for (const match of xml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?context\b[^>]*\bid=["']([^"']+)["'][^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?context>/g)) {
    const id = match[1];
    const body = match[2];
    const instant = textContent(body, "instant");
    const start = textContent(body, "startDate");
    const end = textContent(body, "endDate");
    contexts.set(id, {
      id,
      start: start ?? undefined,
      end: instant ?? end ?? undefined,
      instant: instant ?? undefined,
      periodType: instant ? "instant" : start && end ? "duration" : undefined,
      dimensions: parseContextDimensions(body)
    });
  }
  return contexts;
}

function parseContextDimensions(contextBody: string): SecFilingContextDimension[] {
  const dimensions: SecFilingContextDimension[] = [];
  for (const match of contextBody.matchAll(/<(?:[A-Za-z_][\w.-]*:)?explicitMember\b([^>]*)>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?explicitMember>/g)) {
    const attrs = match[1];
    const dimension = attr(attrs, "dimension");
    const member = decodeXml(stripTags(match[2]).trim());
    if (dimension && member) dimensions.push({ dimension, member });
  }
  for (const match of contextBody.matchAll(/<(?:[A-Za-z_][\w.-]*:)?typedMember\b([^>]*)>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?typedMember>/g)) {
    const attrs = match[1];
    const dimension = attr(attrs, "dimension");
    const typedValue = decodeXml(stripTags(match[2]).trim());
    if (dimension) dimensions.push({ dimension, member: `${dimension}:typed`, typedValue });
  }
  return dimensions;
}

function parseInstanceUnits(xml: string) {
  const units = new Map<string, string>();
  for (const match of xml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?unit\b[^>]*\bid=["']([^"']+)["'][^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?unit>/g)) {
    const id = match[1];
    const measures = Array.from(match[2].matchAll(/<(?:[A-Za-z_][\w.-]*:)?measure>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?measure>/g))
      .map((measure) => decodeXml(stripTags(measure[1]).trim()))
      .filter(Boolean);
    if (measures.length) units.set(id, measures.map(unitLocalName).join("/"));
  }
  return units;
}

function parseLabelLinkbase(xml: string): LabelMap {
  const labels: LabelMap = new Map();
  for (const linkMatch of xml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?labelLink\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?labelLink>/g)) {
    const body = linkMatch[1];
    const locators = parseLocators(body);
    const labelResources = new Map<string, ConceptLabel>();
    for (const labelMatch of body.matchAll(/<(?:[A-Za-z_][\w.-]*:)?label\b([^>]*)>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?label>/g)) {
      const attrs = labelMatch[1];
      const label = attr(attrs, "xlink:label") ?? attr(attrs, "label");
      const text = decodeXml(stripTags(labelMatch[2]).trim()).replace(/\s+/g, " ");
      if (!label || !text) continue;
      labelResources.set(label, { text, role: attr(attrs, "xlink:role") ?? attr(attrs, "role") ?? undefined });
    }
    for (const arc of parseArcs(body, "labelArc")) {
      const concept = locators.get(arc.from)?.concept;
      const resource = labelResources.get(arc.to);
      if (!concept || !resource) continue;
      const existing = labels.get(concept) ?? [];
      existing.push(resource);
      labels.set(concept, existing);
    }
  }
  return labels;
}

function parsePresentationLinkbase(xml: string, labels: LabelMap): PresentationStatement[] {
  const statements: PresentationStatement[] = [];
  for (const linkMatch of xml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?presentationLink\b([^>]*)>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?presentationLink>/g)) {
    const linkAttrs = linkMatch[1];
    const body = linkMatch[2];
    const roleUri = attr(linkAttrs, "xlink:role") ?? attr(linkAttrs, "role") ?? undefined;
    const locators = parseLocators(body);
    const arcs = parseArcs(body, "presentationArc");
    const nodesByLabel = new Map<string, PresentationNode>();
    const parentLabels = new Set<string>();
    const childLabels = new Set<string>();

    for (const [label, locator] of locators) {
      nodesByLabel.set(label, {
        concept: locator.concept,
        taxonomy: locator.taxonomy,
        order: 0,
        children: []
      });
    }

    for (const arc of arcs.sort((a, b) => a.order - b.order)) {
      const parent = nodesByLabel.get(arc.from);
      const child = nodesByLabel.get(arc.to);
      if (!parent || !child) continue;
      child.parentConcept = parent.concept;
      child.order = arc.order;
      child.preferredLabelRole = arc.preferredLabelRole;
      parent.children.push(child);
      parentLabels.add(arc.from);
      childLabels.add(arc.to);
    }

    const roots = Array.from(nodesByLabel.entries())
      .filter(([label]) => parentLabels.has(label) && !childLabels.has(label))
      .map(([, node]) => node)
      .sort((a, b) => a.order - b.order || labelForConcept(a.concept, labels).localeCompare(labelForConcept(b.concept, labels)));
    const fallbackRoots = roots.length ? roots : Array.from(nodesByLabel.values()).filter((node) => !node.parentConcept);
    const statementName = statementNameFromRole(roleUri, fallbackRoots[0]?.concept ? labelForConcept(fallbackRoots[0].concept, labels) : undefined);
    statements.push({
      statementName,
      roleUri,
      sourceTableType: classifySourceTableType(statementName, roleUri),
      nodes: flattenPresentationNodes(fallbackRoots)
    });
  }
  return statements;
}

function parseCalculationLinkbase(xml: string): Map<string, CalculationParent[]> {
  const parentsByChild = new Map<string, CalculationParent[]>();
  for (const linkMatch of xml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?calculationLink\b([^>]*)>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?calculationLink>/g)) {
    const roleUri = attr(linkMatch[1], "xlink:role") ?? attr(linkMatch[1], "role") ?? undefined;
    const body = linkMatch[2];
    const locators = parseLocators(body);
    for (const arc of parseArcs(body, "calculationArc")) {
      const parent = locators.get(arc.from);
      const child = locators.get(arc.to);
      if (!parent || !child) continue;
      const existing = parentsByChild.get(child.concept) ?? [];
      existing.push({ parentConcept: parent.concept, weight: arc.weight, roleUri });
      parentsByChild.set(child.concept, existing);
    }
  }
  return parentsByChild;
}

function parseDefinitionLinkbase(xml: string): SecFilingDefinitionRelationship[] {
  const relationships: SecFilingDefinitionRelationship[] = [];
  for (const linkMatch of xml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?definitionLink\b([^>]*)>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?definitionLink>/g)) {
    const roleUri = attr(linkMatch[1], "xlink:role") ?? attr(linkMatch[1], "role") ?? undefined;
    const body = linkMatch[2];
    const locators = parseLocators(body);
    for (const arc of parseArcs(body, "definitionArc")) {
      const from = locators.get(arc.from);
      const to = locators.get(arc.to);
      if (!from || !to) continue;
      relationships.push({ fromConcept: from.concept, toConcept: to.concept, roleUri, arcrole: arc.arcrole });
    }
  }
  return relationships;
}

function parseLocators(body: string) {
  const locators = new Map<string, LinkbaseLocator>();
  for (const match of body.matchAll(/<(?:[A-Za-z_][\w.-]*:)?loc\b([^>]*)\/?>/g)) {
    const attrs = match[1];
    const label = attr(attrs, "xlink:label") ?? attr(attrs, "label");
    const href = attr(attrs, "xlink:href") ?? attr(attrs, "href");
    const concept = href ? conceptFromHref(href) : null;
    if (!label || !concept) continue;
    locators.set(label, { label, concept: concept.concept, taxonomy: concept.taxonomy });
  }
  return locators;
}

function parseArcs(body: string, tagName: string): LinkbaseArc[] {
  const arcs: LinkbaseArc[] = [];
  const pattern = new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${tagName}\\b([^>]*)\\/?>(?:<\\/(?:[A-Za-z_][\\w.-]*:)?${tagName}>)?`, "g");
  for (const match of body.matchAll(pattern)) {
    const attrs = match[1];
    const from = attr(attrs, "xlink:from") ?? attr(attrs, "from");
    const to = attr(attrs, "xlink:to") ?? attr(attrs, "to");
    if (!from || !to) continue;
    arcs.push({
      from,
      to,
      order: Number(attr(attrs, "order") ?? 0),
      preferredLabelRole: attr(attrs, "preferredLabel") ?? undefined,
      weight: attr(attrs, "weight") !== null ? Number(attr(attrs, "weight")) : undefined,
      arcrole: attr(attrs, "xlink:arcrole") ?? attr(attrs, "arcrole") ?? undefined
    });
  }
  return arcs;
}

function parsePrimaryHtmlStatements(
  html: string,
  instance: ParsedInstance,
  labels: LabelMap,
  calculations: Map<string, CalculationParent[]>,
  metadata: {
    accessionNumber: string;
    form?: string;
    filingDate?: string;
    reportingPeriod?: string;
    primaryDocumentUrl?: string;
    definitionRelationships: SecFilingDefinitionRelationship[];
  }
): SecFilingStatementStructure[] {
  const contexts = instance.contexts.size ? instance.contexts : parseInstanceContexts(html);
  const units = instance.units.size ? instance.units : parseInstanceUnits(html);
  const statements: SecFilingStatementStructure[] = [];
  let tableOrder = 0;

  for (const tableMatch of html.matchAll(/<table\b[\s\S]*?<\/table>/gi)) {
    const tableHtml = tableMatch[0];
    if (!/<ix:(?:nonFraction|nonNumeric)\b/i.test(tableHtml)) continue;
    const tableIndex = tableMatch.index ?? 0;
    const name = statementNameFromHtmlContext(html, tableHtml, tableIndex);
    const sourceTableType = classifySourceTableType(`${name} ${htmlText(tableHtml.slice(0, 3000))}`);
    const rows: SecFilingStatementRow[] = [];
    let rowOrder = 0;

    for (const rowMatch of tableHtml.matchAll(/<tr\b[\s\S]*?<\/tr>/gi)) {
      const rowHtml = rowMatch[0];
      const inlineFacts = parseInlineFactsInHtml(rowHtml, contexts, units);
      if (!inlineFacts.length) continue;
      const rowLabel = rowLabelFromHtmlRow(rowHtml) ?? inlineFacts.map((fact) => labelForConcept(fact.concept, labels)).find(Boolean) ?? "";
      for (const fact of inlineFacts) {
        rows.push(
          statementRowFromFact({
            statementName: name,
            sourceTableType,
            rowLabel: cleanStatementLabel(rowLabel) || labelForConcept(fact.concept, labels),
            fact,
            labels,
            calculations,
            metadata,
            sourceUrl: metadata.primaryDocumentUrl,
            rowOrder: tableOrder * 100_000 + rowOrder
          })
        );
      }
      rowOrder += 1;
    }

    if (rows.length) {
      statements.push({
        statementName: name,
        sourceTableType,
        sourceUrl: metadata.primaryDocumentUrl,
        accession: metadata.accessionNumber,
        reportingPeriod: metadata.reportingPeriod,
        form: metadata.form,
        filingDate: metadata.filingDate,
        rows: rows.slice(0, SEC_FILING_PACKAGE_MAX_ROWS_PER_STATEMENT)
      });
      tableOrder += 1;
    }
  }

  return statements;
}

function parseInlineFactsInHtml(
  html: string,
  contexts: Map<string, ParsedContext>,
  units: Map<string, string>
): ParsedFact[] {
  const facts: ParsedFact[] = [];
  for (const match of html.matchAll(/<ix:(nonFraction|nonNumeric)\b([^>]*)>([\s\S]*?)<\/ix:\1>/gi)) {
    const tagKind = match[1];
    const attrs = match[2];
    const name = attr(attrs, "name");
    const concept = name ? conceptFromQualifiedName(name) : null;
    const contextRef = attr(attrs, "contextRef");
    if (!concept || !contextRef) continue;
    const unitRef = attr(attrs, "unitRef") ?? undefined;
    const rawValue = decodeXml(stripTags(match[3]).trim());
    const value = tagKind.toLowerCase() === "nonfraction" ? parseInlineNumber(rawValue, attrs) : rawValue;
    facts.push({
      concept: concept.concept,
      taxonomy: concept.taxonomy,
      contextRef,
      unitRef,
      unit: unitRef ? units.get(unitRef) ?? unitRef : undefined,
      value,
      decimals: attr(attrs, "decimals") ?? undefined,
      rawValue,
      context: contexts.get(contextRef)
    });
  }
  return facts;
}

function buildPresentationStatementStructures(
  presentationStatements: PresentationStatement[],
  instance: ParsedInstance,
  labels: LabelMap,
  calculations: Map<string, CalculationParent[]>,
  metadata: {
    accessionNumber: string;
    form?: string;
    filingDate?: string;
    reportingPeriod?: string;
    primaryDocumentUrl?: string;
    definitionRelationships: SecFilingDefinitionRelationship[];
  }
): SecFilingStatementStructure[] {
  return presentationStatements.map((statement) => {
    const rows: SecFilingStatementRow[] = [];
    statement.nodes.forEach((node, index) => {
      const facts = instance.factsByConcept.get(node.concept) ?? [];
      const label = labelForConcept(node.concept, labels, node.preferredLabelRole);
      if (!facts.length) {
        rows.push({
          statementName: statement.statementName,
          sourceTableType: statement.sourceTableType,
          rowLabel: label,
          xbrlConcept: node.concept,
          taxonomy: node.taxonomy,
          value: null,
          period: {},
          consolidated: true,
          dimensions: [],
          currentNonCurrentSection: inferCurrentNonCurrentSection([label, node.concept, node.parentConcept ?? ""]),
          rowOrder: index,
          parentSubtotal: presentationParentSubtotal(node, labels),
          accession: metadata.accessionNumber,
          reportingPeriod: metadata.reportingPeriod,
          sourceUrl: metadata.primaryDocumentUrl
        });
        return;
      }
      facts.forEach((fact, factIndex) => {
        rows.push(
          statementRowFromFact({
            statementName: statement.statementName,
            sourceTableType: statement.sourceTableType,
            rowLabel: label,
            fact,
            labels,
            calculations,
            metadata,
            roleUri: statement.roleUri,
            sourceUrl: metadata.primaryDocumentUrl,
            rowOrder: index + factIndex / 1000,
            presentationParent: presentationParentSubtotal(node, labels)
          })
        );
      });
    });
    return {
      statementName: statement.statementName,
      sourceTableType: statement.sourceTableType,
      roleUri: statement.roleUri,
      sourceUrl: metadata.primaryDocumentUrl,
      accession: metadata.accessionNumber,
      reportingPeriod: metadata.reportingPeriod,
      form: metadata.form,
      filingDate: metadata.filingDate,
      rows: rows.slice(0, SEC_FILING_PACKAGE_MAX_ROWS_PER_STATEMENT)
    };
  });
}

function statementRowFromFact(input: {
  statementName: string;
  sourceTableType: SecFilingStatementSourceTableType;
  rowLabel: string;
  fact: ParsedFact;
  labels: LabelMap;
  calculations: Map<string, CalculationParent[]>;
  metadata: {
    accessionNumber: string;
    reportingPeriod?: string;
  };
  roleUri?: string;
  sourceUrl?: string;
  rowOrder: number;
  presentationParent?: SecFilingParentSubtotal;
}): SecFilingStatementRow {
  const context = input.fact.context;
  const dimensions = context?.dimensions ?? [];
  const parentSubtotal = calculationParentSubtotal(input.fact.concept, input.calculations, input.labels, input.roleUri) ?? input.presentationParent;
  return {
    statementName: input.statementName,
    sourceTableType: input.sourceTableType,
    rowLabel: input.rowLabel || labelForConcept(input.fact.concept, input.labels),
    xbrlConcept: input.fact.concept,
    taxonomy: input.fact.taxonomy,
    value: input.fact.value,
    unit: input.fact.unit,
    period: {
      contextRef: input.fact.contextRef,
      start: context?.start,
      end: context?.end,
      instant: context?.instant,
      periodType: context?.periodType
    },
    consolidated: dimensions.length === 0,
    dimensions,
    currentNonCurrentSection: inferCurrentNonCurrentSection([
      input.rowLabel,
      input.fact.concept,
      parentSubtotal?.label ?? "",
      parentSubtotal?.concept ?? ""
    ]),
    rowOrder: input.rowOrder,
    parentSubtotal,
    accession: input.metadata.accessionNumber,
    reportingPeriod: input.metadata.reportingPeriod,
    sourceUrl: input.sourceUrl
  };
}

function calculationParentSubtotal(
  concept: string,
  calculations: Map<string, CalculationParent[]>,
  labels: LabelMap,
  roleUri?: string
): SecFilingParentSubtotal | undefined {
  const candidates = calculations.get(concept) ?? [];
  if (!candidates.length) return undefined;
  const preferred = candidates.find((item) => item.roleUri === roleUri) ?? candidates[0];
  return {
    concept: preferred.parentConcept,
    label: labelForConcept(preferred.parentConcept, labels),
    relationship: "calculation",
    weight: preferred.weight,
    roleUri: preferred.roleUri
  };
}

function presentationParentSubtotal(node: PresentationNode, labels: LabelMap): SecFilingParentSubtotal | undefined {
  if (!node.parentConcept) return undefined;
  return {
    concept: node.parentConcept,
    label: labelForConcept(node.parentConcept, labels),
    relationship: "presentation"
  };
}

function flattenPresentationNodes(roots: PresentationNode[]) {
  const flattened: PresentationNode[] = [];
  const visit = (node: PresentationNode) => {
    flattened.push(node);
    node.children.sort((a, b) => a.order - b.order).forEach(visit);
  };
  roots.sort((a, b) => a.order - b.order).forEach(visit);
  return flattened;
}

function labelForConcept(concept: string, labels: LabelMap, preferredRole?: string) {
  const candidates = labels.get(concept) ?? [];
  const preferred = preferredRole ? candidates.find((item) => item.role === preferredRole) : undefined;
  const sorted = [...candidates].sort((a, b) => labelRoleScore(b.role) - labelRoleScore(a.role));
  return preferred?.text ?? sorted[0]?.text ?? humanizeConcept(concept);
}

function labelRoleScore(role?: string) {
  if (!role) return 0;
  if (/terseLabel$/i.test(role)) return 7;
  if (/totalLabel$/i.test(role)) return 6;
  if (/periodEndLabel$/i.test(role)) return 5;
  if (/periodStartLabel$/i.test(role)) return 5;
  if (/negatedTerseLabel$/i.test(role)) return 4;
  if (/label$/i.test(role)) return 3;
  return 1;
}

function statementNameFromRole(roleUri?: string, fallback?: string) {
  if (!roleUri) return fallback ?? "SEC Filing Statement";
  const last = decodeURIComponent(roleUri.split(/[/#]/).filter(Boolean).at(-1) ?? roleUri);
  const withoutPrefix = last.replace(/^\d+\s*[-_]\s*/, "");
  return humanizeConcept(withoutPrefix.replace(/[_-]/g, " ")) || fallback || "SEC Filing Statement";
}

function statementNameFromHtmlContext(html: string, tableHtml: string, tableIndex: number) {
  const before = htmlText(html.slice(Math.max(0, tableIndex - 2500), tableIndex));
  const tableText = htmlText(tableHtml.slice(0, 2500));
  const source = `${before} ${tableText}`;
  const direct = source.match(/\b(?:consolidated\s+)?(?:balance sheets?|statements?\s+of\s+(?:operations|income|earnings|comprehensive income|cash flows|stockholders'? equity|financial position)|income statements?|cash flow statements?)[^.;]{0,120}/i);
  if (direct) return cleanStatementLabel(direct[0]);
  return cleanStatementLabel(tableText.slice(0, 120)) || "SEC Filing Table";
}

function classifySourceTableType(text: string, roleUri = ""): SecFilingStatementSourceTableType {
  const haystack = `${text} ${roleUri}`.toLowerCase();
  if (/\b(segment|reportable segment|geographic|disaggregation of revenue|external customers)\b/.test(haystack)) return "segment_table";
  if (/\b(rollforward|roll forward|changes in|schedule of)\b/.test(haystack)) return "roll_forward";
  if (/\b(balance sheets?|statements? of operations|statements? of income|income statements?|statements? of cash flows|cash flow statements?|statements? of stockholders|statements? of financial position|comprehensive income)\b/.test(haystack)) return "primary_statement";
  if (/\b(note|notes to|supplemental|schedule)\b/.test(haystack)) return "footnote";
  if (/\b(parenthetical|detail|document and entity information)\b/.test(haystack)) return "support_table";
  return "unknown";
}

function rowLabelFromHtmlRow(rowHtml: string) {
  const cells = Array.from(rowHtml.matchAll(/<t[dh]\b[\s\S]*?<\/t[dh]>/gi))
    .map((match) => cleanStatementLabel(htmlText(match[0])))
    .filter(isUsefulStatementLabel);
  return cells[0] ?? null;
}

function isUsefulStatementLabel(value: string) {
  if (!value || value.length > 180) return false;
  if (!/[A-Za-z]/.test(value)) return false;
  if (/^(?:\$|shares?|amounts?|millions?|thousands?|year|quarter|three months|six months|nine months|unaudited)$/i.test(value)) return false;
  if (/^\(?-?\$?\d[\d,]*(?:\.\d+)?\)?$/.test(value)) return false;
  return true;
}

function inferCurrentNonCurrentSection(parts: string[]): "current" | "non_current" | undefined {
  const text = parts.join(" ").toLowerCase();
  if (/\b(noncurrent|non-current|long-term|long term)\b/.test(text)) return "non_current";
  if (/\bcurrent\b/.test(text)) return "current";
  return undefined;
}

async function fetchSecJson(url: string, headers: Record<string, string>) {
  let cached = responseJsonCache.get(url);
  if (!cached) {
    cached = fetchSecText(url, headers, "application/json").then((text) => {
      if (!text) return null;
      try {
        return JSON.parse(text);
      } catch {
        return null;
      }
    });
    responseJsonCache.set(url, cached);
  }
  return cached;
}

async function fetchSecText(url: string, headers: Record<string, string>, accept: string) {
  let cached = responseTextCache.get(url);
  if (!cached) {
    cached = fetchSecTextUncached(url, headers, accept);
    responseTextCache.set(url, cached);
  }
  return cached;
}

async function fetchSecTextUncached(url: string, headers: Record<string, string>, accept: string) {
  for (let attempt = 0; attempt < 3; attempt += 1) {
    await throttleSecArchiveFetch();
    const response = await fetch(url, { headers: secHeaders(headers, accept) });
    if (response.ok) return response.text();
    if (response.status !== 429) return null;
    await sleep(750 * (attempt + 1));
  }
  return null;
}

function secHeaders(headers: Record<string, string>, accept: string) {
  return {
    ...headers,
    "User-Agent": headers["User-Agent"] || SEC_DEFAULT_USER_AGENT,
    Accept: accept
  };
}

async function throttleSecArchiveFetch() {
  const queued = secArchiveFetchQueue.then(async () => {
    const waitMs = Math.max(0, SEC_ARCHIVE_MIN_INTERVAL_MS - (Date.now() - lastSecArchiveFetchAt));
    if (waitMs > 0) await sleep(waitMs);
    lastSecArchiveFetchAt = Date.now();
  });
  secArchiveFetchQueue = queued.catch(() => undefined);
  await queued;
}

function sleep(ms: number) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function conceptFromHref(href: string): ParsedConceptRef | null {
  const fragment = href.split("#").pop();
  if (!fragment) return null;
  const clean = decodeXml(fragment);
  if (clean.includes(":")) return conceptFromQualifiedName(clean);
  const underscore = clean.indexOf("_");
  if (underscore > 0) {
    return {
      taxonomy: clean.slice(0, underscore),
      concept: clean.slice(underscore + 1),
      raw: clean
    };
  }
  return { concept: clean, raw: clean };
}

function conceptFromQualifiedName(value: string): ParsedConceptRef {
  const clean = decodeXml(value.trim());
  const parts = clean.split(":");
  if (parts.length > 1) return { taxonomy: parts[0], concept: parts.slice(1).join(":"), raw: clean };
  return { concept: clean, raw: clean };
}

function attr(input: string, name: string) {
  const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const pattern = new RegExp(`(?:^|\\s)${escaped}=["']([^"']*)["']`, "i");
  return input.match(pattern)?.[1] ?? null;
}

function textContent(xml: string, localName: string) {
  const pattern = new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${localName}\\b[^>]*>([\\s\\S]*?)<\\/(?:[A-Za-z_][\\w.-]*:)?${localName}>`, "i");
  const match = xml.match(pattern);
  return match ? decodeXml(stripTags(match[1]).trim()) : null;
}

function parseXmlFactValue(raw: string): number | string | null {
  if (!raw || raw === "-") return null;
  const numeric = Number(raw.replace(/,/g, ""));
  return Number.isFinite(numeric) ? numeric : raw;
}

function parseInlineNumber(raw: string, attrs: string) {
  const clean = raw.replace(/,/g, "").trim();
  if (!clean || clean === "-") return null;
  const numeric = Number(clean.replace(/[()]/g, ""));
  if (!Number.isFinite(numeric)) return null;
  const scale = Number(attr(attrs, "scale") ?? 0);
  const sign = attr(attrs, "sign") === "-" || /^\(.+\)$/.test(clean) ? -1 : 1;
  return numeric * Math.pow(10, scale) * sign;
}

function stripTags(value: string) {
  return value.replace(/<[^>]+>/g, " ");
}

function htmlText(html: string) {
  return decodeXml(
    html
      .replace(/<script[\s\S]*?<\/script>/gi, " ")
      .replace(/<style[\s\S]*?<\/style>/gi, " ")
      .replace(/<\/(?:tr|p|div|table|h\d)>/gi, " ")
      .replace(/<br\s*\/?>/gi, " ")
      .replace(/<[^>]+>/g, " ")
  )
    .replace(/\u00a0/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function cleanStatementLabel(value: string) {
  return value
    .replace(/\s+/g, " ")
    .replace(/\bTable of Contents\b/gi, "")
    .replace(/\s*\((?:in\s+)?(?:millions|thousands|USD \$|shares|except per share data)\)\s*$/i, "")
    .replace(/^[,;:\- ]+|[,;:\- ]+$/g, "")
    .trim();
}

function humanizeConcept(concept: string) {
  return concept
    .replace(/([a-z])([A-Z])/g, "$1 $2")
    .replace(/\bAnd\b/g, "and")
    .replace(/\bOf\b/g, "of")
    .replace(/\bFor\b/g, "for")
    .replace(/\bNet\b/g, "net")
    .replace(/\s+/g, " ")
    .trim();
}

function decodeXml(value: string) {
  return value
    .replace(/&amp;/g, "&")
    .replace(/&nbsp;/g, " ")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, "\"")
    .replace(/&apos;/g, "'")
    .replace(/&#x([0-9a-f]+);/gi, (_, hex) => String.fromCharCode(parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (_, decimal) => String.fromCharCode(Number(decimal)));
}

function unitLocalName(value: string) {
  return value.includes(":") ? value.split(":").pop() ?? value : value;
}

function isHtmlFile(name: string) {
  return /\.(?:htm|html)$/i.test(name);
}

function isGeneratedFilingSupportFile(name: string) {
  return /^(?:FilingSummary|MetaLinks|Financial_Report)\.xml$/i.test(name) || /\.(?:xsd|jpg|jpeg|png|gif|css|js)$/i.test(name);
}

function uniqueFilingsByAccession(filings: SecFilingPackageRequest[]) {
  const seen = new Set<string>();
  const result: SecFilingPackageRequest[] = [];
  for (const filing of filings) {
    const key = `${filing.cik}:${normalizeAccession(filing.accessionNumber)}`;
    if (seen.has(key)) continue;
    seen.add(key);
    result.push(filing);
  }
  return result;
}

function unique<T>(items: T[]) {
  return Array.from(new Set(items));
}
