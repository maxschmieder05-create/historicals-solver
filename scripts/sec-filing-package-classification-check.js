const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const Module = require("node:module");
const ts = require("typescript");

const repoRoot = path.resolve(__dirname, "..");
const sourcePath = path.join(repoRoot, "server", "fill-model", "sec-filing-package.ts");

function compileTypeScript(source) {
  return ts.transpileModule(source, {
    compilerOptions: {
      module: ts.ModuleKind.CommonJS,
      target: ts.ScriptTarget.ES2020,
      esModuleInterop: true
    }
  }).outputText;
}

function registerTypeScriptRequire() {
  if (require.extensions[".ts"]) return;
  require.extensions[".ts"] = (mod, file) => {
    mod._compile(compileTypeScript(fs.readFileSync(file, "utf8")), file);
  };
}

function loadTypeScriptModule(file) {
  registerTypeScriptRequire();
  const source = fs.readFileSync(file, "utf8");
  const compiled = compileTypeScript(source);
  const mod = new Module(file, module);
  mod.filename = file;
  mod.paths = Module._nodeModulePaths(path.dirname(file));
  mod._compile(compiled, file);
  return mod.exports;
}

const { classifySourceTableType, __secFilingPackageTestHooks: hooks } = loadTypeScriptModule(sourcePath);

assert.equal(classifySourceTableType("CONSOLIDATED STATEMENTS OF EARNINGS (Unaudited)"), "primary_statement");
assert.equal(classifySourceTableType("Condensed consolidated earnings statement"), "primary_statement");
assert.equal(classifySourceTableType("CONSOLIDATED STATEMENTS OF CASH FLOWS (Unaudited)"), "primary_statement");
assert.equal(classifySourceTableType("Schedule of restructuring charges"), "roll_forward");

const duplicatePrecisionInstance = hooks.parseInstanceXml(`
  <xbrli:xbrl xmlns:xbrli="http://www.xbrl.org/2003/instance" xmlns:us-gaap="http://fasb.org/us-gaap/2024">
    <xbrli:context id="c-1"><xbrli:entity><xbrli:identifier scheme="test">1</xbrli:identifier></xbrli:entity><xbrli:period><xbrli:instant>2024-12-31</xbrli:instant></xbrli:period></xbrli:context>
    <xbrli:unit id="usd"><xbrli:measure>iso4217:USD</xbrli:measure></xbrli:unit>
    <us-gaap:DeferredIncomeTaxAssetsNet contextRef="c-1" unitRef="usd" decimals="-8">1400000000</us-gaap:DeferredIncomeTaxAssetsNet>
    <us-gaap:DeferredIncomeTaxAssetsNet contextRef="c-1" unitRef="usd" decimals="-3">1440418000</us-gaap:DeferredIncomeTaxAssetsNet>
  </xbrli:xbrl>
`);
assert.deepEqual(
  duplicatePrecisionInstance.factsByConcept.get("DeferredIncomeTaxAssetsNet").map((fact) => fact.value),
  [1_440_418_000],
  "When a filing repeats one concept/context at different rounded precisions, the most precise SEC fact must be selected once."
);

const conflictingEqualPrecisionInstance = hooks.parseInstanceXml(`
  <xbrli:xbrl xmlns:xbrli="http://www.xbrl.org/2003/instance" xmlns:us-gaap="http://fasb.org/us-gaap/2024">
    <xbrli:context id="c-1"><xbrli:entity><xbrli:identifier scheme="test">1</xbrli:identifier></xbrli:entity><xbrli:period><xbrli:instant>2024-12-31</xbrli:instant></xbrli:period></xbrli:context>
    <xbrli:unit id="usd"><xbrli:measure>iso4217:USD</xbrli:measure></xbrli:unit>
    <us-gaap:Assets contextRef="c-1" unitRef="usd" decimals="-3">1000000</us-gaap:Assets>
    <us-gaap:Assets contextRef="c-1" unitRef="usd" decimals="-3">1100000</us-gaap:Assets>
  </xbrli:xbrl>
`);
assert.equal(
  conflictingEqualPrecisionInstance.factsByConcept.get("Assets").length,
  2,
  "Equally precise but conflicting SEC facts must remain visible for fail-closed downstream validation."
);

const opaqueSegmentRole = "https://issuer.example/role/R42";
const filingBaseUrl = "https://www.sec.gov/Archives/edgar/data/1/000000000125000001";
const roleMetadata = hooks.parseFilingRoleMetadata(
  `
    <xsd:schema xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:link="http://www.xbrl.org/2003/linkbase">
      <link:roleType id="R42" roleURI="${opaqueSegmentRole}">
        <link:definition>9955585 - Disclosure - Segment Reporting - Regional Results (Details)</link:definition>
      </link:roleType>
    </xsd:schema>
  `,
  `
    <FilingSummary><MyReports><Report>
      <HtmlFileName>R42.htm</HtmlFileName>
      <LongName>9955585 - Disclosure - Segment Reporting - Regional Results (Details)</LongName>
      <Role>${opaqueSegmentRole}</Role>
      <ShortName>Segment Reporting - Regional Results (Details)</ShortName>
      <MenuCategory>Details</MenuCategory>
      <Position>42</Position>
    </Report></MyReports></FilingSummary>
  `,
  filingBaseUrl
);

const catalogLabels = hooks.parseLabelLinkbase(`
  <link:linkbase xmlns:link="http://www.xbrl.org/2003/linkbase" xmlns:xlink="http://www.w3.org/1999/xlink">
    <link:labelLink>
      <link:loc xlink:label="loc_root" xlink:href="test.xsd#us-gaap_SegmentReportingInformationLineItems" />
      <link:loc xlink:label="loc_revenue" xlink:href="test.xsd#us-gaap_RevenueFromContractWithCustomerExcludingAssessedTax" />
      <link:loc xlink:label="loc_axis" xlink:href="test.xsd#us-gaap_StatementBusinessSegmentsAxis" />
      <link:loc xlink:label="loc_group" xlink:href="test.xsd#test_Group1SegmentMember" />
      <link:label xlink:label="lab_root" xlink:role="http://www.xbrl.org/2003/role/terseLabel">Segment information</link:label>
      <link:label xlink:label="lab_revenue" xlink:role="http://www.xbrl.org/2003/role/terseLabel">Revenue</link:label>
      <link:label xlink:label="lab_revenue" xlink:role="http://www.xbrl.org/2003/role/verboseLabel">Net revenue</link:label>
      <link:label xlink:label="lab_axis" xlink:role="http://www.xbrl.org/2003/role/terseLabel">Operating segments</link:label>
      <link:label xlink:label="lab_group" xlink:role="http://www.xbrl.org/2003/role/terseLabel">Group 1</link:label>
      <link:label xlink:label="lab_group" xlink:role="http://www.xbrl.org/2003/role/label">Group 1 Segment [Member]</link:label>
      <link:labelArc xlink:from="loc_root" xlink:to="lab_root" />
      <link:labelArc xlink:from="loc_revenue" xlink:to="lab_revenue" />
      <link:labelArc xlink:from="loc_axis" xlink:to="lab_axis" />
      <link:labelArc xlink:from="loc_group" xlink:to="lab_group" />
    </link:labelLink>
  </link:linkbase>
`);
assert.equal(catalogLabels.get("RevenueFromContractWithCustomerExcludingAssessedTax").length, 2);
assert.equal(hooks.labelForConcept("Group1SegmentMember", catalogLabels), "Group 1");

const catalogPresentation = hooks.parsePresentationLinkbase(
  `
    <link:linkbase xmlns:link="http://www.xbrl.org/2003/linkbase" xmlns:xlink="http://www.w3.org/1999/xlink">
      <link:presentationLink xlink:role="${opaqueSegmentRole}">
        <link:loc xlink:label="root" xlink:href="test.xsd#us-gaap_SegmentReportingInformationLineItems" />
        <link:loc xlink:label="revenue" xlink:href="test.xsd#us-gaap_RevenueFromContractWithCustomerExcludingAssessedTax" />
        <link:presentationArc xlink:from="root" xlink:to="revenue" order="1" preferredLabel="http://www.xbrl.org/2003/role/verboseLabel" />
      </link:presentationLink>
    </link:linkbase>
  `,
  catalogLabels,
  roleMetadata
);
assert.equal(catalogPresentation.length, 1);
assert.equal(catalogPresentation[0].sourceTableType, "segment_table");
assert.equal(catalogPresentation[0].statementName, "Segment Reporting - Regional Results (Details)");
assert.equal(catalogPresentation[0].sourceUrl, `${filingBaseUrl}/R42.htm`);

const catalogInstance = hooks.parseInstanceXml(`
  <xbrli:xbrl xmlns:xbrli="http://www.xbrl.org/2003/instance" xmlns:xbrldi="http://xbrl.org/2006/xbrldi" xmlns:us-gaap="http://fasb.org/us-gaap/2025">
    <xbrli:context id="segment-2025">
      <xbrli:entity><xbrli:identifier scheme="test">1</xbrli:identifier><xbrli:segment>
        <xbrldi:explicitMember dimension="us-gaap:StatementBusinessSegmentsAxis">test:Group1SegmentMember</xbrldi:explicitMember>
      </xbrli:segment></xbrli:entity>
      <xbrli:period><xbrli:startDate>2025-01-01</xbrli:startDate><xbrli:endDate>2025-12-31</xbrli:endDate></xbrli:period>
    </xbrli:context>
    <xbrli:unit id="usd"><xbrli:measure>iso4217:USD</xbrli:measure></xbrli:unit>
    <us-gaap:RevenueFromContractWithCustomerExcludingAssessedTax contextRef="segment-2025" unitRef="usd" decimals="-3">7509000000</us-gaap:RevenueFromContractWithCustomerExcludingAssessedTax>
  </xbrli:xbrl>
`);
const catalogStatements = hooks.buildPresentationStatementStructures(
  catalogPresentation,
  catalogInstance,
  catalogLabels,
  new Map(),
  {
    accessionNumber: "0000000001-25-000001",
    form: "10-K",
    filingDate: "2026-02-01",
    reportingPeriod: "2025-12-31",
    primaryDocumentUrl: `${filingBaseUrl}/issuer-20251231.htm`,
    definitionRelationships: []
  }
);
const catalogRevenue = catalogStatements[0].rows.find((row) => row.value === 7_509_000_000);
assert.ok(catalogRevenue);
assert.equal(catalogRevenue.rowLabel, "Net revenue");
assert.equal(catalogRevenue.unit, "USD");
assert.deepEqual(catalogRevenue.period, {
  contextRef: "segment-2025",
  start: "2025-01-01",
  end: "2025-12-31",
  instant: undefined,
  periodType: "duration"
});
assert.equal(catalogRevenue.dimensions[0].dimensionLabel, "Operating segments");
assert.equal(catalogRevenue.dimensions[0].memberLabel, "Group 1");
assert.equal(catalogRevenue.sourceUrl, `${filingBaseUrl}/R42.htm`);

const discoveredArtifacts = hooks.discoverFilingArtifacts(
  [
    { name: "issuer-20251231.htm" },
    { name: "issuer-20251231.xsd" },
    { name: "FilingSummary.xml" },
    { name: "issuer-20251231_htm.xml" },
    { name: "issuer-20251231_pre.xml" },
    { name: "issuer-20251231_lab.xml" }
  ],
  { cik: "1", accessionNumber: "0000000001-25-000001", primaryDocument: "issuer-20251231.htm" },
  filingBaseUrl
);
assert.equal(discoveredArtifacts.schema.name, "issuer-20251231.xsd");
assert.equal(discoveredArtifacts.filing_summary.name, "FilingSummary.xml");
assert.equal(discoveredArtifacts.instance.name, "issuer-20251231_htm.xml");

console.log("SEC filing package statement classification rules passed.");
