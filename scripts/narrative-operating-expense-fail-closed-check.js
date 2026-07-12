const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const Module = require("node:module");
const ts = require("typescript");

function compileTypeScript(source) {
  return ts.transpileModule(source, {
    compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020, esModuleInterop: true }
  }).outputText;
}

function loadTypeScriptModule(file) {
  if (!require.extensions[".ts"]) {
    require.extensions[".ts"] = (mod, filename) => mod._compile(compileTypeScript(fs.readFileSync(filename, "utf8")), filename);
  }
  const mod = new Module(file, module);
  mod.filename = file;
  mod.paths = Module._nodeModulePaths(path.dirname(file));
  mod._compile(compileTypeScript(fs.readFileSync(file, "utf8")), file);
  return mod.exports;
}

const sourcePath = path.join(__dirname, "..", "server", "fill-model", "fill-model-service.ts");
const hooks = loadTypeScriptModule(sourcePath).__fillModelServiceTestHooks;

const filing = {
  accessionNumber: "0000000000-26-000001",
  filingDate: "2026-05-01",
  form: "10-Q",
  primaryDocument: "company-20260331.htm",
  reportDate: "2026-03-31"
};

function emptyContext() {
  return { duration: new Map(), instant: new Map() };
}

function durationContext(id, start, end) {
  return `<xbrli:context id="${id}"><xbrli:entity><xbrli:identifier scheme="https://www.sec.gov/CIK">0000000000</xbrli:identifier></xbrli:entity><xbrli:period><xbrli:startDate>${start}</xbrli:startDate><xbrli:endDate>${end}</xbrli:endDate></xbrli:period></xbrli:context>`;
}

const usdUnit = `<xbrli:unit id="money"><xbrli:measure>iso4217:USD</xbrli:measure></xbrli:unit>`;
const contexts = `${durationContext("current", "2026-01-01", "2026-03-31")}${durationContext("prior", "2025-01-01", "2025-03-31")}`;

// The display columns intentionally put the current period first and the prior
// period second. The deleted narrative parser always chose the second number.
// Authoritative inline facts instead follow their XBRL contexts, independent of
// presentation order, and preserve the issuer's actual extension concept.
const taggedReversedColumns = `${contexts}${usdUnit}
  <table>
    <tr><th>Operating expenses</th><th>Three months ended March 31, 2026</th><th>Three months ended March 31, 2025</th></tr>
    <tr><td>Fulfillment</td>
      <td><ix:nonFraction name="acme:CloudFulfillmentExpense" contextRef="current" unitRef="money" scale="6">125</ix:nonFraction></td>
      <td><ix:nonFraction name="acme:CloudFulfillmentExpense" contextRef="prior" unitRef="money" scale="6">80</ix:nonFraction></td>
    </tr>
    <tr><td>Total operating expenses</td><td>125</td><td>80</td></tr>
  </table>`;
const taggedContext = emptyContext();
hooks.mergeInlineFacts(taggedReversedColumns, filing, new Set(["1Q25", "1Q26"]), taggedContext, "https://www.sec.gov/example");
assert.equal(taggedContext.duration.get("1Q26")?.get("CloudFulfillmentExpense")?.value, 125_000_000);
assert.equal(taggedContext.duration.get("1Q25")?.get("CloudFulfillmentExpense")?.value, 80_000_000);
assert.equal(taggedContext.duration.get("1Q26")?.has("FulfillmentExpense"), false, "the parser must not invent a standard SEC concept from a row label");
assert.equal(taggedContext.duration.get("1Q26")?.get("CloudFulfillmentExpense")?.unit, "USD");

// A visually plausible narrative table is not a numeric source when its cells
// are untagged and its scale is ambiguous.
const ambiguousNarrative = `
  <table>
    <caption>Operating expenses (amounts may be in millions or thousands)</caption>
    <tr><th></th><th>March 31, 2026</th><th>March 31, 2025</th></tr>
    <tr><td>Fulfillment</td><td>125</td><td>80</td></tr>
    <tr><td>Total operating expenses</td><td>125</td><td>80</td></tr>
  </table>`;
const ambiguousContext = emptyContext();
hooks.mergeInlineFacts(ambiguousNarrative, filing, new Set(["1Q26"]), ambiguousContext);
assert.equal(ambiguousContext.duration.size, 0, "untagged narrative numbers must fail closed");

// UnitRef names and prose labels are not enough: the referenced XBRL unit must
// resolve to authoritative USD/shares semantics, and scale must be a valid XBRL
// integer rather than a human-language guess.
const unresolvedUnit = `${contexts}
  <table><tr><td>Fulfillment</td><td><ix:nonFraction name="acme:CloudFulfillmentExpense" contextRef="current" unitRef="USD" scale="6">125</ix:nonFraction></td></tr></table>`;
const unresolvedUnitContext = emptyContext();
hooks.mergeInlineFacts(unresolvedUnit, filing, new Set(["1Q26"]), unresolvedUnitContext);
assert.equal(unresolvedUnitContext.duration.size, 0, "a USD-looking unitRef without an XBRL unit definition must fail closed");

const invalidScale = `${contexts}${usdUnit}
  <table><tr><td>Fulfillment</td><td><ix:nonFraction name="acme:CloudFulfillmentExpense" contextRef="current" unitRef="money" scale="millions">125</ix:nonFraction></td></tr></table>`;
const invalidScaleContext = emptyContext();
hooks.mergeInlineFacts(invalidScale, filing, new Set(["1Q26"]), invalidScaleContext);
assert.equal(invalidScaleContext.duration.size, 0, "an ambiguous textual scale must fail closed");

const mislabeledSharesUnit = `${contexts}<xbrli:unit id="USD"><xbrli:measure>xbrli:shares</xbrli:measure></xbrli:unit>
  <ix:nonFraction name="acme:CloudFulfillmentExpense" contextRef="current" unitRef="USD" scale="6">125</ix:nonFraction>`;
const mislabeledSharesContext = emptyContext();
hooks.mergeInlineFacts(mislabeledSharesUnit, filing, new Set(["1Q26"]), mislabeledSharesContext);
assert.equal(mislabeledSharesContext.duration.size, 0, "an expense fact measured in shares must not be accepted because its unitRef is named USD");

assert.deepEqual(Array.from(hooks.parseInlineFactUnits(`<xbrli:unit id="fakeUsd"><xbrli:measure>xbrli:shares</xbrli:measure></xbrli:unit>`)), [["fakeUsd", "shares"]]);
assert.equal(hooks.ixNumber("125", ' scale="millions"'), null);

console.log("Narrative operating-expense extraction fails closed unless authoritative inline XBRL semantics are present.");
