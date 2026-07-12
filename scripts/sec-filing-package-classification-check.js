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

console.log("SEC filing package statement classification rules passed.");
