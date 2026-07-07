const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const Module = require("node:module");
const ts = require("typescript");

const repoRoot = path.resolve(__dirname, "..");

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

const { isPrimaryBalanceSheetStructure } = loadTypeScriptModule(path.join(repoRoot, "server", "fill-model", "fill-model-service.ts"));

function statement(statementName) {
  return {
    statementName,
    sourceTableType: "primary_statement",
    accession: "0000000000-26-000001",
    rows: []
  };
}

assert.equal(
  isPrimaryBalanceSheetStructure(statement("CONSOLIDATED BALANCE SHEET - (CONTINUED) (UNAUDITED) LIABILITIES AND EQUITY")),
  true
);
assert.equal(isPrimaryBalanceSheetStructure(statement("CONSOLIDATED BALANCE SHEET ASSETS")), true);
assert.equal(isPrimaryBalanceSheetStructure(statement("CONSOLIDATED STATEMENT OF STOCKHOLDERS' EQUITY")), false);
assert.equal(isPrimaryBalanceSheetStructure(statement("CONSOLIDATED STATEMENTS OF CASH FLOWS")), false);
assert.equal(isPrimaryBalanceSheetStructure(statement("BALANCE SHEET PARENTHETICAL")), false);

console.log("Primary balance-sheet structure regression passed.");
