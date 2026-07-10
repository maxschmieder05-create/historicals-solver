#!/usr/bin/env node

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const Module = require("node:module");
const ts = require("typescript");
const ExcelJS = require("exceljs");

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
  const compiled = compileTypeScript(fs.readFileSync(file, "utf8"));
  const mod = new Module(file, module);
  mod.filename = file;
  mod.paths = Module._nodeModulePaths(path.dirname(file));
  mod._compile(compiled, file);
  return mod.exports;
}

const { __fillModelServiceTestHooks } = loadTypeScriptModule(
  path.join(__dirname, "..", "server", "fill-model", "fill-model-service.ts")
);

const { normalizeKnownExcelFormula, normalizeKnownExcelFormulaPrefixes } = __fillModelServiceTestHooks;

assert.equal(
  normalizeKnownExcelFormula('_xludf.IFS(A1>0,"yes",A1=0,"zero",TRUE,"no")'),
  'IFS(A1>0,"yes",A1=0,"zero",TRUE,"no")'
);
assert.equal(normalizeKnownExcelFormula("_xludf.XLOOKUP(A1,B:B,C:C)"), "XLOOKUP(A1,B:B,C:C)");
assert.equal(
  normalizeKnownExcelFormula("_xludf.MyCustomAccountingFunction(A1)"),
  "_xludf.MyCustomAccountingFunction(A1)",
  "unknown user-defined functions must remain untouched"
);

const workbook = new ExcelJS.Workbook();
const sheet = workbook.addWorksheet("Valuation");
sheet.getCell("A1").value = {
  formula: "_xludf.IFS(B1=1,10,TRUE,20)",
  result: { error: "#NAME?" }
};
sheet.getCell("A2").value = {
  formula: "_xludf.MyCustomAccountingFunction(B2)",
  result: 7
};

assert.equal(normalizeKnownExcelFormulaPrefixes(workbook), 1);
assert.deepEqual(sheet.getCell("A1").value, { formula: "IFS(B1=1,10,TRUE,20)" });
assert.deepEqual(sheet.getCell("A2").value, {
  formula: "_xludf.MyCustomAccountingFunction(B2)",
  result: 7
});

console.log("Known Excel formula-prefix normalization checks passed.");
