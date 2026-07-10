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
const {
  writeSegmentFourthQuarterBridgeFormula,
  writeSegmentMetricPreservingFormula,
  reconcileSegmentMetricRowsToModelRow,
  writeHistoricalEbitdaDaAddback,
  findSegmentResidualRow
} = __fillModelServiceTestHooks;

const workbook = new ExcelJS.Workbook();
const sheet = workbook.addWorksheet("Segment Analysis");
sheet.getCell("C8").value = "Foods and Sundries Revenue";
sheet.getCell("F8").value = 10;
sheet.getCell("G8").value = 20;
sheet.getCell("H8").value = 30;
sheet.getCell("I8").value = { formula: "70-SUM(F8:H8)", result: 10 };
sheet.getCell("J8").value = { formula: "SUM(F8:I8)", result: 70 };

const periods = ["1Q23", "2Q23", "3Q23", "4Q23", "FY23"];
const columns = [6, 7, 8, 9, 10];
const segment = {
  label: "Foods and Sundries",
  family: "merchandise_category",
  values: new Map([
    ["4Q23", 40_000_000],
    ["FY23", 100_000_000]
  ]),
  operatingIncome: new Map(),
  depreciationAmortization: new Map(),
  annualValues: new Map([["FY23", 70_000_000]])
};

assert.equal(writeSegmentFourthQuarterBridgeFormula(sheet, 8, columns, periods, 3, segment, "values", 40), true);
assert.deepEqual(sheet.getCell("I8").value, { formula: "100-SUM(F8:H8)", result: 40 });
assert.equal(
  writeSegmentMetricPreservingFormula(sheet.getCell("J8"), 100, {
    sheet,
    rowNumber: 8,
    columns,
    periods,
    periodIndex: 4,
    segment,
    metric: "values"
  }),
  true
);
assert.deepEqual(sheet.getCell("J8").value, { formula: "SUM(F8:I8)", result: 100 });

sheet.getCell("C35").value = "Canada Operating Income";
for (const address of ["F35", "G35", "H35", "I35"]) sheet.getCell(address).value = 0;
sheet.getCell("J35").value = { formula: "SUM(F35:I35)", result: 0 };
assert.equal(
  writeSegmentMetricPreservingFormula(sheet.getCell("J35"), 50, {
    sheet,
    rowNumber: 35,
    columns,
    periods,
    periodIndex: 4,
    segment: { ...segment, operatingIncome: new Map([["FY23", 50_000_000]]) },
    metric: "operatingIncome"
  }),
  true
);
assert.equal(sheet.getCell("J35").value, 50, "a disclosed annual value must replace a non-reconciling quarterly-sum formula");

sheet.getCell("K8").value = { formula: "'Revenue Build'!K8", result: 1 };
assert.equal(writeSegmentMetricPreservingFormula(sheet.getCell("K8"), 25), true);
assert.deepEqual(sheet.getCell("K8").value, { formula: "'Revenue Build'!K8", result: 25 });

const model = workbook.addWorksheet("Model");
model.getCell("C55").value = "Income Statement";
model.getCell("C59").value = "EBIT";
model.getCell("C60").value = "Depreciation & Amortization";
model.getCell("C61").value = "EBITDA";
model.getCell("F60").value = { formula: "-F34", result: 0 };
const daAuditRows = [];
const daAddback = writeHistoricalEbitdaDaAddback(
  model,
  ["1Q23"],
  [6],
  {
    duration: new Map([
      [
        "1Q23",
        new Map([
          [
            "DepreciationDepletionAndAmortization",
            {
              concept: "DepreciationDepletionAndAmortization",
              label: "Depreciation, depletion and amortization",
              value: 25_000_000,
              periodType: "quarterly",
              form: "10-Q",
              accn: "0000000000-23-000001"
            }
          ]
        ])
      ]
    ])
  },
  daAuditRows
);
assert.equal(daAddback.filledCells, 1);
assert.equal(model.getCell("F60").value, 25);
assert.equal(daAuditRows[0].validationStatus, "OK!");
sheet.getCell("C51").value = "Total D&A";
sheet.getCell("C59").value = "D&A Check";
for (let row = 52; row <= 57; row += 1) sheet.getCell(row, 6).value = 0;
const auditRows = [];
const reconciliation = reconcileSegmentMetricRowsToModelRow(sheet, ["1Q23"], [6], [52, 53, 54, 55, 56, 57], "D&A", auditRows, [
  "Depreciation & Amortization"
]);
assert.equal(reconciliation.filledCells, 1);
assert.equal(
  [52, 53, 54, 55, 56, 57].reduce((total, row) => total + (Number(sheet.getCell(row, 6).value) || 0), 0),
  25
);
assert.match(sheet.getCell("C57").text, /Other \/ Reconciliation D&A/i);
assert.equal(auditRows[0].validationStatus, "OK!");

sheet.getCell("C70").value = "Other International Operating Income";
sheet.getCell("F70").value = 0;
sheet.getCell("C71").value = "451";
sheet.getCell("F71").value = 0;
assert.equal(findSegmentResidualRow(sheet, [70, 71], 6, "Operating Income"), 71);
assert.match(sheet.getCell("C71").text, /Other \/ Reconciliation Operating Income/i);

console.log("Segment formula-write and annual bridge checks passed.");
