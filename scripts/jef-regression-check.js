const fs = require("node:fs/promises");
const path = require("node:path");
const ExcelJS = require("exceljs");

const repoRoot = path.resolve(__dirname, "..");
const goldenWorkbook =
  process.env.JEF_GOLDEN_WORKBOOK || "/Users/maxschmieder/Desktop/Jefferies Financial Group Inc. (JEF)_Valuation Workbook (10-Mar-2026) (1).xlsx";
const inputWorkbook = process.env.JEF_INPUT_WORKBOOK || goldenWorkbook;
const outputWorkbook = process.env.JEF_OUTPUT_WORKBOOK || path.join(repoRoot, "tmp", "jef-regression-output.xlsx");
const apiUrl = process.env.FILL_API_URL || "http://localhost:3000/api/fill-model";
const ticker = process.env.JEF_TICKER || "JEF";

const checks = [
  { sheet: "Model", name: "Income Statement", start: 28, end: 57 },
  { sheet: "Model", name: "Balance Sheet", start: 118, end: 159 },
  { sheet: "Model", name: "Shareholder Equity / Shares", start: 237, end: 287 },
  { sheet: "Model", name: "Debt and Interest Schedule", start: 288, end: 360 },
  { sheet: "Segment Analysis", name: "Segment Revenue", start: 7, end: 17 }
];

function cellValue(cell) {
  const value = cell.value;
  if (value && typeof value === "object") {
    if (typeof value.formula === "string") return `=${value.formula}`;
    if (typeof value.sharedFormula === "string") return `=shared:${value.sharedFormula}`;
    if (typeof value.result === "number") return value.result;
    if (Array.isArray(value.richText)) return value.richText.map((part) => part.text).join("");
  }
  return value;
}

function cellFormula(cell) {
  const value = cell.value;
  if (value && typeof value === "object") {
    if (typeof value.formula === "string") return value.formula;
    if (typeof value.sharedFormula === "string") return `shared:${value.sharedFormula}`;
  }
  return null;
}

function valuesMatch(a, b) {
  if (typeof a === "number" && typeof b === "number") return Math.abs(a - b) < 0.05;
  return String(a ?? "") === String(b ?? "");
}

async function fillWorkbook() {
  await fs.mkdir(path.dirname(outputWorkbook), { recursive: true });
  const bytes = await fs.readFile(inputWorkbook);
  const formData = new FormData();
  formData.append("ticker", ticker);
  formData.append("file", new Blob([bytes]), path.basename(inputWorkbook));

  const response = await fetch(apiUrl, { method: "POST", body: formData });
  const body = Buffer.from(await response.arrayBuffer());
  if (!response.ok) {
    throw new Error(`Fill API failed (${response.status}): ${body.toString("utf8")}`);
  }
  await fs.writeFile(outputWorkbook, body);
}

async function readWorkbook(file) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(file);
  return workbook;
}

function compareLabels(golden, actual, errors) {
  for (const sheetName of ["Model", "Segment Analysis"]) {
    const expectedSheet = golden.getWorksheet(sheetName);
    const actualSheet = actual.getWorksheet(sheetName);
    if (!expectedSheet || !actualSheet) continue;
    const maxRows = Math.max(expectedSheet.rowCount, actualSheet.rowCount);
    for (let row = 1; row <= maxRows; row += 1) {
      for (let col = 1; col <= 5; col += 1) {
        const expected = cellValue(expectedSheet.getCell(row, col));
        const got = cellValue(actualSheet.getCell(row, col));
        if (!valuesMatch(expected, got)) {
          errors.push(`${sheetName}!${expectedSheet.getCell(row, col).address}: label/layout changed from "${expected ?? ""}" to "${got ?? ""}".`);
        }
      }
    }
  }
}

function compareFormulas(golden, actual, errors) {
  for (const sheetName of ["Model", "Segment Analysis"]) {
    const expectedSheet = golden.getWorksheet(sheetName);
    const actualSheet = actual.getWorksheet(sheetName);
    if (!expectedSheet || !actualSheet) continue;
    for (let row = 1; row <= expectedSheet.rowCount; row += 1) {
      for (let col = 1; col <= expectedSheet.columnCount; col += 1) {
        const expected = cellFormula(expectedSheet.getCell(row, col));
        if (!expected) continue;
        const got = cellFormula(actualSheet.getCell(row, col));
        if (sheetName === "Model" && isPpeScheduleUnsupportedInput(row, col) && !got && isBlank(actualSheet.getCell(row, col))) continue;
        if (expected !== got) {
          errors.push(`${sheetName}!${expectedSheet.getCell(row, col).address}: formula changed from "${expected}" to "${got ?? "[hardcoded/blank]"}".`);
        }
      }
    }
  }
}

function compareRanges(golden, actual, errors) {
  for (const check of checks) {
    const expectedSheet = golden.getWorksheet(check.sheet);
    const actualSheet = actual.getWorksheet(check.sheet);
    if (!expectedSheet || !actualSheet) continue;
    for (let row = check.start; row <= check.end; row += 1) {
      for (let col = 6; col <= 20; col += 1) {
        const expected = cellValue(expectedSheet.getCell(row, col));
        const got = cellValue(actualSheet.getCell(row, col));
        if (!valuesMatch(expected, got)) {
          errors.push(`${check.name} ${check.sheet}!${expectedSheet.getCell(row, col).address}: expected "${expected ?? ""}", got "${got ?? ""}".`);
        }
      }
    }
  }
}

function compareCashFlowBlank(actual, errors) {
  const sheet = actual.getWorksheet("Model");
  if (!sheet) return;
  const quarterCols = [6, 7, 8, 9, 11, 12, 13, 14, 16, 17, 18, 19];
  for (let row = 73; row <= 104; row += 1) {
    for (const col of quarterCols) {
      if (cellFormula(sheet.getCell(row, col))) continue;
      const got = cellValue(sheet.getCell(row, col));
      if (got !== null && got !== undefined && got !== "") {
        errors.push(`Cash Flow Statement Model!${sheet.getCell(row, col).address}: expected blank historical input, got "${got}".`);
      }
    }
  }
}

function isBlank(cell) {
  const got = cellValue(cell);
  return got === null || got === undefined || got === "";
}

function isPpeScheduleUnsupportedInput(row, col) {
  const quarterCols = new Set([6, 7, 8, 9, 11, 12, 13, 14, 16, 17, 18, 19]);
  return row >= 196 && row <= 198 && quarterCols.has(col);
}

function comparePpeScheduleBlank(actual, errors) {
  const sheet = actual.getWorksheet("Model");
  if (!sheet) return;
  for (let row = 196; row <= 198; row += 1) {
    for (const col of [6, 7, 8, 9, 11, 12, 13, 14, 16, 17, 18, 19]) {
      const cell = sheet.getCell(row, col);
      const got = cellValue(cell);
      if (!isBlank(cell)) {
        errors.push(`PP&E / Depreciation Schedule Model!${cell.address}: expected blank unsupported historical input, got "${got}".`);
      }
    }
  }
}

async function main() {
  await fillWorkbook();
  const golden = await readWorkbook(goldenWorkbook);
  const actual = await readWorkbook(outputWorkbook);
  const errors = [];
  compareLabels(golden, actual, errors);
  compareFormulas(golden, actual, errors);
  compareRanges(golden, actual, errors);
  compareCashFlowBlank(actual, errors);
  comparePpeScheduleBlank(actual, errors);

  if (errors.length) {
    console.error(errors.slice(0, 40).join("\n"));
    throw new Error(`JEF regression failed with ${errors.length} mismatch(es).`);
  }

  console.log(`JEF regression passed: ${outputWorkbook}`);
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
