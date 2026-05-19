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
  { sheet: "Model", name: "Balance Sheet", start: 118, end: 159 }
];

const preservedChecks = [
  { sheet: "Model", name: "Cash Flow Statement", start: 73, end: 104 },
  { sheet: "Model", name: "PP&E / Depreciation Schedule", start: 193, end: 213 },
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

function isAllowedDividendBridgeUpdate(sheetName, cell, expected, got) {
  if (sheetName !== "Model" || cell.address !== "I257") return false;
  const normalizedExpected = String(expected ?? "").replace(/^=/, "");
  const normalizedGot = String(got ?? "").replace(/^=/, "");
  return normalizedExpected === "-169.4-SUM(F257:H257)" && normalizedGot === "-278.6-SUM(F257:H257)";
}

function normalize(input) {
  return String(input ?? "").toLowerCase().replace(/[^a-z0-9]/g, "");
}

function rowLabel(sheet, row) {
  return String(cellValue(sheet.getCell(row, 3)) ?? "").trim();
}

function isProtectedBeginningBalanceRow(sheet, row) {
  if (normalize(rowLabel(sheet, row)) !== normalize("Beginning Balance")) return false;
  for (let contextRow = row - 1; contextRow >= Math.max(1, row - 12); contextRow -= 1) {
    const label = normalize(rowLabel(sheet, contextRow));
    if (
      label === normalize("Retained Earnings") ||
      label === normalize("AOCI Assumptions") ||
      label === normalize("Revolver Balance") ||
      label === normalize("Total Debt Balance")
    ) {
      return true;
    }
  }
  return false;
}

function isProtectedPpeDepreciationRow(sheet, row) {
  if (normalize(rowLabel(sheet, row)) !== normalize("Depreciation Expense")) return false;
  for (let contextRow = row - 1; contextRow >= Math.max(1, row - 8); contextRow -= 1) {
    const label = normalize(rowLabel(sheet, contextRow));
    if (label === normalize("PP&E / Depreciation Schedule") || label === normalize("PPE / Depreciation Schedule")) return true;
  }
  return false;
}

function isProtectedScheduleClearRow(sheet, row) {
  return isProtectedBeginningBalanceRow(sheet, row) || isProtectedPpeDepreciationRow(sheet, row);
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
        if (expected !== got) {
          if (isAllowedDividendBridgeUpdate(sheetName, expectedSheet.getCell(row, col), expected, got)) continue;
          errors.push(`${sheetName}!${expectedSheet.getCell(row, col).address}: formula changed from "${expected}" to "${got ?? "[hardcoded/blank]"}".`);
        }
      }
    }
  }
}

function compareFilledRanges(golden, actual, errors) {
  for (const check of checks) {
    const expectedSheet = golden.getWorksheet(check.sheet);
    const actualSheet = actual.getWorksheet(check.sheet);
    if (!expectedSheet || !actualSheet) continue;
    for (let row = check.start; row <= check.end; row += 1) {
      for (let col = 6; col <= 20; col += 1) {
        const expected = cellValue(expectedSheet.getCell(row, col));
        const got = cellValue(actualSheet.getCell(row, col));
        if (!valuesMatch(expected, got)) {
          if (isAllowedDividendBridgeUpdate(check.sheet, expectedSheet.getCell(row, col), expected, got)) continue;
          if (check.sheet === "Model" && isProtectedScheduleClearRow(actualSheet, row) && isBlank(actualSheet.getCell(row, col)) && !cellFormula(expectedSheet.getCell(row, col))) continue;
          errors.push(`${check.name} ${check.sheet}!${expectedSheet.getCell(row, col).address}: expected "${expected ?? ""}", got "${got ?? ""}".`);
        }
      }
    }
  }
}

function comparePreservedRanges(input, actual, errors) {
  for (const check of preservedChecks) {
    const inputSheet = input.getWorksheet(check.sheet);
    const actualSheet = actual.getWorksheet(check.sheet);
    if (!inputSheet || !actualSheet) continue;
    for (let row = check.start; row <= check.end; row += 1) {
      for (let col = 1; col <= Math.max(inputSheet.columnCount, actualSheet.columnCount); col += 1) {
        const expected = cellValue(inputSheet.getCell(row, col));
        const got = cellValue(actualSheet.getCell(row, col));
        if (!valuesMatch(expected, got)) {
          errors.push(`${check.name} ${check.sheet}!${inputSheet.getCell(row, col).address}: expected preserved "${expected ?? ""}", got "${got ?? ""}".`);
        }
      }
    }
  }
}

function isBlank(cell) {
  const got = cellValue(cell);
  return got === null || got === undefined || got === "";
}

async function main() {
  await fillWorkbook();
  const golden = await readWorkbook(goldenWorkbook);
  const input = await readWorkbook(inputWorkbook);
  const actual = await readWorkbook(outputWorkbook);
  const errors = [];
  compareLabels(golden, actual, errors);
  compareFormulas(golden, actual, errors);
  compareFilledRanges(golden, actual, errors);
  comparePreservedRanges(input, actual, errors);

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
