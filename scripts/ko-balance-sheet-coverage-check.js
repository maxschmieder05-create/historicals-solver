const path = require("node:path");
const ExcelJS = require("exceljs");
const { postWorkbook } = require("./fill-workbook-api");

const repoRoot = path.resolve(__dirname, "..");
const apiUrl = process.env.FILL_API_URL || "http://localhost:3000/api/fill-model";
const inputWorkbook = process.env.KO_INPUT_WORKBOOK || "/Users/maxschmieder/Downloads/KO_historicals_filled.xlsx";
const outputWorkbook = process.env.KO_BALANCE_SHEET_OUTPUT_WORKBOOK || path.join(repoRoot, "tmp", "ko-balance-sheet-coverage-output.xlsx");

const expectedRows = [
  { labels: ["Cash & Cash Equivalents", "Cash and Cash Equivalents"], value: 13820 },
  { labels: ["Accounts Receivable", "Accounts Receivable, Net", "Trade Receivables"], value: 3675 },
  { labels: ["Inventory", "Inventories"], value: 4730 },
  { labels: ["Prepaid & Other Current Assets", "Prepaid and Other Current Assets", "Other Current Assets"], value: 8165 },
  { labels: ["Total Current Assets"], value: 30390 },
  { labels: ["PP&E, Net", "Property, Plant and Equipment, Net", "Property and Equipment, Net"], value: 9522 },
  { labels: ["Intangible Assets, Net", "Intangibles, Net"], value: 12463 },
  { labels: ["Goodwill"], value: 15411 },
  { labels: ["Other Non-Current Assets", "Other Long-Term Assets", "Other LT Assets"], value: 36431 },
  { labels: ["Total Assets"], value: 104217 },
  { labels: ["Accounts Payable", "Accounts Payable and Accrued Liabilities", "Accounts Payable & Accrued Liabilities"], value: 14409 },
  { labels: ["Accrued Liabilities", "Accrued Expenses", "Accrued Expenses and Other"], value: 717 },
  { labels: ["Other Current Liabilities", "Other Current Liabs"], value: 2427 },
  { labels: ["Revolver", "Short-Term Borrowings", "Short Term Borrowings", "Current Borrowings"], value: 332 },
  { labels: ["LT Debt (Incl. Current Portion)", "Long-Term Debt", "Long Term Debt", "Borrowings"], value: 43558 },
  { labels: ["Deferred Income Taxes", "Deferred Tax Liabilities", "Deferred Taxes"], value: 2615 },
  { labels: ["Other Non-Current Liabilities", "Other Long-Term Liabilities", "Other LT Liabilities"], value: 4425 },
  { labels: ["Total Liabilities"], value: 68483 },
  { labels: ["Common Stock & APIC", "Common Stock and APIC", "Common Stock and Additional Paid-In Capital"], value: 22394 },
  { labels: ["Retained Earnings", "Accumulated Deficit"], value: 82026 },
  { labels: ["Accumulated Other Comprehensive Income (AOCI)", "Accumulated Other Comprehensive Income", "AOCI"], value: -14040 },
  { labels: ["Treasury Stock", "Treasury & Preferred Stock"], value: -56747 },
  { labels: ["Noncontrolling Interests", "Non-Controlling Interests"], value: 2101 },
  { labels: ["Total Equity"], value: 35734 },
  { labels: ["Total Liabilities & Shareholder's Equity", "Total Liabilities and Shareholder's Equity"], value: 104217 },
  { labels: ["Balance Sheet Check"], value: 0, tolerance: 0.1 }
];

const ledgerExpectations = [
  { source: "Cash and cash equivalents", row: "Cash & Cash Equivalents" },
  { source: "Short-term investments", row: "Cash & Cash Equivalents" },
  { source: "Marketable securities", row: "Cash & Cash Equivalents" },
  { source: "Trade accounts receivable", row: "Accounts Receivable" },
  { source: "Inventories", row: "Inventory" },
  { source: "Prepaid expenses and other current assets", row: "Prepaid & Other Current Assets" },
  { source: "Assets held for sale", row: "Prepaid & Other Current Assets" },
  { source: "Equity method investments", row: "Other Non-Current Assets" },
  { source: "Deferred income tax assets", row: "Other Non-Current Assets" },
  { source: "Property, plant and equipment", row: "PP&E, Net" },
  { source: "Trademarks with indefinite lives", row: "Intangible Assets, Net" },
  { source: "Goodwill", row: "Goodwill" },
  { source: "Other noncurrent assets", row: "Other Non-Current Assets" },
  { source: "Accounts payable and accrued expenses", row: "Accounts Payable" },
  { source: "Loans and notes payable", row: "Revolver" },
  { source: "Current maturities of long-term debt", row: "LT Debt (Incl. Current Portion)" },
  { source: "Accrued income taxes", row: "Accrued Liabilities" },
  { source: "Liabilities held for sale", row: "Other Current Liabilities" },
  { source: "Long-term debt", row: "LT Debt (Incl. Current Portion)" },
  { source: "Other noncurrent liabilities", row: "Other Non-Current Liabilities" },
  { source: "Deferred income tax liabilities", row: "Deferred Income Taxes" },
  { source: "Common stock", row: "Common Stock & APIC" },
  { source: "Capital surplus", row: "Common Stock & APIC" },
  { source: "Reinvested earnings", row: "Retained Earnings" },
  { source: "Accumulated other comprehensive income", row: "AOCI" },
  { source: "Treasury stock", row: "Treasury Stock" },
  { source: "Equity attributable to noncontrolling interests", row: "Noncontrolling Interests" }
];

async function main() {
  await postWorkbook({ apiUrl, ticker: "KO", inputWorkbook, outputWorkbook });
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(outputWorkbook);
  const model = workbook.getWorksheet("Model");
  const ledger = workbook.getWorksheet("Balance Sheet Assignment Ledger");
  if (!model) throw new Error("Output workbook does not contain a Model sheet.");
  if (!ledger) throw new Error("Output workbook does not contain a Balance Sheet Assignment Ledger sheet.");

  const col = findPeriodColumn(model, "1Q26");
  if (!col) throw new Error("Could not find 1Q26 column in Model sheet.");

  const errors = [];
  for (const check of expectedRows) {
    const row = findRowInBalanceSheet(model, check.labels);
    if (!row) {
      errors.push(`Could not find model row ${check.labels[0]}.`);
      continue;
    }
    const actual = cellNumber(model.getCell(row, col));
    if (!valuesMatch(actual, check.value, check.tolerance ?? 0.5)) {
      errors.push(`${check.labels[0]} expected ${check.value}, got ${actual ?? "[blank]"} at ${model.getCell(row, col).address}.`);
    }
  }

  const currentLiabilitiesRow = findRowInBalanceSheet(model, ["Total Current Liabilities"]);
  const currentLiabilitiesExDebtRow = findRowInBalanceSheet(model, [
    "Total Current Liabilities (Excl. Debt)",
    "Total Current Liabilities Excl. Debt"
  ]);
  if (currentLiabilitiesRow) {
    const actual = cellNumber(model.getCell(currentLiabilitiesRow, col));
    if (!valuesMatch(actual, 22378)) errors.push(`Total Current Liabilities expected 22378, got ${actual ?? "[blank]"}.`);
  } else if (currentLiabilitiesExDebtRow) {
    const actual = cellNumber(model.getCell(currentLiabilitiesExDebtRow, col));
    if (!valuesMatch(actual, 17553)) errors.push(`Total Current Liabilities (Excl. Debt) expected 17553, got ${actual ?? "[blank]"}.`);
  } else {
    errors.push("Could not find current liabilities subtotal row.");
  }

  const ledgerRows = ledgerRowsForPeriod(ledger, "1Q26");
  if (!ledgerRows.length) errors.push("Balance Sheet Assignment Ledger has no 1Q26 rows.");
  const assignedLedgerRows = ledgerRows.filter((row) => ["mapped_to_model_row", "grouped_into_model_row"].includes(row.assignmentStatus));
  const assetSum = assignedLedgerRows.filter((row) => row.side === "assets").reduce((total, row) => total + row.amount, 0);
  const liabilitiesEquitySum = assignedLedgerRows.filter((row) => row.side === "liabilities_and_equity").reduce((total, row) => total + row.amount, 0);
  if (!valuesMatch(assetSum, 104217000000, 500000)) errors.push(`Assignment ledger asset rows expected 104217000000, got ${assetSum}.`);
  if (!valuesMatch(liabilitiesEquitySum, 104217000000, 500000)) {
    errors.push(`Assignment ledger liabilities/equity rows expected 104217000000, got ${liabilitiesEquitySum}.`);
  }

  for (const expected of ledgerExpectations) {
    const candidates = ledgerRows.filter((candidate) => normalize(candidate.sourceLineItemLabel).includes(normalize(expected.source)));
    const row = candidates.find((candidate) => modelRowsMatch(candidate.assignedModelRow, expected.row)) ?? candidates[0];
    if (!row) {
      errors.push(`Assignment ledger missing source line "${expected.source}".`);
      continue;
    }
    if (!modelRowsMatch(row.assignedModelRow, expected.row)) {
      errors.push(`Assignment ledger maps "${expected.source}" to "${row.assignedModelRow}", expected "${expected.row}".`);
    }
    if (!["mapped_to_model_row", "grouped_into_model_row"].includes(row.assignmentStatus)) {
      errors.push(`Assignment ledger status for "${expected.source}" should be mapped/grouped, got "${row.assignmentStatus}".`);
    }
  }

  if (errors.length) {
    console.error(errors.join("\n"));
    throw new Error(`KO balance sheet coverage regression failed with ${errors.length} issue(s).`);
  }

  console.log(`KO balance sheet coverage regression passed: ${outputWorkbook}`);
}

function ledgerRowsForPeriod(sheet, period) {
  const header = headerMap(sheet);
  const rows = [];
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    if (String(row.getCell(header.get("fiscal period")).value || "").trim() !== period) return;
    rows.push({
      sourceLineItemLabel: String(row.getCell(header.get("source line item label")).value || ""),
      assignedModelRow: String(row.getCell(header.get("assigned model row")).value || ""),
      assignmentStatus: String(row.getCell(header.get("assignment status")).value || ""),
      amount: Number(row.getCell(header.get("amount")).value || 0),
      side: String(row.getCell(header.get("side")).value || "")
    });
  });
  return rows;
}

function headerMap(sheet) {
  const map = new Map();
  sheet.getRow(1).eachCell((cell, col) => {
    map.set(String(cell.value || "").trim().toLowerCase(), col);
  });
  return map;
}

function findPeriodColumn(sheet, period) {
  for (let col = 1; col <= sheet.columnCount; col += 1) {
    const text = String(sheet.getCell(25, col).text || sheet.getCell(25, col).value || "").replace(/[’']/g, "").trim().toUpperCase();
    if (text === period) return col;
  }
  return null;
}

function findRowInBalanceSheet(sheet, labels) {
  const wanted = new Set(labels.map(normalize));
  const sectionStart = findBalanceSheetSectionStart(sheet);
  if (!sectionStart) return null;
  for (let row = sectionStart + 1; row <= sheet.rowCount; row += 1) {
    const label = rowLabel(sheet, row);
    if (/income statement analysis|cash flow statement|working capital|schedule|drivers/i.test(label)) break;
    if (wanted.has(normalize(label))) return row;
  }
  return null;
}

function findBalanceSheetSectionStart(sheet) {
  let fallback = null;
  for (let row = 1; row <= sheet.rowCount; row += 1) {
    const label = rowLabel(sheet, row);
    if (normalize(label) !== normalize("Balance Sheet")) continue;
    fallback = row;
    for (let probe = row + 1; probe <= Math.min(sheet.rowCount, row + 20); probe += 1) {
      if (modelRowsMatch(rowLabel(sheet, probe), "Cash & Cash Equivalents")) return row;
    }
  }
  return fallback;
}

function rowLabel(sheet, row) {
  for (let col = 1; col <= Math.min(sheet.columnCount, 8); col += 1) {
    const text = String(sheet.getCell(row, col).text || sheet.getCell(row, col).value || "").trim();
    if (/^x$/i.test(text)) continue;
    if (text) return text;
  }
  return "";
}

function cellNumber(cell) {
  const value = cell.value;
  if (typeof value === "number") return value;
  if (value && typeof value === "object") {
    if (typeof value.result === "number") return value.result;
    if (typeof value.result === "string") return Number(value.result);
  }
  return null;
}

function valuesMatch(actual, expected, tolerance = 0.5) {
  return typeof actual === "number" && Number.isFinite(actual) && Math.abs(actual - expected) <= tolerance;
}

function normalize(value) {
  return String(value).toLowerCase().replace(/[^a-z0-9]/g, "");
}

function modelRowsMatch(a, b) {
  const left = normalize(a);
  const right = normalize(b);
  if (left === right) return true;
  const aliases = [
    ["cashcashequivalents", "cashandcashequivalents", "cash"],
    ["ppenet", "propertyplantandequipmentnet", "propertyandequipmentnet"],
    ["intangibleassetsnet", "intangiblesnet"],
    ["commonstockapic", "commonstockandapic", "commonstockandadditionalpaidincapital"],
    ["accumulatedothercomprehensiveincomeaoci", "accumulatedothercomprehensiveincome", "aoci"],
    ["ltdebtinclcurrentportion", "longtermdebt", "borrowings"],
    ["revolver", "shorttermborrowings", "currentborrowings"],
    ["othernoncurrentassets", "otherlongtermassets", "otherltassets"],
    ["othernoncurrentliabilities", "otherlongtermliabilities", "otherltliabilities"]
  ];
  return aliases.some((group) => group.includes(left) && group.includes(right));
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
