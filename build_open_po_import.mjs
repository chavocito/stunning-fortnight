import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const templatePath = "C:/Users/FrancisBota/Downloads/Purchase Orders - V1.5.xlsx";
const workDir = "C:/Users/FrancisBota/Documents/Codex/2026-09-10/i/work/open_po_import";
const sourcePath = path.join(workDir, "input", "Open Purchase Orders(Sheet1) (1).csv");
const ownerCrossReferencePath = path.join(workDir, "outputs", "open-po-import-fixed", "open_purchase_orders_acumatica_import.xlsx");
const outputDir = path.join(workDir, "outputs", "actual-open-pos-2026-10-04");

const SOURCE_COLUMNS = [
  "Order Co", "Business Unit", "Order Number", "Line No", "Order Type", "Supplier Number",
  "2nd Item Number", "Item Description", "Unit Of Measure", "Quantity To Receive", "Amount to Receive",
  "Order Date", "Request Date", "Promised Delivery Date",
];

const TARGET_HEADERS = [
  "Legacy", "Type", "Order Nbr.", "Legacy Order Nbr.", "Vendor", "Location", "Owner", "Vendor Ref.", "Date",
  "Promised On", "Description", "Shipping Destination Type", "Ship To", "Shipping Location", "Ship To Warehouse",
  "Line Nbr.", "Branch (Transaction)", "Inventory ID", "Warehouse", "Line Description", "UOM", "Order Qty.",
  "Qty. On Receipts", "Billed Qty.", "Unit Cost", "Ext. Cost", "Discount Percent", "Discount Amount",
  "Discount Code", "Amount", "Alternate ID", "Min. Receipt (%)", "Max. Receipt (%)", "Receipt Action",
  "Tax Category", "Accrual Account", "Accrual Sub.", "Project", "Project Task", "Requested", "Promised",
  "Completed", "Cancelled", "Closed",
];

const COLUMN_LIMITS = new Map([
  ["Order Nbr.", 15], ["Legacy Order Nbr.", 128], ["Vendor", 30], ["Location", 30], ["Owner", 30],
  ["Vendor Ref.", 40], ["Description", 60], ["Ship To", 30], ["Shipping Location", 30],
  ["Ship To Warehouse", 30], ["Branch (Transaction)", 30], ["Inventory ID", 30], ["Warehouse", 30],
  ["Line Description", 256], ["UOM", 6], ["Discount Code", 10], ["Alternate ID", 50],
  ["Tax Category", 15], ["Accrual Account", 10], ["Accrual Sub.", 30], ["Project", 30],
  ["Project Task", 30],
]);

const REQUIRED_COLUMNS = [
  "Legacy", "Type", "Order Nbr.", "Legacy Order Nbr.", "Vendor", "Location", "Owner", "Date",
  "Shipping Location", "Line Nbr.", "Branch (Transaction)", "Inventory ID", "Warehouse", "UOM",
  "Order Qty.", "Ext. Cost",
];

const mappingNotes = [
  ["Legacy", "Default", "TRUE for every imported legacy PO row."],
  ["Type", "Default", "Normal for every row. JDE source types are summarized for validation."],
  ["Order Nbr.", "Order Number", "Prefixes the JDE order number with PO, for example PO24332."],
  ["Legacy Order Nbr.", "Order Number", "Stores the original JDE order number without a prefix."],
  ["Vendor", "Supplier Number", "Must match an existing Acumatica vendor or a vendor crosswalk."],
  ["Location", "Default", "MAIN. Must exist for the Acumatica vendor/location setup."],
  ["Owner", "Previous import workbook: Order Nbr. + Line Nbr.", "Matches the source Order Number and Line No to the previous workbook before applying the PO prefix; unmatched rows remain blank."],
  ["Vendor Ref.", "Blank", "No reliable source values present in the extract."],
  ["Date", "Order Date", "Converted from the source date text."],
  ["Promised On", "Blank", "Promised Delivery Date is mapped to the line-level Promised field."],
  ["Description", "Default", "JDE {type} PO {number} Branch {branch}, truncated to 60 characters."],
  ["Shipping Destination Type", "Default", "Branch."],
  ["Ship To", "Business Unit", "Branch code, preserving the approved branch-based mapping."],
  ["Shipping Location", "Default", "MAIN."],
  ["Ship To Warehouse", "Business Unit", "Branch code used as warehouse default for migration dry run."],
  ["Line Nbr.", "Line No", "JDE PO line number."],
  ["Branch (Transaction)", "Business Unit", "Receiving branch code."],
  ["Inventory ID", "2nd Item Number / outside-process mapping", "Branch 150/1550 item 281680 maps by description to the configured non-stock outside-process item; unmatched descriptions use SUBCON-USD or SUBCON-CAD."],
  ["Warehouse", "Business Unit", "Branch code used as receiving warehouse default for migration dry run."],
  ["Line Description", "Item Description", "Description is trimmed and truncated to 256 characters."],
  ["UOM", "Unit Of Measure", "JDE transaction UOM."],
  ["Order Qty.", "Quantity To Receive", "Open quantity carried forward; amount-only lines use quantity 1."],
  ["Qty. On Receipts", "Default", "0, because this load carries open commitments only."],
  ["Billed Qty.", "Default", "0, because this load carries open commitments only."],
  ["Unit Cost", "Amount to Receive / Quantity To Receive", "Calculated from the open amount and mapped open quantity."],
  ["Ext. Cost", "Amount to Receive", "Open amount carried forward."],
  ["Discount Percent", "Blank", "No reliable source value mapped."],
  ["Discount Amount", "Blank", "No reliable source value mapped."],
  ["Discount Code", "Blank", "No reliable source value mapped."],
  ["Amount", "Amount to Receive", "Open amount carried forward."],
  ["Alternate ID", "Blank", "No reliable source value mapped."],
  ["Min. Receipt (%)", "Blank", "No reliable source value mapped."],
  ["Max. Receipt (%)", "Blank", "No reliable source value mapped."],
  ["Receipt Action", "Blank", "No reliable source value mapped."],
  ["Tax Category", "Blank", "No reliable source value mapped."],
  ["Accrual Account", "Blank", "No reliable source value mapped."],
  ["Accrual Sub.", "Blank", "No reliable source value mapped."],
  ["Project", "Blank", "No reliable source value mapped."],
  ["Project Task", "Blank", "No reliable source value mapped."],
  ["Requested", "Request Date", "Converted from the source date text."],
  ["Promised", "Promised Delivery Date", "Converted from the source date text when present."],
  ["Completed", "Default", "FALSE."],
  ["Cancelled", "Default", "FALSE."],
  ["Closed", "Default", "FALSE."],
];

function excelSerialToDate(value) {
  if (value === null || value === undefined || value === "") return null;
  if (value instanceof Date) return value;
  if (typeof value === "number" && !Number.isNaN(value)) return new Date(Date.UTC(1899, 11, 30) + value * 86400000);
  const match = asText(value).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  return match ? new Date(Date.UTC(Number(match[3]), Number(match[1]) - 1, Number(match[2]))) : value;
}

function asText(value) {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

function cleanText(value) {
  const text = asText(value);
  if (text.length >= 2 && text.startsWith("\"") && text.endsWith("\"")) return text.slice(1, -1).trim();
  return text;
}

function truncate(value, maxLength) {
  const text = cleanText(value);
  return text.length > maxLength ? text.slice(0, maxLength) : text;
}

function toNumber(value) {
  if (value === null || value === undefined || value === "") return 0;
  const num = Number(value);
  return Number.isFinite(num) ? num : 0;
}

function orderQuantity(unitsOpen, amountOpen) {
  const quantity = toNumber(unitsOpen);
  return quantity === 0 && toNumber(amountOpen) !== 0 ? 1 : quantity;
}

console.assert(orderQuantity(0, 944.9) === 1, "Amount-only PO lines must receive quantity 1.");

function columnLetter(indexOneBased) {
  let n = indexOneBased;
  let s = "";
  while (n > 0) {
    const mod = (n - 1) % 26;
    s = String.fromCharCode(65 + mod) + s;
    n = Math.floor((n - mod) / 26);
  }
  return s;
}

function rangeAddress(cols, rows) {
  return `A1:${columnLetter(cols)}${rows}`;
}

function getHeaderIndex(headers) {
  return new Map(headers.map((h, i) => [h, i]));
}

function rowObj(headers, row) {
  const out = {};
  headers.forEach((h, i) => { out[h] = row[i]; });
  return out;
}

function sourceKey(r) {
  return [r.company, r.branch, r.po_order_type, r.po_order_no, r.po_order_suffix].map(asText).join("|");
}

function orderLineKey(r) {
  return `${asText(r.po_order_no)}|${toNumber(r.po_line)}`;
}

console.assert(orderLineKey({ po_order_no: "123", po_line: "2.000" }) === "123|2");

function legacyOrderNumber(r) {
  return asText(r.po_order_no);
}

function orderNumber(r) {
  return `PO${asText(r.po_order_no)}`;
}

console.assert(orderNumber({ po_order_no: 24332 }) === "PO24332");
console.assert(legacyOrderNumber({ po_order_no: 24332 }) === "24332");

function lineDescription(r) {
  return truncate([cleanText(r.item_desc_1), cleanText(r.item_desc_2)].filter(Boolean).join(" "), 256);
}

function normalizeCsvRow(r) {
  const branch = cleanText(r["Business Unit"]);
  const company = cleanText(r["Order Co"]);
  return {
    company,
    branch,
    po_order_no: cleanText(r["Order Number"]),
    po_line: r["Line No"],
    po_order_type: cleanText(r["Order Type"]),
    line_type: "",
    po_order_suffix: "",
    header_vendor_no: cleanText(r["Supplier Number"]),
    header_ordered_by: "",
    buyer: "",
    order_date: r["Order Date"],
    request_date: r["Request Date"],
    promised_date: r["Promised Delivery Date"],
    item_no: cleanText(r["2nd Item Number"]),
    item_id: cleanText(r["2nd Item Number"]),
    item_desc_1: cleanText(r["Item Description"]),
    item_desc_2: "",
    uom: cleanText(r["Unit Of Measure"]),
    units_open: r["Quantity To Receive"],
    amt_open: r["Amount to Receive"],
    location: "",
    currency_code: company === "00015" || branch === "1550" ? "CAD" : "USD",
  };
}

function inventoryId(r) {
  const sourceItem = cleanText(r.item_no);
  if (!["150", "1550"].includes(asText(r.branch)) || sourceItem !== "281680") return sourceItem;

  const description = lineDescription(r).toUpperCase();
  if (/\bHOSING\b/.test(description)) return "HOSING";
  if (/\bINSPECT/.test(description)) return "INSPECTION";
  if (/\b(?:TEST|COMMISSION|COMMISION|START UP)/.test(description)) return "TESTCOMM";
  if (/\b(?:CRATING|HAUL|TRANSPORT|MOVE)\b/.test(description)) return "LOGISTICS";

  const currency = asText(r.currency_code).toUpperCase();
  if (!["USD", "CAD"].includes(currency)) throw new Error(`Unsupported outside-process currency ${currency || "(blank)"}.`);
  return `SUBCON-${currency}`;
}

console.assert(inventoryId({ branch: 1550, item_no: 281680, item_desc_1: "HOSING WORK", currency_code: "CAD" }) === "HOSING");
console.assert(inventoryId({ branch: 1550, item_no: 281680, item_desc_1: "INSPECT PUMP", currency_code: "CAD" }) === "INSPECTION");
console.assert(inventoryId({ branch: 1550, item_no: 281680, item_desc_1: "XTD12 COMMISSIONING", currency_code: "CAD" }) === "TESTCOMM");
console.assert(inventoryId({ branch: 1550, item_no: 281680, item_desc_1: "CRATING", currency_code: "CAD" }) === "LOGISTICS");
console.assert(inventoryId({ branch: 1550, item_no: 281680, item_desc_1: "EAVOR RIG DESIGN", currency_code: "CAD" }) === "SUBCON-CAD");

function sortRows(rows) {
  return rows.sort((a, b) =>
    toNumber(a.branch) - toNumber(b.branch) ||
    toNumber(a.po_order_no) - toNumber(b.po_order_no) ||
    toNumber(a.po_line) - toNumber(b.po_line)
  );
}

function groupBy(rows, keyFn) {
  const map = new Map();
  for (const row of rows) {
    const key = keyFn(row);
    if (!map.has(key)) map.set(key, []);
    map.get(key).push(row);
  }
  return map;
}

function addSheetTable(sheet, startCell, rows, options = {}) {
  if (!rows.length) return;
  const rowCount = rows.length;
  const colCount = rows[0].length;
  const startRow = Number(startCell.match(/\d+/)[0]) - 1;
  const startColLetters = startCell.match(/[A-Z]+/)[0];
  let startCol = 0;
  for (const ch of startColLetters) startCol = startCol * 26 + (ch.charCodeAt(0) - 64);
  startCol -= 1;
  const range = sheet.getRangeByIndexes(startRow, startCol, rowCount, colCount);
  range.values = rows;
  range.format.font = { name: "Arial", size: 10, color: "#1F2933" };
  range.format.verticalAlignment = "center";
  range.format.borders = { preset: "outside", style: "thin", color: "#B7C9D6" };
  const header = sheet.getRangeByIndexes(startRow, startCol, 1, colCount);
  header.format = {
    fill: options.headerFill ?? "#1F4E78",
    font: { name: "Arial", size: 10, bold: true, color: "#FFFFFF" },
  };
  header.format.horizontalAlignment = "center";
  range.format.autofitColumns();
}

function formatImportSheet(sheet, rowCount) {
  sheet.showGridLines = false;
  sheet.freezePanes.freezeRows(1);
  sheet.getRange(rangeAddress(TARGET_HEADERS.length, rowCount)).format.font = { name: "Arial", size: 10, color: "#1F2933" };
  sheet.getRange("A1:AR1").format = {
    fill: "#1F4E78",
    font: { name: "Arial", size: 10, bold: true, color: "#FFFFFF" },
  };
  sheet.getRange("A1:AR1").format.horizontalAlignment = "center";
  sheet.getRange(`I2:I${rowCount}`).setNumberFormat("mm/dd/yy");
  sheet.getRange(`J2:J${rowCount}`).setNumberFormat("mm/dd/yy");
  sheet.getRange(`V2:Z${rowCount}`).setNumberFormat("#,##0.000000");
  sheet.getRange(`AD2:AD${rowCount}`).setNumberFormat("#,##0.0000");
  sheet.getRange(`AN2:AO${rowCount}`).setNumberFormat("mm/dd/yy");
  sheet.getRange("A1:AR1").format.autofitColumns();
  for (const [col, width] of [
    ["A", 10], ["B", 12], ["C", 12], ["D", 32], ["E", 12], ["F", 12], ["G", 12],
    ["I", 12], ["K", 28], ["L", 24], ["M", 12], ["N", 16], ["O", 18], ["P", 10],
    ["Q", 20], ["R", 22], ["S", 14], ["T", 36], ["U", 10], ["V", 14], ["Y", 14],
    ["Z", 14], ["AD", 14], ["AN", 12], ["AO", 12],
  ]) {
    sheet.getRange(`${col}1:${col}${rowCount}`).format.columnWidth = width;
  }
}

function addException(exceptions, type, severity, r, detail) {
  exceptions.push([
    type, severity, r.company ?? "", r.branch ?? "", r.po_order_no ?? "", r.po_order_type ?? "",
    r.po_order_suffix ?? "", r.po_line ?? "", cleanText(r.item_no), r.header_vendor_no ?? r.vendor ?? "",
    detail,
  ]);
}

function buildQaSummary(sourceHeaders, sourceLines, outputRows, exceptions, duplicateKeys, requiredBlanks, ownerMatches) {
  const uniquePOs = new Set(sourceLines.map((r) => asText(r.po_order_no))).size;
  const branchCount = new Set(sourceLines.map((r) => asText(r.branch))).size;
  const sourceOpenUnits = sourceLines.reduce((sum, r) => sum + orderQuantity(r.units_open, r.amt_open), 0);
  const outputOpenUnits = outputRows.reduce((sum, r) => sum + toNumber(r[TARGET_HEADERS.indexOf("Order Qty.")]), 0);
  const sourceOpenAmount = sourceLines.reduce((sum, r) => sum + toNumber(r.amt_open), 0);
  const outputAmount = outputRows.reduce((sum, r) => sum + toNumber(r[TARGET_HEADERS.indexOf("Amount")]), 0);
  return [
    ["Metric", "Value", "Status / Notes"],
    ["Source PO headers", sourceHeaders.length, "Derived from unique PO composite keys in the actual CSV."],
    ["Source PO lines", sourceLines.length, "Actual open-PO CSV rows."],
    ["Output import rows", outputRows.length, "One row per source PO line."],
    ["Unique PO numbers", uniquePOs, "Source number is stored in Legacy Order Nbr.; Order Nbr. uses the PO prefix."],
    ["Branches represented", branchCount, [...new Set(sourceLines.map((r) => asText(r.branch)))].sort().join(", ")],
    ["Duplicate Order Nbr. + Line Nbr.", duplicateKeys.length, duplicateKeys.length === 0 ? "OK" : "Review before import."],
    ["Required field blank cells", requiredBlanks.length, requiredBlanks.length === 0 ? "OK" : "Review exceptions before import."],
    ["Owner cross-reference matches", ownerMatches, "Exact source Order Number + Line No matches from the previous import workbook."],
    ["Owner cross-reference unmatched", sourceLines.length - ownerMatches, "Owner remains blank; no exact previous source order-line match."],
    ["Source open units", sourceOpenUnits, "All branches."],
    ["Output order quantity", outputOpenUnits, Math.abs(sourceOpenUnits - outputOpenUnits) < 0.000001 ? "OK" : "Does not reconcile."],
    ["Source open amount", sourceOpenAmount, "CAD and USD are mixed here only as a technical checksum."],
    ["Output amount", outputAmount, Math.abs(sourceOpenAmount - outputAmount) < 0.000001 ? "OK" : "Does not reconcile."],
    ["Data-quality exceptions", exceptions.length, "See Exceptions sheet."],
    ["Migration warning", "Lookup validation required", "Vendors, items, branch warehouses, owners, and MAIN locations must exist in Acumatica."],
    ["Source file", path.basename(sourcePath), "Actual open-PO list supplied in the input directory."],
    ["Owner cross-reference", path.basename(ownerCrossReferencePath), "Previous mapped import workbook."],
    ["Template workbook", "Purchase Orders - V1.5.xlsx", "Data Migration tab and 44 target columns."],
  ];
}

async function main() {
  await fs.mkdir(outputDir, { recursive: true });
  await fs.mkdir(workDir, { recursive: true });

  const csvText = await fs.readFile(sourcePath, "utf8");
  const sourceWorkbook = await Workbook.fromCSV(csvText, { sheetName: "Open PO Lines" });
  const lineSheet = sourceWorkbook.worksheets.getItem("Open PO Lines");
  const lineValues = lineSheet.getUsedRange(true).values;
  const lineHeaders = lineValues[0];
  const missingColumns = SOURCE_COLUMNS.filter((column) => !lineHeaders.includes(column));
  if (missingColumns.length) throw new Error(`Missing required source columns: ${missingColumns.join(", ")}`);
  const lines = sortRows(lineValues.slice(1).map((row) => normalizeCsvRow(rowObj(lineHeaders, row))));
  const headers = [...groupBy(lines, sourceKey).values()].map((group) => group[0]);
  const headerMap = new Map(headers.map((r) => [sourceKey(r), r]));

  const ownerWorkbook = await SpreadsheetFile.importXlsx(await FileBlob.load(ownerCrossReferencePath));
  const ownerValues = ownerWorkbook.worksheets.getItem("Data Migration").getUsedRange(true).values;
  const ownerHeaders = getHeaderIndex(ownerValues[0]);
  const ownerByOrderLine = new Map();
  for (const row of ownerValues.slice(1)) {
    const key = orderLineKey({ po_order_no: row[ownerHeaders.get("Order Nbr.")], po_line: row[ownerHeaders.get("Line Nbr.")] });
    const owner = asText(row[ownerHeaders.get("Owner")]);
    if (ownerByOrderLine.has(key) && ownerByOrderLine.get(key) !== owner) throw new Error(`Conflicting Owner values for ${key}.`);
    ownerByOrderLine.set(key, owner);
  }
  const ownerMatches = lines.filter((line) => ownerByOrderLine.has(orderLineKey(line)) && ownerByOrderLine.get(orderLineKey(line)) !== "").length;

  const outputRows = [];
  const exceptions = [];

  for (const line of lines) {
    const header = headerMap.get(sourceKey(line));
    if (!header) addException(exceptions, "Missing header match", "High", line, "No matching All PO Headers row found.");
    const r = { ...line, ...(header ?? {}) };
    const openAmount = toNumber(line.amt_open);
    const openQty = orderQuantity(line.units_open, openAmount);
    const mappedInventoryId = inventoryId(line);
    const unitCost = openQty !== 0 ? openAmount / openQty : null;

    const row = TARGET_HEADERS.map((target) => {
      switch (target) {
        case "Legacy": return true;
        case "Type": return "Normal";
        case "Order Nbr.": return orderNumber(line);
        case "Legacy Order Nbr.": return legacyOrderNumber(line);
        case "Vendor": return asText(line.header_vendor_no || line.vendor);
        case "Location": return "MAIN";
        case "Owner": return ownerByOrderLine.get(orderLineKey(line)) || null;
        case "Vendor Ref.": return null;
        case "Date": return excelSerialToDate(line.order_date);
        case "Promised On": return null;
        case "Description": return truncate(`JDE ${line.po_order_type} PO ${line.po_order_no} Branch ${line.branch}`, 60);
        case "Shipping Destination Type": return "Branch";
        case "Ship To": return asText(line.branch);
        case "Shipping Location": return "MAIN";
        case "Ship To Warehouse": return asText(line.branch);
        case "Line Nbr.": return toNumber(line.po_line);
        case "Branch (Transaction)": return asText(line.branch);
        case "Inventory ID": return mappedInventoryId || null;
        case "Warehouse": return asText(line.branch);
        case "Line Description": return lineDescription(line);
        case "UOM": return asText(line.uom);
        case "Order Qty.": return openQty;
        case "Qty. On Receipts": return 0;
        case "Billed Qty.": return 0;
        case "Unit Cost": return unitCost;
        case "Ext. Cost": return openAmount;
        case "Amount": return openAmount;
        case "Requested": return excelSerialToDate(line.request_date);
        case "Promised": return excelSerialToDate(line.promised_date);
        case "Completed": return false;
        case "Cancelled": return false;
        case "Closed": return false;
        default: return null;
      }
    });

    outputRows.push(row);

    if (!mappedInventoryId || !line.item_id) addException(exceptions, "Missing item mapping", "High", line, "Inventory ID or source item_id is blank.");
    if (openAmount < 0) addException(exceptions, "Negative open amount", "Medium", line, "Open amount is negative and should be validated before import.");
    if (toNumber(line.units_open) === 0 && openAmount !== 0) addException(exceptions, "Derived order quantity", "Medium", line, "Source open quantity is zero; target order quantity defaults to 1 for this amount-only line.");

    for (let i = 0; i < TARGET_HEADERS.length; i += 1) {
      const target = TARGET_HEADERS[i];
      const limit = COLUMN_LIMITS.get(target);
      const value = row[i];
      if (limit && value !== null && value !== undefined && String(value).length > limit) {
        addException(exceptions, "Target field length", "High", line, `${target} length ${String(value).length} exceeds max ${limit}.`);
      }
    }
  }

  const duplicateKeys = [];
  const keyCounts = new Map();
  for (const row of outputRows) {
    const key = `${row[TARGET_HEADERS.indexOf("Order Nbr.")]}|${row[TARGET_HEADERS.indexOf("Line Nbr.")]}`;
    keyCounts.set(key, (keyCounts.get(key) ?? 0) + 1);
  }
  for (const [key, count] of keyCounts) if (count > 1) duplicateKeys.push([key, count]);

  const requiredBlanks = [];
  for (let r = 0; r < outputRows.length; r += 1) {
    for (const col of REQUIRED_COLUMNS) {
      const idx = TARGET_HEADERS.indexOf(col);
      const value = outputRows[r][idx];
      if (value === null || value === undefined || value === "") {
        requiredBlanks.push([r + 2, col, outputRows[r][TARGET_HEADERS.indexOf("Order Nbr.")], outputRows[r][TARGET_HEADERS.indexOf("Line Nbr.")]]);
      }
    }
  }

  const templateWorkbook = await SpreadsheetFile.importXlsx(await FileBlob.load(templatePath));
  const importSheet = templateWorkbook.worksheets.getItem("Data Migration");
  importSheet.getRangeByIndexes(1, 0, outputRows.length, TARGET_HEADERS.length).values = outputRows;
  formatImportSheet(importSheet, outputRows.length + 1);
  templateWorkbook.recalculate();

  const importInspect = await templateWorkbook.inspect({
    kind: "table",
    range: "Data Migration!A1:AR8",
    include: "values,formulas",
    tableMaxRows: 8,
    tableMaxCols: 12,
    maxChars: 5000,
  });
  console.log("IMPORT SAMPLE");
  console.log(importInspect.ndjson);

  const importErrors = await templateWorkbook.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!",
    options: { useRegex: true, maxResults: 300 },
    summary: "final formula error scan",
    maxChars: 5000,
  });
  console.log("IMPORT ERROR SCAN");
  console.log(importErrors.ndjson);

  const importPreview = await templateWorkbook.render({ sheetName: "Data Migration", range: "A1:AR35", scale: 1, format: "png" });
  await fs.writeFile(path.join(workDir, "open_purchase_orders_acumatica_import_preview.png"), new Uint8Array(await importPreview.arrayBuffer()));
  await (await SpreadsheetFile.exportXlsx(templateWorkbook)).save(path.join(outputDir, "open_purchase_orders_acumatica_import.xlsx"));

  const qaWorkbook = Workbook.create();
  const summarySheet = qaWorkbook.worksheets.add("Summary");
  const branchSheet = qaWorkbook.worksheets.add("Branch Totals");
  const typeSheet = qaWorkbook.worksheets.add("Type Validation");
  const mappingSheet = qaWorkbook.worksheets.add("Mapping Notes");
  const exceptionSheet = qaWorkbook.worksheets.add("Exceptions");

  for (const sheet of [summarySheet, branchSheet, typeSheet, mappingSheet, exceptionSheet]) {
    sheet.showGridLines = false;
  }

  summarySheet.getRange("A2").values = [["Open Purchase Orders Migration QA"]];
  summarySheet.getRange("A2").format.font = { name: "Arial", size: 14, bold: true, color: "#1F2933" };
  summarySheet.getRange("A3:C3").format.borders = { bottom: { style: "thin", color: "#6B7280" } };
  addSheetTable(summarySheet, "A5", buildQaSummary(headers, lines, outputRows, exceptions, duplicateKeys, requiredBlanks, ownerMatches));
  summarySheet.getRange("A1:A20").format.columnWidth = 30;
  summarySheet.getRange("B1:B23").format.columnWidth = 44;
  summarySheet.getRange("C1:C20").format.columnWidth = 92;
  summarySheet.getRange("B6:B14").setNumberFormat("#,##0");
  summarySheet.getRange("B15:B18").setNumberFormat("#,##0.00");
  summarySheet.freezePanes.freezeRows(5);

  const branchGroups = groupBy(lines, (r) => `${r.branch}|${r.currency_code}`);
  const branchRows = [["Branch", "Currency", "Source Lines", "Source Open Units", "Source Open Amount", "Output Rows", "Output Order Qty.", "Output Amount", "Delta Lines", "Delta Qty.", "Delta Amount"]];
  for (const [key, group] of [...branchGroups.entries()].sort((a, b) => {
    const [branchA, curA] = a[0].split("|");
    const [branchB, curB] = b[0].split("|");
    return Number(branchA) - Number(branchB) || curA.localeCompare(curB);
  })) {
    const [branch, currency] = key.split("|");
    const sourceLines = group.length;
    const sourceUnits = group.reduce((s, r) => s + orderQuantity(r.units_open, r.amt_open), 0);
    const sourceAmount = group.reduce((s, r) => s + toNumber(r.amt_open), 0);
    const outGroup = outputRows.filter((r) =>
      asText(r[TARGET_HEADERS.indexOf("Branch (Transaction)")]) === branch &&
      lines[outputRows.indexOf(r)]?.currency_code === currency
    );
    const outputUnits = outGroup.reduce((s, r) => s + toNumber(r[TARGET_HEADERS.indexOf("Order Qty.")]), 0);
    const outputAmount = outGroup.reduce((s, r) => s + toNumber(r[TARGET_HEADERS.indexOf("Amount")]), 0);
    branchRows.push([branch, currency, sourceLines, sourceUnits, sourceAmount, outGroup.length, outputUnits, outputAmount, outGroup.length - sourceLines, outputUnits - sourceUnits, outputAmount - sourceAmount]);
  }
  branchSheet.getRange("A2").values = [["Branch and Currency Reconciliation"]];
  branchSheet.getRange("A2").format.font = { name: "Arial", size: 14, bold: true, color: "#1F2933" };
  addSheetTable(branchSheet, "A4", branchRows);
  branchSheet.getRange(`C5:C${branchRows.length + 3}`).setNumberFormat("#,##0");
  branchSheet.getRange(`F5:F${branchRows.length + 3}`).setNumberFormat("#,##0");
  branchSheet.getRange(`D5:E${branchRows.length + 3}`).setNumberFormat("#,##0.00");
  branchSheet.getRange(`G5:K${branchRows.length + 3}`).setNumberFormat("#,##0.00");
  branchSheet.getRange("A1:A12").format.columnWidth = 10;
  branchSheet.getRange("B1:B12").format.columnWidth = 10;
  branchSheet.getRange("C1:K12").format.columnWidth = 16;
  branchSheet.freezePanes.freezeRows(4);

  const typeGroups = groupBy(lines, (r) => `${asText(r.po_order_type)}|${asText(r.line_type)}`);
  const typeRows = [["JDE PO Type", "JDE Line Type", "Lines", "Mapped Acumatica Type", "Validation Note"]];
  for (const [key, group] of [...typeGroups.entries()].sort()) {
    const [poType, lineType] = key.split("|");
    typeRows.push([poType, lineType, group.length, "Normal", "Validate before production import; no template dropdown was supplied."]);
  }
  typeSheet.getRange("A2").values = [["JDE Type Mapping Validation"]];
  typeSheet.getRange("A2").format.font = { name: "Arial", size: 14, bold: true, color: "#1F2933" };
  addSheetTable(typeSheet, "A4", typeRows);
  typeSheet.getRange("A1:E20").format.columnWidth = 24;
  typeSheet.freezePanes.freezeRows(4);

  mappingSheet.getRange("A2").values = [["Target Mapping Notes"]];
  mappingSheet.getRange("A2").format.font = { name: "Arial", size: 14, bold: true, color: "#1F2933" };
  addSheetTable(mappingSheet, "A4", [["Target Column", "Source / Default", "Rule"], ...mappingNotes]);
  mappingSheet.getRange("A1:A50").format.columnWidth = 24;
  mappingSheet.getRange("B1:B50").format.columnWidth = 48;
  mappingSheet.getRange("C1:C50").format.columnWidth = 92;
  mappingSheet.freezePanes.freezeRows(4);

  const exceptionRows = [
    ["Exception Type", "Severity", "Company", "Branch", "PO Number", "JDE PO Type", "Suffix", "Line", "Inventory ID", "Vendor", "Detail"],
    ...exceptions,
    ...duplicateKeys.map(([key, count]) => ["Duplicate Order Nbr. + Line Nbr.", "High", "", "", key.split("|")[0], "", "", key.split("|")[1], "", "", `Duplicate count ${count}.`]),
    ...requiredBlanks.map(([rowNumber, col, orderNo, lineNo]) => ["Required target blank", "High", "", "", orderNo, "", "", lineNo, "", "", `Row ${rowNumber} has blank required field ${col}.`]),
  ];
  exceptionSheet.getRange("A2").values = [["Import Exceptions"]];
  exceptionSheet.getRange("A2").format.font = { name: "Arial", size: 14, bold: true, color: "#1F2933" };
  addSheetTable(exceptionSheet, "A4", exceptionRows);
  exceptionSheet.getRange("A1:A80").format.columnWidth = 26;
  exceptionSheet.getRange("B1:H80").format.columnWidth = 12;
  exceptionSheet.getRange("I1:I80").format.columnWidth = 34;
  exceptionSheet.getRange("J1:J80").format.columnWidth = 14;
  exceptionSheet.getRange("K1:K80").format.columnWidth = 74;
  exceptionSheet.freezePanes.freezeRows(4);
  exceptionSheet.getRange(`A5:K${exceptionRows.length + 3}`).conditionalFormats.add("containsText", {
    text: "High",
    format: { fill: "#FDE2E1", font: { color: "#9F1239", bold: true } },
  });

  qaWorkbook.recalculate();
  const qaInspect = await qaWorkbook.inspect({
    kind: "table",
    range: "Summary!A5:C23",
    include: "values,formulas",
    tableMaxRows: 20,
    tableMaxCols: 3,
    maxChars: 5000,
  });
  console.log("QA SUMMARY");
  console.log(qaInspect.ndjson);

  const qaErrors = await qaWorkbook.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!",
    options: { useRegex: true, maxResults: 300 },
    summary: "final formula error scan",
    maxChars: 5000,
  });
  console.log("QA ERROR SCAN");
  console.log(qaErrors.ndjson);

  for (const sheetName of ["Summary", "Branch Totals", "Type Validation", "Mapping Notes", "Exceptions"]) {
    const preview = await qaWorkbook.render({
      sheetName,
      ...(sheetName === "Exceptions" ? { range: "A1:K40" } : { autoCrop: "all" }),
      scale: 1,
      format: "png",
    });
    await fs.writeFile(path.join(workDir, `${sheetName.replaceAll(" ", "_").toLowerCase()}_preview.png`), new Uint8Array(await preview.arrayBuffer()));
  }
  await (await SpreadsheetFile.exportXlsx(qaWorkbook)).save(path.join(outputDir, "open_purchase_orders_migration_QA.xlsx"));

  const report = {
    outputDir,
    importWorkbook: path.join(outputDir, "open_purchase_orders_acumatica_import.xlsx"),
    qaWorkbook: path.join(outputDir, "open_purchase_orders_migration_QA.xlsx"),
    sourceHeaderRows: headers.length,
    sourceLineRows: lines.length,
    sourceFile: sourcePath,
    ownerCrossReferenceFile: ownerCrossReferencePath,
    ownerCrossReferenceMatches: ownerMatches,
    ownerCrossReferenceUnmatched: lines.length - ownerMatches,
    outputRows: outputRows.length,
    uniquePOs: new Set(lines.map((r) => asText(r.po_order_no))).size,
    duplicateOrderLineKeys: duplicateKeys.length,
    requiredBlankCells: requiredBlanks.length,
    blankOwnerCells: requiredBlanks.filter((row) => row[1] === "Owner").length,
    blankInventoryIdCells: requiredBlanks.filter((row) => row[1] === "Inventory ID").length,
    remainingSourceItem281680: outputRows.filter((row) => asText(row[TARGET_HEADERS.indexOf("Inventory ID")]) === "281680").length,
    exceptionRows: exceptionRows.length - 1,
    branchCurrencies: branchRows.slice(1),
  };
  await fs.writeFile(path.join(workDir, "verification_report.json"), JSON.stringify(report, null, 2));
  console.log("VERIFICATION REPORT");
  console.log(JSON.stringify(report, null, 2));
}

await main();
