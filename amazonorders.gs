// Amazon Orders Import: CSV upload and append to Tiller Transactions sheet.
// Independent of Code.gs; menu in Code.gs calls importAmazonCSV_LocalUpload() defined here.
// Runtime config is read from the "AMZ Import" sheet; defaults below are used only when creating that sheet.
//
// Global scope: all top-level const identifiers use an AMZ_ prefix to avoid collisions with
// QuickSearchSidebar.js when both are in one project (same TillerTools repo). You can copy this
// file plus Amazon HTML + a minimal Code.gs into a standalone project if needed.

const AMZ_IMPORT_SHEET_NAME = "AMZ Import";

/**
 * Sidebar "Offset category" `<select>` first option value (must match AmazonOrdersSidebar.html).
 * Empty string is reserved for explicit "Blank" (no category on offset rows); this sentinel means the user did not choose yet.
 */
const AMZ_OFFSET_CATEGORY_SELECT_VALUE = "__AMZ_OFFSET_SELECT__";

/** Header that identifies a Digital Content Orders CSV (takes precedence if both markers exist). */
const AMZ_DIGITAL_MARKER_HEADER = "Digital Order Item ID";
/** Header that identifies a standard Order History CSV. */
const AMZ_STANDARD_MARKER_HEADER = "Carrier Name & Tracking Number";

/**
 * Default Tiller **Transactions** header for the Date Added column; seeds AMZ Import TABLE4 and matches bundled templates (often column P).
 * Change the cell on AMZ Import if your sheet uses a different header text.
 */
const AMZ_DEFAULT_TILLER_LABEL_DATE_ADDED = "Date Added";

/** Value written to Transactions Source for every row this add-on inserts (orders, returns, offsets, etc.). */
const AMZ_TRANSACTIONS_SOURCE_VALUE = "AmazonCSV";

/** Legacy reference (unused); Transactions column names live on AMZ Import Tiller labels. */
const AMZ_TILLER_CONFIG = {
  SHEET_NAME: "Transactions",
  COLUMNS: {
    DATE: "Date",
    DESCRIPTION: "Description",
    AMOUNT: "Amount",
    TRANSACTION_ID: "Transaction ID",
    FULL_DESCRIPTION: "Full Description",
    DATE_ADDED: AMZ_DEFAULT_TILLER_LABEL_DATE_ADDED,
    MONTH: "Month",
    WEEK: "Week",
    ACCOUNT: "Account",
    ACCOUNT_NUMBER: "Account #",
    INSTITUTION: "Institution",
    ACCOUNT_ID: "Account ID",
    SOURCE: "Source",
    METADATA: "Metadata"
  }
};

// Defaults for seeding the "AMZ Import" sheet when it does not exist.
const AMZ_IMPORT_DEFAULTS = {
  INTRO_ROW: "The below settings are for managing Amazon Order CSV Import.",
  TABLE1_INTRO: "Add one row for each credit card you use with Amazon, and the appropriate values from your Tiller Accounts tab.",
  TABLE1_HEADERS: ["Payment Type", "Account", "Account #", "Institution", "Account ID", "Use for Digital orders?"],
  TABLE1_ROWS: [
    ["Visa - 8534", "Chase Amazon Visa", "xxxx8534", "Chase", "636838acde7b2a0033ff46d5", "Yes"],
    ["Not Available", "Amex", "xxxx4004", "American Express", "636838eea1c01b00330ba247", "No"],
    ["Gift Certificate/Card and Visa - 8534", "Chase Amazon Visa", "xxxx8534", "Chase", "636838acde7b2a0033ff46d5", "No"],
    ["AmericanExpress - 2008", "Amex", "xxxx4004", "American Express", "636838eea1c01b00330ba247", "No"],
    ["AmericanExpress - 2008 and Gift Certificate/Card", "Amex", "xxxx4004", "American Express", "636838eea1c01b00330ba247", "No"]
  ],
  TABLE_CSV_INTRO:
    "Each row maps one Amazon export column. Source file = which CSV; Header = exact column title from row 1 of that file; Name in code = logical field (usually leave as-is); Metadata field name = key in the imported Metadata JSON (blank if none). Rows with Source _file_detection define which column identifies Order History vs Digital orders.",
  TABLE_CSV_HEADERS: ["Source file", "Header", "Name in code", "Metadata field name"],
  TABLE4_TITLE: "Sheet and Column labels used from Tiller",
  TABLE4_HEADERS: ["Name in Code", "Tiller label"],
  TABLE4_ROWS: [
    ["SHEET_NAME", "Transactions"],
    ["DATE", "Date"],
    ["DESCRIPTION", "Description"],
    ["AMOUNT", "Amount"],
    ["TRANSACTION_ID", "Transaction ID"],
    ["FULL_DESCRIPTION", "Full Description"],
    ["DATE_ADDED", AMZ_DEFAULT_TILLER_LABEL_DATE_ADDED],
    ["MONTH", "Month"],
    ["WEEK", "Week"],
    ["ACCOUNT", "Account"],
    ["ACCOUNT_NUMBER", "Account #"],
    ["INSTITUTION", "Institution"],
    ["ACCOUNT_ID", "Account ID"],
    ["SOURCE", "Source"],
    ["METADATA", "Metadata"]
  ]
};

/** Seed data: Order History vs Digital column triples [standardHeader, digitalHeader, nameInCode] for default CSV map. */
const AMZ_SEED_UNIFIED_OH_DO_ROWS = [
  ["Order Date", "Order Date", "Order Date"],
  ["Order ID", "Order ID", "Order ID"],
  ["Product Name", "Product Name", "Product Name"],
  ["Total Amount", "Transaction Amount", "Total Amount"],
  ["ASIN", "ASIN", "ASIN"],
  ["Payment Method Type", "Payment Information", "Payment Method Type"],
  ["Carrier Name & Tracking Number", "", "Carrier Name & Tracking Number"],
  ["Original Quantity", "Original Quantity", "Original Quantity"],
  ["Purchase Order Number", "", "Purchase Order Number"],
  ["Ship Date", "", "Ship Date"],
  ["Shipping Charge", "", "Shipping Charge"],
  ["Total Discounts", "", "Total Discounts"],
  ["Unit Price", "Price", "Unit Price"],
  ["Unit Price Tax", "Price Tax", "Unit Price Tax"],
  ["Website", "", "Website"]
];

/** Seed data: metadata keys matching OH/DO column pairs in {@link AMZ_SEED_UNIFIED_OH_DO_ROWS}. */
const AMZ_SEED_UNIFIED_METADATA_ROWS = [
  ["Order ID", "Order ID", "id"],
  ["Original Quantity", "Original Quantity", "quantity"],
  ["Unit Price", "Price", "item-price"],
  ["Unit Price Tax", "Price Tax", "unit-price-tax"],
  ["Shipping Charge", "", "shipping-charge"],
  ["Total Discounts", "", "total-discounts"],
  ["Total Amount", "Transaction Amount", "total"],
  ["Ship Date", "", "ship-date"],
  ["Carrier Name & Tracking Number", "", "tracking"],
  ["Payment Method Type", "Payment Information", "payment-type"],
  ["Website", "", "site"],
  ["Purchase Order Number", "", "purchase-order"],
  ["purchase", "", "type"]
];

/** Basename, lowercased, forward slashes (for Source file column). */
function amzNormalizeCsvSourceKey_(raw) {
  const s = String(raw || "").trim();
  if (!s) return "";
  const base = s.replace(/\\/g, "/").split("/").pop();
  return String(base).trim().toLowerCase();
}

/**
 * Canonical source keys used in code: order history.csv, digital content orders.csv, refund details.csv, digital returns.csv, _file_detection
 */
function amzCanonicalCsvSourceFile_(raw) {
  const n = amzNormalizeCsvSourceKey_(raw);
  if (!n) return "";
  if (n === "orders history.csv" || n === "order history.csv") return "order history.csv";
  if (n === "digital content orders.csv") return "digital content orders.csv";
  if (n === "returns.csv" || n === "refund details.csv") return "refund details.csv";
  if (n === "digital returns.csv") return "digital returns.csv";
  if (n === "_file_detection" || n === "file_detection") return "_file_detection";
  return n;
}

function amzMetadataKeyForOhDoColumns_(amazonCol, digitalCol) {
  const a = amazonCol != null ? String(amazonCol).trim() : "";
  const d = digitalCol != null ? String(digitalCol).trim() : "";
  for (let i = 0; i < AMZ_SEED_UNIFIED_METADATA_ROWS.length; i++) {
    const r = AMZ_SEED_UNIFIED_METADATA_ROWS[i];
    const ra = r[0] != null ? String(r[0]).trim() : "";
    const rd = r[1] != null ? String(r[1]).trim() : "";
    const k = r[2] != null ? String(r[2]).trim() : "";
    if (ra === a && rd === d) return k;
  }
  return "";
}

/**
 * Default rows for unified Source file / Header / Name in code / Metadata table.
 */
function amzGetDefaultUnifiedCsvMapRows_() {
  const OH = "Order History.csv";
  const DO = "Digital Content Orders.csv";
  const RD = "Refund Details.csv";
  const DR = "Digital Returns.csv";
  const rows = [];
  rows.push(["_file_detection", AMZ_STANDARD_MARKER_HEADER, "standard", ""]);
  rows.push(["_file_detection", AMZ_DIGITAL_MARKER_HEADER, "digital", ""]);
  for (let i = 0; i < AMZ_SEED_UNIFIED_OH_DO_ROWS.length; i++) {
    const r = AMZ_SEED_UNIFIED_OH_DO_ROWS[i];
    const amz = r[0] != null ? String(r[0]).trim() : "";
    const dig = r[1] != null ? String(r[1]).trim() : "";
    const fn = r[2] != null ? String(r[2]).trim() : "";
    if (!fn) continue;
    const mk = amzMetadataKeyForOhDoColumns_(amz, dig);
    if (amz) rows.push([OH, amz, fn, mk]);
    if (dig) rows.push([DO, dig, fn, mk]);
    if (!amz && dig) rows.push([DO, dig, fn, mk]);
  }
  rows.push([OH, "", "purchase", "type"]);
  rows.push([RD, "Order ID", "Order ID", ""]);
  rows.push([RD, "Refund Amount", "Refund Amount", ""]);
  rows.push([RD, "Website", "Website", ""]);
  rows.push([RD, "Refund Date", "Refund Date", ""]);
  rows.push([RD, "Creation Date", "Creation Date", ""]);
  rows.push([RD, "Contract ID", "Contract ID", ""]);
  rows.push(["Returns.csv", "Order ID", "Order ID", ""]);
  rows.push(["Returns.csv", "Refund Amount", "Refund Amount", ""]);
  rows.push(["Returns.csv", "Website", "Website", ""]);
  rows.push(["Returns.csv", "Refund Date", "Refund Date", ""]);
  rows.push(["Returns.csv", "Creation Date", "Creation Date", ""]);
  rows.push(["Returns.csv", "Contract ID", "Contract ID", ""]);
  rows.push([DR, "ASIN", "ASIN", ""]);
  rows.push([DR, "Order ID", "Order ID", ""]);
  rows.push([DR, "Return Date", "Return Date", ""]);
  rows.push([DR, "Transaction Amount", "Transaction Amount", ""]);
  return rows;
}

/**
 * Header string for a logical Name in code on a given Amazon source file (unified map), or "".
 * @param {*} config - from readAmzImportConfig
 * @param {string} sourceCanon - e.g. refund details.csv
 * @param {string} nameInCode - e.g. Order ID
 */
function amzGetSourceMapHeader(config, sourceCanon, nameInCode) {
  if (!config || !config.csvMapBySource || !sourceCanon || !nameInCode) return "";
  const m = config.csvMapBySource[sourceCanon];
  if (!m) return "";
  const h = m[nameInCode];
  return h != null ? String(h).trim() : "";
}

// Required keys in Tiller labels section of AMZ Import.
const AMZ_REQUIRED_TILLER_LABEL_KEYS = [
  "SHEET_NAME", "DATE", "DESCRIPTION", "AMOUNT", "TRANSACTION_ID",
  "FULL_DESCRIPTION", "DATE_ADDED", "MONTH", "WEEK", "ACCOUNT",
  "ACCOUNT_NUMBER", "INSTITUTION", "ACCOUNT_ID", "SOURCE", "METADATA"
];

/** Transactions columns written on each import row (excludes SHEET_NAME) — resolve via {@link amzGetTillerColumnIndex_}. */
const AMZ_WRITTEN_TILLER_LABEL_KEYS = [
  "DATE",
  "DESCRIPTION",
  "AMOUNT",
  "TRANSACTION_ID",
  "FULL_DESCRIPTION",
  "DATE_ADDED",
  "MONTH",
  "WEEK",
  "ACCOUNT",
  "ACCOUNT_NUMBER",
  "INSTITUTION",
  "ACCOUNT_ID",
  "SOURCE",
  "METADATA"
];

const AMZ_IMPORT_INVALID_MSG = "AMZ Import configuration settings are missing or invalid. Suggest deleting that tab to load default values.";
const AMZ_IMPORT_MISSING_CSV_MAP_MSG =
  "AMZ Import must include the CSV column map (header row: Source file, Header, Name in code, Metadata field name). Delete the AMZ Import tab to recreate defaults.";

/** Normalize payment type strings so sheet vs CSV match (NBSP, repeated spaces). */
function amzNormalizePaymentTypeKey(s) {
  return String(s || "")
    .replace(/\u00A0/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

/**
 * True if column A marks the start of a section after the payment table (not a payment type row).
 * Without this, a missing blank row between tables makes the insert target the wrong row.
 */
function amzIsAmzPaymentTableBoundaryRow(fcNormalized) {
  const fc = amzNormalizePaymentTypeKey(fcNormalized);
  if (!fc) return true;
  if (fc === "Source file") return true;
  if (AMZ_IMPORT_DEFAULTS.TABLE_CSV_INTRO && fc === AMZ_IMPORT_DEFAULTS.TABLE_CSV_INTRO) return true;
  if (AMZ_IMPORT_DEFAULTS.TABLE4_TITLE && fc === AMZ_IMPORT_DEFAULTS.TABLE4_TITLE) return true;
  if (AMZ_IMPORT_DEFAULTS.TABLE1_INTRO && fc === AMZ_IMPORT_DEFAULTS.TABLE1_INTRO) return true;
  if (fc.indexOf("Edit the left column only") === 0) return true;
  if (fc.indexOf("Only edit left column") === 0) return true;
  if (fc === "Name in Code" || fc === "Tiller label") return true;
  return false;
}

/** Whole Foods / Amazon Fresh rows in Order History use Website = panda01 (case-insensitive). */
const AMZ_WHOLE_FOODS_WEBSITE = "panda01";

/** Amazon import helpers: distinct names from QuickSearchSidebar.js (same project global scope). */
const AMZ_ACCOUNTS_SHEET_NAME = "Accounts";
/** Accounts sheet: Unique Account Identifier (match 4-digit token from payment string). */
const AMZ_ACCOUNTS_COL_UNIQUE_ID = 6;
/** Account Id (for AMZ Import Account ID column). */
const AMZ_ACCOUNTS_COL_ACCOUNT_ID_FIELD = 7;
const AMZ_ACCOUNTS_COL_ACCOUNT = 10;
const AMZ_ACCOUNTS_COL_ACCOUNT_NUMBER = 11;
const AMZ_ACCOUNTS_COL_INSTITUTION = 12;

const AMZ_CATEGORIES_SHEET_NAME = "Categories";

/** Expected ZIP entry names (basename match is case-insensitive). */
const AMZ_ZIP_ORDER_HISTORY = "order history.csv";
const AMZ_ZIP_DIGITAL_ORDERS = "digital content orders.csv";
const AMZ_ZIP_DIGITAL_RETURNS = "digital returns.csv";
const AMZ_ZIP_REFUND_DETAILS = "refund details.csv";

/** Digital refund / digital return main-row label (Description + Full Description). */
const AMZ_DESC_DIGITAL_REFUND = "[AMZD] Refund";

function amzFormatPhysicalRefundDescription_() {
  return "[AMZ]  Refund";
}

/** Refund main-row text: `[AMZ]  Refund Order ID {id}`. Used for both Description and Full Description on import. */
function amzFormatPhysicalRefundFullDescription_(orderId) {
  const oid = String(orderId == null ? "" : orderId).trim();
  return amzFormatPhysicalRefundDescription_() + " Order ID " + oid;
}

function amzFormatPhysicalRefundOffsetLine_(orderId) {
  const oid = String(orderId == null ? "" : orderId).trim();
  return "[AMZ]  Refund offset Order ID " + oid;
}

/**
 * Prefix for imported purchase lines: [AMZD] + one space, or [AMZ] + two spaces so "Order ID …" aligns.
 * @param {boolean} isDigital
 * @returns {string}
 */
function amzPurchaseLinePrefix_(isDigital) {
  return isDigital ? "[AMZD] " : "[AMZ]  ";
}

/** Description column only: prefix + product title (no Order ID or ASIN). */
function amzFormatPurchaseDescription_(isDigital, productName) {
  const pn = String(productName == null ? "" : productName).trim();
  return amzPurchaseLinePrefix_(isDigital) + pn;
}

/** Full Description: prefix + Order ID + title + ASIN. */
function amzFormatPurchaseFullDescription_(isDigital, orderId, productName, asin) {
  const oid = String(orderId == null ? "" : orderId).trim();
  const pn = String(productName == null ? "" : productName).trim();
  const a = String(asin == null ? "" : asin).trim();
  return (
    amzPurchaseLinePrefix_(isDigital) +
    "Order ID " +
    oid +
    ": " +
    pn +
    " (" +
    a +
    ")"
  );
}

/** Purchase balancing offset row (same string for Description and Full Description). */
function amzFormatPurchaseOffsetLine_(isDigital, orderId, itemCount) {
  const oid = String(orderId == null ? "" : orderId).trim();
  let line = amzPurchaseLinePrefix_(isDigital) + "Purchase offset Order ID " + oid;
  const n = itemCount == null ? NaN : Number(itemCount);
  if (!isNaN(n) && n >= 1) line += n === 1 ? " (1 item)" : " (" + n + " items)";
  return line;
}

/** Digital returns balancing offset (same string for Description and Full Description). */
function amzFormatDigitalReturnOffsetLine_(orderId) {
  const oid = String(orderId == null ? "" : orderId).trim();
  return amzPurchaseLinePrefix_(true) + "Return offset Order ID " + oid;
}

function amzGenerateGuid() {
  return Utilities.getUuid();
}

/** Empty Tiller account fields when payment cannot be resolved (user may edit manually). */
function amzEmptyAccountRow() {
  return { ACCOUNT: "", ACCOUNT_NUMBER: "", INSTITUTION: "", ACCOUNT_ID: "" };
}

/**
 * Last row wins per Order ID. Requires standard or digital file type and mapped Order ID + Payment columns.
 * @returns {Object<string, string>} orderId -> payment method string
 */
function amzOrderIdToPaymentStringMapFromCsv(csvText, config, isDigital) {
  const map = {};
  if (!csvText || String(csvText).trim() === "") return map;
  const csv = Utilities.parseCsv(csvText);
  if (!csv.length || !csv[0].length) return map;
  const headers = csv[0];
  const col = {};
  headers.forEach(function (h, i) {
    if (h != null && String(h).trim() !== "") col[String(h).trim()] = i;
  });
  const kind = amzDetectAmazonCsvFileType(col, config);
  if (isDigital && kind !== "digital") return map;
  if (!isDigital && kind !== "standard") return map;
  const orderIdCol = amzGetCoreCsvColumn(config, "Order ID", isDigital);
  const payCol = amzGetCoreCsvColumn(config, "Payment Method Type", isDigital);
  if (!orderIdCol || !payCol || col[orderIdCol] === undefined || col[payCol] === undefined) return map;
  for (let i = 1; i < csv.length; i++) {
    const r = csv[i];
    const oid = String(r[col[orderIdCol]] == null ? "" : r[col[orderIdCol]]).trim();
    if (!oid) continue;
    map[oid] = String(r[col[payCol]] == null ? "" : r[col[payCol]]).trim();
  }
  return map;
}

function amzResolvePhysicalRefundAccountRow(paymentStr, paymentAccounts) {
  const pt = String(paymentStr || "").trim();
  if (pt) {
    const row = amzLookupPaymentAccountRow(paymentAccounts, pt);
    if (row) return row;
  }
  return amzEmptyAccountRow();
}

/**
 * Digital returns: prefer Order ID join to Digital Content Orders payment string, then digital user row.
 */
function amzResolveDigitalReturnAccountRow(orderId, paymentByOrder, paymentAccounts, digitalUserAccount) {
  const oid = String(orderId == null ? "" : orderId).trim();
  const pt = oid && paymentByOrder[oid] != null ? String(paymentByOrder[oid]).trim() : "";
  if (pt) {
    const row = amzLookupPaymentAccountRow(paymentAccounts, pt);
    if (row) return row;
  }
  if (digitalUserAccount != null) {
    return {
      ACCOUNT: digitalUserAccount.ACCOUNT != null ? String(digitalUserAccount.ACCOUNT) : "",
      ACCOUNT_NUMBER: digitalUserAccount.ACCOUNT_NUMBER != null ? String(digitalUserAccount.ACCOUNT_NUMBER) : "",
      INSTITUTION: digitalUserAccount.INSTITUTION != null ? String(digitalUserAccount.INSTITUTION) : "",
      ACCOUNT_ID: digitalUserAccount.ACCOUNT_ID != null ? String(digitalUserAccount.ACCOUNT_ID) : ""
    };
  }
  return amzEmptyAccountRow();
}

/**
 * @param {*} options - Optional JSON string or object with cutoffDateIso, offsetCategory, skipPanda01, skipNonPanda01,
 *   deferTransactionsSheetPostProcess, bundleImportTimestampIso
 */
function amzParseImportAmazonOptions(options) {
  let opts = {};
  if (options == null || options === "") return opts;
  if (typeof options === "string") {
    try {
      opts = JSON.parse(options);
    } catch (e) {
      opts = {};
    }
  } else if (typeof options === "object") {
    opts = options;
  }
  return opts;
}

/**
 * Parse "yyyy-MM-dd HH:mm:ss" as wall time in the **active spreadsheet** timezone (same as {@link amzFormatImportTimestampStr_}).
 */
function amzParseImportTimestampToDate(s) {
  const str = String(s || "").trim().replace("T", " ");
  if (!/^(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2}):(\d{2})$/.test(str)) return new Date();
  const tz = amzActiveSpreadsheetTimeZoneOrDefault_();
  try {
    return Utilities.parseDate(str, tz, "yyyy-MM-dd HH:mm:ss");
  } catch (e) {
    return new Date();
  }
}

/**
 * Strip BOM, Excel text apostrophe, outer quotes, and formula-style wrappers so date strings parse.
 * @param {*} raw
 * @returns {string}
 */
function amzNormalizeAmazonCsvDateString_(raw) {
  let s = String(raw == null ? "" : raw).trim();
  if (!s) return "";
  if (s.charCodeAt(0) === 0xfeff) s = s.slice(1).trim();
  // Excel "force text" / locale apostrophe before ISO or dates
  while (s.length && (s.charAt(0) === "'" || s.charAt(0) === "\u2019")) s = s.slice(1).trim();
  // Optional spreadsheet formula as text: ="2021-06-14..."
  const eqQuote = s.match(/^="([^"]*)"/);
  if (eqQuote) s = String(eqQuote[1] || "").trim();
  else if (/^=\s*"/.test(s)) s = s.replace(/^=\s*"/, "").replace(/"\s*$/, "").trim();
  if (s.length >= 2 && s.charAt(0) === '"' && s.charAt(s.length - 1) === '"') {
    s = s.slice(1, -1).replace(/""/g, '"').trim();
  }
  return s;
}

/**
 * Parse Amazon CSV date cells (Refund Date, etc.) to script-local start-of-day.
 * Empty or unparseable → null (do not substitute "today").
 * @param {*} raw
 * @returns {Date|null}
 */
function amzParseAmazonCsvDateLoose_(raw) {
  if (raw == null) return null;
  if (raw instanceof Date) {
    if (isNaN(raw.getTime())) return null;
    const out = new Date(raw.getFullYear(), raw.getMonth(), raw.getDate());
    out.setHours(0, 0, 0, 0);
    return out;
  }
  const s = amzNormalizeAmazonCsvDateString_(raw);
  if (!s) return null;
  const iso = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (iso) {
    const y = Number(iso[1]);
    const mo = Number(iso[2]) - 1;
    const d = Number(iso[3]);
    if (y >= 1900 && y <= 2100 && mo >= 0 && mo <= 11 && d >= 1 && d <= 31) {
      const out = new Date(y, mo, d);
      out.setHours(0, 0, 0, 0);
      if (out.getFullYear() === y && out.getMonth() === mo && out.getDate() === d) return out;
    }
  }
  const dt = new Date(s);
  if (isNaN(dt.getTime())) return null;
  const out = new Date(dt.getFullYear(), dt.getMonth(), dt.getDate());
  out.setHours(0, 0, 0, 0);
  return out;
}

/**
 * Refund Details.csv: calendar date for the transaction row. Prefer Refund Date; if empty or
 * unparseable (e.g. "Not Applicable"), use Creation Date when present.
 * @param {Array} r - one CSV row
 * @param {Object<string, number>} col - trimmed header → column index
 * @param {*} [config] - readAmzImportConfig; unified map may rename these columns
 * @returns {Date|null}
 */
function amzResolveRefundDetailsOrderDate_(r, col, config) {
  const refundDateCol = amzGetSourceMapHeader(config, "refund details.csv", "Refund Date") || "Refund Date";
  const creationDateCol = amzGetSourceMapHeader(config, "refund details.csv", "Creation Date") || "Creation Date";
  const hasRefundDate = col[refundDateCol] !== undefined;
  const hasCreationDate = col[creationDateCol] !== undefined;
  if (hasRefundDate) {
    const fromRefund = amzParseAmazonCsvDateLoose_(r[col[refundDateCol]]);
    if (fromRefund) return fromRefund;
  }
  if (hasCreationDate) {
    const fromCreation = amzParseAmazonCsvDateLoose_(r[col[creationDateCol]]);
    if (fromCreation) return fromCreation;
  }
  return null;
}

/**
 * Sum Refund Amount for one Order ID after dropping duplicate CSV lines that repeat the same
 * refund event (same amount and resolved Refund/Creation date — Amazon often repeats rows e.g. by Quantity).
 * @param {Array<Array>} rows - CSV rows for one order
 * @param {Object<string, number>} col - header → index
 * @param {string} refundAmountCol - header name for refund amount
 * @param {*} config - readAmzImportConfig
 * @returns {number}
 */
function amzDedupedRefundSumForOrder_(rows, col, refundAmountCol, config) {
  const seen = Object.create(null);
  let sum = 0;
  for (let j = 0; j < rows.length; j++) {
    const v = parseFloat(rows[j][col[refundAmountCol]]);
    if (isNaN(v)) continue;
    const d = amzResolveRefundDetailsOrderDate_(rows[j], col, config);
    const amtKey = Number(v).toFixed(2);
    const datePart = d && !isNaN(d.getTime()) ? String(d.getTime()) : "nodate:" + j;
    const key = amtKey + "|" + datePart;
    if (seen[key]) continue;
    seen[key] = 1;
    sum += v;
  }
  return sum;
}

/** Max full skipped CSV rows to echo into the import log (status panel). */
const AMZ_MAX_SKIPPED_CSV_ROW_DUMPS = 20;

/**
 * Append one skipped-row detail line for the sidebar status log (not filtered as "Server:" lines).
 * @param {Array<string>} timing
 * @param {{ n: number }} dumpCounter - incremented when a line is appended
 * @param {string} pipelineLabel
 * @param {string} reason
 * @param {*} rowOrRows - one CSV row (array of cells) or array of rows (grouped lines)
 */
function amzLogSkippedCsvDataIfUnderCap_(timing, dumpCounter, pipelineLabel, reason, rowOrRows) {
  if (!timing || dumpCounter.n >= AMZ_MAX_SKIPPED_CSV_ROW_DUMPS) return;
  dumpCounter.n += 1;
  timing.push(
    "Skipped row detail " +
      dumpCounter.n +
      "/" +
      AMZ_MAX_SKIPPED_CSV_ROW_DUMPS +
      " [" +
      pipelineLabel +
      "]: " +
      reason +
      " — " +
      JSON.stringify(rowOrRows)
  );
}

function amzPushSkippedCsvDumpCapNoticeIfNeeded_(timing, dumpCounter) {
  if (timing && dumpCounter && dumpCounter.n >= AMZ_MAX_SKIPPED_CSV_ROW_DUMPS) {
    timing.push("Further skipped CSV row details omitted (limit " + AMZ_MAX_SKIPPED_CSV_ROW_DUMPS + ").");
  }
}

/**
 * Start of the user's cutoff calendar day (midnight local) for CSV row comparisons.
 * Rows with Order/Refund/Return date before this are excluded; on this day are included.
 * @param {Date|null} userCutoff - from cutoffDateIso (often noon) or months lookback
 * @returns {Date|null}
 */
function amzCutoffStartOfDay_(userCutoff) {
  if (!userCutoff || isNaN(userCutoff.getTime())) return null;
  const d = new Date(userCutoff.getFullYear(), userCutoff.getMonth(), userCutoff.getDate());
  d.setHours(0, 0, 0, 0);
  return d;
}

/**
 * @param {string|{ importTimestampStr?: string, sourceFilterOnly?: boolean, legacyMetadataContains?: string }|null|undefined} filterSpec
 * @returns {{ importTimestampStr?: string, sourceFilterOnly?: boolean, legacyMetadataContains?: string }|null}
 */
function amzNormalizePostImportFilterSpec_(filterSpec) {
  if (filterSpec == null || filterSpec === "") return null;
  if (typeof filterSpec === "string") {
    const t = String(filterSpec).trim();
    return t ? { importTimestampStr: t } : null;
  }
  if (typeof filterSpec === "object") return filterSpec;
  return null;
}

/**
 * Sort + filter on Transactions (no AMZ Import re-read). Filter modes:
 * - Normal import: Source = {@link AMZ_TRANSACTIONS_SOURCE_VALUE} and Date Added = run instant ({@code importTimestampStr}).
 * - {@code sourceFilterOnly}: Source only (e.g. TestFilterSort).
 * - {@code legacyMetadataContains}: Metadata whenTextContains (old rows with importer prefix).
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object} tillerCols
 * @param {Object} tillerLabels
 * @param {string} sheetName
 * @param {string|Object|null|undefined} filterSpec - string treated as {@code importTimestampStr}; empty skips filter (sort still runs).
 * @returns {Array<string>} Log lines
 */
function amzApplyTransactionsSortAndFilterCore_(ss, sheet, tillerCols, tillerLabels, sheetName, filterSpec) {
  const lines = [];
  const spec = amzNormalizePostImportFilterSpec_(filterSpec);
  const sourceOnly = spec && spec.sourceFilterOnly === true;
  const legacyMeta =
    spec && spec.legacyMetadataContains != null ? String(spec.legacyMetadataContains).trim() : "";
  const importTs = spec && spec.importTimestampStr != null ? String(spec.importTimestampStr).trim() : "";
  const willApplyRunFilter = importTs !== "" && !sourceOnly && legacyMeta === "";
  const willApplySourceOnly = sourceOnly && legacyMeta === "";
  const willApplyLegacyMeta = legacyMeta !== "" && !sourceOnly;
  const willSetAnyFilter = willApplyRunFilter || willApplySourceOnly || willApplyLegacyMeta;

  const metaCol = amzGetTillerColumnIndex_(tillerCols, tillerLabels.METADATA);
  const sourceCol = amzGetTillerColumnIndex_(tillerCols, tillerLabels.SOURCE);
  const dateAddedCol = amzGetTillerColumnIndex_(tillerCols, tillerLabels.DATE_ADDED);

  // Sort must run on the full data range; an active filter can block rows from reordering.
  try {
    const existingFilter = sheet.getFilter();
    if (existingFilter) {
      existingFilter.remove();
    }
  } catch (e) {
    // continue
  }
  SpreadsheetApp.flush();

  const tSortStart = Date.now();
  const dateCol = amzGetTillerColumnIndex_(tillerCols, tillerLabels.DATE);
  let sortLastRow = 0;
  try {
    sortLastRow = sheet.getLastRow();
  } catch (e) {
    sortLastRow = 0;
  }
  const sortLastCol = Math.max(
    amzNumColsForTransactionRowsSafe(sheet, tillerCols),
    dateCol != null ? dateCol : 1
  );

  if (sortLastRow >= 2 && sortLastCol >= 1 && dateCol != null && dateCol >= 1) {
    try {
      // Prefer Range.sort on data rows only — much faster than Sheet.sort on large grids.
      const sortRange = sheet.getRange(2, 1, sortLastRow, sortLastCol);
      const colInRange = dateCol - sortRange.getColumn() + 1;
      sortRange.sort([{ column: colInRange, ascending: false }]);
      SpreadsheetApp.flush();
    } catch (e) {
      const prevFrozen = sheet.getFrozenRows();
      let frozenTouched = false;
      if (prevFrozen < 1) {
        try {
          sheet.setFrozenRows(1);
          frozenTouched = true;
          SpreadsheetApp.flush();
        } catch (frErr) {
          /* ignore */
        }
      }
      try {
        sheet.sort(dateCol, false);
        SpreadsheetApp.flush();
      } catch (e2) {
        lines.push(
          "Error: sort by Date failed: " +
            (e.message || String(e)) +
            " / sheet.sort: " +
            (e2.message || String(e2))
        );
      }
      if (frozenTouched) {
        try {
          sheet.setFrozenRows(prevFrozen);
          SpreadsheetApp.flush();
        } catch (e3) {
          /* ignore */
        }
      }
    }
  }
  const tSortEnd = Date.now();
  lines.push(
    "Server: sort sheet by Date: " + ((tSortEnd - tSortStart) / 1000).toFixed(2) + " s"
  );

  if (willSetAnyFilter) {
    let lastRowRaw = 0;
    let maxRows = 0;
    let maxCols = 0;
    let numColsForFilter = 0;
    try {
      try {
        SpreadsheetApp.flush();
      } catch (flushErr) {
        /* ignore */
      }

      const sh = ss.getSheetByName(sheetName) || sheet;
      try {
        lastRowRaw = sh.getLastRow();
      } catch (e) {
        lastRowRaw = 0;
      }
      try {
        maxRows = sh.getMaxRows();
      } catch (e) {
        maxRows = 0;
      }
      try {
        maxCols = sh.getMaxColumns();
      } catch (e) {
        maxCols = 0;
      }
      const numColsBase = amzNumColsForTransactionRowsSafe(sh, tillerCols);
      let lastRowForFilter = Math.max(lastRowRaw, 1);
      if (maxRows > 0) {
        lastRowForFilter = Math.min(lastRowForFilter, maxRows);
      }
      let minNeedCol = 1;
      if (willApplyLegacyMeta && metaCol != null && metaCol >= 1) minNeedCol = Math.max(minNeedCol, metaCol);
      if (willApplySourceOnly || willApplyRunFilter) {
        if (sourceCol != null && sourceCol >= 1) minNeedCol = Math.max(minNeedCol, sourceCol);
      }
      if (willApplyRunFilter) {
        if (dateAddedCol != null && dateAddedCol >= 1) minNeedCol = Math.max(minNeedCol, dateAddedCol);
      }
      numColsForFilter = Math.max(numColsBase, minNeedCol, 1);
      if (maxCols > 0) {
        numColsForFilter = Math.min(numColsForFilter, maxCols);
      }

      amzEnsureSheetGridCovers(sh, lastRowForFilter, numColsForFilter);

      if (lastRowForFilter < 1 || numColsForFilter < 1) {
        lines.push(
          "Warning: Transactions filter skipped: invalid dimensions (lastRowForFilter=" +
            lastRowForFilter +
            " numColsForFilter=" +
            numColsForFilter +
            ")."
        );
      } else if (minNeedCol > numColsForFilter) {
        lines.push(
          "Warning: Transactions filter skipped: required column exceeds usable width (need col " +
            minNeedCol +
            " filterCols=" +
            numColsForFilter +
            " maxCols=" +
            maxCols +
            ")."
        );
      } else if (
        willApplyLegacyMeta &&
        (!metaCol || metaCol < 1 || metaCol > numColsForFilter)
      ) {
        lines.push("Warning: Metadata filter skipped: Metadata column missing or out of range.");
      } else if (
        (willApplySourceOnly || willApplyRunFilter) &&
        (!sourceCol || sourceCol < 1 || sourceCol > numColsForFilter)
      ) {
        lines.push('Warning: Transactions filter skipped: Source column missing (add "Source" and AMZ Import SOURCE row).');
      } else if (
        willApplyRunFilter &&
        (!dateAddedCol || dateAddedCol < 1 || dateAddedCol > numColsForFilter)
      ) {
        lines.push("Warning: Transactions filter skipped: Date Added column missing or out of range.");
      } else {
        const dataRange = sh.getRange(1, 1, lastRowForFilter, numColsForFilter);
        const filter = dataRange.createFilter();
        if (willApplyLegacyMeta) {
          const criteria = SpreadsheetApp.newFilterCriteria().whenTextContains(legacyMeta).build();
          const colInRange = metaCol - dataRange.getColumn() + 1;
          filter.setColumnFilterCriteria(colInRange, criteria);
        } else {
          const srcCrit = SpreadsheetApp.newFilterCriteria()
            .whenTextEqualTo(AMZ_TRANSACTIONS_SOURCE_VALUE)
            .build();
          filter.setColumnFilterCriteria(sourceCol - dataRange.getColumn() + 1, srcCrit);
          if (willApplyRunFilter) {
            const runDate = amzParseImportTimestampToDate(importTs);
            const tz = amzActiveSpreadsheetTimeZoneOrDefault_();
            let serial = amzSheetsDateTimeSerialInTimeZone_(runDate, tz);
            if (serial === "" || (typeof serial === "number" && isNaN(serial))) {
              lines.push("Warning: Date Added filter skipped: could not compute serial for import timestamp.");
            } else {
              if (typeof serial === "number") {
                serial = Math.round(serial * 1e10) / 1e10;
              }
              const daCrit = SpreadsheetApp.newFilterCriteria().whenNumberEqualTo(serial).build();
              filter.setColumnFilterCriteria(dateAddedCol - dataRange.getColumn() + 1, daCrit);
            }
          }
        }
      }
    } catch (e) {
      let msg =
        "Warning: Transactions filter failed: " +
        (e.message || String(e)) +
        " (lastRow=" +
        lastRowRaw +
        " maxRows=" +
        maxRows +
        " filterCols=" +
        numColsForFilter +
        " maxCols=" +
        maxCols +
        ")";
      if (maxRows === 0 || maxCols === 0) {
        msg +=
          " [grid reported maxRows/maxCols=0; filter range no longer clamps to those zeros]";
      }
      lines.push(msg);
    }
  }

  return lines;
}

/**
 * Sort Transactions by Date (desc) and apply post-import filter (Source + Date Added, or legacy Metadata substring).
 * Loads AMZ Import config and Transactions sheet — use amzApplyTransactionsSortAndFilterCore_ from import paths
 * that already have sheet/tillerCols to avoid a second config read.
 * @param {string} importTimestampStr - Run id for Date Added filter (yyyy-MM-dd HH:mm:ss).
 * @param {{ metadataFilterContains?: string, sourceFilterOnly?: boolean }|undefined} opts - Test: {@code sourceFilterOnly}; legacy: {@code metadataFilterContains}.
 * @returns {Array<string>} Log lines for the sidebar
 */
function amzApplyTransactionsSortAndFilter(importTimestampStr, opts) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const amzResult = getOrCreateAmzImportSheet();
  const config = readAmzImportConfig(amzResult.sheet);
  const tillerLabels = config.tillerLabels;
  const sheetName = tillerLabels.SHEET_NAME;
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return ["Error: Sheet '" + sheetName + "' not found."];
  }
  const tillerCols = amzGetTillerColumnMap(sheet);
  if (amzGetTillerColumnIndex_(tillerCols, tillerLabels.DATE) == null) {
    return ["Error: Transactions sheet is missing Date column."];
  }
  let filterSpec = null;
  if (opts && opts.sourceFilterOnly === true) {
    filterSpec = { sourceFilterOnly: true };
  } else if (opts && opts.metadataFilterContains != null) {
    filterSpec = { legacyMetadataContains: String(opts.metadataFilterContains).trim() };
  } else {
    const t = String(importTimestampStr || "").trim();
    filterSpec = t ? { importTimestampStr: t } : null;
  }
  return amzApplyTransactionsSortAndFilterCore_(ss, sheet, tillerCols, tillerLabels, sheetName, filterSpec);
}

/**
 * Sidebar troubleshooting: re-run sort + Source filter without re-importing.
 * Shows all rows with Source = {@link AMZ_TRANSACTIONS_SOURCE_VALUE}.
 * @returns {string} Log text for the sidebar
 */
function TestFilterSort() {
  const lines = [];
  lines.push("=== TestFilterSort ===");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const amzResult = getOrCreateAmzImportSheet();
  const config = readAmzImportConfig(amzResult.sheet);
  const err = validateAmzImportConfig(config);
  if (err) {
    lines.push(err);
    return lines.join("\n");
  }
  const tillerLabels = config.tillerLabels;
  const sheetName = tillerLabels.SHEET_NAME;
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    lines.push("Error: Sheet '" + sheetName + "' not found.");
    return lines.join("\n");
  }
  const tillerCols = amzGetTillerColumnMap(sheet);
  const srcCol = amzGetTillerColumnIndex_(tillerCols, tillerLabels.SOURCE);
  if (!srcCol) {
    lines.push(
      "Error: Source column not found (Tiller label: \"" + String(tillerLabels.SOURCE) + "\")."
    );
    return lines.join("\n");
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    lines.push("No transaction rows below the header.");
    lines.push("Running sort only (no filter).");
    const post = amzApplyTransactionsSortAndFilter("");
    for (let pi = 0; pi < post.length; pi++) lines.push(post[pi]);
    return lines.join("\n");
  }
  lines.push("Filter (test): Source equals \"" + AMZ_TRANSACTIONS_SOURCE_VALUE + "\"");
  const post = amzApplyTransactionsSortAndFilter("", { sourceFilterOnly: true });
  for (let pi = 0; pi < post.length; pi++) lines.push(post[pi]);
  return lines.join("\n");
}

/** First 4-digit token (e.g. card last-4) in a payment method string. */
function amzFindFirstFourDigitInString(s) {
  const m = String(s == null ? "" : s).match(/\b(\d{4})\b/);
  return m ? m[1] : null;
}

/**
 * @returns {Array<{ uniqueId: string, accountId: string, account: string, accountNumber: string, institution: string }>}
 */
function amzReadAccountsForLookup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(AMZ_ACCOUNTS_SHEET_NAME);
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  const rows = [];
  for (let r = 1; r < data.length; r++) {
    const f = data[r][AMZ_ACCOUNTS_COL_UNIQUE_ID - 1];
    if (f == null || String(f).trim() === "") continue;
    rows.push({
      uniqueId: String(f).trim(),
      accountId: data[r][AMZ_ACCOUNTS_COL_ACCOUNT_ID_FIELD - 1] != null ? String(data[r][AMZ_ACCOUNTS_COL_ACCOUNT_ID_FIELD - 1]).trim() : "",
      account: data[r][AMZ_ACCOUNTS_COL_ACCOUNT - 1] != null ? String(data[r][AMZ_ACCOUNTS_COL_ACCOUNT - 1]).trim() : "",
      accountNumber: data[r][AMZ_ACCOUNTS_COL_ACCOUNT_NUMBER - 1] != null ? String(data[r][AMZ_ACCOUNTS_COL_ACCOUNT_NUMBER - 1]).trim() : "",
      institution: data[r][AMZ_ACCOUNTS_COL_INSTITUTION - 1] != null ? String(data[r][AMZ_ACCOUNTS_COL_INSTITUTION - 1]).trim() : ""
    });
  }
  return rows;
}

function amzFindAccountRowByFourDigits(accounts, fourDigit) {
  if (!fourDigit) return null;
  for (let i = 0; i < accounts.length; i++) {
    if (accounts[i].uniqueId.indexOf(fourDigit) >= 0) {
      return accounts[i];
    }
  }
  return null;
}

/**
 * Finds a payment row from AMZ Import Table 1 by exact or case-insensitive Payment Type match.
 */
function amzLookupPaymentAccountRow(paymentAccounts, paymentTypeFromCsv) {
  const pt = amzNormalizePaymentTypeKey(paymentTypeFromCsv);
  if (!pt || !paymentAccounts) return null;
  if (Object.prototype.hasOwnProperty.call(paymentAccounts, pt)) return paymentAccounts[pt];
  const lower = pt.toLowerCase();
  const keys = Object.keys(paymentAccounts);
  for (let i = 0; i < keys.length; i++) {
    if (amzNormalizePaymentTypeKey(keys[i]).toLowerCase() === lower) return paymentAccounts[keys[i]];
  }
  return null;
}

function amzPaymentTypeHasRow(paymentAccounts, paymentTypeFromCsv) {
  return amzLookupPaymentAccountRow(paymentAccounts, paymentTypeFromCsv) != null;
}

/**
 * ZIP / sidebar: offset category may be a real name or "" (Blank — no Category on offset rows).
 * Rejects only {@link AMZ_OFFSET_CATEGORY_SELECT_VALUE} ("-- Select Value --") and missing payload field.
 * @param {*} offsetCategoryRaw - {@code payload.offsetCategory} / bundle field
 * @returns {string|null} error message, or null if ok
 */
function amzValidateBundleOffsetCategory_(offsetCategoryRaw) {
  if (offsetCategoryRaw == null) {
    return 'Choose Blank or a category for offset rows—not “-- Select Value --”.';
  }
  const s = String(offsetCategoryRaw).trim();
  if (s === AMZ_OFFSET_CATEGORY_SELECT_VALUE) {
    return 'Choose Blank or a category for offset rows—not “-- Select Value --”.';
  }
  return null;
}

/** Category names from Categories sheet column A (header row 1 = "Category"). */
function amzGetCategoriesListForSidebar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(AMZ_CATEGORIES_SHEET_NAME);
  if (!sh) return [];
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const vals = sh.getRange(2, 1, lastRow, 1).getValues();
  const out = [];
  for (let i = 0; i < vals.length; i++) {
    const c = vals[i][0];
    if (c != null && String(c).trim() !== "") out.push(String(c).trim());
  }
  return out;
}

const AMZ_SIDEBAR_WELCOME_PROP = "AMZ_SIDEBAR_WELCOME_BANNER";

/**
 * Default cutoff (180 days ago), category list, and one-time welcome flag after AMZ Import tab is created.
 * @returns {{ defaultCutoffIso: string, categories: string[], showAmzWelcomeBanner: boolean }}
 */
function getAmazonSidebarInit() {
  const d = new Date();
  d.setDate(d.getDate() - 180);
  const tz = amzActiveSpreadsheetTimeZoneOrDefault_();
  const defaultCutoffIso = Utilities.formatDate(d, tz, "yyyy-MM-dd");
  const props = PropertiesService.getScriptProperties();
  const showBanner = props.getProperty(AMZ_SIDEBAR_WELCOME_PROP) === "1";
  if (showBanner) {
    props.deleteProperty(AMZ_SIDEBAR_WELCOME_PROP);
  }
  return {
    defaultCutoffIso: defaultCutoffIso,
    categories: amzGetCategoriesListForSidebar(),
    showAmzWelcomeBanner: showBanner
  };
}

/**
 * Lists payment method strings in Order History after date/Website filters; suggests Accounts matches for missing AMZ rows.
 * @returns {string} JSON string
 */
function analyzePaymentMethodsForOrderHistory(csvText, cutoffDateIso, includePhysicalOrders, includeWholeFoods) {
  const amzResult = getOrCreateAmzImportSheet();
  const config = readAmzImportConfig(amzResult.sheet);
  const err = validateAmzImportConfig(config);
  if (err) {
    return JSON.stringify({ ok: false, error: err });
  }
  const csv = Utilities.parseCsv(csvText);
  if (!csv.length || !csv[0].length) {
    return JSON.stringify({ ok: false, error: "CSV is empty." });
  }
  const headers = csv[0];
  const col = {};
  headers.forEach(function (h, i) {
    if (h != null && String(h).trim() !== "") col[String(h).trim()] = i;
  });
  if (amzDetectAmazonCsvFileType(col, config) !== "standard") {
    return JSON.stringify({ ok: false, error: "Expected Orders (standard) CSV for payment analysis." });
  }
  const headerErr = amzValidateMappedCsvHeadersPresent(col, config, false);
  if (headerErr) {
    return JSON.stringify({ ok: false, error: headerErr });
  }

  let cutoff = null;
  if (cutoffDateIso) {
    cutoff = new Date(String(cutoffDateIso) + "T12:00:00");
    if (isNaN(cutoff.getTime())) cutoff = null;
  }
  const cutoffStart = amzCutoffStartOfDay_(cutoff);

  const orderDateCol = amzGetCoreCsvColumn(config, "Order Date", false);
  const paymentMethodColName = amzGetCoreCsvColumn(config, "Payment Method Type", false);
  const websiteColName = amzGetCoreCsvColumn(config, "Website", false);
  if (!orderDateCol || !paymentMethodColName || col[paymentMethodColName] === undefined) {
    return JSON.stringify({ ok: false, error: "Missing Order Date or Payment Method Type mapping." });
  }

  const skipPanda01 = includeWholeFoods === false;
  const skipNonPanda01 = includePhysicalOrders === false;

  const paymentTypes = {};
  for (let i = 1; i < csv.length; i++) {
    const r = csv[i];
    let orderDate = new Date(r[col[orderDateCol]]);
    orderDate.setHours(0, 0, 0, 0);
    if (cutoffStart && orderDate < cutoffStart) continue;

    let isWf = false;
    if (websiteColName && col[websiteColName] !== undefined) {
      const wf = String(r[col[websiteColName]] || "").trim().toLowerCase();
      isWf = wf === AMZ_WHOLE_FOODS_WEBSITE;
    }
    if (skipPanda01 && isWf) continue;
    if (skipNonPanda01 && !isWf) continue;

    const pt = String(r[col[paymentMethodColName]] || "").trim();
    if (pt) paymentTypes[pt] = true;
  }

  const existing = config.paymentAccounts;
  const accounts = amzReadAccountsForLookup();
  const missing = [];
  const paymentTypesList = Object.keys(paymentTypes).sort();
  const paymentRows = [];
  paymentTypesList.forEach(function (pt) {
    const existingRow = amzLookupPaymentAccountRow(existing, pt);
    if (existingRow) {
      paymentRows.push({
        paymentType: pt,
        ACCOUNT: existingRow.ACCOUNT,
        ACCOUNT_NUMBER: existingRow.ACCOUNT_NUMBER,
        INSTITUTION: existingRow.INSTITUTION,
        ACCOUNT_ID: existingRow.ACCOUNT_ID,
        status: "configured"
      });
      return;
    }
    const four = amzFindFirstFourDigitInString(pt);
    const acc = amzFindAccountRowByFourDigits(accounts, four);
    missing.push({
      paymentType: pt,
      fourDigit: four,
      suggested: acc
        ? {
            ACCOUNT: acc.account,
            ACCOUNT_NUMBER: acc.accountNumber,
            INSTITUTION: acc.institution,
            ACCOUNT_ID: acc.accountId
          }
        : null
    });
    if (acc) {
      paymentRows.push({
        paymentType: pt,
        ACCOUNT: acc.account,
        ACCOUNT_NUMBER: acc.accountNumber,
        INSTITUTION: acc.institution,
        ACCOUNT_ID: acc.accountId,
        status: "suggested"
      });
    } else {
      paymentRows.push({
        paymentType: pt,
        ACCOUNT: "Unknown",
        ACCOUNT_NUMBER: "Unknown",
        INSTITUTION: "Unknown",
        ACCOUNT_ID: "",
        status: "unknown"
      });
    }
  });

  return JSON.stringify({
    ok: true,
    paymentTypesFound: paymentTypesList,
    missing: missing,
    paymentRows: paymentRows
  });
}

/**
 * Appends rows to AMZ Import Table 1 for suggested payment types (Use for Digital = No).
 * @param {string} rowsJson - JSON array of { paymentType, ACCOUNT, ACCOUNT_NUMBER, INSTITUTION, ACCOUNT_ID }
 */
function insertSuggestedAmzPaymentRows(rowsJson) {
  let rows;
  try {
    rows = JSON.parse(rowsJson);
  } catch (e) {
    return "Error: invalid JSON.";
  }
  if (!rows || !rows.length) return "Nothing to insert.";
  try {
    const amzResult = getOrCreateAmzImportSheet();
    const configForDedupe = readAmzImportConfig(amzResult.sheet);
    const have = configForDedupe.paymentAccounts || {};
    rows = rows.filter(function (r) {
      const pt = r.paymentType != null ? amzNormalizePaymentTypeKey(r.paymentType) : "";
      return pt && !amzPaymentTypeHasRow(have, pt);
    });
    if (!rows.length) return "All suggested payment types already exist on AMZ Import.";
    const sheet = amzResult.sheet;
    const data = sheet.getDataRange().getValues();
    let tableStart = -1;
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === "Payment Type") {
        tableStart = i;
        break;
      }
    }
    if (tableStart < 0) return "Error: AMZ Import payment table not found.";
    let lastData = tableStart;
    for (let k = tableStart + 1; k < data.length; k++) {
      const fc = amzNormalizePaymentTypeKey(data[k][0]);
      if (!fc) break;
      if (amzIsAmzPaymentTableBoundaryRow(fc)) break;
      lastData = k;
    }
    const insertAfterRow = lastData + 1;
    const numRows = rows.length;
    const numCols = 6;
    const out = [];
    for (let r = 0; r < numRows; r++) {
      const row = rows[r];
      out.push([
        row.paymentType != null ? String(row.paymentType) : "",
        row.ACCOUNT != null ? String(row.ACCOUNT) : "",
        row.ACCOUNT_NUMBER != null ? String(row.ACCOUNT_NUMBER) : "",
        row.INSTITUTION != null ? String(row.INSTITUTION) : "",
        row.ACCOUNT_ID != null ? String(row.ACCOUNT_ID) : "",
        "No"
      ]);
    }
    sheet.insertRowsAfter(insertAfterRow, numRows);
    // getRange(row, column, numRows, numColumns) — third arg is row COUNT, not end row.
    sheet.getRange(insertAfterRow + 1, 1, numRows, numCols).setValues(out);
    SpreadsheetApp.flush();
    return "Inserted " + numRows + " payment row(s) on the AMZ Import tab.";
  } catch (e) {
    return "Error: " + (e.message || String(e));
  }
}

function amzGetWeekStartDate(date) {
  const d = new Date(date);
  const day = d.getDay();
  d.setDate(d.getDate() - day);
  d.setHours(0, 0, 0, 0);
  return d;
}

/**
 * Active spreadsheet timezone, or script TZ if no spreadsheet (e.g. rare edge cases).
 * Date Added must match **{@link Spreadsheet#getSpreadsheetTimeZone}**, not only {@link Session#getScriptTimeZone},
 * or calendar day / serial will be wrong when those differ.
 * @returns {string}
 */
function amzActiveSpreadsheetTimeZoneOrDefault_() {
  try {
    return SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  } catch (e) {
    return Session.getScriptTimeZone();
  }
}

/**
 * Import run clock string for metadata {@code importRunAt} and bundle options — always **spreadsheet** timezone
 * so it matches {@link amzSheetsDateTimeSerialInTimeZone_} on Date Added (not {@link Session#getScriptTimeZone}).
 * @param {Date} runTimestamp
 * @returns {string}
 */
function amzFormatImportTimestampStr_(runTimestamp) {
  return Utilities.formatDate(runTimestamp, amzActiveSpreadsheetTimeZoneOrDefault_(), "yyyy-MM-dd HH:mm:ss");
}

/** Log line prefix; row count is summed by {@link amzImportCommitCountFromLog_} for sidebar summaries. */
const AMZ_IMPORT_COMMIT_COUNT_PREFIX = "AMZ_IMPORT_COMMIT_COUNT:";

/**
 * @param {string} logText
 * @returns {number}
 */
function amzImportCommitCountFromLog_(logText) {
  if (logText == null || logText === "") return 0;
  let sum = 0;
  const re = /AMZ_IMPORT_COMMIT_COUNT:(\d+)/g;
  let m;
  const s = String(logText);
  while ((m = re.exec(s)) !== null) {
    const n = Number(m[1]);
    if (!isNaN(n)) sum += n;
  }
  return sum;
}

/**
 * Calendar date at local midnight for instant {@code d} in timezone {@code tz} (not the transaction “order” date).
 * @param {Date} d
 * @param {string} tz - e.g. {@link Spreadsheet#getSpreadsheetTimeZone}
 * @returns {Date}
 */
function amzCalendarDateInTimeZone_(d, tz) {
  if (!(d instanceof Date) || isNaN(d.getTime())) return new Date(NaN);
  const y = Number(Utilities.formatDate(d, tz, "yyyy"));
  const mo = Number(Utilities.formatDate(d, tz, "M"));
  const da = Number(Utilities.formatDate(d, tz, "d"));
  const out = new Date(y, mo - 1, da);
  out.setHours(0, 0, 0, 0);
  return out;
}

/**
 * Calendar date (midnight) in the **spreadsheet** timezone for “today.”
 * Amazon import rows instead use {@code runTimestamp} → full datetime serial on Date Added for per-run filtering.
 */
function amzDateAddedForImportRun_() {
  return amzCalendarDateInTimeZone_(new Date(), amzActiveSpreadsheetTimeZoneOrDefault_());
}

/**
 * Google Sheets / Excel serial date (days since 1899-12-30) for local calendar Y/M/D in the **active spreadsheet** timezone
 * ({@link Spreadsheet#getSpreadsheetTimeZone}), matching sheet display and {@code NOW()}, not {@link Session#getScriptTimeZone}.
 *
 * Why not pass JS Date into {@code Range#setValues}? On some Transactions sheets the Date and Week columns are
 * formatted or treated such that {@code setValues} with a Date leaves the cell empty on readback, while Amount
 * (number) and Month ({@code Utilities.formatDate} string) still write. Serial numbers persist and display as
 * dates when the column format allows; {@code amzCoercePaddedRowsDateWeekToSerial_} runs immediately before each
 * transaction-batch {@code setValues} (also coerces **Date Added** when present in {@code ci}).
 *
 * @param {Date} d
 * @returns {number|string} serial, or "" if invalid
 */
function amzSheetsDateSerial_(d) {
  return amzSheetsDateSerialInTimeZone_(d, amzActiveSpreadsheetTimeZoneOrDefault_());
}

/**
 * Like {@link amzSheetsDateSerial_} but calendar Y/M/D for {@code d} are taken in {@code tz} (match spreadsheet display).
 * Uses {@link Utilities#parseDate} for midnight in {@code tz}; {@code new Date(y,m,d)} would use the script runtime TZ and
 * skew Date Added vs {@code NOW()} when script TZ ≠ spreadsheet TZ (often ~1 hour with US zones or DST quirks).
 * @param {Date} d
 * @param {string} tz
 * @returns {number|string}
 */
function amzSheetsDateSerialInTimeZone_(d, tz) {
  if (!(d instanceof Date) || isNaN(d.getTime())) return "";
  const y = Utilities.formatDate(d, tz, "yyyy");
  const mo = Utilities.formatDate(d, tz, "MM");
  const da = Utilities.formatDate(d, tz, "dd");
  const midnightStr = y + "-" + mo + "-" + da + " 00:00:00";
  try {
    const cal = Utilities.parseDate(midnightStr, tz, "yyyy-MM-dd HH:mm:ss");
    const anchorMs = amzSheetsSerialAnchorMsInTimeZone_(tz);
    return (cal.getTime() - anchorMs) / 86400000;
  } catch (e) {
    return "";
  }
}

/**
 * Sheets serial for a full date+time in {@code tz}. Same epoch as {@link amzSheetsDateSerialInTimeZone_}.
 * Used for Date Added on Amazon import rows so post-import filters can match the run instant with {@code whenNumberEqualTo}.
 * @param {Date} d
 * @param {string} tz
 * @returns {number|string}
 */
function amzSheetsDateTimeSerialInTimeZone_(d, tz) {
  if (!(d instanceof Date) || isNaN(d.getTime())) return "";
  const y = Utilities.formatDate(d, tz, "yyyy");
  const mo = Utilities.formatDate(d, tz, "MM");
  const da = Utilities.formatDate(d, tz, "dd");
  const HH = Utilities.formatDate(d, tz, "HH");
  const mm = Utilities.formatDate(d, tz, "mm");
  const ss = Utilities.formatDate(d, tz, "ss");
  const str = y + "-" + mo + "-" + da + " " + HH + ":" + mm + ":" + ss;
  try {
    const instant = Utilities.parseDate(str, tz, "yyyy-MM-dd HH:mm:ss");
    const anchorMs = amzSheetsSerialAnchorMsInTimeZone_(tz);
    return (instant.getTime() - anchorMs) / 86400000;
  } catch (e) {
    return "";
  }
}

/**
 * Milliseconds for Sheets serial 0: 1899-12-30 00:00:00 interpreted in {@code tz}.
 * @param {string} tz
 * @returns {number}
 */
function amzSheetsSerialAnchorMsInTimeZone_(tz) {
  try {
    return Utilities.parseDate("1899-12-30 00:00:00", tz, "yyyy-MM-dd HH:mm:ss").getTime();
  } catch (e) {
    const anchor = new Date(1899, 11, 30);
    anchor.setHours(0, 0, 0, 0);
    return anchor.getTime();
  }
}

/**
 * Metadata cell value: JSON only (starts with "{"), with {@code importRunAt} for grep and legacy substring filters.
 * @param {Object} amazonInner - object stored under {@code amazon}
 * @param {string} importTimestampStr
 * @returns {string}
 */
function amzImportMetadataJson_(amazonInner, importTimestampStr) {
  return JSON.stringify({
    amazon: amazonInner,
    importRunAt: String(importTimestampStr || "").trim()
  });
}

/**
 * Replaces Date instances at the Date, Week, and (when mapped) Date Added column indices with {@link amzSheetsDateSerial_}
 * values in-place, so import batches survive sheets where native Date values would otherwise store as blank (see {@code amzSheetsDateSerial_}).
 * @param {Array<Array<*>>} padded
 * @param {Object} ci - {@link amzWrittenTillerIndices_} (DATE, WEEK 1-based; DATE_ADDED optional)
 */
function amzCoercePaddedRowsDateWeekToSerial_(padded, ci) {
  if (!padded || !ci || typeof ci.DATE !== "number" || typeof ci.WEEK !== "number") return;
  const di = ci.DATE - 1;
  const wi = ci.WEEK - 1;
  const ai = typeof ci.DATE_ADDED === "number" ? ci.DATE_ADDED - 1 : -1;
  const addedTz = amzActiveSpreadsheetTimeZoneOrDefault_();
  for (let r = 0; r < padded.length; r++) {
    const row = padded[r];
    if (!row) continue;
    if (row[di] instanceof Date && !isNaN(row[di].getTime())) row[di] = amzSheetsDateSerial_(row[di]);
    if (row[wi] instanceof Date && !isNaN(row[wi].getTime())) row[wi] = amzSheetsDateSerial_(row[wi]);
    if (ai >= 0 && row[ai] instanceof Date && !isNaN(row[ai].getTime()))
      row[ai] = amzSheetsDateTimeSerialInTimeZone_(row[ai], addedTz);
  }
}

/** Ensure row-1 header scan spans all Tiller columns even if getLastColumn() is temporarily narrow. */
const AMZ_TRANSACTIONS_HEADER_MIN_COLS = 40;

/** Row 1 header name -> 1-based column index (local to Amazon; Quick Search has its own copy). */
function amzGetTillerColumnMap(sheet) {
  const headerWidth = Math.max(sheet.getLastColumn(), AMZ_TRANSACTIONS_HEADER_MIN_COLS);
  const headers = sheet.getRange(1, 1, 1, headerWidth).getDisplayValues()[0];
  const map = {};
  headers.forEach((h, i) => {
    if (h === "" || h === null) return;
    const k = String(h).trim();
    if (!k) return;
    // First column wins per label (left-to-right). If the same header text appeared twice,
    // only the leftmost column is used.
    if (map[k] !== undefined) return;
    map[k] = i + 1;
  });
  return map;
}

/**
 * 1-based column index for a Transactions header; exact match first, then case-insensitive (Tiller header casing varies).
 * @param {Object} tillerCols - from amzGetTillerColumnMap
 * @param {string} headerLabel - e.g. config tillerLabels.FULL_DESCRIPTION
 * @returns {number|null}
 */
function amzGetTillerColumnIndex_(tillerCols, headerLabel) {
  if (!tillerCols || headerLabel == null || String(headerLabel).trim() === "") return null;
  const exact = tillerCols[headerLabel];
  if (typeof exact === "number" && !isNaN(exact)) return exact;
  const want = String(headerLabel).trim().toLowerCase();
  const keys = Object.keys(tillerCols);
  for (let i = 0; i < keys.length; i++) {
    if (String(keys[i]).trim().toLowerCase() === want) return tillerCols[keys[i]];
  }
  return null;
}

/**
 * Ensures every written Tiller label maps to a Transactions header (case-insensitive).
 * @returns {string|null} Error message or null.
 */
function amzValidateTransactionsImportColumns_(tillerCols, tillerLabels) {
  const missing = [];
  for (let i = 0; i < AMZ_WRITTEN_TILLER_LABEL_KEYS.length; i++) {
    const key = AMZ_WRITTEN_TILLER_LABEL_KEYS[i];
    const lab = tillerLabels[key];
    if (lab == null || String(lab).trim() === "") {
      missing.push('AMZ Import Name in Code "' + key + '" is empty');
      continue;
    }
    if (amzGetTillerColumnIndex_(tillerCols, lab) == null) {
      missing.push(
        'no column header matching "' + String(lab).trim() + '" (' + key + ")"
      );
    }
  }
  if (missing.length) {
    return (
      "Error: Transactions sheet headers or AMZ Import Tiller labels: " +
      missing.join("; ") +
      "."
    );
  }
  return null;
}

/**
 * 1-based indices for columns written on import rows (call after validation).
 * @returns {Object<string, number>}
 */
function amzWrittenTillerIndices_(tillerCols, tillerLabels) {
  const o = {};
  for (let i = 0; i < AMZ_WRITTEN_TILLER_LABEL_KEYS.length; i++) {
    const key = AMZ_WRITTEN_TILLER_LABEL_KEYS[i];
    o[key] = amzGetTillerColumnIndex_(tillerCols, tillerLabels[key]);
  }
  return o;
}

/**
 * Set Description + Full Description on an import row array (0-based indices).
 */
function amzSetRowDescriptionFields_(row, tillerCols, tillerLabels, descriptionText, fullDescriptionText) {
  const dCol = amzGetTillerColumnIndex_(tillerCols, tillerLabels.DESCRIPTION);
  const fCol = amzGetTillerColumnIndex_(tillerCols, tillerLabels.FULL_DESCRIPTION);
  if (dCol != null) row[dCol - 1] = descriptionText;
  if (fCol != null) row[fCol - 1] = fullDescriptionText;
}

/** Largest 1-based column index referenced by the Tiller label map (for row width vs setValues). */
function amzMaxTillerColumnIndex(tillerCols) {
  let m = 1;
  const keys = Object.keys(tillerCols);
  for (let i = 0; i < keys.length; i++) {
    const c = tillerCols[keys[i]];
    if (typeof c === "number" && !isNaN(c)) m = Math.max(m, c);
  }
  return m;
}

/** Width for new transaction rows: never smaller than the rightmost mapped Tiller column. */
function amzNumColsForTransactionRows(sheet, tillerCols) {
  return Math.max(sheet.getLastColumn(), amzMaxTillerColumnIndex(tillerCols), 1);
}

/**
 * Same as amzNumColsForTransactionRows but never throws if getLastColumn() fails (some grids throw
 * "coordinates outside dimensions" when the sheet column extent is inconsistent).
 */
function amzNumColsForTransactionRowsSafe(sheet, tillerCols) {
  const fromMap = Math.max(amzMaxTillerColumnIndex(tillerCols), 1);
  try {
    const lc = sheet.getLastColumn();
    if (lc > 0) return Math.max(fromMap, lc);
  } catch (e) {
    /* ignore */
  }
  return fromMap;
}

/**
 * When getMaxRows/getMaxColumns report smaller than needed, nudge the grid so getRange/createFilter
 * can address the bottom-right cell. Uses setColumnWidth/setRowHeight on that corner only (avoids
 * resizing every row on large sheets).
 */
function amzEnsureSheetGridCovers(sh, lastRow, lastCol) {
  if (!sh || lastRow < 1 || lastCol < 1) return;
  let mr = 0;
  let mc = 0;
  try {
    mr = sh.getMaxRows();
  } catch (e) {
    return;
  }
  try {
    mc = sh.getMaxColumns();
  } catch (e) {
    return;
  }
  if (mr >= lastRow && mc >= lastCol) return;
  const defW = 21;
  const defH = 21;
  try {
    if (mc < lastCol) {
      sh.setColumnWidth(lastCol, defW);
    }
    if (mr < lastRow) {
      sh.setRowHeight(lastRow, defH);
    }
  } catch (e) {
    /* ignore */
  }
}

/** Largest 1-based column index among Tiller label fields written on each import row (not SHEET_NAME). */
function amzMaxRequiredTillerColumn(tillerCols, tillerLabels) {
  let m = 1;
  for (let i = 0; i < AMZ_WRITTEN_TILLER_LABEL_KEYS.length; i++) {
    const c = amzGetTillerColumnIndex_(tillerCols, tillerLabels[AMZ_WRITTEN_TILLER_LABEL_KEYS[i]]);
    if (typeof c === "number" && !isNaN(c)) m = Math.max(m, c);
  }
  return m;
}

/** Row width for setValues: sheet width, all header map columns, every required Tiller field, optional Category. */
function amzNumColsForImportRows(sheet, tillerCols, tillerLabels, categoryColNum) {
  const base = Math.max(
    amzNumColsForTransactionRows(sheet, tillerCols),
    amzMaxRequiredTillerColumn(tillerCols, tillerLabels)
  );
  const cat =
    categoryColNum != null && typeof categoryColNum === "number" && !isNaN(categoryColNum) ? categoryColNum : 0;
  return Math.max(base, cat, 1);
}

/** Log lines for sidebar: resolved column indices, duplicate-index warnings (diagnose blank Date / Date Added / Week). */
function amzTransactionsColumnDebugLines(sheet, tillerCols, tillerLabels, numCols, categoryColNum) {
  const dateCol = amzGetTillerColumnIndex_(tillerCols, tillerLabels.DATE);
  const monthCol = amzGetTillerColumnIndex_(tillerCols, tillerLabels.MONTH);
  const weekCol = amzGetTillerColumnIndex_(tillerCols, tillerLabels.WEEK);
  const dateAddedCol = amzGetTillerColumnIndex_(tillerCols, tillerLabels.DATE_ADDED);
  const descCol = amzGetTillerColumnIndex_(tillerCols, tillerLabels.DESCRIPTION);
  const fullDescCol = amzGetTillerColumnIndex_(tillerCols, tillerLabels.FULL_DESCRIPTION);
  const amountCol = amzGetTillerColumnIndex_(tillerCols, tillerLabels.AMOUNT);
  const txnIdCol = amzGetTillerColumnIndex_(tillerCols, tillerLabels.TRANSACTION_ID);
  const acctCol = amzGetTillerColumnIndex_(tillerCols, tillerLabels.ACCOUNT);
  const sourceCol = amzGetTillerColumnIndex_(tillerCols, tillerLabels.SOURCE);

  let dh = "";
  if (dateCol != null && dateCol >= 1) {
    try {
      dh = sheet.getRange(1, dateCol, 1, 1).getDisplayValues()[0][0];
    } catch (e1) {
      dh = "?";
    }
  }

  const lines = [
    "Server: Transactions columns — Date: " +
      dateCol +
      ' (header="' +
      String(dh) +
      '" expected "' +
      String(tillerLabels.DATE) +
      '") Date Added: ' +
      dateAddedCol +
      " Month: " +
      monthCol +
      " Week: " +
      weekCol +
      " Desc: " +
      descCol +
      " Full Desc: " +
      fullDescCol +
      " Amount: " +
      amountCol +
      " Txn ID: " +
      txnIdCol +
      " Account: " +
      acctCol +
      " Source: " +
      sourceCol +
      " numCols: " +
      numCols +
      " dateInRange: " +
      (dateCol != null && dateCol <= numCols)
  ];

  const dupPairs = [
    ["Date", dateCol],
    ["Date Added", dateAddedCol],
    ["Month", monthCol],
    ["Week", weekCol],
    ["Description", descCol],
    ["Full Description", fullDescCol],
    ["Amount", amountCol],
    ["Txn ID", txnIdCol],
    ["Account", acctCol]
  ];
  const seen = {};
  const dupMsgs = [];
  for (let pi = 0; pi < dupPairs.length; pi++) {
    const label = dupPairs[pi][0];
    const c = dupPairs[pi][1];
    if (c == null || typeof c !== "number") continue;
    if (seen[c]) dupMsgs.push(seen[c] + " & " + label + " → col " + c);
    else seen[c] = label;
  }
  if (dupMsgs.length) {
    lines.push(
      "Server: WARNING: duplicate column indices (text may overwrite dates): " + dupMsgs.join("; ")
    );
  }

  if (categoryColNum != null && typeof categoryColNum === "number" && !isNaN(categoryColNum)) {
    lines.push(
      "Server: offset Category column: " + categoryColNum + " inRange: " + (categoryColNum <= numCols)
    );
  }
  return lines;
}

function amzMaxRowLength(rows) {
  let m = 0;
  for (let i = 0; i < rows.length; i++) {
    m = Math.max(m, rows[i].length);
  }
  return m;
}

/** Pad each row to writeCols; avoids setValues dropping trailing cells when JS arrays grew past numCols. */
function amzPadRowsToWriteCols(rows, writeCols) {
  const out = [];
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r].slice();
    while (row.length < writeCols) row.push("");
    out.push(row);
  }
  return out;
}

function amzIsCellNonempty(v) {
  if (v === "" || v === null) return false;
  if (typeof v === "string" && String(v).trim() === "") return false;
  return true;
}

/**
 * Last row (1-based) that has a value in the given column. Scans from the bottom in chunks so a
 * stray cell far below real data does not make imports append after row 50,000 while ~13k rows
 * hold transactions (sheet.getLastRow() uses any column on the sheet).
 */
function amzGetLastRowWithValueInColumn(sheet, colNum) {
  if (colNum === undefined || colNum < 1) return 1;
  const sheetLast = sheet.getLastRow();
  if (sheetLast < 2) return 1;
  const CHUNK = 5000;
  let end = sheetLast;
  while (end >= 2) {
    const start = Math.max(2, end - CHUNK + 1);
    const numRows = end - start + 1;
    const chunk = sheet.getRange(start, colNum, numRows, 1).getValues();
    for (let i = chunk.length - 1; i >= 0; i--) {
      if (amzIsCellNonempty(chunk[i][0])) {
        return start + i;
      }
    }
    end = start - 1;
  }
  return 1;
}

/**
 * Last row that has a Date value — used for append position and duplicate scan extent.
 * Date-only avoids treating a stray cell in another column (or Full Description) far below as
 * the "end" of the sheet.
 */
function amzGetLastTransactionDataRow(sheet, tillerCols, tillerLabels) {
  const dateCol = amzGetTillerColumnIndex_(tillerCols, tillerLabels.DATE);
  if (!dateCol) return 1;
  return amzGetLastRowWithValueInColumn(sheet, dateCol);
}

/**
 * Dedup keys from Transactions Metadata only (Full Description may be user-edited).
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object} tillerCols
 * @param {Object} tillerLabels
 * @param {number} lastDataRow
 * @param {Set<string>} existingSet
 */
function amzAppendDuplicateKeysFromTransactions_(sheet, tillerCols, tillerLabels, lastDataRow, existingSet) {
  if (lastDataRow < 2) return;
  const metaCol = amzGetTillerColumnIndex_(tillerCols, tillerLabels.METADATA);
  if (!metaCol) return;
  const metas = sheet.getRange(2, metaCol, lastDataRow, 1).getValues();
  for (let j = 0; j < metas.length; j++) {
    amzAddDuplicateKeysFromImportMetadataCell_(metas[j][0], existingSet);
  }
}

/**
 * Last parenthetical on a line — Amazon full descriptions put ASIN/ISBN in the final "(…)".
 * @param {string} s
 * @returns {string}
 */
function amzLastParenToken_(s) {
  const str = String(s || "");
  const open = str.lastIndexOf("(");
  const close = str.lastIndexOf(")");
  if (open < 0 || close <= open) return "";
  return str.substring(open + 1, close).trim();
}

/**
 * Dedup keys from Full Description for legacy rows where Metadata was never filled.
 * Physical: legacy {@code Amazon Order ID …: … (token)} or current {@code [AMZ]  Order ID …: … (token)}.
 * Skips lines starting with {@code [AMZD]} (legacy importer had no digital orders).
 * Returns: {@code Amazon Order ID … with Contract ID …}.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object} tillerCols
 * @param {Object} tillerLabels
 * @param {number} lastDataRow
 * @param {Set<string>} existingSet
 */
function amzAppendLegacyDuplicateKeysFromFullDescription_(sheet, tillerCols, tillerLabels, lastDataRow, existingSet) {
  if (lastDataRow < 2) return;
  const fdCol = amzGetTillerColumnIndex_(tillerCols, tillerLabels.FULL_DESCRIPTION);
  if (!fdCol) return;
  const vals = sheet.getRange(2, fdCol, lastDataRow, 1).getValues();
  for (let j = 0; j < vals.length; j++) {
    amzAddDedupKeysFromFullDescriptionLine_(vals[j][0], existingSet);
  }
}

/**
 * @param {*} lineValue
 * @param {Set<string>} existingSet
 */
function amzAddDedupKeysFromFullDescriptionLine_(lineValue, existingSet) {
  if (lineValue == null || lineValue === "") return;
  const t = String(lineValue).trim();
  if (!t) return;
  if (/^\[AMZD\]/i.test(t)) return;

  const mRet = t.match(/Amazon Order ID\s+([\d-]+)\s+with Contract ID\s+([0-9a-fA-F-]{36})/i);
  if (mRet) {
    const oid = mRet[1].trim();
    const cid = mRet[2].trim().toLowerCase();
    if (oid && cid) existingSet.add("legacy-return|" + oid + "|" + cid);
    return;
  }

  let m = t.match(/^\[AMZ\]\s+Order ID\s+([\d-]+)\s*:\s*(.+)$/);
  if (m) {
    const token = amzLastParenToken_(m[2]);
    if (token) {
      const norm = amzNormalizePurchaseDedupToken_(token);
      if (norm) existingSet.add("physical-purchase-line|" + m[1].trim() + "|" + norm);
    }
    return;
  }

  m = t.match(/^Amazon Order ID\s+([\d-]+)\s*:\s*(.+)$/);
  if (m) {
    const token = amzLastParenToken_(m[2]);
    if (token) {
      const norm = amzNormalizePurchaseDedupToken_(token);
      if (norm) existingSet.add("physical-purchase-line|" + m[1].trim() + "|" + norm);
    }
  }
}

/**
 * @param {string} s
 * @returns {string}
 */
function amzNormalizeAsinDedupKey_(s) {
  return String(s || "").trim().toUpperCase();
}

/**
 * Normalizes ASIN/ISBN token for physical line-item dedup (matches import + Metadata + Full Description scan).
 * All-numeric tokens of length at most 10 are left-padded to ISBN-10 width so legacy vs new imports align.
 * @param {string} s
 * @returns {string}
 */
function amzNormalizePurchaseDedupToken_(s) {
  const raw = String(s == null ? "" : s).trim();
  if (!raw) return "";
  if (/^\d+$/.test(raw)) {
    if (raw.length <= 10) return raw.padStart(10, "0");
    return raw;
  }
  return raw.toUpperCase();
}

/**
 * Add stable dedup key(s) for one parsed {@code parsed.amazon} object (current and older metadata shapes).
 * @param {Object} amz
 * @param {Set<string>} setObj
 */
function amzAddDedupKeysForAmazonMeta_(amz, setObj) {
  if (!amz || amz.type == null || String(amz.type).trim() === "") return;
  const t = String(amz.type);
  if (t === "refund-detail" && amz.orderId != null) {
    const amt = Number(amz.refundAmount);
    const amtKey = isNaN(amt) ? String(amz.refundAmount) : amt.toFixed(2);
    setObj.add("refund-detail|" + String(amz.orderId).trim() + "|" + amtKey);
    return;
  }
  if (t === "return" && amz.id != null) {
    const oid = String(amz.id).trim();
    const cidRaw =
      amz["contract-id"] != null
        ? String(amz["contract-id"]).trim()
        : amz.contractId != null
          ? String(amz.contractId).trim()
          : "";
    const cid = cidRaw.toLowerCase();
    if (oid && cid) setObj.add("legacy-return|" + oid + "|" + cid);
    return;
  }
  if (t === "digital-return" && amz.id != null) {
    setObj.add("digital-return|" + String(amz.id).trim() + "|" + String(amz.asin || "").trim());
    return;
  }
  if (t === "purchase-offset" && amz.id != null) {
    setObj.add("purchase-offset|" + String(amz.id).trim());
    return;
  }
  if (t === "digital-purchase-offset" && amz.id != null) {
    setObj.add("digital-purchase-offset|" + String(amz.id).trim());
    return;
  }
  if (t === "physical-refund-offset" && amz.id != null) {
    setObj.add("physical-refund-offset|" + String(amz.id).trim());
    return;
  }
  if (t === "digital-return-offset" && amz.id != null) {
    setObj.add("digital-return-offset|" + String(amz.id).trim());
    return;
  }
  if (t === "offset" && amz.id != null) {
    setObj.add("purchase-offset|" + String(amz.id).trim());
    return;
  }
  if (t === "digital-purchase" && amz.id != null) {
    setObj.add("digital-purchase|" + String(amz.id).trim());
    return;
  }
  if (t === "purchase") {
    const oid = amz.id != null ? String(amz.id).trim() : "";
    if (!oid) return;
    // Digital Content Orders (aggregated per Order ID) always set lineItemCount; physical line-item rows do not.
    if (amz.lineItemCount != null && String(amz.lineItemCount).trim() !== "") {
      setObj.add("digital-purchase|" + oid);
      return;
    }
    const lineAsin = amz.isbn != null && String(amz.isbn).trim() !== "" ? amz.isbn : amz.asin;
    const asinK = amzNormalizePurchaseDedupToken_(lineAsin);
    setObj.add("physical-purchase-line|" + oid + "|" + asinK);
  }
}

/**
 * @param {*} metaCellValue
 * @param {Set<string>} setObj
 */
function amzAddDuplicateKeysFromImportMetadataCell_(metaCellValue, setObj) {
  if (metaCellValue == null || metaCellValue === "") return;
  const s = String(metaCellValue);
  const brace = s.indexOf("{");
  if (brace < 0) return;
  let parsed;
  try {
    parsed = JSON.parse(s.substring(brace));
  } catch (e) {
    return;
  }
  const amz = parsed && parsed.amazon;
  if (!amz) return;
  amzAddDedupKeysForAmazonMeta_(amz, setObj);
}

/**
 * True if {@code existingSet} has any {@code legacy-return|<orderId>|…} key from Metadata or Full Description scan.
 * Refund Details import uses this so legacy return credits still dedupe when the sheet amount is wrong (legacy
 * summed duplicate CSV rows) and {@code refund-detail|orderId|amount} would not match the deduped CSV total.
 * @param {Set<string>} existingSet
 * @param {*} orderIdRaw
 * @returns {boolean}
 */
function amzAnyLegacyReturnKeyForOrderId_(existingSet, orderIdRaw) {
  const oid = String(orderIdRaw == null ? "" : orderIdRaw).trim();
  if (!oid) return false;
  const p = "legacy-return|" + oid + "|";
  let found = false;
  existingSet.forEach(function (k) {
    if (!found && k.indexOf(p) === 0) found = true;
  });
  return found;
}

/**
 * Sidebar: toast total ZIP import duration (seconds).
 * @param {number} elapsedSec
 * @param {boolean} success
 * @param {string} [errSummary]
 */
function amzNotifyImportBundleFinished(elapsedSec, success, errSummary) {
  const ss = SpreadsheetApp.getActive();
  const raw = typeof elapsedSec === "number" ? elapsedSec : 0;
  const rounded = Math.round(raw * 10) / 10;
  if (success) {
    ss.toast("Import finished in " + rounded + " s", "Amazon import", 5);
  } else {
    const msg = errSummary ? String(errSummary).slice(0, 100) : "Import failed";
    ss.toast(msg + " (" + rounded + " s)", "Amazon import", 8);
  }
}

/**
 * Gets or creates the "AMZ Import" sheet. If it does not exist, creates it and fills with defaults.
 * @returns {{ sheet: GoogleAppsScript.Spreadsheet.Sheet, wasCreated: boolean }}
 */
function getOrCreateAmzImportSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existing = ss.getSheetByName(AMZ_IMPORT_SHEET_NAME);
  if (existing) return { sheet: existing, wasCreated: false };

  const sheet = ss.insertSheet(AMZ_IMPORT_SHEET_NAME);
  let row = 1;
  sheet.getRange(row, 1).setValue(AMZ_IMPORT_DEFAULTS.INTRO_ROW);
  row += 2;

  sheet.getRange(row, 1).setValue(AMZ_IMPORT_DEFAULTS.TABLE1_INTRO);
  row += 1;
  sheet.getRange(row, 1, 1, AMZ_IMPORT_DEFAULTS.TABLE1_HEADERS.length).setValues([AMZ_IMPORT_DEFAULTS.TABLE1_HEADERS]);
  sheet.getRange(row, 1, 1, AMZ_IMPORT_DEFAULTS.TABLE1_HEADERS.length).setFontWeight("bold");
  row += 1;
  if (AMZ_IMPORT_DEFAULTS.TABLE1_ROWS.length) {
    sheet.getRange(row, 1, AMZ_IMPORT_DEFAULTS.TABLE1_ROWS.length, AMZ_IMPORT_DEFAULTS.TABLE1_HEADERS.length)
      .setValues(AMZ_IMPORT_DEFAULTS.TABLE1_ROWS);
    row += AMZ_IMPORT_DEFAULTS.TABLE1_ROWS.length;
  }
  row += 1;

  const csvMapRows = amzGetDefaultUnifiedCsvMapRows_();
  sheet.getRange(row, 1).setValue(AMZ_IMPORT_DEFAULTS.TABLE_CSV_INTRO);
  row += 1;
  sheet.getRange(row, 1, 1, AMZ_IMPORT_DEFAULTS.TABLE_CSV_HEADERS.length).setValues([AMZ_IMPORT_DEFAULTS.TABLE_CSV_HEADERS]);
  sheet.getRange(row, 1, 1, AMZ_IMPORT_DEFAULTS.TABLE_CSV_HEADERS.length).setFontWeight("bold");
  row += 1;
  sheet.getRange(row, 1, csvMapRows.length, AMZ_IMPORT_DEFAULTS.TABLE_CSV_HEADERS.length).setValues(csvMapRows);
  row += csvMapRows.length;
  row += 1;

  sheet.getRange(row, 1).setValue(AMZ_IMPORT_DEFAULTS.TABLE4_TITLE);
  row += 1;
  sheet.getRange(row, 1, 1, AMZ_IMPORT_DEFAULTS.TABLE4_HEADERS.length).setValues([AMZ_IMPORT_DEFAULTS.TABLE4_HEADERS]);
  sheet.getRange(row, 1, 1, AMZ_IMPORT_DEFAULTS.TABLE4_HEADERS.length).setFontWeight("bold");
  row += 1;
  sheet.getRange(row, 1, AMZ_IMPORT_DEFAULTS.TABLE4_ROWS.length, AMZ_IMPORT_DEFAULTS.TABLE4_HEADERS.length)
    .setValues(AMZ_IMPORT_DEFAULTS.TABLE4_ROWS);

  return { sheet: sheet, wasCreated: true };
}

/**
 * Reads config from the "AMZ Import" sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} amzSheet
 * @returns {{
 *   paymentAccounts: Object,
 *   coreMappingStandard: Object,
 *   coreMappingDigital: Object,
 *   metadataMapping: Array<{ key: string, standardCol: string, digitalCol: string }>,
 *   digitalUserAccount: Object|null,
 *   digitalUserYesCount: number,
 *   tillerLabels: Object|null,
 *   csvMapPresent: boolean,
 *   csvMapBySource: Object<string, Object<string, string>>,
 *   csvDetection: { standardHeader?: string, digitalHeader?: string }|null
 * }}
 */
function readAmzImportConfig(amzSheet) {
  const data = amzSheet.getDataRange().getValues();
  const paymentAccounts = {};
  const coreMappingStandard = {};
  const coreMappingDigital = {};
  const metadataMapping = [];
  let digitalUserAccount = null;
  let digitalUserYesCount = 0;
  let tillerLabels = null;
  let csvMapPresent = false;
  /** @type {Object<string, Object<string, string>>} */
  const csvMapBySource = {};
  let csvDetection = null;

  let i = 0;
  while (i < data.length) {
    const row = data[i];
    const first = row[0] ? String(row[0]).trim() : "";
    const second = row[1] ? String(row[1]).trim() : "";
    const third = row[2] != null ? String(row[2]).trim() : "";
    const fourth = row[3] != null ? String(row[3]).trim() : "";

    if (first === "Source file" && second === "Header" && third === "Name in code" && fourth === "Metadata field name") {
      csvMapPresent = true;
      const buf = [];
      i += 1;
      while (i < data.length) {
        const r = data[i];
        const c0 = r[0] != null ? String(r[0]).trim() : "";
        const c1 = r[1] != null ? String(r[1]).trim() : "";
        const c2 = r[2] != null ? String(r[2]).trim() : "";
        const c3 = r[3] != null ? String(r[3]).trim() : "";
        if (!c0 && !c1 && !c2 && !c3) break;
        if (c0.indexOf("Each row maps one Amazon export column") === 0) break;
        if (c0 === "Sheet and Column labels used from Tiller") break;
        if (c0 === "Name in Code" && String(r[1] || "").trim() === "Tiller label") break;
        if (amzIsAmzPaymentTableBoundaryRow(c0) && c0 !== "_file_detection") break;
        buf.push({ rawSource: c0, header: c1, nameCode: c2, meta: c3 });
        i += 1;
      }
      const metaAcc = {};
      for (let b = 0; b < buf.length; b++) {
        const u = buf[b];
        const canon = amzCanonicalCsvSourceFile_(u.rawSource);
        const storeKey = canon === "returns.csv" ? "refund details.csv" : canon;
        if (canon === "_file_detection") {
          if (!csvDetection) csvDetection = {};
          if (u.nameCode === "standard" && u.header) csvDetection.standardHeader = u.header;
          if (u.nameCode === "digital" && u.header) csvDetection.digitalHeader = u.header;
          continue;
        }
        if (u.header && u.nameCode) {
          if (!csvMapBySource[storeKey]) csvMapBySource[storeKey] = {};
          csvMapBySource[storeKey][u.nameCode] = u.header;
        }
        if (u.meta) {
          if (!metaAcc[u.meta]) metaAcc[u.meta] = {};
          const g = metaAcc[u.meta];
          if (storeKey === "order history.csv") {
            if (u.header) g.stdH = u.header;
            else if (u.nameCode) g.stdLit = u.nameCode;
          } else if (storeKey === "digital content orders.csv") {
            if (u.header) g.digH = u.header;
            else if (u.nameCode) g.digLit = u.nameCode;
          }
        }
      }
      for (const mk in metaAcc) {
        const g = metaAcc[mk];
        const std = (g.stdH || g.stdLit || "").trim();
        const dig = (g.digH || g.digLit || "").trim();
        if (std || dig) metadataMapping.push({ key: mk, standardCol: std, digitalCol: dig });
      }
      const oh = csvMapBySource["order history.csv"];
      const dgo = csvMapBySource["digital content orders.csv"];
      if (oh)
        Object.keys(oh).forEach(function (k) {
          coreMappingStandard[k] = oh[k];
        });
      if (dgo)
        Object.keys(dgo).forEach(function (k) {
          coreMappingDigital[k] = dgo[k];
        });
      continue;
    }

    if (first === "Payment Type") {
      i += 1;
      while (i < data.length) {
        const fc = amzNormalizePaymentTypeKey(data[i][0]);
        if (!fc) break;
        if (amzIsAmzPaymentTableBoundaryRow(fc)) break;
        const r = data[i];
        const paymentType = amzNormalizePaymentTypeKey(r[0]);
        const userForDigital = r[5] != null && String(r[5]).trim().toLowerCase() === "yes";
        if (userForDigital) {
          digitalUserYesCount += 1;
          digitalUserAccount = {
            ACCOUNT: r[1] != null ? String(r[1]) : "",
            ACCOUNT_NUMBER: r[2] != null ? String(r[2]) : "",
            INSTITUTION: r[3] != null ? String(r[3]) : "",
            ACCOUNT_ID: r[4] != null ? String(r[4]) : ""
          };
        }
        paymentAccounts[paymentType] = {
          ACCOUNT: r[1] != null ? String(r[1]) : "",
          ACCOUNT_NUMBER: r[2] != null ? String(r[2]) : "",
          INSTITUTION: r[3] != null ? String(r[3]) : "",
          ACCOUNT_ID: r[4] != null ? String(r[4]) : ""
        };
        i += 1;
      }
      continue;
    }

    if (first === "Name in Code" && second === "Tiller label") {
      tillerLabels = {};
      i += 1;
      while (i < data.length && (data[i][0] && String(data[i][0]).trim() !== "")) {
        const r = data[i];
        const nameInCode = r[0] != null ? String(r[0]).trim() : "";
        const tillerLabel = r[1] != null ? String(r[1]).trim() : "";
        if (nameInCode) tillerLabels[nameInCode] = tillerLabel;
        i += 1;
      }
      continue;
    }
    i += 1;
  }

  return {
    paymentAccounts,
    coreMappingStandard,
    coreMappingDigital,
    metadataMapping,
    digitalUserAccount,
    digitalUserYesCount,
    tillerLabels,
    csvMapPresent: csvMapPresent,
    csvMapBySource: csvMapBySource,
    csvDetection: csvDetection
  };
}

/**
 * @param {Object} col - CSV header name -> column index
 * @param {*} [config] - from readAmzImportConfig; optional.csvDetection overrides marker column names
 * @returns {"digital"|"standard"|null}
 */
function amzDetectAmazonCsvFileType(col, config) {
  const digH =
    config && config.csvDetection && config.csvDetection.digitalHeader
      ? String(config.csvDetection.digitalHeader).trim()
      : AMZ_DIGITAL_MARKER_HEADER;
  const stdH =
    config && config.csvDetection && config.csvDetection.standardHeader
      ? String(config.csvDetection.standardHeader).trim()
      : AMZ_STANDARD_MARKER_HEADER;
  if (col[digH] !== undefined && col[digH] !== null) return "digital";
  if (col[stdH] !== undefined && col[stdH] !== null) return "standard";
  return null;
}

/**
 * Validates core mappings from the unified CSV map: every non-empty mapped header for this file type must exist in the CSV.
 * Metadata mappings are not validated here: if a mapping "column" name is absent from the CSV headers,
 * {@link amzBuildAmazonMetadataObject} uses that string as a literal (e.g. type = "purchase").
 * @param {Object} col
 * @param {{ coreMappingStandard: Object, coreMappingDigital: Object, metadataMapping: Array }} config
 * @param {boolean} isDigital
 * @returns {string|null} Error message or null if OK.
 */
function amzValidateMappedCsvHeadersPresent(col, config, isDigital) {
  const map = isDigital ? config.coreMappingDigital : config.coreMappingStandard;
  const other = isDigital ? config.coreMappingStandard : config.coreMappingDigital;
  const fieldNames = {};
  Object.keys(map).forEach(function (k) { fieldNames[k] = true; });
  Object.keys(other).forEach(function (k) { fieldNames[k] = true; });

  for (const fieldName in fieldNames) {
    const resolved = isDigital
      ? (config.coreMappingDigital[fieldName] || "")
      : (config.coreMappingStandard[fieldName] || "");
    if (!resolved || String(resolved).trim() === "") continue;
    if (col[resolved] === undefined || col[resolved] === null) {
      return "Missing required CSV column for this file type: \"" + resolved + "\" (logical field: " + fieldName + "). Check AMZ Import CSV column map.";
    }
  }

  return null;
}

/**
 * Resolves core mapping for one logical field for the current file type.
 */
function amzGetCoreCsvColumn(config, fieldName, isDigital) {
  if (isDigital) {
    const d = config.coreMappingDigital[fieldName];
    if (d != null && String(d).trim() !== "") return String(d).trim();
    return config.coreMappingStandard[fieldName] || "";
  }
  const s = config.coreMappingStandard[fieldName];
  if (s != null && String(s).trim() !== "") return String(s).trim();
  return config.coreMappingDigital[fieldName] || "";
}

/**
 * Validates that all required AMZ Import config is present. No fallbacks.
 * @param {*} config
 * @returns {string|null} Null if valid; otherwise the error message to show.
 */
function validateAmzImportConfig(config) {
  if (!config) return AMZ_IMPORT_INVALID_MSG;
  if (!config.paymentAccounts || Object.keys(config.paymentAccounts).length === 0) {
    return AMZ_IMPORT_INVALID_MSG;
  }
  if (!config.tillerLabels || typeof config.tillerLabels !== "object") {
    return AMZ_IMPORT_INVALID_MSG;
  }
  for (let k = 0; k < AMZ_REQUIRED_TILLER_LABEL_KEYS.length; k++) {
    const key = AMZ_REQUIRED_TILLER_LABEL_KEYS[k];
    const val = config.tillerLabels[key];
    if (val === undefined || val === null || String(val).trim() === "") {
      return AMZ_IMPORT_INVALID_MSG;
    }
  }
  if (config.csvMapPresent !== true) {
    return AMZ_IMPORT_MISSING_CSV_MAP_MSG;
  }
  const need = ["Order Date", "Order ID", "Product Name", "Total Amount", "ASIN"];
  for (let n = 0; n < need.length; n++) {
    const fieldName = need[n];
    const hasStd = config.coreMappingStandard[fieldName] && String(config.coreMappingStandard[fieldName]).trim() !== "";
    const hasDig = config.coreMappingDigital[fieldName] && String(config.coreMappingDigital[fieldName]).trim() !== "";
    if (!hasStd && !hasDig) return AMZ_IMPORT_INVALID_MSG;
  }
  return null;
}

/**
 * Resolves which CSV header name exists in col for metadata: digital column first, then standard,
 * then (for digital) coreMappingDigital when standardCol matches a logical field's Amazon column.
 * @param {{ key: string, standardCol: string, digitalCol: string }} m
 * @param {Object} col - header name -> index
 * @param {boolean} isDigital
 * @param {Object|null} coreMappingStandard
 * @param {Object|null} coreMappingDigital
 * @returns {string} Header to use for col[...] lookup (may still be missing from col)
 */
function amzResolveMetadataColumnName(m, col, isDigital, coreMappingStandard, coreMappingDigital) {
  let src = isDigital ? m.digitalCol : m.standardCol;
  if (src == null || String(src).trim() === "") {
    src = isDigital ? m.standardCol : m.digitalCol;
  }
  src = src != null ? String(src).trim() : "";
  if (!src) return "";
  if (col[src] !== undefined && col[src] !== null) return src;
  if (isDigital && coreMappingStandard && coreMappingDigital) {
    const keys = Object.keys(coreMappingStandard);
    for (let i = 0; i < keys.length; i++) {
      const fn = keys[i];
      const std = coreMappingStandard[fn] != null ? String(coreMappingStandard[fn]).trim() : "";
      if (std === src) {
        const dig = coreMappingDigital[fn] != null ? String(coreMappingDigital[fn]).trim() : "";
        if (dig && col[dig] !== undefined && col[dig] !== null) return dig;
      }
    }
  }
  return src;
}

/**
 * Builds the metadata amazon object for one CSV row using the metadata mapping.
 * @param {Array} csvRow - CSV row (array of values)
 * @param {Object} col - Map of CSV column name -> index
 * @param {Array} metadataMapping - Array of { key, standardCol, digitalCol }
 * @param {boolean} isDigital
 * @param {Object} [coreMappingStandard] - Order History CSV column per logical field (for digital fallback)
 * @param {Object} [coreMappingDigital] - Digital orders CSV column per logical field
 * @returns {Object}
 */
function amzBuildAmazonMetadataObject(csvRow, col, metadataMapping, isDigital, coreMappingStandard, coreMappingDigital) {
  const numericKeys = ["quantity", "item-price", "unit-price-tax", "shipping-charge", "total-discounts", "total"];
  const obj = {};
  metadataMapping.forEach(function (m) {
    const k = m.key;
    const src = amzResolveMetadataColumnName(m, col, isDigital, coreMappingStandard, coreMappingDigital);
    if (!src) return;

    const colIndex = col[src];
    let val;
    if (colIndex !== undefined && colIndex !== null) {
      const raw = csvRow[colIndex];
      if (raw === "" || raw === undefined || raw === null) {
        val = numericKeys.indexOf(k) >= 0 ? 0 : "";
      } else if (numericKeys.indexOf(k) >= 0) {
        val = parseFloat(raw);
        if (isNaN(val)) val = 0;
      } else {
        val = String(raw).trim();
      }
    } else {
      if (numericKeys.indexOf(k) >= 0) val = 0;
      else val = src;
    }
    obj[k] = val;
  });
  return obj;
}

/**
 * For Digital Content Orders CSVs where Amazon emits multiple lines per Order ID (e.g. item + tax),
 * merge metadata: sum any field whose mapped column is the same as totalAmountColName; otherwise use the first row.
 * @param {Array<Array>} rows - CSV rows belonging to one Order ID
 * @param {Object} col - header name -> column index
 * @param {Array} metadataMapping
 * @param {boolean} isDigital
 * @param {string} totalAmountColName - mapped "Total Amount" / Transaction Amount header (digital = Transaction Amount)
 * @param {Object} [coreMappingStandard]
 * @param {Object} [coreMappingDigital]
 */
function amzBuildAmazonMetadataObjectFromRows(
  rows,
  col,
  metadataMapping,
  isDigital,
  totalAmountColName,
  coreMappingStandard,
  coreMappingDigital
) {
  if (!rows || rows.length === 0) return {};
  if (rows.length === 1) {
    const one = amzBuildAmazonMetadataObject(
      rows[0], col, metadataMapping, isDigital, coreMappingStandard, coreMappingDigital
    );
    one.lineItemCount = 1;
    return one;
  }
  const numericKeys = ["quantity", "item-price", "unit-price-tax", "shipping-charge", "total-discounts", "total"];
  const obj = {};
  metadataMapping.forEach(function (m) {
    const k = m.key;
    const resolvedSrc = amzResolveMetadataColumnName(m, col, isDigital, coreMappingStandard, coreMappingDigital);
    if (!resolvedSrc) return;

    if (resolvedSrc === totalAmountColName) {
      let sum = 0;
      for (let i = 0; i < rows.length; i++) {
        const raw = rows[i][col[resolvedSrc]];
        if (raw !== "" && raw !== undefined && raw !== null) {
          const v = parseFloat(raw);
          if (!isNaN(v)) sum += v;
        }
      }
      if (numericKeys.indexOf(k) >= 0) {
        obj[k] = sum;
      } else {
        obj[k] = String(sum);
      }
    } else {
      const part = amzBuildAmazonMetadataObject(rows[0], col, [m], isDigital, coreMappingStandard, coreMappingDigital);
      if (part[k] !== undefined) obj[k] = part[k];
    }
  });
  obj.lineItemCount = rows.length;
  return obj;
}

/**
 * Opens the Amazon Orders import sidebar (HTML from AmazonOrdersSidebar.html).
 * Creates AMZ Import on first run (no popup — sidebar explains defaults and can add payment rows).
 */
function openAmazonOrdersSidebar() {
  const result = getOrCreateAmzImportSheet();
  if (result.wasCreated) {
    PropertiesService.getScriptProperties().setProperty(AMZ_SIDEBAR_WELCOME_PROP, "1");
  }
  const config = readAmzImportConfig(result.sheet);
  const err = validateAmzImportConfig(config);
  if (err) {
    SpreadsheetApp.getUi().alert(err);
    return;
  }
  if (config.tillerLabels && config.tillerLabels.SHEET_NAME) {
    const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.tillerLabels.SHEET_NAME);
    if (targetSheet) targetSheet.activate();
  }
  const html = HtmlService.createHtmlOutputFromFile("AmazonOrdersSidebar").setTitle("Tiller™ Amazon Import");
  SpreadsheetApp.getUi().showSidebar(html);
}

/** @deprecated Use openAmazonOrdersSidebar — kept for any saved menu bindings. */
function importAmazonCSV_LocalUpload() {
  openAmazonOrdersSidebar();
}

/**
 * Imports Amazon CSV data into the Tiller Transactions sheet.
 * Called from HTML via google.script.run.importAmazonRecent(text, months, optionsJson).
 * @param {string} csvText - Raw CSV file content
 * @param {number|null} months - Optional months lookback; null = all rows (ignored if options.cutoffDateIso set)
 * @param {string|undefined} options - Optional JSON: cutoffDateIso, offsetCategory, skipPanda01, skipNonPanda01
 * @returns {string} Newline-separated summary and timing lines for the dialog log
 */
function importAmazonRecent(csvText, months, options) {
  const t0 = Date.now();
  const timing = [];

  const opts = amzParseImportAmazonOptions(options);
  const offsetCategory =
    opts.offsetCategory != null && String(opts.offsetCategory).trim() !== ""
      ? String(opts.offsetCategory).trim()
      : "";
  const skipPanda01 = opts.skipPanda01 === true;
  const skipNonPanda01 = opts.skipNonPanda01 === true;
  const deferPost = opts.deferTransactionsSheetPostProcess === true;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const amzResult = getOrCreateAmzImportSheet();
  const config = readAmzImportConfig(amzResult.sheet);
  const paymentAccounts = config.paymentAccounts;
  const tillerLabels = config.tillerLabels;

  const configErr = validateAmzImportConfig(config);
  if (configErr) return configErr;

  const sheetName = tillerLabels.SHEET_NAME;
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return "Error: Sheet '" + sheetName + "' not found.";

  const tillerCols = amzGetTillerColumnMap(sheet);
  const colMapErr0 = amzValidateTransactionsImportColumns_(tillerCols, tillerLabels);
  if (colMapErr0) return colMapErr0;
  const categoryColNum = offsetCategory
    ? amzGetTillerColumnIndex_(tillerCols, "Category") ||
      amzGetTillerColumnIndex_(tillerCols, "Categories")
    : null;

  const tParseStart = Date.now();
  const csv = Utilities.parseCsv(csvText);
  const tParseEnd = Date.now();
  timing.push("Server: parse CSV: " + ((tParseEnd - tParseStart) / 1000).toFixed(2) + " s");

  if (!csv.length || !csv[0].length) {
    return "Error: CSV is empty or has no header row.";
  }

  const headers = csv[0];
  const col = {};
  headers.forEach(function (h, i) {
    if (h != null && String(h).trim() !== "") col[String(h).trim()] = i;
  });

  const fileKind = amzDetectAmazonCsvFileType(col, config);
  if (!fileKind) {
    const digH =
      config.csvDetection && config.csvDetection.digitalHeader
        ? String(config.csvDetection.digitalHeader).trim()
        : AMZ_DIGITAL_MARKER_HEADER;
    const stdH =
      config.csvDetection && config.csvDetection.standardHeader
        ? String(config.csvDetection.standardHeader).trim()
        : AMZ_STANDARD_MARKER_HEADER;
    return "Could not detect file type. The CSV must include column \"" + digH + "\" (Digital orders) or \"" + stdH + "\" (Orders).";
  }
  const isDigital = fileKind === "digital";
  const detectedLabel = isDigital
    ? "Detected file type: Digital orders"
    : "Detected file type: Orders";

  if (isDigital) {
    if (config.digitalUserYesCount === 0) {
      return detectedLabel + "\n" + "Digital orders import requires exactly one row on AMZ Import (payment table) with \"Use for Digital orders?\" set to Yes.";
    }
    if (config.digitalUserYesCount > 1) {
      return detectedLabel + "\n" + "Digital orders import: multiple rows have \"Use for Digital orders?\" set to Yes. Only one row should be Yes.";
    }
    if (!config.digitalUserAccount) {
      return detectedLabel + "\n" + "Digital orders import: could not read account fields from the row with Use for Digital orders? = Yes.";
    }
  }

  const headerErr = amzValidateMappedCsvHeadersPresent(col, config, isDigital);
  if (headerErr) return detectedLabel + "\n" + headerErr;

  let cutoff = null;
  if (opts.cutoffDateIso) {
    cutoff = new Date(String(opts.cutoffDateIso) + "T12:00:00");
    if (isNaN(cutoff.getTime())) cutoff = null;
  } else if (months) {
    cutoff = new Date();
    cutoff.setMonth(cutoff.getMonth() - months);
  }

  const cutoffStart = amzCutoffStartOfDay_(cutoff);

  const lastDataRow = amzGetLastTransactionDataRow(sheet, tillerCols, tillerLabels);
  const existingFullDescSet = new Set();
  const tDupStart = Date.now();
  amzAppendDuplicateKeysFromTransactions_(sheet, tillerCols, tillerLabels, lastDataRow, existingFullDescSet);
  amzAppendLegacyDuplicateKeysFromFullDescription_(sheet, tillerCols, tillerLabels, lastDataRow, existingFullDescSet);
  const tDupEnd = Date.now();
  timing.push(
    "Server: scan Metadata + Full Description for dedup keys (" +
    existingFullDescSet.size + " entries): " +
    ((tDupEnd - tDupStart) / 1000).toFixed(2) + " s"
  );

  const orderDateCol = amzGetCoreCsvColumn(config, "Order Date", isDigital);
  const orderIdCol = amzGetCoreCsvColumn(config, "Order ID", isDigital);
  const productNameCol = amzGetCoreCsvColumn(config, "Product Name", isDigital);
  const asinCol = amzGetCoreCsvColumn(config, "ASIN", isDigital);
  const totalAmountCol = amzGetCoreCsvColumn(config, "Total Amount", isDigital);
  const paymentMethodColName = amzGetCoreCsvColumn(config, "Payment Method Type", isDigital);

  if (!orderDateCol || !orderIdCol || !productNameCol || !asinCol || !totalAmountCol) {
    return "AMZ Import is missing required core column mappings for this file type (CSV column map).";
  }
  if (!isDigital && (!paymentMethodColName || col[paymentMethodColName] === undefined)) {
    return "Your CSV must include the payment column mapped for Payment Method Type. Please request a new Orders export from Amazon.";
  }

  const numCols = amzNumColsForImportRows(sheet, tillerCols, tillerLabels, categoryColNum);
  const ci = amzWrittenTillerIndices_(tillerCols, tillerLabels);
  const colDebug = amzTransactionsColumnDebugLines(sheet, tillerCols, tillerLabels, numCols, categoryColNum);
  for (let di = 0; di < colDebug.length; di++) timing.push(colDebug[di]);

  if (csv.length > 1 && col[orderDateCol] !== undefined) {
    const rawOd = csv[1][col[orderDateCol]];
    const sd = new Date(rawOd);
    sd.setHours(0, 0, 0, 0);
    timing.push(
      "Server: sample CSV row 1 Order Date — raw=" +
        JSON.stringify(rawOd) +
        " getTime=" +
        sd.getTime() +
        " valid=" +
        !isNaN(sd.getTime())
    );
  }

  const output = [];
  /** Per Order ID: { totalAmount (negative sum), orderDate, payKey, lineItemCount } for offset rows (one offset per order). */
  const perOrderOffset = {};
  let duplicateCount = 0;
  const skippedRowDump = { n: 0 };
  let runTimestamp = new Date();
  let importTimestampStr = amzFormatImportTimestampStr_(runTimestamp);
  if (opts.bundleImportTimestampIso) {
    importTimestampStr = String(opts.bundleImportTimestampIso).trim();
    runTimestamp = amzParseImportTimestampToDate(importTimestampStr);
  }
  const dateAddedForRun = new Date(runTimestamp.getTime());

  const tLoopStart = Date.now();
  const csvDataRowCount = csv.length - 1;
  let aggregatedOrderCount = 0;

  function pushOneRow(r, rowsForMeta, orderDate, orderID, productName, asin, amount, accountRow, payKey) {
    const month = Utilities.formatDate(orderDate, amzActiveSpreadsheetTimeZoneOrDefault_(), "yyyy-MM");
    const week = amzGetWeekStartDate(orderDate);
    const descShort = amzFormatPurchaseDescription_(isDigital, productName);
    const descFull = amzFormatPurchaseFullDescription_(isDigital, orderID, productName, asin);

    const amazonMeta = rowsForMeta && rowsForMeta.length > 1
      ? amzBuildAmazonMetadataObjectFromRows(
          rowsForMeta,
          col,
          config.metadataMapping,
          isDigital,
          totalAmountCol,
          config.coreMappingStandard,
          config.coreMappingDigital
        )
      : amzBuildAmazonMetadataObject(
          r,
          col,
          config.metadataMapping,
          isDigital,
          config.coreMappingStandard,
          config.coreMappingDigital
        );
    if (isDigital) {
      amazonMeta.type = "digital-purchase";
      amazonMeta.id = String(orderID == null ? "" : orderID).trim();
    } else {
      if (!amazonMeta.type) amazonMeta.type = "purchase";
      amazonMeta.id = String(orderID == null ? "" : orderID).trim();
      amazonMeta.asin = String(asin == null ? "" : asin).trim();
    }
    const metadataValue = amzImportMetadataJson_(amazonMeta, importTimestampStr);

    const rowOut = new Array(numCols).fill("");
    amzSetRowDescriptionFields_(rowOut, tillerCols, tillerLabels, descShort, descFull);
    rowOut[ci.DATE - 1] = orderDate;
    rowOut[ci.AMOUNT - 1] = amount;
    rowOut[ci.TRANSACTION_ID - 1] = amzGenerateGuid();
    rowOut[ci.DATE_ADDED - 1] = dateAddedForRun;
    rowOut[ci.SOURCE - 1] = AMZ_TRANSACTIONS_SOURCE_VALUE;
    rowOut[ci.MONTH - 1] = month;
    rowOut[ci.WEEK - 1] = week;
    rowOut[ci.ACCOUNT - 1] = accountRow.ACCOUNT;
    rowOut[ci.ACCOUNT_NUMBER - 1] = accountRow.ACCOUNT_NUMBER;
    rowOut[ci.INSTITUTION - 1] = accountRow.INSTITUTION;
    rowOut[ci.ACCOUNT_ID - 1] = accountRow.ACCOUNT_ID;
    rowOut[ci.METADATA - 1] = metadataValue;
    output.push(rowOut);
  }

  if (isDigital) {
    const groups = {};
    for (let i = 1; i < csv.length; i++) {
      const r = csv[i];
      const oid = String(r[col[orderIdCol]] == null ? "" : r[col[orderIdCol]]).trim();
      if (!oid) {
        amzLogSkippedCsvDataIfUnderCap_(
          timing,
          skippedRowDump,
          "Digital orders",
          "missing Order ID",
          r
        );
        continue;
      }
      if (!groups[oid]) groups[oid] = [];
      groups[oid].push(r);
    }
    const orderIds = Object.keys(groups);
    aggregatedOrderCount = orderIds.length;
    for (let g = 0; g < orderIds.length; g++) {
      const rows = groups[orderIds[g]];
      const r = rows[0];
      let sumAmount = 0;
      for (let j = 0; j < rows.length; j++) {
        const v = parseFloat(rows[j][col[totalAmountCol]]);
        if (!isNaN(v)) sumAmount += v;
      }
      const amount = sumAmount * -1;

      const orderDateParsed = amzParseAmazonCsvDateLoose_(r[col[orderDateCol]]);
      if (!orderDateParsed) {
        amzLogSkippedCsvDataIfUnderCap_(
          timing,
          skippedRowDump,
          "Digital orders",
          "Order Date empty or not parseable",
          rows
        );
        continue;
      }
      const orderDate = orderDateParsed;
      if (cutoffStart && orderDate < cutoffStart) continue;

      const orderID = r[col[orderIdCol]];
      const productName = r[col[productNameCol]];
      const asin = r[col[asinCol]];

      const dupKeyPurchase = "digital-purchase|" + String(orderID == null ? "" : orderID).trim();
      if (existingFullDescSet.has(dupKeyPurchase)) {
        duplicateCount += 1;
        continue;
      }

      const accountRow = config.digitalUserAccount;
      const payKey = "Digital";
      const oidKey = String(orderID == null ? "" : orderID).trim();
      if (!perOrderOffset[oidKey]) {
        perOrderOffset[oidKey] = {
          totalAmount: 0,
          orderDate: new Date(orderDate.getTime()),
          payKey: payKey,
          lineItemCount: 0
        };
      }
      perOrderOffset[oidKey].totalAmount += amount;
      perOrderOffset[oidKey].lineItemCount = rows.length;

      pushOneRow(r, rows, orderDate, orderID, productName, asin, amount, accountRow, payKey);
      existingFullDescSet.add(dupKeyPurchase);
    }
  } else {
    const websiteColName = amzGetCoreCsvColumn(config, "Website", false);
    for (let i = 1; i < csv.length; i++) {
      const r = csv[i];
      const orderDateParsed = amzParseAmazonCsvDateLoose_(r[col[orderDateCol]]);
      if (!orderDateParsed) {
        amzLogSkippedCsvDataIfUnderCap_(
          timing,
          skippedRowDump,
          "Orders",
          "Order Date empty or not parseable",
          r
        );
        continue;
      }
      const orderDate = orderDateParsed;
      if (cutoffStart && orderDate < cutoffStart) continue;

      if (websiteColName && col[websiteColName] !== undefined) {
        const wf = String(r[col[websiteColName]] || "").trim().toLowerCase();
        const isWf = wf === AMZ_WHOLE_FOODS_WEBSITE;
        if (skipPanda01 && isWf) continue;
        if (skipNonPanda01 && !isWf) continue;
      }

      const orderID = r[col[orderIdCol]];
      const productName = r[col[productNameCol]];
      const asin = r[col[asinCol]];
      const oidTrim = String(orderID == null ? "" : orderID).trim();
      if (!oidTrim) {
        amzLogSkippedCsvDataIfUnderCap_(timing, skippedRowDump, "Orders", "missing Order ID", r);
        continue;
      }

      const tokenK = amzNormalizePurchaseDedupToken_(asin);
      const dupKeyPurchase = "physical-purchase-line|" + oidTrim + "|" + tokenK;

      if (existingFullDescSet.has(dupKeyPurchase)) {
        duplicateCount += 1;
        continue;
      }

      const paymentMethodType = String(r[col[paymentMethodColName]] || "").trim();
      const accountRow = paymentAccounts[paymentMethodType];
      if (!accountRow) {
        return "Payment type \"" + paymentMethodType + "\" not found. Import was stopped. Add new payment type to AMZ Import tab.";
      }

      const amount = parseFloat(r[col[totalAmountCol]]) * -1;
      const payKey = paymentMethodType;
      const oidKey = oidTrim;
      if (!perOrderOffset[oidKey]) {
        perOrderOffset[oidKey] = {
          totalAmount: 0,
          orderDate: new Date(orderDate.getTime()),
          payKey: paymentMethodType,
          lineItemCount: 0
        };
      } else if (perOrderOffset[oidKey].payKey !== paymentMethodType) {
        // Unusual: same Order ID, different payment strings — keep first row's payment for account routing
      }
      perOrderOffset[oidKey].totalAmount += amount;
      perOrderOffset[oidKey].lineItemCount += 1;

      pushOneRow(r, null, orderDate, orderID, productName, asin, amount, accountRow, payKey);
      existingFullDescSet.add(dupKeyPurchase);
    }
  }

  const tLoopEnd = Date.now();
  timing.push(
    "Server: main loop (" + csvDataRowCount + " CSV data rows" +
    (isDigital ? ", " + aggregatedOrderCount + " unique Order IDs after grouping" : "") +
    ", " + output.length + " new rows): " +
    ((tLoopEnd - tLoopStart) / 1000).toFixed(2) + " s"
  );

  if (!output.length) {
    let msg = "No new transactions found";
    if (duplicateCount > 0) {
      msg += "\n" + duplicateCount + " duplicate transactions were not imported.";
    }
    amzPushSkippedCsvDumpCapNoticeIfNeeded_(timing, skippedRowDump);
    return detectedLabel + "\n" + msg + (timing.length ? "\n" + timing.join("\n") : "");
  }

  // One balancing offset per Order ID; Date/Month/Week = that order's date. Date Added = calendar date when import runs.
  const offsetRows = [];

  const orderIdsForOffset = Object.keys(perOrderOffset);
  let offsetSkippedZeroNet = 0;
  /** Offsets written with empty Account fields because AMZ Import had no row for this payment / digital user. */
  let offsetBlankAccountFields = 0;
  for (let oi = 0; oi < orderIdsForOffset.length; oi++) {
    const oidKey = orderIdsForOffset[oi];
    const po = perOrderOffset[oidKey];
    const total = po.totalAmount;
    if (total === 0) {
      offsetSkippedZeroNet += 1;
      continue;
    }
    const resolvedAccount = isDigital ? config.digitalUserAccount : paymentAccounts[po.payKey];
    // Every order that received line items in this run must get an offset when net ≠ 0. If payment → account
    // mapping is missing, still write the offset with blank Account / Institution fields so the user can fix
    // AMZ Import and assign the row manually without hunting for a "missing" offset.
    if (!resolvedAccount) offsetBlankAccountFields += 1;
    const accountRow = resolvedAccount || amzEmptyAccountRow();

    const orderDateForOffset = new Date(po.orderDate.getTime());
    orderDateForOffset.setHours(0, 0, 0, 0);
    const offMonth = Utilities.formatDate(orderDateForOffset, amzActiveSpreadsheetTimeZoneOrDefault_(), "yyyy-MM");
    const offWeek = amzGetWeekStartDate(orderDateForOffset);

    const itemCountForOffset =
      po.lineItemCount != null && Number(po.lineItemCount) >= 1 ? Number(po.lineItemCount) : 1;
    const offDesc = amzFormatPurchaseOffsetLine_(isDigital, oidKey, itemCountForOffset);
    const offset = new Array(numCols).fill("");
    amzSetRowDescriptionFields_(offset, tillerCols, tillerLabels, offDesc, offDesc);
    offset[ci.DATE - 1] = orderDateForOffset;
    offset[ci.AMOUNT - 1] = Math.abs(total);
    offset[ci.TRANSACTION_ID - 1] = amzGenerateGuid();
    offset[ci.DATE_ADDED - 1] = dateAddedForRun;
    offset[ci.MONTH - 1] = offMonth;
    offset[ci.WEEK - 1] = offWeek;
    offset[ci.ACCOUNT - 1] = accountRow.ACCOUNT;
    offset[ci.ACCOUNT_NUMBER - 1] = accountRow.ACCOUNT_NUMBER;
    offset[ci.INSTITUTION - 1] = accountRow.INSTITUTION;
    offset[ci.ACCOUNT_ID - 1] = accountRow.ACCOUNT_ID;
    offset[ci.SOURCE - 1] = AMZ_TRANSACTIONS_SOURCE_VALUE;
    const offsetAmazonMeta = isDigital
      ? { id: String(oidKey), type: "digital-purchase-offset", lineItemCount: itemCountForOffset }
      : { id: String(oidKey), type: "purchase-offset", lineItemCount: itemCountForOffset };
    offset[ci.METADATA - 1] = amzImportMetadataJson_(offsetAmazonMeta, importTimestampStr);
    if (offsetCategory && categoryColNum) {
      offset[categoryColNum - 1] = offsetCategory;
    }
    offsetRows.push(offset);
  }

  const rowsToWrite = offsetRows.length ? output.concat(offsetRows) : output;
  const tWriteStart = Date.now();
  const appendAfterRow = amzGetLastTransactionDataRow(sheet, tillerCols, tillerLabels);
  const startRow = appendAfterRow + 1;
  const maxRowLen = amzMaxRowLength(rowsToWrite);
  const writeCols = Math.max(numCols, maxRowLen);
  if (writeCols > numCols) {
    timing.push(
      "Server: WARNING: row array width " + maxRowLen + " > numCols " + numCols + "; writing " + writeCols + " columns"
    );
  }
  const padded = amzPadRowsToWriteCols(rowsToWrite, writeCols);
  amzCoercePaddedRowsDateWeekToSerial_(padded, ci); /* see amzSheetsDateSerial_ */

  sheet.getRange(startRow, 1, padded.length, writeCols).setValues(padded);
  SpreadsheetApp.flush();
  timing.push(AMZ_IMPORT_COMMIT_COUNT_PREFIX + padded.length);
  const tWriteEnd = Date.now();
  timing.push(
    "Server: write new rows to sheet (" + output.length + " import" +
    (offsetRows.length ? ", " + offsetRows.length + " offset" : "") +
    ", " + padded.length + " total, " + writeCols + " cols): " +
    ((tWriteEnd - tWriteStart) / 1000).toFixed(2) + " s"
  );

  try {
    const d = sheet.getRange(startRow, ci.DATE).getValues()[0][0];
    const dLabel =
      d instanceof Date
        ? "Date isDate=" + !isNaN(d.getTime())
        : typeof d === "number"
          ? "type=number serial=" + d
          : "type=" + typeof d + (d != null ? " valLen=" + String(d).length : "");
    timing.push("Server: post-write first row Date column — " + dLabel);
  } catch (eRw) {
    timing.push("Server: post-write readback Date: " + (eRw.message || String(eRw)));
  }

  if (!deferPost) {
    const postLines = amzApplyTransactionsSortAndFilterCore_(
      ss,
      sheet,
      tillerCols,
      tillerLabels,
      sheetName,
      String(importTimestampStr).trim()
    );
    for (let pi = 0; pi < postLines.length; pi++) timing.push(postLines[pi]);
  }

  const tEnd = Date.now();
  timing.push(
    "Server: TOTAL importAmazonRecent time: " +
    ((tEnd - t0) / 1000).toFixed(2) + " s"
  );

  let summary = output.length + " transactions imported";
  if (duplicateCount > 0) {
    summary += "\n" + duplicateCount + " duplicate transactions were not imported.";
  }
  if (offsetSkippedZeroNet > 0) {
    summary +=
      "\n" +
      offsetSkippedZeroNet +
      " order(s) had a net total of $0 after summing line items — no offset row (offsets only apply when net is non-zero).";
  }
  if (offsetBlankAccountFields > 0) {
    summary +=
      "\n" +
      offsetBlankAccountFields +
      " offset row(s) have blank Account fields — add the payment type on AMZ Import (or assign accounts manually on Transactions).";
  }
  amzPushSkippedCsvDumpCapNoticeIfNeeded_(timing, skippedRowDump);
  timing.unshift(summary);
  timing.unshift(detectedLabel);

  return timing.join("\n");
}

/**
 * Digital Returns.csv: ASIN, Order ID, Return Date, Transaction Amount — grouped by Order ID (sum amounts).
 * Optional Digital Content Orders CSV joins Order ID to Payment Information for account routing; else Use-for-Digital row; else blank accounts.
 * @param {string} options - JSON with cutoffDateIso, offsetCategory
 * @param {string} [digitalOrdersCsv] - Optional: same-ZIP Digital Content Orders for payment lookup
 */
function importDigitalReturnsCsv(csvText, options, digitalOrdersCsv) {
  const t0 = Date.now();
  const timing = [];
  const opts = amzParseImportAmazonOptions(options);
  const offsetCategory =
    opts.offsetCategory != null && String(opts.offsetCategory).trim() !== ""
      ? String(opts.offsetCategory).trim()
      : "";
  const deferPost = opts.deferTransactionsSheetPostProcess === true;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const amzResult = getOrCreateAmzImportSheet();
  const config = readAmzImportConfig(amzResult.sheet);
  const configErr = validateAmzImportConfig(config);
  if (configErr) return configErr;

  const paymentAccounts = config.paymentAccounts;
  const paymentByOrder =
    digitalOrdersCsv != null && String(digitalOrdersCsv).trim() !== ""
      ? amzOrderIdToPaymentStringMapFromCsv(digitalOrdersCsv, config, true)
      : {};

  const tillerLabels = config.tillerLabels;
  const sheetName = tillerLabels.SHEET_NAME;
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return "Error: Sheet '" + sheetName + "' not found.";

  const tillerCols = amzGetTillerColumnMap(sheet);
  const colMapErrDr = amzValidateTransactionsImportColumns_(tillerCols, tillerLabels);
  if (colMapErrDr) return colMapErrDr;
  const categoryColNum = offsetCategory
    ? amzGetTillerColumnIndex_(tillerCols, "Category") ||
      amzGetTillerColumnIndex_(tillerCols, "Categories")
    : null;

  const csv = Utilities.parseCsv(csvText);
  if (!csv.length || !csv[0].length) {
    return "Error: Digital returns CSV is empty.";
  }
  const headers = csv[0];
  const col = {};
  headers.forEach(function (h, i) {
    if (h != null && String(h).trim() !== "") col[String(h).trim()] = i;
  });

  const DR = "digital returns.csv";
  const orderIdCol = amzGetSourceMapHeader(config, DR, "Order ID") || "Order ID";
  const returnDateCol = amzGetSourceMapHeader(config, DR, "Return Date") || "Return Date";
  const asinCol = amzGetSourceMapHeader(config, DR, "ASIN") || "ASIN";
  const totalAmountCol = amzGetSourceMapHeader(config, DR, "Transaction Amount") || "Transaction Amount";
  const needHdrs = [asinCol, orderIdCol, returnDateCol, totalAmountCol];
  for (let n = 0; n < needHdrs.length; n++) {
    if (col[needHdrs[n]] === undefined) {
      return "Digital returns CSV must include columns: " + needHdrs.join(", ") + ".";
    }
  }

  let cutoff = null;
  if (opts.cutoffDateIso) {
    cutoff = new Date(String(opts.cutoffDateIso) + "T12:00:00");
    if (isNaN(cutoff.getTime())) cutoff = null;
  }
  const cutoffStart = amzCutoffStartOfDay_(cutoff);

  const lastDataRow = amzGetLastTransactionDataRow(sheet, tillerCols, tillerLabels);
  const existingFullDescSet = new Set();
  amzAppendDuplicateKeysFromTransactions_(sheet, tillerCols, tillerLabels, lastDataRow, existingFullDescSet);
  amzAppendLegacyDuplicateKeysFromFullDescription_(sheet, tillerCols, tillerLabels, lastDataRow, existingFullDescSet);

  const groups = {};
  const skipRowDump = { n: 0 };
  let skippedMissingDigitalReturnOrderId = 0;
  for (let i = 1; i < csv.length; i++) {
    const r = csv[i];
    const oid = String(r[col[orderIdCol]] == null ? "" : r[col[orderIdCol]]).trim();
    if (!oid) {
      skippedMissingDigitalReturnOrderId += 1;
      amzLogSkippedCsvDataIfUnderCap_(timing, skipRowDump, "Digital returns", "missing Order ID", r);
      continue;
    }
    if (!groups[oid]) groups[oid] = [];
    groups[oid].push(r);
  }

  const numCols = amzNumColsForImportRows(sheet, tillerCols, tillerLabels, categoryColNum);
  const ciDr = amzWrittenTillerIndices_(tillerCols, tillerLabels);
  const colDebug = amzTransactionsColumnDebugLines(sheet, tillerCols, tillerLabels, numCols, categoryColNum);
  for (let di = 0; di < colDebug.length; di++) timing.push(colDebug[di]);

  if (csv.length > 1 && col[returnDateCol] !== undefined) {
    const rawRd = csv[1][col[returnDateCol]];
    const sd = new Date(rawRd);
    sd.setHours(0, 0, 0, 0);
    timing.push(
      "Server: sample CSV row 1 Return Date — raw=" +
        JSON.stringify(rawRd) +
        " getTime=" +
        sd.getTime() +
        " valid=" +
        !isNaN(sd.getTime())
    );
  }

  const output = [];
  const perOrderOffset = {};
  let runTimestamp = new Date();
  let importTimestampStr = amzFormatImportTimestampStr_(runTimestamp);
  if (opts.bundleImportTimestampIso) {
    importTimestampStr = String(opts.bundleImportTimestampIso).trim();
    runTimestamp = amzParseImportTimestampToDate(importTimestampStr);
  }
  const dateAddedForRun = new Date(runTimestamp.getTime());
  let duplicateCount = 0;
  const payKey = "Digital";

  const orderIds = Object.keys(groups);
  for (let g = 0; g < orderIds.length; g++) {
    const rows = groups[orderIds[g]];
    const r = rows[0];
    let sumAmount = 0;
    for (let j = 0; j < rows.length; j++) {
      const v = parseFloat(rows[j][col[totalAmountCol]]);
      if (!isNaN(v)) sumAmount += v;
    }
    const amount = sumAmount * -1;

    const orderDate = amzParseAmazonCsvDateLoose_(r[col[returnDateCol]]);
    if (!orderDate) {
      amzLogSkippedCsvDataIfUnderCap_(
        timing,
        skipRowDump,
        "Digital returns",
        "Return Date empty or not parseable",
        rows
      );
      continue;
    }
    if (cutoffStart && orderDate < cutoffStart) continue;
    if (sumAmount === 0) {
      amzLogSkippedCsvDataIfUnderCap_(
        timing,
        skipRowDump,
        "Digital returns",
        "Transaction Amount summed to $0",
        rows
      );
      continue;
    }

    const orderID = r[col[orderIdCol]];
    const asin = r[col[asinCol]];

    const dupKeyDigitalReturn =
      "digital-return|" + String(orderID == null ? "" : orderID).trim() + "|" + String(asin || "").trim();
    if (existingFullDescSet.has(dupKeyDigitalReturn)) {
      duplicateCount += 1;
      continue;
    }

    const accountRow = amzResolveDigitalReturnAccountRow(orderID, paymentByOrder, paymentAccounts, config.digitalUserAccount);

    const amazonMeta = { id: String(orderID), asin: String(asin || ""), type: "digital-return" };
    const metadataValue = amzImportMetadataJson_(amazonMeta, importTimestampStr);

    const month = Utilities.formatDate(orderDate, amzActiveSpreadsheetTimeZoneOrDefault_(), "yyyy-MM");
    const week = amzGetWeekStartDate(orderDate);
    const rowOut = new Array(numCols).fill("");
    amzSetRowDescriptionFields_(rowOut, tillerCols, tillerLabels, AMZ_DESC_DIGITAL_REFUND, AMZ_DESC_DIGITAL_REFUND);
    rowOut[ciDr.DATE - 1] = orderDate;
    rowOut[ciDr.AMOUNT - 1] = amount;
    rowOut[ciDr.TRANSACTION_ID - 1] = amzGenerateGuid();
    rowOut[ciDr.DATE_ADDED - 1] = dateAddedForRun;
    rowOut[ciDr.SOURCE - 1] = AMZ_TRANSACTIONS_SOURCE_VALUE;
    rowOut[ciDr.MONTH - 1] = month;
    rowOut[ciDr.WEEK - 1] = week;
    rowOut[ciDr.ACCOUNT - 1] = accountRow.ACCOUNT;
    rowOut[ciDr.ACCOUNT_NUMBER - 1] = accountRow.ACCOUNT_NUMBER;
    rowOut[ciDr.INSTITUTION - 1] = accountRow.INSTITUTION;
    rowOut[ciDr.ACCOUNT_ID - 1] = accountRow.ACCOUNT_ID;
    rowOut[ciDr.METADATA - 1] = metadataValue;
    output.push(rowOut);
    existingFullDescSet.add(dupKeyDigitalReturn);

    const oidKey = String(orderID == null ? "" : orderID).trim();
    if (!perOrderOffset[oidKey]) {
      perOrderOffset[oidKey] = {
        totalAmount: 0,
        orderDate: new Date(orderDate.getTime()),
        payKey: payKey
      };
    }
    perOrderOffset[oidKey].totalAmount += amount;
  }

  if (!output.length) {
    let msg = "Digital returns: no new transactions";
    if (duplicateCount > 0) msg += "\n" + duplicateCount + " duplicates skipped.";
    if (skippedMissingDigitalReturnOrderId > 0) {
      msg +=
        "\n" + skippedMissingDigitalReturnOrderId + " data row(s) skipped: missing Order ID.";
    }
    amzPushSkippedCsvDumpCapNoticeIfNeeded_(timing, skipRowDump);
    timing.unshift(msg);
    return timing.join("\n");
  }

  const offsetRows = [];
  const orderIdsForOffset = Object.keys(perOrderOffset);
  for (let oi = 0; oi < orderIdsForOffset.length; oi++) {
    const oidKey = orderIdsForOffset[oi];
    const po = perOrderOffset[oidKey];
    const total = po.totalAmount;
    if (total === 0) continue;
    const orderDateForOffset = new Date(po.orderDate.getTime());
    orderDateForOffset.setHours(0, 0, 0, 0);
    const offMonth = Utilities.formatDate(orderDateForOffset, amzActiveSpreadsheetTimeZoneOrDefault_(), "yyyy-MM");
    const offWeek = amzGetWeekStartDate(orderDateForOffset);
    const offset = new Array(numCols).fill("");
    const digRetOffDesc = amzFormatDigitalReturnOffsetLine_(oidKey);
    amzSetRowDescriptionFields_(offset, tillerCols, tillerLabels, digRetOffDesc, digRetOffDesc);
    offset[ciDr.DATE - 1] = orderDateForOffset;
    offset[ciDr.AMOUNT - 1] = Math.abs(total);
    offset[ciDr.TRANSACTION_ID - 1] = amzGenerateGuid();
    offset[ciDr.DATE_ADDED - 1] = dateAddedForRun;
    offset[ciDr.MONTH - 1] = offMonth;
    offset[ciDr.WEEK - 1] = offWeek;
    const offsetAcct = amzResolveDigitalReturnAccountRow(oidKey, paymentByOrder, paymentAccounts, config.digitalUserAccount);
    offset[ciDr.ACCOUNT - 1] = offsetAcct.ACCOUNT;
    offset[ciDr.ACCOUNT_NUMBER - 1] = offsetAcct.ACCOUNT_NUMBER;
    offset[ciDr.INSTITUTION - 1] = offsetAcct.INSTITUTION;
    offset[ciDr.ACCOUNT_ID - 1] = offsetAcct.ACCOUNT_ID;
    offset[ciDr.SOURCE - 1] = AMZ_TRANSACTIONS_SOURCE_VALUE;
    const digRetOffMeta = { id: String(oidKey), type: "digital-return-offset" };
    offset[ciDr.METADATA - 1] = amzImportMetadataJson_(digRetOffMeta, importTimestampStr);
    if (offsetCategory && categoryColNum) {
      offset[categoryColNum - 1] = offsetCategory;
    }
    offsetRows.push(offset);
  }

  const rowsToWrite = offsetRows.length ? output.concat(offsetRows) : output;
  const appendAfterRow = amzGetLastTransactionDataRow(sheet, tillerCols, tillerLabels);
  const startRow = appendAfterRow + 1;
  const maxRowLen = amzMaxRowLength(rowsToWrite);
  const writeCols = Math.max(numCols, maxRowLen);
  if (writeCols > numCols) {
    timing.push(
      "Server: WARNING: row array width " + maxRowLen + " > numCols " + numCols + "; writing " + writeCols + " columns"
    );
  }
  const padded = amzPadRowsToWriteCols(rowsToWrite, writeCols);
  amzCoercePaddedRowsDateWeekToSerial_(padded, ciDr); /* see amzSheetsDateSerial_ */
  sheet.getRange(startRow, 1, padded.length, writeCols).setValues(padded);
  SpreadsheetApp.flush();
  timing.push(AMZ_IMPORT_COMMIT_COUNT_PREFIX + padded.length);

  if (!deferPost) {
    const postLines = amzApplyTransactionsSortAndFilterCore_(
      ss,
      sheet,
      tillerCols,
      tillerLabels,
      sheetName,
      String(importTimestampStr).trim()
    );
    for (let pi = 0; pi < postLines.length; pi++) timing.push(postLines[pi]);
  }

  let summary = "Digital returns: " + output.length + " transactions imported";
  if (duplicateCount > 0) summary += "\n" + duplicateCount + " duplicates skipped.";
  if (skippedMissingDigitalReturnOrderId > 0) {
    summary +=
      "\n" + skippedMissingDigitalReturnOrderId + " data row(s) skipped: missing Order ID.";
  }
  amzPushSkippedCsvDumpCapNoticeIfNeeded_(timing, skipRowDump);
  timing.push(
    "Server: TOTAL importDigitalReturnsCsv: " + ((Date.now() - t0) / 1000).toFixed(2) + " s"
  );
  timing.unshift(summary);
  return timing.join("\n");
}

/**
 * Refund Details.csv (Order History refunds): Order ID, Refund Amount, Website;
 * Refund Date and/or Creation Date (date from Refund Date when parseable, else Creation Date).
 * Per–Order ID refund totals dedupe CSV lines that repeat the same amount and date (see amzDedupedRefundSumForOrder_).
 * Website is stored in metadata / full description only — not filtered by Orders vs Whole Foods toggles.
 * When Order History.csv is available in the same bundle, join by Order ID to resolve payment method; else account fields are left blank for manual fix.
 * @param {string} orderHistoryCsv - Optional raw Order History CSV from ZIP (for payment join)
 */
function importRefundDetailsCsv(csvText, options, orderHistoryCsv) {
  const t0 = Date.now();
  const timing = [];
  const opts = amzParseImportAmazonOptions(options);
  const offsetCategory =
    opts.offsetCategory != null && String(opts.offsetCategory).trim() !== ""
      ? String(opts.offsetCategory).trim()
      : "";
  const deferPost = opts.deferTransactionsSheetPostProcess === true;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const amzResult = getOrCreateAmzImportSheet();
  const config = readAmzImportConfig(amzResult.sheet);
  const configErr = validateAmzImportConfig(config);
  if (configErr) return configErr;

  const paymentAccounts = config.paymentAccounts;
  const paymentByOrder =
    orderHistoryCsv != null && String(orderHistoryCsv).trim() !== ""
      ? amzOrderIdToPaymentStringMapFromCsv(orderHistoryCsv, config, false)
      : {};

  const tillerLabels = config.tillerLabels;
  const sheetName = tillerLabels.SHEET_NAME;
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return "Error: Sheet '" + sheetName + "' not found.";

  const tillerCols = amzGetTillerColumnMap(sheet);
  const colMapErrRd = amzValidateTransactionsImportColumns_(tillerCols, tillerLabels);
  if (colMapErrRd) return colMapErrRd;
  const categoryColNum = offsetCategory
    ? amzGetTillerColumnIndex_(tillerCols, "Category") ||
      amzGetTillerColumnIndex_(tillerCols, "Categories")
    : null;

  const csv = Utilities.parseCsv(csvText);
  if (!csv.length || !csv[0].length) {
    return "Error: Orders returns CSV is empty.";
  }
  const headers = csv[0];
  const col = {};
  headers.forEach(function (h, i) {
    if (h != null && String(h).trim() !== "") col[String(h).trim()] = i;
  });

  const RD = "refund details.csv";
  const orderIdCol = amzGetSourceMapHeader(config, RD, "Order ID") || "Order ID";
  const refundAmountCol = amzGetSourceMapHeader(config, RD, "Refund Amount") || "Refund Amount";
  const websiteCol = amzGetSourceMapHeader(config, RD, "Website") || "Website";
  const refundDateCol = amzGetSourceMapHeader(config, RD, "Refund Date") || "Refund Date";
  const creationDateCol = amzGetSourceMapHeader(config, RD, "Creation Date") || "Creation Date";
  const contractIdHeader = amzGetSourceMapHeader(config, RD, "Contract ID");
  const contractIdCol =
    contractIdHeader && col[contractIdHeader] !== undefined ? contractIdHeader : null;

  const needCore = [orderIdCol, refundAmountCol, websiteCol];
  for (let n = 0; n < needCore.length; n++) {
    if (col[needCore[n]] === undefined) {
      return "Orders returns CSV must include columns: " + needCore.join(", ") + ".";
    }
  }
  if (col[refundDateCol] === undefined && col[creationDateCol] === undefined) {
    return (
      "Orders returns CSV must include at least one date column: " + refundDateCol + " or " + creationDateCol + "."
    );
  }

  let cutoff = null;
  if (opts.cutoffDateIso) {
    cutoff = new Date(String(opts.cutoffDateIso) + "T12:00:00");
    if (isNaN(cutoff.getTime())) cutoff = null;
  }
  const cutoffStart = amzCutoffStartOfDay_(cutoff);

  const lastDataRow = amzGetLastTransactionDataRow(sheet, tillerCols, tillerLabels);
  const existingFullDescSet = new Set();
  amzAppendDuplicateKeysFromTransactions_(sheet, tillerCols, tillerLabels, lastDataRow, existingFullDescSet);
  amzAppendLegacyDuplicateKeysFromFullDescription_(sheet, tillerCols, tillerLabels, lastDataRow, existingFullDescSet);

  const groups = {};
  const skipRowDump = { n: 0 };
  let skippedMissingRefundOrderId = 0;
  for (let i = 1; i < csv.length; i++) {
    const r = csv[i];
    const oid = String(r[col[orderIdCol]] == null ? "" : r[col[orderIdCol]]).trim();
    if (!oid) {
      skippedMissingRefundOrderId += 1;
      amzLogSkippedCsvDataIfUnderCap_(timing, skipRowDump, "Orders returns", "missing Order ID", r);
      continue;
    }

    // Do not apply Order History Website (Amazon.com vs Whole Foods) filters here — they are tied to
    // the "Orders" checkbox and would drop all Amazon.com refunds when only Orders Returns is selected.

    if (!groups[oid]) groups[oid] = [];
    groups[oid].push(r);
  }

  const numCols = amzNumColsForImportRows(sheet, tillerCols, tillerLabels, categoryColNum);
  const ciRd = amzWrittenTillerIndices_(tillerCols, tillerLabels);
  const colDebug = amzTransactionsColumnDebugLines(sheet, tillerCols, tillerLabels, numCols, categoryColNum);
  for (let di = 0; di < colDebug.length; di++) timing.push(colDebug[di]);

  if (csv.length > 1) {
    if (col[refundDateCol] !== undefined) {
      const rawFd = csv[1][col[refundDateCol]];
      const sd = new Date(rawFd);
      sd.setHours(0, 0, 0, 0);
      timing.push(
        "Server: sample CSV row 1 Refund Date — raw=" +
          JSON.stringify(rawFd) +
          " getTime=" +
          sd.getTime() +
          " valid=" +
          !isNaN(sd.getTime())
      );
    }
    if (col[creationDateCol] !== undefined) {
      const rawCd = csv[1][col[creationDateCol]];
      const resolved = amzResolveRefundDetailsOrderDate_(csv[1], col, config);
      timing.push(
        "Server: sample CSV row 1 Creation Date — raw=" +
          JSON.stringify(rawCd) +
          " resolvedOrderDate=" +
          (resolved ? resolved.getTime() : "null")
      );
    }
  }

  const output = [];
  const perOrderOffset = {};
  let runTimestamp = new Date();
  let importTimestampStr = amzFormatImportTimestampStr_(runTimestamp);
  if (opts.bundleImportTimestampIso) {
    importTimestampStr = String(opts.bundleImportTimestampIso).trim();
    runTimestamp = amzParseImportTimestampToDate(importTimestampStr);
  }
  const dateAddedForRun = new Date(runTimestamp.getTime());
  let duplicateCount = 0;
  let skippedInvalidRefundDate = 0;

  const orderIds = Object.keys(groups);
  for (let g = 0; g < orderIds.length; g++) {
    const rows = groups[orderIds[g]];
    const r = rows[0];
    const sumRefund = amzDedupedRefundSumForOrder_(rows, col, refundAmountCol, config);
    const amount = sumRefund;

    const orderDate = amzResolveRefundDetailsOrderDate_(r, col, config);
    if (!orderDate) {
      skippedInvalidRefundDate += 1;
      amzLogSkippedCsvDataIfUnderCap_(
        timing,
        skipRowDump,
        "Orders returns",
        "Could not parse a transaction date from Refund Date or Creation Date",
        rows
      );
      continue;
    }
    if (cutoffStart && orderDate < cutoffStart) continue;
    if (sumRefund === 0) {
      continue;
    }

    const orderID = r[col[orderIdCol]];
    const websiteVal = String(r[col[websiteCol]] || "").trim();

    const oidStr = String(orderID == null ? "" : orderID).trim();
    const paymentHint = oidStr && paymentByOrder[oidStr] != null ? String(paymentByOrder[oidStr]).trim() : "";
    const accountRow = amzResolvePhysicalRefundAccountRow(paymentHint, paymentAccounts);

    const dupKeyRefund =
      "refund-detail|" + oidStr + "|" + Number(sumRefund).toFixed(2);
    if (existingFullDescSet.has(dupKeyRefund)) {
      duplicateCount += 1;
      continue;
    }
    if (amzAnyLegacyReturnKeyForOrderId_(existingFullDescSet, oidStr)) {
      duplicateCount += 1;
      continue;
    }
    if (contractIdCol) {
      let hitLegacy = false;
      const seenLegacyC = {};
      for (let jc = 0; jc < rows.length; jc++) {
        const cid = String(rows[jc][col[contractIdCol]] == null ? "" : rows[jc][col[contractIdCol]]).trim().toLowerCase();
        if (!cid || seenLegacyC[cid]) continue;
        seenLegacyC[cid] = 1;
        if (existingFullDescSet.has("legacy-return|" + oidStr + "|" + cid)) {
          hitLegacy = true;
          break;
        }
      }
      if (hitLegacy) {
        duplicateCount += 1;
        continue;
      }
    }

    const amazonMeta = {
      orderId: String(orderID),
      refundAmount: sumRefund,
      refundDate: String(
        col[refundDateCol] !== undefined && r[col[refundDateCol]] != null ? r[col[refundDateCol]] : ""
      ),
      website: websiteVal,
      type: "refund-detail"
    };
    const metadataValue = amzImportMetadataJson_(amazonMeta, importTimestampStr);

    const month = Utilities.formatDate(orderDate, amzActiveSpreadsheetTimeZoneOrDefault_(), "yyyy-MM");
    const week = amzGetWeekStartDate(orderDate);
    const rowOut = new Array(numCols).fill("");
    const refundDescLine = amzFormatPhysicalRefundFullDescription_(orderID);
    amzSetRowDescriptionFields_(rowOut, tillerCols, tillerLabels, refundDescLine, refundDescLine);
    rowOut[ciRd.DATE - 1] = orderDate;
    rowOut[ciRd.AMOUNT - 1] = amount;
    rowOut[ciRd.TRANSACTION_ID - 1] = amzGenerateGuid();
    rowOut[ciRd.DATE_ADDED - 1] = dateAddedForRun;
    rowOut[ciRd.SOURCE - 1] = AMZ_TRANSACTIONS_SOURCE_VALUE;
    rowOut[ciRd.MONTH - 1] = month;
    rowOut[ciRd.WEEK - 1] = week;
    rowOut[ciRd.ACCOUNT - 1] = accountRow.ACCOUNT;
    rowOut[ciRd.ACCOUNT_NUMBER - 1] = accountRow.ACCOUNT_NUMBER;
    rowOut[ciRd.INSTITUTION - 1] = accountRow.INSTITUTION;
    rowOut[ciRd.ACCOUNT_ID - 1] = accountRow.ACCOUNT_ID;
    rowOut[ciRd.METADATA - 1] = metadataValue;
    output.push(rowOut);
    existingFullDescSet.add(dupKeyRefund);
    if (contractIdCol) {
      const seenOutC = {};
      for (let jo = 0; jo < rows.length; jo++) {
        const cid = String(rows[jo][col[contractIdCol]] == null ? "" : rows[jo][col[contractIdCol]]).trim().toLowerCase();
        if (!cid || seenOutC[cid]) continue;
        seenOutC[cid] = 1;
        existingFullDescSet.add("legacy-return|" + oidStr + "|" + cid);
      }
    }

    const oidKey = oidStr;
    if (!perOrderOffset[oidKey]) {
      perOrderOffset[oidKey] = {
        totalAmount: 0,
        orderDate: new Date(orderDate.getTime()),
        payKey: paymentHint
      };
    }
    perOrderOffset[oidKey].totalAmount += amount;
  }

  if (!output.length) {
    let msg = "Orders returns: no new transactions";
    if (duplicateCount > 0) msg += "\n" + duplicateCount + " duplicates skipped.";
    if (skippedMissingRefundOrderId > 0) {
      msg +=
        "\n" + skippedMissingRefundOrderId + " data row(s) skipped: missing Order ID.";
    }
    if (skippedInvalidRefundDate > 0) {
      msg +=
        "\n" +
        skippedInvalidRefundDate +
        " row(s) skipped: could not parse Refund Date or Creation Date (rows are not dated to today).";
    }
    amzPushSkippedCsvDumpCapNoticeIfNeeded_(timing, skipRowDump);
    timing.unshift(msg);
    return timing.join("\n");
  }

  const offsetRows = [];
  const orderIdsForOffset = Object.keys(perOrderOffset);
  for (let oi = 0; oi < orderIdsForOffset.length; oi++) {
    const oidKey = orderIdsForOffset[oi];
    const po = perOrderOffset[oidKey];
    const total = po.totalAmount;
    if (total === 0) continue;
    const orderDateForOffset = new Date(po.orderDate.getTime());
    orderDateForOffset.setHours(0, 0, 0, 0);
    const offMonth = Utilities.formatDate(orderDateForOffset, amzActiveSpreadsheetTimeZoneOrDefault_(), "yyyy-MM");
    const offWeek = amzGetWeekStartDate(orderDateForOffset);
    const offset = new Array(numCols).fill("");
    const physRefOffDesc = amzFormatPhysicalRefundOffsetLine_(oidKey);
    amzSetRowDescriptionFields_(offset, tillerCols, tillerLabels, physRefOffDesc, physRefOffDesc);
    offset[ciRd.DATE - 1] = orderDateForOffset;
    offset[ciRd.AMOUNT - 1] = -Math.abs(total);
    offset[ciRd.TRANSACTION_ID - 1] = amzGenerateGuid();
    offset[ciRd.DATE_ADDED - 1] = dateAddedForRun;
    offset[ciRd.MONTH - 1] = offMonth;
    offset[ciRd.WEEK - 1] = offWeek;
    const payHint = paymentByOrder[oidKey] != null ? String(paymentByOrder[oidKey]).trim() : "";
    const acct = amzResolvePhysicalRefundAccountRow(payHint, paymentAccounts);
    offset[ciRd.ACCOUNT - 1] = acct.ACCOUNT;
    offset[ciRd.ACCOUNT_NUMBER - 1] = acct.ACCOUNT_NUMBER;
    offset[ciRd.INSTITUTION - 1] = acct.INSTITUTION;
    offset[ciRd.ACCOUNT_ID - 1] = acct.ACCOUNT_ID;
    offset[ciRd.SOURCE - 1] = AMZ_TRANSACTIONS_SOURCE_VALUE;
    const physRefOffMeta = { id: String(oidKey), type: "physical-refund-offset" };
    offset[ciRd.METADATA - 1] = amzImportMetadataJson_(physRefOffMeta, importTimestampStr);
    if (offsetCategory && categoryColNum) {
      offset[categoryColNum - 1] = offsetCategory;
    }
    offsetRows.push(offset);
  }

  const rowsToWrite = offsetRows.length ? output.concat(offsetRows) : output;
  const appendAfterRow = amzGetLastTransactionDataRow(sheet, tillerCols, tillerLabels);
  const startRow = appendAfterRow + 1;
  const maxRowLen = amzMaxRowLength(rowsToWrite);
  const writeCols = Math.max(numCols, maxRowLen);
  if (writeCols > numCols) {
    timing.push(
      "Server: WARNING: row array width " + maxRowLen + " > numCols " + numCols + "; writing " + writeCols + " columns"
    );
  }
  const padded = amzPadRowsToWriteCols(rowsToWrite, writeCols);
  amzCoercePaddedRowsDateWeekToSerial_(padded, ciRd); /* see amzSheetsDateSerial_ */
  sheet.getRange(startRow, 1, padded.length, writeCols).setValues(padded);
  SpreadsheetApp.flush();
  timing.push(AMZ_IMPORT_COMMIT_COUNT_PREFIX + padded.length);

  if (!deferPost) {
    const postLines = amzApplyTransactionsSortAndFilterCore_(
      ss,
      sheet,
      tillerCols,
      tillerLabels,
      sheetName,
      String(importTimestampStr).trim()
    );
    for (let pi = 0; pi < postLines.length; pi++) timing.push(postLines[pi]);
  }

  let summary = "Orders returns: " + output.length + " transactions imported";
  if (duplicateCount > 0) summary += "\n" + duplicateCount + " duplicates skipped.";
  if (skippedMissingRefundOrderId > 0) {
    summary +=
      "\n" + skippedMissingRefundOrderId + " data row(s) skipped: missing Order ID.";
  }
  if (skippedInvalidRefundDate > 0) {
    summary +=
      "\n" +
      skippedInvalidRefundDate +
      " row(s) skipped: could not parse Refund Date or Creation Date (rows are not dated to today).";
  }
  amzPushSkippedCsvDumpCapNoticeIfNeeded_(timing, skipRowDump);
  timing.push(
    "Server: TOTAL importRefundDetailsCsv: " + ((Date.now() - t0) / 1000).toFixed(2) + " s"
  );
  timing.unshift(summary);
  return timing.join("\n");
}

/**
 * Shared bundle timestamp + import flags → JSON options string (defer sheet post-process).
 * @param {Object} bundleLike - cutoffDateIso, offsetCategory, includeWholeFoods, includePhysicalOrders, bundleImportTimestampIso
 */
function amzImportBundleOptsStr_(bundleLike, bundleImportTimestampIso) {
  const optsObj = {
    cutoffDateIso: bundleLike.cutoffDateIso || "",
    offsetCategory: bundleLike.offsetCategory || "",
    skipPanda01: bundleLike.includeWholeFoods === false,
    skipNonPanda01: bundleLike.includePhysicalOrders === false,
    deferTransactionsSheetPostProcess: true,
    bundleImportTimestampIso: bundleImportTimestampIso
  };
  return JSON.stringify(optsObj);
}

/**
 * After deferred imports: flush, sort/filter Transactions for this bundle timestamp.
 * @param {string} bundleImportTimestampIso
 * @returns {string} log block (starts with "\n\n=== Transactions sheet ===\n") or errors
 */
function amzImportBundleTransactionsPostLog_(bundleImportTimestampIso) {
  SpreadsheetApp.flush();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const amzResult = getOrCreateAmzImportSheet();
  const config = readAmzImportConfig(amzResult.sheet);
  const tillerLabels = config.tillerLabels;
  const sheetName = tillerLabels.SHEET_NAME;
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return "\n\n=== Transactions sheet ===\nError: Sheet '" + sheetName + "' not found.";
  }
  const tillerCols = amzGetTillerColumnMap(sheet);
  if (amzGetTillerColumnIndex_(tillerCols, tillerLabels.DATE) == null) {
    return "\n\n=== Transactions sheet ===\nError: Transactions sheet is missing Date column.";
  }
  const postLines = amzApplyTransactionsSortAndFilterCore_(
    ss,
    sheet,
    tillerCols,
    tillerLabels,
    sheetName,
    String(bundleImportTimestampIso).trim()
  );
  return "\n\n=== Transactions sheet ===\n" + postLines.join("\n");
}

/**
 * One step of a chunked ZIP import (one RPC per step) to cap request payload size.
 * Payload shape: step, bundleImportTimestampIso (optional; set on first chunk), cutoff/offset/include* flags,
 * and CSV fields needed for that step only (orderHistoryCsv, digitalOrdersCsv, digitalReturnsCsv, refundDetailsCsv).
 * @param {string} payloadJson
 * @returns {Object} { bundleImportTimestampIso, step, log: string, didImport: boolean }
 */
function importAmazonBundleChunk(payloadJson) {
  let payload;
  try {
    payload = JSON.parse(payloadJson);
  } catch (e) {
    throw new Error("Invalid import chunk payload.");
  }
  const step = String(payload.step || "").trim();
  if (step !== "finalize") {
    const catErr = amzValidateBundleOffsetCategory_(payload.offsetCategory);
    if (catErr) throw new Error(catErr);
  }
  let ts = "";
  const rawTs = payload.bundleImportTimestampIso;
  if (rawTs != null && rawTs !== "") {
    ts = String(rawTs).trim();
  }
  if (!ts) {
    ts = amzFormatImportTimestampStr_(new Date());
  }
  const optsStr = amzImportBundleOptsStr_(payload, ts);

  if (step === "finalize") {
    const postLog = amzImportBundleTransactionsPostLog_(ts);
    const finLog = postLog.replace(/^\n+/, "");
    return {
      bundleImportTimestampIso: ts,
      step: step,
      log: finLog,
      didImport: false,
      rowsWritten: amzImportCommitCountFromLog_(finLog)
    };
  }

  let logText = "";
  let didImport = false;
  const runOH = payload.includePhysicalOrders || payload.includeWholeFoods;

  if (step === "orderHistory") {
    if (runOH && payload.orderHistoryCsv) {
      didImport = true;
      logText = "=== Orders ===\n" + importAmazonRecent(payload.orderHistoryCsv, null, optsStr);
    } else if (runOH) {
      logText = "=== Orders ===\n" + "(skipped: Order History.csv not found in ZIP)";
    } else {
      logText = "=== Orders ===\n" + "(skipped: not selected)";
    }
  } else if (step === "digitalOrders") {
    if (payload.includeDigitalOrders && payload.digitalOrdersCsv) {
      didImport = true;
      logText = "=== Digital orders ===\n" + importAmazonRecent(payload.digitalOrdersCsv, null, optsStr);
    } else if (payload.includeDigitalOrders) {
      logText = "=== Digital orders ===\n" + "(skipped: file not in ZIP)";
    } else {
      logText = "=== Digital orders ===\n" + "(skipped: not selected)";
    }
  } else if (step === "digitalReturns") {
    if (payload.includeDigitalReturns && payload.digitalReturnsCsv) {
      didImport = true;
      logText =
        "=== Digital returns ===\n" +
        importDigitalReturnsCsv(payload.digitalReturnsCsv, optsStr, payload.digitalOrdersCsv || null);
    } else if (payload.includeDigitalReturns) {
      logText = "=== Digital returns ===\n" + "(skipped: file not in ZIP)";
    } else {
      logText = "=== Digital returns ===\n" + "(skipped: not selected)";
    }
  } else if (step === "refundDetails") {
    if (payload.includeRefundDetails && payload.refundDetailsCsv) {
      didImport = true;
      logText =
        "=== Orders returns ===\n" +
        importRefundDetailsCsv(payload.refundDetailsCsv, optsStr, payload.orderHistoryCsv || null);
    } else if (payload.includeRefundDetails) {
      logText = "=== Orders returns ===\n" + "(skipped: file not in ZIP)";
    } else {
      logText = "=== Orders returns ===\n" + "(skipped: not selected)";
    }
  } else {
    throw new Error("Unknown import chunk step: " + step);
  }

  return {
    bundleImportTimestampIso: ts,
    step: step,
    log: logText,
    didImport: didImport,
    rowsWritten: amzImportCommitCountFromLog_(logText)
  };
}

/**
 * ZIP wizard: run selected pipelines with one cutoff and offset category.
 * @param {string} bundleJson
 */
function importAmazonBundle(bundleJson) {
  let bundle;
  try {
    bundle = JSON.parse(bundleJson);
  } catch (e) {
    return "Error: invalid import bundle.";
  }

  const catErr0 = amzValidateBundleOffsetCategory_(bundle.offsetCategory);
  if (catErr0) return "Error: " + catErr0;

  const bundleStarted = new Date();
  const bundleImportTimestampIso = amzFormatImportTimestampStr_(bundleStarted);
  const optsStr = amzImportBundleOptsStr_(bundle, bundleImportTimestampIso);
  const sections = [];
  let importRuns = 0;

  const runOH = bundle.includePhysicalOrders || bundle.includeWholeFoods;
  if (runOH && bundle.orderHistoryCsv) {
    importRuns += 1;
    sections.push("=== Orders ===\n" + importAmazonRecent(bundle.orderHistoryCsv, null, optsStr));
  } else if (runOH) {
    sections.push("=== Orders ===\n" + "(skipped: Order History.csv not found in ZIP)");
  }

  if (bundle.includeDigitalOrders && bundle.digitalOrdersCsv) {
    importRuns += 1;
    sections.push("=== Digital orders ===\n" + importAmazonRecent(bundle.digitalOrdersCsv, null, optsStr));
  } else if (bundle.includeDigitalOrders) {
    sections.push("=== Digital orders ===\n" + "(skipped: file not in ZIP)");
  }

  if (bundle.includeDigitalReturns && bundle.digitalReturnsCsv) {
    importRuns += 1;
    sections.push(
      "=== Digital returns ===\n" + importDigitalReturnsCsv(bundle.digitalReturnsCsv, optsStr, bundle.digitalOrdersCsv || null)
    );
  } else if (bundle.includeDigitalReturns) {
    sections.push("=== Digital returns ===\n" + "(skipped: file not in ZIP)");
  }

  if (bundle.includeRefundDetails && bundle.refundDetailsCsv) {
    importRuns += 1;
    sections.push(
      "=== Orders returns ===\n" + importRefundDetailsCsv(bundle.refundDetailsCsv, optsStr, bundle.orderHistoryCsv || null)
    );
  }

  if (!sections.length) {
    return "No import steps were run. Select at least one order type and ensure the ZIP contains the matching CSV(s).";
  }

  let out = sections.join("\n\n");
  if (importRuns > 0) {
    out += amzImportBundleTransactionsPostLog_(bundleImportTimestampIso);
  }
  return out;
}
